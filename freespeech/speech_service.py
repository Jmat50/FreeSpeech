from __future__ import annotations

import ctypes
from datetime import datetime
import os
from pathlib import Path
import queue
import re
import tempfile
import threading
import time
from typing import Callable

from .backends import build_backend
from .config import APP_DIR, Settings


LogFn = Callable[[str], None]
SettingsProvider = Callable[[], Settings]
PlaybackStateFn = Callable[[bool], None]


class WindowsMciPlayer:
    def __init__(self) -> None:
        self._alias = "freespeech_audio"
        self._lock = threading.Lock()
        self._current_alias: str | None = None
        self._current_device_id: str | None = None
        self._current_file: Path | None = None

    @staticmethod
    def _send(command: str) -> str:
        buffer = ctypes.create_unicode_buffer(512)
        err = ctypes.windll.winmm.mciSendStringW(command, buffer, len(buffer), 0)
        if err != 0:
            error_text = ctypes.create_unicode_buffer(512)
            ctypes.windll.winmm.mciGetErrorStringW(err, error_text, len(error_text))
            raise RuntimeError(error_text.value or f"MCI error code {err}")
        return buffer.value

    @staticmethod
    def _send_quiet(command: str) -> None:
        try:
            WindowsMciPlayer._send(command)
        except Exception:
            pass

    def play_bytes(self, audio: bytes) -> None:
        if not audio:
            raise RuntimeError("Audio buffer is empty.")

        with self._lock:
            self._stop_unlocked()
            fd, temp_path = tempfile.mkstemp(prefix="freespeech_", suffix=".mp3")
            os.close(fd)
            path = Path(temp_path)
            path.write_bytes(audio)
            quoted = str(path.resolve()).replace('"', '""')
            unique_alias = f"{self._alias}_{int(time.time() * 1000)}"
            open_response = self._send(f'open "{quoted}" type mpegvideo alias {unique_alias}')
            self._send(f"play {unique_alias}")
            self._current_alias = unique_alias
            resolved_device_id = str(open_response or "").strip()
            self._current_device_id = resolved_device_id if resolved_device_id else None
            self._current_file = path

    def stop(self) -> None:
        with self._lock:
            self._stop_unlocked()

    def is_playing(self) -> bool:
        with self._lock:
            targets: list[str] = []
            for candidate in (self._current_alias, self._current_device_id, self._alias):
                value = str(candidate or "").strip()
                if value and value not in targets:
                    targets.append(value)
            if not targets:
                return False

            for target in targets:
                try:
                    mode = self._send(f"status {target} mode").strip().lower()
                except Exception:
                    continue
                if mode in {"playing", "seeking", "paused", "open"}:
                    return True
            return False

    def _stop_unlocked(self) -> None:
        aliases: list[str] = []
        active_alias = str(self._current_alias or "").strip()
        if active_alias:
            aliases.append(active_alias)
        active_device_id = str(self._current_device_id or "").strip()
        if active_device_id and active_device_id not in aliases:
            aliases.append(active_device_id)
        if self._alias not in aliases:
            aliases.append(self._alias)

        for alias in aliases:
            self._send_quiet(f"stop {alias}")
            self._send_quiet(f"close {alias}")

        # Fallback for systems where alias/device-id addressing is flaky.
        self._send_quiet("stop all")
        self._send_quiet("close all")

        self._current_alias = None
        self._current_device_id = None
        if self._current_file is not None:
            try:
                self._current_file.unlink(missing_ok=True)
            except Exception:
                pass
            self._current_file = None


class SpeechService:
    def __init__(
        self,
        settings_provider: SettingsProvider,
        logger: LogFn,
        playback_state_callback: PlaybackStateFn | None = None,
    ) -> None:
        self._settings_provider = settings_provider
        self._logger = logger
        self._playback_state_callback = playback_state_callback

        self._queue: "queue.Queue[str]" = queue.Queue()
        self._control_queue: "queue.Queue[tuple[str, threading.Event | None]]" = queue.Queue()
        self._stop_event = threading.Event()
        self._backend_lock = threading.Lock()
        self._backend = None
        self._player = WindowsMciPlayer()
        self._state_lock = threading.Lock()
        self._playback_token = 0
        self._playback_active = False
        self._stop_generation = 0
        self._prewarm_lock = threading.Lock()
        self._prewarm_started = False

        self._worker = threading.Thread(target=self._run, daemon=True)
        self._worker.start()

    def _current_backend(self, settings: Settings):
        with self._backend_lock:
            if self._backend is not None:
                return self._backend

            backend = build_backend(settings)
            self._backend = backend
            return backend

    def enqueue_speak(self, text: str, replace: bool = True) -> None:
        clean = str(text or "").strip()
        if not clean:
            return

        if bool(replace):
            while True:
                try:
                    self._queue.get_nowait()
                except queue.Empty:
                    break

        self._queue.put(clean)

    def stop_audio(self) -> None:
        while True:
            try:
                self._queue.get_nowait()
            except queue.Empty:
                break
        with self._state_lock:
            self._playback_token += 1
            self._stop_generation += 1

        # MCI playback control can be thread-sensitive on some systems.
        # Route stop through the worker thread first, then fall back locally.
        ack = threading.Event()
        try:
            self._control_queue.put_nowait(("stop", ack))
        except Exception:
            ack.set()
        if not ack.wait(timeout=1.0):
            self._player.stop()

        deadline = time.monotonic() + 1.2
        while time.monotonic() < deadline:
            if not self._player.is_playing():
                break
            time.sleep(0.06)
            self._player.stop()
        self._notify_playback_state(False)
        self._logger("Playback stopped.")

    def is_playing(self) -> bool:
        with self._state_lock:
            return bool(self._playback_active)

    def close(self) -> None:
        self._stop_event.set()
        self.stop_audio()

    def prewarm_backend_async(self) -> None:
        if self._stop_event.is_set():
            return
        with self._prewarm_lock:
            if self._prewarm_started:
                return
            self._prewarm_started = True
        threading.Thread(target=self._prewarm_backend_worker, daemon=True).start()

    def _prewarm_backend_worker(self) -> None:
        started_at = time.monotonic()
        try:
            settings = self._settings_provider()
            backend = self._current_backend(settings)
            # Tiny synthesis to warm network/session path so first real speak starts faster.
            _ = backend.synthesize(
                text="Ready.",
                voice=settings.voice,
                rate=settings.rate,
                pitch=settings.pitch,
                volume=settings.volume,
            )
            elapsed = max(0.0, time.monotonic() - started_at)
            self._logger(f"TTS engine preloaded ({elapsed:.2f}s).")
        except Exception as exc:
            self._logger(f"TTS warm-up skipped: {exc}")

    def _run(self) -> None:
        while not self._stop_event.is_set():
            self._drain_control_queue()
            try:
                text = self._queue.get(timeout=0.05)
            except queue.Empty:
                continue

            try:
                start_generation = self._current_stop_generation()
                settings = self._settings_provider()
                backend = self._current_backend(settings)
                audio = backend.synthesize(
                    text=text,
                    voice=settings.voice,
                    rate=settings.rate,
                    pitch=settings.pitch,
                    volume=settings.volume,
                )
                # Stop may have been pressed while synthesis was running.
                if self._stop_requested_since(start_generation):
                    continue
                self._save_generated_audio_if_enabled(audio, text, settings)
                # Stop may have been pressed while optional MP3 save was running.
                if self._stop_requested_since(start_generation):
                    continue
                self._player.play_bytes(audio)
                # Stop may have been pressed while opening/starting playback.
                if self._stop_requested_since(start_generation):
                    self._player.stop()
                    continue
                playback_token = self._next_playback_token()
                estimated_seconds = self._estimate_playback_seconds(text)
                self._notify_playback_state(True)
                threading.Thread(
                    target=self._watch_playback_until_complete,
                    args=(playback_token, estimated_seconds),
                    daemon=True,
                ).start()
                self._logger(
                    f"Speaking {min(len(text), settings.max_chars)} chars with "
                    f"{settings.voice}."
                )
            except Exception as exc:
                self._notify_playback_state(False)
                self._logger(f"Synthesis failed: {exc}")

    def _drain_control_queue(self) -> None:
        while True:
            try:
                command, ack = self._control_queue.get_nowait()
            except queue.Empty:
                return
            try:
                if command == "stop":
                    self._player.stop()
            finally:
                if ack is not None:
                    try:
                        ack.set()
                    except Exception:
                        pass

    def _save_generated_audio_if_enabled(self, audio: bytes, text: str, settings: Settings) -> None:
        if not bool(settings.save_generated_speech):
            return
        if not audio:
            return

        target_raw = str(settings.generated_speech_dir or "").strip()
        target_dir = Path(target_raw) if target_raw else (APP_DIR / "generated_speech")
        target_dir.mkdir(parents=True, exist_ok=True)

        stamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
        snippet = re.sub(r"[^A-Za-z0-9_-]+", "_", str(text or "").strip())[:36].strip("_")
        file_name = f"{stamp}.mp3" if not snippet else f"{stamp}_{snippet}.mp3"
        output_path = target_dir / file_name
        output_path.write_bytes(audio)
        self._logger(f"Saved MP3: {output_path}")

    def _next_playback_token(self) -> int:
        with self._state_lock:
            self._playback_token += 1
            return self._playback_token

    def _current_stop_generation(self) -> int:
        with self._state_lock:
            return int(self._stop_generation)

    def _stop_requested_since(self, generation: int) -> bool:
        return self._current_stop_generation() != int(generation)

    def _notify_playback_state(self, active: bool) -> None:
        callback = self._playback_state_callback
        if callback is None:
            return
        should_emit = False
        with self._state_lock:
            normalized = bool(active)
            if normalized != self._playback_active:
                self._playback_active = normalized
                should_emit = True
        if not should_emit:
            return
        try:
            callback(bool(active))
        except Exception:
            pass

    @staticmethod
    def _estimate_playback_seconds(text: str) -> float:
        normalized = re.sub(r"\s+", " ", str(text or "").strip())
        if not normalized:
            return 1.0
        chars = len(normalized)
        punctuation = sum(normalized.count(ch) for ch in ".!?;,")
        estimate = (chars / 16.0) + (punctuation * 0.08)
        return max(1.0, min(240.0, float(estimate)))

    def _watch_playback_until_complete(self, token: int, estimated_seconds: float) -> None:
        start_time = time.monotonic()
        saw_playing_state = False
        false_streak = 0
        false_streak_required = 4
        minimum_before_force_complete = max(1.0, float(estimated_seconds))

        while not self._stop_event.is_set():
            with self._state_lock:
                if token != self._playback_token:
                    return
            try:
                playing = bool(self._player.is_playing())
            except Exception:
                playing = False

            elapsed = time.monotonic() - start_time
            if playing:
                saw_playing_state = True
                false_streak = 0
                time.sleep(0.15)
                continue

            false_streak += 1
            if false_streak < false_streak_required:
                time.sleep(0.15)
                continue

            # If MCI mode is flaky and never reports "playing", keep stop visible
            # for a reasonable estimated speech duration.
            if (not saw_playing_state) and elapsed < minimum_before_force_complete:
                time.sleep(0.15)
                continue

            break
        with self._state_lock:
            if token != self._playback_token:
                return
        self._notify_playback_state(False)
