from __future__ import annotations

import asyncio
from dataclasses import dataclass
from typing import Any

from .config import Settings


class BackendError(RuntimeError):
    pass


@dataclass(frozen=True)
class VoiceInfo:
    short_name: str
    locale: str = ""
    gender: str = ""


def to_rate(value: int) -> str:
    return f"{int(value):+d}%"


def to_pitch(value: int) -> str:
    return f"{int(value):+d}Hz"


def to_volume(value: int) -> str:
    return f"{int(value):+d}%"


class PythonEdgeBackend:
    id = "python"
    label = "Python (rany2/edge-tts)"

    @staticmethod
    def _load_module() -> Any:
        try:
            import edge_tts  # type: ignore
        except ModuleNotFoundError as exc:
            raise BackendError(
                "Python backend requires `edge-tts`. Install with: pip install edge-tts"
            ) from exc
        return edge_tts

    def list_voices(self) -> list[VoiceInfo]:
        edge_tts = self._load_module()

        async def _list() -> list[dict[str, Any]]:
            return await edge_tts.list_voices()

        raw = asyncio.run(_list())
        voices: list[VoiceInfo] = []
        for item in raw:
            if not isinstance(item, dict):
                continue
            short_name = str(item.get("ShortName") or item.get("Name") or "").strip()
            if not short_name:
                continue
            voices.append(
                VoiceInfo(
                    short_name=short_name,
                    locale=str(item.get("Locale") or "").strip(),
                    gender=str(item.get("Gender") or "").strip(),
                )
            )
        voices.sort(key=lambda entry: entry.short_name)
        return voices

    def synthesize(
        self,
        text: str,
        voice: str,
        rate: int,
        pitch: int,
        volume: int,
    ) -> bytes:
        clean_text = str(text or "").strip()
        if not clean_text:
            raise BackendError("Cannot synthesize empty text.")

        edge_tts = self._load_module()

        async def _synthesize() -> bytes:
            communicate = edge_tts.Communicate(
                text=clean_text,
                voice=voice,
                rate=to_rate(rate),
                pitch=to_pitch(pitch),
                volume=to_volume(volume),
            )
            audio = bytearray()
            async for chunk in communicate.stream():
                if chunk.get("type") == "audio":
                    audio.extend(chunk.get("data") or b"")
            if not audio:
                raise BackendError("Edge TTS returned an empty audio stream.")
            return bytes(audio)

        return asyncio.run(_synthesize())


def build_backend(settings: Settings) -> PythonEdgeBackend:
    _ = settings
    return PythonEdgeBackend()
