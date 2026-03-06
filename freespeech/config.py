from __future__ import annotations

from dataclasses import asdict, dataclass, fields
import json
import os
from pathlib import Path
from typing import Any


APP_NAME = "FreeSpeech"
APP_DIR = Path(os.environ.get("APPDATA", str(Path.home()))) / APP_NAME
CONFIG_PATH = APP_DIR / "settings.json"


@dataclass
class Settings:
    always_on_top: bool = True
    dark_mode: bool = True
    dark_mode_manual: bool = False
    ui_scale_percent: int = 100
    capture_delay_ms: int = 100
    max_chars: int = 4000
    start_on_boot: bool = False
    save_generated_speech: bool = False
    generated_speech_dir: str = ""

    voice: str = "en-US-JennyNeural"
    rate: int = 0
    pitch: int = 0
    volume: int = 0


def ensure_app_dir() -> None:
    APP_DIR.mkdir(parents=True, exist_ok=True)


def _sanitize_int(value: Any, default: int, minimum: int, maximum: int) -> int:
    try:
        parsed = int(value)
    except (TypeError, ValueError):
        return default
    return max(minimum, min(maximum, parsed))


def _sanitize_bool(value: Any, default: bool) -> bool:
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)):
        return int(value) != 0

    normalized = str(value or "").strip().lower()
    if normalized in {"1", "true", "yes", "on"}:
        return True
    if normalized in {"0", "false", "no", "off", ""}:
        return False
    return bool(default)


def _sanitize_settings(settings: Settings) -> Settings:
    settings.always_on_top = _sanitize_bool(settings.always_on_top, True)
    settings.dark_mode = _sanitize_bool(settings.dark_mode, True)
    settings.dark_mode_manual = _sanitize_bool(settings.dark_mode_manual, False)
    settings.ui_scale_percent = _sanitize_int(settings.ui_scale_percent, 100, 70, 200)
    settings.capture_delay_ms = _sanitize_int(settings.capture_delay_ms, 100, 0, 1500)
    settings.max_chars = _sanitize_int(settings.max_chars, 4000, 100, 100000)
    settings.start_on_boot = _sanitize_bool(settings.start_on_boot, False)
    settings.save_generated_speech = _sanitize_bool(settings.save_generated_speech, False)
    settings.generated_speech_dir = str(settings.generated_speech_dir or "").strip()

    settings.voice = str(settings.voice or "").strip() or "en-US-JennyNeural"
    settings.rate = _sanitize_int(settings.rate, 0, -100, 100)
    settings.pitch = _sanitize_int(settings.pitch, 0, -100, 100)
    settings.volume = _sanitize_int(settings.volume, 0, -100, 100)
    return settings


def _filtered_payload(payload: dict[str, Any]) -> dict[str, Any]:
    allowed = {item.name for item in fields(Settings)}
    return {k: v for k, v in payload.items() if k in allowed}


def load_settings() -> Settings:
    ensure_app_dir()
    if not CONFIG_PATH.exists():
        return Settings()

    try:
        raw = CONFIG_PATH.read_text(encoding="utf-8")
        payload = json.loads(raw)
        if not isinstance(payload, dict):
            return Settings()
        return _sanitize_settings(Settings(**_filtered_payload(payload)))
    except Exception:
        return Settings()


def save_settings(settings: Settings) -> None:
    ensure_app_dir()
    cleaned = _sanitize_settings(settings)
    CONFIG_PATH.write_text(
        json.dumps(asdict(cleaned), indent=2, ensure_ascii=True),
        encoding="utf-8",
    )
