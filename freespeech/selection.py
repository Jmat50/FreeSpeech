from __future__ import annotations

import ctypes
import threading
import time
import uuid

import pyperclip
from pynput.keyboard import Controller, Key


def is_shift_pressed() -> bool:
    return bool(ctypes.windll.user32.GetAsyncKeyState(0x10) & 0x8000)


class SelectionCapture:
    def __init__(self) -> None:
        self._keyboard = Controller()
        self._lock = threading.Lock()

    def _send_copy(self) -> None:
        with self._keyboard.pressed(Key.ctrl):
            self._keyboard.press("c")
            self._keyboard.release("c")

    def capture(self, delay_ms: int = 100) -> str:
        with self._lock:
            marker = f"__freespeech_marker__{uuid.uuid4()}__"
            previous_clipboard = ""
            clipboard_available = False

            try:
                previous_clipboard = pyperclip.paste()
                clipboard_available = True
            except Exception:
                clipboard_available = False

            if clipboard_available:
                try:
                    pyperclip.copy(marker)
                except Exception:
                    clipboard_available = False

            if delay_ms > 0:
                time.sleep(delay_ms / 1000.0)

            self._send_copy()
            deadline = time.monotonic() + 0.6
            captured = ""

            while time.monotonic() < deadline:
                time.sleep(0.02)
                try:
                    current = pyperclip.paste()
                except Exception:
                    continue
                if not current:
                    continue
                if clipboard_available and current == marker:
                    continue
                captured = str(current).strip()
                if captured:
                    break

            if clipboard_available:
                try:
                    pyperclip.copy(previous_clipboard)
                except Exception:
                    pass

            return captured
