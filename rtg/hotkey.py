"""Global hotkey, active even when the window is minimised.

Built on pynput, which installs an OS-level keyboard hook, so the shortcut
fires no matter which application has focus. The listener runs in its own
thread: the callback is therefore *not* on the Qt thread and must hand over
through a signal (see ``rtg.gui.main_window.HotkeyBridge``).

This module stays free of Qt so the parsing rules can be unit tested without a
display.
"""

from __future__ import annotations

import threading
from typing import Callable, List, Optional

from .errors import ResultGeneratorError

DEFAULT_HOTKEY = "Ctrl+Shift+S"


class HotkeyError(ResultGeneratorError):
    """Raised for an unusable hotkey or a listener that cannot start."""


#: Modifier spellings accepted from the user, mapped to pynput syntax.
_MODIFIERS = {
    "ctrl": "<ctrl>", "control": "<ctrl>", "ctl": "<ctrl>",
    "shift": "<shift>", "shft": "<shift>",
    "alt": "<alt>", "altgr": "<alt_gr>",
    "win": "<cmd>", "windows": "<cmd>", "meta": "<cmd>",
    "cmd": "<cmd>", "super": "<cmd>",
}

#: Display order, so "Shift+Ctrl+S" is always shown as "Ctrl+Shift+S".
_MODIFIER_ORDER = ("<ctrl>", "<alt>", "<alt_gr>", "<shift>", "<cmd>")
_MODIFIER_NAMES = {
    "<ctrl>": "Ctrl", "<alt>": "Alt", "<alt_gr>": "AltGr",
    "<shift>": "Shift", "<cmd>": "Win",
}

#: Non-character keys pynput knows by name.
_SPECIAL = {
    "esc": "<esc>", "escape": "<esc>",
    "space": "<space>", "spasi": "<space>",
    "tab": "<tab>", "enter": "<enter>", "return": "<enter>",
    "backspace": "<backspace>", "delete": "<delete>", "del": "<delete>",
    "insert": "<insert>", "ins": "<insert>",
    "home": "<home>", "end": "<end>",
    "pageup": "<page_up>", "page_up": "<page_up>", "pgup": "<page_up>",
    "pagedown": "<page_down>", "page_down": "<page_down>", "pgdn": "<page_down>",
    "up": "<up>", "down": "<down>", "left": "<left>", "right": "<right>",
    "printscreen": "<print_screen>", "print_screen": "<print_screen>",
    "prtsc": "<print_screen>", "pause": "<pause>",
}
_SPECIAL.update({f"f{n}": f"<f{n}>" for n in range(1, 25)})

#: Reverse of _SPECIAL, for showing a parsed hotkey back to the user.
_DISPLAY_SPECIAL = {
    "<esc>": "Esc", "<space>": "Space", "<tab>": "Tab", "<enter>": "Enter",
    "<backspace>": "Backspace", "<delete>": "Delete", "<insert>": "Insert",
    "<home>": "Home", "<end>": "End", "<page_up>": "PageUp",
    "<page_down>": "PageDown", "<up>": "Up", "<down>": "Down",
    "<left>": "Left", "<right>": "Right", "<print_screen>": "PrintScreen",
    "<pause>": "Pause",
}
_DISPLAY_SPECIAL.update({f"<f{n}>": f"F{n}" for n in range(1, 25)})


def _split(text: str) -> List[str]:
    parts = [part.strip().lower() for part in str(text).replace("-", "+").split("+")]
    return [part for part in parts if part]


def parse_hotkey(text: str) -> str:
    """Translate ``"Ctrl+Shift+S"`` into pynput's ``"<ctrl>+<shift>+s"``.

    Raises :class:`HotkeyError` with a message meant for the user.
    """
    parts = _split(text)
    if not parts:
        raise HotkeyError("Hotkey tidak boleh kosong.")

    modifiers: List[str] = []
    keys: List[str] = []
    for part in parts:
        if part in _MODIFIERS:
            token = _MODIFIERS[part]
            if token not in modifiers:
                modifiers.append(token)
        elif part in _SPECIAL:
            keys.append(_SPECIAL[part])
        elif len(part) == 1 and (part.isalnum() or part in "`-=[]\\;',./"):
            keys.append(part)
        else:
            raise HotkeyError(f"Tombol '{part}' tidak dikenali untuk hotkey.")

    if not keys:
        raise HotkeyError(
            "Hotkey harus punya satu tombol selain modifier "
            "(contoh: Ctrl+Shift+S)."
        )
    if len(keys) > 1:
        raise HotkeyError(
            "Hotkey hanya boleh punya satu tombol utama, ditemukan: "
            + ", ".join(keys)
        )
    if not modifiers:
        raise HotkeyError(
            "Hotkey harus memakai minimal satu modifier (Ctrl, Alt, Shift, "
            "atau Win) supaya tidak bentrok dengan pengetikan biasa."
        )

    ordered = [m for m in _MODIFIER_ORDER if m in modifiers]
    return "+".join(ordered + keys)


def format_hotkey(text: str) -> str:
    """Normalise a hotkey for display: ``"shift+ctrl+s"`` -> ``"Ctrl+Shift+S"``."""
    parsed = parse_hotkey(text)
    out: List[str] = []
    for token in parsed.split("+"):
        if token in _MODIFIER_NAMES:
            out.append(_MODIFIER_NAMES[token])
        elif token in _DISPLAY_SPECIAL:
            out.append(_DISPLAY_SPECIAL[token])
        else:
            out.append(token.upper())
    return "+".join(out)


def is_valid(text: str) -> bool:
    try:
        parse_hotkey(text)
    except HotkeyError:
        return False
    return True


class HotkeyListener:
    """Keeps a global hotkey registered until it is stopped.

    The callback runs on the listener's own thread. Keep it short and marshal
    to the UI thread; a slow callback delays every later key event.
    """

    def __init__(self, sequence: str, callback: Callable[[], None]) -> None:
        self.sequence = sequence
        self.callback = callback
        self._listener = None
        self._lock = threading.Lock()

    @property
    def running(self) -> bool:
        return self._listener is not None

    def start(self) -> None:
        """Register the hotkey. Safe to call when already running."""
        with self._lock:
            if self._listener is not None:
                return
            combination = parse_hotkey(self.sequence)
            try:
                from pynput import keyboard  # noqa: PLC0415 - optional at import time
            except ImportError as exc:  # pragma: no cover - environment dependent
                raise HotkeyError(
                    "Modul 'pynput' belum terpasang. Jalankan: pip install pynput",
                    detail=str(exc),
                ) from exc
            try:
                listener = keyboard.GlobalHotKeys({combination: self._fire})
                listener.daemon = True
                listener.start()
            except Exception as exc:  # pragma: no cover - OS dependent
                raise HotkeyError(
                    f"Gagal mendaftarkan hotkey '{format_hotkey(self.sequence)}': {exc}",
                    detail=str(exc),
                ) from exc
            self._listener = listener

    def _fire(self) -> None:
        try:
            self.callback()
        except Exception:  # pragma: no cover - never kill the listener thread
            pass

    def stop(self) -> None:
        """Unregister the hotkey. Safe to call when not running."""
        with self._lock:
            listener, self._listener = self._listener, None
        if listener is not None:
            try:
                listener.stop()
            except Exception:  # pragma: no cover
                pass

    def restart(self, sequence: Optional[str] = None) -> None:
        """Re-register, optionally with a different key combination."""
        if sequence is not None:
            parse_hotkey(sequence)       # validate before dropping the old one
            self.sequence = sequence
        self.stop()
        self.start()

    def __enter__(self) -> "HotkeyListener":
        self.start()
        return self

    def __exit__(self, *exc_info) -> bool:
        self.stop()
        return False
