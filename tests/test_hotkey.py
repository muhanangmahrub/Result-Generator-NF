"""Unit tests for hotkey parsing and the listener lifecycle.

The listener is exercised against a fake pynput so the tests never install a
real OS-level keyboard hook.
"""

from __future__ import annotations

import sys
import types

import pytest

from rtg.hotkey import (
    DEFAULT_HOTKEY,
    HotkeyError,
    HotkeyListener,
    format_hotkey,
    is_valid,
    parse_hotkey,
)


# --------------------------------------------------------------------------- #
# Parsing
# --------------------------------------------------------------------------- #

@pytest.mark.parametrize(
    "raw, expected",
    [
        ("Ctrl+Shift+S", "<ctrl>+<shift>+s"),
        ("ctrl+shift+s", "<ctrl>+<shift>+s"),
        ("  CTRL + SHIFT + S ", "<ctrl>+<shift>+s"),
        ("shift+ctrl+s", "<ctrl>+<shift>+s"),      # order is normalised
        ("Ctrl-Shift-S", "<ctrl>+<shift>+s"),      # dashes accepted
        ("Alt+F9", "<alt>+<f9>"),
        ("Ctrl+Alt+PrintScreen", "<ctrl>+<alt>+<print_screen>"),
        ("Win+D", "<cmd>+d"),
        ("Control+Shift+1", "<ctrl>+<shift>+1"),
        ("Ctrl+Alt+Space", "<ctrl>+<alt>+<space>"),
        ("Ctrl+Shift+Insert", "<ctrl>+<shift>+<insert>"),
        ("Ctrl+Shift+PgDn", "<ctrl>+<shift>+<page_down>"),
        ("Ctrl+.", "<ctrl>+."),
    ],
)
def test_parse_hotkey(raw, expected):
    assert parse_hotkey(raw) == expected


def test_default_hotkey_is_valid():
    assert parse_hotkey(DEFAULT_HOTKEY) == "<ctrl>+<shift>+s"


@pytest.mark.parametrize(
    "raw, message",
    [
        ("", "tidak boleh kosong"),
        ("   ", "tidak boleh kosong"),
        ("Ctrl", "satu tombol selain modifier"),
        ("Ctrl+Shift", "satu tombol selain modifier"),
        ("S", "minimal satu modifier"),
        ("Ctrl+S+D", "hanya boleh punya satu tombol utama"),
        ("Ctrl+Shift+Banana", "tidak dikenali"),
    ],
)
def test_parse_hotkey_errors(raw, message):
    with pytest.raises(HotkeyError) as excinfo:
        parse_hotkey(raw)
    assert message in excinfo.value.message


def test_modifier_without_key_is_rejected_because_it_would_block_typing():
    # A bare letter would fire while the user types in any application.
    assert not is_valid("s")
    assert is_valid("Ctrl+Shift+S")


@pytest.mark.parametrize(
    "raw, expected",
    [
        ("ctrl+shift+s", "Ctrl+Shift+S"),
        ("shift+ctrl+s", "Ctrl+Shift+S"),
        ("alt+f9", "Alt+F9"),
        ("ctrl+alt+printscreen", "Ctrl+Alt+PrintScreen"),
        ("win+d", "Win+D"),
        ("ctrl+shift+pgup", "Ctrl+Shift+PageUp"),
    ],
)
def test_format_hotkey(raw, expected):
    assert format_hotkey(raw) == expected


def test_format_hotkey_is_idempotent():
    once = format_hotkey("shift+ctrl+s")
    assert format_hotkey(once) == once


# --------------------------------------------------------------------------- #
# Listener
# --------------------------------------------------------------------------- #

class FakeGlobalHotKeys:
    """Stand-in for pynput.keyboard.GlobalHotKeys."""

    instances: list = []

    def __init__(self, mapping):
        self.mapping = mapping
        self.started = False
        self.stopped = False
        self.daemon = False
        FakeGlobalHotKeys.instances.append(self)

    def start(self):
        self.started = True

    def stop(self):
        self.stopped = True

    def press(self, combination):
        """Pretend the user pressed the registered combination."""
        self.mapping[combination]()


@pytest.fixture
def fake_pynput(monkeypatch):
    FakeGlobalHotKeys.instances = []
    keyboard = types.ModuleType("pynput.keyboard")
    keyboard.GlobalHotKeys = FakeGlobalHotKeys
    package = types.ModuleType("pynput")
    package.keyboard = keyboard
    monkeypatch.setitem(sys.modules, "pynput", package)
    monkeypatch.setitem(sys.modules, "pynput.keyboard", keyboard)
    return FakeGlobalHotKeys


def test_listener_registers_the_parsed_combination(fake_pynput):
    calls = []
    listener = HotkeyListener("Ctrl+Shift+S", lambda: calls.append(1))
    listener.start()
    try:
        assert listener.running
        hook = fake_pynput.instances[-1]
        assert list(hook.mapping) == ["<ctrl>+<shift>+s"]
        assert hook.started and hook.daemon

        hook.press("<ctrl>+<shift>+s")
        assert calls == [1]
    finally:
        listener.stop()
    assert not listener.running
    assert fake_pynput.instances[-1].stopped


def test_start_twice_registers_only_once(fake_pynput):
    listener = HotkeyListener("Ctrl+Shift+S", lambda: None)
    listener.start()
    listener.start()
    try:
        assert len(fake_pynput.instances) == 1
    finally:
        listener.stop()


def test_stop_is_safe_when_not_running(fake_pynput):
    HotkeyListener("Ctrl+Shift+S", lambda: None).stop()  # must not raise


def test_restart_with_a_new_sequence(fake_pynput):
    listener = HotkeyListener("Ctrl+Shift+S", lambda: None)
    listener.start()
    listener.restart("Alt+F9")
    try:
        assert listener.sequence == "Alt+F9"
        assert list(fake_pynput.instances[-1].mapping) == ["<alt>+<f9>"]
        assert fake_pynput.instances[0].stopped
    finally:
        listener.stop()


def test_restart_validates_before_dropping_the_running_hotkey(fake_pynput):
    listener = HotkeyListener("Ctrl+Shift+S", lambda: None)
    listener.start()
    try:
        with pytest.raises(HotkeyError):
            listener.restart("Banana")
        # The old hotkey is still registered.
        assert listener.running
        assert listener.sequence == "Ctrl+Shift+S"
        assert not fake_pynput.instances[0].stopped
    finally:
        listener.stop()


def test_invalid_sequence_fails_on_start(fake_pynput):
    listener = HotkeyListener("Ctrl", lambda: None)
    with pytest.raises(HotkeyError):
        listener.start()
    assert not listener.running


def test_callback_errors_do_not_kill_the_listener_thread(fake_pynput):
    def boom():
        raise RuntimeError("callback meledak")

    listener = HotkeyListener("Ctrl+Shift+S", boom)
    listener.start()
    try:
        fake_pynput.instances[-1].press("<ctrl>+<shift>+s")   # must not raise
        assert listener.running
    finally:
        listener.stop()


def test_context_manager(fake_pynput):
    with HotkeyListener("Ctrl+Shift+S", lambda: None) as listener:
        assert listener.running
    assert not listener.running
