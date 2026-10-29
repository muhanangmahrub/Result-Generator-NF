"""Theme and icon helpers (offscreen)."""

from __future__ import annotations

import sys

import pytest

pytest.importorskip("PySide6")

from PySide6.QtGui import QIcon, QPalette                  # noqa: E402
from PySide6.QtWidgets import QApplication                 # noqa: E402

from rtg.gui import icons, theme                           # noqa: E402


@pytest.fixture(scope="session")
def qapp():
    return QApplication.instance() or QApplication([])


@pytest.mark.parametrize("tokens", [theme.LIGHT, theme.DARK])
def test_qss_is_fully_expanded(tokens):
    qss = theme.build_qss(tokens)
    assert "QMainWindow" in qss
    assert "%(" not in qss                      # every placeholder was filled
    assert tokens["accent"] in qss


def test_light_and_dark_define_the_same_tokens():
    assert set(theme.LIGHT) == set(theme.DARK)


def test_apply_theme_sets_stylesheet_and_accent(qapp):
    tokens = theme.apply_theme(qapp)
    assert "QToolBar" in qapp.styleSheet()
    highlight = qapp.palette().color(QPalette.ColorRole.Highlight).name().lower()
    assert highlight == tokens["accent"].lower()


def test_icon_returns_a_qicon(qapp):
    assert isinstance(icons.icon("save"), QIcon)


def test_unknown_icon_name_is_empty_not_an_error(qapp):
    assert icons.icon("nope-not-there").isNull()


def test_missing_qtawesome_gives_empty_icons(qapp, monkeypatch):
    monkeypatch.setitem(sys.modules, "qtawesome", None)
    icons.clear_cache()
    try:
        assert icons.icon("save").isNull()
    finally:
        icons.clear_cache()


def test_every_named_icon_resolves(qapp):
    pytest.importorskip("qtawesome")
    icons.clear_cache()
    missing = [name for name in icons.ICONS if icons.icon(name).isNull()]
    assert missing == []
