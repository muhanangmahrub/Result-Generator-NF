"""Icons for the GUI, drawn from the Phosphor set through qtawesome.

qtawesome is optional: without it every icon is an empty ``QIcon`` and the
buttons simply keep their text/tooltips.
"""

from __future__ import annotations

from functools import lru_cache
from typing import Dict, Optional, Tuple

from PySide6.QtGui import QIcon

from . import theme

#: semantic name -> candidate qtawesome names, tried in order.
ICONS: Dict[str, Tuple[str, ...]] = {
    "new": ("ph.file-plus", "mdi6.file-plus-outline"),
    "open": ("ph.folder-open", "mdi6.folder-open-outline"),
    "save": ("ph.floppy-disk", "mdi6.content-save-outline"),
    "reimport": ("ph.arrows-clockwise", "mdi6.refresh"),
    "metadata": ("ph.sliders-horizontal", "mdi6.tune"),
    "capture": ("ph.camera", "mdi6.camera-outline"),
    "generate": ("ph.file-text", "mdi6.file-document-outline"),
    "pdf": ("ph.file-pdf", "mdi6.file-pdf-box"),
    "excel": ("ph.table", "mdi6.file-excel-outline"),
    "settings": ("ph.gear-six", "mdi6.cog-outline"),
    "add": ("ph.plus", "mdi6.plus"),
    "caption": ("ph.pencil-simple", "mdi6.pencil-outline"),
    "move": ("ph.arrow-bend-up-right", "mdi6.arrow-right-bold-outline"),
    "delete": ("ph.trash", "mdi6.trash-can-outline"),
    "search": ("ph.magnifying-glass", "mdi6.magnify"),
    "module": ("ph.table", "mdi6.table"),
    "scenario": ("ph.folder-simple", "mdi6.folder-outline"),
    "done": ("ph.check-circle", "mdi6.check-circle-outline"),
    "empty": ("ph.circle", "mdi6.circle-outline"),
    "review": ("ph.warning", "mdi6.alert-outline"),
    "warning": ("ph.warning-circle", "mdi6.alert-circle-outline"),
    "orphan": ("ph.link-break", "mdi6.link-off"),
    "folder": ("ph.folder-notch-open", "mdi6.folder-open-outline"),
    "info": ("ph.info", "mdi6.information-outline"),
    "checks": ("ph.checks", "mdi6.check-all"),
}


@lru_cache(maxsize=None)
def _cached(name: str, color: str, disabled: str) -> QIcon:
    try:
        import qtawesome as qta
    except ImportError:
        return QIcon()
    for candidate in ICONS.get(name, ()):
        try:
            return qta.icon(candidate, color=color, color_disabled=disabled)
        except Exception:        # unknown glyph name in this qtawesome version
            continue
    return QIcon()


def icon(name: str, color: Optional[str] = None) -> QIcon:
    """The icon for ``name`` in the current theme, or an empty icon."""
    tokens = theme.tokens()
    return _cached(name, color or tokens["text2"], tokens["text3"])


def clear_cache() -> None:
    """Forget cached icons (call after the theme changed)."""
    _cached.cache_clear()
