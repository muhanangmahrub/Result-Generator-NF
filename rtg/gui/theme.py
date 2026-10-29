"""Colours, fonts and the single stylesheet of the GUI (light + dark).

All colours come from the token dicts below; nothing outside this module should
hard-code a hex value.
"""

from __future__ import annotations

from typing import Dict

from PySide6.QtGui import QColor, QFont, QFontDatabase, QGuiApplication, QPalette
from PySide6.QtCore import Qt
from PySide6.QtWidgets import QApplication

LIGHT: Dict[str, str] = dict(
    window="#F5F5F7", surface="#FFFFFF", sidebar="#EFEFF2",
    border="#E3E3E8", separator="#D8D8DE", text="#1D1D1F",
    text2="#6E6E73", text3="#A1A1A6", accent="#007AFF",
    accent_hover="#0062CC", accent_soft="rgba(0,122,255,0.14)",
    success="#34C759", warning="#FF9F0A", danger="#FF3B30",
    hover="rgba(0,0,0,0.05)", on_accent="#FFFFFF",
    select_soft="#CEDFF4",   # accent_soft blended over the sidebar (opaque)
)

DARK: Dict[str, str] = dict(
    window="#1E1E1E", surface="#262628", sidebar="#2A2A2D",
    border="#3A3A3C", separator="#38383A", text="#F5F5F7",
    text2="#A1A1A6", text3="#6E6E73", accent="#0A84FF",
    accent_hover="#409CFF", accent_soft="rgba(10,132,255,0.22)",
    success="#30D158", warning="#FF9F0A", danger="#FF453A",
    hover="rgba(255,255,255,0.06)", on_accent="#FFFFFF",
    select_soft="#233E5B",
)

FONT_CANDIDATES = ("Segoe UI Variable Text", "Segoe UI", "Helvetica Neue", "Arial")


def is_dark() -> bool:
    app = QGuiApplication.instance()
    if app is None:
        return False
    try:
        return app.styleHints().colorScheme() == Qt.ColorScheme.Dark
    except AttributeError:       # Qt < 6.5
        return False


def tokens() -> Dict[str, str]:
    return DARK if is_dark() else LIGHT


def pick_font(point_size: float = 10.0) -> QFont:
    families = set(QFontDatabase.families())
    for name in FONT_CANDIDATES:
        if name in families:
            return _sized(QFont(name), point_size)
    return _sized(QFont(), point_size)


def _sized(font: QFont, point_size: float) -> QFont:
    font.setPointSizeF(point_size)
    return font


def build_qss(t: Dict[str, str]) -> str:
    """One stylesheet for the whole application, built from the tokens."""
    return """
QMainWindow, QDialog { background: %(window)s; }
QWidget { color: %(text)s; }
QToolTip { background: %(surface)s; color: %(text)s; border: 1px solid %(border)s;
           padding: 4px 6px; }

QToolBar { background: %(window)s; border: none; border-bottom: 1px solid %(separator)s;
           padding: 4px 8px; spacing: 2px; }
QToolBar::separator { background: %(separator)s; width: 1px; margin: 6px 6px; }
QToolButton { border: none; border-radius: 6px; padding: 5px; background: transparent; }
QToolButton:hover { background: %(hover)s; }
QToolButton:pressed { background: %(accent_soft)s; }
QToolButton:disabled { background: transparent; color: %(text3)s; }
QToolButton[primary="true"] { background: %(accent)s; color: %(on_accent)s; padding: 5px 12px; }
QToolButton[primary="true"]:hover { background: %(accent_hover)s; }
QToolButton[soft="true"] { background: %(accent_soft)s; color: %(accent)s; padding: 5px 12px; }
QToolButton[soft="true"]:hover { background: %(accent_soft)s; }

QPushButton { background: %(surface)s; border: 1px solid %(border)s; border-radius: 6px;
              padding: 4px 14px; min-height: 20px; }
QPushButton:hover { background: %(hover)s; }
QPushButton:pressed { background: %(accent_soft)s; }
QPushButton:disabled { color: %(text3)s; }
QPushButton:default, QPushButton[primary="true"] { background: %(accent)s;
              border: 1px solid %(accent)s; color: %(on_accent)s; }
QPushButton:default:hover, QPushButton[primary="true"]:hover { background: %(accent_hover)s; }

QLineEdit, QComboBox, QSpinBox, QDateEdit, QPlainTextEdit, QTextBrowser {
    background: %(surface)s; border: 1px solid %(border)s; border-radius: 6px;
    padding: 4px 8px; selection-background-color: %(accent)s;
    selection-color: %(on_accent)s; }
QLineEdit:focus, QComboBox:focus, QSpinBox:focus, QDateEdit:focus,
QPlainTextEdit:focus { border: 1px solid %(accent)s; }
QLineEdit#search { background: %(hover)s; border: 1px solid transparent; border-radius: 7px; }
QLineEdit#search:focus { border: 1px solid %(accent)s; }
QTextBrowser { border: none; padding: 0; }
QComboBox::drop-down { border: none; width: 20px; }
QComboBox QAbstractItemView { background: %(surface)s; border: 1px solid %(border)s;
    selection-background-color: %(accent)s; selection-color: %(on_accent)s; outline: 0; }

QWidget#sidebarPane { background: %(sidebar)s; }
QTreeWidget#sidebar { background: %(sidebar)s; border: none; outline: 0;
    selection-background-color: transparent; }
QTreeWidget#sidebar::item { height: 26px; padding-left: 2px; }
QTreeWidget#sidebar::item:hover { background: %(hover)s; }
QTreeWidget#sidebar::item:selected:active { background: %(accent)s; color: %(on_accent)s; }
QTreeWidget#sidebar::item:selected:!active { background: %(accent_soft)s; color: %(text)s; }
QTreeWidget#sidebar::branch { background: transparent; }
QTreeWidget#sidebar::branch:selected:active { background: %(accent)s; }
QTreeWidget#sidebar::branch:selected:!active { background: %(select_soft)s; }

QTabWidget::pane { border: none; background: %(surface)s; }
QTabWidget::tab-bar { alignment: center; }
QTabBar { background: transparent; }
QTabBar::tab { background: %(hover)s; color: %(text2)s; padding: 4px 16px; margin: 8px 0;
               border: none; }
QTabBar::tab:first { border-top-left-radius: 7px; border-bottom-left-radius: 7px; }
QTabBar::tab:last { border-top-right-radius: 7px; border-bottom-right-radius: 7px; }
QTabBar::tab:only-one { border-radius: 7px; }
QTabBar::tab:selected { background: %(accent)s; color: %(on_accent)s; font-weight: 600; }
QTabBar::tab:hover:!selected { color: %(text)s; }

QHeaderView::section { background: %(surface)s; color: %(text2)s; border: none;
    border-bottom: 1px solid %(separator)s; padding: 4px 8px; }
QTableWidget { background: %(surface)s; border: 1px solid %(border)s; border-radius: 6px;
    gridline-color: %(border)s; selection-background-color: %(accent_soft)s;
    selection-color: %(text)s; }
QListWidget { background: %(window)s; border: none; outline: 0; }
QListWidget::item:selected { background: %(accent_soft)s; color: %(text)s; border-radius: 6px; }

QScrollBar:vertical { background: transparent; width: 10px; margin: 0; }
QScrollBar:horizontal { background: transparent; height: 10px; margin: 0; }
QScrollBar::handle:vertical { background: %(text3)s; border-radius: 4px; min-height: 30px;
    margin: 1px; }
QScrollBar::handle:horizontal { background: %(text3)s; border-radius: 4px; min-width: 30px;
    margin: 1px; }
QScrollBar::add-line, QScrollBar::sub-line { width: 0; height: 0; }
QScrollBar::add-page, QScrollBar::sub-page { background: transparent; }

QSplitter::handle { background: %(separator)s; }
QSplitter::handle:horizontal { width: 1px; }

QStatusBar { background: %(window)s; color: %(text2)s; border-top: 1px solid %(separator)s; }
QStatusBar::item { border: none; }

QMenuBar { background: %(window)s; }
QMenuBar::item { padding: 4px 10px; border-radius: 4px; background: transparent; }
QMenuBar::item:selected { background: %(hover)s; }
QMenu { background: %(surface)s; border: 1px solid %(border)s; border-radius: 8px;
    padding: 4px; }
QMenu::item { padding: 5px 22px 5px 12px; border-radius: 4px; }
QMenu::item:selected { background: %(accent)s; color: %(on_accent)s; }
QMenu::item:disabled { color: %(text3)s; }
QMenu::separator { height: 1px; background: %(separator)s; margin: 4px 8px; }

QProgressBar#progress { background: %(hover)s; border: none; border-radius: 3px;
    max-height: 6px; min-height: 6px; }
QProgressBar#progress::chunk { background: %(success)s; border-radius: 3px; }

QLabel[role="secondary"] { color: %(text2)s; }
QLabel[role="title"] { font-size: 15pt; font-weight: 600; }
QLabel[role="danger"] { color: %(danger)s; }
""" % t


def apply_theme(app: QApplication) -> Dict[str, str]:
    """Fusion style + palette + stylesheet in the system's light/dark scheme."""
    t = tokens()
    app.setStyle("Fusion")

    palette = QPalette()
    palette.setColor(QPalette.ColorRole.Window, QColor(t["window"]))
    palette.setColor(QPalette.ColorRole.WindowText, QColor(t["text"]))
    palette.setColor(QPalette.ColorRole.Base, QColor(t["surface"]))
    palette.setColor(QPalette.ColorRole.AlternateBase, QColor(t["sidebar"]))
    palette.setColor(QPalette.ColorRole.Text, QColor(t["text"]))
    palette.setColor(QPalette.ColorRole.Button, QColor(t["surface"]))
    palette.setColor(QPalette.ColorRole.ButtonText, QColor(t["text"]))
    palette.setColor(QPalette.ColorRole.Highlight, QColor(t["accent"]))
    palette.setColor(QPalette.ColorRole.HighlightedText, QColor(t["on_accent"]))
    palette.setColor(QPalette.ColorRole.ToolTipBase, QColor(t["surface"]))
    palette.setColor(QPalette.ColorRole.ToolTipText, QColor(t["text"]))
    palette.setColor(QPalette.ColorRole.PlaceholderText, QColor(t["text3"]))
    app.setPalette(palette)

    app.setFont(pick_font())
    app.setStyleSheet(build_qss(t))
    return t
