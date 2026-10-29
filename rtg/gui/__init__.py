"""PySide6 front-end for Result Test Generator."""

from __future__ import annotations

import sys
from pathlib import Path
from typing import List, Optional

__all__ = ["run", "MainWindow"]


def run(argv: Optional[List[str]] = None) -> int:
    """Start the GUI. An optional argument is a project folder to open."""
    from PySide6.QtWidgets import QApplication, QMessageBox

    from ..project import Project, ProjectError
    from .main_window import MainWindow

    argv = list(sys.argv if argv is None else argv)
    app = QApplication(argv)
    app.setApplicationName("Result Test Generator")
    app.setOrganizationName("ResultTestGenerator")

    if "--selftest" in argv[1:]:
        return _selftest(app)

    project = None
    error = None
    if len(argv) > 1 and argv[1]:
        try:
            project = Project.open(Path(argv[1]))
        except ProjectError as exc:
            error = exc

    from .theme import apply_theme

    apply_theme(app)
    window = MainWindow(project)

    def retheme(*_args) -> None:
        apply_theme(app)
        window.apply_icons()

    try:
        app.styleHints().colorSchemeChanged.connect(retheme)
    except AttributeError:       # Qt < 6.5 has no colour-scheme signal
        pass
    window.show()
    if project is not None:
        window.announce_scenario_changes()
    if error is not None:
        QMessageBox.warning(window, "Gagal membuka project", error.message)
    return app.exec()


def _selftest(app) -> int:
    """Smoke test for a packaged build: theme, window and bundled icons load.

    Run ``"Result Test Generator.exe" --selftest`` (with QT_QPA_PLATFORM=offscreen
    on a machine without a display); exit code 0 means the build is sane.
    """
    from .icons import icon
    from .main_window import MainWindow
    from .theme import apply_theme

    apply_theme(app)
    window = MainWindow(None)
    window.show()
    app.processEvents()
    ok = not icon("save").isNull()          # fails when qtawesome fonts were not bundled
    window.close()
    return 0 if ok else 1


def __getattr__(name: str):
    if name == "MainWindow":
        from .main_window import MainWindow

        return MainWindow
    raise AttributeError(name)
