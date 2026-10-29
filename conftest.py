"""Make the project root importable and keep Qt head-less during tests."""

import os
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

# Must be set before PySide6 creates a QApplication, otherwise the GUI tests
# would try to open real windows.
os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
