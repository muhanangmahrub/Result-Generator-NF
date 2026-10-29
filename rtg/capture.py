"""Screen capture: full virtual desktop, one monitor, or a rectangle.

Pure capture logic, no Qt, so it can be unit tested without a GUI. The region
selector in the GUI never grabs the screen itself: it takes one full capture
first and then crops it, which is how Snipping Tool behaves and keeps the
selection overlay out of the picture.

Physical vs logical pixels: mss works in *physical* pixels while Qt reports
geometry in *logical* pixels. :func:`crop` therefore takes the scale factor the
caller measured, instead of guessing it here.
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import List, Optional, Tuple

from .errors import ResultGeneratorError

#: Capture modes stored in project.json.
MODE_FULLSCREEN = "fullscreen"
MODE_MONITOR = "monitor"
MODE_REGION = "region"
CAPTURE_MODES = (MODE_FULLSCREEN, MODE_MONITOR, MODE_REGION)

MODE_LABELS = {
    MODE_FULLSCREEN: "Seluruh layar",
    MODE_MONITOR: "Monitor tertentu",
    MODE_REGION: "Pilih area (drag)",
}


class CaptureError(ResultGeneratorError):
    """Raised when the screen cannot be captured."""


@dataclass(frozen=True)
class Monitor:
    """One monitor as reported by mss, in physical pixels."""

    index: int               # 1-based, matching mss; 0 means the whole desktop
    left: int
    top: int
    width: int
    height: int

    @property
    def label(self) -> str:
        if self.index == 0:
            return f"Seluruh layar ({self.width}x{self.height})"
        return f"Monitor {self.index} - {self.width}x{self.height} @ ({self.left},{self.top})"

    @property
    def box(self) -> Tuple[int, int, int, int]:
        return (self.left, self.top, self.width, self.height)

    def to_dict(self) -> dict:
        return {
            "index": self.index, "left": self.left, "top": self.top,
            "width": self.width, "height": self.height,
        }


def _screenshotter():
    """Return the mss class to instantiate.

    mss >= 10 renamed ``mss.mss`` to ``mss.MSS`` and deprecated the old name;
    accept whichever the installed version offers.
    """
    try:
        import mss  # noqa: PLC0415 - imported lazily so the parser works without it
    except ImportError as exc:  # pragma: no cover - depends on the environment
        raise CaptureError(
            "Modul 'mss' belum terpasang. Jalankan: pip install mss Pillow",
            detail=str(exc),
        ) from exc
    return getattr(mss, "MSS", None) or mss.mss


def _pil_image():
    try:
        from PIL import Image  # noqa: PLC0415
    except ImportError as exc:  # pragma: no cover
        raise CaptureError(
            "Modul 'Pillow' belum terpasang. Jalankan: pip install Pillow",
            detail=str(exc),
        ) from exc
    return Image


def list_monitors() -> List[Monitor]:
    """Every monitor, with index 0 as the whole virtual desktop."""
    screenshotter = _screenshotter()
    try:
        with screenshotter() as sct:
            return [
                Monitor(
                    index=index,
                    left=int(mon["left"]), top=int(mon["top"]),
                    width=int(mon["width"]), height=int(mon["height"]),
                )
                for index, mon in enumerate(sct.monitors)
            ]
    except Exception as exc:  # pragma: no cover - driver/display dependent
        raise CaptureError(
            f"Gagal membaca daftar monitor: {exc}", detail=str(exc)
        ) from exc


def grab(monitor_index: int = 0):
    """Capture a monitor and return a Pillow ``Image`` in RGB.

    ``monitor_index`` 0 is the whole virtual desktop (every monitor combined).
    """
    screenshotter = _screenshotter()
    Image = _pil_image()
    try:
        with screenshotter() as sct:
            monitors = sct.monitors
            if not 0 <= monitor_index < len(monitors):
                raise CaptureError(
                    f"Monitor {monitor_index} tidak ada. Tersedia: "
                    f"0..{len(monitors) - 1}."
                )
            shot = sct.grab(monitors[monitor_index])
    except CaptureError:
        raise
    except Exception as exc:  # pragma: no cover - driver/display dependent
        raise CaptureError(f"Gagal mengambil screenshot: {exc}", detail=str(exc)) from exc
    return Image.frombytes("RGB", shot.size, shot.bgra, "raw", "BGRX")


def crop(image, box: Tuple[int, int, int, int], *, scale: float = 1.0):
    """Crop ``(left, top, width, height)`` out of a captured image.

    ``box`` is in the caller's coordinate space relative to the image origin;
    ``scale`` converts it to the image's physical pixels (2.0 on a 200% display).
    """
    left, top, width, height = box
    if width <= 0 or height <= 0:
        raise CaptureError("Area yang dipilih kosong.")
    right = int(round((left + width) * scale))
    bottom = int(round((top + height) * scale))
    left = int(round(left * scale))
    top = int(round(top * scale))
    left = max(0, min(left, image.width))
    top = max(0, min(top, image.height))
    right = max(left + 1, min(right, image.width))
    bottom = max(top + 1, min(bottom, image.height))
    return image.crop((left, top, right, bottom))


def save_png(image, path: Path) -> Path:
    """Write the image as PNG, creating the parent folder if needed."""
    path = Path(path)
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        image.save(str(path), format="PNG", optimize=True)
    except OSError as exc:
        raise CaptureError(
            f"Gagal menyimpan screenshot ke {path}: {exc}", detail=str(exc)
        ) from exc
    return path


def capture_to_file(
    path: Path,
    *,
    mode: str = MODE_FULLSCREEN,
    monitor_index: int = 0,
    box: Optional[Tuple[int, int, int, int]] = None,
    scale: float = 1.0,
) -> Path:
    """Capture according to ``mode`` and write a PNG.

    ``MODE_REGION`` needs ``box``; the GUI normally captures once and crops
    itself, this path exists for scripted use and tests.
    """
    if mode not in CAPTURE_MODES:
        raise CaptureError(
            f"Mode capture tidak dikenali: {mode!r}. "
            f"Pilihan: {', '.join(CAPTURE_MODES)}."
        )
    if mode == MODE_MONITOR:
        image = grab(monitor_index)
    elif mode == MODE_REGION:
        if box is None:
            raise CaptureError("Mode 'region' butuh area (box) yang dipilih.")
        image = crop(grab(0), box, scale=scale)
    else:
        image = grab(0)
    return save_png(image, path)
