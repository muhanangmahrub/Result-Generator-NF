"""Unit tests for screen capture.

The grabbing tests need a real display; they are skipped when mss cannot see
one. The cropping and saving tests run anywhere.
"""

from __future__ import annotations

import pytest

pytest.importorskip("mss")
pytest.importorskip("PIL")

from PIL import Image                                            # noqa: E402

from rtg.capture import (                                        # noqa: E402
    CAPTURE_MODES,
    MODE_FULLSCREEN,
    MODE_MONITOR,
    MODE_REGION,
    CaptureError,
    capture_to_file,
    crop,
    grab,
    list_monitors,
    save_png,
)


@pytest.fixture(scope="module")
def monitors():
    try:
        found = list_monitors()
    except CaptureError as exc:                                   # pragma: no cover
        pytest.skip(f"tidak ada display: {exc}")
    if len(found) < 2:                                            # pragma: no cover
        pytest.skip("tidak ada monitor yang terdeteksi")
    return found


@pytest.fixture
def picture():
    return Image.new("RGB", (400, 200), (10, 20, 30))


# --------------------------------------------------------------------------- #
# Monitors and grabbing
# --------------------------------------------------------------------------- #

def test_list_monitors_starts_with_the_whole_desktop(monitors):
    assert monitors[0].index == 0
    assert monitors[0].width > 0 and monitors[0].height > 0
    assert "Seluruh layar" in monitors[0].label
    assert monitors[1].index == 1
    assert monitors[1].label.startswith("Monitor 1")
    assert monitors[0].box == (
        monitors[0].left, monitors[0].top, monitors[0].width, monitors[0].height
    )


def test_grab_whole_desktop(monitors):
    image = grab(0)
    assert image.mode == "RGB"
    assert (image.width, image.height) == (monitors[0].width, monitors[0].height)


def test_grab_single_monitor(monitors):
    image = grab(1)
    assert (image.width, image.height) == (monitors[1].width, monitors[1].height)


def test_grab_unknown_monitor(monitors):
    with pytest.raises(CaptureError, match="tidak ada"):
        grab(99)


# --------------------------------------------------------------------------- #
# Cropping
# --------------------------------------------------------------------------- #

def test_crop_basic(picture):
    out = crop(picture, (10, 20, 100, 50))
    assert (out.width, out.height) == (100, 50)


def test_crop_applies_the_dpi_scale(picture):
    # A 200% display: the user drags 50x25 logical pixels, which is 100x50 real ones.
    out = crop(picture, (5, 10, 50, 25), scale=2.0)
    assert (out.width, out.height) == (100, 50)


def test_crop_is_clamped_to_the_image(picture):
    out = crop(picture, (350, 150, 500, 500))
    assert out.width == 50 and out.height == 50
    assert crop(picture, (-50, -50, 100, 100)).width <= picture.width


@pytest.mark.parametrize("box", [(0, 0, 0, 10), (0, 0, 10, 0), (0, 0, -5, -5)])
def test_crop_rejects_an_empty_area(picture, box):
    with pytest.raises(CaptureError, match="kosong"):
        crop(picture, box)


# --------------------------------------------------------------------------- #
# Saving
# --------------------------------------------------------------------------- #

def test_save_png_creates_parent_folders(picture, tmp_path):
    target = tmp_path / "a" / "b" / "shot.png"
    assert save_png(picture, target) == target
    assert target.is_file()

    with Image.open(target) as reopened:
        assert reopened.format == "PNG"
        assert reopened.size == (400, 200)


def test_save_png_on_an_unwritable_path(picture, tmp_path):
    blocker = tmp_path / "blocker"
    blocker.write_text("i am a file", encoding="utf-8")
    with pytest.raises(CaptureError, match="Gagal menyimpan"):
        save_png(picture, blocker / "nested" / "shot.png")


# --------------------------------------------------------------------------- #
# capture_to_file
# --------------------------------------------------------------------------- #

def test_capture_to_file_fullscreen(monitors, tmp_path):
    path = capture_to_file(tmp_path / "full.png", mode=MODE_FULLSCREEN)
    with Image.open(path) as image:
        assert image.size == (monitors[0].width, monitors[0].height)


def test_capture_to_file_monitor(monitors, tmp_path):
    path = capture_to_file(tmp_path / "mon.png", mode=MODE_MONITOR, monitor_index=1)
    with Image.open(path) as image:
        assert image.size == (monitors[1].width, monitors[1].height)


def test_capture_to_file_region(monitors, tmp_path):
    path = capture_to_file(
        tmp_path / "region.png", mode=MODE_REGION, box=(0, 0, 120, 80)
    )
    with Image.open(path) as image:
        assert image.size == (120, 80)


def test_capture_to_file_region_without_a_box(tmp_path):
    with pytest.raises(CaptureError, match="butuh area"):
        capture_to_file(tmp_path / "x.png", mode=MODE_REGION)


def test_capture_to_file_unknown_mode(tmp_path):
    with pytest.raises(CaptureError, match="tidak dikenali"):
        capture_to_file(tmp_path / "x.png", mode="telepati")


def test_capture_modes_are_the_three_promised_ones():
    assert set(CAPTURE_MODES) == {MODE_FULLSCREEN, MODE_MONITOR, MODE_REGION}
