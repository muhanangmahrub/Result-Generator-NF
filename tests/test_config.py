"""Unit tests for the alias configuration loader."""

from __future__ import annotations

import json

import pytest

from rtg.config import PACKAGE_DEFAULT_CONFIG, load_config, write_default_config


def test_defaults_load():
    config = load_config()
    assert config.required_fields == ("test_case_id", "test_case", "test_step")
    assert "test_case_id" in config.aliases
    assert ("test", "case", "id") in config.aliases["test_case_id"]
    assert config.header_scan_rows == 30
    assert PACKAGE_DEFAULT_CONFIG in config.sources


def test_aliases_are_normalised_and_sorted_longest_first():
    config = load_config()
    aliases = config.aliases["test_case_id"]
    lengths = [len(item) for item in aliases]
    assert lengths == sorted(lengths, reverse=True)
    # 'Pass/Fail (P/F)' in the config is stored normalised.
    assert ("pass", "fail", "pf") in config.aliases["pass_fail"]


def test_fill_down_groups():
    config = load_config()
    keys = [group.key for group in config.fill_down_groups]
    assert keys == ["scenario_id", "test_case_id"]
    assert "test_case" in config.fill_down_fields
    assert "test_step" not in config.fill_down_fields


def test_user_override_extends_defaults(tmp_path):
    override = tmp_path / "column_aliases.json"
    override.write_text(
        json.dumps({"aliases": {"test_case_id": ["Kode TC"]}, "max_blank_rows": 3}),
        encoding="utf-8",
    )
    config = load_config(override)
    assert ("kode", "tc") in config.aliases["test_case_id"]
    # defaults survive the merge
    assert ("test", "case", "id") in config.aliases["test_case_id"]
    assert config.max_blank_rows == 3


def test_user_override_discovered_in_project_folder(tmp_path):
    config_dir = tmp_path / "config"
    config_dir.mkdir()
    (config_dir / "column_aliases.json").write_text(
        json.dumps({"aliases": {"notes": ["Komentar"]}}), encoding="utf-8"
    )
    config = load_config(search_from=tmp_path)
    assert ("komentar",) in config.aliases["notes"]


def test_missing_explicit_config_raises(tmp_path):
    with pytest.raises(FileNotFoundError):
        load_config(tmp_path / "nope.json")


def test_write_default_config(tmp_path):
    written = write_default_config(tmp_path / "config")
    # A folder path that does not exist yet still gets the default file name.
    assert written == tmp_path / "config" / "column_aliases.json"
    assert written.is_file()
    assert json.loads(written.read_text(encoding="utf-8"))["aliases"]
    with pytest.raises(FileExistsError):
        write_default_config(written)
    assert write_default_config(written, overwrite=True) == written
