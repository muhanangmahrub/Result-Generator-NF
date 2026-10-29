"""Loading and merging of the user-editable header alias configuration."""

from __future__ import annotations

import json
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple

from .normalize import normalize_header, tokens_of

PACKAGE_DEFAULT_CONFIG = Path(__file__).with_name("data") / "column_aliases.json"

#: Relative locations searched for a user override, in order.
USER_CONFIG_CANDIDATES: Tuple[str, ...] = (
    "config/column_aliases.json",
    "column_aliases.json",
)


@dataclass(frozen=True)
class FillDownGroup:
    """A set of columns that are only filled on the first row of a group.

    When ``key`` changes to a new non-empty value, the carried values of this
    group *and of every group after it* are cleared, so a new Test Case ID can
    never inherit the previous test case's name.
    """

    key: str
    fields: Tuple[str, ...]


@dataclass
class FieldConfig:
    """Normalised alias tables plus parser tuning knobs."""

    aliases: Dict[str, Tuple[Tuple[str, ...], ...]] = field(default_factory=dict)
    metadata_aliases: Dict[str, Tuple[Tuple[str, ...], ...]] = field(default_factory=dict)
    required_fields: Tuple[str, ...] = ("test_case_id", "test_case", "test_step")
    fill_down_groups: Tuple[FillDownGroup, ...] = ()
    header_scan_rows: int = 30
    max_blank_rows: int = 10
    metadata_scan_columns: int = 10
    sources: Tuple[Path, ...] = ()

    @property
    def known_fields(self) -> Tuple[str, ...]:
        return tuple(self.aliases.keys())

    @property
    def fill_down_fields(self) -> Tuple[str, ...]:
        out: List[str] = []
        for group in self.fill_down_groups:
            for name in group.fields:
                if name not in out:
                    out.append(name)
        return tuple(out)


def _normalize_alias_table(raw: Dict[str, Sequence[str]]) -> Dict[str, Tuple[Tuple[str, ...], ...]]:
    """Normalise every alias and sort longest-first for stable matching."""
    table: Dict[str, Tuple[Tuple[str, ...], ...]] = {}
    for field_name, aliases in raw.items():
        seen: List[Tuple[str, ...]] = []
        for alias in aliases:
            toks = tokens_of(normalize_header(alias))
            if toks and toks not in seen:
                seen.append(toks)
        # Longest alias first: "test case id" must be tried before "test case".
        seen.sort(key=lambda toks: (-len(toks), toks))
        table[field_name] = tuple(seen)
    return table


def _merge_raw(base: Dict[str, Any], override: Dict[str, Any]) -> Dict[str, Any]:
    """Merge a user config over the defaults.

    Alias lists are *extended* (user entries first) rather than replaced, so a
    user can add a project-specific header without losing the defaults. To drop
    a default alias entirely, remove it from the packaged file.
    """
    merged = dict(base)
    for key, value in override.items():
        if key.startswith("_"):
            continue
        if key in ("aliases", "metadata_aliases") and isinstance(value, dict):
            table = {k: list(v) for k, v in base.get(key, {}).items()}
            for field_name, aliases in value.items():
                existing = table.get(field_name, [])
                table[field_name] = list(aliases) + [a for a in existing if a not in aliases]
            merged[key] = table
        else:
            merged[key] = value
    return merged


def _read_json(path: Path) -> Dict[str, Any]:
    with path.open("r", encoding="utf-8-sig") as handle:
        data = json.load(handle)
    if not isinstance(data, dict):
        raise ValueError(f"Isi config harus berupa object JSON: {path}")
    return data


def load_config(
    path: Optional[Path] = None,
    *,
    search_from: Optional[Path] = None,
    use_defaults: bool = True,
) -> FieldConfig:
    """Load the alias configuration.

    Args:
        path: explicit config file; when given it is the only override read.
        search_from: directory whose ``config/column_aliases.json`` is used as
            an override when ``path`` is not given (typically the project folder).
        use_defaults: start from the packaged defaults (almost always True).
    """
    raw: Dict[str, Any] = {}
    sources: List[Path] = []

    if use_defaults:
        raw = _read_json(PACKAGE_DEFAULT_CONFIG)
        sources.append(PACKAGE_DEFAULT_CONFIG)

    override_path: Optional[Path] = None
    if path is not None:
        override_path = Path(path)
        if not override_path.is_file():
            raise FileNotFoundError(f"File config tidak ditemukan: {override_path}")
    elif search_from is not None:
        for candidate in USER_CONFIG_CANDIDATES:
            candidate_path = Path(search_from) / candidate
            if candidate_path.is_file():
                override_path = candidate_path
                break

    if override_path is not None:
        raw = _merge_raw(raw, _read_json(override_path))
        sources.append(override_path)

    groups = tuple(
        FillDownGroup(key=item.get("key", ""), fields=tuple(item.get("fields", ())))
        for item in raw.get("fill_down_groups", ())
    )

    return FieldConfig(
        aliases=_normalize_alias_table(raw.get("aliases", {})),
        metadata_aliases=_normalize_alias_table(raw.get("metadata_aliases", {})),
        required_fields=tuple(raw.get("required_fields", ("test_case_id", "test_case", "test_step"))),
        fill_down_groups=groups,
        header_scan_rows=int(raw.get("header_scan_rows", 30)),
        max_blank_rows=int(raw.get("max_blank_rows", 10)),
        metadata_scan_columns=int(raw.get("metadata_scan_columns", 10)),
        sources=tuple(sources),
    )


def write_default_config(destination: Path, *, overwrite: bool = False) -> Path:
    """Copy the packaged defaults to ``destination`` for the user to edit."""
    destination = Path(destination)
    # A path without a .json suffix is treated as the target folder, whether or
    # not it exists yet.
    if destination.is_dir() or destination.suffix.lower() != ".json":
        destination = destination / "column_aliases.json"
    if destination.exists() and not overwrite:
        raise FileExistsError(f"File sudah ada: {destination}")
    destination.parent.mkdir(parents=True, exist_ok=True)
    destination.write_text(
        PACKAGE_DEFAULT_CONFIG.read_text(encoding="utf-8"), encoding="utf-8"
    )
    return destination
