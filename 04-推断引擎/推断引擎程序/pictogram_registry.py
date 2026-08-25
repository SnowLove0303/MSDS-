"""Single source of truth for GHS pictogram codes, labels and PNG assets."""

from __future__ import annotations

import json
import os
import re
from pathlib import Path
from typing import Iterable


ENGINE_DIR = Path(__file__).resolve().parent
LIBRARY_PATH = ENGINE_DIR / "s2_result_library.json"
_PROJECT_ROOT = Path(os.environ.get(
    "MSDS_ROOT", str(Path(__file__).resolve().parents[2])))
ASSET_DIR = _PROJECT_ROOT / "04-推断引擎" / "判断skill" / "templates" / "pictograms"
_CODE_RE = re.compile(r"GHS0[1-9]", re.IGNORECASE)


def load_library() -> dict:
    return json.loads(LIBRARY_PATH.read_text(encoding="utf-8"))


def pictogram_definitions() -> dict[str, dict]:
    return load_library().get("pictograms", {})


def normalize_code(value: object) -> str:
    match = _CODE_RE.search(str(value or ""))
    return match.group(0).upper() if match else ""


def extract_codes(value: object) -> list[str]:
    """Extract ordered, de-duplicated GHS codes from a value."""
    found: list[str] = []
    for match in _CODE_RE.finditer(str(value or "")):
        code = match.group(0).upper()
        if code not in found:
            found.append(code)
    return found


def ordered_codes(codes: Iterable[object]) -> list[str]:
    valid = {code for code in (normalize_code(x) for x in codes) if code}
    return sorted(valid, key=lambda code: int(code[-2:]))


def asset_path(code: object) -> Path | None:
    normalized = normalize_code(code)
    definition = pictogram_definitions().get(normalized)
    if not definition:
        return None
    path = ASSET_DIR / str(definition.get("asset", ""))
    return path if path.is_file() else None


def asset_bytes(code: object) -> tuple[bytes, str] | None:
    path = asset_path(code)
    if path is None:
        return None
    return path.read_bytes(), path.suffix.lstrip(".").lower() or "png"


def label(code: object) -> str:
    normalized = normalize_code(code)
    return str(pictogram_definitions().get(normalized, {}).get("short", normalized))


def display_value(codes: Iterable[object]) -> str:
    ordered = ordered_codes(codes)
    return "；".join(f"{code} {label(code)}" for code in ordered)


def display_items(value: object) -> list[tuple[str, str, bytes, str]]:
    """Return (code, label, bytes, extension) for GUI/Word rendering."""
    items: list[tuple[str, str, bytes, str]] = []
    for code in ordered_codes(extract_codes(value)):
        blob = asset_bytes(code)
        if blob is not None:
            items.append((code, label(code), blob[0], blob[1]))
    return items


__all__ = [
    "ASSET_DIR", "asset_bytes", "asset_path", "display_items", "display_value",
    "extract_codes", "label", "normalize_code", "ordered_codes",
]
