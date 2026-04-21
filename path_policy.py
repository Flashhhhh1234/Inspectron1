"""Centralized path policy for shared project storage.

All persisted paths are stored relative to a configured UNC base path.
Absolute paths are reconstructed at runtime.
"""

import json
import os
from typing import Optional


DEFAULT_BASE_PATH = r"\\innsk01-fs01\iCenter\Operations\01_Projects"
_BASE_PATH_CACHE: Optional[str] = None


def _normalize_path(value: str) -> str:
    return os.path.normpath(str(value).strip())


def _config_paths():
    current_dir = os.path.dirname(os.path.abspath(__file__))
    return (
        os.path.join(current_dir, "assets", "postgres.json"),
        os.path.join(os.path.dirname(current_dir), "assets", "postgres.json"),
    )


def _load_base_path_from_config() -> Optional[str]:
    for config_path in _config_paths():
        if not os.path.exists(config_path):
            continue

        try:
            with open(config_path, "r", encoding="utf-8") as handle:
                raw = json.load(handle)
        except Exception:
            continue

        if not isinstance(raw, dict):
            continue

        postgres_block = raw.get("postgres", raw)
        if not isinstance(postgres_block, dict):
            continue

        base_path = postgres_block.get("base_path") or raw.get("base_path")
        if base_path:
            return _normalize_path(base_path)

    return None


def get_base_path(force_refresh: bool = False) -> str:
    global _BASE_PATH_CACHE
    if _BASE_PATH_CACHE is None or force_refresh:
        configured = _load_base_path_from_config()
        _BASE_PATH_CACHE = configured or _normalize_path(DEFAULT_BASE_PATH)
    return _BASE_PATH_CACHE


def is_within_base_path(path: Optional[str]) -> bool:
    if not path:
        return False

    candidate = _normalize_path(path)
    if not os.path.isabs(candidate):
        candidate = _normalize_path(os.path.join(get_base_path(), candidate))

    base = _normalize_path(get_base_path())
    try:
        return os.path.commonpath([os.path.normcase(candidate), os.path.normcase(base)]) == os.path.normcase(base)
    except ValueError:
        return False


def to_relative_path(path: Optional[str]) -> Optional[str]:
    if path is None:
        return None

    raw = str(path).strip()
    if not raw:
        return None

    normalized = _normalize_path(raw)
    if not os.path.isabs(normalized):
        return normalized

    if not is_within_base_path(normalized):
        raise ValueError(f"Path must be under base_path '{get_base_path()}': {normalized}")

    relative = os.path.relpath(normalized, get_base_path())
    return "." if relative in ("", ".") else _normalize_path(relative)


def to_absolute_path(path: Optional[str]) -> Optional[str]:
    if path is None:
        return None

    raw = str(path).strip()
    if not raw:
        return ""

    normalized = _normalize_path(raw)
    if os.path.isabs(normalized):
        return normalized if is_within_base_path(normalized) else None

    return _normalize_path(os.path.join(get_base_path(), normalized))


def to_relative_storage_location(storage_location: Optional[str]) -> str:
    if storage_location is None:
        return "."

    raw = str(storage_location).strip()
    if not raw:
        return "."

    normalized = _normalize_path(raw)
    if not os.path.isabs(normalized):
        return "." if normalized in ("", ".") else normalized

    if not is_within_base_path(normalized):
        raise ValueError(
            f"Storage location must be under base_path '{get_base_path()}': {normalized}"
        )

    relative = os.path.relpath(normalized, get_base_path())
    return "." if relative in ("", ".") else _normalize_path(relative)


def resolve_storage_location(stored_value: Optional[str]) -> str:
    if stored_value is None:
        return get_base_path()

    raw = str(stored_value).strip()
    if not raw:
        return get_base_path()

    normalized = _normalize_path(raw)

    if os.path.isabs(normalized):
        if is_within_base_path(normalized):
            return normalized
        return get_base_path()

    if normalized in (".", ""):
        return get_base_path()

    return _normalize_path(os.path.join(get_base_path(), normalized))
