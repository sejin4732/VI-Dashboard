from __future__ import annotations

import json
import os
import sys
from functools import lru_cache
from pathlib import Path


DEFAULT_APP_VERSION = "0.1.0"
CONFIG_FILE_NAME = "config.json"


def is_frozen_app() -> bool:
    return bool(getattr(sys, "frozen", False))


def source_root() -> Path:
    return Path(__file__).resolve().parent


def application_root() -> Path:
    if is_frozen_app():
        return Path(sys.executable).resolve().parent
    return source_root()


def _env_config_path() -> Path | None:
    raw_value = os.environ.get("VICT_CONFIG", "").strip()
    if not raw_value:
        return None
    return Path(raw_value).expanduser().resolve()


def config_path() -> Path | None:
    env_path = _env_config_path()
    if env_path is not None and env_path.exists():
        return env_path

    runtime_candidate = application_root() / CONFIG_FILE_NAME
    if runtime_candidate.exists():
        return runtime_candidate.resolve()

    source_candidate = source_root() / CONFIG_FILE_NAME
    if source_candidate.exists():
        return source_candidate.resolve()
    return None


@lru_cache(maxsize=1)
def runtime_config() -> dict[str, object]:
    resolved_config_path = config_path()
    if resolved_config_path is None or not resolved_config_path.exists():
        return {}
    try:
        return json.loads(resolved_config_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {}


def clear_runtime_path_caches() -> None:
    runtime_config.cache_clear()


def has_runtime_config() -> bool:
    return bool(runtime_config())


def _config_base_dir() -> Path:
    resolved_config_path = config_path()
    if resolved_config_path is None:
        return application_root()
    return resolved_config_path.parent.resolve()


def _resolve_path(value: str | Path, *, base_dir: Path) -> Path:
    resolved = Path(value).expanduser()
    if not resolved.is_absolute():
        resolved = base_dir / resolved
    return resolved.resolve()


def _configured_path(name: str, default_path: Path) -> Path:
    config = runtime_config()
    raw_value = str(config.get(name, "") or "").strip()
    if raw_value:
        return _resolve_path(raw_value, base_dir=_config_base_dir())
    return default_path.resolve()


def data_root() -> Path:
    if is_frozen_app() or has_runtime_config():
        default_path = application_root() / "data"
    else:
        default_path = source_root()
    return _configured_path("data_root", default_path)


def bom_lake_root() -> Path:
    if is_frozen_app() or has_runtime_config():
        default_path = application_root() / "data" / "bom_lake"
    else:
        default_path = source_root() / "bom_lake"
    return _configured_path("bom_lake_root", default_path)


def cache_root() -> Path:
    if is_frozen_app() or has_runtime_config():
        default_path = application_root() / "cache"
    else:
        default_path = source_root() / "workflow" / "dashboard_outputs" / "cache"
    return _configured_path("cache_root", default_path)


def output_root() -> Path:
    if is_frozen_app() or has_runtime_config():
        default_path = application_root() / "outputs"
    else:
        default_path = source_root() / "VI Result"
    return _configured_path("output_root", default_path)


def log_root() -> Path:
    if is_frozen_app() or has_runtime_config():
        default_path = application_root() / "logs"
    else:
        default_path = source_root() / "_WORKSPACE" / "logs_archive"
    return _configured_path("log_root", default_path)


def scenario_templates_root() -> Path:
    configured_root = data_root() / "scenario_templates"
    if configured_root.exists() or is_frozen_app() or has_runtime_config():
        return configured_root.resolve()
    return source_root()


def sample_files_root() -> Path:
    configured_root = data_root() / "sample_files"
    if configured_root.exists() or is_frozen_app() or has_runtime_config():
        return configured_root.resolve()
    return source_root()


def dashboard_output_root() -> Path:
    if is_frozen_app() or has_runtime_config():
        return output_root()
    return (source_root() / "workflow" / "dashboard_outputs").resolve()


def report_cache_root() -> Path:
    if is_frozen_app() or has_runtime_config():
        return (cache_root() / "report_cache").resolve()
    return cache_root()


def recommendation_cache_root() -> Path:
    return report_cache_root()


def bom_pair_cache_root() -> Path:
    if is_frozen_app() or has_runtime_config():
        return (cache_root() / "bom_pair_cache").resolve()
    return cache_root()


def duckdb_temp_root() -> Path:
    if is_frozen_app() or has_runtime_config():
        return (cache_root() / ".duckdb_tmp").resolve()
    return (dashboard_output_root() / ".duckdb_tmp").resolve()


def portable_path_token(path: Path | None, *, preferred_root: Path | None = None) -> str:
    if path is None:
        return ""
    resolved_path = Path(path).expanduser().resolve()
    for candidate_root in [preferred_root, application_root()]:
        if candidate_root is None:
            continue
        try:
            relative_path = resolved_path.relative_to(Path(candidate_root).expanduser().resolve())
            return relative_path.as_posix()
        except ValueError:
            continue
    return str(resolved_path)

