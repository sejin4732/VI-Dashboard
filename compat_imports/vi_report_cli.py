"""Compatibility wrapper for the root-level VI report CLI."""

from __future__ import annotations

import importlib.util
from pathlib import Path
import sys
from types import ModuleType


def _load_root_module() -> ModuleType:
    module_path = Path(__file__).resolve().parent.parent / "vi_report_cli.py"
    spec = importlib.util.spec_from_file_location("_compat_root_vi_report_cli", module_path)
    if spec is None or spec.loader is None:
        raise ImportError(f"Could not load root vi_report_cli from {module_path}")
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


_root_module = _load_root_module()
__all__ = [name for name in dir(_root_module) if not name.startswith("_")]
globals().update({name: getattr(_root_module, name) for name in __all__})
