from __future__ import annotations

import importlib.util
import sys
import types
from pathlib import Path

import streamlit as st


def _ensure_package(package_name: str, package_path: Path) -> None:
    existing = sys.modules.get(package_name)
    if existing is not None and hasattr(existing, "__path__"):
        existing_paths = [str(path) for path in getattr(existing, "__path__", [])]
        if str(package_path) not in existing_paths:
            existing.__path__ = [str(package_path), *existing_paths]
        return

    package = types.ModuleType(package_name)
    package.__path__ = [str(package_path)]
    package.__package__ = package_name
    sys.modules[package_name] = package


def _load_module_from_path(module_name: str, module_path: Path):
    spec = importlib.util.spec_from_file_location(module_name, module_path)
    if spec is None or spec.loader is None:
        raise ImportError(f"Cannot load module {module_name} from {module_path}")
    module = importlib.util.module_from_spec(spec)
    sys.modules[module_name] = module
    spec.loader.exec_module(module)
    return module


def _load_local_dashboard_page():
    app_root = Path(__file__).resolve().parent
    dashboard_root = app_root / "dashboard"
    core_root = dashboard_root / "core"
    data_path = core_root / "vi_dashboard_data.py"
    page_path = core_root / "vi_dashboard_page.py"

    if not data_path.exists():
        raise FileNotFoundError(f"Dashboard data module not found: {data_path}")
    if not page_path.exists():
        raise FileNotFoundError(f"Dashboard page module not found: {page_path}")

    _ensure_package("dashboard", dashboard_root)
    _ensure_package("dashboard.core", core_root)

    data_module = _load_module_from_path("dashboard.core.vi_dashboard_data", data_path)
    page_module = _load_module_from_path("dashboard.core.vi_dashboard_page", page_path)

    print(
        "[DashboardLoader] loaded local modules "
        f"| data={Path(getattr(data_module, '__file__', data_path)).resolve()} "
        f"| page={Path(getattr(page_module, '__file__', page_path)).resolve()}",
        flush=True,
    )
    return page_module


dashboard_page = _load_local_dashboard_page()
parse_dashboard_args = dashboard_page.parse_dashboard_args
render_dashboard = dashboard_page.render_dashboard


st.set_page_config(
    page_title="VI Dashboard",
    layout="wide",
)


def main() -> None:
    args = parse_dashboard_args()
    render_dashboard(args.report_file)


if __name__ == "__main__":
    main()
