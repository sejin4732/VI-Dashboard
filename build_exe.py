from __future__ import annotations

import subprocess
import sys
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parent
WORKSPACE_ROOT = PROJECT_ROOT / "_WORKSPACE"
BUILD_ROOT = WORKSPACE_ROOT / "build"
DIST_ROOT = BUILD_ROOT / "dist"
WORK_ROOT = BUILD_ROOT / "work"
SPEC_ROOT = BUILD_ROOT / "spec"


def main() -> int:
    BUILD_ROOT.mkdir(parents=True, exist_ok=True)
    DIST_ROOT.mkdir(parents=True, exist_ok=True)
    WORK_ROOT.mkdir(parents=True, exist_ok=True)
    SPEC_ROOT.mkdir(parents=True, exist_ok=True)

    command = [
        sys.executable,
        "-m",
        "PyInstaller",
        "--noconfirm",
        "--clean",
        "--onedir",
        "--windowed",
        "--name",
        "VICT",
        "--distpath",
        str(DIST_ROOT),
        "--workpath",
        str(WORK_ROOT),
        "--specpath",
        str(SPEC_ROOT),
        "--paths",
        str(PROJECT_ROOT),
        "--hidden-import",
        "pandas",
        "--hidden-import",
        "duckdb",
        "--hidden-import",
        "openpyxl",
        "--hidden-import",
        "xlsxwriter",
        "--hidden-import",
        "streamlit",
        "--hidden-import",
        "plotly",
        "--hidden-import",
        "altair",
        "--hidden-import",
        "pyarrow",
        "--hidden-import",
        "tornado",
        "--hidden-import",
        "watchdog",
        "--hidden-import",
        "streamlit.web.cli",
        "--collect-all",
        "duckdb",
        "--collect-all",
        "pandas",
        "--collect-all",
        "openpyxl",
        "--collect-all",
        "xlsxwriter",
        "--collect-all",
        "streamlit",
        "--collect-all",
        "plotly",
        "--collect-all",
        "altair",
        "--collect-all",
        "pyarrow",
        "--collect-all",
        "tornado",
        "--collect-all",
        "watchdog",
        "--copy-metadata",
        "streamlit",
        "--copy-metadata",
        "plotly",
        "--copy-metadata",
        "altair",
        "--copy-metadata",
        "pyarrow",
        str(PROJECT_ROOT / "vi_report_gui.py"),
    ]
    print("[Build] PyInstaller build start", flush=True)
    print(" ".join(f'"{part}"' if " " in part else part for part in command), flush=True)
    subprocess.run(command, check=True, cwd=str(PROJECT_ROOT))
    exe_path = DIST_ROOT / "VICT" / "VICT.exe"
    if not exe_path.exists():
        raise FileNotFoundError(f"Expected EXE not found after build: {exe_path}")
    print(f"[Build] EXE ready: {exe_path}", flush=True)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
