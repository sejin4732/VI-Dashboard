from __future__ import annotations

import json
import os
import shutil
import sys
import time
from contextlib import contextmanager
from pathlib import Path

from runtime_paths import clear_runtime_path_caches
from workflow import vi_calculation_engine as vi_engine


PROJECT_ROOT = Path(__file__).resolve().parent
WORKSPACE_ROOT = PROJECT_ROOT / "_WORKSPACE"
DEPLOY_ROOT = PROJECT_ROOT / "_DEPLOY_VICT"

BUILD_DIST_DIR = WORKSPACE_ROOT / "build" / "dist" / "VICT"
OUTPUTS_ARCHIVE_ROOT = WORKSPACE_ROOT / "outputs_archive"
LOGS_ARCHIVE_ROOT = WORKSPACE_ROOT / "logs_archive"
ARCHIVE_ROOT = WORKSPACE_ROOT / "archive"

CONFIG_PAYLOAD = {
    "app_version": "0.1.0",
    "data_root": "./data",
    "bom_lake_root": "./data/bom_lake",
    "cache_root": "./cache",
    "output_root": "./outputs",
    "log_root": "./logs",
}

RUNTIME_FILES = [
    "bom_lake.py",
    "historical_bom_lake.py",
    "runtime_paths.py",
    "streamlit_vi_dashboard.py",
    "vi_calculation_engine.py",
    "vi_report_cli.py",
    "vi_report_gui.py",
]
RUNTIME_DIRS = [
    "compat_imports",
    "dashboard",
    "workflow",
]
SAMPLE_FILE_GLOBS = [
    "sample_*.xlsx",
    "sample_*.parquet",
    "generated_*.xlsx",
]
SCENARIO_TEMPLATE_GLOBS = [
    "VI_Scenario_Template*.xlsx",
    "*Scenario*.xlsx",
]


def log(message: str) -> None:
    print(message, flush=True)


def ensure_workspace_structure() -> None:
    for path in [
        WORKSPACE_ROOT,
        WORKSPACE_ROOT / "build",
        OUTPUTS_ARCHIVE_ROOT,
        LOGS_ARCHIVE_ROOT,
        ARCHIVE_ROOT,
    ]:
        path.mkdir(parents=True, exist_ok=True)


def _timestamp() -> str:
    return time.strftime("%Y%m%d_%H%M%S")


def archive_existing_deploy_folder() -> None:
    if not DEPLOY_ROOT.exists():
        return
    archived_path = ARCHIVE_ROOT / f"_DEPLOY_VICT_{_timestamp()}"
    shutil.move(str(DEPLOY_ROOT), str(archived_path))
    log(f"[Deploy] Archived existing deploy folder to: {archived_path}")


def snapshot_existing_generated_outputs() -> None:
    source_dashboard_outputs = PROJECT_ROOT / "workflow" / "dashboard_outputs"
    if source_dashboard_outputs.exists():
        target = OUTPUTS_ARCHIVE_ROOT / f"dashboard_outputs_{_timestamp()}"
        shutil.copytree(
            source_dashboard_outputs,
            target,
            dirs_exist_ok=True,
            ignore=shutil.ignore_patterns("__pycache__", "*.pyc"),
        )
        log(f"[Workspace] Archived current dashboard outputs to: {target}")


def _copy_file(src: Path, dst: Path) -> None:
    dst.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(src, dst)


def _copy_tree(src: Path, dst: Path, *, ignore_patterns: tuple[str, ...] = ()) -> None:
    shutil.copytree(
        src,
        dst,
        dirs_exist_ok=True,
        ignore=shutil.ignore_patterns(*ignore_patterns) if ignore_patterns else None,
    )


def create_deploy_structure() -> None:
    for path in [
        DEPLOY_ROOT,
        DEPLOY_ROOT / "data" / "bom_lake" / "historical",
        DEPLOY_ROOT / "data" / "bom_lake" / "current_cache",
        DEPLOY_ROOT / "data" / "scenario_templates",
        DEPLOY_ROOT / "data" / "sample_files",
        DEPLOY_ROOT / "cache" / "bom_pair_cache",
        DEPLOY_ROOT / "cache" / "report_cache",
        DEPLOY_ROOT / "outputs",
        DEPLOY_ROOT / "logs",
    ]:
        path.mkdir(parents=True, exist_ok=True)


def copy_build_artifacts() -> None:
    if not BUILD_DIST_DIR.exists():
        raise FileNotFoundError(
            f"Build output not found: {BUILD_DIST_DIR}\nRun build_exe.py first."
        )
    _copy_tree(BUILD_DIST_DIR, DEPLOY_ROOT, ignore_patterns=("__pycache__", "*.pyc"))
    exe_path = DEPLOY_ROOT / "VICT.exe"
    if not exe_path.exists():
        raise FileNotFoundError(f"Expected EXE not copied: {exe_path}")
    log(f"[Deploy] Copied build artifacts: {exe_path}")


def copy_runtime_sources() -> None:
    for file_name in RUNTIME_FILES:
        source_path = PROJECT_ROOT / file_name
        if source_path.exists():
            _copy_file(source_path, DEPLOY_ROOT / file_name)
    for dir_name in RUNTIME_DIRS:
        source_dir = PROJECT_ROOT / dir_name
        if not source_dir.exists():
            continue
        ignore_patterns = ("__pycache__", "*.pyc")
        if dir_name == "workflow":
            ignore_patterns = ("__pycache__", "*.pyc", "dashboard_outputs")
        _copy_tree(source_dir, DEPLOY_ROOT / dir_name, ignore_patterns=ignore_patterns)
    log("[Deploy] Copied runtime source files")


def copy_bom_lake_data() -> None:
    source_root = PROJECT_ROOT / "bom_lake"
    target_root = DEPLOY_ROOT / "data" / "bom_lake"
    if not source_root.exists():
        return
    _copy_tree(source_root, target_root, ignore_patterns=("__pycache__", "*.pyc"))
    log(f"[Deploy] Copied BOM lake data: {target_root}")


def copy_reference_files() -> None:
    sample_target = DEPLOY_ROOT / "data" / "sample_files"
    template_target = DEPLOY_ROOT / "data" / "scenario_templates"

    copied_samples = 0
    for pattern in SAMPLE_FILE_GLOBS:
        for source_path in PROJECT_ROOT.glob(pattern):
            _copy_file(source_path, sample_target / source_path.name)
            copied_samples += 1

    copied_templates = 0
    seen_template_paths: set[Path] = set()
    for base_dir in [PROJECT_ROOT, PROJECT_ROOT / "workflow"]:
        if not base_dir.exists():
            continue
        for pattern in SCENARIO_TEMPLATE_GLOBS:
            for source_path in base_dir.glob(pattern):
                if source_path in seen_template_paths:
                    continue
                seen_template_paths.add(source_path)
                _copy_file(source_path, template_target / source_path.name)
                copied_templates += 1

    log(f"[Deploy] Copied sample files: {copied_samples}")
    log(f"[Deploy] Copied scenario templates: {copied_templates}")


def write_config_file() -> Path:
    config_path = DEPLOY_ROOT / "config.json"
    config_path.write_text(json.dumps(CONFIG_PAYLOAD, ensure_ascii=True, indent=2), encoding="utf-8")
    log(f"[Deploy] Wrote config: {config_path}")
    return config_path


def write_readme_file() -> None:
    readme_path = DEPLOY_ROOT / "README_실행방법.txt"
    readme_text = """VICT 실행 방법

1. 폴더 전체를 원하는 위치에 그대로 압축 해제하거나 복사합니다.
2. 폴더 구조를 변경하지 말고 VICT.exe를 실행합니다.
3. Base/Current BOM parquet를 선택하거나 기본으로 포함된 데이터를 사용합니다.
4. 결과 파일은 outputs 폴더에 저장됩니다.
5. GUI에서 Dashboard 열기 기능으로 대시보드를 실행할 수 있습니다.
6. 문제가 생기면 logs 폴더 안의 로그 파일을 함께 전달해 주세요.

중요 안내
- Python 설치가 필요 없습니다.
- data 폴더와 cache 폴더를 삭제하지 마세요.
- Teams 공유 폴더에서 받은 뒤 다른 위치로 복사해도 실행할 수 있습니다.
- data/bom_lake 안의 parquet와 cache 안의 사전 생성 캐시를 그대로 두면 초기 대기 시간을 줄일 수 있습니다.
수동 디버그 명령
- Source:
  python -m streamlit run streamlit_vi_dashboard.py --server.address localhost --server.port 8501 --server.headless true -- --report-file "outputs/report.xlsx"
- Deploy:
  VICT.exe --launch-dashboard --report-file "outputs/report.xlsx" --port 8501
"""
    readme_path.write_text(readme_text, encoding="utf-8")
    log(f"[Deploy] Wrote README: {readme_path}")


def _discover_month_dirs(root: Path) -> list[tuple[int, int]]:
    months: list[tuple[int, int]] = []
    if not root.exists():
        return months
    for year_dir in sorted(root.glob("bom_year=*")):
        year_text = year_dir.name.replace("bom_year=", "").strip()
        if not year_text.isdigit():
            continue
        for month_dir in sorted(year_dir.glob("bom_month=*")):
            month_text = month_dir.name.replace("bom_month=", "").strip()
            if not month_text.isdigit():
                continue
            if any(month_dir.glob("*.parquet")):
                months.append((int(year_text), int(month_text)))
    return months


def _select_prebuild_pair(deploy_lake_root: Path) -> tuple[int, int, str, int, int] | None:
    historical_months = _discover_month_dirs(deploy_lake_root / "historical")
    current_cache_months = _discover_month_dirs(deploy_lake_root / "current_cache")
    historical_set = set(historical_months)
    current_cache_set = set(current_cache_months)

    if (2025, 12) in historical_set and (2026, 4) in current_cache_set:
        return 2025, 12, "current_cache", 2026, 4
    if (2025, 12) in historical_set and current_cache_months:
        latest_current_year, latest_current_month = sorted(current_cache_months)[-1]
        return 2025, 12, "current_cache", latest_current_year, latest_current_month
    if historical_months and current_cache_months:
        base_year, base_month = sorted(historical_months)[-1]
        current_year, current_month = sorted(current_cache_months)[-1]
        return base_year, base_month, "current_cache", current_year, current_month
    if len(historical_months) >= 2:
        ordered = sorted(historical_months)
        base_year, base_month = ordered[-2]
        current_year, current_month = ordered[-1]
        return base_year, base_month, "historical", current_year, current_month
    return None


@contextmanager
def deploy_runtime_context(config_path: Path):
    original_config = os.environ.get("VICT_CONFIG")
    try:
        os.environ["VICT_CONFIG"] = str(config_path.resolve())
        clear_runtime_path_caches()
        vi_engine.load_bom_snapshot_compact.cache_clear()
        vi_engine._snapshot_columns.cache_clear()
        yield
    finally:
        if original_config is None:
            os.environ.pop("VICT_CONFIG", None)
        else:
            os.environ["VICT_CONFIG"] = original_config
        clear_runtime_path_caches()
        vi_engine.load_bom_snapshot_compact.cache_clear()
        vi_engine._snapshot_columns.cache_clear()


def prebuild_bom_pair_cache(config_path: Path) -> None:
    deploy_lake_root = DEPLOY_ROOT / "data" / "bom_lake"
    selected_pair = _select_prebuild_pair(deploy_lake_root)
    if selected_pair is None:
        log("[Deploy] No BOM snapshot pair available for prebuild")
        return

    base_year, base_month, current_kind, current_year, current_month = selected_pair
    log(
        "[Deploy] Prebuilding bom_pair cache "
        f"| base={base_year:04d}-{base_month:02d} "
        f"| current={current_year:04d}-{current_month:02d} "
        f"| current_kind={current_kind}"
    )
    with deploy_runtime_context(config_path):
        vi_engine.load_or_build_bom_pair_compact_cache(
            history_root_text=str(deploy_lake_root),
            base_year=base_year,
            base_month=base_month,
            base_snapshot_kind="historical",
            base_snapshot_path=None,
            current_year=current_year,
            current_month=current_month,
            current_snapshot_kind=current_kind,
            current_snapshot_path=None,
        )


def copy_portable_report_caches() -> None:
    source_cache_root = PROJECT_ROOT / "workflow" / "dashboard_outputs" / "cache"
    target_cache_root = DEPLOY_ROOT / "cache" / "report_cache"
    copied_count = 0
    if not source_cache_root.exists():
        log("[Deploy] Portable report cache copied: 0")
        return
    for cache_dir in source_cache_root.iterdir():
        if not cache_dir.is_dir() or cache_dir.name.startswith("bom_pair_"):
            continue
        metadata_path = cache_dir / "metadata.json"
        if not metadata_path.exists():
            continue
        try:
            metadata = json.loads(metadata_path.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            continue
        scenario_path = str(metadata.get("scenario_path", "") or "")
        if ":" in scenario_path or scenario_path.startswith("/"):
            continue
        _copy_tree(cache_dir, target_cache_root / cache_dir.name, ignore_patterns=("__pycache__", "*.pyc"))
        copied_count += 1
    log(f"[Deploy] Portable report cache copied: {copied_count}")


def main() -> int:
    ensure_workspace_structure()
    snapshot_existing_generated_outputs()
    archive_existing_deploy_folder()
    create_deploy_structure()
    copy_build_artifacts()
    copy_runtime_sources()
    copy_bom_lake_data()
    copy_reference_files()
    config_path = write_config_file()
    write_readme_file()
    prebuild_bom_pair_cache(config_path)
    copy_portable_report_caches()
    log(f"[Deploy] Final folder ready: {DEPLOY_ROOT}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
