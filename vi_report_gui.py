from __future__ import annotations

import contextlib
import ctypes
import hashlib
import importlib
import importlib.util
import os
import queue
import socket
import subprocess
import sys
import threading
import time
import traceback
import unicodedata
import urllib.error
import urllib.request
import webbrowser
from pathlib import Path
from typing import TYPE_CHECKING, Any, Callable

if TYPE_CHECKING:
    from pandas import DataFrame as PandasDataFrame
else:
    PandasDataFrame = Any

try:
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk
except Exception as exc:  # pragma: no cover
    raise RuntimeError("tkinter is required to run the VI GUI.") from exc


BOOTSTRAP_DIR = Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parent
SCRIPT_DIR = BOOTSTRAP_DIR
PARENT_DIR = SCRIPT_DIR.parent
for path in [SCRIPT_DIR, PARENT_DIR]:
    if str(path) not in sys.path:
        sys.path.insert(0, str(path))

from runtime_paths import (
    application_root,
    is_frozen_app,
    log_root,
    output_root,
    sample_files_root,
    scenario_templates_root,
)

SCRIPT_DIR = application_root()


RUNTIME_PACKAGES = {
    "duckdb": "duckdb",
    "openpyxl": "openpyxl",
    "pandas": "pandas",
    "xlsxwriter": "xlsxwriter",
}

default_output_path: Callable[..., Path]
default_recommendation_output_path: Callable[..., Path]
default_scenario_path: Callable[..., Path | None]
open_workbook_for_review: Callable[..., Any]
parse_month_text: Callable[..., tuple[int, int]]
run_recommendation_flow: Callable[..., Any]
run_report_flow: Callable[..., Any]
scenario_file_hash: Callable[..., str]
save_current_snapshot_from_excel: Callable[..., dict[str, Any]]
save_historical_snapshot_from_excel: Callable[..., dict[str, Any]]
inspect_parquet_snapshot: Callable[..., dict[str, Any]]
scenario_template_columns: dict[str, list[str]]
pd: Any

TEMPLATE_HEADER_LABELS = {
    "VI_Item_Master": {
        "vi_item_id": "VI 아이템 ID",
        "vi_item_name": "VI 아이템명",
        "owner": "담당자",
        "note": "비고",
    },
    "BasePN_Recommendation_Request": {
        "vi_item_id": "VI 아이템 ID",
        "vi_item_name": "VI아이템명",
        "enabled": "사용 여부",
        "subsidiary": "대상 법인",
        "base_part_no": "기준 품번(Base P/N)",
        "new_part_no": "적용 New P/N",
        "model_type": "모델 타입(C/O,H/P)",
        "indoor_tool": "실내기 툴(SJ,SK,S0,SA)",
        "desc_match_mode": "Desc. 키워드 매칭방식",
        "desc_contains_any": "Desc. 키워드",
        "spec_contains_any": "Spec. 키워드(OR)",
        "parent_assy_desc_contains_any": "상위 ASSY Desc. 키워드",
        "note": "비고",
    },
    "BasePN_Recommendation_Result": {
        "vi_item_id": "VI 아이템 ID",
        "subsidiary": "대상 법인",
        "base_part_no": "기준 품번(Base P/N)",
        "new_part_no": "적용 New P/N",
        "note": "비고",
        "vi_item_name": "VI 아이템명",
        "recommended_base_part_no": "추천 Base P/N",
        "parent_assy_desc_text": "상위 Assy Description",
        "recommended_desc": "Description",
        "recommended_subsidiaries": "사용 중인 법인",
        "bom_subsidiary_model_count_detail": "원본 BOM 법인별 모델 수",
        "base_bom_total_model_count": "Base BOM전체 모델수",
        "item_target_model_count": "VI대상 모델 수",
        "item_target_subsidiary_model_count_detail": "대상 법인별 모델 수",
        "recommended_model_count": "Base P/N 사용 모델 수",
        "recommended_model_share_formula": "추천 비중 계산식",
        "recommended_model_share_pct": "Base P/N 사용 비중",
        "recommended_subsidiary_share_detail": "법인별 사용 비중",
        "base_pn_subsidiaries": "법인",
        "base_pn_models": "모델명",
        "recommendation_reason": "추천 근거",
        "user_confirmed": "사용 여부",
        "confirmed_new_part_no": "확정 New P/N",
    },
    "Volume": {
        "month": "생산월(YYYY-MM)",
        "subsidiary": "법인",
        "model_suffix": "모델명",
        "production_qty": "생산수량",
    },
}

TEMPLATE_COLUMN_GUIDES = {
    "VI_Item_Master": {
        "vi_item_id": {
            "input_type": "필수 입력",
            "owner": "사용자",
            "description": "VI 아이템 고유 ID입니다. 같은 ID를 추천요청 시트에서도 사용합니다.",
        },
        "vi_item_name": {
            "input_type": "필수 입력",
            "owner": "사용자",
            "description": "VI 아이템 이름입니다.",
        },
        "owner": {
            "input_type": "선택 입력",
            "owner": "사용자",
            "description": "담당자 메모용입니다.",
        },
        "note": {
            "input_type": "선택 입력",
            "owner": "사용자",
            "description": "비고 메모용입니다.",
        },
    },
    "BasePN_Recommendation_Request": {
        "vi_item_id": {
            "input_type": "필수 입력",
            "owner": "사용자",
            "description": "VI_Item_Master에 있는 VI 아이템 ID를 적습니다.",
        },
        "vi_item_name": {
            "input_type": "자동 생성",
            "owner": "시스템",
            "description": "VI 아이템 ID를 기준으로 Master 시트에서 자동으로 채워집니다.",
        },
        "enabled": {
            "input_type": "선택 입력",
            "owner": "사용자",
            "description": "비우면 Y로 처리됩니다.",
        },
        "subsidiary": {
            "input_type": "조건 입력",
            "owner": "사용자",
            "description": "추천 대상을 특정 법인으로 제한할 때 입력합니다. 비우면 전체입니다.",
        },
        "base_part_no": {
            "input_type": "조건 입력",
            "owner": "사용자",
            "description": "특정 Base P/N으로 직접 추천을 찾을 때 입력합니다.",
        },
        "new_part_no": {
            "input_type": "조건 입력",
            "owner": "사용자",
            "description": "Base P/N과 함께 이미 알고 있는 적용 New P/N이 있으면 바로 입력합니다.",
        },
        "model_type": {
            "input_type": "조건 입력",
            "owner": "사용자",
            "description": "C/O 또는 H/P 조건입니다.",
        },
        "indoor_tool": {
            "input_type": "조건 입력",
            "owner": "사용자",
            "description": "SJ, SK, S0, SA 조건입니다.",
        },
        "desc_match_mode": {
            "input_type": "선택 입력",
            "owner": "사용자",
            "description": "설명/스펙/상위 ASSY 키워드의 매칭 방식입니다. 비우면 OR입니다.",
        },
        "desc_contains_any": {
            "input_type": "조건 입력",
            "owner": "사용자",
            "description": "대상 설명 키워드입니다. 쉼표로 여러 개를 넣을 수 있습니다.",
        },
        "spec_contains_any": {
            "input_type": "조건 입력",
            "owner": "사용자",
            "description": "스펙 키워드입니다.",
        },
        "parent_assy_desc_contains_any": {
            "input_type": "조건 입력",
            "owner": "사용자",
            "description": "상위 ASSY 설명 키워드입니다.",
        },
        "note": {
            "input_type": "선택 입력",
            "owner": "사용자",
            "description": "요청 메모입니다.",
        },
    },
    "BasePN_Recommendation_Result": {
        "vi_item_id": {
            "input_type": "자동 생성",
            "owner": "시스템",
            "description": "추천 실행 후 자동으로 채워집니다.",
        },
        "subsidiary": {
            "input_type": "자동 생성",
            "owner": "시스템",
            "description": "추천 실행 후 자동으로 채워집니다.",
        },
        "base_part_no": {
            "input_type": "자동 생성",
            "owner": "시스템",
            "description": "추천 실행 후 자동으로 채워집니다.",
        },
        "new_part_no": {
            "input_type": "자동 생성",
            "owner": "시스템",
            "description": "추천 실행 후 자동으로 채워집니다.",
        },
        "note": {
            "input_type": "자동 생성",
            "owner": "시스템",
            "description": "추천 실행 후 자동으로 채워집니다.",
        },
        "vi_item_name": {
            "input_type": "자동 생성",
            "owner": "시스템",
            "description": "추천 실행 후 자동으로 채워집니다.",
        },
        "recommended_base_part_no": {
            "input_type": "자동 생성",
            "owner": "시스템",
            "description": "시스템 추천 결과입니다.",
        },
        "parent_assy_desc_text": {
            "input_type": "자동 생성",
            "owner": "시스템",
            "description": "상위 ASSY Desc 조건으로 매칭된 경우 가장 가까운 상위 ASSY 설명입니다.",
        },
        "recommended_desc": {
            "input_type": "자동 생성",
            "owner": "시스템",
            "description": "시스템 추천 결과입니다.",
        },
        "recommended_subsidiaries": {
            "input_type": "자동 생성",
            "owner": "시스템",
            "description": "시스템 추천 결과입니다.",
        },
        "bom_subsidiary_model_count_detail": {
            "input_type": "자동 생성",
            "owner": "시스템",
            "description": "시스템 추천 결과입니다.",
        },
        "base_bom_total_model_count": {
            "input_type": "자동 생성",
            "owner": "시스템",
            "description": "선택한 Base BOM 전체의 모델 수입니다.",
        },
        "item_target_model_count": {
            "input_type": "자동 생성",
            "owner": "시스템",
            "description": "시스템 추천 결과입니다.",
        },
        "item_target_subsidiary_model_count_detail": {
            "input_type": "자동 생성",
            "owner": "시스템",
            "description": "시스템 추천 결과입니다.",
        },
        "recommended_model_count": {
            "input_type": "자동 생성",
            "owner": "시스템",
            "description": "시스템 추천 결과입니다.",
        },
        "recommended_model_share_formula": {
            "input_type": "자동 생성",
            "owner": "시스템",
            "description": "시스템 추천 결과입니다.",
        },
        "recommended_model_share_pct": {
            "input_type": "자동 생성",
            "owner": "시스템",
            "description": "시스템 추천 결과입니다.",
        },
        "recommended_subsidiary_share_detail": {
            "input_type": "자동 생성",
            "owner": "시스템",
            "description": "시스템 추천 결과입니다.",
        },
        "base_pn_subsidiaries": {
            "input_type": "자동 생성",
            "owner": "시스템",
            "description": "해당 Base P/N이 존재하는 법인 목록입니다.",
        },
        "base_pn_models": {
            "input_type": "자동 생성",
            "owner": "시스템",
            "description": "해당 Base P/N이 존재하는 모델 목록입니다.",
        },
        "recommendation_reason": {
            "input_type": "자동 생성",
            "owner": "시스템",
            "description": "시스템 추천 근거입니다.",
        },
        "user_confirmed": {
            "input_type": "검토 입력",
            "owner": "사용자",
            "description": "추천 검토 후 사용할 행이면 Y를 입력합니다.",
        },
        "confirmed_new_part_no": {
            "input_type": "검토 입력",
            "owner": "사용자",
            "description": "최종 적용할 New P/N을 입력합니다.",
        },
    },
    "Volume": {
        "month": {
            "input_type": "필수 입력",
            "owner": "사용자",
            "description": "생산월입니다. YYYY-MM 형식으로 입력합니다.",
        },
        "subsidiary": {
            "input_type": "필수 입력",
            "owner": "사용자",
            "description": "법인명입니다.",
        },
        "model_suffix": {
            "input_type": "필수 입력",
            "owner": "사용자",
            "description": "모델명입니다.",
        },
        "production_qty": {
            "input_type": "필수 입력",
            "owner": "사용자",
            "description": "생산수량입니다.",
        },
    },
}


def ensure_runtime_package(module_name: str, package_name: str) -> None:
    try:
        importlib.import_module(module_name)
        return
    except ModuleNotFoundError:
        pass
    if is_frozen_app():
        raise RuntimeError(
            f"Required packaged module is missing in the EXE runtime: {module_name}"
        )
    subprocess.check_call([sys.executable, "-m", "pip", "install", package_name])
    importlib.invalidate_caches()


def _fallback_default_output_path(base_year: int, base_month: int, current_year: int, current_month: int) -> Path:
    return output_root() / f"VI_Result_{base_year:04d}-{base_month:02d}_to_{current_year:04d}-{current_month:02d}.xlsx"


def _fallback_default_recommendation_output_path(scenario_path: Path, base_year: int, base_month: int) -> Path:
    return scenario_path.with_name(f"{scenario_path.stem}_Recommendation_{base_year:04d}-{base_month:02d}.xlsx")


def _fallback_default_scenario_path() -> Path | None:
    candidates = [
        scenario_templates_root() / "VI_Scenario_Template_Advanced_v3.xlsx",
        scenario_templates_root() / "VI_Scenario_Template_Advanced_v2.xlsx",
        scenario_templates_root() / "VI_Scenario_Template_Advanced.xlsx",
        scenario_templates_root() / "VI_Scenario_Template_Custom.xlsx",
        scenario_templates_root() / "VI_Scenario_Template.xlsx",
    ]
    for candidate in candidates:
        if candidate.exists():
            return candidate
    xlsx_files = sorted(scenario_templates_root().glob("*.xlsx"))
    if xlsx_files:
        return xlsx_files[0]
    sample_xlsx_files = sorted(sample_files_root().glob("*.xlsx"))
    if sample_xlsx_files:
        return sample_xlsx_files[0]
    return None


def _fallback_open_workbook_for_review(path: Path) -> None:
    try:
        os.startfile(str(path))
    except Exception:
        return None


def _fallback_parse_month_text(value: str) -> tuple[int, int]:
    text = str(value).strip()
    cleaned = text.replace("/", "-").replace(".", "-")
    parts = [part.strip() for part in cleaned.split("-") if part.strip()]
    if len(parts) == 2:
        year = int(parts[0])
        month = int(parts[1])
    elif len(text) == 6 and text.isdigit():
        year = int(text[:4])
        month = int(text[4:])
    else:
        raise ValueError(f"Invalid month text: {value}")
    if month < 1 or month > 12:
        raise ValueError(f"Invalid month text: {value}")
    return year, month


def _fallback_scenario_file_hash(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _fallback_inspect_parquet_snapshot(parquet_path: Path) -> dict[str, Any]:
    import re

    resolved_path = Path(parquet_path).expanduser().resolve()
    bom_year = 0
    bom_month = 0
    row_count = 0
    model_count = 0
    label = resolved_path.name

    path_match = re.search(r"bom_year=(\d{4}).*?bom_month=(\d{1,2})", str(resolved_path).replace("\\", "/"))
    if path_match:
        bom_year = int(path_match.group(1))
        bom_month = int(path_match.group(2))

    con = None
    try:
        import duckdb

        con = duckdb.connect()
        row = con.execute(
            """
            SELECT
                COALESCE(MAX(TRY_CAST(bom_year AS INTEGER)), 0) AS bom_year,
                COALESCE(MAX(TRY_CAST(bom_month AS INTEGER)), 0) AS bom_month,
                COUNT(*) AS row_count,
                COUNT(DISTINCT model_suffix) AS model_count
            FROM read_parquet(?)
            """,
            [str(resolved_path)],
        ).fetchone()
        if row is not None:
            bom_year = int(row[0] or bom_year or 0)
            bom_month = int(row[1] or bom_month or 0)
            row_count = int(row[2] or 0)
            model_count = int(row[3] or 0)
    finally:
        if con is not None:
            con.close()

    if bom_year and bom_month:
        label = f"{bom_year:04d}-{bom_month:02d} | {resolved_path.name}"

    return {
        "label": label,
        "parquet_path": str(resolved_path),
        "bom_year": bom_year,
        "bom_month": bom_month,
        "row_count": row_count,
        "model_count": model_count,
    }


def _module_attr(module: Any, name: str, fallback: Any | None = None) -> Any:
    if hasattr(module, name):
        return getattr(module, name)
    if fallback is not None:
        return fallback
    module_file = getattr(module, "__file__", "<unknown>")
    raise ImportError(f"cannot import name '{name}' from '{module.__name__}' ({module_file})")


def _display_text_width(value: object) -> int:
    text = "" if value is None else str(value)
    width = 0
    for char in text:
        width += 2 if unicodedata.east_asian_width(char) in {"W", "F"} else 1
    return width


def _auto_fit_openpyxl_worksheet(worksheet: Any) -> None:
    from openpyxl.utils import get_column_letter

    for column_cells in worksheet.columns:
        max_width = 0
        column_letter = get_column_letter(column_cells[0].column)
        for cell in column_cells:
            max_width = max(max_width, _display_text_width(cell.value))
        worksheet.column_dimensions[column_letter].width = min(max(max_width + 2, 12), 48)


def _friendly_template_frame(frame: PandasDataFrame, sheet_name: str) -> PandasDataFrame:
    header_map = TEMPLATE_HEADER_LABELS.get(sheet_name, {})
    return frame.rename(columns={column: header_map.get(column, column) for column in frame.columns})


def _build_template_sample_frame(columns: list[str], rows: list[dict[str, object]]) -> PandasDataFrame:
    ordered_rows = [[row.get(column, "") for column in columns] for row in rows]
    return pd.DataFrame(ordered_rows, columns=columns)


def _build_scenario_input_guide() -> PandasDataFrame:
    rows: list[dict[str, str]] = []
    for sheet_name, columns in scenario_template_columns.items():
        header_map = TEMPLATE_HEADER_LABELS.get(sheet_name, {})
        guide_map = TEMPLATE_COLUMN_GUIDES.get(sheet_name, {})
        for column in columns:
            info = guide_map.get(
                column,
                {
                    "input_type": "선택 입력",
                    "owner": "사용자",
                    "description": "",
                },
            )
            rows.append(
                {
                    "시트명": sheet_name,
                    "표시 헤더": header_map.get(column, column),
                    "엔진 컬럼명": column,
                    "입력구분": str(info["input_type"]),
                    "작성주체": str(info["owner"]),
                    "설명": str(info["description"]),
                }
            )
    return pd.DataFrame(rows)


def _apply_template_header_styles(worksheet: Any, sheet_name: str) -> None:
    from openpyxl.comments import Comment
    from openpyxl.styles import Border, Font, PatternFill, Side

    required_by_sheet = {
        "BasePN_Recommendation_Request": {"vi_item_name", "base_part_no", "new_part_no"},
        "BasePN_Recommendation_Result": {"user_confirmed", "confirmed_new_part_no"},
    }
    optional_by_sheet = {
        "BasePN_Recommendation_Request": {
            "enabled",
            "subsidiary",
            "model_type",
            "indoor_tool",
            "desc_match_mode",
            "desc_contains_any",
            "spec_contains_any",
            "parent_assy_desc_contains_any",
            "note",
        },
    }
    fill_required = PatternFill(fill_type="solid", fgColor="8B0000")
    fill_optional = PatternFill(fill_type="solid", fgColor="DDEBF7")
    fill_system = PatternFill(fill_type="solid", fgColor="E7E6E6")
    required_font = Font(color="FFFFFF", bold=True)
    normal_font = Font(bold=True)
    thick_red = Side(style="thick", color="C00000")
    required_border = Border(left=thick_red, right=thick_red, top=thick_red, bottom=thick_red)

    original_columns = scenario_template_columns.get(sheet_name, [])
    required_columns = required_by_sheet.get(sheet_name, set())
    optional_columns = optional_by_sheet.get(sheet_name, set())
    for index, column in enumerate(original_columns, start=1):
        cell = worksheet.cell(row=1, column=index)
        if column in required_columns:
            cell.fill = fill_required
            cell.font = required_font
            cell.border = required_border
            cell.comment = Comment("[필수 입력]", "VICT")
        elif column in optional_columns:
            cell.fill = fill_optional
            cell.font = normal_font
        else:
            cell.fill = fill_system
            cell.font = normal_font


def load_vi_report_api() -> None:
    global pd
    for module_name, package_name in RUNTIME_PACKAGES.items():
        ensure_runtime_package(module_name, package_name)
    pd = importlib.import_module("pandas")

    global default_output_path
    global default_recommendation_output_path
    global default_scenario_path
    global open_workbook_for_review
    global parse_month_text
    global run_recommendation_flow
    global run_report_flow
    global scenario_file_hash
    global inspect_parquet_snapshot
    global save_current_snapshot_from_excel
    global save_historical_snapshot_from_excel
    global scenario_template_columns

    try:
        vi_report_cli_module = importlib.import_module("vi_report_cli")
        bom_lake_module = importlib.import_module("bom_lake")
        vi_engine_module = importlib.import_module("vi_calculation_engine")
    except ModuleNotFoundError:
        vi_report_cli_module = importlib.import_module("compat_imports.vi_report_cli")
        bom_lake_module = importlib.import_module("compat_imports.bom_lake")
        vi_engine_module = importlib.import_module("compat_imports.vi_calculation_engine")

    default_output_path = _module_attr(
        vi_report_cli_module,
        "default_output_path",
        _fallback_default_output_path,
    )
    default_recommendation_output_path = _module_attr(
        vi_report_cli_module,
        "default_recommendation_output_path",
        _fallback_default_recommendation_output_path,
    )
    default_scenario_path = _module_attr(
        vi_report_cli_module,
        "default_scenario_path",
        _fallback_default_scenario_path,
    )
    open_workbook_for_review = _module_attr(
        vi_report_cli_module,
        "open_workbook_for_review",
        _fallback_open_workbook_for_review,
    )
    parse_month_text = _module_attr(
        vi_report_cli_module,
        "parse_month_text",
        _fallback_parse_month_text,
    )
    run_recommendation_flow = _module_attr(vi_report_cli_module, "run_recommendation_flow")
    run_report_flow = _module_attr(vi_report_cli_module, "run_report_flow")
    scenario_file_hash = _module_attr(
        vi_report_cli_module,
        "scenario_file_hash",
        _fallback_scenario_file_hash,
    )
    inspect_parquet_snapshot = _module_attr(
        bom_lake_module,
        "inspect_parquet_snapshot",
        _fallback_inspect_parquet_snapshot,
    )
    save_current_snapshot_from_excel = _module_attr(bom_lake_module, "save_current_snapshot_from_excel")
    save_historical_snapshot_from_excel = _module_attr(bom_lake_module, "save_historical_snapshot_from_excel")
    RECOMMENDATION_REQUEST_COLUMNS = _module_attr(vi_engine_module, "RECOMMENDATION_REQUEST_COLUMNS")
    RECOMMENDATION_RESULT_COLUMNS = _module_attr(vi_engine_module, "RECOMMENDATION_RESULT_COLUMNS")
    VI_ITEM_MASTER_COLUMNS = _module_attr(vi_engine_module, "VI_ITEM_MASTER_COLUMNS")
    VOLUME_COLUMNS = _module_attr(vi_engine_module, "VOLUME_COLUMNS")

    scenario_template_columns = {
        "BasePN_Recommendation_Request": [column for column in RECOMMENDATION_REQUEST_COLUMNS if column != "vi_item_id"],
        "BasePN_Recommendation_Result": [column for column in RECOMMENDATION_RESULT_COLUMNS if column != "vi_item_id"],
        "Volume": list(VOLUME_COLUMNS),
    }


def _stop_existing_dashboard_processes(dashboard_script: Path) -> None:
    script_path = str(dashboard_script.resolve()).replace("'", "''")
    command = (
        f"$target = '{script_path}'; "
        "Get-CimInstance Win32_Process | "
        "Where-Object { $_.CommandLine -and $_.CommandLine -like '*streamlit*run*' -and $_.CommandLine -like ('*' + $target + '*') } | "
        "ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }"
    )
    subprocess.run(
        ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", command],
        check=False,
        capture_output=True,
        text=True,
    )
    time.sleep(0.5)


def _dashboard_launch_log_path() -> Path:
    return log_root() / "dashboard_launch.log"


def _dashboard_stdout_log_path() -> Path:
    return log_root() / "dashboard_streamlit_stdout.log"


def _dashboard_stderr_log_path() -> Path:
    return log_root() / "dashboard_streamlit_stderr.log"


def _write_dashboard_launch_log(message: str) -> None:
    try:
        log_path = _dashboard_launch_log_path()
        log_path.parent.mkdir(parents=True, exist_ok=True)
        with log_path.open("a", encoding="utf-8") as handle:
            handle.write(message.rstrip() + "\n")
    except Exception:
        return None


def _write_dashboard_error_log(message: str) -> None:
    try:
        log_path = _dashboard_stderr_log_path()
        log_path.parent.mkdir(parents=True, exist_ok=True)
        with log_path.open("a", encoding="utf-8") as handle:
            handle.write(message.rstrip() + "\n")
    except Exception:
        return None


def _is_dashboard_port_available(port: int) -> bool:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
        sock.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
        try:
            sock.bind(("127.0.0.1", int(port)))
        except OSError:
            return False
    return True


def _select_dashboard_port() -> int:
    for port in [8501, 8502, 8503]:
        if _is_dashboard_port_available(port):
            return port
    raise RuntimeError("No available dashboard port found in 8501, 8502, 8503.")


def _dashboard_url(port: int) -> str:
    return f"http://localhost:{int(port)}"


def _wait_for_dashboard_ready(port: int, timeout_seconds: float = 20.0) -> bool:
    deadline = time.perf_counter() + float(timeout_seconds)
    url = _dashboard_url(port)
    while time.perf_counter() < deadline:
        try:
            with urllib.request.urlopen(url, timeout=1.5) as response:
                status_code = int(getattr(response, "status", 200) or 200)
                if 200 <= status_code < 400:
                    return True
        except urllib.error.URLError:
            pass
        except Exception:
            pass
        time.sleep(0.5)
    return False


def _open_dashboard_browser(port: int) -> None:
    url = _dashboard_url(port)
    _write_dashboard_launch_log(f"[Dashboard Browser] url={url}")
    webbrowser.open(url, new=2)


def _windows_creation_flags() -> int:
    return int(getattr(subprocess, "CREATE_NO_WINDOW", 0))


def _format_launch_command(args: list[str]) -> str:
    return subprocess.list2cmdline([str(part) for part in args])


def _extract_cli_arg_value(flag: str) -> str:
    if flag not in sys.argv:
        return ""
    flag_index = sys.argv.index(flag)
    if flag_index + 1 >= len(sys.argv):
        return ""
    return str(sys.argv[flag_index + 1])


def launch_streamlit_dashboard(report_path: Path) -> None:
    dashboard_script = SCRIPT_DIR / "streamlit_vi_dashboard.py"
    if not dashboard_script.exists():
        messagebox.showerror("대시보드 실행 실패", "streamlit_vi_dashboard.py 파일이 없습니다.")
        return
    try:
        ensure_runtime_package("streamlit", "streamlit")
        ensure_runtime_package("plotly", "plotly")
    except Exception as exc:
        _write_dashboard_error_log(traceback.format_exc())
        messagebox.showerror(
            "대시보드 실행 실패",
            "Streamlit 또는 Plotly를 준비하지 못했습니다.\n\n"
            f"{sys.executable} -m pip install streamlit plotly\n\n오류 내용:\n{exc}",
        )
        return

    report_path = report_path.expanduser().resolve()
    dashboard_cwd = dashboard_script.parent.resolve()
    stdout_log_path = _dashboard_stdout_log_path()
    stderr_log_path = _dashboard_stderr_log_path()
    stdout_log_path.parent.mkdir(parents=True, exist_ok=True)
    stderr_log_path.parent.mkdir(parents=True, exist_ok=True)
    selected_port = _select_dashboard_port()
    if is_frozen_app():
        launch_args = [
            sys.executable,
            "--launch-dashboard",
            "--report-file",
            str(report_path),
            "--port",
            str(selected_port),
        ]
    else:
        launch_args = [
            sys.executable,
            "-m",
            "streamlit",
            "run",
            dashboard_script.name,
            "--global.developmentMode",
            "false",
            "--server.address",
            "localhost",
            "--server.port",
            str(selected_port),
            "--server.headless",
            "true",
            "--browser.gatherUsageStats",
            "false",
            "--",
            "--report-file",
            str(report_path),
        ]
    launch_header = (
        "[Dashboard Launch]\n"
        f"python={Path(sys.executable).resolve()}\n"
        f"script={dashboard_script.resolve()}\n"
        f"cwd={dashboard_cwd}\n"
        f"report={report_path}\n"
        f"port={selected_port}\n"
        f"stdout_log={stdout_log_path.resolve()}\n"
        f"stderr_log={stderr_log_path.resolve()}\n"
        f"command={_format_launch_command(launch_args)}\n"
    )
    print(launch_header, flush=True)
    _write_dashboard_launch_log(launch_header)
    try:
        popen_kwargs: dict[str, object] = {
            "cwd": str(dashboard_cwd),
            "creationflags": _windows_creation_flags(),
        }
        stdout_handle = None
        stderr_handle = None
        if not is_frozen_app():
            stdout_handle = stdout_log_path.open("w", encoding="utf-8")
            stderr_handle = stderr_log_path.open("w", encoding="utf-8")
            popen_kwargs["stdout"] = stdout_handle
            popen_kwargs["stderr"] = stderr_handle
            popen_kwargs["text"] = True
        process = subprocess.Popen(launch_args, **popen_kwargs)
        if stdout_handle is not None:
            stdout_handle.close()
        if stderr_handle is not None:
            stderr_handle.close()
        launch_result = f"[Dashboard Launch PID] pid={process.pid}\n"
        print(launch_result, flush=True)
        _write_dashboard_launch_log(launch_result)
        if _wait_for_dashboard_ready(selected_port):
            _open_dashboard_browser(selected_port)
        else:
            _write_dashboard_error_log(
                "[Dashboard Ready Timeout]\n"
                f"port={selected_port}\n"
                f"report={report_path}\n"
                f"command={_format_launch_command(launch_args)}\n"
            )
            messagebox.showerror(
                "??쒕낫???ㅽ뻾 ?ㅽ뙣",
                "??쒕낫?쒕? ?ㅽ뻾?덉쑝?섎굹 ?쒓컙 ?덈줈 ?덉쭅 ?곸냽?섏? 紐삵뻽?듬땲??\n"
                f"stderr 濡쒓렇: {stderr_log_path}",
            )
    except Exception as exc:
        messagebox.showerror("대시보드 실행 실패", f"대시보드를 실행하지 못했습니다.\n{exc}")


class QueueWriter:
    def __init__(self, output_queue: queue.Queue[str]) -> None:
        self.output_queue = output_queue
        self._buffer = ""

    def write(self, text: str) -> int:
        if not text:
            return 0
        self._buffer += text
        while "\n" in self._buffer:
            line, self._buffer = self._buffer.split("\n", 1)
            if line.strip():
                self.output_queue.put(f"{line}\n")
        return len(text)

    def flush(self) -> None:
        if self._buffer.strip():
            self.output_queue.put(f"{self._buffer.rstrip()}\n")
        self._buffer = ""
        return None


class NullWriter:
    def write(self, _text: str) -> int:
        return 0

    def flush(self) -> None:
        return None


def _raise_async_exception(thread: threading.Thread | None, exception_type: type[BaseException]) -> bool:
    if thread is None or not thread.is_alive() or thread.ident is None:
        return False
    thread_id = ctypes.c_ulong(thread.ident)
    result = ctypes.pythonapi.PyThreadState_SetAsyncExc(thread_id, ctypes.py_object(exception_type))
    if result == 0:
        return False
    if result > 1:
        ctypes.pythonapi.PyThreadState_SetAsyncExc(thread_id, None)
        return False
    return True


def _path_hash(path: Path, row: dict[str, object] | None = None) -> str:
    if row is not None:
        return (
            f"{row.get('parquet_path', path)}|{row.get('processed_at', '')}|"
            f"{row.get('row_count', '')}|{row.get('model_count', '')}"
        )
    try:
        stat = path.stat()
        return f"{path}|{stat.st_mtime_ns}|{stat.st_size}"
    except Exception:
        return str(path)


def _next_available_path(path: Path) -> Path:
    if not path.exists():
        return path
    for index in range(1, 1000):
        candidate = path.with_name(f"{path.stem}_{index:03d}{path.suffix}")
        if not candidate.exists():
            return candidate
    raise FileExistsError(f"Could not find available output path for {path}")


def _format_elapsed_seconds(seconds: float) -> str:
    if seconds < 60:
        return f"{seconds:.1f}초"
    minutes = int(seconds // 60)
    remain_seconds = seconds - (minutes * 60)
    return f"{minutes}분 {remain_seconds:.1f}초"


class VIReportGUI:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("VI 추천 / 집계 / 대시보드")
        self.root.geometry("1180x920")

        self.log_queue: queue.Queue[str] = queue.Queue()
        self.worker_thread: threading.Thread | None = None
        self.cancel_requested = threading.Event()
        self.review_continue_event = threading.Event()
        self.review_stop_requested = threading.Event()
        self._run_token_counter = 0
        self._active_run_token: int | None = None
        self.review_waiting = False
        self.review_resume_config: dict[str, object] | None = None
        self.review_resume_workbook_path: Path | None = None

        self.path_store: dict[str, Path | None] = {}

        default_scenario = default_scenario_path()
        self.workflow_var = tk.StringVar(value="recommend_and_report")
        self.scenario_path_var = tk.StringVar(value=default_scenario.name if default_scenario else "")
        self.recommend_output_var = tk.StringVar()
        self.report_output_var = tk.StringVar()
        self.dashboard_report_path_var = tk.StringVar()
        self.status_var = tk.StringVar(value="대기 중")

        self.convert_file_var = tk.StringVar()
        self.convert_kind_var = tk.StringVar(value="historical")
        self.convert_snapshot_name_var = tk.StringVar()
        self.convert_status_var = tk.StringVar(value="")

        self.base_source_var = tk.StringVar(value="parquet")
        self.base_parquet_file_var = tk.StringVar()
        self.base_file_var = tk.StringVar()
        self.base_snapshot_name_var = tk.StringVar()

        self.current_source_var = tk.StringVar(value="parquet")
        self.current_parquet_file_var = tk.StringVar()
        self.current_file_var = tk.StringVar()
        self.current_snapshot_name_var = tk.StringVar()

        if default_scenario:
            self.path_store["scenario_path"] = Path(default_scenario).expanduser().resolve()
        else:
            self.path_store["scenario_path"] = None

        self._build_layout()
        self._refresh_default_paths(force=True)
        self._on_workflow_changed()
        self.root.after(150, self._drain_log_queue)

    def _build_layout(self) -> None:
        outer = ttk.Frame(self.root)
        outer.pack(fill="both", expand=True)
        outer.columnconfigure(0, weight=1)
        outer.rowconfigure(0, weight=1)

        canvas = tk.Canvas(outer, highlightthickness=0)
        scrollbar = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
        container = ttk.Frame(canvas, padding=12)

        canvas.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")
        canvas.configure(yscrollcommand=scrollbar.set)

        self._container_window_id = canvas.create_window((0, 0), window=container, anchor="nw")
        self._scroll_canvas = canvas

        container.bind("<Configure>", self._on_container_configure)
        canvas.bind("<Configure>", self._on_canvas_configure)
        canvas.bind_all("<MouseWheel>", self._on_mousewheel)

        container.columnconfigure(0, weight=1)
        container.rowconfigure(1, weight=1)

        ttk.Label(container, text="VI Workflow", font=("Malgun Gothic", 16, "bold")).grid(row=0, column=0, sticky="w")
        notebook = ttk.Notebook(container)
        notebook.grid(row=1, column=0, sticky="nsew", pady=(10, 0))

        convert_tab = ttk.Frame(notebook, padding=14)
        workflow_tab = ttk.Frame(notebook, padding=14)
        notebook.add(convert_tab, text="BOM → Parquet 변환")
        notebook.add(workflow_tab, text="추천/집계/대시보드")

        self._build_convert_tab(convert_tab)
        self._build_workflow_tab(workflow_tab)

    def _on_container_configure(self, _event: tk.Event) -> None:
        if hasattr(self, "_scroll_canvas"):
            self._scroll_canvas.configure(scrollregion=self._scroll_canvas.bbox("all"))

    def _on_canvas_configure(self, event: tk.Event) -> None:
        if hasattr(self, "_scroll_canvas") and hasattr(self, "_container_window_id"):
            self._scroll_canvas.itemconfigure(self._container_window_id, width=event.width)

    def _on_mousewheel(self, event: tk.Event) -> None:
        if not hasattr(self, "_scroll_canvas"):
            return
        delta = 0
        if getattr(event, "delta", 0):
            delta = -int(event.delta / 120)
        if delta:
            self._scroll_canvas.yview_scroll(delta, "units")

    def _set_selected_path(self, key: str, variable: tk.StringVar, path: Path | None) -> None:
        resolved = Path(path).expanduser().resolve() if path is not None else None
        self.path_store[key] = resolved
        variable.set(resolved.name if resolved is not None else "")

    def _get_selected_path(self, key: str) -> Path | None:
        return self.path_store.get(key)

    def _clear_selected_path(self, key: str, variable: tk.StringVar) -> None:
        self.path_store[key] = None
        variable.set("")

    def _build_convert_tab(self, parent: ttk.Frame) -> None:
        parent.columnconfigure(1, weight=1)
        row = 0
        self._add_entry(parent, row, "BOM Excel 파일", self.convert_file_var, self._browse_convert_file)
        row += 1
        ttk.Label(parent, text="Parquet 구분").grid(row=row, column=0, sticky="w", padx=(0, 10), pady=6)
        kind_frame = ttk.Frame(parent)
        kind_frame.grid(row=row, column=1, columnspan=2, sticky="w", pady=6)
        ttk.Radiobutton(kind_frame, text="Base BOM용", value="historical", variable=self.convert_kind_var).pack(side="left")
        ttk.Radiobutton(kind_frame, text="Current BOM용", value="current_cache", variable=self.convert_kind_var).pack(side="left", padx=(18, 0))
        row += 1
        self._add_plain_entry(parent, row, "저장할 Parquet 파일 이름", self.convert_snapshot_name_var)
        row += 1
        self.convert_button = ttk.Button(parent, text="변환 실행", command=self._convert_clicked)
        self.convert_button.grid(row=row, column=2, sticky="e", pady=8)
        row += 1
        ttk.Label(parent, textvariable=self.convert_status_var, wraplength=920).grid(
            row=row, column=0, columnspan=3, sticky="ew", pady=(8, 0)
        )

    def _build_workflow_tab(self, parent: ttk.Frame) -> None:
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(4, weight=1)

        mode_frame = ttk.LabelFrame(parent, text="실행 모드", padding=10)
        mode_frame.grid(row=0, column=0, sticky="ew")
        ttk.Radiobutton(mode_frame, text="추천부터 집계까지", value="recommend_and_report", variable=self.workflow_var, command=self._on_workflow_changed).pack(anchor="w")
        ttk.Radiobutton(mode_frame, text="집계만", value="report_only", variable=self.workflow_var, command=self._on_workflow_changed).pack(anchor="w", pady=(6, 0))
        ttk.Radiobutton(mode_frame, text="대시보드만 보기", value="dashboard_only", variable=self.workflow_var, command=self._on_workflow_changed).pack(anchor="w", pady=(6, 0))

        form = ttk.LabelFrame(parent, text="입력값", padding=10)
        form.grid(row=1, column=0, sticky="nsew", pady=(10, 0))
        form.columnconfigure(1, weight=1)
        row = 0
        self.template_button = ttk.Button(form, text="Template", command=self._create_template_clicked)
        self.template_button.grid(row=row, column=2, sticky="e", pady=(0, 6))
        row += 1
        self.scenario_entry, self.scenario_button = self._add_entry(form, row, "시나리오/추천 검토 파일", self.scenario_path_var, self._browse_scenario_file)
        row += 1

        self._build_bom_source_frame(form, row, "Base BOM Source", "base")
        row += 1
        self._build_bom_source_frame(form, row, "Current BOM Source", "current")
        row += 1

        self.recommend_output_entry, self.recommend_output_button = self._add_entry(form, row, "추천 검토 파일", self.recommend_output_var, self._browse_recommend_output)
        row += 1
        self.report_output_entry, self.report_output_button = self._add_entry(form, row, "최종 집계 파일", self.report_output_var, self._browse_report_output)
        row += 1
        self.dashboard_report_entry, self.dashboard_report_button = self._add_entry(form, row, "기존 결과 파일", self.dashboard_report_path_var, self._browse_dashboard_report_file)

        action_frame = ttk.Frame(parent)
        action_frame.grid(row=2, column=0, sticky="ew", pady=(12, 8))
        action_frame.columnconfigure(0, weight=1)
        ttk.Label(action_frame, textvariable=self.status_var).grid(row=0, column=0, sticky="w")
        self.review_continue_button = ttk.Button(
            action_frame,
            text="검토 완료 후 계속",
            command=self._continue_after_review,
            state="disabled",
        )
        self.review_continue_button.grid(row=0, column=1, sticky="e", padx=(0, 8))
        self.stop_button = ttk.Button(action_frame, text="실행 중단", command=self._stop_clicked, state="disabled")
        self.stop_button.grid(row=0, column=2, sticky="e", padx=(0, 8))
        self.run_button = ttk.Button(action_frame, text="실행", command=self._run_clicked)
        self.run_button.grid(row=0, column=3, sticky="e")

        log_frame = ttk.LabelFrame(parent, text="간단 실행 로그", padding=10)
        log_frame.grid(row=4, column=0, sticky="nsew")
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        self.log_text = tk.Text(log_frame, height=10, wrap="word", font=("Consolas", 10))
        self.log_text.grid(row=0, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.log_text.configure(yscrollcommand=scrollbar.set, state="disabled")

    def _build_bom_source_frame(self, parent: ttk.Frame, row: int, title: str, prefix: str) -> None:
        frame = ttk.LabelFrame(parent, text=title, padding=10)
        frame.grid(row=row, column=0, columnspan=3, sticky="ew", pady=(8, 4))
        frame.columnconfigure(1, weight=1)

        source_var = self.base_source_var if prefix == "base" else self.current_source_var
        ttk.Radiobutton(frame, text="Parquet file", value="parquet", variable=source_var, command=self._on_source_changed).grid(row=0, column=0, sticky="w")
        ttk.Radiobutton(frame, text="Excel file", value="excel", variable=source_var, command=self._on_source_changed).grid(row=0, column=1, sticky="w")

        widgets_parquet: list[ttk.Widget] = []
        widgets_excel: list[ttk.Widget] = []
        manual_parquet_var = self.base_parquet_file_var if prefix == "base" else self.current_parquet_file_var
        widgets_parquet.extend(self._add_source_entry(frame, 1, "Parquet 파일", manual_parquet_var, lambda p=prefix: self._browse_parquet_file(p)))
        setattr(self, f"{prefix}_parquet_widgets", widgets_parquet)

        file_var = self.base_file_var if prefix == "base" else self.current_file_var
        name_var = self.base_snapshot_name_var if prefix == "base" else self.current_snapshot_name_var

        widgets_excel.extend(self._add_source_entry(frame, 2, "Excel 파일", file_var, lambda p=prefix: self._browse_source_file(p)))
        widgets_excel.extend(self._add_source_entry(frame, 3, "저장할 Parquet 파일 이름", name_var, None))
        setattr(self, f"{prefix}_excel_widgets", widgets_excel)

    def _add_entry(self, parent: ttk.Widget, row: int, label: str, variable: tk.StringVar, command) -> tuple[ttk.Entry, ttk.Button]:
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", padx=(0, 10), pady=5)
        entry = ttk.Entry(parent, textvariable=variable)
        entry.grid(row=row, column=1, sticky="ew", pady=5)
        entry.bind("<FocusOut>", lambda _event: self._refresh_default_paths(force=False))
        button = ttk.Button(parent, text="찾아보기", command=command)
        button.grid(row=row, column=2, sticky="e", pady=5)
        return entry, button

    def _add_plain_entry(self, parent: ttk.Widget, row: int, label: str, variable: tk.StringVar) -> ttk.Entry:
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", padx=(0, 10), pady=5)
        entry = ttk.Entry(parent, textvariable=variable)
        entry.grid(row=row, column=1, columnspan=2, sticky="ew", pady=5)
        return entry

    def _add_combo(self, parent: ttk.Widget, row: int, label: str, variable: tk.StringVar, values: list[str]) -> ttk.Combobox:
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", padx=(0, 10), pady=5)
        combo = ttk.Combobox(parent, textvariable=variable, values=values, state="readonly")
        combo.grid(row=row, column=1, columnspan=2, sticky="ew", pady=5)
        return combo

    def _add_source_entry(self, parent: ttk.Widget, row: int, label: str, variable: tk.StringVar, command) -> list[ttk.Widget]:
        label_widget = ttk.Label(parent, text=label)
        label_widget.grid(row=row, column=0, sticky="w", padx=(0, 10), pady=5)
        entry = ttk.Entry(parent, textvariable=variable)
        entry.grid(row=row, column=1, sticky="ew", pady=5)
        widgets: list[ttk.Widget] = [label_widget, entry]
        if command is not None:
            button = ttk.Button(parent, text="찾아보기", command=command)
            button.grid(row=row, column=2, sticky="e", pady=5)
            widgets.append(button)
        return widgets

    def _append_to_log_widget(self, widget: tk.Text, message: str) -> None:
        widget.configure(state="normal")
        widget.insert("end", message)
        widget.see("end")
        widget.configure(state="disabled")

    def _append_log(self, message: str) -> None:
        self._append_to_log_widget(self.log_text, message)

    def _clear_logs(self) -> None:
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")

    def _drain_log_queue(self) -> None:
        while True:
            try:
                message = self.log_queue.get_nowait()
            except queue.Empty:
                break
            self._append_log(message)
        self.root.after(150, self._drain_log_queue)

    def _set_status(self, text: str) -> None:
        self.status_var.set(text)

    def _push_log(self, message: str) -> None:
        text = message if message.endswith("\n") else f"{message}\n"
        self.log_queue.put(text)

    def _begin_run_token(self) -> int:
        self._run_token_counter += 1
        self._active_run_token = self._run_token_counter
        return self._run_token_counter

    def _invalidate_active_run(self) -> None:
        self._active_run_token = None

    def _run_token_is_active(self, run_token: int) -> bool:
        return self._active_run_token == run_token

    def _root_after_if_active(self, run_token: int, callback: Callable[[], None]) -> None:
        def wrapped() -> None:
            if not self._run_token_is_active(run_token):
                return
            callback()

        self.root.after(0, wrapped)

    def _reset_execution_controls(self, *, clear_flags: bool = True) -> None:
        self.run_button.configure(state="normal")
        self.stop_button.configure(state="disabled")
        self.review_continue_button.configure(state="disabled")
        self.review_waiting = False
        if clear_flags:
            self.cancel_requested.clear()
            self.review_continue_event.clear()
            self.review_stop_requested.clear()

    def _prepare_execution_controls(self) -> None:
        self.run_button.configure(state="disabled")
        self.stop_button.configure(state="normal")
        self.review_continue_button.configure(state="disabled")
        self.review_waiting = False
        self.cancel_requested.clear()
        self.review_continue_event.clear()
        self.review_stop_requested.clear()

    def _continue_after_review(self) -> None:
        if self.review_waiting:
            self._push_log("추천 파일 검토 완료 요청을 확인했습니다. 다음 단계로 진행합니다.")
            self.review_continue_event.set()
            return
        if self.review_resume_config is None:
            return
        if self.worker_thread and self.worker_thread.is_alive():
            return
        self._push_log("추천 파일 검토 완료 요청을 확인했습니다. 저장된 검토 파일로 최종 집계를 다시 진행합니다.")
        self._start_workflow_thread(self.review_resume_config, clear_logs=False, initial_status="최종 집계 파일 생성 중...")

    def _stop_clicked(self) -> None:
        if self.worker_thread is None or not self.worker_thread.is_alive():
            return
        self.cancel_requested.set()
        self.review_stop_requested.set()
        self._invalidate_active_run()
        self._clear_review_resume_state()
        self._reset_execution_controls(clear_flags=False)
        self.convert_button.configure(state="normal")
        interrupted = _raise_async_exception(self.worker_thread, SystemExit)
        if interrupted:
            self._set_status("대기 중")
            self._push_log("사용자가 실행 즉시 중단을 요청했습니다. 화면을 실행 전 상태로 되돌립니다.")
        else:
            self._set_status("대기 중")
            self._push_log("사용자가 실행 중단을 요청했습니다. 백그라운드 작업 종료를 기다리는 동안 화면은 초기 상태로 복귀합니다.")

    def _check_cancel_requested(self) -> None:
        if self.cancel_requested.is_set():
            raise SystemExit

    def _browse_excel_file(self, title: str, current_value: str) -> str:
        initial = Path(current_value).parent if current_value else sample_files_root()
        return filedialog.askopenfilename(
            title=title,
            initialdir=str(initial),
            filetypes=[("Excel files", "*.xlsx *.xlsm *.xls"), ("All files", "*.*")],
        )

    def _browse_convert_file(self) -> None:
        current = self._get_selected_path("convert_file_path")
        selected = self._browse_excel_file("BOM Excel 파일 선택", str(current) if current else "")
        if selected:
            path = Path(selected)
            self._set_selected_path("convert_file_path", self.convert_file_var, path)
            self.convert_snapshot_name_var.set(path.stem)

    def _browse_source_file(self, prefix: str) -> None:
        variable = self.base_file_var if prefix == "base" else self.current_file_var
        key = f"{prefix}_excel_file_path"
        current = self._get_selected_path(key)
        selected = self._browse_excel_file(f"{prefix.title()} BOM Excel 파일 선택", str(current) if current else "")
        if selected:
            path = Path(selected)
            self._set_selected_path(key, variable, path)
            name_var = self.base_snapshot_name_var if prefix == "base" else self.current_snapshot_name_var
            name_var.set(path.stem)
            self._refresh_default_paths(force=False)

    def _browse_parquet_file(self, prefix: str) -> None:
        variable = self.base_parquet_file_var if prefix == "base" else self.current_parquet_file_var
        key = f"{prefix}_manual_parquet_path"
        current = self._get_selected_path(key)
        initial = str(current.parent if current else sample_files_root())
        selected = filedialog.askopenfilename(
            title=f"{prefix.title()} Parquet 파일 선택",
            initialdir=initial,
            filetypes=[("Parquet files", "*.parquet"), ("All files", "*.*")],
        )
        if selected:
            self._set_selected_path(key, variable, Path(selected))
            self._refresh_default_paths(force=False)

    def _browse_scenario_file(self) -> None:
        current = self._get_selected_path("scenario_path")
        selected = self._browse_excel_file("시나리오/추천 검토 파일 선택", str(current) if current else "")
        if selected:
            self._set_selected_path("scenario_path", self.scenario_path_var, Path(selected))
            self._refresh_default_paths(force=False)

    def _browse_recommend_output(self) -> None:
        default_path = self._current_recommend_output_path()
        selected = filedialog.asksaveasfilename(
            title="추천 검토 파일 저장 위치",
            initialdir=str(default_path.parent),
            initialfile=default_path.name,
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
        )
        if selected:
            self._set_selected_path("recommend_output_path", self.recommend_output_var, Path(selected))

    def _browse_report_output(self) -> None:
        default_path = self._current_report_output_path()
        selected = filedialog.asksaveasfilename(
            title="최종 집계 파일 저장 위치",
            initialdir=str(default_path.parent),
            initialfile=default_path.name,
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
        )
        if selected:
            self._set_selected_path("report_output_path", self.report_output_var, Path(selected))

    def _browse_dashboard_report_file(self) -> None:
        current = self._get_selected_path("dashboard_report_path")
        selected = self._browse_excel_file("기존 결과 파일 선택", str(current) if current else "")
        if selected:
            self._set_selected_path("dashboard_report_path", self.dashboard_report_path_var, Path(selected))

    def _default_template_path(self) -> Path:
        return scenario_templates_root() / "VI_Scenario_Template_Auto.xlsx"

    def _create_template_clicked(self) -> None:
        default_path = self._default_template_path()
        selected = filedialog.asksaveasfilename(
            title="VI Scenario Template 저장 위치",
            initialdir=str(default_path.parent),
            initialfile=default_path.name,
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
        )
        if not selected:
            return

        output_path = Path(selected)
        if output_path.suffix.lower() != ".xlsx":
            output_path = output_path.with_suffix(".xlsx")

        self._write_scenario_template(output_path)
        self._set_selected_path("scenario_path", self.scenario_path_var, output_path)
        self._refresh_default_paths(force=False)
        open_workbook_for_review(output_path)
        messagebox.showinfo(
            "Template 생성 완료",
            f"VI Scenario Template 파일을 만들었습니다.\n\n{output_path}",
        )

    def _write_scenario_template(self, output_path: Path) -> None:
        readme = pd.DataFrame(
            [
                {
                    "항목": "사용 순서",
                    "설명": "1) BasePN_Recommendation_Request 입력  2) Volume 입력  3) 추천 실행  4) BasePN_Recommendation_Result에서 검토 입력  5) 최종 집계 실행",
                },
                {
                    "항목": "필수 시트",
                    "설명": "BasePN_Recommendation_Request, Volume",
                },
                {
                    "항목": "추천 결과 시트",
                    "설명": "BasePN_Recommendation_Result는 추천 실행 후 자동으로 채워집니다. 이후 사용자 검토 입력은 '적용 확정 여부(Y/N)', '확정 New P/N'만 채우면 됩니다.",
                },
                {
                    "항목": "헤더 색상 의미",
                    "설명": "노랑=필수 입력, 파랑=조건 입력, 초록=선택 입력, 주황=검토 입력, 회색=자동 생성",
                },
                {
                    "항목": "주의",
                    "설명": "헤더명은 그대로 두고 데이터만 입력하세요. 빈 조건은 자동으로 제외되고, 채운 조건만 AND로 함께 적용됩니다. Base P/N을 입력하면 해당 품번만 직접 매칭합니다.",
                },
            ]
        )
        input_guide = _build_scenario_input_guide()

        sample_request = _build_template_sample_frame(
            scenario_template_columns["BasePN_Recommendation_Request"],
            [
                {
                    "vi_item_name": "예시 항목",
                    "enabled": "Y",
                    "subsidiary": "ALL",
                    "base_part_no": "",
                    "new_part_no": "",
                    "model_type": "",
                    "indoor_tool": "",
                    "desc_match_mode": "OR",
                    "desc_contains_any": "COMPRESSOR",
                    "spec_contains_any": "",
                    "parent_assy_desc_contains_any": "",
                    "note": "예시 요청 행",
                }
            ],
        )

        sample_volume = _build_template_sample_frame(
            scenario_template_columns["Volume"],
            [
                {
                    "month": "2026-05",
                    "subsidiary": "BRAZIL",
                    "model_suffix": "EXAMPLEMODEL01",
                    "production_qty": 100,
                }
            ],
        )

        empty_recommendation_result = pd.DataFrame(columns=scenario_template_columns["BasePN_Recommendation_Result"])

        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            readme.to_excel(writer, sheet_name="README", index=False)
            input_guide.to_excel(writer, sheet_name="InputGuide", index=False)
            _friendly_template_frame(sample_request, "BasePN_Recommendation_Request").to_excel(
                writer,
                sheet_name="BasePN_Recommendation_Request",
                index=False,
            )
            _friendly_template_frame(empty_recommendation_result, "BasePN_Recommendation_Result").to_excel(
                writer,
                sheet_name="BasePN_Recommendation_Result",
                index=False,
            )
            _friendly_template_frame(sample_volume, "Volume").to_excel(
                writer,
                sheet_name="Volume",
                index=False,
            )
            for worksheet in writer.book.worksheets:
                worksheet.freeze_panes = "A2"
                if worksheet.title in scenario_template_columns:
                    _apply_template_header_styles(worksheet, worksheet.title)
                _auto_fit_openpyxl_worksheet(worksheet)

    def _refresh_snapshots(self) -> None:
        self._refresh_default_paths(force=False)

    def _on_source_changed(self) -> None:
        for prefix in ["base", "current"]:
            source_var = self.base_source_var if prefix == "base" else self.current_source_var
            parquet_widgets = getattr(self, f"{prefix}_parquet_widgets")
            excel_widgets = getattr(self, f"{prefix}_excel_widgets")
            parquet_mode = source_var.get() == "parquet"
            for widget in parquet_widgets:
                widget.configure(state="normal" if parquet_mode else "disabled")
            for widget in excel_widgets:
                widget.configure(state="disabled" if parquet_mode else "normal")
        self._refresh_default_paths(force=False)

    def _on_workflow_changed(self) -> None:
        workflow = self.workflow_var.get()
        dashboard_only = workflow == "dashboard_only"
        recommend_mode = workflow == "recommend_and_report"
        calculation_state = "disabled" if dashboard_only else "normal"
        for widget in [
            self.scenario_entry,
            self.scenario_button,
            self.report_output_entry,
            self.report_output_button,
        ]:
            widget.configure(state=calculation_state)
        recommend_state = "normal" if recommend_mode else "disabled"
        self.recommend_output_entry.configure(state=recommend_state)
        self.recommend_output_button.configure(state=recommend_state)
        dashboard_state = "normal" if dashboard_only else "disabled"
        self.dashboard_report_entry.configure(state=dashboard_state)
        self.dashboard_report_button.configure(state=dashboard_state)
        self._on_source_changed()

    def _current_recommend_output_path(self) -> Path:
        fallback = default_scenario_path() or (scenario_templates_root() / "scenario.xlsx")
        scenario_path = self._get_selected_path("scenario_path") or fallback
        base_year, base_month = self._resolve_base_month_for_default()
        return _next_available_path(default_recommendation_output_path(scenario_path, base_year, base_month))

    def _current_report_output_path(self) -> Path:
        base_year, base_month = self._resolve_base_month_for_default()
        current_year, current_month = self._resolve_current_month_for_default(base_year, base_month)
        return _next_available_path(default_output_path(base_year, base_month, current_year, current_month))

    def _resolve_base_month_for_default(self) -> tuple[int, int]:
        if self.base_source_var.get() == "parquet":
            manual_parquet_path = self._get_selected_path("base_manual_parquet_path")
            if manual_parquet_path is not None:
                try:
                    info = inspect_parquet_snapshot(manual_parquet_path)
                    return int(info["bom_year"]), int(info["bom_month"])
                except Exception:
                    pass
        return 2025, 1

    def _resolve_current_month_for_default(self, base_year: int, base_month: int) -> tuple[int, int]:
        if self.current_source_var.get() == "parquet":
            manual_parquet_path = self._get_selected_path("current_manual_parquet_path")
            if manual_parquet_path is not None:
                try:
                    info = inspect_parquet_snapshot(manual_parquet_path)
                    return int(info["bom_year"]), int(info["bom_month"])
                except Exception:
                    pass
        return base_year, base_month

    def _refresh_default_paths(self, force: bool = False) -> None:
        if force or not self.recommend_output_var.get().strip():
            try:
                self._set_selected_path("recommend_output_path", self.recommend_output_var, self._current_recommend_output_path())
            except Exception:
                pass
        if force or not self.report_output_var.get().strip():
            try:
                self._set_selected_path("report_output_path", self.report_output_var, self._current_report_output_path())
            except Exception:
                pass

    def _convert_clicked(self) -> None:
        if self.worker_thread and self.worker_thread.is_alive():
            messagebox.showinfo("실행 중", "이미 실행 중입니다. 완료될 때까지 기다려 주세요.")
            return
        file_path = self._get_selected_path("convert_file_path")
        if file_path is None:
            messagebox.showerror("입력 오류", "BOM Excel 파일을 선택해 주세요.")
            return
        if not file_path.exists():
            messagebox.showerror("입력 오류", f"BOM Excel 파일이 없습니다.\n{file_path}")
            return
        snapshot_name = self.convert_snapshot_name_var.get().strip() or file_path.stem
        config = {
            "file_path": file_path,
            "snapshot_kind": self.convert_kind_var.get(),
            "snapshot_name": snapshot_name,
        }
        self._clear_logs()
        self.convert_button.configure(state="disabled")
        self._prepare_execution_controls()
        self._set_status("BOM 변환 중...")
        self._push_log("BOM Excel 파일을 Parquet로 변환하는 작업을 시작했습니다.")
        run_token = self._begin_run_token()
        self.worker_thread = threading.Thread(target=self._run_conversion, args=(config, run_token), daemon=True)
        self.worker_thread.start()

    def _run_conversion(self, config: dict[str, object], run_token: int) -> None:
        writer = QueueWriter(self.log_queue)
        try:
            self._check_cancel_requested()
            with contextlib.redirect_stdout(writer), contextlib.redirect_stderr(writer):
                result = self._convert_excel_to_parquet(config)
            self._check_cancel_requested()
            parquet_path = Path(result["parquet_path"])
            row_count = len(result["data"].index)
            model_count = int(result["data"]["model_suffix"].nunique())
            message = f"저장 완료: {parquet_path}\nrows={row_count:,}, models={model_count:,}"
            self._root_after_if_active(run_token, lambda: self.convert_status_var.set(message))
            self._root_after_if_active(run_token, self._refresh_snapshots)
            self._root_after_if_active(run_token, lambda: self.convert_button.configure(state="normal"))
            self._root_after_if_active(run_token, self._reset_execution_controls)
            self._root_after_if_active(run_token, lambda: self._set_status("BOM 변환 완료"))
        except SystemExit:
            self.log_queue.put("사용자가 실행을 즉시 중단했습니다.\n")
            self.root.after(0, lambda: self.convert_button.configure(state="normal"))
        except Exception as exc:
            self.log_queue.put(f"{exc}\n\n{traceback.format_exc()}")
            self._root_after_if_active(run_token, lambda: self.convert_button.configure(state="normal"))
            self._root_after_if_active(run_token, self._reset_execution_controls)
            self._root_after_if_active(run_token, lambda: self._finish_error(str(exc)))
        finally:
            if self.worker_thread is threading.current_thread():
                self.worker_thread = None
            if self._run_token_is_active(run_token):
                self._invalidate_active_run()

    def _convert_excel_to_parquet(self, config: dict[str, object]) -> dict[str, Any]:
        kind = str(config["snapshot_kind"])
        file_path = Path(config["file_path"])
        kind_label = "기준월 이력" if kind == "historical" else "당월 캐시"
        print(f"[1/3] Excel -> Parquet 변환 시작: {kind_label}", flush=True)
        print(f" - 파일: {file_path}", flush=True)
        print(f" - Parquet 파일 이름: {config['snapshot_name']}", flush=True)
        print(" - Excel 엔진: calamine", flush=True)
        if kind == "historical":
            result = save_historical_snapshot_from_excel(
                file_path=file_path,
                snapshot_name=str(config["snapshot_name"]),
                excel_engine="calamine",
            )
        else:
            result = save_current_snapshot_from_excel(
                file_path=file_path,
                snapshot_name=str(config["snapshot_name"]),
                excel_engine="calamine",
            )
        print(
            f"[3/3] 변환 완료: {result['parquet_path']} | "
            f"행 수={len(result['data'].index):,} | 모델 수={result['data']['model_suffix'].nunique():,}",
            flush=True,
        )
        return result

    def _validate_inputs(self) -> dict[str, object] | None:
        workflow = self.workflow_var.get()
        if workflow == "dashboard_only":
            report_path = self._get_selected_path("dashboard_report_path")
            if report_path is None:
                messagebox.showerror("입력 오류", "기존 결과 파일을 선택해 주세요.")
                return None
            if not report_path.exists():
                messagebox.showerror("입력 오류", f"기존 결과 파일이 없습니다.\n{report_path}")
                return None
            return {"workflow": workflow, "dashboard_report_path": report_path}

        scenario_path = self._get_selected_path("scenario_path")
        if scenario_path is None:
            messagebox.showerror("입력 오류", "시나리오/추천 검토 파일을 선택해 주세요.")
            return None
        if not scenario_path.exists():
            messagebox.showerror("입력 오류", f"시나리오/추천 검토 파일이 없습니다.\n{scenario_path}")
            return None

        base_config = self._validate_bom_source("base")
        current_config = self._validate_bom_source("current")
        if base_config is None or current_config is None:
            return None

        report_output_path = self._get_selected_path("report_output_path")
        if report_output_path is None:
            messagebox.showerror("입력 오류", "최종 집계 파일 저장 위치를 지정해 주세요.")
            return None

        recommend_output_path = None
        if workflow == "recommend_and_report":
            recommend_output_path = self._get_selected_path("recommend_output_path")
            if recommend_output_path is None:
                messagebox.showerror("입력 오류", "추천 검토 파일 저장 위치를 지정해 주세요.")
                return None
            recommend_output_path = _next_available_path(recommend_output_path)
            self._set_selected_path("recommend_output_path", self.recommend_output_var, recommend_output_path)

        report_output_path = _next_available_path(report_output_path)
        self._set_selected_path("report_output_path", self.report_output_var, report_output_path)
        return {
            "workflow": workflow,
            "scenario_path": scenario_path,
            "recommend_output_path": recommend_output_path,
            "report_output_path": report_output_path,
            "base": base_config,
            "current": current_config,
        }

    def _validate_bom_source(self, prefix: str) -> dict[str, object] | None:
        source_var = self.base_source_var if prefix == "base" else self.current_source_var
        if source_var.get() == "parquet":
            manual_key = f"{prefix}_manual_parquet_path"
            manual_path = self._get_selected_path(manual_key)
            if manual_path is None:
                messagebox.showerror("입력 오류", f"{prefix.title()} Parquet 파일을 선택해 주세요.")
                return None
            if not manual_path.exists():
                messagebox.showerror("입력 오류", f"{prefix.title()} Parquet 파일이 없습니다.\n{manual_path}")
                return None
            try:
                info = inspect_parquet_snapshot(manual_path)
            except Exception as exc:
                messagebox.showerror("입력 오류", f"{prefix.title()} Parquet 파일 정보를 읽지 못했습니다.\n{exc}")
                return None
            return {
                "source": "parquet",
                "snapshot": info,
                "year": int(info["bom_year"]),
                "month": int(info["bom_month"]),
                "snapshot_path": manual_path,
                "source_hash": _path_hash(manual_path),
            }

        file_var = self.base_file_var if prefix == "base" else self.current_file_var
        name_var = self.base_snapshot_name_var if prefix == "base" else self.current_snapshot_name_var
        file_path = self._get_selected_path(f"{prefix}_excel_file_path")
        if file_path is None:
            messagebox.showerror("입력 오류", f"{prefix.title()} BOM Excel 파일을 선택해 주세요.")
            return None
        if not file_path.exists():
            messagebox.showerror("입력 오류", f"{prefix.title()} BOM Excel 파일이 없습니다.\n{file_path}")
            return None
        return {
            "source": "excel",
            "file_path": file_path,
            "snapshot_name": name_var.get().strip() or file_path.stem,
        }

    def _run_clicked(self) -> None:
        if self.worker_thread and self.worker_thread.is_alive():
            messagebox.showinfo("실행 중", "이미 실행 중입니다. 완료될 때까지 기다려 주세요.")
            return
        config = self._validate_inputs()
        if config is None:
            return
        self._start_workflow_thread(config, clear_logs=True, initial_status="실행 중...")

    def _start_workflow_thread(
        self,
        config: dict[str, object],
        *,
        clear_logs: bool,
        initial_status: str,
    ) -> None:
        if clear_logs:
            self._clear_logs()
            self._push_log("VI 추천/집계 작업을 시작했습니다.")
        self._prepare_execution_controls()
        self._set_status(initial_status)
        run_token = self._begin_run_token()
        self.worker_thread = threading.Thread(target=self._run_workflow, args=(config, run_token), daemon=True)
        self.worker_thread.start()

    def _store_review_resume_state(self, config: dict[str, object], workbook_path: Path) -> None:
        resume_config = dict(config)
        resume_config["workflow"] = "report_only"
        resume_config["scenario_path"] = workbook_path
        self.review_resume_config = resume_config
        self.review_resume_workbook_path = workbook_path

    def _clear_review_resume_state(self) -> None:
        self.review_resume_config = None
        self.review_resume_workbook_path = None

    def _activate_review_resume_mode(self, workbook_path: Path, message: str) -> None:
        self.worker_thread = None
        self.review_waiting = False
        self.cancel_requested.clear()
        self.review_continue_event.clear()
        self.review_stop_requested.clear()
        self.workflow_var.set("report_only")
        self._on_workflow_changed()
        self._set_selected_path("scenario_path", self.scenario_path_var, workbook_path)
        self.run_button.configure(state="normal")
        self.stop_button.configure(state="disabled")
        self.review_continue_button.configure(state="normal")
        self._set_status("추천 검토 파일 수정 후 집계만 다시 진행할 수 있습니다.")
        self._push_log(message)

    def _wait_for_review_continue(self, workbook_path: Path) -> str:
        opened_event = threading.Event()

        def prompt() -> None:
            try:
                self.review_waiting = True
                self.review_continue_button.configure(state="normal")
                self._set_status("추천 파일을 검토한 뒤 '검토 완료 후 계속'을 눌러 주세요.")
                open_workbook_for_review(workbook_path)
            finally:
                opened_event.set()

        self.root.after(0, prompt)
        opened_event.wait()

        while True:
            if self.review_continue_event.wait(0.2):
                self.root.after(0, lambda: self.review_continue_button.configure(state="disabled"))
                self.review_waiting = False
                return "continue"
            if self.review_stop_requested.is_set():
                self.review_waiting = False
                return "resume_pending"
            if self.cancel_requested.is_set():
                self.root.after(0, lambda: self.review_continue_button.configure(state="disabled"))
                self.review_waiting = False
                return "cancelled"

    def _finish_success(self, report_output_path: Path, final_rule_count: int) -> None:
        self._reset_execution_controls()
        self._set_status("완료")
        if final_rule_count == 0:
            messagebox.showwarning(
                "대시보드 실행 건너뜀",
                "대상 규칙 수가 0입니다. 추천 결과의 사용 여부 및 New P/N 입력 상태를 확인하세요.",
            )
            return
        launch_streamlit_dashboard(report_output_path)

    def _finish_dashboard_only(self, report_path: Path) -> None:
        self._reset_execution_controls()
        self._set_status("대시보드 실행 완료")
        launch_streamlit_dashboard(report_path)

    def _finish_error(self, message: str) -> None:
        self._reset_execution_controls()
        self._set_status("오류 발생 - 다시 실행 가능")
        messagebox.showerror("실행 실패", f"{message}\n\n입력값과 로그를 확인한 뒤 다시 실행할 수 있습니다.")

    def _prepare_bom_for_run(self, prefix: str, source_config: dict[str, object]) -> dict[str, object]:
        source_label = "기준월" if prefix == "base" else "당월"
        if source_config["source"] == "parquet":
            snapshot = source_config["snapshot"]
            self._push_log(f"[준비] {source_label} Parquet 선택: {snapshot['label']}")
            return source_config

        kind = "historical" if prefix == "base" else "current_cache"
        conversion_config = {
            "file_path": source_config["file_path"],
            "snapshot_kind": kind,
            "snapshot_name": source_config["snapshot_name"],
        }
        self._push_log(f"[준비] {source_label} BOM을 Excel에서 준비하고 있습니다.")
        result = self._convert_excel_to_parquet(conversion_config)
        parquet_path = Path(result["parquet_path"])
        prepared = dict(source_config)
        prepared.update(
            {
                "source": "excel",
                "snapshot_path": parquet_path,
                "source_hash": _path_hash(parquet_path),
                "year": int(result["bom_year"]),
                "month": int(result["bom_month"]),
                "row_count": len(result["data"].index),
                "model_count": int(result["data"]["model_suffix"].nunique()),
            }
        )
        self._push_log(f"[준비] {source_label} BOM 준비 완료: {parquet_path.name}")
        self.root.after(0, self._refresh_snapshots)
        return prepared

    def _run_workflow(self, config: dict[str, object], run_token: int) -> None:
        writer = QueueWriter(self.log_queue)
        try:
            with contextlib.redirect_stdout(writer), contextlib.redirect_stderr(writer):
                workflow = str(config["workflow"])
                self._check_cancel_requested()
                if workflow == "dashboard_only":
                    self._push_log("[1/1] 대시보드 전용 모드로 결과 파일을 엽니다.")
                    self._root_after_if_active(run_token, lambda: self._finish_dashboard_only(config["dashboard_report_path"]))
                    return

                if workflow == "recommend_and_report":
                    self._push_log("[1/4] 추천 파일을 생성합니다.")
                    self._root_after_if_active(run_token, lambda: self._set_status("추천 파일 생성 중..."))
                else:
                    self._push_log("[1/3] 추천 없이 최종 집계를 진행합니다.")
                    self._root_after_if_active(run_token, lambda: self._set_status("집계 준비 중..."))
                base = self._prepare_bom_for_run("base", config["base"])
                current = self._prepare_bom_for_run("current", config["current"])
                self._check_cancel_requested()

                scenario_path = Path(config["scenario_path"])
                scenario_hash = scenario_file_hash(scenario_path)
                base_year = int(base["year"])
                base_month = int(base["month"])
                current_year = int(current["year"])
                current_month = int(current["month"])
                base_snapshot_path = Path(base["snapshot_path"])
                current_snapshot_path = Path(current["snapshot_path"])
                base_source_hash = str(base["source_hash"])
                current_source_hash = str(current["source_hash"])
                force = True

                if workflow == "recommend_and_report":
                    recommend_output_path = Path(config["recommend_output_path"])
                    recommendation_started_at = time.perf_counter()
                    run_recommendation_flow(
                        scenario_path=scenario_path,
                        scenario_hash=scenario_hash,
                        base_year=base_year,
                        base_month=base_month,
                        history_root=None,
                        recommend_output_path=recommend_output_path,
                        force=force,
                        base_snapshot_kind="historical",
                        base_source_hash=base_source_hash,
                        base_snapshot_path=base_snapshot_path,
                    )
                    self._check_cancel_requested()
                    recommendation_elapsed = time.perf_counter() - recommendation_started_at
                    self._push_log(
                        f"[2/4] 추천 파일 생성 완료: {recommend_output_path.name} | 추천 시간 {_format_elapsed_seconds(recommendation_elapsed)}"
                    )
                    self._push_log("[3/4] 추천 파일을 자동으로 열었습니다. 수정 후 저장/닫고 '검토 완료 후 계속'을 눌러 주세요.")
                    self._root_after_if_active(run_token, lambda: self._set_status("추천 파일 검토를 기다리는 중..."))
                    self._store_review_resume_state(config, recommend_output_path)
                    review_state = self._wait_for_review_continue(recommend_output_path)
                    if review_state == "resume_pending":
                        self._root_after_if_active(
                            run_token,
                            lambda: self._activate_review_resume_mode(
                                recommend_output_path,
                                "추천 검토를 중단한 상태로 저장했습니다. 같은 파일을 다시 열어 New P/N을 채운 뒤 '검토 완료 후 계속' 또는 '실행'으로 집계만 이어가세요.",
                            ),
                        )
                        return
                    if review_state == "cancelled":
                        raise SystemExit
                    if review_state != "continue":
                        self._root_after_if_active(run_token, lambda: self._finish_error("사용자가 추천 검토 단계에서 실행을 중단했습니다."))
                        return
                    scenario_path = recommend_output_path
                    scenario_hash = scenario_file_hash(scenario_path)
                    self._check_cancel_requested()
                    self._push_log("[4/4] 최종 집계를 시작합니다.")
                    self._root_after_if_active(run_token, lambda: self._set_status("최종 집계 파일 생성 중..."))
                else:
                    self._push_log("[2/3] 최종 집계를 시작합니다.")
                    self._root_after_if_active(run_token, lambda: self._set_status("최종 집계 파일 생성 중..."))

                report_output_path = Path(config["report_output_path"])
                report_started_at = time.perf_counter()
                try:
                    report, _cache_key = run_report_flow(
                        scenario_path=scenario_path,
                        scenario_hash=scenario_hash,
                        base_year=base_year,
                        base_month=base_month,
                        current_year=current_year,
                        current_month=current_month,
                        history_root=None,
                        detail=True,
                        output_path=report_output_path,
                        force=force,
                        base_snapshot_kind="historical",
                        base_source_hash=base_source_hash,
                        base_snapshot_path=base_snapshot_path,
                        current_snapshot_kind="current_cache",
                        current_source_hash=current_source_hash,
                        use_cache=True,
                        current_snapshot_path=current_snapshot_path,
                    )
                except ValueError as exc:
                    if "사용 여부가 Y인 행에 적용 New P/N이 비어 있습니다." in str(exc) and self.review_resume_workbook_path is not None:
                        workbook_path = self.review_resume_workbook_path
                        self._root_after_if_active(
                            run_token,
                            lambda: self._activate_review_resume_mode(
                                workbook_path,
                                "추천 검토 파일에 New P/N이 비어 있는 행이 있어 집계를 멈췄습니다. 같은 파일을 다시 열어 수정하고 저장한 뒤 '검토 완료 후 계속' 또는 '실행'으로 집계만 다시 진행하세요.",
                            ),
                        )
                        return
                    raise
                self._check_cancel_requested()
                report_elapsed = time.perf_counter() - report_started_at
                final_rule_count = len(report.get("FinalVITargetRule", []))
                self._push_log(
                    f"[완료] 최종 집계 완료: {report_output_path.name} | 대상 규칙 수 {final_rule_count} | 집계 시간 {_format_elapsed_seconds(report_elapsed)}"
                )
                if final_rule_count == 0:
                    self._push_log("[Dashboard] 대상 규칙 수가 0이라 자동 실행을 건너뜁니다.")
                self._clear_review_resume_state()
            self._root_after_if_active(run_token, lambda: self._finish_success(report_output_path, final_rule_count))
        except SystemExit:
            self.log_queue.put("사용자가 실행을 즉시 중단했습니다.\n")
        except Exception as exc:
            self.log_queue.put(f"{exc}\n\n{traceback.format_exc()}")
            self._root_after_if_active(run_token, lambda: self._finish_error(str(exc)))
        finally:
            if self.worker_thread is threading.current_thread():
                self.worker_thread = None
            if self._run_token_is_active(run_token):
                self._invalidate_active_run()


def run_frozen_dashboard(report_file: str, port: int) -> int:
    app_root = Path(sys.executable).resolve().parent
    dashboard_script = app_root / "streamlit_vi_dashboard.py"
    report_path = Path(report_file).expanduser().resolve()
    stdout_log_path = _dashboard_stdout_log_path()
    stderr_log_path = _dashboard_stderr_log_path()
    stdout_log_path.parent.mkdir(parents=True, exist_ok=True)
    stderr_log_path.parent.mkdir(parents=True, exist_ok=True)

    with stdout_log_path.open("w", encoding="utf-8") as stdout_handle, stderr_log_path.open("w", encoding="utf-8") as stderr_handle:
        try:
            if str(app_root) not in sys.path:
                sys.path.insert(0, str(app_root))
            if not dashboard_script.exists():
                raise FileNotFoundError(f"Dashboard script not found: {dashboard_script}")
            os.chdir(str(app_root))
            _write_dashboard_launch_log(
                "[Dashboard Launch]\n"
                f"python={Path(sys.executable).resolve()}\n"
                f"script={dashboard_script.resolve()}\n"
                f"cwd={app_root}\n"
                f"report={report_path}\n"
                f"port={int(port)}\n"
            )
            os.environ.setdefault("STREAMLIT_BROWSER_GATHER_USAGE_STATS", "false")
            os.environ.setdefault("STREAMLIT_SERVER_HEADLESS", "true")
            ensure_runtime_package("streamlit", "streamlit")
            ensure_runtime_package("plotly", "plotly")
            import streamlit.web.cli as stcli

            with contextlib.redirect_stdout(stdout_handle), contextlib.redirect_stderr(stderr_handle):
                sys.argv = [
                    "streamlit",
                    "run",
                    str(dashboard_script),
                    "--global.developmentMode",
                    "false",
                    "--server.address",
                    "localhost",
                    "--server.port",
                    str(int(port)),
                    "--server.headless",
                    "true",
                    "--browser.gatherUsageStats",
                    "false",
                    "--",
                    "--report-file",
                    str(report_path),
                ]
                return int(stcli.main() or 0)
        except Exception:
            stderr_handle.write(traceback.format_exc())
            stderr_handle.flush()
            raise


def main() -> None:
    if "--launch-dashboard" in sys.argv:
        report_file = _extract_cli_arg_value("--report-file")
        port_text = _extract_cli_arg_value("--port") or "8501"
        if not report_file:
            raise SystemExit("Missing --report-file for dashboard launch.")
        raise SystemExit(run_frozen_dashboard(report_file, int(port_text)))
    root = tk.Tk()
    try:
        load_vi_report_api()
    except Exception as exc:
        messagebox.showerror(
            "필수 패키지 설치 실패",
            "VI 실행에 필요한 Python 패키지를 준비하지 못했습니다.\n\n"
            f"현재 Python:\n{sys.executable}\n\n"
            f"{sys.executable} -m pip install -r requirements.txt\n\n오류 내용:\n{exc}",
        )
        root.destroy()
        return
    VIReportGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
