from __future__ import annotations

import argparse
import hashlib
import json
import os
import shutil
import sys
import time
from datetime import datetime
from pathlib import Path
from typing import Callable

import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows

try:
    import duckdb
except ImportError as exc:  # pragma: no cover
    duckdb = None
    DUCKDB_IMPORT_ERROR = exc
else:  # pragma: no cover
    DUCKDB_IMPORT_ERROR = None

bootstrap_dir = Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parent
parent_dir = bootstrap_dir.parent
for path in [bootstrap_dir, parent_dir]:
    if str(path) not in sys.path:
        sys.path.insert(0, str(path))

from runtime_paths import (
    application_root,
    bom_lake_root as runtime_bom_lake_root,
    dashboard_output_root as runtime_dashboard_output_root,
    output_root as runtime_output_root,
    portable_path_token,
    recommendation_cache_root,
    report_cache_root,
    sample_files_root,
    scenario_templates_root,
)

try:
    from bom_lake import cache_current_month, discover_current_snapshots, save_current_snapshot_from_excel
    import vi_calculation_engine as _vi_engine
except ModuleNotFoundError:
    from compat_imports.bom_lake import cache_current_month, discover_current_snapshots, save_current_snapshot_from_excel  # type: ignore[reportMissingImports]
    from compat_imports import vi_calculation_engine as _vi_engine  # type: ignore[reportMissingImports]


RECOMMENDATION_RESULT_COLUMNS = _vi_engine.RECOMMENDATION_RESULT_COLUMNS
build_rules_from_recommendation_result = _vi_engine.build_rules_from_recommendation_result
build_base_pn_recommendations = _vi_engine.build_base_pn_recommendations
build_vi_scenario_outputs = _vi_engine.build_vi_scenario_outputs
find_invalid_confirmed_recommendation_rows = _vi_engine.find_invalid_confirmed_recommendation_rows
load_scenario_workbook = _vi_engine.load_scenario_workbook
save_vi_scenario_workbook = _vi_engine.save_vi_scenario_workbook


def _validate_required_request_rows(scenario_inputs) -> None:
    request = scenario_inputs.recommendation_request.copy()
    if request.empty:
        raise ValueError("BasePN_Recommendation_Request 시트에 입력된 행이 없습니다.")
    usable = request.copy()
    key_columns = ["vi_item_name", "subsidiary", "base_part_no", "new_part_no", "desc_contains_any", "spec_contains_any", "parent_assy_desc_contains_any"]
    row_mask = pd.Series([False] * len(usable.index), index=usable.index)
    for column in key_columns:
        if column in usable.columns:
            row_mask = row_mask | usable[column].fillna("").astype(str).str.strip().ne("")
    usable = usable[row_mask].copy()
    required_columns = ["vi_item_name", "base_part_no", "new_part_no"]
    errors: list[str] = []
    for row_number, (_, row) in enumerate(usable.iterrows(), start=2):
        for column in required_columns:
            if str(row.get(column, "")).strip():
                continue
            errors.append(f"sheet=BasePN_Recommendation_Request, row={row_number}, missing_column={column}")
    if errors:
        raise ValueError("필수 입력 누락:\n" + "\n".join(errors))


def _fallback_format_recommendation_result_for_review(recommendation_result: pd.DataFrame) -> pd.DataFrame:
    review_columns = [
        "model_suffix",
        "subsidiary",
        "parent_assy_part_no",
        "parent_assy_desc_text",
        "ancestor_desc_path",
        "recommended_base_part_no",
        "recommended_desc",
        "confirmed_new_part_no",
        "user_confirmed",
        "expected_vi_amount_usd",
        "vi_item_name",
        "base_bom_total_model_count",
        "item_target_model_count",
        "recommended_model_count",
        "recommended_model_share_pct",
        "recommendation_reason",
    ]
    if recommendation_result.empty:
        return pd.DataFrame(columns=review_columns)

    formatted = recommendation_result.copy()
    for column in review_columns:
        if column not in formatted.columns:
            formatted[column] = ""
    formatted["user_confirmed"] = formatted["user_confirmed"].fillna("").astype(str).str.strip()
    return formatted[review_columns].copy()


format_recommendation_result_for_review = getattr(
    _vi_engine,
    "format_recommendation_result_for_review",
    _fallback_format_recommendation_result_for_review,
)


RECOMMENDATION_SHEETS = ["BasePN_Recommendation_Result"]
SUMMARY_SHEETS = ["FinalVITargetRule", "ViItemSummary", "ViSubsidiarySummary", "BaseTargetModelList", "BasePNTargetSummary", "ModelChangeSummary"]
DETAIL_SHEETS = ["ModelChangeDetail", "ViModelDetail", "ViExceptionLog"]
ALL_SHEETS = SUMMARY_SHEETS + DETAIL_SHEETS

COMMON_DISPLAY_HEADERS = {
    "source_sheet": "출처",
    "vi_item_id": "아이템ID",
    "vi_item_name": "아이템명",
    "rule_seq": "규칙번호",
    "enabled": "사용여부",
    "rule_type": "기준유형",
    "subsidiary": "법인",
    "base_part_no": "Base P/N",
    "new_part_no": "New P/N",
    "model_type": "모델타입",
    "indoor_tool": "실내기툴",
    "desc_match_mode": "Desc 매칭방식",
    "desc_contains_any": "Desc 키워드",
    "spec_contains_any": "Spec 키워드",
    "parent_assy_desc_contains_any": "상위 ASSY Desc 키워드",
    "parent_assy_part_no": "상위 ASSY P/N",
    "parent_assy_desc_text": "상위 ASSY Desc",
    "note": "비고",
    "recommended_base_part_no": "추천 Base P/N",
    "confirmed_new_part_no": "확정 New P/N",
    "recommended_desc": "추천품번 설명",
    "recommended_subsidiaries": "존재법인",
    "base_bom_total_model_count": "Base BOM전체 모델수",
    "recommended_model_count": "추천품번 모델수",
    "recommended_model_share_pct": "추천품번 비중",
    "recommended_subsidiary_share_detail": "법인별 비중",
    "recommendation_reason": "추천근거",
    "user_confirmed": "사용 여부",
    "target_model_count": "대상모델수",
    "applied_model_count": "적용모델수",
    "not_applied_model_count": "미적용모델수",
    "mixed_review_model_count": "혼재검토모델수",
    "missing_or_model_changed_count": "판정불가모델수",
    "application_rate": "적용률",
    "realized_vi_amount": "실현VI금액",
    "remaining_vi_amount": "잔여VI금액",
    "expected_vi_amount_usd": "예상 VI 금액[$]",
    "expected_vi_calc_status": "예상 VI 계산상태",
    "expected_vi_calc_message": "예상 VI 계산메시지",
    "total_opportunity": "총VI기회",
    "model_suffix": "모델명",
    "base_unit_price": "Base 단가",
    "new_unit_price": "New 단가",
    "base_bom_qty": "Base BOM수량",
    "target_source": "대상출처",
    "base_part_desc": "Base P/N 설명",
    "base_pn_model_count": "Base P/N 모델수",
    "item_target_model_count": "아이템 대상모델수",
    "base_pn_model_share_pct": "Base P/N 비중",
    "base_target_model_count": "Base 대상모델수",
    "existing_model_count": "유지모델수",
    "removed_model_count": "삭제모델수",
    "added_model_count": "추가모델수",
    "base_target_applied_count": "Base 적용모델수",
    "base_target_not_applied_count": "Base 미적용모델수",
    "base_target_mixed_review_count": "Base 혼재검토모델수",
    "base_target_missing_or_model_changed_count": "Base 판정불가모델수",
    "added_applied_count": "추가 적용모델수",
    "added_not_applied_count": "추가 미적용모델수",
    "added_mixed_review_count": "추가 혼재검토모델수",
    "added_not_target_or_unknown_count": "추가 판정불가모델수",
    "model_change_status": "모델변화상태",
    "base_exists": "Base 존재",
    "current_exists": "Current 존재",
    "base_target_flag": "Base 대상여부",
    "current_base_part_exists": "Current Base 존재",
    "current_new_part_exists": "Current New 존재",
    "application_status": "적용상태",
    "bom_qty": "BOM수량",
    "production_qty": "생산수량",
    "issue_flag": "확인필요",
    "issue_message": "확인필요사항",
    "unit_saving": "대당개선액",
    "issue_type": "확인유형",
}

COMMON_COLUMN_DESCRIPTIONS = {
    "source_sheet": "이 최종 규칙이 어느 시트에서 왔는지 보여줍니다.",
    "vi_item_id": "VI 아이템 고유 ID입니다.",
    "vi_item_name": "VI 아이템명입니다.",
    "rule_seq": "규칙 식별용 순번입니다.",
    "enabled": "규칙 활성 여부입니다.",
    "rule_type": "규칙 적용 방식입니다.",
    "subsidiary": "적용 대상 법인입니다. 빈 값은 전체 법인 의미입니다.",
    "base_part_no": "기준이 되는 Base P/N입니다.",
    "new_part_no": "적용하려는 New P/N입니다.",
    "model_type": "모델 타입 분류값입니다.",
    "indoor_tool": "실내기 툴 분류값입니다.",
    "desc_match_mode": "Desc/Spec/상위 ASSY 키워드에 적용된 OR/AND 매칭 방식입니다.",
    "desc_contains_any": "최종 규칙에 보존된 Desc 키워드 조건입니다.",
    "spec_contains_any": "최종 규칙에 보존된 Spec 키워드 조건입니다.",
    "parent_assy_desc_contains_any": "최종 규칙에 보존된 상위 ASSY Desc 키워드 조건입니다.",
    "parent_assy_part_no": "현재 추적 중인 BOM 행의 바로 상위 ASSY P/N입니다.",
    "parent_assy_desc_text": "현재 추적 중인 BOM 행의 바로 상위 ASSY 설명입니다.",
    "note": "원본 규칙 또는 검토 메모입니다.",
    "target_model_count": "해당 집계 기준에서 대상 모델 총 개수입니다.",
    "applied_model_count": "현재 BOM에 New P/N만 존재해 적용된 모델 수입니다.",
    "not_applied_model_count": "현재 BOM에 Base P/N만 존재해 미적용된 모델 수입니다.",
    "mixed_review_model_count": "Base/New P/N이 함께 있어 추가 검토가 필요한 모델 수입니다.",
    "missing_or_model_changed_count": "현재 모델이 사라졌거나 P/N 판정이 어려운 모델 수입니다.",
    "application_rate": "적용 모델 수 / 대상 모델 수 비율입니다.",
    "realized_vi_amount": "이미 적용된 모델 기준 실현 VI 금액입니다.",
    "remaining_vi_amount": "아직 미적용 모델 기준 잠재 VI 금액입니다.",
    "total_opportunity": "실현 VI와 잔여 VI의 합계입니다.",
    "model_suffix": "모델 식별자입니다.",
    "base_unit_price": "Base P/N 단가입니다.",
    "new_unit_price": "New P/N 단가입니다.",
    "base_bom_qty": "Base BOM에서의 수량입니다.",
    "target_source": "대상 모델이 어떤 규칙 로직으로 잡혔는지 표시합니다.",
    "base_part_desc": "Base P/N 설명입니다.",
    "base_pn_model_count": "해당 Base P/N이 걸린 대상 모델 수입니다.",
    "parent_assy_desc_text": "상위 ASSY Desc 조건으로 매칭된 경우 가장 가까운 상위 ASSY 설명입니다.",
    "base_bom_total_model_count": "선택한 Base BOM 전체의 모델 수입니다.",
    "item_target_model_count": "해당 VI 아이템/법인 범위의 전체 대상 모델 수입니다.",
    "base_pn_model_share_pct": "대상 모델 중 해당 Base P/N이 차지하는 비중입니다.",
    "base_target_model_count": "기준월(Base BOM)에서 대상이었던 모델 수입니다.",
    "existing_model_count": "기준월과 현재월 모두 존재하는 모델 수입니다.",
    "removed_model_count": "기준월에는 있었지만 현재월에는 없는 모델 수입니다.",
    "added_model_count": "현재월에 새로 나타난 모델 수입니다.",
    "base_target_applied_count": "기준 대상 모델 중 적용 완료 모델 수입니다.",
    "base_target_not_applied_count": "기준 대상 모델 중 미적용 모델 수입니다.",
    "base_target_mixed_review_count": "기준 대상 모델 중 Base/New가 함께 있는 모델 수입니다.",
    "base_target_missing_or_model_changed_count": "기준 대상 모델 중 현재 판정 불가 모델 수입니다.",
    "added_applied_count": "신규 모델 중 적용 완료 판정된 모델 수입니다.",
    "added_not_applied_count": "신규 모델 중 미적용 판정된 모델 수입니다.",
    "added_mixed_review_count": "신규 모델 중 Base/New가 함께 있는 모델 수입니다.",
    "added_not_target_or_unknown_count": "신규 모델 중 타깃 외 또는 판정 불가 모델 수입니다.",
    "model_change_status": "모델이 유지/삭제/추가 중 무엇인지 표시합니다.",
    "base_exists": "기준월 BOM에 존재했는지 여부입니다.",
    "current_exists": "현재월 BOM에 존재하는지 여부입니다.",
    "base_target_flag": "기준월 타깃 모델인지 여부입니다.",
    "current_base_part_exists": "현재월 BOM에 Base P/N이 있는지 여부입니다.",
    "current_new_part_exists": "현재월 BOM에 New P/N이 있는지 여부입니다.",
    "application_status": "최종 적용 상태 판정값입니다.",
    "bom_qty": "VI 금액 계산에 사용된 BOM 수량입니다.",
    "production_qty": "물량 시트에서 매칭된 생산 수량입니다.",
    "issue_flag": "예외 또는 확인 필요 사항 존재 여부입니다.",
    "issue_message": "예외 또는 확인 필요 사유입니다.",
    "unit_saving": "Base 단가 - New 단가입니다.",
    "issue_type": "예외 유형 분류입니다.",
}

REQUEST_REVIEW_DISPLAY_HEADERS = {
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
}

RECOMMENDATION_REVIEW_DISPLAY_HEADERS = {
    "model_suffix": "모델명",
    "subsidiary": "법인",
    "parent_assy_part_no": "상위 ASSY P/N",
    "parent_assy_desc_text": "상위 ASSY Desc",
    "ancestor_desc_path": "상위 ASSY 경로",
    "recommended_base_part_no": "추천 Base P/N",
    "recommended_desc": "Base P/N Desc",
    "confirmed_new_part_no": "확정 New P/N",
    "user_confirmed": "사용 여부",
    "expected_vi_amount_usd": "예상 VI 금액[$]",
    "vi_item_name": "VI 아이템",
    "base_bom_total_model_count": "Base BOM전체 모델수",
    "item_target_model_count": "VI대상 모델 수",
    "recommended_model_count": "Base P/N 사용 모델 수",
    "recommended_model_share_pct": "Base P/N 사용 비중",
    "recommendation_reason": "추천 근거",
}

RECOMMENDATION_REVIEW_COLUMN_GUIDES = {
    "모델명": ("자동 생성", "시스템", "추천 실행 시 모델별로 한 줄씩 채워집니다."),
    "법인": ("자동 생성", "시스템", "추천 실행 시 자동으로 채워집니다."),
    "상위 ASSY P/N": ("자동 생성", "시스템", "가장 가까운 상위 ASSY P/N입니다."),
    "상위 ASSY Desc": ("자동 생성", "시스템", "가장 가까운 상위 ASSY Desc입니다."),
    "상위 ASSY 경로": ("자동 생성", "시스템", "가능한 경우 상위 ASSY 경로를 보여줍니다."),
    "사용 여부": ("검토 입력", "사용자", "최종 적용할 행이면 Y를 입력합니다."),
    "추천 Base P/N": ("자동 생성", "시스템", "시스템 추천 결과입니다."),
    "확정 New P/N": ("검토 입력", "사용자", "최종 적용할 New P/N을 입력합니다."),
    "예상 VI 금액[$]": ("자동 생성", "시스템", "최종 집계 단계에서 계산되면 참고용으로 보여줍니다."),
    "VI 아이템": ("자동 생성", "시스템", "내부 규칙 그룹 이름입니다."),
    "Base P/N Desc": ("자동 생성", "시스템", "시스템 추천 결과입니다."),
    "Base BOM전체 모델수": ("자동 생성", "시스템", "선택한 Base BOM 전체의 모델 수입니다."),
    "VI대상 모델 수": ("자동 생성", "시스템", "시스템 계산 결과입니다."),
    "Base P/N 사용 모델 수": ("자동 생성", "시스템", "시스템 계산 결과입니다."),
    "Base P/N 사용 비중": ("자동 생성", "시스템", "시스템 계산 결과입니다."),
    "추천 근거": ("자동 생성", "시스템", "왜 이 품번이 추천되었는지 보여줍니다."),
}


def require_duckdb() -> None:
    if duckdb is None:
        raise RuntimeError("duckdb is required. Install it with: pip install duckdb") from DUCKDB_IMPORT_ERROR


def parse_bool(value: str | bool) -> bool:
    return str(value).strip().lower() in {"1", "true", "y", "yes"}


def parse_month_text(value: str) -> tuple[int, int]:
    text = str(value).strip()
    parts = text.split("-")
    if len(parts) != 2 or not parts[0].isdigit() or not parts[1].isdigit():
        raise ValueError(f"Invalid month format: '{value}'. Use YYYY-MM.")
    return int(parts[0]), int(parts[1])


def month_sort_key(value: str) -> tuple[int, int]:
    try:
        year_text, month_text = str(value).split("-")
        return int(year_text), int(month_text)
    except Exception:
        return 0, 0


def discover_snapshot_months(history_root: Path | None = None) -> list[str]:
    root = (history_root if history_root is not None else runtime_bom_lake_root()) / "historical"
    months: set[str] = set()
    if not root.exists():
        return []
    for year_dir in root.glob("bom_year=*"):
        year_text = year_dir.name.replace("bom_year=", "").strip()
        if not year_text.isdigit():
            continue
        for month_dir in year_dir.glob("bom_month=*"):
            month_text = month_dir.name.replace("bom_month=", "").strip()
            if not month_text.isdigit():
                continue
            if any(month_dir.glob("*.parquet")):
                months.add(f"{int(year_text):04d}-{int(month_text):02d}")
    return sorted(months, key=month_sort_key)


def choose_month_interactive(label: str, available_months: list[str], default_value: str | None = None) -> tuple[int, int]:
    if not available_months:
        entered = prompt_text(f"{label} month", default_value)
        return parse_month_text(entered)

    print(f"\n{label} BOM candidates", flush=True)
    for index, month in enumerate(available_months, start=1):
        marker = " (default)" if month == default_value else ""
        print(f"{index:>2}. {month}{marker}", flush=True)

    default_index = available_months.index(default_value) + 1 if default_value in available_months else 1
    entered = prompt_text(f"Select {label} BOM by number or YYYY-MM", str(default_index))
    if entered.isdigit():
        selected_index = int(entered)
        if 1 <= selected_index <= len(available_months):
            return parse_month_text(available_months[selected_index - 1])
        raise ValueError(f"Invalid {label} BOM selection number: {entered}")
    return parse_month_text(entered)


def default_output_path(base_year: int, base_month: int, current_year: int, current_month: int) -> Path:
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return runtime_output_root() / f"VI_Result_{base_year:04d}-{base_month:02d}_to_{current_year:04d}-{current_month:02d}_{stamp}.xlsx"


def default_recommendation_output_path(scenario_path: Path, base_year: int, base_month: int) -> Path:
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return scenario_path.with_name(f"{scenario_path.stem}_Recommendation_{base_year:04d}-{base_month:02d}_{stamp}.xlsx")


def scenario_file_hash(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _cache_metadata_path_token(path: Path | None) -> str:
    if path is None:
        return ""
    return portable_path_token(path, preferred_root=application_root())


def dashboard_output_root() -> Path:
    return runtime_dashboard_output_root()


def build_cache_key(
    scenario_hash: str,
    base_year: int,
    base_month: int,
    current_year: int,
    current_month: int,
    detail: bool,
    base_source_hash: str = "",
    current_source_hash: str = "",
) -> str:
    raw_key = "|".join(
        [
            scenario_hash,
            f"{base_year:04d}-{base_month:02d}",
            f"{current_year:04d}-{current_month:02d}",
            f"detail={str(detail).lower()}",
            f"base_source_hash={base_source_hash}",
            f"current_source_hash={current_source_hash}",
        ]
    )
    return hashlib.sha256(raw_key.encode("utf-8")).hexdigest()[:20]


def build_recommendation_cache_key(
    scenario_hash: str,
    base_year: int,
    base_month: int,
    base_source_hash: str = "",
) -> str:
    raw_key = "|".join(
        [
            scenario_hash,
            f"{base_year:04d}-{base_month:02d}",
            f"base_source_hash={base_source_hash}",
            "recommendation",
        ]
    )
    return hashlib.sha256(raw_key.encode("utf-8")).hexdigest()[:20]


def cache_dir_for_key(cache_key: str) -> Path:
    return report_cache_root() / cache_key


def report_sheet_names(detail: bool) -> list[str]:
    return ALL_SHEETS if detail else SUMMARY_SHEETS


def filter_report_for_export(report: dict[str, pd.DataFrame], detail: bool) -> dict[str, pd.DataFrame]:
    filtered: dict[str, pd.DataFrame] = {}
    for name in report_sheet_names(detail):
        filtered[name] = report.get(name, pd.DataFrame()).copy()
    return filtered


def _display_header(column_name: object) -> str:
    text = str(column_name).strip()
    if text in COMMON_DISPLAY_HEADERS:
        return COMMON_DISPLAY_HEADERS[text]
    return text.replace("_", " ").strip().title()


def _build_report_guide_sheet(report: dict[str, pd.DataFrame]) -> pd.DataFrame:
    rows: list[dict[str, str]] = []
    for sheet_name, frame in report.items():
        for column in frame.columns:
            engine_header = str(column)
            rows.append(
                {
                    "시트명": sheet_name,
                    "표시헤더": _display_header(engine_header),
                    "엔진헤더": engine_header,
                    "입력구분": "자동 계산",
                    "작성주체": "시스템",
                    "설명": COMMON_COLUMN_DESCRIPTIONS.get(engine_header, ""),
                }
            )
    if not rows:
        return pd.DataFrame(columns=["시트명", "표시헤더", "엔진헤더", "입력구분", "작성주체", "설명"])
    return pd.DataFrame(rows)


def build_excel_friendly_report(report: dict[str, pd.DataFrame]) -> dict[str, pd.DataFrame]:
    friendly_report: dict[str, pd.DataFrame] = {
        "ReportGuide": _build_report_guide_sheet(report),
    }
    for sheet_name, frame in report.items():
        renamed = frame.rename(columns={column: _display_header(column) for column in frame.columns})
        friendly_report[sheet_name] = renamed
    return friendly_report


def log_timing(label: str, started_at: float) -> None:
    print(f"[Timing] {label}: {time.perf_counter() - started_at:.2f}s", flush=True)


def prompt_text(label: str, default: str | None = None) -> str:
    suffix = f" [{default}]" if default else ""
    entered = input(f"{label}{suffix}: ").strip()
    if entered:
        return entered
    if default is not None:
        return default
    raise ValueError(f"{label} is required.")


def prompt_enter(message: str) -> None:
    input(f"{message}\nPress Enter to continue...")


def default_scenario_path() -> Path | None:
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


def choose_excel_file(initial_dir: Path | None = None, title: str = "Select Excel Workbook") -> Path | None:
    try:
        import tkinter as tk
        from tkinter import filedialog
    except Exception:
        return None

    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    try:
        selected = filedialog.askopenfilename(
            title=title,
            initialdir=str(initial_dir) if initial_dir else str(application_root()),
            filetypes=[
                ("Excel files", "*.xlsx *.xlsm *.xls"),
                ("All files", "*.*"),
            ],
        )
    finally:
        root.destroy()

    if not selected:
        return None
    return Path(selected)


def choose_save_excel_file(default_path: Path, title: str = "Save Workbook") -> Path | None:
    try:
        import tkinter as tk
        from tkinter import filedialog
    except Exception:
        return None

    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    try:
        selected = filedialog.asksaveasfilename(
            title=title,
            initialdir=str(default_path.parent),
            initialfile=default_path.name,
            defaultextension=".xlsx",
            filetypes=[
                ("Excel files", "*.xlsx"),
                ("All files", "*.*"),
            ],
        )
    finally:
        root.destroy()

    if not selected:
        return None
    return Path(selected)


def write_frame_parquet(frame: pd.DataFrame, output_path: Path) -> None:
    require_duckdb()
    output_path.parent.mkdir(parents=True, exist_ok=True)
    con = duckdb.connect()
    try:
        con.register("frame_view", frame)
        con.execute("COPY frame_view TO ? (FORMAT PARQUET)", [str(output_path)])
    finally:
        con.close()


def read_frame_parquet(input_path: Path) -> pd.DataFrame:
    require_duckdb()
    con = duckdb.connect()
    try:
        return con.execute("SELECT * FROM read_parquet(?)", [str(input_path)]).df()
    finally:
        con.close()


def save_named_parquets(data: dict[str, pd.DataFrame], output_dir: Path, active_names: list[str], all_names: list[str]) -> None:
    output_dir.mkdir(parents=True, exist_ok=True)
    active_set = set(active_names)
    for name in all_names:
        parquet_path = output_dir / f"{name}.parquet"
        if name in active_set:
            write_frame_parquet(data.get(name, pd.DataFrame()), parquet_path)
        elif parquet_path.exists():
            parquet_path.unlink()


def load_named_parquets(output_dir: Path, active_names: list[str]) -> dict[str, pd.DataFrame]:
    loaded: dict[str, pd.DataFrame] = {}
    for name in active_names:
        parquet_path = output_dir / f"{name}.parquet"
        if not parquet_path.exists():
            raise FileNotFoundError(f"Cached parquet is missing: {parquet_path}")
        loaded[name] = read_frame_parquet(parquet_path)
    return loaded


def save_report_parquets(report: dict[str, pd.DataFrame], output_dir: Path, detail: bool) -> None:
    save_named_parquets(report, output_dir, report_sheet_names(detail), ALL_SHEETS)


def load_report_parquets(output_dir: Path, detail: bool) -> dict[str, pd.DataFrame]:
    report = load_named_parquets(output_dir, report_sheet_names(detail))
    if detail:
        return report
    report["ModelChangeDetail"] = pd.DataFrame()
    report["ViModelDetail"] = pd.DataFrame()
    report["ViExceptionLog"] = pd.DataFrame()
    return report


def mirror_cache_outputs(cache_dir: Path, active_names: list[str], all_names: list[str]) -> None:
    target_dir = dashboard_output_root()
    target_dir.mkdir(parents=True, exist_ok=True)
    active_set = set(active_names)
    for name in all_names:
        source = cache_dir / f"{name}.parquet"
        target = target_dir / f"{name}.parquet"
        if name in active_set and source.exists():
            shutil.copy2(source, target)
        elif target.exists():
            target.unlink()


def save_cache_metadata(cache_dir: Path, metadata: dict[str, object]) -> None:
    cache_dir.mkdir(parents=True, exist_ok=True)
    (cache_dir / "metadata.json").write_text(json.dumps(metadata, ensure_ascii=True, indent=2), encoding="utf-8")


def export_excel(report: dict[str, pd.DataFrame], output_path: Path) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    save_vi_scenario_workbook(build_excel_friendly_report(report), output_path)


def auto_fit_worksheet_columns(worksheet) -> None:
    for column_cells in worksheet.columns:
        max_length = 0
        column_letter = get_column_letter(column_cells[0].column)
        for cell in column_cells:
            value = "" if cell.value is None else str(cell.value)
            if len(value) > max_length:
                max_length = len(value)
        worksheet.column_dimensions[column_letter].width = min(max(max_length + 2, 10), 60)


def build_progress_logger() -> Callable[[float, str], None]:
    def update_progress(value: float, message: str) -> None:
        print(f"[Progress] {int(max(0.0, min(1.0, value)) * 100):3d}% {message}", flush=True)

    return update_progress


def _replace_or_create_sheet(workbook, sheet_name: str, index: int | None = None):
    if sheet_name in workbook.sheetnames:
        current_index = workbook.sheetnames.index(sheet_name)
        worksheet = workbook[sheet_name]
        workbook.remove(worksheet)
        return workbook.create_sheet(sheet_name, current_index if index is None else index)
    if index is None:
        return workbook.create_sheet(sheet_name)
    return workbook.create_sheet(sheet_name, index)


def _write_dataframe_to_worksheet(worksheet, frame: pd.DataFrame) -> None:
    excel_safe = frame.copy()
    excel_safe = excel_safe.replace([float("inf"), float("-inf")], "")
    excel_safe = excel_safe.astype(object).where(pd.notna(excel_safe), "")
    for row in dataframe_to_rows(excel_safe, index=False, header=True):
        worksheet.append(row)
    worksheet.freeze_panes = "A2"
    worksheet.auto_filter.ref = worksheet.dimensions
    auto_fit_worksheet_columns(worksheet)


def _build_recommendation_review_guide_sheet() -> pd.DataFrame:
    rows: list[dict[str, str]] = [
        {
            "구분": "수정할 컬럼",
            "컬럼": "적용 확정 여부(Y/N), 확정 New P/N",
            "설명": "사용자가 실제로 입력해야 하는 컬럼입니다.",
        },
        {
            "구분": "확정 여부 입력",
            "컬럼": "적용 확정 여부(Y/N)",
            "설명": "적용할 추천 행이면 Y, 제외할 행이면 N 또는 빈칸으로 둡니다.",
        },
        {
            "구분": "최종 품번 입력",
            "컬럼": "확정 New P/N",
            "설명": "최종 적용할 New P/N을 입력합니다.",
        },
        {
            "구분": "자동 계산 컬럼",
            "컬럼": "그 외 나머지 컬럼",
            "설명": "시스템이 채운 참고 정보이므로 직접 수정하지 않는 것을 권장합니다.",
        },
        {
            "구분": "헤더 색상 의미",
            "컬럼": "주황 / 회색",
            "설명": "주황=검토 입력, 회색=자동 생성",
        },
    ]
    for display_header, guide in RECOMMENDATION_REVIEW_COLUMN_GUIDES.items():
        rows.append(
            {
                "구분": guide[0],
                "컬럼": display_header,
                "설명": guide[2],
            }
        )
    return pd.DataFrame(rows)


def _style_recommendation_review_headers(worksheet) -> None:
    from openpyxl.comments import Comment
    from openpyxl.styles import Border, Font, PatternFill, Side

    fill_by_type = {
        "검토 입력": PatternFill(fill_type="solid", fgColor="8B0000"),
        "자동 생성": PatternFill(fill_type="solid", fgColor="E7E6E6"),
    }
    thick_red = Side(style="thick", color="C00000")
    required_border = Border(left=thick_red, right=thick_red, top=thick_red, bottom=thick_red)
    for column_cells in worksheet.iter_cols(min_row=1, max_row=1):
        cell = column_cells[0]
        guide = RECOMMENDATION_REVIEW_COLUMN_GUIDES.get(str(cell.value).strip())
        input_type = guide[0] if guide else "자동 생성"
        cell.fill = fill_by_type[input_type]
        if input_type == "검토 입력":
            cell.font = Font(color="FFFFFF", bold=True)
            cell.border = required_border
            cell.comment = Comment("[필수 입력]", "VICT")
        else:
            cell.font = Font(bold=True)


def _extract_volume_seed_pairs(recommendation_result: pd.DataFrame) -> list[tuple[str, str]]:
    if recommendation_result.empty or "_volume_seed_pairs" not in recommendation_result.columns:
        return []
    pairs: set[tuple[str, str]] = set()
    for raw_value in recommendation_result["_volume_seed_pairs"].fillna("").astype(str):
        for line in raw_value.splitlines():
            text = line.strip()
            if not text or "\t" not in text:
                continue
            subsidiary, model_suffix = [part.strip() for part in text.split("\t", 1)]
            if subsidiary and model_suffix:
                pairs.add((subsidiary, model_suffix))
    return sorted(pairs)


def _build_recommendation_volume_frame(existing_volume: pd.DataFrame, recommendation_result: pd.DataFrame) -> pd.DataFrame:
    if existing_volume.empty:
        volume_frame = pd.DataFrame(columns=["month", "subsidiary", "model_suffix", "production_qty"])
    else:
        volume_frame = existing_volume.copy()
    for column in ["month", "subsidiary", "model_suffix", "production_qty"]:
        if column not in volume_frame.columns:
            volume_frame[column] = ""
    existing_pairs = {
        (str(row["subsidiary"]).strip(), str(row["model_suffix"]).strip())
        for _, row in volume_frame.iterrows()
        if str(row["subsidiary"]).strip() and str(row["model_suffix"]).strip()
    }
    new_rows = [
        {
            "month": "",
            "subsidiary": subsidiary,
            "model_suffix": model_suffix,
            "production_qty": "",
        }
        for subsidiary, model_suffix in _extract_volume_seed_pairs(recommendation_result)
        if (subsidiary, model_suffix) not in existing_pairs
    ]
    if new_rows:
        volume_frame = pd.concat([volume_frame, pd.DataFrame(new_rows)], ignore_index=True)
    return volume_frame[["month", "subsidiary", "model_suffix", "production_qty"]].copy()


def replace_sheet_with_dataframe(
    workbook_path: Path,
    output_path: Path,
    sheet_name: str,
    frame: pd.DataFrame,
    remove_sheets: set[str] | None = None,
) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(workbook_path, output_path)
    workbook = load_workbook(output_path)
    for remove_sheet_name in remove_sheets or set():
        if remove_sheet_name in workbook.sheetnames and remove_sheet_name != sheet_name:
            workbook.remove(workbook[remove_sheet_name])
    review_guide_index = 0 if "ReviewGuide" not in workbook.sheetnames else workbook.sheetnames.index("ReviewGuide")
    worksheet = _replace_or_create_sheet(workbook, sheet_name)
    _write_dataframe_to_worksheet(worksheet, frame)
    if sheet_name == "BasePN_Recommendation_Result":
        _style_recommendation_review_headers(worksheet)
        guide_sheet = _replace_or_create_sheet(workbook, "ReviewGuide", review_guide_index)
        _write_dataframe_to_worksheet(guide_sheet, _build_recommendation_review_guide_sheet())
    workbook.save(output_path)


def replace_sheets_with_dataframes(
    workbook_path: Path,
    output_path: Path,
    sheet_frames: dict[str, pd.DataFrame],
    remove_sheets: set[str] | None = None,
) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(workbook_path, output_path)
    workbook = load_workbook(output_path)
    for remove_sheet_name in remove_sheets or set():
        if remove_sheet_name in workbook.sheetnames and remove_sheet_name not in sheet_frames:
            workbook.remove(workbook[remove_sheet_name])
    review_guide_index = 0 if "ReviewGuide" not in workbook.sheetnames else workbook.sheetnames.index("ReviewGuide")
    for sheet_name, frame in sheet_frames.items():
        worksheet = _replace_or_create_sheet(workbook, sheet_name)
        _write_dataframe_to_worksheet(worksheet, frame)
        if sheet_name == "BasePN_Recommendation_Result":
            _style_recommendation_review_headers(worksheet)
    if "BasePN_Recommendation_Result" in sheet_frames:
        guide_sheet = _replace_or_create_sheet(workbook, "ReviewGuide", review_guide_index)
        _write_dataframe_to_worksheet(guide_sheet, _build_recommendation_review_guide_sheet())
    workbook.save(output_path)


def open_workbook_for_review(path: Path) -> None:
    try:
        os.startfile(str(path))
        print(f"Opened workbook for review: {path}", flush=True)
    except Exception:
        print(f"Please open this workbook manually for review: {path}", flush=True)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="CLI-first VI recommendation/report workflow",
        formatter_class=argparse.RawTextHelpFormatter,
        epilog=(
            "Examples:\n"
            "  python vi_report_cli.py --mode recommend -b 2025-01\n"
            "  python vi_report_cli.py --mode report -b 2025-01 -c 2025-12 --detail false\n"
            "  python vi_report_cli.py --mode both -b 2025-01 -c 2025-12\n"
        ),
    )
    parser.add_argument("--mode", choices=["recommend", "report", "both"], default="both")
    parser.add_argument("-b", "--base", help="Base BOM month in YYYY-MM")
    parser.add_argument("-c", "--current", help="Current BOM month in YYYY-MM")
    parser.add_argument("--current-file", help="Optional Current BOM Excel file path. When provided, the file is loaded into current_cache and used for report calculation.")
    parser.add_argument("--current-sheet", help="Optional Current BOM sheet name when --current-file is used.")
    parser.add_argument("--base-year", type=int)
    parser.add_argument("--base-month", type=int)
    parser.add_argument("--current-year", type=int)
    parser.add_argument("--current-month", type=int)
    parser.add_argument("-s", "--scenario", help="Scenario workbook path")
    parser.add_argument("--output", help="Report workbook output path")
    parser.add_argument("--recommend-output", help="Recommendation review workbook output path")
    parser.add_argument("--history-root", help="Optional BOM lake root")
    parser.add_argument("-d", "--detail", default="false", help="true/false. Default false")
    parser.add_argument("--force", action="store_true", help="Ignore cache and recalculate")
    parser.add_argument("--no-review-pause", action="store_true", help="Do not pause for workbook review in default both flow")
    return parser.parse_args()


def validate_month(year: int, month: int, label: str) -> None:
    if year < 1900 or year > 9999:
        raise ValueError(f"{label} year is invalid: {year}")
    if month < 1 or month > 12:
        raise ValueError(f"{label} month is invalid: {month}")


def resolve_month_input(
    combined_text: str | None,
    year: int | None,
    month: int | None,
    label: str,
    default_text: str,
) -> tuple[int, int]:
    if combined_text:
        resolved_year, resolved_month = parse_month_text(combined_text)
    elif year is not None and month is not None:
        resolved_year, resolved_month = year, month
    else:
        resolved_year, resolved_month = parse_month_text(prompt_text(f"{label} month", default_text))
    validate_month(resolved_year, resolved_month, label)
    return resolved_year, resolved_month


def run_recommendation_flow(
    *,
    scenario_path: Path,
    scenario_hash: str,
    base_year: int,
    base_month: int,
    history_root: Path | None,
    recommend_output_path: Path,
    force: bool,
    base_snapshot_kind: str = "historical",
    base_source_hash: str = "",
    base_snapshot_path: Path | None = None,
) -> tuple[pd.DataFrame, str]:
    scenario_inputs = load_scenario_workbook(scenario_path)
    _validate_required_request_rows(scenario_inputs)
    recommendation_cache_key = build_recommendation_cache_key(
        scenario_hash,
        base_year,
        base_month,
        base_source_hash=base_source_hash,
    )
    recommendation_cache_dir = recommendation_cache_root() / f"recommend_{recommendation_cache_key}"

    recommendation_df = pd.DataFrame()
    cache_ready = (
        not force
        and recommendation_cache_dir.exists()
        and (recommendation_cache_dir / "BasePN_Recommendation_Result.parquet").exists()
    )
    if cache_ready:
        print(f"[Cache] recommendation hit: {recommendation_cache_key}", flush=True)
        load_started_at = time.perf_counter()
        recommendation_data = load_named_parquets(recommendation_cache_dir, RECOMMENDATION_SHEETS)
        recommendation_df = recommendation_data["BasePN_Recommendation_Result"]
        log_timing("load cached recommendation outputs", load_started_at)
        missing_columns = [column for column in RECOMMENDATION_RESULT_COLUMNS if column not in recommendation_df.columns]
        if missing_columns:
            print(
                f"[Cache] recommendation schema mismatch. Rebuilding cache for missing columns: {', '.join(missing_columns)}",
                flush=True,
            )
            cache_ready = False

    if not cache_ready:
        if force:
            print(f"[Cache] recommendation bypassed with --force: {recommendation_cache_key}", flush=True)
        elif recommendation_df.empty:
            print(f"[Cache] recommendation miss: {recommendation_cache_key}", flush=True)
        recommend_started_at = time.perf_counter()
        recommendation_df = build_base_pn_recommendations(
            history_root=history_root,
            base_year=base_year,
            base_month=base_month,
            scenario_workbook_path=scenario_path,
            base_snapshot_kind=base_snapshot_kind,
            base_snapshot_path=base_snapshot_path,
        )
        log_timing("build recommendation result", recommend_started_at)
        recommendation_data = {"BasePN_Recommendation_Result": recommendation_df}
        parquet_started_at = time.perf_counter()
        save_named_parquets(recommendation_data, recommendation_cache_dir, RECOMMENDATION_SHEETS, RECOMMENDATION_SHEETS)
        log_timing("export cached recommendation outputs", parquet_started_at)
        save_cache_metadata(
            recommendation_cache_dir,
            {
                "cache_key": recommendation_cache_key,
                "scenario_path": _cache_metadata_path_token(scenario_path),
                "scenario_hash": scenario_hash,
                "base_year": base_year,
                "base_month": base_month,
                "base_snapshot_kind": base_snapshot_kind,
                "base_source_hash": base_source_hash,
                "base_snapshot_path": _cache_metadata_path_token(base_snapshot_path),
                "mode": "recommend",
            },
        )

    mirror_started_at = time.perf_counter()
    mirror_cache_outputs(recommendation_cache_dir, RECOMMENDATION_SHEETS, RECOMMENDATION_SHEETS)
    log_timing("mirror recommendation parquet", mirror_started_at)

    if not recommendation_df.empty and "subsidiary" in recommendation_df.columns:
        if bool(recommendation_df["subsidiary"].fillna("").astype(str).str.strip().eq("").any()):
            print("Blank subsidiary in recommendation result. Treating as ALL subsidiaries.", flush=True)

    workbook_started_at = time.perf_counter()
    print("[Review] Writing recommendation workbook for user confirmation.", flush=True)
    review_frame = format_recommendation_result_for_review(recommendation_df).rename(
        columns={column: RECOMMENDATION_REVIEW_DISPLAY_HEADERS.get(column, column) for column in RECOMMENDATION_REVIEW_DISPLAY_HEADERS}
    )
    scenario_inputs = load_scenario_workbook(scenario_path)
    request_frame = scenario_inputs.recommendation_request.copy()
    request_frame["desc_match_mode"] = request_frame["desc_match_mode"].replace({"ANY": "OR", "ALL": "AND"})
    request_frame = request_frame[list(REQUEST_REVIEW_DISPLAY_HEADERS)].rename(columns=REQUEST_REVIEW_DISPLAY_HEADERS)
    volume_frame = _build_recommendation_volume_frame(scenario_inputs.volume, recommendation_df)
    replace_sheets_with_dataframes(
        scenario_path,
        recommend_output_path,
        {
            "BasePN_Recommendation_Request": request_frame,
            "BasePN_Recommendation_Result": review_frame,
            "Volume": volume_frame,
        },
        remove_sheets={"VI_Target_Rule"},
    )
    log_timing("write recommendation workbook", workbook_started_at)
    return recommendation_df, recommendation_cache_key


def run_report_flow(
    *,
    scenario_path: Path,
    scenario_hash: str,
    base_year: int,
    base_month: int,
    current_year: int,
    current_month: int,
    history_root: Path | None,
    detail: bool,
    output_path: Path,
    force: bool,
    base_snapshot_kind: str = "historical",
    base_source_hash: str = "",
    base_snapshot_path: Path | None = None,
    current_snapshot_kind: str = "historical",
    current_source_hash: str = "",
    use_cache: bool = True,
    current_snapshot_path: Path | None = None,
) -> tuple[dict[str, pd.DataFrame], str]:
    scenario_inputs = load_scenario_workbook(scenario_path)
    invalid_confirmed_rows = find_invalid_confirmed_recommendation_rows(scenario_inputs.recommendation_result)
    if invalid_confirmed_rows:
        print("Invalid confirmed recommendation rows detected:", flush=True)
        for line in invalid_confirmed_rows:
            print(f"- {line}", flush=True)
        raise ValueError(
            "사용 여부가 Y인 행에 적용 New P/N이 비어 있습니다.\n" + "\n".join(invalid_confirmed_rows)
        )

    target_rule_filled_rows = 0
    if not scenario_inputs.raw_vi_target_rule.empty:
        target_rule_filled_rows = int(
            (
                scenario_inputs.raw_vi_target_rule["vi_item_id"].fillna("").astype(str).str.strip().ne("")
                & scenario_inputs.raw_vi_target_rule["base_part_no"].fillna("").astype(str).str.strip().ne("")
                & scenario_inputs.raw_vi_target_rule["new_part_no"].fillna("").astype(str).str.strip().ne("")
                & scenario_inputs.raw_vi_target_rule["enabled"].fillna("").astype(str).str.strip().str.upper().ne("N")
            ).sum()
        )
    recommendation_rules = build_rules_from_recommendation_result(scenario_inputs.recommendation_result)
    confirmed_recommendation_rows = len(recommendation_rules.index)
    final_rule_count = len(scenario_inputs.final_vi_target_rule.index)

    print("Final rule validation", flush=True)
    print(f"- Confirmed recommendation usable rows count: {confirmed_recommendation_rows}", flush=True)
    print(f"- Backward-compatible VI_Target_Rule usable rows count: {target_rule_filled_rows}", flush=True)
    print(f"- Final rule count: {final_rule_count}", flush=True)

    if confirmed_recommendation_rows > 0:
        print("Confirmed BasePN_Recommendation_Result rows will be used as final rules.", flush=True)
    elif target_rule_filled_rows > 0:
        print("No usable confirmed recommendation rows. Existing VI_Target_Rule rows are used only for backward compatibility.", flush=True)
    if target_rule_filled_rows == 0 and confirmed_recommendation_rows == 0:
        print(
            "Warning: No final VI rules found. BasePN_Recommendation_Result에서 사용 여부=Y, 확정 New P/N을 입력해야 최종 집계됩니다.",
            flush=True,
        )
    for message in scenario_inputs.validation_messages:
        print(f"Warning: {message}", flush=True)

    cache_key = build_cache_key(
        scenario_hash,
        base_year,
        base_month,
        current_year,
        current_month,
        detail,
        base_source_hash=base_source_hash,
        current_source_hash=current_source_hash,
    )
    cache_dir = cache_dir_for_key(cache_key)

    if (
        use_cache
        and
        not force
        and cache_dir.exists()
        and all((cache_dir / f"{name}.parquet").exists() for name in report_sheet_names(detail))
    ):
        print(f"[Cache] report hit: {cache_key}", flush=True)
        load_started_at = time.perf_counter()
        report = load_report_parquets(cache_dir, detail)
        log_timing("load cached parquet outputs", load_started_at)
    else:
        if not use_cache:
            print(f"[Cache] report disabled for this run: {cache_key}", flush=True)
        elif force:
            print("[Cache] report bypassed with --force", flush=True)
            print(f"[Cache] report key: {cache_key}", flush=True)
            print("[Cache] bom_pair compact cache remains enabled during --force report rebuild.", flush=True)
        else:
            print(f"[Cache] report miss: {cache_key}", flush=True)

        calculation_started_at = time.perf_counter()
        report = build_vi_scenario_outputs(
            history_root=history_root,
            base_year=base_year,
            base_month=base_month,
            current_year=current_year,
            current_month=current_month,
            scenario_workbook_path=scenario_path,
            base_snapshot_kind=base_snapshot_kind,
            base_snapshot_path=base_snapshot_path,
            current_snapshot_kind=current_snapshot_kind,
            current_snapshot_path=current_snapshot_path,
            detail=detail,
            progress_callback=build_progress_logger(),
        )
        log_timing("calculation", calculation_started_at)
        report = filter_report_for_export(report, detail)

        if use_cache:
            parquet_started_at = time.perf_counter()
            save_report_parquets(report, cache_dir, detail)
            log_timing("export cached parquet outputs", parquet_started_at)

            save_cache_metadata(
                cache_dir,
                {
                    "cache_key": cache_key,
                    "scenario_path": _cache_metadata_path_token(scenario_path),
                    "scenario_hash": scenario_hash,
                    "base_year": base_year,
                    "base_month": base_month,
                    "base_snapshot_kind": base_snapshot_kind,
                    "base_source_hash": base_source_hash,
                    "base_snapshot_path": _cache_metadata_path_token(base_snapshot_path),
                    "current_year": current_year,
                    "current_month": current_month,
                    "current_snapshot_kind": current_snapshot_kind,
                    "current_source_hash": current_source_hash,
                    "current_snapshot_path": _cache_metadata_path_token(current_snapshot_path),
                    "detail": detail,
                    "mode": "report",
                },
            )

    if use_cache:
        mirror_started_at = time.perf_counter()
        mirror_cache_outputs(cache_dir, report_sheet_names(detail), ALL_SHEETS)
        log_timing("mirror runtime parquet outputs", mirror_started_at)

    if not report.get("ViItemSummary", pd.DataFrame()).empty and report.get("BasePNTargetSummary", pd.DataFrame()).empty:
        print("Warning: BasePNTargetSummary is empty while ViItemSummary is not empty.", flush=True)
    if final_rule_count > 0 and report.get("BaseTargetModelList", pd.DataFrame()).empty:
        print(
            "Warning: Final VI rules exist, but no base target models were matched. Check that recommended Base P/N exists in the selected Base BOM month and that BasePN_Recommendation_Result has 사용 여부=Y and 확정 New P/N.",
            flush=True,
        )

    excel_started_at = time.perf_counter()
    export_excel(report, output_path)
    log_timing("export Excel", excel_started_at)
    return report, cache_key


def load_current_bom_file_for_report(
    *,
    current_file_path: Path,
    current_sheet_name: str | None,
    current_year: int | None,
    current_month: int | None,
    history_root: Path | None,
    snapshot_name: str | None = None,
    overwrite_existing_month: bool = True,
) -> tuple[int, int, str, Path]:
    cache_started_at = time.perf_counter()
    current_result = cache_current_month(
        file_path=current_file_path,
        root=history_root,
        preferred_sheet_name=current_sheet_name,
        bom_year=current_year,
        bom_month=current_month,
        persist_as_snapshot=False,
        overwrite_existing_month=overwrite_existing_month,
        snapshot_name=snapshot_name,
    )
    _vi_engine.load_bom_snapshot_compact.cache_clear()
    _vi_engine._snapshot_columns.cache_clear()
    log_timing("cache current bom file", cache_started_at)
    print(
        f"Loaded Current BOM file into current_cache: {current_result['parquet_path']} ({current_result['snapshot_month']})",
        flush=True,
    )
    return (
        int(current_result["bom_year"]),
        int(current_result["bom_month"]),
        current_result["snapshot_month"],
        Path(str(current_result["parquet_path"])).expanduser().resolve(),
    )


def register_current_snapshot_from_excel(
    *,
    current_file_path: Path,
    snapshot_name: str,
    current_year: int,
    current_month: int,
    current_sheet_name: str | None,
    history_root: Path | None,
    overwrite_existing_month: bool = False,
) -> dict[str, object]:
    cache_started_at = time.perf_counter()
    print(
        f"Converting Current BOM Excel to parquet snapshot: {snapshot_name} ({current_year:04d}-{current_month:02d})",
        flush=True,
    )
    current_result = save_current_snapshot_from_excel(
        file_path=current_file_path,
        snapshot_name=snapshot_name,
        bom_year=current_year,
        bom_month=current_month,
        root=history_root,
        preferred_sheet_name=current_sheet_name,
        overwrite_existing_month=overwrite_existing_month,
    )
    _vi_engine.load_bom_snapshot_compact.cache_clear()
    _vi_engine._snapshot_columns.cache_clear()
    log_timing("register current snapshot", cache_started_at)
    print(
        "Current snapshot saved: "
        f"{current_result['parquet_path']} | rows={len(current_result['data'].index):,} | "
        f"models={current_result['data']['model_suffix'].nunique():,}",
        flush=True,
    )
    return current_result


def main() -> int:
    require_duckdb()
    args = parse_args()
    total_started_at = time.perf_counter()
    try:
        interactive_default_flow = len(sys.argv) == 1
        history_root = Path(args.history_root).expanduser().resolve() if args.history_root else None
        available_months = discover_snapshot_months(history_root)

        if interactive_default_flow:
            print("Interactive VI workflow", flush=True)
            print("1. Select Base BOM from existing parquet snapshots", flush=True)
            print("2. Select scenario template workbook", flush=True)
            print("3. Review recommended Base P/N", flush=True)
            print("4. In BasePN_Recommendation_Result, set 사용 여부=Y and enter 확정 New P/N", flush=True)
            print("5. Select Current BOM from existing parquet snapshots or load a Current BOM file", flush=True)
            print("6. Create final item application summary", flush=True)

        if interactive_default_flow and not args.base and args.base_year is None and args.base_month is None:
            default_base = available_months[0] if available_months else "2025-01"
            base_year, base_month = choose_month_interactive("Base", available_months, default_base)
        else:
            base_year, base_month = resolve_month_input(
                args.base,
                args.base_year,
                args.base_month,
                "Base",
                available_months[0] if available_months else "2025-01",
            )

        current_year = current_month = None
        current_snapshot_kind = "historical"
        current_source_hash = ""
        current_snapshot_path: Path | None = None
        current_file_path: Path | None = Path(args.current_file).expanduser().resolve() if args.current_file else None

        scenario_default = default_scenario_path()
        if args.scenario:
            scenario_path = Path(args.scenario).expanduser().resolve()
        else:
            selected_scenario = choose_excel_file(
                scenario_default.parent if scenario_default is not None else None,
                title="Select Scenario Workbook",
            )
            if selected_scenario is not None:
                print(f"Selected scenario workbook: {selected_scenario}", flush=True)
                scenario_path = selected_scenario.expanduser().resolve()
            else:
                scenario_text = prompt_text(
                    "Scenario workbook path",
                    str(scenario_default) if scenario_default is not None else None,
                )
                scenario_path = Path(scenario_text).expanduser().resolve()
        if not scenario_path.exists():
            raise FileNotFoundError(f"Scenario workbook not found: {scenario_path}")

        detail = parse_bool(args.detail)

        delay_report_output_prompt = interactive_default_flow and args.mode == "both" and not args.output

        report_output_path: Path | None = None
        if args.mode in {"report", "both"} and not delay_report_output_prompt:
            output_default = default_output_path(
                base_year,
                base_month,
                current_year or base_year,
                current_month or base_month,
            )
            if args.output:
                report_output_path = Path(args.output).expanduser().resolve()
            else:
                selected_output = choose_save_excel_file(output_default, "Save VI Report Workbook")
                if selected_output is not None:
                    print(f"Selected report workbook: {selected_output}", flush=True)
                    report_output_path = selected_output.expanduser().resolve()
                else:
                    output_text = prompt_text("Report workbook path", str(output_default))
                    report_output_path = Path(output_text).expanduser().resolve()

        recommend_output_path: Path | None = None
        if args.mode in {"recommend", "both"}:
            recommend_default = default_recommendation_output_path(scenario_path, base_year, base_month)
            if args.recommend_output:
                recommend_output_path = Path(args.recommend_output).expanduser().resolve()
            else:
                selected_output = choose_save_excel_file(recommend_default, "Save Recommendation Review Workbook")
                if selected_output is not None:
                    print(f"Selected recommendation workbook: {selected_output}", flush=True)
                    recommend_output_path = selected_output.expanduser().resolve()
                else:
                    output_text = prompt_text("Recommendation workbook path", str(recommend_default))
                    recommend_output_path = Path(output_text).expanduser().resolve()

        hash_started_at = time.perf_counter()
        scenario_hash = scenario_file_hash(scenario_path)
        log_timing("hash scenario workbook", hash_started_at)

        recommendation_df = pd.DataFrame()
        recommendation_cache_key = ""
        if args.mode in {"recommend", "both"} and recommend_output_path is not None:
            recommendation_df, recommendation_cache_key = run_recommendation_flow(
                scenario_path=scenario_path,
                scenario_hash=scenario_hash,
                base_year=base_year,
                base_month=base_month,
                history_root=history_root,
                recommend_output_path=recommend_output_path,
                force=args.force,
            )
            if interactive_default_flow and args.mode == "both" and not args.no_review_pause:
                open_workbook_for_review(recommend_output_path)
                prompt_enter(
                    "BasePN_Recommendation_Result에서 사용 여부=Y, 확정 New P/N을 입력하고 저장한 뒤 돌아와 Enter를 누르세요. VI_Target_Rule은 구버전 파일 호환용으로만 읽습니다."
                )
                scenario_path = recommend_output_path
                hash_started_at = time.perf_counter()
                scenario_hash = scenario_file_hash(scenario_path)
                log_timing("rehash reviewed scenario workbook", hash_started_at)

        if args.mode in {"report", "both"}:
            if interactive_default_flow and current_file_path is None and not args.current and args.current_year is None and args.current_month is None:
                current_source = prompt_text("Current BOM source (snapshot/file)", "snapshot").strip().lower()
                if current_source == "file":
                    selected_current = choose_excel_file(
                        sample_files_root(),
                        title="Select Current BOM Workbook",
                    )
                    if selected_current is not None:
                        print(f"Selected Current BOM file: {selected_current}", flush=True)
                        current_file_path = selected_current.expanduser().resolve()
                    else:
                        current_text = prompt_text("Current BOM file path")
                        current_file_path = Path(current_text).expanduser().resolve()
            if current_file_path is not None:
                current_source_hash = scenario_file_hash(current_file_path)
                current_year, current_month, _, current_snapshot_path = load_current_bom_file_for_report(
                    current_file_path=current_file_path,
                    current_sheet_name=args.current_sheet,
                    current_year=args.current_year,
                    current_month=args.current_month,
                    history_root=history_root,
                )
                current_snapshot_kind = "current_cache"
            else:
                if interactive_default_flow and not args.current and args.current_year is None and args.current_month is None:
                    base_text = f"{base_year:04d}-{base_month:02d}"
                    filtered_current_months = [month for month in available_months if month_sort_key(month) >= month_sort_key(base_text)]
                    default_current = filtered_current_months[-1] if filtered_current_months else (available_months[-1] if available_months else "2025-12")
                    current_year, current_month = choose_month_interactive("Current", filtered_current_months or available_months, default_current)
                else:
                    current_year, current_month = resolve_month_input(
                        args.current,
                        args.current_year,
                        args.current_month,
                        "Current",
                        available_months[-1] if available_months else "2025-12",
                    )
                matching_current_snapshots = [
                    row
                    for row in discover_current_snapshots(history_root)
                    if int(row.get("bom_year", 0) or 0) == int(current_year or 0)
                    and int(row.get("bom_month", 0) or 0) == int(current_month or 0)
                ]
                if matching_current_snapshots:
                    selected_snapshot = matching_current_snapshots[0]
                    current_snapshot_kind = "current_cache"
                    current_snapshot_path = Path(str(selected_snapshot["parquet_path"])).expanduser().resolve()
                    print(
                        "Using included current_cache snapshot for Current BOM: "
                        f"{current_snapshot_path} ({selected_snapshot.get('snapshot_month', '')})",
                        flush=True,
                    )

        if args.mode in {"report", "both"} and report_output_path is None:
            output_default = default_output_path(
                base_year,
                base_month,
                current_year or base_year,
                current_month or base_month,
            )
            selected_output = choose_save_excel_file(output_default, "Save VI Report Workbook")
            if selected_output is not None:
                print(f"Selected report workbook: {selected_output}", flush=True)
                report_output_path = selected_output.expanduser().resolve()
            else:
                output_text = prompt_text("Report workbook path", str(output_default))
                report_output_path = Path(output_text).expanduser().resolve()

        report: dict[str, pd.DataFrame] | None = None
        report_cache_key = ""
        if args.mode in {"report", "both"} and report_output_path is not None and current_year is not None and current_month is not None:
            report, report_cache_key = run_report_flow(
                scenario_path=scenario_path,
                scenario_hash=scenario_hash,
                base_year=base_year,
                base_month=base_month,
                current_year=current_year,
                current_month=current_month,
                history_root=history_root,
                detail=detail,
                output_path=report_output_path,
                force=args.force,
                current_snapshot_kind=current_snapshot_kind,
                current_source_hash=current_source_hash,
                use_cache=True,
                current_snapshot_path=current_snapshot_path,
            )

        total_elapsed = time.perf_counter() - total_started_at
        print("", flush=True)
        print("Final Summary", flush=True)
        print(f"- Mode: {args.mode}", flush=True)
        print(f"- Total seconds: {total_elapsed:.2f}", flush=True)

        if args.mode in {"recommend", "both"} and recommend_output_path is not None:
            recommended_row_count = len(recommendation_df.index)
            unique_recommended_part_count = (
                recommendation_df["recommended_base_part_no"].astype(str).str.strip().replace("", pd.NA).dropna().nunique()
                if not recommendation_df.empty and "recommended_base_part_no" in recommendation_df.columns
                else 0
            )
            print(f"- Recommendation workbook: {recommend_output_path}", flush=True)
            print(f"- Recommendation row count: {recommended_row_count}", flush=True)
            print(f"- Recommended Base P/N count: {unique_recommended_part_count}", flush=True)
            print(f"- Recommendation cache key: {recommendation_cache_key or '-'}", flush=True)

        if args.mode in {"report", "both"} and report is not None and report_output_path is not None:
            item_count = len(report.get("ViItemSummary", pd.DataFrame()).index)
            subsidiary_count = len(report.get("ViSubsidiarySummary", pd.DataFrame()).index)
            base_target_model_count = len(report.get("BaseTargetModelList", pd.DataFrame()).index)
            base_pn_target_count = len(report.get("BasePNTargetSummary", pd.DataFrame()).index)
            model_change_frame = report.get("ModelChangeDetail", pd.DataFrame())
            if model_change_frame.empty:
                model_change_summary = report.get("ModelChangeSummary", pd.DataFrame())
                current_compared_model_count = 0
                if not model_change_summary.empty:
                    current_compared_model_count = int(
                        (
                            model_change_summary["existing_model_count"].fillna(0)
                            + model_change_summary["removed_model_count"].fillna(0)
                            + model_change_summary["added_model_count"].fillna(0)
                        ).sum()
                    )
            else:
                current_compared_model_count = len(model_change_frame.index)
            print(f"- Report workbook: {report_output_path}", flush=True)
            print(f"- Item count: {item_count}", flush=True)
            print(f"- Subsidiary count: {subsidiary_count}", flush=True)
            print(f"- Base target model count: {base_target_model_count}", flush=True)
            print(f"- Base P/N target summary count: {base_pn_target_count}", flush=True)
            print(f"- Current compared model count: {current_compared_model_count}", flush=True)
            print(f"- Current source: {'uploaded file' if current_snapshot_kind == 'current_cache' else 'historical snapshot'}", flush=True)
            print(f"- Report cache key: {report_cache_key or '-'}", flush=True)
            print(f"- Detail mode: {str(detail).lower()}", flush=True)
        return 0
    except Exception as exc:
        print(f"Error: {exc}", file=sys.stderr, flush=True)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
