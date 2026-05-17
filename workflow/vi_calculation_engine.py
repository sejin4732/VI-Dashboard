from __future__ import annotations

import hashlib
import json
import sys
import time
from dataclasses import dataclass
from functools import lru_cache
from pathlib import Path
from typing import TYPE_CHECKING, Any, Callable

import pandas as pd

try:
    import duckdb
except ImportError as exc:  # pragma: no cover
    duckdb = None
    DUCKDB_IMPORT_ERROR = exc
else:  # pragma: no cover
    DUCKDB_IMPORT_ERROR = None

if TYPE_CHECKING:
    DuckDBConnection = Any
else:
    DuckDBConnection = Any

bootstrap_dir = Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parent
parent_dir = bootstrap_dir.parent
for path in [bootstrap_dir, parent_dir]:
    if str(path) not in sys.path:
        sys.path.insert(0, str(path))

from runtime_paths import (
    application_root,
    bom_lake_root as runtime_bom_lake_root,
    bom_pair_cache_root,
    dashboard_output_root,
    duckdb_temp_root,
    portable_path_token,
)

try:
    from bom_lake import save_report_workbook, round_financial_columns
except ModuleNotFoundError:
    from compat_imports.bom_lake import save_report_workbook, round_financial_columns  # type: ignore[reportMissingImports]


VI_ITEM_MASTER_COLUMNS = ["vi_item_id", "vi_item_name", "owner", "note"]
RECOMMENDATION_REQUEST_COLUMNS = [
    "vi_item_id",
    "vi_item_name",
    "enabled",
    "subsidiary",
    "base_part_no",
    "new_part_no",
    "model_type",
    "indoor_tool",
    "desc_match_mode",
    "desc_contains_any",
    "spec_contains_any",
    "parent_assy_desc_contains_any",
    "note",
]
RECOMMENDATION_RESULT_COLUMNS = [
    "vi_item_id",
    "model_suffix",
    "subsidiary",
    "parent_assy_part_no",
    "parent_assy_desc_text",
    "ancestor_desc_path",
    "recommended_base_part_no",
    "recommended_desc",
    "confirmed_new_part_no",
    "user_confirmed",
    "vi_item_name",
    "base_part_no",
    "new_part_no",
    "note",
    "desc_match_mode",
    "desc_contains_any",
    "spec_contains_any",
    "parent_assy_desc_contains_any",
    "model_type",
    "indoor_tool",
    "recommended_subsidiaries",
    "bom_subsidiary_model_count_detail",
    "base_bom_total_model_count",
    "item_target_model_count",
    "item_target_subsidiary_model_count_detail",
    "recommended_model_count",
    "recommended_model_share_formula",
    "recommended_model_share_pct",
    "recommended_subsidiary_share_detail",
    "base_pn_subsidiaries",
    "base_pn_models",
    "expected_vi_amount_usd",
    "recommendation_reason",
]
RECOMMENDATION_REVIEW_EXPORT_COLUMNS = [
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
VI_TARGET_RULE_COLUMNS = [
    "vi_item_id",
    "rule_seq",
    "enabled",
    "rule_type",
    "subsidiary",
    "base_part_no",
    "new_part_no",
    "model_type",
    "indoor_tool",
    "set_model_only",
    "target_side",
    "anchor_desc_contains_any",
    "desc_match_mode",
    "desc_contains_any",
    "spec_contains_any",
    "parent_assy_desc_contains_any",
    "desc_not_contains_any",
    "material_contains_any",
    "level_min",
    "level_max",
    "tool_location",
    "custom_rule_key",
    "note",
]
VOLUME_COLUMNS = ["month", "subsidiary", "model_suffix", "production_qty"]

SHEET_ALIASES = {
    "vi_item_master": "VI_Item_Master",
    "vi target master": "VI_Item_Master",
    "vi아이템마스터": "VI_Item_Master",
    "아이템마스터": "VI_Item_Master",
    "basepn_recommendation_request": "BasePN_Recommendation_Request",
    "basepn recommendation request": "BasePN_Recommendation_Request",
    "추천요청": "BasePN_Recommendation_Request",
    "basepn추천요청": "BasePN_Recommendation_Request",
    "basepn_recommendation_result": "BasePN_Recommendation_Result",
    "basepn recommendation result": "BasePN_Recommendation_Result",
    "추천결과": "BasePN_Recommendation_Result",
    "basepn추천결과": "BasePN_Recommendation_Result",
    "vi_target_rule": "VI_Target_Rule",
    "vi target rule": "VI_Target_Rule",
    "vi적용규칙": "VI_Target_Rule",
    "적용규칙": "VI_Target_Rule",
    "volume": "Volume",
    "물량": "Volume",
}

COLUMN_ALIASES = {
    "viitemid": "vi_item_id",
    "vi item id": "vi_item_id",
    "vi_item_id": "vi_item_id",
    "viitemname": "vi_item_name",
    "vi item name": "vi_item_name",
    "vi_item_name": "vi_item_name",
    "vi item 명": "vi_item_name",
    "vi item 이름": "vi_item_name",
    "vi item name(kor)": "vi_item_name",
    "owner": "owner",
    "note": "note",
    "vi 아이템 id": "vi_item_id",
    "vi 아이템명": "vi_item_name",
    "담당자": "owner",
    "비고": "note",
    "vi아이템명": "vi_item_name",
    "vi item 명": "vi_item_name",
    "요청 순번": "rule_seq",
    "사용 여부": "enabled",
    "대상 법인": "subsidiary",
    "기준 품번(base p/n)": "base_part_no",
    "적용 new p/n": "new_part_no",
    "추천 base p/n": "recommended_base_part_no",
    "확정 new p/n": "confirmed_new_part_no",
    "모델 타입(c/o,h/p)": "model_type",
    "실내기 툴(sj,sk,s0,sa)": "indoor_tool",
    "설명 키워드 매칭방식(any/all)": "desc_match_mode",
    "desc. 키워드 매칭방식": "desc_match_mode",
    "세트모델만(y)": "set_model_only",
    "기준 설명 키워드": "anchor_desc_contains_any",
    "대상 설명 키워드": "desc_contains_any",
    "desc. 키워드": "desc_contains_any",
    "스펙 키워드": "spec_contains_any",
    "spec. 키워드(or)": "spec_contains_any",
    "상위 assy 설명 키워드": "parent_assy_desc_contains_any",
    "상위 assy desc. 키워드": "parent_assy_desc_contains_any",
    "상위 assy description": "parent_assy_desc_text",
    "상위 assy desc": "parent_assy_desc_text",
    "추천 품번 설명": "recommended_desc",
    "description": "recommended_desc",
    "추천 존재 법인": "recommended_subsidiaries",
    "사용 중인 법인": "recommended_subsidiaries",
    "원본 bom 법인별 모델 수": "bom_subsidiary_model_count_detail",
    "base bom전체 모델수": "base_bom_total_model_count",
    "대상 모델 수": "item_target_model_count",
    "vi대상 모델 수": "item_target_model_count",
    "대상 법인별 모델 수": "item_target_subsidiary_model_count_detail",
    "추천 비중 계산식": "recommended_model_share_formula",
    "추천 품번 비중": "recommended_model_share_pct",
    "base p/n 사용 비중": "recommended_model_share_pct",
    "법인별 추천 비중": "recommended_subsidiary_share_detail",
    "법인별 사용 비중": "recommended_subsidiary_share_detail",
    "추천 근거": "recommendation_reason",
    "추천 품번 모델 수": "recommended_model_count",
    "base p/n 사용 모델 수": "recommended_model_count",
    "적용 확정 여부(y/n)": "user_confirmed",
    "사용 여부(y/n)": "user_confirmed",
    "생산월(yyyy-mm)": "month",
    "ruleseq": "rule_seq",
    "rule seq": "rule_seq",
    "rule_seq": "rule_seq",
    "순번": "rule_seq",
    "enabled": "enabled",
    "사용여부": "enabled",
    "활성": "enabled",
    "ruletype": "rule_type",
    "rule type": "rule_type",
    "rule_type": "rule_type",
    "적용구분": "rule_type",
    "규칙유형": "rule_type",
    "subsidiary": "subsidiary",
    "법인": "subsidiary",
    "basepartno": "base_part_no",
    "base part no": "base_part_no",
    "base p/n": "base_part_no",
    "base_part_no": "base_part_no",
    "기준pn": "base_part_no",
    "기준 p/n": "base_part_no",
    "기준품번": "base_part_no",
    "newpartno": "new_part_no",
    "new part no": "new_part_no",
    "new p/n": "new_part_no",
    "new_part_no": "new_part_no",
    "변경pn": "new_part_no",
    "변경 p/n": "new_part_no",
    "변경품번": "new_part_no",
    "modeltype": "model_type",
    "model type": "model_type",
    "model_type": "model_type",
    "모델타입": "model_type",
    "model type (c/o or h/p)": "model_type",
    "c/o or h/p": "model_type",
    "indoortool": "indoor_tool",
    "indoor tool": "indoor_tool",
    "indoor_tool": "indoor_tool",
    "실내기tool": "indoor_tool",
    "indoor tool명": "indoor_tool",
    "indoor tool name": "indoor_tool",
    "실내기툴": "indoor_tool",
    "indoor type": "indoor_tool",
    "indoor type (sj etc)": "indoor_tool",
    "descmatchmode": "desc_match_mode",
    "desc match mode": "desc_match_mode",
    "desc matching mode": "desc_match_mode",
    "desc keywords mode": "desc_match_mode",
    "desc 키워드 매칭방식": "desc_match_mode",
    "세트모델만": "set_model_only",
    "setmodelonly": "set_model_only",
    "set_model_only": "set_model_only",
    "대상side": "target_side",
    "targetside": "target_side",
    "target_side": "target_side",
    "anchordesccontainsany": "anchor_desc_contains_any",
    "anchor desc contains any": "anchor_desc_contains_any",
    "anchor_desc_contains_any": "anchor_desc_contains_any",
    "anchor키워드": "anchor_desc_contains_any",
    "anchor desc 키워드": "anchor_desc_contains_any",
    "desccontainsany": "desc_contains_any",
    "desc contains any": "desc_contains_any",
    "desc_contains_any": "desc_contains_any",
    "포함desc": "desc_contains_any",
    "포함 desc": "desc_contains_any",
    "desckeywords": "desc_contains_any",
    "desc keywords": "desc_contains_any",
    "desc 키워드": "desc_contains_any",
    "include desc keywords": "desc_contains_any",
    "find desc keywords": "desc_contains_any",
    "spec": "spec_contains_any",
    "speckeywords": "spec_contains_any",
    "spec keywords": "spec_contains_any",
    "spec_contains_any": "spec_contains_any",
    "스펙키워드": "spec_contains_any",
    "spec키워드": "spec_contains_any",
    "spec 키워드": "spec_contains_any",
    "spec. 키워드or": "spec_contains_any",
    "spec 키워드or": "spec_contains_any",
    "parentassydesc": "parent_assy_desc_contains_any",
    "parent assy desc": "parent_assy_desc_contains_any",
    "parent_assy_desc_contains_any": "parent_assy_desc_contains_any",
    "상위assydesc": "parent_assy_desc_contains_any",
    "상위assy desc": "parent_assy_desc_contains_any",
    "상위assy키워드": "parent_assy_desc_contains_any",
    "상위 assy 키워드": "parent_assy_desc_contains_any",
    "상위 assy desc 키워드": "parent_assy_desc_contains_any",
    "상위부품desc": "parent_assy_desc_contains_any",
    "상위부품 desc": "parent_assy_desc_contains_any",
    "descnotcontainsany": "desc_not_contains_any",
    "desc not contains any": "desc_not_contains_any",
    "desc_not_contains_any": "desc_not_contains_any",
    "제외desc": "desc_not_contains_any",
    "제외 desc": "desc_not_contains_any",
    "materialcontainsany": "material_contains_any",
    "material contains any": "material_contains_any",
    "material_contains_any": "material_contains_any",
    "재질키워드": "material_contains_any",
    "재질 조건": "material_contains_any",
    "levelmin": "level_min",
    "level min": "level_min",
    "level_min": "level_min",
    "최소레벨": "level_min",
    "levelmax": "level_max",
    "level max": "level_max",
    "level_max": "level_max",
    "최대레벨": "level_max",
    "priority": "priority",
    "우선순위": "priority",
    "toollocation": "tool_location",
    "tool location": "tool_location",
    "tool_location": "tool_location",
    "tool위치": "tool_location",
    "tool 위치": "tool_location",
    "customrulekey": "custom_rule_key",
    "custom rule key": "custom_rule_key",
    "custom_rule_key": "custom_rule_key",
    "커스텀규칙": "custom_rule_key",
    "custom 규칙": "custom_rule_key",
    "month": "month",
    "월": "month",
    "modelsuffix": "model_suffix",
    "model suffix": "model_suffix",
    "model_suffix": "model_suffix",
    "모델명": "model_suffix",
    "candidaterank": "candidate_rank",
    "candidate rank": "candidate_rank",
    "candidate_rank": "candidate_rank",
    "후보순위": "candidate_rank",
    "후보 순위": "candidate_rank",
    "recommendedbasepartno": "recommended_base_part_no",
    "recommended base part no": "recommended_base_part_no",
    "recommended base p/n": "recommended_base_part_no",
    "recommended_base_part_no": "recommended_base_part_no",
    "추천basepn": "recommended_base_part_no",
    "추천 base p/n": "recommended_base_part_no",
    "추천기준품번": "recommended_base_part_no",
    "추천 기준 p/n": "recommended_base_part_no",
    "recommendeddesc": "recommended_desc",
    "recommended desc": "recommended_desc",
    "recommended_desc": "recommended_desc",
    "추천desc": "recommended_desc",
    "추천 desc": "recommended_desc",
    "추천 설명": "recommended_desc",
    "recommendedsubsidiaries": "recommended_subsidiaries",
    "recommended subsidiaries": "recommended_subsidiaries",
    "recommended_subsidiaries": "recommended_subsidiaries",
    "추천법인": "recommended_subsidiaries",
    "추천 법인": "recommended_subsidiaries",
    "존재법인": "recommended_subsidiaries",
    "존재 법인": "recommended_subsidiaries",
    "사용 중인 법인": "recommended_subsidiaries",
    "bomsubsidiarymodelcountdetail": "bom_subsidiary_model_count_detail",
    "bom subsidiary model count detail": "bom_subsidiary_model_count_detail",
    "bom_subsidiary_model_count_detail": "bom_subsidiary_model_count_detail",
    "원본bom법인별모델수": "bom_subsidiary_model_count_detail",
    "원본 bom 법인별 모델 수": "bom_subsidiary_model_count_detail",
    "itemtargetmodelcount": "item_target_model_count",
    "item target model count": "item_target_model_count",
    "item_target_model_count": "item_target_model_count",
    "대상모델수": "item_target_model_count",
    "대상 모델 수": "item_target_model_count",
    "vi대상모델수": "item_target_model_count",
    "vi대상 모델 수": "item_target_model_count",
    "itemtargetsubsidiarymodelcountdetail": "item_target_subsidiary_model_count_detail",
    "item target subsidiary model count detail": "item_target_subsidiary_model_count_detail",
    "item_target_subsidiary_model_count_detail": "item_target_subsidiary_model_count_detail",
    "적용대상법인별모델수": "item_target_subsidiary_model_count_detail",
    "적용 대상 법인별 모델 수": "item_target_subsidiary_model_count_detail",
    "recommendedmodelshareformula": "recommended_model_share_formula",
    "recommended model share formula": "recommended_model_share_formula",
    "recommended_model_share_formula": "recommended_model_share_formula",
    "추천모델비중산식": "recommended_model_share_formula",
    "추천 모델 비중 산식": "recommended_model_share_formula",
    "recommendedmodelsharepct": "recommended_model_share_pct",
    "recommended model share pct": "recommended_model_share_pct",
    "recommended_model_share_pct": "recommended_model_share_pct",
    "추천모델비중": "recommended_model_share_pct",
    "추천 모델 비중": "recommended_model_share_pct",
    "basepn사용비중": "recommended_model_share_pct",
    "base p/n 사용 비중": "recommended_model_share_pct",
    "recommendedsubsidiarysharedetail": "recommended_subsidiary_share_detail",
    "recommended subsidiary share detail": "recommended_subsidiary_share_detail",
    "recommended_subsidiary_share_detail": "recommended_subsidiary_share_detail",
    "법인별모델비중": "recommended_subsidiary_share_detail",
    "법인별 모델 비중": "recommended_subsidiary_share_detail",
    "법인별사용비중": "recommended_subsidiary_share_detail",
    "법인별 사용 비중": "recommended_subsidiary_share_detail",
    "recommendationreason": "recommendation_reason",
    "recommendation reason": "recommendation_reason",
    "recommendation_reason": "recommendation_reason",
    "추천근거": "recommendation_reason",
    "추천 근거": "recommendation_reason",
    "recommendedmodelcount": "recommended_model_count",
    "recommended model count": "recommended_model_count",
    "recommended_model_count": "recommended_model_count",
    "추천모델수": "recommended_model_count",
    "추천 모델 수": "recommended_model_count",
    "basepn사용모델수": "recommended_model_count",
    "base p/n 사용 모델 수": "recommended_model_count",
    "userconfirmed": "user_confirmed",
    "user confirmed": "user_confirmed",
    "user_confirmed": "user_confirmed",
    "사용자확정여부": "user_confirmed",
    "사용자 확정 여부": "user_confirmed",
    "confirmedbasepartno": "confirmed_base_part_no",
    "confirmed base part no": "confirmed_base_part_no",
    "confirmed base p/n": "confirmed_base_part_no",
    "confirmed_base_part_no": "confirmed_base_part_no",
    "확정basepn": "confirmed_base_part_no",
    "확정 기준 p/n": "confirmed_base_part_no",
    "확정 base p/n": "confirmed_base_part_no",
    "confirmednewpartno": "confirmed_new_part_no",
    "confirmed new part no": "confirmed_new_part_no",
    "confirmed new p/n": "confirmed_new_part_no",
    "confirmed_new_part_no": "confirmed_new_part_no",
    "확정newpn": "confirmed_new_part_no",
    "확정 변경 p/n": "confirmed_new_part_no",
    "확정 new p/n": "confirmed_new_part_no",
    "법인목록": "base_pn_subsidiaries",
    "모델명목록": "base_pn_models",
    "reviewnote": "review_note",
    "review note": "review_note",
    "review_note": "review_note",
    "검토메모": "review_note",
    "검토 메모": "review_note",
    "productionqty": "production_qty",
    "production qty": "production_qty",
    "production_qty": "production_qty",
    "생산수량": "production_qty",
}

SUPPORTED_CUSTOM_RULE_KEYS = {"model_type_match", "indoor_tool_match"}
STATUS_PRIORITY = {
    "mixed_review": 5,
    "missing_or_model_changed": 4,
    "not_applied": 3,
    "applied": 2,
    "target_only": 1,
}

HISTORICAL_DIR_NAME = "historical"
CURRENT_CACHE_DIR_NAME = "current_cache"
COMPACT_CONTEXT_SCHEMA_VERSION = "compact_context_v2"

MODEL_TYPE_SQL = """
CASE
    WHEN LENGTH(TRIM(CAST(model_suffix AS VARCHAR))) < 4 THEN 'UNKNOWN'
    WHEN UPPER(SUBSTR(TRIM(CAST(model_suffix AS VARCHAR)), 4, 1)) = 'Q' THEN 'C/O'
    WHEN UPPER(SUBSTR(TRIM(CAST(model_suffix AS VARCHAR)), 4, 1)) = 'W' THEN 'H/P'
    ELSE 'OTHER'
END
"""

INDOOR_TOOL_SQL = """
CASE
    WHEN LENGTH(TRIM(CAST(model_suffix AS VARCHAR))) < 7 THEN 'UNKNOWN'
    WHEN UPPER(SUBSTR(TRIM(CAST(model_suffix AS VARCHAR)), 7, 1)) = 'J' THEN 'SJ'
    WHEN UPPER(SUBSTR(TRIM(CAST(model_suffix AS VARCHAR)), 7, 1)) = 'K' THEN 'SK'
    WHEN UPPER(SUBSTR(TRIM(CAST(model_suffix AS VARCHAR)), 7, 1)) = '0' THEN 'S0'
    WHEN UPPER(SUBSTR(TRIM(CAST(model_suffix AS VARCHAR)), 7, 1)) = 'A' THEN 'SA'
    ELSE 'OTHER'
END
"""

ProgressCallback = Callable[[float, str], None]
CustomRuleHandler = Callable[[pd.Series, pd.DataFrame], pd.DataFrame]


@dataclass
class ScenarioInputs:
    vi_item_master: pd.DataFrame
    recommendation_request: pd.DataFrame
    recommendation_result: pd.DataFrame
    vi_target_rule: pd.DataFrame
    volume: pd.DataFrame
    raw_vi_target_rule: pd.DataFrame
    final_vi_target_rule: pd.DataFrame
    final_vi_target_rule_audit: pd.DataFrame
    validation_messages: list[str]


CONFIRMED_USER_VALUES = {"Y", "YES", "TRUE", "1", "확정", "적용"}
FINAL_VI_TARGET_RULE_AUDIT_COLUMNS = [
    "source_sheet",
    "vi_item_id",
    "rule_seq",
    "enabled",
    "rule_type",
    "subsidiary",
    "base_part_no",
    "new_part_no",
    "model_type",
    "indoor_tool",
    "desc_match_mode",
    "desc_contains_any",
    "spec_contains_any",
    "parent_assy_desc_contains_any",
    "note",
]


def require_duckdb() -> None:
    if duckdb is None:
        raise RuntimeError("duckdb is required. Install it with: pip install duckdb") from DUCKDB_IMPORT_ERROR


def dashboard_outputs_dir() -> Path:
    return dashboard_output_root()


def duckdb_temp_dir() -> Path:
    return duckdb_temp_root()


def create_duckdb_connection() -> DuckDBConnection:
    require_duckdb()
    temp_dir = duckdb_temp_dir()
    temp_dir.mkdir(parents=True, exist_ok=True)
    con = duckdb.connect()
    con.execute("PRAGMA threads=2")
    con.execute("PRAGMA preserve_insertion_order=false")
    con.execute(f"PRAGMA temp_directory='{temp_dir.as_posix()}'")
    return con


def script_dir() -> Path:
    return application_root()


def default_lake_root(root: Path | None = None) -> Path:
    return Path(root) if root is not None else runtime_bom_lake_root()


def _snapshot_glob(root: Path | None, snapshot_kind: str) -> str:
    lake_root = default_lake_root(root)
    if snapshot_kind == "current_cache":
        return str(lake_root / CURRENT_CACHE_DIR_NAME / "bom_year=*" / "bom_month=*" / "*.parquet")
    return str(lake_root / HISTORICAL_DIR_NAME / "bom_year=*" / "bom_month=*" / "*.parquet")


def _normalize_text(value: object) -> str:
    return "" if value is None else str(value).strip()


def _normalize_key(value: object) -> str:
    return _normalize_text(value).lower()


def _normalize_filter_value(value: object) -> str:
    text = _normalize_text(value)
    if text.upper() in {"", "ALL", "ANY", "*", "전체"}:
        return ""
    return text


def _normalize_rmc_flag_value(value: object) -> str:
    text = _normalize_text(value).upper()
    if text in {"Y", "YES", "TRUE", "1"}:
        return "Y"
    return "N"


def _sql_identifier(identifier: str) -> str:
    escaped = str(identifier).replace('"', '""')
    return f'"{escaped}"'


def normalize_filter_value(value: object) -> str:
    return _normalize_filter_value(value)


def _split_filter_values(value: object) -> list[str]:
    text = _normalize_filter_value(value)
    if not text:
        return []
    normalized = text.replace("\n", ",").replace(";", ",").replace("|", ",")
    values: list[str] = []
    seen: set[str] = set()
    for part in normalized.split(","):
        item = _normalize_text(part)
        if not item:
            continue
        key = item.upper()
        if key in seen:
            continue
        seen.add(key)
        values.append(item)
    return values


def _split_filter_values_upper(value: object) -> list[str]:
    return [item.upper() for item in _split_filter_values(value)]


def _append_in_condition(
    conditions: list[str],
    params: list[object],
    column_sql: str,
    values: list[str],
) -> None:
    usable_values = [str(value).strip() for value in values if str(value).strip()]
    if not usable_values:
        return
    if len(usable_values) == 1:
        conditions.append(f"{column_sql} = ?")
        params.append(usable_values[0])
        return
    placeholders = ", ".join("?" for _ in usable_values)
    conditions.append(f"{column_sql} IN ({placeholders})")
    params.extend(usable_values)


def _normalize_header(value: object) -> str:
    text = _normalize_key(value)
    compact = text.replace(".", "").replace("_", "").replace("-", "").replace(" ", "")
    return COLUMN_ALIASES.get(compact, COLUMN_ALIASES.get(text, text))


def _normalize_sheet_name(value: object) -> str:
    text = _normalize_key(value)
    compact = text.replace(".", "").replace("_", "").replace("-", "").replace(" ", "")
    return SHEET_ALIASES.get(compact, SHEET_ALIASES.get(text, str(value)))


def normalize_month_value(value: object) -> str:
    text = _normalize_text(value)
    if not text:
        return ""
    parsed = pd.to_datetime(text, errors="coerce")
    if pd.notna(parsed):
        return parsed.strftime("%Y-%m")
    text = text.replace("/", "-").replace(".", "-")
    parts = [part for part in text.split("-") if part]
    if len(parts) >= 2 and parts[0].isdigit() and parts[1].isdigit():
        return f"{int(parts[0]):04d}-{int(parts[1]):02d}"
    return text


def get_model_type(model_suffix: str) -> str:
    text = _normalize_text(model_suffix)
    if len(text) < 4:
        return "UNKNOWN"
    code = text[3].upper()
    if code == "Q":
        return "C/O"
    if code == "W":
        return "H/P"
    return "OTHER"


def get_indoor_tool(model_suffix: str) -> str:
    text = _normalize_text(model_suffix)
    if len(text) < 7:
        return "UNKNOWN"
    code = text[6].upper()
    if code == "J":
        return "SJ"
    if code == "K":
        return "SK"
    if code == "0":
        return "S0"
    if code == "A":
        return "SA"
    return "OTHER"


def _empty_frame(columns: list[str]) -> pd.DataFrame:
    return pd.DataFrame(columns=columns)


def _is_confirmed_recommendation(value: object) -> bool:
    return _normalize_text(value).upper() in CONFIRMED_USER_VALUES


def _empty_rule_audit_frame() -> pd.DataFrame:
    return _empty_frame(FINAL_VI_TARGET_RULE_AUDIT_COLUMNS)


def _notify_progress(progress_callback: ProgressCallback | None, value: float, message: str) -> None:
    if progress_callback is not None:
        progress_callback(max(0.0, min(1.0, float(value))), message)


def _log_timing(label: str, started_at: float) -> None:
    print(f"[Timing] {label}: {time.perf_counter() - started_at:.2f}s", flush=True)


def _log_elapsed(label: str, elapsed_seconds: float) -> None:
    print(f"[Timing] {label}: {float(elapsed_seconds):.2f}s", flush=True)


def _standardize_columns(frame: pd.DataFrame, required_columns: list[str]) -> pd.DataFrame:
    rename_map: dict[str, str] = {}
    claimed: set[str] = set()
    for column in frame.columns:
        normalized = _normalize_header(column)
        if normalized in required_columns and normalized not in claimed:
            rename_map[column] = normalized
            claimed.add(normalized)
    standardized = frame.rename(columns=rename_map).copy()
    for column in required_columns:
        if column not in standardized.columns:
            standardized[column] = ""
    return standardized[required_columns].copy()


def _normalize_string_columns(frame: pd.DataFrame, columns: list[str]) -> pd.DataFrame:
    normalized = frame.copy()
    for column in columns:
        if column not in normalized.columns:
            normalized[column] = ""
        series = normalized[column].astype("string")
        normalized[column] = series.fillna("").str.strip().astype(str)
    return normalized


def _auto_vi_item_id(vi_item_name: object, row_number: int, prefix: str) -> str:
    base_name = _normalize_text(vi_item_name)
    if base_name:
        digest = hashlib.sha1(base_name.encode("utf-8")).hexdigest()[:8].upper()
        return f"{prefix}_{digest}"
    return f"{prefix}_{int(row_number):03d}"


def _ensure_internal_vi_item_ids(frame: pd.DataFrame, *, prefix: str) -> pd.DataFrame:
    if frame.empty:
        return frame.copy()
    working = frame.copy()
    if "vi_item_id" not in working.columns:
        working["vi_item_id"] = ""
    if "vi_item_name" not in working.columns:
        working["vi_item_name"] = ""
    missing_mask = working["vi_item_id"].fillna("").astype(str).str.strip().eq("")
    if not bool(missing_mask.any()):
        return working
    generated_values = [
        _auto_vi_item_id(working.iloc[index].get("vi_item_name", ""), index + 1, prefix)
        for index in range(len(working.index))
    ]
    generated_series = pd.Series(generated_values, index=working.index, dtype="object")
    working.loc[missing_mask, "vi_item_id"] = generated_series.loc[missing_mask]
    return working


@lru_cache(maxsize=8)
def _snapshot_columns(history_root_text: str, snapshot_kind: str = "historical", snapshot_path_text: str = "") -> tuple[str, ...]:
    require_duckdb()
    parquet_source = snapshot_path_text or _snapshot_glob(Path(history_root_text), snapshot_kind)
    con = create_duckdb_connection()
    try:
        frame = con.execute(
            "DESCRIBE SELECT * FROM read_parquet(?)",
            [parquet_source],
        ).df()
    finally:
        con.close()
    return tuple(frame["column_name"].astype(str).tolist())


def _first_available_column(columns: tuple[str, ...], candidates: list[str]) -> str | None:
    lookup = set(columns)
    for candidate in candidates:
        if candidate in lookup:
            return candidate
    return None


def _normalize_desc_match_mode(value: object) -> str:
    text = _normalize_text(value).upper()
    if text in {"ALL", "AND"}:
        return "AND"
    return "OR"


def _build_level_num_expr(level_expr: str) -> str:
    return f"""
    CASE
        WHEN REGEXP_EXTRACT(TRIM({level_expr}), '([0-9]+)$', 1) <> ''
            THEN TRY_CAST(REGEXP_EXTRACT(TRIM({level_expr}), '([0-9]+)$', 1) AS INTEGER)
        WHEN REGEXP_EXTRACT(TRIM({level_expr}), '^([.]+)', 1) <> ''
            THEN LENGTH(REGEXP_EXTRACT(TRIM({level_expr}), '^([.]+)', 1))
        ELSE TRY_CAST(TRY_CAST(TRIM({level_expr}) AS DOUBLE) AS INTEGER)
    END
    """


def _build_recommendation_source_query(columns: tuple[str, ...]) -> str:
    desc_column = _first_available_column(columns, ["description", "desc", "item_desc", "part_desc"])
    material_column = _first_available_column(columns, ["material", "material_name", "resin_material", "raw_material_desc"])
    spec_column = _first_available_column(columns, ["spec_text", "spec", "specification", "material_spec"])
    level_column = _first_available_column(columns, ["level", "lv", "bom_level", "lvl"])
    row_order_column = _first_available_column(columns, ["row_order"])
    desc_upper_column = _first_available_column(columns, ["desc_text_upper"])
    spec_upper_column = _first_available_column(columns, ["spec_text_upper"])
    level_num_column = _first_available_column(columns, ["level_num"])

    desc_expr = f"TRIM(CAST({desc_column} AS VARCHAR))" if desc_column else "''"
    material_expr = f"TRIM(CAST({material_column} AS VARCHAR))" if material_column else "''"
    spec_expr = f"TRIM(CAST({spec_column} AS VARCHAR))" if spec_column else "''"
    level_expr = f"CAST({level_column} AS VARCHAR)" if level_column else "''"
    row_order_expr = f"TRY_CAST({row_order_column} AS BIGINT)" if row_order_column else "ROW_NUMBER() OVER ()"
    desc_upper_expr = (
        f"COALESCE(NULLIF(TRIM(CAST({desc_upper_column} AS VARCHAR)), ''), UPPER(COALESCE(NULLIF({desc_expr}, ''), '')))"
        if desc_upper_column
        else f"UPPER(COALESCE(NULLIF({desc_expr}, ''), ''))"
    )
    spec_upper_expr = (
        f"COALESCE(NULLIF(TRIM(CAST({spec_upper_column} AS VARCHAR)), ''), UPPER(COALESCE(NULLIF({spec_expr}, ''), '')))"
        if spec_upper_column
        else f"UPPER(COALESCE(NULLIF({spec_expr}, ''), ''))"
    )
    level_num_expr = f"COALESCE(TRY_CAST({level_num_column} AS INTEGER), {_build_level_num_expr(level_expr)})" if level_num_column else _build_level_num_expr(level_expr)

    return f"""
        WITH raw_rows AS (
            SELECT
                ROW_NUMBER() OVER () AS raw_row_id,
                COALESCE(NULLIF(TRIM(CAST(subsidiary AS VARCHAR)), ''), 'Unknown') AS subsidiary,
                TRIM(CAST(model_suffix AS VARCHAR)) AS model_suffix,
                TRIM(CAST(part_no AS VARCHAR)) AS part_no,
                {MODEL_TYPE_SQL} AS model_type,
                {INDOOR_TOOL_SQL} AS indoor_tool,
                CASE
                    WHEN LENGTH(TRIM(CAST(model_suffix AS VARCHAR))) >= 3
                     AND SUBSTR(TRIM(CAST(model_suffix AS VARCHAR)), 3, 1) = '-' THEN 'Y'
                    ELSE 'N'
                END AS set_model_flag,
                COALESCE(NULLIF({desc_expr}, ''), '') AS desc_text,
                {desc_upper_expr} AS desc_text_upper,
                COALESCE(NULLIF({material_expr}, ''), '') AS material_text,
                COALESCE(NULLIF({spec_expr}, ''), '') AS spec_text,
                {spec_upper_expr} AS spec_text_upper,
                COALESCE(NULLIF({level_expr}, ''), '') AS level_text,
                {level_num_expr} AS level_num,
                COALESCE(TRY_CAST(seq AS DOUBLE), 0) AS seq,
                COALESCE({row_order_expr}, 0) AS row_order,
                COALESCE(unit_price, 0) AS unit_price,
                COALESCE(qty, 0) AS qty
            FROM read_parquet(?)
            WHERE bom_year = ?
              AND bom_month = ?
              AND model_suffix IS NOT NULL
              AND part_no IS NOT NULL
        ),
        parented AS (
            SELECT
                child.*,
                COALESCE(parent.desc_text, '') AS parent_assy_desc_text,
                UPPER(COALESCE(parent.desc_text, '')) AS parent_assy_desc_upper
            FROM raw_rows child
            LEFT JOIN raw_rows parent
              ON parent.subsidiary = child.subsidiary
             AND parent.model_suffix = child.model_suffix
             AND parent.level_num = child.level_num - 1
             AND (
                    parent.seq < child.seq
                 OR (parent.seq = child.seq AND parent.row_order < child.row_order)
             )
            QUALIFY ROW_NUMBER() OVER (
                PARTITION BY child.raw_row_id
                ORDER BY parent.seq DESC NULLS LAST, parent.row_order DESC NULLS LAST
            ) = 1
        ),
        aggregated AS (
            SELECT
                subsidiary,
                model_suffix,
                part_no,
                model_type,
                indoor_tool,
                set_model_flag,
                desc_text,
                MAX(desc_text_upper) AS desc_text_upper,
                STRING_AGG(DISTINCT NULLIF(material_text, ''), ', ') AS material_text,
                STRING_AGG(DISTINCT NULLIF(spec_text, ''), ', ') AS spec_text,
                MAX(spec_text_upper) AS spec_text_upper,
                STRING_AGG(DISTINCT NULLIF(parent_assy_desc_text, ''), ', ') AS parent_assy_desc_text,
                MAX(parent_assy_desc_upper) AS parent_assy_desc_upper,
                MIN(level_text) AS level_text,
                MIN(seq) AS seq,
                MIN(row_order) AS row_order,
                MAX(COALESCE(unit_price, 0)) AS unit_price,
                SUM(COALESCE(qty, 0)) AS qty
            FROM parented
            GROUP BY 1,2,3,4,5,6,7
        )
        SELECT * FROM aggregated
    """


def _build_recommendation_raw_bom_query(columns: tuple[str, ...]) -> str:
    desc_column = _first_available_column(columns, ["description", "desc", "item_desc", "part_desc"])
    spec_column = _first_available_column(columns, ["spec_text", "spec", "specification", "material_spec"])
    level_column = _first_available_column(columns, ["level", "lv", "bom_level", "lvl"])
    row_order_column = _first_available_column(columns, ["row_order"])
    desc_upper_column = _first_available_column(columns, ["desc_text_upper"])
    spec_upper_column = _first_available_column(columns, ["spec_text_upper"])
    level_num_column = _first_available_column(columns, ["level_num"])

    desc_expr = f"TRIM(CAST({desc_column} AS VARCHAR))" if desc_column else "''"
    spec_expr = f"TRIM(CAST({spec_column} AS VARCHAR))" if spec_column else "''"
    level_expr = f"CAST({level_column} AS VARCHAR)" if level_column else "''"
    row_order_expr = f"TRY_CAST({row_order_column} AS BIGINT)" if row_order_column else "ROW_NUMBER() OVER ()"
    desc_upper_expr = (
        f"COALESCE(NULLIF(TRIM(CAST({desc_upper_column} AS VARCHAR)), ''), UPPER(COALESCE(NULLIF({desc_expr}, ''), '')))"
        if desc_upper_column
        else f"UPPER(COALESCE(NULLIF({desc_expr}, ''), ''))"
    )
    spec_upper_expr = (
        f"COALESCE(NULLIF(TRIM(CAST({spec_upper_column} AS VARCHAR)), ''), UPPER(COALESCE(NULLIF({spec_expr}, ''), '')))"
        if spec_upper_column
        else f"UPPER(COALESCE(NULLIF({spec_expr}, ''), ''))"
    )
    level_num_expr = (
        f"COALESCE(TRY_CAST({level_num_column} AS INTEGER), {_build_level_num_expr(level_expr)})"
        if level_num_column
        else _build_level_num_expr(level_expr)
    )

    return f"""
        SELECT
            ROW_NUMBER() OVER () AS raw_row_id,
            COALESCE(NULLIF(TRIM(CAST(subsidiary AS VARCHAR)), ''), 'Unknown') AS subsidiary,
            TRIM(CAST(model_suffix AS VARCHAR)) AS model_suffix,
            TRIM(CAST(part_no AS VARCHAR)) AS part_no,
            UPPER(TRIM(CAST(part_no AS VARCHAR))) AS part_no_upper,
            {MODEL_TYPE_SQL} AS model_type,
            {INDOOR_TOOL_SQL} AS indoor_tool,
            COALESCE(NULLIF({desc_expr}, ''), '') AS desc_text,
            {desc_upper_expr} AS desc_text_upper,
            {spec_upper_expr} AS spec_text_upper,
            {level_num_expr} AS level_num,
            COALESCE(TRY_CAST(seq AS DOUBLE), 0) AS seq,
            COALESCE({row_order_expr}, 0) AS row_order
        FROM read_parquet(?)
        WHERE bom_year = ?
          AND bom_month = ?
          AND model_suffix IS NOT NULL
          AND part_no IS NOT NULL
    """


def _build_keyword_like_clause(column_sql: str, raw_keywords: object, match_mode: str, params: list[object]) -> str:
    keywords = [keyword.upper() for keyword in _split_keywords(raw_keywords)]
    if not keywords:
        return "TRUE"
    operator = " AND " if match_mode == "AND" else " OR "
    parts: list[str] = []
    for keyword in keywords:
        parts.append(f"{column_sql} LIKE ?")
        params.append(f"%{keyword}%")
    return "(" + operator.join(parts) + ")"


@lru_cache(maxsize=32)
def load_bom_snapshot_recommendation_view(
    history_root_text: str,
    bom_year: int,
    bom_month: int,
    snapshot_kind: str = "historical",
    snapshot_path_text: str = "",
) -> pd.DataFrame:
    require_duckdb()
    parquet_source = snapshot_path_text or _snapshot_glob(Path(history_root_text), snapshot_kind)
    columns = _snapshot_columns(history_root_text, snapshot_kind, snapshot_path_text)
    started_at = time.perf_counter()
    query = _build_recommendation_source_query(columns) + "\nORDER BY subsidiary, model_suffix, part_no"
    con = create_duckdb_connection()
    try:
        frame = con.execute(
            query,
            [parquet_source, int(bom_year), int(bom_month)],
        ).df()
    finally:
        con.close()
    for text_column in ["material_text", "spec_text", "parent_assy_desc_text", "level_text"]:
        if text_column not in frame.columns:
            frame[text_column] = ""
        frame[text_column] = frame[text_column].fillna("").astype(str)
    _log_timing(f"load recommendation snapshot {bom_year}-{bom_month:02d} ({snapshot_kind})", started_at)
    return frame


def _split_keywords(value: object) -> list[str]:
    text = _normalize_text(value)
    if not text:
        return []
    normalized = str(text).replace("\n", ",").replace(";", ",").replace("|", ",")
    return [part.strip().lower() for part in normalized.split(",") if part.strip()]


def _contains_any_mask(series: pd.Series, raw_keywords: object) -> pd.Series:
    keywords = _split_keywords(raw_keywords)
    if not keywords:
        return pd.Series([True] * len(series.index), index=series.index)
    lowered = series.fillna("").astype(str).str.lower()
    mask = pd.Series([False] * len(series.index), index=series.index)
    for keyword in keywords:
        mask = mask | lowered.str.contains(keyword, regex=False)
    return mask


def _contains_all_mask(series: pd.Series, raw_keywords: object) -> pd.Series:
    keywords = _split_keywords(raw_keywords)
    if not keywords:
        return pd.Series([True] * len(series.index), index=series.index)
    lowered = series.fillna("").astype(str).str.lower()
    mask = pd.Series([True] * len(series.index), index=series.index)
    for keyword in keywords:
        mask = mask & lowered.str.contains(keyword, regex=False)
    return mask


def _contains_none_mask(series: pd.Series, raw_keywords: object) -> pd.Series:
    keywords = _split_keywords(raw_keywords)
    if not keywords:
        return pd.Series([True] * len(series.index), index=series.index)
    lowered = series.fillna("").astype(str).str.lower()
    mask = pd.Series([True] * len(series.index), index=series.index)
    for keyword in keywords:
        mask = mask & (~lowered.str.contains(keyword, regex=False))
    return mask


def _format_share_detail(counts_by_subsidiary: dict[str, int], totals_by_subsidiary: dict[str, int]) -> str:
    parts: list[str] = []
    for subsidiary in sorted(set(totals_by_subsidiary) | set(counts_by_subsidiary)):
        matched = int(counts_by_subsidiary.get(subsidiary, 0))
        total = int(totals_by_subsidiary.get(subsidiary, 0))
        pct = (matched / total * 100.0) if total > 0 else 0.0
        parts.append(f"{subsidiary} {pct:.1f}% ({matched}/{total})")
    return ", ".join(parts)


def _format_model_count_detail(counts_by_subsidiary: dict[str, int]) -> str:
    return ", ".join(
        f"{subsidiary} {int(counts_by_subsidiary[subsidiary])}"
        for subsidiary in sorted(counts_by_subsidiary)
    )


def _format_model_list(values: list[str]) -> str:
    return ", ".join(sorted({str(value).strip() for value in values if str(value).strip()}))


def build_base_pn_recommendations(
    history_root: Path | None,
    base_year: int,
    base_month: int,
    scenario_workbook_path: Path,
    item_master_sheet: str | None = None,
    recommendation_request_sheet: str | None = None,
    base_snapshot_kind: str = "historical",
    base_snapshot_path: Path | None = None,
) -> pd.DataFrame:
    workbook_started_at = time.perf_counter()
    scenario_inputs = load_scenario_workbook(
        scenario_workbook_path,
        item_master_sheet=item_master_sheet,
        recommendation_request_sheet=recommendation_request_sheet,
    )
    _log_timing("load recommendation workbook", workbook_started_at)

    request = scenario_inputs.recommendation_request.copy()
    if request.empty:
        return _empty_frame(RECOMMENDATION_RESULT_COLUMNS)

    request["enabled"] = request["enabled"].replace("", "Y").astype(str).str.upper()
    request = request[request["enabled"] != "N"].copy()
    request["subsidiary"] = request["subsidiary"].map(normalize_filter_value)
    request["desc_match_mode"] = request["desc_match_mode"].map(_normalize_desc_match_mode)

    def has_recommendation_filters(row: pd.Series) -> bool:
        return any(
            [
                bool(_split_filter_values(row.get("subsidiary", ""))),
                bool(_normalize_text(row.get("base_part_no", ""))),
                bool(_normalize_text(row.get("new_part_no", ""))),
                bool(_normalize_filter_value(row.get("model_type", ""))),
                bool(_normalize_filter_value(row.get("indoor_tool", ""))),
                bool(_normalize_filter_value(row.get("desc_contains_any", ""))),
                bool(_normalize_filter_value(row.get("spec_contains_any", ""))),
                bool(_normalize_filter_value(row.get("parent_assy_desc_contains_any", ""))),
            ]
        )

    request = request[request.apply(has_recommendation_filters, axis=1)].copy()
    if request.empty:
        return _empty_frame(RECOMMENDATION_RESULT_COLUMNS)

    history_root_text = str(default_lake_root(history_root))
    parquet_source = str(base_snapshot_path) if base_snapshot_path is not None else _snapshot_glob(Path(history_root_text), base_snapshot_kind)
    columns = _snapshot_columns(history_root_text, base_snapshot_kind, str(base_snapshot_path) if base_snapshot_path is not None else "")
    raw_bom_query = _build_recommendation_raw_bom_query(columns)

    con = create_duckdb_connection()
    try:
        print(
            f"[Recommendation] Filtering start | parquet={parquet_source} | month={base_year:04d}-{base_month:02d}",
            flush=True,
        )
        print("[DuckDB] Materializing recommendation raw_bom", flush=True)
        materialize_started_at = time.perf_counter()
        con.execute(
            f"""
            CREATE OR REPLACE TEMP TABLE recommendation_raw_bom AS
            {raw_bom_query}
            """,
            [parquet_source, int(base_year), int(base_month)],
        )
        _log_timing("materialize recommendation raw_bom", materialize_started_at)
        raw_bom_row_count = int(con.execute("SELECT COUNT(*) FROM recommendation_raw_bom").fetchone()[0] or 0)
        print(f"[DuckDB] recommendation raw_bom rows: {raw_bom_row_count:,}", flush=True)
        base_bom_total_model_count = int(
            con.execute("SELECT COUNT(DISTINCT model_suffix) FROM recommendation_raw_bom").fetchone()[0] or 0
        )
        print("[DuckDB] Executing recommendation BOM summary SQL", flush=True)
        bom_totals_rows = con.execute(
            """
            SELECT subsidiary, COUNT(DISTINCT model_suffix) AS model_count
            FROM recommendation_raw_bom
            GROUP BY 1
            ORDER BY 1
            """
        ).fetchall()
        bom_totals_by_subsidiary = {
            str(subsidiary_name): int(model_count)
            for subsidiary_name, model_count in bom_totals_rows
        }
        bom_subsidiary_model_count_detail = _format_model_count_detail(bom_totals_by_subsidiary)
        if not bom_totals_by_subsidiary:
            return _empty_frame(RECOMMENDATION_RESULT_COLUMNS)

        result_rows: list[pd.DataFrame] = []
        recommend_started_at = time.perf_counter()

        for request_row in request.to_dict("records"):
            row = pd.Series(request_row)
            subsidiaries = _split_filter_values(row.get("subsidiary", ""))
            subsidiary = ", ".join(subsidiaries)
            base_part_no = _normalize_text(row.get("base_part_no", ""))
            direct_new_part_no = _normalize_text(row.get("new_part_no", ""))
            model_types = _split_filter_values_upper(row.get("model_type", ""))
            indoor_tools = _split_filter_values_upper(row.get("indoor_tool", ""))
            model_type = ", ".join(model_types)
            indoor_tool = ", ".join(indoor_tools)
            desc_match_mode = _normalize_desc_match_mode(row.get("desc_match_mode", ""))
            desc_keywords = _normalize_filter_value(row.get("desc_contains_any", ""))
            spec_keywords = _normalize_filter_value(row.get("spec_contains_any", ""))
            parent_assy_desc_keywords = _normalize_filter_value(row.get("parent_assy_desc_contains_any", ""))

            scope_conditions = ["TRUE"]
            scope_params: list[object] = []
            candidate_base_conditions = ["TRUE"]
            candidate_base_params: list[object] = []

            if subsidiaries:
                placeholders = ", ".join(["?"] * len(subsidiaries))
                scope_conditions.append(f"subsidiary IN ({placeholders})")
                scope_params.extend(subsidiaries)
                candidate_base_conditions.append(f"subsidiary IN ({placeholders})")
                candidate_base_params.extend(subsidiaries)
            if model_types:
                _append_in_condition(scope_conditions, scope_params, "model_type", model_types)
                _append_in_condition(candidate_base_conditions, candidate_base_params, "model_type", model_types)
            if indoor_tools:
                _append_in_condition(scope_conditions, scope_params, "indoor_tool", indoor_tools)
                _append_in_condition(candidate_base_conditions, candidate_base_params, "indoor_tool", indoor_tools)
            if base_part_no:
                candidate_base_conditions.append("part_no_upper = ?")
                candidate_base_params.append(base_part_no.upper())

            candidate_base_conditions.append(
                _build_keyword_like_clause("desc_text_upper", desc_keywords, desc_match_mode, candidate_base_params)
            )
            candidate_base_conditions.append(
                _build_keyword_like_clause("spec_text_upper", spec_keywords, desc_match_mode, candidate_base_params)
            )

            print(
                "[DuckDB] Executing recommendation filter SQL "
                f"| vi_item_id={_normalize_text(row.get('vi_item_id', ''))} "
                f"| base_part_no={base_part_no or '-'} "
                f"| model_type={model_type or '-'} "
                f"| indoor_tool={indoor_tool or '-'} "
                f"| desc_match={desc_match_mode}",
                flush=True,
            )
            con.execute(
                f"""
                CREATE OR REPLACE TEMP TABLE recommendation_candidate_base AS
                SELECT
                    raw_row_id,
                    subsidiary,
                    model_suffix,
                    part_no,
                    part_no_upper,
                    model_type,
                    indoor_tool,
                    desc_text,
                    desc_text_upper,
                    spec_text_upper,
                    level_num,
                    seq,
                    row_order
                FROM recommendation_raw_bom
                WHERE {' AND '.join(candidate_base_conditions)}
                """,
                candidate_base_params,
            )
            candidate_base_count = int(
                con.execute("SELECT COUNT(*) FROM recommendation_candidate_base").fetchone()[0] or 0
            )
            print(
                f"[DuckDB] candidate_base rows | vi_item_id={_normalize_text(row.get('vi_item_id', ''))} "
                f"| rows={candidate_base_count:,}",
                flush=True,
            )

            if parent_assy_desc_keywords:
                parent_filter_params: list[object] = []
                parent_filter_clause = _build_keyword_like_clause(
                    "parent_assy_desc_upper",
                    parent_assy_desc_keywords,
                    desc_match_mode,
                    parent_filter_params,
                )
                con.execute(
                    f"""
                    CREATE OR REPLACE TEMP TABLE recommendation_candidate_filtered AS
                    WITH parented AS (
                        SELECT
                            candidate.raw_row_id,
                            candidate.part_no,
                            candidate.desc_text,
                            candidate.subsidiary,
                            candidate.model_suffix,
                            COALESCE(parent.part_no, '') AS parent_assy_part_no,
                            COALESCE(parent.desc_text, '') AS parent_assy_desc_text,
                            UPPER(COALESCE(parent.desc_text, '')) AS parent_assy_desc_upper
                        FROM recommendation_candidate_base candidate
                        LEFT JOIN recommendation_raw_bom parent
                          ON parent.subsidiary = candidate.subsidiary
                         AND parent.model_suffix = candidate.model_suffix
                         AND parent.level_num = candidate.level_num - 1
                         AND (
                                parent.seq < candidate.seq
                             OR (parent.seq = candidate.seq AND parent.row_order < candidate.row_order)
                         )
                        QUALIFY ROW_NUMBER() OVER (
                            PARTITION BY candidate.raw_row_id
                            ORDER BY parent.seq DESC NULLS LAST, parent.row_order DESC NULLS LAST
                        ) = 1
                    )
                    SELECT
                        raw_row_id,
                        part_no,
                        desc_text,
                        subsidiary,
                        model_suffix,
                        parent_assy_part_no,
                        parent_assy_desc_text
                    FROM parented
                    WHERE {parent_filter_clause}
                    """,
                    parent_filter_params,
                )
                print(
                    f"[DuckDB] parent join used | vi_item_id={_normalize_text(row.get('vi_item_id', ''))}",
                    flush=True,
                )
            else:
                con.execute(
                    """
                    CREATE OR REPLACE TEMP TABLE recommendation_candidate_filtered AS
                    SELECT
                        raw_row_id,
                        part_no,
                        desc_text,
                        subsidiary,
                        model_suffix,
                        '' AS parent_assy_part_no,
                        '' AS parent_assy_desc_text
                    FROM recommendation_candidate_base
                    """
                )
                print(
                    f"[DuckDB] parent join skipped | vi_item_id={_normalize_text(row.get('vi_item_id', ''))}",
                    flush=True,
                )

            target_model_count_rows = con.execute(
                """
                SELECT subsidiary, COUNT(DISTINCT model_suffix) AS model_count
                FROM recommendation_candidate_filtered
                GROUP BY 1
                ORDER BY 1
                """
            ).fetchall()
            totals_by_subsidiary = {
                str(subsidiary_name): int(model_count)
                for subsidiary_name, model_count in target_model_count_rows
            }
            total_target_model_count = int(
                con.execute(
                    """
                    SELECT COUNT(DISTINCT model_suffix)
                    FROM recommendation_candidate_filtered
                    """
                ).fetchone()[0]
                or 0
            )
            item_target_subsidiary_model_count_detail = _format_model_count_detail(totals_by_subsidiary)
            candidate_filtered_count = int(
                con.execute("SELECT COUNT(*) FROM recommendation_candidate_filtered").fetchone()[0] or 0
            )
            print(
                f"[DuckDB] candidate rows after parent filter | vi_item_id={_normalize_text(row.get('vi_item_id', ''))} "
                f"| rows={candidate_filtered_count:,}",
                flush=True,
            )

            filtered_detail = con.execute(
                """
                SELECT DISTINCT raw_row_id, part_no, desc_text, subsidiary, model_suffix, parent_assy_part_no, parent_assy_desc_text
                FROM recommendation_candidate_filtered
                ORDER BY part_no, subsidiary, model_suffix, parent_assy_part_no, parent_assy_desc_text
                """
            ).df()

            if filtered_detail.empty:
                reason_parts = ["NO_MATCH"]
                if base_part_no:
                    reason_parts.append(f"base_part_no={base_part_no}")
                if model_type:
                    reason_parts.append(f"model_type={model_type}")
                if indoor_tool:
                    reason_parts.append(f"indoor_tool={indoor_tool}")
                if desc_keywords:
                    reason_parts.append(f"desc={desc_keywords}")
                    reason_parts.append(f"desc_match={desc_match_mode}")
                if spec_keywords:
                    reason_parts.append(f"spec={spec_keywords}")
                if parent_assy_desc_keywords:
                    reason_parts.append(f"parent_assy_desc={parent_assy_desc_keywords}")
                result_rows.append(
                    pd.DataFrame(
                        [
                            {
                                "vi_item_id": _normalize_text(row.get("vi_item_id", "")),
                                "model_suffix": "",
                                "subsidiary": subsidiary,
                                "parent_assy_part_no": "",
                                "parent_assy_desc_text": "",
                                "ancestor_desc_path": "",
                                "recommended_base_part_no": base_part_no,
                                "recommended_desc": "",
                                "confirmed_new_part_no": direct_new_part_no,
                                "user_confirmed": "Y",
                                "base_part_no": base_part_no,
                                "new_part_no": direct_new_part_no,
                                "note": _normalize_text(row.get("note", "")),
                                "desc_match_mode": desc_match_mode,
                                "desc_contains_any": desc_keywords,
                                "spec_contains_any": spec_keywords,
                                "parent_assy_desc_contains_any": parent_assy_desc_keywords,
                                "model_type": model_type,
                                "indoor_tool": indoor_tool,
                                "vi_item_name": _normalize_text(row.get("vi_item_name", "")),
                                "recommended_subsidiaries": "",
                                "bom_subsidiary_model_count_detail": bom_subsidiary_model_count_detail,
                                "base_bom_total_model_count": base_bom_total_model_count,
                                "item_target_model_count": total_target_model_count,
                                "item_target_subsidiary_model_count_detail": item_target_subsidiary_model_count_detail,
                                "recommended_model_count": 0,
                                "recommended_model_share_formula": "recommended_model_count / item_target_model_count",
                                "recommended_model_share_pct": "0.0%",
                                "recommended_subsidiary_share_detail": "",
                                "base_pn_subsidiaries": "",
                                "base_pn_models": "",
                                "expected_vi_amount_usd": "",
                                "_volume_seed_pairs": "",
                                "recommendation_reason": ", ".join(reason_parts),
                            }
                        ]
                    )
                )
                continue

            matched_by_subsidiary = (
                filtered_detail.groupby(["part_no", "desc_text", "subsidiary"])["model_suffix"]
                .nunique()
                .reset_index(name="matched_model_count")
            )
            subsidiary_share_detail = (
                matched_by_subsidiary.groupby(["part_no", "desc_text"], as_index=False)
                .apply(
                    lambda frame: _format_share_detail(
                        {
                            str(subsidiary_name): int(count)
                            for subsidiary_name, count in zip(frame["subsidiary"], frame["matched_model_count"])
                        },
                        totals_by_subsidiary,
                    ),
                    include_groups=False,
                )
                .rename(columns={None: "recommended_subsidiary_share_detail"})
            )
            grouped = (
                filtered_detail.groupby(["part_no", "desc_text"], as_index=False)
                .agg(
                    recommended_model_count=("model_suffix", "nunique"),
                    recommended_subsidiaries=("subsidiary", lambda s: _format_model_list(s.tolist())),
                    base_pn_subsidiaries=("subsidiary", lambda s: _format_model_list(s.tolist())),
                    base_pn_models=("model_suffix", lambda s: _format_model_list(s.tolist())),
                )
                .sort_values(["recommended_model_count", "part_no"], ascending=[False, True])
                .reset_index(drop=True)
            )
            volume_seed_by_key = (
                filtered_detail.assign(
                    _volume_seed_pair=filtered_detail["subsidiary"].astype(str).str.strip() + "\t" + filtered_detail["model_suffix"].astype(str).str.strip()
                )
                .groupby(["part_no", "desc_text"])["_volume_seed_pair"]
                .apply(lambda series: "\n".join(sorted({value for value in series.tolist() if value.strip()})))
                .reset_index(name="_volume_seed_pairs")
            )
            detailed_rows = filtered_detail.rename(
                columns={
                    "part_no": "recommended_base_part_no",
                    "desc_text": "recommended_desc",
                }
            ).copy()
            detailed_rows["ancestor_desc_path"] = ""
            detailed_rows = detailed_rows.merge(
                grouped,
                how="left",
                left_on=["recommended_base_part_no", "recommended_desc"],
                right_on=["part_no", "desc_text"],
            ).drop(columns=["part_no", "desc_text"], errors="ignore")
            detailed_rows = detailed_rows.merge(
                subsidiary_share_detail,
                how="left",
                left_on=["recommended_base_part_no", "recommended_desc"],
                right_on=["part_no", "desc_text"],
            ).drop(columns=["part_no", "desc_text"], errors="ignore")
            detailed_rows = detailed_rows.drop(columns=["_volume_seed_pairs"], errors="ignore").merge(
                volume_seed_by_key,
                how="left",
                left_on=["recommended_base_part_no", "recommended_desc"],
                right_on=["part_no", "desc_text"],
            )
            detailed_rows = detailed_rows.drop(columns=["part_no", "desc_text"], errors="ignore")
            detailed_rows["_volume_seed_pairs"] = detailed_rows["_volume_seed_pairs"].fillna("")
            reason_parts: list[str] = []
            if base_part_no:
                reason_parts.append(f"base_part_no={base_part_no}")
            if model_type:
                reason_parts.append(f"model_type={model_type}")
            if indoor_tool:
                reason_parts.append(f"indoor_tool={indoor_tool}")
            if desc_keywords:
                reason_parts.append(f"desc={desc_keywords}")
                reason_parts.append(f"desc_match={desc_match_mode}")
            if spec_keywords:
                reason_parts.append(f"spec={spec_keywords}")
            if parent_assy_desc_keywords:
                reason_parts.append(f"parent_assy_desc={parent_assy_desc_keywords}")
            detailed_rows["vi_item_id"] = _normalize_text(row.get("vi_item_id", ""))
            detailed_rows["new_part_no"] = direct_new_part_no
            detailed_rows["note"] = _normalize_text(row.get("note", ""))
            detailed_rows["desc_match_mode"] = desc_match_mode
            detailed_rows["desc_contains_any"] = desc_keywords
            detailed_rows["spec_contains_any"] = spec_keywords
            detailed_rows["parent_assy_desc_contains_any"] = parent_assy_desc_keywords
            detailed_rows["model_type"] = model_type
            detailed_rows["indoor_tool"] = indoor_tool
            detailed_rows["vi_item_name"] = _normalize_text(row.get("vi_item_name", ""))
            detailed_rows["bom_subsidiary_model_count_detail"] = bom_subsidiary_model_count_detail
            detailed_rows["base_bom_total_model_count"] = base_bom_total_model_count
            detailed_rows["item_target_model_count"] = total_target_model_count
            detailed_rows["item_target_subsidiary_model_count_detail"] = item_target_subsidiary_model_count_detail
            detailed_rows["recommended_model_share_formula"] = "recommended_model_count / item_target_model_count"
            detailed_rows["recommended_model_share_pct"] = detailed_rows["recommended_model_count"].map(
                lambda count: f"{(float(count) / total_target_model_count * 100.0):.1f}%"
                if total_target_model_count > 0
                else "0.0%"
            )
            detailed_rows["recommendation_reason"] = ", ".join(reason_parts)
            detailed_rows["user_confirmed"] = "Y"
            detailed_rows["confirmed_new_part_no"] = direct_new_part_no
            detailed_rows["base_part_no"] = detailed_rows["recommended_base_part_no"]
            detailed_rows["expected_vi_amount_usd"] = ""
            result_rows.append(detailed_rows)

        _log_timing("build base pn recommendations", recommend_started_at)
        print("[Recommendation] Filtering end", flush=True)
    finally:
        con.close()

    if not result_rows:
        return _empty_frame(RECOMMENDATION_RESULT_COLUMNS)

    result = pd.concat(result_rows, ignore_index=True)
    for column in RECOMMENDATION_RESULT_COLUMNS:
        if column not in result.columns:
            result[column] = ""
    ordered_columns = RECOMMENDATION_RESULT_COLUMNS + [
        column for column in result.columns if column not in RECOMMENDATION_RESULT_COLUMNS
    ]
    return result[ordered_columns].copy()


def build_rules_from_recommendation_result(recommendation_result: pd.DataFrame) -> pd.DataFrame:
    rules, _, _ = _build_rules_from_recommendation_result_meta(recommendation_result)
    return rules


def format_recommendation_result_for_review(recommendation_result: pd.DataFrame) -> pd.DataFrame:
    if recommendation_result.empty:
        return _empty_frame(RECOMMENDATION_REVIEW_EXPORT_COLUMNS)

    standardized = _standardize_columns(recommendation_result, RECOMMENDATION_RESULT_COLUMNS)
    standardized = _normalize_string_columns(standardized, RECOMMENDATION_RESULT_COLUMNS)
    return standardized[RECOMMENDATION_REVIEW_EXPORT_COLUMNS].copy()


def find_invalid_confirmed_recommendation_rows(recommendation_result: pd.DataFrame) -> list[str]:
    if recommendation_result.empty:
        return []

    standardized = _standardize_columns(recommendation_result, RECOMMENDATION_RESULT_COLUMNS)
    standardized = _normalize_string_columns(standardized, RECOMMENDATION_RESULT_COLUMNS)
    invalid_rows: list[str] = []
    for row_number, (_, row) in enumerate(standardized.iterrows(), start=1):
        if not _is_confirmed_recommendation(row.get("user_confirmed", "")):
            continue
        new_part_no = _normalize_text(row.get("confirmed_new_part_no", "")) or _normalize_text(row.get("new_part_no", ""))
        if new_part_no:
            continue
        invalid_rows.append(
            f"행 {row_number}: VI 아이템 ID={_normalize_text(row.get('vi_item_id', '')) or '-'}, "
            f"대상 법인={_normalize_text(row.get('subsidiary', '')) or '-'}, "
            f"추천 Base P/N={_normalize_text(row.get('recommended_base_part_no', '')) or _normalize_text(row.get('base_part_no', '')) or '-'}"
        )
    return invalid_rows


def _build_rules_from_recommendation_result_meta(
    recommendation_result: pd.DataFrame,
) -> tuple[pd.DataFrame, pd.DataFrame, list[str]]:
    if recommendation_result.empty:
        return _empty_frame(VI_TARGET_RULE_COLUMNS), _empty_rule_audit_frame(), []

    warnings: list[str] = []
    rows: list[dict[str, str]] = []
    audit_rows: list[dict[str, str]] = []
    standardized = _standardize_columns(recommendation_result, RECOMMENDATION_RESULT_COLUMNS)
    standardized = _normalize_string_columns(standardized, RECOMMENDATION_RESULT_COLUMNS)
    standardized["subsidiary"] = standardized["subsidiary"].map(normalize_filter_value)

    for row_number, (_, row) in enumerate(standardized.iterrows(), start=1):
        if not _is_confirmed_recommendation(row.get("user_confirmed", "")):
            continue

        vi_item_id = _normalize_text(row.get("vi_item_id", "")) or _auto_vi_item_id(
            row.get("vi_item_name", ""),
            row_number,
            "REC",
        )
        subsidiary = normalize_filter_value(row.get("subsidiary", ""))
        base_part_no = _normalize_text(row.get("recommended_base_part_no", "")) or _normalize_text(row.get("base_part_no", ""))
        new_part_no = _normalize_text(row.get("confirmed_new_part_no", "")) or _normalize_text(row.get("new_part_no", ""))

        if not base_part_no:
            warnings.append(f"Recommendation row {row_number}: user_confirmed is Y but recommended_base_part_no is blank.")
            continue
        if not new_part_no:
            warnings.append(f"Recommendation row {row_number}: user_confirmed is Y but New P/N is blank.")
            continue
        if not subsidiary:
            warnings.append(f"Recommendation row {row_number}: blank subsidiary treated as ALL.")

        note = "from BasePN_Recommendation_Result"

        built_row = {column: "" for column in VI_TARGET_RULE_COLUMNS}
        built_row.update(
            {
                "vi_item_id": vi_item_id,
                "rule_seq": f"REC_{row_number}",
                "enabled": "Y",
                "rule_type": "PART_NO",
                "subsidiary": subsidiary,
                "base_part_no": base_part_no,
                "new_part_no": new_part_no,
                "desc_match_mode": _normalize_desc_match_mode(row.get("desc_match_mode", "")),
                "desc_contains_any": _normalize_text(row.get("desc_contains_any", "")),
                "spec_contains_any": _normalize_text(row.get("spec_contains_any", "")),
                "parent_assy_desc_contains_any": _normalize_text(row.get("parent_assy_desc_contains_any", "")),
                "model_type": _normalize_text(row.get("model_type", "")),
                "indoor_tool": _normalize_text(row.get("indoor_tool", "")),
                "note": note,
            }
        )
        rows.append(built_row)
        audit_rows.append(
            {
                "source_sheet": "BasePN_Recommendation_Result",
                "vi_item_id": vi_item_id,
                "rule_seq": built_row["rule_seq"],
                "enabled": built_row["enabled"],
                "rule_type": built_row["rule_type"],
                "subsidiary": subsidiary,
                "base_part_no": base_part_no,
                "new_part_no": new_part_no,
                "model_type": _normalize_text(row.get("model_type", "")),
                "indoor_tool": _normalize_text(row.get("indoor_tool", "")),
                "desc_match_mode": _normalize_desc_match_mode(row.get("desc_match_mode", "")),
                "desc_contains_any": _normalize_text(row.get("desc_contains_any", "")),
                "spec_contains_any": _normalize_text(row.get("spec_contains_any", "")),
                "parent_assy_desc_contains_any": _normalize_text(row.get("parent_assy_desc_contains_any", "")),
                "note": note,
            }
        )

    rules = pd.DataFrame(rows, columns=VI_TARGET_RULE_COLUMNS) if rows else _empty_frame(VI_TARGET_RULE_COLUMNS)
    audit = pd.DataFrame(audit_rows, columns=FINAL_VI_TARGET_RULE_AUDIT_COLUMNS) if audit_rows else _empty_rule_audit_frame()
    return rules, audit, warnings


def _build_final_vi_target_rule(
    vi_target_rule: pd.DataFrame,
    recommendation_result: pd.DataFrame,
) -> tuple[pd.DataFrame, pd.DataFrame, list[str]]:
    validation_messages: list[str] = []

    recommendation_rules, recommendation_audit, recommendation_warnings = _build_rules_from_recommendation_result_meta(
        recommendation_result
    )
    validation_messages.extend(recommendation_warnings)

    if not recommendation_rules.empty:
        return recommendation_rules, recommendation_audit, validation_messages

    original_rules = _standardize_columns(vi_target_rule, VI_TARGET_RULE_COLUMNS)
    original_rules = _normalize_string_columns(original_rules, VI_TARGET_RULE_COLUMNS)
    original_rules["subsidiary"] = original_rules["subsidiary"].map(normalize_filter_value)
    original_rules = original_rules[
        original_rules["vi_item_id"].ne("")
        & original_rules["base_part_no"].ne("")
        & original_rules["new_part_no"].ne("")
        & original_rules["enabled"].fillna("").astype(str).str.strip().str.upper().ne("N")
    ].copy()
    if original_rules.empty:
        return _empty_frame(VI_TARGET_RULE_COLUMNS), _empty_rule_audit_frame(), validation_messages

    validation_messages.append(
        "No usable confirmed rows in BasePN_Recommendation_Result. Using existing VI_Target_Rule rows for backward compatibility."
    )
    original_rules = original_rules.copy()
    original_rules["source_sheet"] = "VI_Target_Rule"
    final_rule = original_rules[VI_TARGET_RULE_COLUMNS].copy()
    final_audit = original_rules[
        [
            "source_sheet",
            "vi_item_id",
            "rule_seq",
            "enabled",
            "rule_type",
            "subsidiary",
            "base_part_no",
            "new_part_no",
            "model_type",
            "indoor_tool",
            "desc_match_mode",
            "desc_contains_any",
            "spec_contains_any",
            "parent_assy_desc_contains_any",
            "note",
        ]
    ].copy()
    return final_rule, final_audit, validation_messages


def load_scenario_workbook(
    workbook_path: Path,
    item_master_sheet: str | None = None,
    recommendation_request_sheet: str | None = None,
    recommendation_result_sheet: str | None = None,
    target_rule_sheet: str | None = None,
    volume_sheet: str | None = None,
) -> ScenarioInputs:
    sheets = pd.read_excel(workbook_path, sheet_name=None, engine="openpyxl", dtype=object)
    normalized_lookup = {_normalize_sheet_name(name): name for name in sheets}

    item_key = item_master_sheet or "VI_Item_Master"
    recommendation_request_key = recommendation_request_sheet or "BasePN_Recommendation_Request"
    recommendation_result_key = recommendation_result_sheet or "BasePN_Recommendation_Result"
    rule_key = target_rule_sheet or "VI_Target_Rule"
    volume_key = volume_sheet or "Volume"

    item_source = sheets.get(normalized_lookup.get(_normalize_sheet_name(item_key), item_key), pd.DataFrame())
    recommendation_request_source = sheets.get(
        normalized_lookup.get(_normalize_sheet_name(recommendation_request_key), recommendation_request_key),
        pd.DataFrame(),
    )
    recommendation_result_source = sheets.get(
        normalized_lookup.get(_normalize_sheet_name(recommendation_result_key), recommendation_result_key),
        pd.DataFrame(),
    )
    rule_source = sheets.get(normalized_lookup.get(_normalize_sheet_name(rule_key), rule_key), pd.DataFrame())
    volume_source = sheets.get(normalized_lookup.get(_normalize_sheet_name(volume_key), volume_key), pd.DataFrame())

    recommendation_request_source = recommendation_request_source.rename(columns={"사용 여부": "enabled"})
    recommendation_result_source = recommendation_result_source.rename(columns={"사용 여부": "user_confirmed"})

    vi_item_master = _standardize_columns(item_source, VI_ITEM_MASTER_COLUMNS)
    recommendation_request = _standardize_columns(recommendation_request_source, RECOMMENDATION_REQUEST_COLUMNS)
    recommendation_result = _standardize_columns(recommendation_result_source, RECOMMENDATION_RESULT_COLUMNS)
    vi_target_rule = _standardize_columns(rule_source, VI_TARGET_RULE_COLUMNS)
    volume = _standardize_columns(volume_source, VOLUME_COLUMNS)

    vi_item_master = _normalize_string_columns(vi_item_master, VI_ITEM_MASTER_COLUMNS)
    recommendation_request = _normalize_string_columns(recommendation_request, RECOMMENDATION_REQUEST_COLUMNS)
    recommendation_result = _normalize_string_columns(recommendation_result, RECOMMENDATION_RESULT_COLUMNS)
    vi_target_rule = _normalize_string_columns(vi_target_rule, VI_TARGET_RULE_COLUMNS)

    vi_item_master = _ensure_internal_vi_item_ids(vi_item_master, prefix="MASTER")
    recommendation_request = _ensure_internal_vi_item_ids(recommendation_request, prefix="REQ")
    recommendation_result = _ensure_internal_vi_item_ids(recommendation_result, prefix="REC")
    vi_target_rule = _ensure_internal_vi_item_ids(vi_target_rule, prefix="RULE")

    vi_item_master = vi_item_master[
        vi_item_master["vi_item_id"].ne("") | vi_item_master["vi_item_name"].ne("")
    ].copy()
    recommendation_request = recommendation_request[recommendation_request["vi_item_name"].ne("")].copy()
    vi_target_rule = vi_target_rule[
        vi_target_rule["vi_item_id"].ne("") | vi_target_rule["vi_item_name"].ne("")
    ].copy()

    recommendation_request["subsidiary"] = recommendation_request["subsidiary"].map(normalize_filter_value)
    recommendation_result["subsidiary"] = recommendation_result["subsidiary"].map(normalize_filter_value)
    vi_target_rule["subsidiary"] = vi_target_rule["subsidiary"].map(normalize_filter_value)
    recommendation_request["desc_match_mode"] = recommendation_request["desc_match_mode"].map(_normalize_desc_match_mode)
    recommendation_result["desc_match_mode"] = recommendation_result["desc_match_mode"].map(_normalize_desc_match_mode)
    vi_target_rule["desc_match_mode"] = vi_target_rule["desc_match_mode"].map(_normalize_desc_match_mode)

    item_name_lookup = (
        vi_item_master[["vi_item_id", "vi_item_name"]]
        .drop_duplicates(subset=["vi_item_id"])
        .set_index("vi_item_id")["vi_item_name"]
        .to_dict()
    )
    recommendation_request["vi_item_name"] = recommendation_request["vi_item_name"].where(
        recommendation_request["vi_item_name"].ne(""),
        recommendation_request["vi_item_id"].map(item_name_lookup).fillna(""),
    )
    recommendation_result["vi_item_name"] = recommendation_result["vi_item_name"].where(
        recommendation_result["vi_item_name"].ne(""),
        recommendation_result["vi_item_id"].map(item_name_lookup).fillna(""),
    )

    raw_vi_target_rule = vi_target_rule.copy()
    final_vi_target_rule, final_vi_target_rule_audit, validation_messages = _build_final_vi_target_rule(
        vi_target_rule,
        recommendation_result,
    )

    volume["month"] = volume["month"].map(normalize_month_value)
    volume["subsidiary"] = volume["subsidiary"].fillna("").astype(str).str.strip()
    volume["model_suffix"] = volume["model_suffix"].fillna("").astype(str).str.strip()
    volume["production_qty"] = pd.to_numeric(volume["production_qty"], errors="coerce").fillna(0.0)

    return ScenarioInputs(
        vi_item_master=vi_item_master,
        recommendation_request=recommendation_request,
        recommendation_result=recommendation_result,
        vi_target_rule=final_vi_target_rule,
        volume=volume,
        raw_vi_target_rule=raw_vi_target_rule,
        final_vi_target_rule=final_vi_target_rule,
        final_vi_target_rule_audit=final_vi_target_rule_audit,
        validation_messages=validation_messages,
    )


def _prepare_rules(rule_frame: pd.DataFrame) -> pd.DataFrame:
    rules = rule_frame.copy()
    for column in VI_TARGET_RULE_COLUMNS:
        if column not in rules.columns:
            rules[column] = ""
        rules[column] = rules[column].fillna("").astype(str).str.strip()

    rules["subsidiary"] = rules["subsidiary"].map(normalize_filter_value)
    if "enabled" in rules.columns:
        rules.loc[rules["enabled"].eq(""), "enabled"] = "Y"
    rules["rule_type"] = rules["rule_type"].str.upper()
    if "desc_match_mode" in rules.columns:
        rules["desc_match_mode"] = rules["desc_match_mode"].map(_normalize_desc_match_mode)
    missing_rule_type = rules["rule_type"].eq("") & (
        rules["base_part_no"].ne("") | rules["new_part_no"].ne("")
    )
    rules.loc[missing_rule_type, "rule_type"] = "PART_NO"
    rules["effective_custom_rule_key"] = rules["custom_rule_key"].str.lower()
    rules.loc[rules["rule_type"] == "MODEL_TYPE", "effective_custom_rule_key"] = "model_type_match"
    rules.loc[rules["rule_type"] == "TOOL_TYPE", "effective_custom_rule_key"] = "indoor_tool_match"
    rules["effective_rule_type"] = rules["rule_type"].where(rules["rule_type"] == "PART_NO", "CUSTOM")

    expanded_rows: list[pd.Series] = []
    for _, source_row in rules.iterrows():
        subsidiaries = _split_filter_values(source_row.get("subsidiary", ""))
        if not subsidiaries:
            expanded_rows.append(source_row.copy())
            continue
        for subsidiary in subsidiaries:
            expanded_row = source_row.copy()
            expanded_row["subsidiary"] = subsidiary
            expanded_rows.append(expanded_row)
    if expanded_rows:
        rules = pd.DataFrame(expanded_rows, columns=rules.columns)
    else:
        rules = rules.iloc[0:0].copy()

    rules = rules[rules["enabled"].str.upper().ne("N")].copy()
    return rules


def _compact_snapshot_required_columns() -> list[str]:
    return [
        "raw_row_id",
        "subsidiary",
        "model_suffix",
        "part_no",
        "qty",
        "unit_price",
        "rmc_flag",
        "normalized_rmc_flag",
        "desc_text",
        "spec_text",
        "model_type",
        "indoor_tool",
        "level_num",
        "seq",
        "row_order",
        "parent_assy_part_no",
        "parent_assy_desc_text",
        "ancestor_desc_path",
        "ancestor_part_no_path",
        "price_source_part_no",
        "price_source_desc",
        "price_source_unit_price",
        "price_source_rmc_flag",
    ]


def _resolve_snapshot_source_path(
    history_root_text: str,
    snapshot_kind: str,
    bom_year: int,
    bom_month: int,
    snapshot_path: Path | None,
) -> Path | None:
    if snapshot_path is not None:
        resolved_path = Path(snapshot_path).expanduser().resolve()
        if resolved_path.exists():
            return resolved_path
        return resolved_path

    lake_root = default_lake_root(Path(history_root_text))
    snapshot_root = lake_root / (CURRENT_CACHE_DIR_NAME if snapshot_kind == "current_cache" else HISTORICAL_DIR_NAME)
    month_dir_candidates = [
        snapshot_root / f"bom_year={int(bom_year)}" / f"bom_month={int(bom_month)}",
        snapshot_root / f"bom_year={int(bom_year):04d}" / f"bom_month={int(bom_month):02d}",
    ]
    candidates: list[Path] = []
    for month_dir in month_dir_candidates:
        candidates.extend(sorted(month_dir.glob("*.parquet")))
    if not candidates:
        return None
    if len(candidates) == 1:
        return candidates[0].resolve()
    return max(candidates, key=lambda path: path.stat().st_mtime_ns).resolve()


def _resolve_snapshot_file_metadata(
    history_root_text: str,
    snapshot_kind: str,
    bom_year: int,
    bom_month: int,
    snapshot_path: Path | None,
    fallback_source: str,
) -> dict[str, object]:
    resolved_path = _resolve_snapshot_source_path(
        history_root_text,
        snapshot_kind,
        bom_year,
        bom_month,
        snapshot_path,
    )
    preferred_root = default_lake_root(Path(history_root_text)).parent
    if resolved_path is not None and resolved_path.exists():
        stat = resolved_path.stat()
        return {
            "snapshot_path": portable_path_token(resolved_path, preferred_root=preferred_root),
            "snapshot_mtime_ns": int(stat.st_mtime_ns),
            "snapshot_size": int(stat.st_size),
        }
    return {
        "snapshot_path": portable_path_token(resolved_path, preferred_root=preferred_root) if resolved_path is not None else str(fallback_source),
        "snapshot_mtime_ns": 0,
        "snapshot_size": 0,
    }


def _build_bom_pair_cache_metadata(
    *,
    history_root_text: str,
    base_snapshot_path: Path | None,
    base_snapshot_kind: str,
    base_year: int,
    base_month: int,
    current_snapshot_path: Path | None,
    current_snapshot_kind: str,
    current_year: int,
    current_month: int,
    compact_schema_version: str = COMPACT_CONTEXT_SCHEMA_VERSION,
) -> dict[str, object]:
    base_metadata = _resolve_snapshot_file_metadata(
        history_root_text,
        base_snapshot_kind,
        base_year,
        base_month,
        base_snapshot_path,
        str(base_snapshot_path) if base_snapshot_path is not None else "",
    )
    current_metadata = _resolve_snapshot_file_metadata(
        history_root_text,
        current_snapshot_kind,
        current_year,
        current_month,
        current_snapshot_path,
        str(current_snapshot_path) if current_snapshot_path is not None else "",
    )
    return {
        "compact_schema_version": compact_schema_version,
        "base_snapshot_path": base_metadata["snapshot_path"],
        "base_snapshot_mtime_ns": base_metadata["snapshot_mtime_ns"],
        "base_snapshot_size": base_metadata["snapshot_size"],
        "base_year": int(base_year),
        "base_month": int(base_month),
        "current_snapshot_path": current_metadata["snapshot_path"],
        "current_snapshot_mtime_ns": current_metadata["snapshot_mtime_ns"],
        "current_snapshot_size": current_metadata["snapshot_size"],
        "current_year": int(current_year),
        "current_month": int(current_month),
    }


def build_bom_pair_cache_key(
    *,
    history_root_text: str,
    base_snapshot_path: Path | None,
    base_snapshot_kind: str,
    base_year: int,
    base_month: int,
    current_snapshot_path: Path | None,
    current_snapshot_kind: str,
    current_year: int,
    current_month: int,
    compact_schema_version: str = COMPACT_CONTEXT_SCHEMA_VERSION,
) -> str:
    metadata = _build_bom_pair_cache_metadata(
        history_root_text=history_root_text,
        base_snapshot_path=base_snapshot_path,
        base_snapshot_kind=base_snapshot_kind,
        base_year=base_year,
        base_month=base_month,
        current_snapshot_path=current_snapshot_path,
        current_snapshot_kind=current_snapshot_kind,
        current_year=current_year,
        current_month=current_month,
        compact_schema_version=compact_schema_version,
    )
    payload = json.dumps(metadata, sort_keys=True, ensure_ascii=True).encode("utf-8")
    return hashlib.sha256(payload).hexdigest()[:16]


def _write_dataframe_to_parquet(frame: pd.DataFrame, path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    con = create_duckdb_connection()
    try:
        con.register("frame_to_save", frame)
        con.execute(
            f"COPY frame_to_save TO '{path.resolve().as_posix()}' (FORMAT PARQUET)"
        )
    finally:
        con.close()


def _read_dataframe_from_parquet(path: Path) -> pd.DataFrame:
    con = create_duckdb_connection()
    try:
        return con.execute("SELECT * FROM read_parquet(?)", [str(path.resolve())]).df()
    finally:
        con.close()


def load_or_build_bom_pair_compact_cache(
    *,
    history_root_text: str,
    base_year: int,
    base_month: int,
    base_snapshot_kind: str,
    base_snapshot_path: Path | None,
    current_year: int,
    current_month: int,
    current_snapshot_kind: str,
    current_snapshot_path: Path | None,
    compact_schema_version: str = COMPACT_CONTEXT_SCHEMA_VERSION,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    cache_started_at = time.perf_counter()
    metadata = _build_bom_pair_cache_metadata(
        history_root_text=history_root_text,
        base_snapshot_path=base_snapshot_path,
        base_snapshot_kind=base_snapshot_kind,
        base_year=base_year,
        base_month=base_month,
        current_snapshot_path=current_snapshot_path,
        current_snapshot_kind=current_snapshot_kind,
        current_year=current_year,
        current_month=current_month,
        compact_schema_version=compact_schema_version,
    )
    cache_key = build_bom_pair_cache_key(
        history_root_text=history_root_text,
        base_snapshot_path=base_snapshot_path,
        base_snapshot_kind=base_snapshot_kind,
        base_year=base_year,
        base_month=base_month,
        current_snapshot_path=current_snapshot_path,
        current_snapshot_kind=current_snapshot_kind,
        current_year=current_year,
        current_month=current_month,
        compact_schema_version=compact_schema_version,
    )
    cache_dir = bom_pair_cache_root() / f"bom_pair_{cache_key}"
    metadata_path = cache_dir / "metadata.json"
    base_cache_path = cache_dir / "base_compact.parquet"
    current_cache_path = cache_dir / "current_compact.parquet"

    cached_metadata: dict[str, object] = {}
    if metadata_path.exists():
        try:
            cached_metadata = json.loads(metadata_path.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            cached_metadata = {}

    if (
        base_cache_path.exists()
        and current_cache_path.exists()
        and cached_metadata == metadata
    ):
        print(f"[Cache] bom_pair hit: {cache_key}", flush=True)
        base_frame = _read_dataframe_from_parquet(base_cache_path)
        current_frame = _read_dataframe_from_parquet(current_cache_path)
        required_columns = set(_compact_snapshot_required_columns())
        if required_columns.issubset(base_frame.columns) and required_columns.issubset(current_frame.columns):
            _log_timing("load BOM pair compact cache", cache_started_at)
            return (
                base_frame[_compact_snapshot_required_columns()].copy(),
                current_frame[_compact_snapshot_required_columns()].copy(),
            )
        print(f"[Cache] bom_pair miss: {cache_key} (invalid cached schema)", flush=True)
    else:
        print(f"[Cache] bom_pair miss: {cache_key}", flush=True)
    build_started_at = time.perf_counter()
    base_source_path = _resolve_snapshot_source_path(
        history_root_text,
        base_snapshot_kind,
        base_year,
        base_month,
        base_snapshot_path,
    )
    current_source_path = _resolve_snapshot_source_path(
        history_root_text,
        current_snapshot_kind,
        current_year,
        current_month,
        current_snapshot_path,
    )
    print(
        f"[Cache] bom_pair build start | base={base_year:04d}-{base_month:02d} | source={base_source_path or '-'}",
        flush=True,
    )
    base_build_started_at = time.perf_counter()
    base_frame = load_bom_snapshot_compact(
        history_root_text,
        base_year,
        base_month,
        snapshot_kind=base_snapshot_kind,
        snapshot_path_text=str(base_snapshot_path) if base_snapshot_path is not None else "",
    )
    _log_timing("build base compact snapshot", base_build_started_at)
    print(
        f"[Cache] bom_pair build continue | current={current_year:04d}-{current_month:02d} | source={current_source_path or '-'}",
        flush=True,
    )
    current_build_started_at = time.perf_counter()
    current_frame = load_bom_snapshot_compact(
        history_root_text,
        current_year,
        current_month,
        snapshot_kind=current_snapshot_kind,
        snapshot_path_text=str(current_snapshot_path) if current_snapshot_path is not None else "",
    )
    _log_timing("build current compact snapshot", current_build_started_at)
    cache_dir.mkdir(parents=True, exist_ok=True)
    print("[Cache] write compact parquet start", flush=True)
    _write_dataframe_to_parquet(base_frame[_compact_snapshot_required_columns()].copy(), base_cache_path)
    _write_dataframe_to_parquet(current_frame[_compact_snapshot_required_columns()].copy(), current_cache_path)
    metadata_path.write_text(json.dumps(metadata, indent=2, ensure_ascii=True), encoding="utf-8")
    print("[Cache] write compact parquet done", flush=True)
    _log_timing("build BOM pair compact cache", build_started_at)
    _log_timing("load BOM pair compact cache", cache_started_at)
    return base_frame, current_frame


@lru_cache(maxsize=64)
def load_bom_snapshot_compact(
    history_root_text: str,
    bom_year: int,
    bom_month: int,
    snapshot_kind: str = "historical",
    snapshot_path_text: str = "",
) -> pd.DataFrame:
    require_duckdb()
    parquet_source = snapshot_path_text or _snapshot_glob(Path(history_root_text), snapshot_kind)
    resolved_source_path = _resolve_snapshot_source_path(
        history_root_text,
        snapshot_kind,
        bom_year,
        bom_month,
        Path(snapshot_path_text) if snapshot_path_text else None,
    )
    columns = _snapshot_columns(history_root_text, snapshot_kind, snapshot_path_text)
    desc_column = _first_available_column(columns, ["description", "desc", "item_desc", "part_desc"])
    spec_column = _first_available_column(columns, ["spec_text", "spec", "specification", "material_spec"])
    level_column = _first_available_column(columns, ["level", "lv", "bom_level", "lvl"])
    level_num_column = _first_available_column(columns, ["level_num"])
    row_order_column = _first_available_column(columns, ["row_order"])
    rmc_column = _first_available_column(columns, ["rmc_flag", "RMC Flag", "RMC_FLAG", "RMC", "rmc"])
    desc_expr = f"TRIM(CAST({desc_column} AS VARCHAR))" if desc_column else "''"
    spec_expr = f"TRIM(CAST({spec_column} AS VARCHAR))" if spec_column else "''"
    level_expr = f"CAST({level_column} AS VARCHAR)" if level_column else "''"
    rmc_expr = f"TRIM(CAST({_sql_identifier(rmc_column)} AS VARCHAR))" if rmc_column else "''"
    level_num_expr = (
        f"COALESCE(TRY_CAST({level_num_column} AS INTEGER), {_build_level_num_expr(level_expr)})"
        if level_num_column
        else _build_level_num_expr(level_expr)
    )
    row_order_expr = f"TRY_CAST({row_order_column} AS BIGINT)" if row_order_column else "ROW_NUMBER() OVER ()"
    started_at = time.perf_counter()
    print(
        f"[Compact] read parquet start | {bom_year:04d}-{bom_month:02d} | source={resolved_source_path or parquet_source}",
        flush=True,
    )
    query = f"""
        SELECT
            ROW_NUMBER() OVER () AS raw_row_id,
            COALESCE(NULLIF(TRIM(CAST(subsidiary AS VARCHAR)), ''), 'Unknown') AS subsidiary,
            TRIM(CAST(model_suffix AS VARCHAR)) AS model_suffix,
            TRIM(CAST(part_no AS VARCHAR)) AS part_no,
            COALESCE(qty, 0) AS qty,
            COALESCE(unit_price, 0) AS unit_price,
            COALESCE(NULLIF({rmc_expr}, ''), '') AS rmc_flag,
            CASE
                WHEN UPPER(COALESCE(NULLIF({rmc_expr}, ''), '')) IN ('Y', 'YES', 'TRUE', '1') THEN 'Y'
                ELSE 'N'
            END AS normalized_rmc_flag,
            COALESCE(NULLIF({desc_expr}, ''), '') AS desc_text,
            COALESCE(NULLIF({spec_expr}, ''), '') AS spec_text,
            {MODEL_TYPE_SQL} AS model_type,
            {INDOOR_TOOL_SQL} AS indoor_tool,
            {level_num_expr} AS level_num,
            COALESCE(TRY_CAST(seq AS DOUBLE), 0) AS seq,
            COALESCE({row_order_expr}, 0) AS row_order
        FROM read_parquet(?)
        WHERE bom_year = ?
          AND bom_month = ?
          AND model_suffix IS NOT NULL
          AND part_no IS NOT NULL
        ORDER BY subsidiary, model_suffix, seq, row_order, part_no
    """
    con = create_duckdb_connection()
    try:
        query_started_at = time.perf_counter()
        frame = con.execute(
            query,
            [parquet_source, int(bom_year), int(bom_month)],
        ).df()
    finally:
        con.close()
    _log_timing(f"query compact snapshot {bom_year}-{bom_month:02d} ({snapshot_kind})", query_started_at)
    print(
        f"[Compact] read parquet done | rows={len(frame.index):,}",
        flush=True,
    )
    print(
        f"[Compact] annotate context start | rows={len(frame.index):,}",
        flush=True,
    )
    annotate_started_at = time.perf_counter()
    frame = _annotate_bom_context_columns(frame)
    _log_timing(f"annotate compact context {bom_year}-{bom_month:02d} ({snapshot_kind})", annotate_started_at)
    print("[Compact] annotate context done", flush=True)
    for column in _compact_snapshot_required_columns():
        if column not in frame.columns:
            frame[column] = ""
    frame = frame[_compact_snapshot_required_columns()].copy()
    _log_timing(f"load compact snapshot {bom_year}-{bom_month:02d} ({snapshot_kind})", started_at)
    return frame


def _build_volume_lookup(volume: pd.DataFrame, month_value: str) -> pd.DataFrame:
    if volume.empty:
        return _empty_frame(["month", "subsidiary", "model_suffix", "production_qty"])
    lookup = volume.copy()
    lookup["month"] = lookup["month"].map(normalize_month_value)
    lookup["subsidiary"] = lookup["subsidiary"].fillna("").astype(str).str.strip()
    lookup["model_suffix"] = lookup["model_suffix"].fillna("").astype(str).str.strip()
    lookup["production_qty"] = pd.to_numeric(lookup["production_qty"], errors="coerce").fillna(0.0)
    lookup = lookup[lookup["month"] == month_value].copy()
    return (
        lookup.groupby(["month", "subsidiary", "model_suffix"], as_index=False)["production_qty"]
        .sum()
        .reset_index(drop=True)
    )


def _empty_outputs(detail: bool) -> dict[str, pd.DataFrame]:
    return {
        "FinalVITargetRule": _empty_rule_audit_frame(),
        "ViItemSummary": _empty_frame(
            [
                "vi_item_id",
                "vi_item_name",
                "target_model_count",
                "applied_model_count",
                "not_applied_model_count",
                "mixed_review_model_count",
                "missing_or_model_changed_count",
                "application_rate",
                "realized_vi_amount",
                "remaining_vi_amount",
                "expected_vi_amount_usd",
                "total_opportunity",
            ]
        ),
        "ViSubsidiarySummary": _empty_frame(
            [
                "subsidiary",
                "target_model_count",
                "applied_model_count",
                "not_applied_model_count",
                "mixed_review_model_count",
                "missing_or_model_changed_count",
                "application_rate",
                "realized_vi_amount",
                "remaining_vi_amount",
                "expected_vi_amount_usd",
                "total_opportunity",
            ]
        ),
        "BaseTargetModelList": _empty_frame(
            [
                "vi_item_id",
                "vi_item_name",
                "rule_type",
                "subsidiary",
                "model_suffix",
                "model_type",
                "indoor_tool",
                "parent_assy_part_no",
                "parent_assy_desc_text",
                "ancestor_desc_path",
                "base_part_no",
                "new_part_no",
                "base_unit_price",
                "base_bom_qty",
                "target_source",
            ]
        ),
        "BasePNTargetSummary": _empty_frame(
            [
                "vi_item_id",
                "vi_item_name",
                "subsidiary",
                "base_part_no",
                "base_part_desc",
                "base_pn_model_count",
                "item_target_model_count",
                "base_pn_model_share_pct",
                "applied_model_count",
                "not_applied_model_count",
                "mixed_review_model_count",
                "missing_or_model_changed_count",
                "realized_vi_amount",
                "remaining_vi_amount",
            ]
        ),
        "ModelChangeSummary": _empty_frame(
            [
                "vi_item_id",
                "vi_item_name",
                "subsidiary",
                "base_target_model_count",
                "existing_model_count",
                "removed_model_count",
                "added_model_count",
                "base_target_applied_count",
                "base_target_not_applied_count",
                "base_target_mixed_review_count",
                "base_target_missing_or_model_changed_count",
                "added_applied_count",
                "added_not_applied_count",
                "added_mixed_review_count",
                "added_not_target_or_unknown_count",
            ]
        ),
        "ModelChangeDetail": _empty_frame(
            [
                "vi_item_id",
                "vi_item_name",
                "rule_type",
                "subsidiary",
                "model_suffix",
                "model_type",
                "indoor_tool",
                "parent_assy_part_no",
                "parent_assy_desc_text",
                "ancestor_desc_path",
                "model_change_status",
                "base_exists",
                "current_exists",
                "base_target_flag",
                "base_part_no",
                "new_part_no",
                "current_base_part_exists",
                "current_new_part_exists",
                "application_status",
                "base_target_parent_assy_part_no",
                "base_target_parent_assy_desc_text",
                "base_target_ancestor_desc_path",
                "current_new_parent_assy_part_no",
                "current_new_parent_assy_desc_text",
                "current_new_ancestor_desc_path",
                "matched_context_method",
                "new_part_found_under_desc_context",
                "new_part_found_anywhere_in_model",
                "base_unit_price",
                "new_unit_price",
                "base_rmc_flag",
                "new_rmc_flag",
                "base_price_source_part_no",
                "base_price_source_desc",
                "base_price_source_unit_price",
                "base_price_source_rmc_flag",
                "new_price_source_part_no",
                "new_price_source_desc",
                "new_price_source_unit_price",
                "new_price_source_rmc_flag",
                "bom_qty",
                "production_qty",
                "realized_vi_amount",
                "remaining_vi_amount",
                "expected_vi_amount_usd",
                "expected_vi_calc_status",
                "expected_vi_calc_message",
                "issue_flag",
                "issue_message",
            ]
        )
        if detail
        else _empty_frame([]),
        "ViModelDetail": _empty_frame(
            [
                "vi_item_id",
                "vi_item_name",
                "rule_type",
                "subsidiary",
                "model_suffix",
                "model_type",
                "indoor_tool",
                "parent_assy_part_no",
                "parent_assy_desc_text",
                "ancestor_desc_path",
                "application_status",
                "base_part_no",
                "new_part_no",
                "base_target_parent_assy_part_no",
                "base_target_parent_assy_desc_text",
                "base_target_ancestor_desc_path",
                "current_new_parent_assy_part_no",
                "current_new_parent_assy_desc_text",
                "current_new_ancestor_desc_path",
                "matched_context_method",
                "new_part_found_under_desc_context",
                "new_part_found_anywhere_in_model",
                "base_unit_price",
                "new_unit_price",
                "base_rmc_flag",
                "new_rmc_flag",
                "base_price_source_part_no",
                "base_price_source_desc",
                "base_price_source_unit_price",
                "base_price_source_rmc_flag",
                "new_price_source_part_no",
                "new_price_source_desc",
                "new_price_source_unit_price",
                "new_price_source_rmc_flag",
                "unit_saving",
                "bom_qty",
                "production_qty",
                "realized_vi_amount",
                "remaining_vi_amount",
                "expected_vi_amount_usd",
                "expected_vi_calc_status",
                "expected_vi_calc_message",
                "issue_flag",
                "issue_message",
            ]
        )
        if detail
        else _empty_frame([]),
        "ViExceptionLog": _empty_frame(
            [
                "vi_item_id",
                "vi_item_name",
                "rule_type",
                "subsidiary",
                "model_suffix",
                "issue_type",
                "issue_message",
            ]
        )
        if detail
        else _empty_frame([]),
    }


def _append_issue_frame(base_frame: pd.DataFrame, mask: pd.Series, issue_type: str, issue_message: str) -> pd.DataFrame:
    if base_frame.empty or not bool(mask.any()):
        return _empty_frame(["vi_item_id", "vi_item_name", "rule_type", "subsidiary", "model_suffix", "issue_type", "issue_message"])
    frame = base_frame.loc[mask, ["vi_item_id", "vi_item_name", "rule_type", "subsidiary", "model_suffix"]].copy()
    frame["issue_type"] = issue_type
    frame["issue_message"] = issue_message
    return frame


def _rule_has_row_context(rule: pd.Series) -> bool:
    return any(
        bool(_normalize_filter_value(rule.get(column, "")))
        for column in ["desc_contains_any", "spec_contains_any", "parent_assy_desc_contains_any"]
    )


def _first_non_empty_value(series: pd.Series) -> str:
    for value in series.fillna("").astype(str):
        text = value.strip()
        if text:
            return text
    return ""


def _first_non_na_value(series: pd.Series) -> object:
    for value in series:
        if pd.notna(value):
            return value
    return pd.NA


def _join_unique_text(values: list[object]) -> str:
    return " | ".join(sorted({str(value).strip() for value in values if str(value).strip()}))


def _normalize_context_text(value: object) -> str:
    text = _normalize_text(value)
    if not text:
        return ""
    return " > ".join(part.strip().lower() for part in text.split(">") if part.strip())


def _business_context_value(row_dict: dict[str, object]) -> tuple[str, str]:
    ancestor_desc_path = _normalize_text(row_dict.get("ancestor_desc_path", ""))
    if ancestor_desc_path:
        return ancestor_desc_path, "ancestor_desc_path"
    return _normalize_text(row_dict.get("parent_assy_desc_text", "")), "parent_assy_desc_text"


def _business_context_key(row_dict: dict[str, object]) -> tuple[object, ...]:
    context_value, _ = _business_context_value(row_dict)
    return (
        row_dict.get("subsidiary", ""),
        row_dict.get("model_suffix", ""),
        row_dict.get("model_type", ""),
        row_dict.get("indoor_tool", ""),
        _normalize_context_text(context_value),
    )


def _aggregate_contexts_by_business_key(contexts: pd.DataFrame) -> dict[tuple[object, ...], dict[str, object]]:
    if contexts.empty:
        return {}

    aggregated: dict[tuple[object, ...], dict[str, object]] = {}
    for _, context_row in contexts.iterrows():
        context_dict = context_row.to_dict()
        key = _business_context_key(context_dict)
        context_value, matched_context_method = _business_context_value(context_dict)
        entry = aggregated.setdefault(
            key,
            {
                "subsidiary": context_dict.get("subsidiary", ""),
                "model_suffix": context_dict.get("model_suffix", ""),
                "model_type": context_dict.get("model_type", ""),
                "indoor_tool": context_dict.get("indoor_tool", ""),
                "parent_assy_part_no": [],
                "parent_assy_desc_text": [],
                "ancestor_desc_path": [],
                "ancestor_part_no_path": [],
                "business_context_value": context_value,
                "matched_context_method": matched_context_method,
                "base_part_desc": "",
                "base_unit_price": pd.NA,
                "base_rmc_flag": "N",
                "base_price_source_part_no": "",
                "base_price_source_desc": "",
                "base_price_source_unit_price": pd.NA,
                "base_price_source_rmc_flag": "N",
                "base_bom_qty": 0.0,
                "base_new_qty": 0.0,
                "current_base_part_exists": 0,
                "current_new_part_exists": 0,
                "new_unit_price": pd.NA,
                "new_rmc_flag": "N",
                "new_price_source_part_no": "",
                "new_price_source_desc": "",
                "new_price_source_unit_price": pd.NA,
                "new_price_source_rmc_flag": "N",
                "current_base_qty": 0.0,
                "current_new_qty": 0.0,
            },
        )
        entry["parent_assy_part_no"].append(context_dict.get("parent_assy_part_no", ""))
        entry["parent_assy_desc_text"].append(context_dict.get("parent_assy_desc_text", ""))
        entry["ancestor_desc_path"].append(context_dict.get("ancestor_desc_path", ""))
        entry["ancestor_part_no_path"].append(context_dict.get("ancestor_part_no_path", ""))
        if not entry["base_part_desc"]:
            entry["base_part_desc"] = _normalize_text(context_dict.get("base_part_desc", ""))
        if pd.isna(entry["base_unit_price"]) and not pd.isna(context_dict.get("base_unit_price", pd.NA)):
            entry["base_unit_price"] = context_dict.get("base_unit_price", pd.NA)
        if _normalize_rmc_flag_value(entry["base_rmc_flag"]) != "Y":
            entry["base_rmc_flag"] = _normalize_rmc_flag_value(context_dict.get("base_rmc_flag", entry["base_rmc_flag"]))
        if not entry["base_price_source_part_no"]:
            entry["base_price_source_part_no"] = _normalize_text(context_dict.get("base_price_source_part_no", ""))
        if not entry["base_price_source_desc"]:
            entry["base_price_source_desc"] = _normalize_text(context_dict.get("base_price_source_desc", ""))
        if pd.isna(entry["base_price_source_unit_price"]) and not pd.isna(context_dict.get("base_price_source_unit_price", pd.NA)):
            entry["base_price_source_unit_price"] = context_dict.get("base_price_source_unit_price", pd.NA)
        if _normalize_rmc_flag_value(entry["base_price_source_rmc_flag"]) != "Y":
            entry["base_price_source_rmc_flag"] = _normalize_rmc_flag_value(context_dict.get("base_price_source_rmc_flag", entry["base_price_source_rmc_flag"]))
        if pd.isna(entry["new_unit_price"]) and not pd.isna(context_dict.get("new_unit_price", pd.NA)):
            entry["new_unit_price"] = context_dict.get("new_unit_price", pd.NA)
        if _normalize_rmc_flag_value(entry["new_rmc_flag"]) != "Y":
            entry["new_rmc_flag"] = _normalize_rmc_flag_value(context_dict.get("new_rmc_flag", entry["new_rmc_flag"]))
        if not entry["new_price_source_part_no"]:
            entry["new_price_source_part_no"] = _normalize_text(context_dict.get("new_price_source_part_no", ""))
        if not entry["new_price_source_desc"]:
            entry["new_price_source_desc"] = _normalize_text(context_dict.get("new_price_source_desc", ""))
        if pd.isna(entry["new_price_source_unit_price"]) and not pd.isna(context_dict.get("new_price_source_unit_price", pd.NA)):
            entry["new_price_source_unit_price"] = context_dict.get("new_price_source_unit_price", pd.NA)
        if _normalize_rmc_flag_value(entry["new_price_source_rmc_flag"]) != "Y":
            entry["new_price_source_rmc_flag"] = _normalize_rmc_flag_value(context_dict.get("new_price_source_rmc_flag", entry["new_price_source_rmc_flag"]))
        entry["base_bom_qty"] += float(context_dict.get("base_bom_qty", 0.0) or 0.0)
        entry["base_new_qty"] += float(context_dict.get("base_new_qty", 0.0) or 0.0)
        entry["current_base_qty"] += float(context_dict.get("current_base_qty", 0.0) or 0.0)
        entry["current_new_qty"] += float(context_dict.get("current_new_qty", 0.0) or 0.0)
        entry["current_base_part_exists"] = max(
            int(entry["current_base_part_exists"]),
            int(context_dict.get("current_base_part_exists", 0) or 0),
        )
        entry["current_new_part_exists"] = max(
            int(entry["current_new_part_exists"]),
            int(context_dict.get("current_new_part_exists", 0) or 0),
        )

    for entry in aggregated.values():
        entry["parent_assy_part_no"] = _join_unique_text(entry["parent_assy_part_no"])
        entry["parent_assy_desc_text"] = _join_unique_text(entry["parent_assy_desc_text"])
        entry["ancestor_desc_path"] = _join_unique_text(entry["ancestor_desc_path"])
        entry["ancestor_part_no_path"] = _join_unique_text(entry["ancestor_part_no_path"])
    return aggregated


def _build_model_presence_lookup(matches: pd.DataFrame) -> dict[tuple[str, str], dict[str, object]]:
    if matches.empty:
        return {}

    working = matches.copy()
    working["subsidiary"] = working["subsidiary"].fillna("").astype(str).str.strip()
    working["model_suffix"] = working["model_suffix"].fillna("").astype(str).str.strip()
    if "parent_assy_part_no" not in working.columns:
        working["parent_assy_part_no"] = ""
    if "parent_assy_desc_text" not in working.columns:
        working["parent_assy_desc_text"] = ""
    if "ancestor_desc_path" not in working.columns:
        working["ancestor_desc_path"] = ""
    if "ancestor_part_no_path" not in working.columns:
        working["ancestor_part_no_path"] = ""

    lookup: dict[tuple[str, str], dict[str, object]] = {}
    for (subsidiary, model_suffix), group in working.groupby(["subsidiary", "model_suffix"], dropna=False, sort=True):
        group = group.copy()
        group["qty"] = pd.to_numeric(group["qty"], errors="coerce").fillna(0.0)
        group["unit_price"] = pd.to_numeric(group["unit_price"], errors="coerce")
        lookup[(str(subsidiary), str(model_suffix))] = {
            "exists": int(bool(group["qty"].gt(0).any())),
            "parent_assy_part_no": _join_unique_text(group["parent_assy_part_no"].tolist()),
            "parent_assy_desc_text": _join_unique_text(group["parent_assy_desc_text"].tolist()),
            "ancestor_desc_path": _join_unique_text(group["ancestor_desc_path"].tolist()),
            "ancestor_part_no_path": _join_unique_text(group["ancestor_part_no_path"].tolist()),
            "new_unit_price": group["unit_price"].max() if group["unit_price"].notna().any() else pd.NA,
            "new_rmc_flag": _first_non_empty_value(group["normalized_rmc_flag"]),
            "new_price_source_part_no": _first_non_empty_value(group["price_source_part_no"]),
            "new_price_source_desc": _first_non_empty_value(group["price_source_desc"]),
            "new_price_source_unit_price": _first_non_na_value(group["price_source_unit_price"]),
            "new_price_source_rmc_flag": _first_non_empty_value(group["price_source_rmc_flag"]),
            "qty": float(group["qty"].sum()),
        }
    return lookup


def _annotate_bom_context_columns(frame: pd.DataFrame) -> pd.DataFrame:
    annotated = frame.copy()
    for column in [
        "parent_assy_part_no",
        "parent_assy_desc_text",
        "ancestor_desc_path",
        "ancestor_part_no_path",
        "price_source_part_no",
        "price_source_desc",
        "price_source_unit_price",
        "price_source_rmc_flag",
    ]:
        annotated[column] = ""
    if annotated.empty:
        return annotated

    parent_assy_part_no_values: list[str] = []
    parent_assy_desc_text_values: list[str] = []
    ancestor_desc_path_values: list[str] = []
    ancestor_part_no_path_values: list[str] = []
    price_source_part_no_values: list[str] = []
    price_source_desc_values: list[str] = []
    price_source_unit_price_values: list[object] = []
    price_source_rmc_flag_values: list[str] = []

    current_group: tuple[str, str] | None = None
    stack: list[dict[str, object]] = []
    total_rows = len(annotated.index)
    row_iterator = annotated[
        ["subsidiary", "model_suffix", "level_num", "part_no", "desc_text", "normalized_rmc_flag", "unit_price"]
    ].itertuples(index=False, name=None)
    for row_index, row in enumerate(row_iterator, start=1):
        subsidiary = _normalize_text(row[0])
        model_suffix = _normalize_text(row[1])
        level_value = row[2]
        part_no = _normalize_text(row[3])
        desc_text = _normalize_text(row[4])
        normalized_rmc_flag = _normalize_rmc_flag_value(row[5])
        try:
            unit_price = float(row[6]) if row[6] is not None and not pd.isna(row[6]) else pd.NA
        except (TypeError, ValueError):
            unit_price = pd.NA
        group_key = (subsidiary, model_suffix)
        if group_key != current_group:
            current_group = group_key
            stack = []

        level_num: int | None
        if pd.isna(level_value):
            level_num = None
        else:
            try:
                level_num = int(level_value)
            except (TypeError, ValueError):
                level_num = None

        if level_num is None:
            stack = []
        else:
            while stack and int(stack[-1]["level_num"]) >= level_num:
                stack.pop()

        parent_node = stack[-1] if stack else None
        parent_assy_part_no = str(parent_node["part_no"]).strip() if parent_node else ""
        parent_assy_desc_text = str(parent_node["desc_text"]).strip() if parent_node else ""
        ancestor_desc_path = str(parent_node["desc_path"]).strip() if parent_node else ""
        ancestor_part_no_path = str(parent_node["part_path"]).strip() if parent_node else ""

        parent_assy_part_no_values.append(parent_assy_part_no)
        parent_assy_desc_text_values.append(parent_assy_desc_text)
        ancestor_desc_path_values.append(ancestor_desc_path)
        ancestor_part_no_path_values.append(ancestor_part_no_path)

        price_source_node: dict[str, object] | None = None
        if normalized_rmc_flag == "Y":
            price_source_node = {
                "part_no": part_no,
                "desc_text": desc_text,
                "unit_price": unit_price,
                "normalized_rmc_flag": "Y",
            }
        else:
            for candidate_node in reversed(stack):
                if _normalize_rmc_flag_value(candidate_node.get("normalized_rmc_flag", "")) == "Y":
                    price_source_node = candidate_node
                    break

        price_source_part_no_values.append(_normalize_text(price_source_node.get("part_no", "")) if price_source_node else "")
        price_source_desc_values.append(_normalize_text(price_source_node.get("desc_text", "")) if price_source_node else "")
        price_source_unit_price_values.append(price_source_node.get("unit_price", pd.NA) if price_source_node else pd.NA)
        price_source_rmc_flag_values.append(
            _normalize_rmc_flag_value(price_source_node.get("normalized_rmc_flag", "")) if price_source_node else "N"
        )

        if level_num is not None:
            desc_path = ancestor_desc_path
            part_path = ancestor_part_no_path
            if desc_text:
                desc_path = f"{ancestor_desc_path} > {desc_text}" if ancestor_desc_path else desc_text
            if part_no:
                part_path = f"{ancestor_part_no_path} > {part_no}" if ancestor_part_no_path else part_no
            stack.append(
                {
                    "level_num": level_num,
                    "part_no": part_no,
                    "desc_text": desc_text,
                    "desc_path": desc_path,
                    "part_path": part_path,
                    "unit_price": unit_price,
                    "normalized_rmc_flag": normalized_rmc_flag,
                }
            )
        if row_index % 50000 == 0:
            print(
                f"[Compact] annotate context progress: {row_index:,} / {total_rows:,}",
                flush=True,
            )

    annotated["parent_assy_part_no"] = parent_assy_part_no_values
    annotated["parent_assy_desc_text"] = parent_assy_desc_text_values
    annotated["ancestor_desc_path"] = ancestor_desc_path_values
    annotated["ancestor_part_no_path"] = ancestor_part_no_path_values
    annotated["price_source_part_no"] = price_source_part_no_values
    annotated["price_source_desc"] = price_source_desc_values
    annotated["price_source_unit_price"] = price_source_unit_price_values
    annotated["price_source_rmc_flag"] = price_source_rmc_flag_values
    return annotated


def _build_rule_candidate_query(
    source_table: str,
    rule: pd.Series,
    part_candidates: list[str] | None,
    require_positive_qty: bool,
    include_parent_columns: bool,
    apply_desc_spec_filters: bool = True,
) -> tuple[str, list[object]]:
    params: list[object] = []
    conditions: list[str] = []
    subsidiary = normalize_filter_value(rule.get("subsidiary", ""))
    model_types = _split_filter_values_upper(rule.get("model_type", ""))
    indoor_tools = _split_filter_values_upper(rule.get("indoor_tool", ""))
    desc_keywords = _normalize_filter_value(rule.get("desc_contains_any", ""))
    spec_keywords = _normalize_filter_value(rule.get("spec_contains_any", ""))
    parent_keywords = _normalize_filter_value(rule.get("parent_assy_desc_contains_any", ""))
    match_mode = _normalize_desc_match_mode(rule.get("desc_match_mode", ""))

    if subsidiary:
        conditions.append("subsidiary = ?")
        params.append(subsidiary)
    if part_candidates is not None:
        usable_candidates = [str(value).strip().upper() for value in part_candidates if str(value).strip()]
        if usable_candidates:
            placeholders = ", ".join("?" for _ in usable_candidates)
            conditions.append(f"UPPER(part_no) IN ({placeholders})")
            params.extend(usable_candidates)
        else:
            conditions.append("FALSE")
    if model_types:
        _append_in_condition(conditions, params, "model_type", model_types)
    if indoor_tools:
        _append_in_condition(conditions, params, "indoor_tool", indoor_tools)
    if apply_desc_spec_filters and desc_keywords:
        conditions.append(_build_keyword_like_clause("UPPER(COALESCE(desc_text, ''))", desc_keywords, match_mode, params))
    if apply_desc_spec_filters and spec_keywords:
        conditions.append(_build_keyword_like_clause("UPPER(COALESCE(spec_text, ''))", spec_keywords, match_mode, params))
    if require_positive_qty:
        conditions.append("COALESCE(qty, 0) > 0")

    where_clause = " AND ".join(conditions) if conditions else "TRUE"
    parent_context_expr = "UPPER(COALESCE(NULLIF(ancestor_desc_path, ''), parent_assy_desc_text, ''))"

    if include_parent_columns:
        parent_params = list(params)
        parent_clause = "TRUE"
        if parent_keywords:
            parent_clause = _build_keyword_like_clause(parent_context_expr, parent_keywords, match_mode, parent_params)
        query = f"""
            SELECT *
            FROM {source_table}
            WHERE {where_clause}
              AND {parent_clause}
            ORDER BY subsidiary, model_suffix, seq, row_order, part_no
        """
        return query, parent_params

    query = f"""
        SELECT
            raw_row_id,
            subsidiary,
            model_suffix,
            part_no,
            qty,
            unit_price,
            rmc_flag,
            normalized_rmc_flag,
            desc_text,
            spec_text,
            model_type,
            indoor_tool,
            level_num,
            seq,
            row_order,
            '' AS parent_assy_part_no,
            '' AS parent_assy_desc_text,
            '' AS ancestor_desc_path,
            '' AS ancestor_part_no_path,
            '' AS price_source_part_no,
            '' AS price_source_desc,
            NULL AS price_source_unit_price,
            '' AS price_source_rmc_flag
        FROM {source_table}
        WHERE {where_clause}
        ORDER BY subsidiary, model_suffix, seq, row_order, part_no
    """
    return query, params


def _query_rule_candidate_rows(
    con: DuckDBConnection,
    source_table: str,
    rule: pd.Series,
    part_candidates: list[str] | None,
    require_positive_qty: bool,
    include_parent_columns: bool,
    apply_desc_spec_filters: bool = True,
) -> pd.DataFrame:
    query, params = _build_rule_candidate_query(
        source_table,
        rule,
        part_candidates=part_candidates,
        require_positive_qty=require_positive_qty,
        include_parent_columns=include_parent_columns,
        apply_desc_spec_filters=apply_desc_spec_filters,
    )
    return con.execute(query, params).df()


def _materialize_part_candidate_pool(
    con: DuckDBConnection,
    *,
    pool_table: str,
    source_table: str,
    part_numbers: list[str],
    require_positive_qty: bool,
) -> None:
    normalized_parts = sorted({str(value).strip().upper() for value in part_numbers if str(value).strip()})
    if not normalized_parts:
        con.execute(f"CREATE OR REPLACE TEMP TABLE {pool_table} AS SELECT * FROM {source_table} WHERE 1=0")
        return

    parts_table = f"{pool_table}_parts"
    con.register(parts_table, pd.DataFrame({"part_no_upper": normalized_parts}))
    qty_clause = "AND COALESCE(src.qty, 0) > 0" if require_positive_qty else ""
    try:
        con.execute(
            f"""
            CREATE OR REPLACE TEMP TABLE {pool_table} AS
            SELECT src.*
            FROM {source_table} src
            INNER JOIN {parts_table} parts
              ON UPPER(src.part_no) = parts.part_no_upper
            WHERE 1=1
              {qty_clause}
            """
        )
    finally:
        try:
            con.unregister(parts_table)
        except Exception:
            pass


def _summarize_context_matches(
    matches: pd.DataFrame,
    *,
    base_part_no: str,
    new_part_no: str,
    include_parent_columns: bool,
) -> pd.DataFrame:
    if matches.empty:
        return _empty_frame(
            [
                "subsidiary",
                "model_suffix",
                "model_type",
                "indoor_tool",
                "parent_assy_part_no",
                "parent_assy_desc_text",
                "ancestor_desc_path",
                "ancestor_part_no_path",
                "base_part_desc",
                "base_unit_price",
                "base_rmc_flag",
                "base_price_source_part_no",
                "base_price_source_desc",
                "base_price_source_unit_price",
                "base_price_source_rmc_flag",
                "base_bom_qty",
                "base_new_qty",
                "current_base_part_exists",
                "current_new_part_exists",
                "new_unit_price",
                "new_rmc_flag",
                "new_price_source_part_no",
                "new_price_source_desc",
                "new_price_source_unit_price",
                "new_price_source_rmc_flag",
                "current_base_qty",
                "current_new_qty",
            ]
        )

    working = matches.copy()
    if not include_parent_columns:
        working["parent_assy_part_no"] = ""
        working["parent_assy_desc_text"] = ""
        working["ancestor_desc_path"] = ""
        working["ancestor_part_no_path"] = ""

    group_columns = [
        "subsidiary",
        "model_suffix",
        "model_type",
        "indoor_tool",
        "parent_assy_part_no",
        "parent_assy_desc_text",
        "ancestor_desc_path",
        "ancestor_part_no_path",
    ]
    rows: list[dict[str, object]] = []
    for group_key, frame in working.groupby(group_columns, dropna=False, sort=True):
        if not isinstance(group_key, tuple):
            group_key = (group_key,)
        group_row = dict(zip(group_columns, group_key))
        frame = frame.copy()
        frame["qty"] = pd.to_numeric(frame["qty"], errors="coerce").fillna(0.0)
        frame["unit_price"] = pd.to_numeric(frame["unit_price"], errors="coerce")

        base_mask = frame["part_no"].fillna("").astype(str).eq(base_part_no)
        new_mask = frame["part_no"].fillna("").astype(str).eq(new_part_no)
        rows.append(
            {
                **group_row,
                "base_part_desc": _first_non_empty_value(frame.loc[base_mask, "desc_text"]),
                "base_unit_price": frame.loc[base_mask, "unit_price"].max() if bool(base_mask.any()) else pd.NA,
                "base_rmc_flag": _first_non_empty_value(frame.loc[base_mask, "normalized_rmc_flag"]) if bool(base_mask.any()) else "N",
                "base_price_source_part_no": _first_non_empty_value(frame.loc[base_mask, "price_source_part_no"]) if bool(base_mask.any()) else "",
                "base_price_source_desc": _first_non_empty_value(frame.loc[base_mask, "price_source_desc"]) if bool(base_mask.any()) else "",
                "base_price_source_unit_price": _first_non_na_value(frame.loc[base_mask, "price_source_unit_price"]) if bool(base_mask.any()) else pd.NA,
                "base_price_source_rmc_flag": _first_non_empty_value(frame.loc[base_mask, "price_source_rmc_flag"]) if bool(base_mask.any()) else "N",
                "base_bom_qty": float(frame.loc[base_mask, "qty"].sum()) if bool(base_mask.any()) else 0.0,
                "base_new_qty": float(frame.loc[new_mask, "qty"].sum()) if bool(new_mask.any()) else 0.0,
                "current_base_part_exists": int(bool(base_mask.any() and frame.loc[base_mask, "qty"].gt(0).any())),
                "current_new_part_exists": int(bool(new_mask.any() and frame.loc[new_mask, "qty"].gt(0).any())),
                "new_unit_price": frame.loc[new_mask, "unit_price"].max() if bool(new_mask.any()) else pd.NA,
                "new_rmc_flag": _first_non_empty_value(frame.loc[new_mask, "normalized_rmc_flag"]) if bool(new_mask.any()) else "N",
                "new_price_source_part_no": _first_non_empty_value(frame.loc[new_mask, "price_source_part_no"]) if bool(new_mask.any()) else "",
                "new_price_source_desc": _first_non_empty_value(frame.loc[new_mask, "price_source_desc"]) if bool(new_mask.any()) else "",
                "new_price_source_unit_price": _first_non_na_value(frame.loc[new_mask, "price_source_unit_price"]) if bool(new_mask.any()) else pd.NA,
                "new_price_source_rmc_flag": _first_non_empty_value(frame.loc[new_mask, "price_source_rmc_flag"]) if bool(new_mask.any()) else "N",
                "current_base_qty": float(frame.loc[base_mask, "qty"].sum()) if bool(base_mask.any()) else 0.0,
                "current_new_qty": float(frame.loc[new_mask, "qty"].sum()) if bool(new_mask.any()) else 0.0,
            }
        )

    return pd.DataFrame(rows)


def _normalized_compare_text(value: object) -> str:
    return " ".join(_normalize_text(value).lower().split())


def _calculate_expected_vi_fields(calc_df: pd.DataFrame) -> tuple[list[float], list[str], list[str]]:
    expected_amounts: list[float] = []
    statuses: list[str] = []
    messages: list[str] = []

    for row in calc_df.itertuples(index=False):
        application_status = _normalize_text(getattr(row, "application_status", ""))
        production_qty = getattr(row, "production_qty", pd.NA)
        bom_qty = float(getattr(row, "bom_qty", 0.0) or 0.0)
        base_rmc_flag = _normalize_rmc_flag_value(getattr(row, "base_rmc_flag", ""))
        new_rmc_flag = _normalize_rmc_flag_value(getattr(row, "new_rmc_flag", ""))
        base_unit_price = getattr(row, "base_unit_price", pd.NA)
        new_unit_price = getattr(row, "new_unit_price", pd.NA)
        base_source_unit_price = getattr(row, "base_price_source_unit_price", pd.NA)
        new_source_unit_price = getattr(row, "new_price_source_unit_price", pd.NA)
        base_source_desc = getattr(row, "base_price_source_desc", "")
        new_source_desc = getattr(row, "new_price_source_desc", "")
        base_source_found = _normalize_rmc_flag_value(getattr(row, "base_price_source_rmc_flag", "")) == "Y"
        new_source_found = _normalize_rmc_flag_value(getattr(row, "new_price_source_rmc_flag", "")) == "Y"

        expected_amount = 0.0
        status = ""
        message = ""

        if application_status == "applied":
            status = "already_applied"
        elif application_status == "not_applied":
            if pd.isna(production_qty):
                status = "missing_volume"
                message = "Production volume is missing."
            elif bom_qty <= 0:
                status = "missing_bom_qty"
                message = "BOM qty is missing."
            elif {base_rmc_flag, new_rmc_flag} == {"Y", "N"}:
                status = "rmc_flag_mismatch"
                message = "RMC Flag mismatch for price comparison."
            else:
                if base_rmc_flag == "Y":
                    effective_base_price = base_unit_price
                else:
                    if not base_source_found:
                        effective_base_price = pd.NA
                        status = "no_rmc_price_source"
                        message = "No RMC=Y ancestor found for price comparison."
                    else:
                        effective_base_price = base_source_unit_price

                if not status:
                    if new_rmc_flag == "Y":
                        effective_new_price = new_unit_price
                    else:
                        if not new_source_found:
                            effective_new_price = pd.NA
                            status = "no_rmc_price_source"
                            message = "No RMC=Y ancestor found for price comparison."
                        else:
                            effective_new_price = new_source_unit_price
                else:
                    effective_new_price = pd.NA

                if not status and base_rmc_flag == "N" and new_rmc_flag == "N":
                    if _normalized_compare_text(base_source_desc) != _normalized_compare_text(new_source_desc):
                        status = "ancestor_desc_mismatch"
                        message = "Ancestor Desc mismatch for RMC=N price comparison."

                if not status and (pd.isna(effective_base_price) or pd.isna(effective_new_price)):
                    status = "missing_price"
                    message = "Unit price missing for expected VI calculation."

                if not status:
                    expected_amount = float((float(effective_base_price) - float(effective_new_price)) * bom_qty * float(production_qty))
                    status = "calculated"
        expected_amounts.append(expected_amount)
        statuses.append(status)
        messages.append(message)

    return expected_amounts, statuses, messages


def calculate_vi_outputs(
    base_bom_compact: pd.DataFrame,
    current_bom_compact: pd.DataFrame,
    vi_item_master: pd.DataFrame,
    vi_target_rule: pd.DataFrame,
    volume: pd.DataFrame,
    current_month: str,
    detail: bool = False,
    progress_callback: ProgressCallback | None = None,
) -> dict[str, pd.DataFrame]:
    require_duckdb()
    if base_bom_compact.empty:
        return _empty_outputs(detail)

    current_month = normalize_month_value(current_month)
    rules = _prepare_rules(vi_target_rule)
    item_master = vi_item_master.copy()
    for column in VI_ITEM_MASTER_COLUMNS:
        if column not in item_master.columns:
            item_master[column] = ""
        item_master[column] = item_master[column].fillna("").astype(str).str.strip()

    supported_rules = rules[
        (rules["effective_rule_type"] == "PART_NO")
        | (
            (rules["effective_rule_type"] == "CUSTOM")
            & (rules["effective_custom_rule_key"].isin(SUPPORTED_CUSTOM_RULE_KEYS))
        )
    ].copy()
    unsupported_rules = rules[
        (rules["effective_rule_type"] == "CUSTOM")
        & (~rules["effective_custom_rule_key"].isin(SUPPORTED_CUSTOM_RULE_KEYS))
    ].copy()

    volume_lookup = _build_volume_lookup(volume, current_month)
    volume_by_model = {
        (str(row["subsidiary"]).strip(), str(row["model_suffix"]).strip()): float(row["production_qty"])
        for _, row in volume_lookup.iterrows()
    }
    item_name_lookup = (
        item_master[["vi_item_id", "vi_item_name"]]
        .drop_duplicates(subset=["vi_item_id"])
        .set_index("vi_item_id")["vi_item_name"]
        .to_dict()
    )
    current_presence_keys = set()
    if not current_bom_compact.empty:
        current_presence_keys = set(
            zip(
                current_bom_compact["subsidiary"].fillna("").astype(str).str.strip(),
                current_bom_compact["model_suffix"].fillna("").astype(str).str.strip(),
            )
        )

    join_started_at = time.perf_counter()
    per_rule_started_at = time.perf_counter()
    base_target_query_seconds = 0.0
    current_base_query_seconds = 0.0
    current_new_query_seconds = 0.0
    current_new_anywhere_query_seconds = 0.0
    context_summary_seconds = 0.0
    con = create_duckdb_connection()
    try:
        con.register("base_compact", base_bom_compact)
        con.register("current_compact", current_bom_compact)
        part_no_rules = supported_rules[supported_rules["effective_rule_type"].eq("PART_NO")].copy()
        base_part_candidates_all = [
            _normalize_text(value)
            for value in part_no_rules["base_part_no"].tolist()
            if _normalize_text(value)
        ]
        new_part_candidates_all = [
            _normalize_text(value)
            for value in part_no_rules["new_part_no"].tolist()
            if _normalize_text(value)
        ]
        candidate_pool_started_at = time.perf_counter()
        _materialize_part_candidate_pool(
            con,
            pool_table="base_candidate_pool",
            source_table="base_compact",
            part_numbers=base_part_candidates_all,
            require_positive_qty=False,
        )
        _materialize_part_candidate_pool(
            con,
            pool_table="current_base_candidate_pool",
            source_table="current_compact",
            part_numbers=base_part_candidates_all,
            require_positive_qty=True,
        )
        _materialize_part_candidate_pool(
            con,
            pool_table="current_new_candidate_pool",
            source_table="current_compact",
            part_numbers=new_part_candidates_all,
            require_positive_qty=True,
        )
        con.execute(
            """
            CREATE OR REPLACE TEMP TABLE current_new_anywhere_pool AS
            SELECT *
            FROM current_new_candidate_pool
            """
        )
        _log_timing("prepare candidate pools", candidate_pool_started_at)
        _notify_progress(progress_callback, 0.5, "Generating base target model list")
        base_target_rows: list[dict[str, object]] = []
        base_compare_rows: list[dict[str, object]] = []
        added_compare_rows: list[dict[str, object]] = []

        def _strict_context_key(row_dict: dict[str, object]) -> tuple[object, ...]:
            return (
                row_dict.get("subsidiary", ""),
                row_dict.get("model_suffix", ""),
                row_dict.get("model_type", ""),
                row_dict.get("indoor_tool", ""),
                row_dict.get("parent_assy_part_no", ""),
                row_dict.get("parent_assy_desc_text", ""),
                row_dict.get("ancestor_desc_path", ""),
            )

        total_rules = max(len(supported_rules.index), 1)
        for rule_index, (_, rule) in enumerate(supported_rules.iterrows(), start=1):
            progress_value = 0.5 + (0.24 * rule_index / total_rules)
            _notify_progress(progress_callback, progress_value, f"Evaluating final rules ({rule_index}/{total_rules})")
            vi_item_id = _normalize_text(rule.get("vi_item_id", ""))
            vi_item_name = item_name_lookup.get(vi_item_id, "")
            rule_type = _normalize_text(rule.get("rule_type", "")).upper()
            rule_subsidiary_scope = normalize_filter_value(rule.get("subsidiary", ""))
            base_part_no = _normalize_text(rule.get("base_part_no", ""))
            new_part_no = _normalize_text(rule.get("new_part_no", ""))
            effective_rule_type = _normalize_text(rule.get("effective_rule_type", "")).upper()
            include_parent_columns = _rule_has_row_context(rule)

            base_part_candidates = [base_part_no] if effective_rule_type == "PART_NO" and base_part_no else None
            base_source_table = "base_candidate_pool" if effective_rule_type == "PART_NO" else "base_compact"
            current_base_source_table = "current_base_candidate_pool" if effective_rule_type == "PART_NO" else "current_compact"
            current_new_source_table = "current_new_candidate_pool" if effective_rule_type == "PART_NO" else "current_compact"
            current_new_anywhere_source_table = "current_new_anywhere_pool" if effective_rule_type == "PART_NO" else "current_compact"
            query_started_at = time.perf_counter()
            base_matches = _query_rule_candidate_rows(
                con,
                base_source_table,
                rule,
                part_candidates=base_part_candidates,
                require_positive_qty=False,
                include_parent_columns=include_parent_columns,
                apply_desc_spec_filters=True,
            )
            base_target_query_seconds += time.perf_counter() - query_started_at
            summary_started_at = time.perf_counter()
            base_contexts = _summarize_context_matches(
                base_matches,
                base_part_no=base_part_no,
                new_part_no=new_part_no,
                include_parent_columns=include_parent_columns,
            )
            context_summary_seconds += time.perf_counter() - summary_started_at
            if effective_rule_type == "PART_NO":
                query_started_at = time.perf_counter()
                current_base_matches = _query_rule_candidate_rows(
                    con,
                    current_base_source_table,
                    rule,
                    part_candidates=[base_part_no] if base_part_no else [],
                    require_positive_qty=True,
                    include_parent_columns=include_parent_columns,
                    apply_desc_spec_filters=True,
                )
                current_base_query_seconds += time.perf_counter() - query_started_at
                query_started_at = time.perf_counter()
                current_new_matches = _query_rule_candidate_rows(
                    con,
                    current_new_source_table,
                    rule,
                    part_candidates=[new_part_no] if new_part_no else [],
                    require_positive_qty=True,
                    include_parent_columns=include_parent_columns,
                    apply_desc_spec_filters=False,
                )
                current_new_query_seconds += time.perf_counter() - query_started_at
                query_started_at = time.perf_counter()
                current_new_anywhere_matches = _query_rule_candidate_rows(
                    con,
                    current_new_anywhere_source_table,
                    rule,
                    part_candidates=[new_part_no] if new_part_no else [],
                    require_positive_qty=True,
                    include_parent_columns=False,
                    apply_desc_spec_filters=False,
                )
                current_new_anywhere_query_seconds += time.perf_counter() - query_started_at
            else:
                query_started_at = time.perf_counter()
                current_base_matches = _query_rule_candidate_rows(
                    con,
                    current_base_source_table,
                    rule,
                    part_candidates=None,
                    require_positive_qty=True,
                    include_parent_columns=include_parent_columns,
                    apply_desc_spec_filters=True,
                )
                current_base_query_seconds += time.perf_counter() - query_started_at
                current_new_matches = _empty_frame(list(current_base_matches.columns))
                current_new_anywhere_matches = _empty_frame(list(current_base_matches.columns))

            summary_started_at = time.perf_counter()
            current_base_contexts = _summarize_context_matches(
                current_base_matches,
                base_part_no=base_part_no,
                new_part_no=new_part_no,
                include_parent_columns=include_parent_columns,
            )
            current_new_contexts = _summarize_context_matches(
                current_new_matches,
                base_part_no=base_part_no,
                new_part_no=new_part_no,
                include_parent_columns=include_parent_columns,
            )
            current_new_anywhere_lookup = _build_model_presence_lookup(current_new_anywhere_matches)
            context_summary_seconds += time.perf_counter() - summary_started_at

            if not base_contexts.empty:
                for _, context_row in base_contexts.iterrows():
                    base_target_rows.append(
                        {
                            "vi_item_id": vi_item_id,
                            "vi_item_name": vi_item_name,
                            "rule_type": rule_type,
                            "subsidiary": context_row["subsidiary"],
                            "model_suffix": context_row["model_suffix"],
                            "model_type": context_row["model_type"],
                            "indoor_tool": context_row["indoor_tool"],
                            "parent_assy_part_no": context_row["parent_assy_part_no"],
                            "parent_assy_desc_text": context_row["parent_assy_desc_text"],
                            "base_part_no": base_part_no,
                            "new_part_no": new_part_no,
                            "base_unit_price": context_row["base_unit_price"],
                            "base_bom_qty": context_row["base_bom_qty"],
                            "target_source": "BASE_TARGET",
                            "ancestor_desc_path": context_row.get("ancestor_desc_path", ""),
                        }
                    )

            current_base_context_lookup = _aggregate_contexts_by_business_key(current_base_contexts)
            current_new_context_lookup = _aggregate_contexts_by_business_key(current_new_contexts)
            base_business_keys = set()
            base_model_keys = set()
            for _, context_row in base_contexts.iterrows():
                context_dict = context_row.to_dict()
                strict_key = _strict_context_key(context_dict)
                key = _business_context_key(context_dict)
                base_business_keys.add(key)
                base_model_keys.add((context_dict["subsidiary"], context_dict["model_suffix"]))
                current_base_context = current_base_context_lookup.get(key, {})
                current_new_context = current_new_context_lookup.get(key, {})
                model_key = (str(context_dict["subsidiary"]).strip(), str(context_dict["model_suffix"]).strip())
                new_anywhere_context = current_new_anywhere_lookup.get(model_key, {})
                matched_context_method = current_new_context.get(
                    "matched_context_method",
                    current_base_context.get("matched_context_method", _business_context_value(context_dict)[1]),
                )
                base_context_value, _ = _business_context_value(context_dict)
                base_compare_rows.append(
                    {
                        "vi_item_id": vi_item_id,
                        "vi_item_name": vi_item_name,
                        "rule_type": rule_type,
                        "rule_subsidiary_scope": rule_subsidiary_scope,
                        "subsidiary": context_dict["subsidiary"],
                        "model_suffix": context_dict["model_suffix"],
                        "model_type": context_dict["model_type"],
                        "indoor_tool": context_dict["indoor_tool"],
                        "parent_assy_part_no": context_dict["parent_assy_part_no"],
                        "parent_assy_desc_text": context_dict["parent_assy_desc_text"],
                        "ancestor_desc_path": context_dict.get("ancestor_desc_path", ""),
                        "base_part_no": base_part_no,
                        "new_part_no": new_part_no,
                        "base_part_desc": context_dict["base_part_desc"],
                        "base_exists": 1,
                        "current_exists": int(model_key in current_presence_keys),
                        "base_target_flag": 1,
                        "current_base_part_exists": int(current_base_context.get("current_base_part_exists", 0) or 0),
                        "current_new_part_exists": int(current_new_context.get("current_new_part_exists", 0) or 0),
                        "base_target_parent_assy_part_no": context_dict["parent_assy_part_no"],
                        "base_target_parent_assy_desc_text": context_dict["parent_assy_desc_text"],
                        "base_target_ancestor_desc_path": context_dict.get("ancestor_desc_path", ""),
                        "current_new_parent_assy_part_no": current_new_context.get("parent_assy_part_no", ""),
                        "current_new_parent_assy_desc_text": current_new_context.get("parent_assy_desc_text", ""),
                        "current_new_ancestor_desc_path": current_new_context.get("ancestor_desc_path", ""),
                        "matched_context_method": matched_context_method,
                        "new_part_found_under_desc_context": "Y" if int(current_new_context.get("current_new_part_exists", 0) or 0) else "N",
                        "new_part_found_anywhere_in_model": "Y" if int(new_anywhere_context.get("exists", 0) or 0) else "N",
                        "base_unit_price": context_dict["base_unit_price"],
                        "new_unit_price": current_new_context.get("new_unit_price", new_anywhere_context.get("new_unit_price", pd.NA)),
                        "base_rmc_flag": context_dict.get("base_rmc_flag", "N"),
                        "new_rmc_flag": current_new_context.get("new_rmc_flag", new_anywhere_context.get("new_rmc_flag", "N")),
                        "base_price_source_part_no": context_dict.get("base_price_source_part_no", ""),
                        "base_price_source_desc": context_dict.get("base_price_source_desc", ""),
                        "base_price_source_unit_price": context_dict.get("base_price_source_unit_price", pd.NA),
                        "base_price_source_rmc_flag": context_dict.get("base_price_source_rmc_flag", "N"),
                        "new_price_source_part_no": current_new_context.get("new_price_source_part_no", new_anywhere_context.get("new_price_source_part_no", "")),
                        "new_price_source_desc": current_new_context.get("new_price_source_desc", new_anywhere_context.get("new_price_source_desc", "")),
                        "new_price_source_unit_price": current_new_context.get("new_price_source_unit_price", new_anywhere_context.get("new_price_source_unit_price", pd.NA)),
                        "new_price_source_rmc_flag": current_new_context.get("new_price_source_rmc_flag", new_anywhere_context.get("new_price_source_rmc_flag", "N")),
                        "base_bom_qty": float(context_dict.get("base_bom_qty", 0.0) or 0.0),
                        "base_new_qty": float(context_dict.get("base_new_qty", 0.0) or 0.0),
                        "current_base_qty": float(current_base_context.get("current_base_qty", 0.0) or 0.0),
                        "current_new_qty": float(current_new_context.get("current_new_qty", 0.0) or 0.0),
                        "production_qty": volume_by_model.get(model_key, pd.NA),
                        "_base_strict_context_key": str(strict_key),
                        "_business_context_value": base_context_value,
                    }
                )

            current_context_lookup = dict(current_base_context_lookup)
            current_context_lookup.update(current_new_context_lookup)
            for business_key, context_dict in current_context_lookup.items():
                model_key = (context_dict["subsidiary"], context_dict["model_suffix"])
                if business_key in base_business_keys or model_key in base_model_keys:
                    continue
                new_anywhere_context = current_new_anywhere_lookup.get(model_key, {})
                added_compare_rows.append(
                    {
                        "vi_item_id": vi_item_id,
                        "vi_item_name": vi_item_name,
                        "rule_type": rule_type,
                        "rule_subsidiary_scope": rule_subsidiary_scope,
                        "subsidiary": context_dict["subsidiary"],
                        "model_suffix": context_dict["model_suffix"],
                        "model_type": context_dict["model_type"],
                        "indoor_tool": context_dict["indoor_tool"],
                        "parent_assy_part_no": context_dict["parent_assy_part_no"],
                        "parent_assy_desc_text": context_dict["parent_assy_desc_text"],
                        "ancestor_desc_path": context_dict.get("ancestor_desc_path", ""),
                        "base_exists": 0,
                        "current_exists": 1,
                        "base_target_flag": 0,
                        "base_part_no": base_part_no,
                        "new_part_no": new_part_no,
                        "base_part_desc": context_dict["base_part_desc"],
                        "base_target_parent_assy_part_no": "",
                        "base_target_parent_assy_desc_text": "",
                        "base_target_ancestor_desc_path": "",
                        "current_new_parent_assy_part_no": context_dict.get("parent_assy_part_no", ""),
                        "current_new_parent_assy_desc_text": context_dict.get("parent_assy_desc_text", ""),
                        "current_new_ancestor_desc_path": context_dict.get("ancestor_desc_path", ""),
                        "matched_context_method": context_dict.get("matched_context_method", ""),
                        "new_part_found_under_desc_context": "Y" if int(context_dict.get("current_new_part_exists", 0) or 0) else "N",
                        "new_part_found_anywhere_in_model": "Y" if int(new_anywhere_context.get("exists", 0) or 0) else "N",
                        "current_base_part_exists": int(context_dict.get("current_base_part_exists", 0) or 0),
                        "current_new_part_exists": int(context_dict.get("current_new_part_exists", 0) or 0),
                        "base_unit_price": pd.NA,
                        "new_unit_price": context_dict.get("new_unit_price", new_anywhere_context.get("new_unit_price", pd.NA)),
                        "base_rmc_flag": "N",
                        "new_rmc_flag": context_dict.get("new_rmc_flag", new_anywhere_context.get("new_rmc_flag", "N")),
                        "base_price_source_part_no": "",
                        "base_price_source_desc": "",
                        "base_price_source_unit_price": pd.NA,
                        "base_price_source_rmc_flag": "N",
                        "new_price_source_part_no": context_dict.get("new_price_source_part_no", new_anywhere_context.get("new_price_source_part_no", "")),
                        "new_price_source_desc": context_dict.get("new_price_source_desc", new_anywhere_context.get("new_price_source_desc", "")),
                        "new_price_source_unit_price": context_dict.get("new_price_source_unit_price", new_anywhere_context.get("new_price_source_unit_price", pd.NA)),
                        "new_price_source_rmc_flag": context_dict.get("new_price_source_rmc_flag", new_anywhere_context.get("new_price_source_rmc_flag", "N")),
                        "base_bom_qty": 0.0,
                        "base_new_qty": 0.0,
                        "current_base_qty": float(context_dict.get("current_base_qty", 0.0) or 0.0),
                        "current_new_qty": float(context_dict.get("current_new_qty", 0.0) or 0.0),
                        "production_qty": volume_by_model.get(model_key, pd.NA),
                    }
                )

        base_target_df = pd.DataFrame(base_target_rows)
        _log_timing("base target list generation", join_started_at)

        current_compare_started_at = time.perf_counter()
        _notify_progress(progress_callback, 0.64, "Comparing base target models with current models")
        base_compare_df = pd.DataFrame(base_compare_rows)
        _log_timing("current model comparison", current_compare_started_at)

        classify_started_at = time.perf_counter()
        _notify_progress(progress_callback, 0.76, "Classifying added and removed models")
        added_compare_df = pd.DataFrame(added_compare_rows)
        _log_timing("added/removed model classification", classify_started_at)
    finally:
        con.close()
    _log_timing("join operations", join_started_at)
    _log_timing("per-rule evaluation total", per_rule_started_at)
    _log_elapsed("base target query total", base_target_query_seconds)
    _log_elapsed("current base query total", current_base_query_seconds)
    _log_elapsed("current new query total", current_new_query_seconds + current_new_anywhere_query_seconds)
    _log_elapsed("context summary total", context_summary_seconds)
    _notify_progress(progress_callback, 0.82, "Computing business flags")

    if base_target_df.empty:
        base_target_df = _empty_frame(
            [
                "vi_item_id",
                "vi_item_name",
                "rule_type",
                "subsidiary",
                "model_suffix",
                "model_type",
                "indoor_tool",
                "parent_assy_part_no",
                "parent_assy_desc_text",
                "base_part_no",
                "new_part_no",
                "base_unit_price",
                "base_bom_qty",
                "target_source",
            ]
        )
    if base_compare_df.empty:
        base_compare_df = _empty_frame(
            [
                "vi_item_id",
                "vi_item_name",
                "rule_type",
                "rule_subsidiary_scope",
                "subsidiary",
                "model_suffix",
                "model_type",
                "indoor_tool",
                "parent_assy_part_no",
                "parent_assy_desc_text",
                "ancestor_desc_path",
                "base_part_no",
                "new_part_no",
                "base_part_desc",
                "base_exists",
                "current_exists",
                "base_target_flag",
                "current_base_part_exists",
                "current_new_part_exists",
                "base_target_parent_assy_part_no",
                "base_target_parent_assy_desc_text",
                "base_target_ancestor_desc_path",
                "current_new_parent_assy_part_no",
                "current_new_parent_assy_desc_text",
                "current_new_ancestor_desc_path",
                "matched_context_method",
                "new_part_found_under_desc_context",
                "new_part_found_anywhere_in_model",
                "base_unit_price",
                "new_unit_price",
                "base_rmc_flag",
                "new_rmc_flag",
                "base_price_source_part_no",
                "base_price_source_desc",
                "base_price_source_unit_price",
                "base_price_source_rmc_flag",
                "new_price_source_part_no",
                "new_price_source_desc",
                "new_price_source_unit_price",
                "new_price_source_rmc_flag",
                "base_bom_qty",
                "base_new_qty",
                "current_base_qty",
                "current_new_qty",
                "production_qty",
            ]
        )
    if added_compare_df.empty:
        added_compare_df = _empty_frame(list(base_compare_df.columns))

    if base_compare_df.empty and added_compare_df.empty:
        outputs = _empty_outputs(detail)
        if detail and not unsupported_rules.empty:
            unsupported_named = unsupported_rules.merge(item_master[["vi_item_id", "vi_item_name"]], how="left", on="vi_item_id")
            outputs["ViExceptionLog"] = unsupported_named.assign(
                model_suffix="",
                issue_type="missing_custom_rule",
                issue_message="No supported set-based custom handler is registered for this custom_rule_key.",
            )[["vi_item_id", "vi_item_name", "rule_type", "subsidiary", "model_suffix", "issue_type", "issue_message"]]
        return outputs

    base_target_model_list = base_target_df[
        [
            "vi_item_id",
            "vi_item_name",
            "rule_type",
            "subsidiary",
            "model_suffix",
            "model_type",
            "indoor_tool",
            "parent_assy_part_no",
            "parent_assy_desc_text",
            "ancestor_desc_path",
            "base_part_no",
            "new_part_no",
            "base_unit_price",
            "base_bom_qty",
            "target_source",
        ]
    ].copy()
    if not base_target_model_list.empty:
        base_target_model_list["base_unit_price"] = pd.to_numeric(base_target_model_list["base_unit_price"], errors="coerce")
        base_target_model_list["base_bom_qty"] = pd.to_numeric(base_target_model_list["base_bom_qty"], errors="coerce").fillna(0.0)

    calc_df = pd.concat([base_compare_df, added_compare_df], ignore_index=True, sort=False)
    numeric_columns = [
        "base_exists",
        "current_exists",
        "base_target_flag",
        "current_base_part_exists",
        "current_new_part_exists",
        "base_bom_qty",
        "base_new_qty",
        "current_base_qty",
        "current_new_qty",
    ]
    for column in numeric_columns:
        calc_df[column] = pd.to_numeric(calc_df[column], errors="coerce").fillna(0.0)
    calc_df["base_unit_price"] = pd.to_numeric(calc_df["base_unit_price"], errors="coerce")
    calc_df["new_unit_price"] = pd.to_numeric(calc_df["new_unit_price"], errors="coerce")
    calc_df["base_price_source_unit_price"] = pd.to_numeric(calc_df["base_price_source_unit_price"], errors="coerce")
    calc_df["new_price_source_unit_price"] = pd.to_numeric(calc_df["new_price_source_unit_price"], errors="coerce")
    calc_df["production_qty"] = pd.to_numeric(calc_df["production_qty"], errors="coerce")

    base_target_mask = calc_df["base_target_flag"].eq(1)
    added_mask = ~base_target_mask
    no_part_rule = calc_df["base_part_no"].fillna("").eq("") & calc_df["new_part_no"].fillna("").eq("")

    calc_df["model_change_status"] = "existing_model"
    calc_df.loc[base_target_mask & calc_df["current_exists"].eq(0), "model_change_status"] = "removed_model"
    calc_df.loc[added_mask, "model_change_status"] = "added_model"

    calc_df["application_status"] = "missing_or_model_changed"
    calc_df.loc[base_target_mask & no_part_rule, "application_status"] = "target_only"
    calc_df.loc[base_target_mask & calc_df["current_exists"].eq(0), "application_status"] = "missing_or_model_changed"
    calc_df.loc[base_target_mask & calc_df["current_exists"].eq(1) & calc_df["current_new_part_exists"].eq(1) & calc_df["current_base_part_exists"].eq(0), "application_status"] = "applied"
    calc_df.loc[base_target_mask & calc_df["current_exists"].eq(1) & calc_df["current_base_part_exists"].eq(1) & calc_df["current_new_part_exists"].eq(0), "application_status"] = "not_applied"
    calc_df.loc[base_target_mask & calc_df["current_exists"].eq(1) & calc_df["current_base_part_exists"].eq(1) & calc_df["current_new_part_exists"].eq(1), "application_status"] = "mixed_review"

    calc_df.loc[added_mask & no_part_rule, "application_status"] = "added_not_target_or_unknown"
    calc_df.loc[added_mask & calc_df["current_new_part_exists"].eq(1) & calc_df["current_base_part_exists"].eq(0), "application_status"] = "added_applied"
    calc_df.loc[added_mask & calc_df["current_base_part_exists"].eq(1) & calc_df["current_new_part_exists"].eq(0), "application_status"] = "added_not_applied"
    calc_df.loc[added_mask & calc_df["current_base_part_exists"].eq(1) & calc_df["current_new_part_exists"].eq(1), "application_status"] = "added_mixed_review"

    calc_df["bom_qty"] = 0.0
    calc_df.loc[calc_df["current_new_qty"] > 0, "bom_qty"] = calc_df.loc[calc_df["current_new_qty"] > 0, "current_new_qty"]
    calc_df.loc[calc_df["bom_qty"].eq(0) & (calc_df["current_base_qty"] > 0), "bom_qty"] = calc_df.loc[calc_df["bom_qty"].eq(0) & (calc_df["current_base_qty"] > 0), "current_base_qty"]
    calc_df.loc[calc_df["bom_qty"].eq(0) & (calc_df["base_new_qty"] > 0), "bom_qty"] = calc_df.loc[calc_df["bom_qty"].eq(0) & (calc_df["base_new_qty"] > 0), "base_new_qty"]
    calc_df.loc[calc_df["bom_qty"].eq(0) & (calc_df["base_bom_qty"] > 0), "bom_qty"] = calc_df.loc[calc_df["bom_qty"].eq(0) & (calc_df["base_bom_qty"] > 0), "base_bom_qty"]

    financial_status_mask = calc_df["application_status"].isin(["applied", "not_applied"])
    calc_df["unit_saving"] = 0.0
    calculable = (
        financial_status_mask
        & calc_df["base_unit_price"].notna()
        & calc_df["new_unit_price"].notna()
        & calc_df["production_qty"].notna()
        & calc_df["bom_qty"].gt(0)
    )
    calc_df.loc[calculable, "unit_saving"] = calc_df.loc[calculable, "base_unit_price"] - calc_df.loc[calculable, "new_unit_price"]
    calc_df["realized_vi_amount"] = 0.0
    calc_df["remaining_vi_amount"] = 0.0
    calc_df["expected_vi_amount_usd"] = 0.0
    calc_df["expected_vi_calc_status"] = ""
    calc_df["expected_vi_calc_message"] = ""
    calc_df.loc[calculable & calc_df["application_status"].eq("applied"), "realized_vi_amount"] = (
        calc_df.loc[calculable & calc_df["application_status"].eq("applied"), "unit_saving"]
        * calc_df.loc[calculable & calc_df["application_status"].eq("applied"), "bom_qty"]
        * calc_df.loc[calculable & calc_df["application_status"].eq("applied"), "production_qty"]
    )
    calc_df.loc[calculable & calc_df["application_status"].eq("not_applied"), "remaining_vi_amount"] = (
        calc_df.loc[calculable & calc_df["application_status"].eq("not_applied"), "unit_saving"]
        * calc_df.loc[calculable & calc_df["application_status"].eq("not_applied"), "bom_qty"]
        * calc_df.loc[calculable & calc_df["application_status"].eq("not_applied"), "production_qty"]
    )
    calc_df.loc[calculable & calc_df["application_status"].eq("not_applied"), "expected_vi_amount_usd"] = (
        calc_df.loc[calculable & calc_df["application_status"].eq("not_applied"), "remaining_vi_amount"]
    )
    expected_amounts, expected_statuses, expected_messages = _calculate_expected_vi_fields(calc_df)
    calc_df["expected_vi_amount_usd"] = expected_amounts
    calc_df["expected_vi_calc_status"] = expected_statuses
    calc_df["expected_vi_calc_message"] = expected_messages

    aggregation_started_at = time.perf_counter()
    base_summary_basis = calc_df[base_target_mask].copy()
    if not base_summary_basis.empty:
        base_summary_basis["status_priority"] = base_summary_basis["application_status"].map(STATUS_PRIORITY).fillna(0)
        base_summary_basis = base_summary_basis.sort_values(
            ["vi_item_id", "subsidiary", "model_suffix", "status_priority"],
            ascending=[True, True, True, False],
        ).drop_duplicates(subset=["vi_item_id", "subsidiary", "model_suffix"], keep="first")

    vi_item_summary = (
        base_summary_basis.groupby(["vi_item_id", "vi_item_name"], as_index=False)
        .agg(
            target_model_count=("model_suffix", "count"),
            applied_model_count=("application_status", lambda s: int((s == "applied").sum())),
            not_applied_model_count=("application_status", lambda s: int((s == "not_applied").sum())),
            mixed_review_model_count=("application_status", lambda s: int((s == "mixed_review").sum())),
            missing_or_model_changed_count=("application_status", lambda s: int((s == "missing_or_model_changed").sum())),
            realized_vi_amount=("realized_vi_amount", "sum"),
            remaining_vi_amount=("remaining_vi_amount", "sum"),
            expected_vi_amount_usd=("expected_vi_amount_usd", "sum"),
        )
        .reset_index(drop=True)
    )
    vi_item_summary["application_rate"] = (
        vi_item_summary["applied_model_count"] / vi_item_summary["target_model_count"].replace(0, pd.NA)
    ).fillna(0.0)
    vi_item_summary["total_opportunity"] = vi_item_summary["realized_vi_amount"] + vi_item_summary["remaining_vi_amount"]

    vi_subsidiary_summary = (
        base_summary_basis.groupby(["subsidiary"], as_index=False)
        .agg(
            target_model_count=("model_suffix", "count"),
            applied_model_count=("application_status", lambda s: int((s == "applied").sum())),
            not_applied_model_count=("application_status", lambda s: int((s == "not_applied").sum())),
            mixed_review_model_count=("application_status", lambda s: int((s == "mixed_review").sum())),
            missing_or_model_changed_count=("application_status", lambda s: int((s == "missing_or_model_changed").sum())),
            realized_vi_amount=("realized_vi_amount", "sum"),
            remaining_vi_amount=("remaining_vi_amount", "sum"),
            expected_vi_amount_usd=("expected_vi_amount_usd", "sum"),
        )
        .reset_index(drop=True)
    )
    vi_subsidiary_summary["application_rate"] = (
        vi_subsidiary_summary["applied_model_count"] / vi_subsidiary_summary["target_model_count"].replace(0, pd.NA)
    ).fillna(0.0)
    vi_subsidiary_summary["total_opportunity"] = vi_subsidiary_summary["realized_vi_amount"] + vi_subsidiary_summary["remaining_vi_amount"]

    model_change_summary = (
        calc_df.groupby(["vi_item_id", "vi_item_name", "subsidiary"], as_index=False)
        .agg(
            base_target_model_count=("base_target_flag", lambda s: int((s == 1).sum())),
            existing_model_count=("model_change_status", lambda s: int((s == "existing_model").sum())),
            removed_model_count=("model_change_status", lambda s: int((s == "removed_model").sum())),
            added_model_count=("model_change_status", lambda s: int((s == "added_model").sum())),
            base_target_applied_count=("application_status", lambda s: int((s == "applied").sum())),
            base_target_not_applied_count=("application_status", lambda s: int((s == "not_applied").sum())),
            base_target_mixed_review_count=("application_status", lambda s: int((s == "mixed_review").sum())),
            base_target_missing_or_model_changed_count=("application_status", lambda s: int((s == "missing_or_model_changed").sum())),
            added_applied_count=("application_status", lambda s: int((s == "added_applied").sum())),
            added_not_applied_count=("application_status", lambda s: int((s == "added_not_applied").sum())),
            added_mixed_review_count=("application_status", lambda s: int((s == "added_mixed_review").sum())),
            added_not_target_or_unknown_count=("application_status", lambda s: int((s == "added_not_target_or_unknown").sum())),
        )
        .reset_index(drop=True)
    )

    base_pn_target_summary = _empty_frame(
        [
            "vi_item_id",
            "vi_item_name",
            "subsidiary",
            "base_part_no",
            "base_part_desc",
            "base_pn_model_count",
            "item_target_model_count",
            "base_pn_model_share_pct",
            "applied_model_count",
            "not_applied_model_count",
            "mixed_review_model_count",
            "missing_or_model_changed_count",
            "realized_vi_amount",
            "remaining_vi_amount",
        ]
    )
    base_pn_basis = calc_df[base_target_mask].copy()
    if not base_pn_basis.empty:
        base_pn_basis["status_priority"] = base_pn_basis["application_status"].map(STATUS_PRIORITY).fillna(0)
        item_scope_basis = (
            base_pn_basis.sort_values(
                ["vi_item_id", "rule_subsidiary_scope", "subsidiary", "model_suffix", "status_priority"],
                ascending=[True, True, True, True, False],
            )
            .drop_duplicates(subset=["vi_item_id", "rule_subsidiary_scope", "subsidiary", "model_suffix"], keep="first")
            .reset_index(drop=True)
        )
        item_target_counts = (
            item_scope_basis.groupby(["vi_item_id", "rule_subsidiary_scope"], as_index=False)
            .agg(item_target_model_count=("model_suffix", "count"))
            .reset_index(drop=True)
        )

        base_pn_basis = (
            base_pn_basis[base_pn_basis["base_part_no"].fillna("").ne("")]
            .sort_values(
                ["vi_item_id", "rule_subsidiary_scope", "base_part_no", "subsidiary", "model_suffix", "status_priority"],
                ascending=[True, True, True, True, True, False],
            )
            .drop_duplicates(
                subset=["vi_item_id", "rule_subsidiary_scope", "base_part_no", "subsidiary", "model_suffix"],
                keep="first",
            )
            .reset_index(drop=True)
        )
        if not base_pn_basis.empty:
            base_pn_target_summary = (
                base_pn_basis.groupby(
                    ["vi_item_id", "vi_item_name", "rule_subsidiary_scope", "base_part_no"],
                    as_index=False,
                )
                .agg(
                    base_part_desc=(
                        "base_part_desc",
                        lambda s: next((str(value).strip() for value in s if str(value).strip()), ""),
                    ),
                    base_pn_model_count=("model_suffix", "count"),
                    applied_model_count=("application_status", lambda s: int((s == "applied").sum())),
                    not_applied_model_count=("application_status", lambda s: int((s == "not_applied").sum())),
                    mixed_review_model_count=("application_status", lambda s: int((s == "mixed_review").sum())),
                    missing_or_model_changed_count=(
                        "application_status",
                        lambda s: int((s == "missing_or_model_changed").sum()),
                    ),
                    realized_vi_amount=("realized_vi_amount", "sum"),
                    remaining_vi_amount=("remaining_vi_amount", "sum"),
                    expected_vi_amount_usd=("expected_vi_amount_usd", "sum"),
                )
                .merge(item_target_counts, how="left", on=["vi_item_id", "rule_subsidiary_scope"])
                .rename(columns={"rule_subsidiary_scope": "subsidiary"})
                .reset_index(drop=True)
            )
            base_pn_target_summary["item_target_model_count"] = pd.to_numeric(
                base_pn_target_summary["item_target_model_count"], errors="coerce"
            ).fillna(0.0)
            base_pn_target_summary["base_pn_model_share_pct"] = (
                base_pn_target_summary["base_pn_model_count"]
                / base_pn_target_summary["item_target_model_count"].replace(0, pd.NA)
                * 100.0
            ).fillna(0.0)
            base_pn_target_summary = base_pn_target_summary[
                [
                    "vi_item_id",
                    "vi_item_name",
                    "subsidiary",
                    "base_part_no",
                    "base_part_desc",
                    "base_pn_model_count",
                    "item_target_model_count",
                    "base_pn_model_share_pct",
                    "applied_model_count",
                    "not_applied_model_count",
                    "mixed_review_model_count",
                    "missing_or_model_changed_count",
                    "realized_vi_amount",
                    "remaining_vi_amount",
                    "expected_vi_amount_usd",
                ]
            ].copy()
    _log_timing("summary aggregation", aggregation_started_at)
    _notify_progress(progress_callback, 0.92, "Building output tables")

    outputs = {
        "ViItemSummary": round_financial_columns(vi_item_summary),
        "ViSubsidiarySummary": round_financial_columns(vi_subsidiary_summary),
        "BaseTargetModelList": round_financial_columns(base_target_model_list),
        "BasePNTargetSummary": round_financial_columns(base_pn_target_summary),
        "ModelChangeSummary": model_change_summary,
    }

    unsupported_issue_frame = _empty_frame(["vi_item_id", "vi_item_name", "rule_type", "subsidiary", "model_suffix", "issue_type", "issue_message"])
    if not unsupported_rules.empty:
        unsupported_named = unsupported_rules.merge(
            item_master[["vi_item_id", "vi_item_name"]],
            how="left",
            on="vi_item_id",
        )
        unsupported_issue_frame = unsupported_named.assign(
            model_suffix="",
            issue_type="missing_custom_rule",
            issue_message="No supported set-based custom handler is registered for this custom_rule_key.",
        )[
            ["vi_item_id", "vi_item_name", "rule_type", "subsidiary", "model_suffix", "issue_type", "issue_message"]
        ]

    if not detail:
        _log_elapsed("detail table generation", 0.0)
        outputs["ModelChangeDetail"] = _empty_frame([])
        outputs["ViModelDetail"] = _empty_frame([])
        outputs["ViExceptionLog"] = _empty_frame([])
        return outputs

    detail_started_at = time.perf_counter()
    model_change_detail = calc_df[
        [
            "vi_item_id",
            "vi_item_name",
            "rule_type",
            "subsidiary",
            "model_suffix",
            "model_type",
            "indoor_tool",
            "parent_assy_part_no",
            "parent_assy_desc_text",
            "ancestor_desc_path",
            "model_change_status",
            "base_exists",
            "current_exists",
            "base_target_flag",
            "base_part_no",
            "new_part_no",
            "current_base_part_exists",
            "current_new_part_exists",
            "application_status",
            "base_target_parent_assy_part_no",
            "base_target_parent_assy_desc_text",
            "base_target_ancestor_desc_path",
            "current_new_parent_assy_part_no",
            "current_new_parent_assy_desc_text",
            "current_new_ancestor_desc_path",
            "matched_context_method",
            "new_part_found_under_desc_context",
            "new_part_found_anywhere_in_model",
            "base_unit_price",
            "new_unit_price",
            "base_rmc_flag",
            "new_rmc_flag",
            "base_price_source_part_no",
            "base_price_source_desc",
            "base_price_source_unit_price",
            "base_price_source_rmc_flag",
            "new_price_source_part_no",
            "new_price_source_desc",
            "new_price_source_unit_price",
            "new_price_source_rmc_flag",
            "bom_qty",
            "production_qty",
            "realized_vi_amount",
            "remaining_vi_amount",
            "expected_vi_amount_usd",
            "expected_vi_calc_status",
            "expected_vi_calc_message",
        ]
    ].copy()

    target_only_mask = model_change_detail["application_status"].eq("target_only")
    missing_base_price_mask = model_change_detail["base_part_no"].fillna("").ne("") & model_change_detail["base_unit_price"].isna()
    missing_new_price_mask = model_change_detail["new_part_no"].fillna("").ne("") & model_change_detail["new_unit_price"].isna()
    missing_bom_qty_mask = model_change_detail["bom_qty"].fillna(0).le(0)
    missing_volume_mask = model_change_detail["production_qty"].isna()
    removed_model_mask = model_change_detail["model_change_status"].eq("removed_model")
    new_part_wrong_context_mask = (
        model_change_detail["base_target_flag"].eq(1)
        & model_change_detail["current_new_part_exists"].eq(0)
        & model_change_detail["new_part_found_anywhere_in_model"].eq("Y")
    )
    missing_model_mask = (
        model_change_detail["application_status"].eq("missing_or_model_changed")
        & ~new_part_wrong_context_mask
    )

    issue_message_series = pd.Series("", index=model_change_detail.index, dtype="object")
    issue_message_series = issue_message_series.mask(target_only_mask, issue_message_series + "Rule does not define base_part_no/new_part_no. ")
    issue_message_series = issue_message_series.mask(missing_base_price_mask, issue_message_series + "Base unit price is missing. ")
    issue_message_series = issue_message_series.mask(missing_new_price_mask, issue_message_series + "New unit price is missing. ")
    issue_message_series = issue_message_series.mask(missing_bom_qty_mask, issue_message_series + "BOM qty is missing. ")
    issue_message_series = issue_message_series.mask(missing_volume_mask, issue_message_series + "Production volume is missing. ")
    issue_message_series = issue_message_series.mask(removed_model_mask, issue_message_series + "Current model is missing entirely. ")
    issue_message_series = issue_message_series.mask(
        new_part_wrong_context_mask,
        issue_message_series + "New P/N exists in model but not under matched ASSY Desc context. ",
    )
    issue_message_series = issue_message_series.mask(missing_model_mask, issue_message_series + "Current BOM has neither base_part_no nor new_part_no. ")
    expected_vi_issue_mask = model_change_detail["expected_vi_calc_status"].isin(
        ["rmc_flag_mismatch", "ancestor_desc_mismatch", "no_rmc_price_source", "missing_price"]
    )
    issue_message_series = issue_message_series.mask(
        expected_vi_issue_mask,
        issue_message_series + model_change_detail["expected_vi_calc_message"].fillna("").astype(str) + " ",
    )
    model_change_detail["issue_message"] = issue_message_series.str.strip()
    model_change_detail["issue_flag"] = model_change_detail["issue_message"].ne("")

    vi_model_detail = model_change_detail[
        [
            "vi_item_id",
            "vi_item_name",
            "rule_type",
            "subsidiary",
            "model_suffix",
            "model_type",
            "indoor_tool",
            "parent_assy_part_no",
            "parent_assy_desc_text",
            "ancestor_desc_path",
            "application_status",
            "base_part_no",
            "new_part_no",
            "base_target_parent_assy_part_no",
            "base_target_parent_assy_desc_text",
            "base_target_ancestor_desc_path",
            "current_new_parent_assy_part_no",
            "current_new_parent_assy_desc_text",
            "current_new_ancestor_desc_path",
            "matched_context_method",
            "new_part_found_under_desc_context",
            "new_part_found_anywhere_in_model",
            "base_unit_price",
            "new_unit_price",
            "base_rmc_flag",
            "new_rmc_flag",
            "base_price_source_part_no",
            "base_price_source_desc",
            "base_price_source_unit_price",
            "base_price_source_rmc_flag",
            "new_price_source_part_no",
            "new_price_source_desc",
            "new_price_source_unit_price",
            "new_price_source_rmc_flag",
            "bom_qty",
            "production_qty",
            "realized_vi_amount",
            "remaining_vi_amount",
            "expected_vi_amount_usd",
            "expected_vi_calc_status",
            "expected_vi_calc_message",
            "issue_flag",
            "issue_message",
        ]
    ].copy()
    vi_model_detail["unit_saving"] = calc_df["unit_saving"].values
    vi_model_detail = vi_model_detail[
        [
            "vi_item_id",
            "vi_item_name",
            "rule_type",
            "subsidiary",
            "model_suffix",
            "model_type",
            "indoor_tool",
            "parent_assy_part_no",
            "parent_assy_desc_text",
            "ancestor_desc_path",
            "application_status",
            "base_part_no",
            "new_part_no",
            "base_target_parent_assy_part_no",
            "base_target_parent_assy_desc_text",
            "base_target_ancestor_desc_path",
            "current_new_parent_assy_part_no",
            "current_new_parent_assy_desc_text",
            "current_new_ancestor_desc_path",
            "matched_context_method",
            "new_part_found_under_desc_context",
            "new_part_found_anywhere_in_model",
            "base_unit_price",
            "new_unit_price",
            "base_rmc_flag",
            "new_rmc_flag",
            "base_price_source_part_no",
            "base_price_source_desc",
            "base_price_source_unit_price",
            "base_price_source_rmc_flag",
            "new_price_source_part_no",
            "new_price_source_desc",
            "new_price_source_unit_price",
            "new_price_source_rmc_flag",
            "unit_saving",
            "bom_qty",
            "production_qty",
            "realized_vi_amount",
            "remaining_vi_amount",
            "expected_vi_amount_usd",
            "expected_vi_calc_status",
            "expected_vi_calc_message",
            "issue_flag",
            "issue_message",
        ]
    ]

    issue_frames = [
        _append_issue_frame(model_change_detail, target_only_mask, "target_only", "Rule has no base_part_no/new_part_no; target model only."),
        _append_issue_frame(model_change_detail, missing_base_price_mask, "missing_price", "Base unit price is missing."),
        _append_issue_frame(model_change_detail, missing_new_price_mask, "missing_price", "New unit price is missing."),
        _append_issue_frame(model_change_detail, missing_bom_qty_mask, "missing_bom_qty", "Could not resolve BOM qty from base/current BOM rows."),
        _append_issue_frame(model_change_detail, missing_volume_mask, "missing_volume", "No volume row matched month/subsidiary/model_suffix."),
        _append_issue_frame(model_change_detail, removed_model_mask, "removed_model", "Current model is missing entirely for this base target model."),
        _append_issue_frame(
            model_change_detail,
            new_part_wrong_context_mask,
            "new_part_context_mismatch",
            "New P/N exists in model but not under matched ASSY Desc context.",
        ),
        _append_issue_frame(model_change_detail, missing_model_mask, "missing_model", "Current BOM has neither base_part_no nor new_part_no for this target model."),
        unsupported_issue_frame,
    ]
    exception_log = pd.concat(issue_frames, ignore_index=True)
    if not exception_log.empty:
        exception_log = exception_log.drop_duplicates().reset_index(drop=True)

    added_without_current_base_mask = (
        model_change_detail["model_change_status"].fillna("").astype(str).str.upper().isin({"ADDED", "ADDED_MODEL"})
        & model_change_detail["current_base_part_exists"].fillna(0).eq(0)
    )
    # Exclude added models without Base P/N in current BOM to reduce noise
    model_change_detail_output = model_change_detail[~added_without_current_base_mask].copy()

    outputs["ModelChangeDetail"] = round_financial_columns(model_change_detail_output)
    outputs["ViModelDetail"] = round_financial_columns(vi_model_detail)
    outputs["ViExceptionLog"] = exception_log
    _log_timing("detail table generation", detail_started_at)
    return outputs


def build_vi_scenario_outputs(
    history_root: Path | None,
    base_year: int,
    base_month: int,
    current_year: int,
    current_month: int,
    scenario_workbook_path: Path,
    item_master_sheet: str | None = None,
    target_rule_sheet: str | None = None,
    volume_sheet: str | None = None,
    base_snapshot_kind: str = "historical",
    base_snapshot_path: Path | None = None,
    current_snapshot_kind: str = "historical",
    current_snapshot_path: Path | None = None,
    detail: bool = False,
    progress_callback: ProgressCallback | None = None,
) -> dict[str, pd.DataFrame]:
    _notify_progress(progress_callback, 0.05, "Loading scenario workbook")
    workbook_started_at = time.perf_counter()
    scenario_inputs = load_scenario_workbook(
        scenario_workbook_path,
        item_master_sheet=item_master_sheet,
        target_rule_sheet=target_rule_sheet,
        volume_sheet=volume_sheet,
    )
    _log_timing("load scenario workbook", workbook_started_at)

    history_root_text = str(default_lake_root(history_root))
    _notify_progress(progress_callback, 0.2, "Loading Base BOM compact snapshot")
    cache_started_at = time.perf_counter()
    base_bom_compact, current_bom_compact = load_or_build_bom_pair_compact_cache(
        history_root_text=history_root_text,
        base_year=base_year,
        base_month=base_month,
        base_snapshot_kind=base_snapshot_kind,
        base_snapshot_path=base_snapshot_path,
        current_year=current_year,
        current_month=current_month,
        current_snapshot_kind=current_snapshot_kind,
        current_snapshot_path=current_snapshot_path,
    )
    _log_timing("load/build BOM pair compact cache", cache_started_at)

    calculation_started_at = time.perf_counter()
    outputs = calculate_vi_outputs(
        base_bom_compact=base_bom_compact,
        current_bom_compact=current_bom_compact,
        vi_item_master=scenario_inputs.vi_item_master,
        vi_target_rule=scenario_inputs.vi_target_rule,
        volume=scenario_inputs.volume,
        current_month=f"{current_year:04d}-{current_month:02d}",
        detail=detail,
        progress_callback=progress_callback,
    )
    _log_timing("calculate_vi_outputs total", calculation_started_at)
    outputs["FinalVITargetRule"] = scenario_inputs.final_vi_target_rule_audit.copy()
    return outputs


def save_vi_scenario_workbook(report: dict[str, pd.DataFrame], output_path: Path) -> None:
    started_at = time.perf_counter()
    save_report_workbook(report, output_path)
    _log_timing("export workbook", started_at)
