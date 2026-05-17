"""Load and normalize VI dashboard workbook data for the Streamlit dashboard."""

from __future__ import annotations

from datetime import datetime
from pathlib import Path
from typing import Any

import pandas as pd

from runtime_paths import application_root, log_root


DASHBOARD_DATA_VERSION = "detail_part_level_v3"

DEFAULT_SUBSIDIARY_OPTIONS = ["LGETA", "LGESP", "LGESR", "LGETH", "LGEKR", "LGEIL"]
DEFAULT_VI_ITEM_OPTIONS = ["VI-001", "VI-002", "VI-003", "VI-004", "VI-005"]

SHEET_CANDIDATES = {
    "item_summary": ["ViItemSummary", "VIItemSummary", "VI_Item_Summary"],
    "subsidiary_summary": ["ViSubsidiarySummary", "VISubsidiarySummary", "VI_Subsidiary_Summary"],
    "detail": ["ModelChangeDetail", "ViModelDetail"],
    "base_target": ["BaseTargetModelList", "BasePNTargetSummary"],
    "model_change_summary": ["ModelChangeSummary"],
}

COLUMN_ALIASES = {
    "viitemid": "vi_item_id",
    "아이템id": "vi_item_id",
    "viitemname": "vi_item_name",
    "아이템명": "vi_item_name",
    "modelsuffix": "model_name",
    "modelname": "model_name",
    "모델명": "model_name",
    "subsidiary": "subsidiary",
    "법인": "subsidiary",
    "parentassypn": "parent_assy_part_no",
    "parentassypartno": "parent_assy_part_no",
    "상위assypn": "parent_assy_part_no",
    "parentassydesc": "parent_assy_desc_text",
    "parentassydesctext": "parent_assy_desc_text",
    "상위assydesc": "parent_assy_desc_text",
    "basepartno": "base_part_no",
    "basepn": "base_part_no",
    "newpartno": "new_part_no",
    "newpn": "new_part_no",
    "변경pn": "new_part_no",
    "applicationstatus": "application_status",
    "적용상태": "application_status",
    "적용상태값": "application_status",
    "targetmodelcount": "target_model_count",
    "대상모델수": "target_model_count",
    "appliedmodelcount": "applied_model_count",
    "적용모델수": "applied_model_count",
    "notappliedmodelcount": "not_applied_model_count",
    "미적용모델수": "not_applied_model_count",
    "realizedviamount": "realized_vi_amount",
    "실현vi금액": "realized_vi_amount",
    "remainingviamount": "remaining_vi_amount",
    "잔여vi금액": "remaining_vi_amount",
    "viamount": "vi_amount",
    "vi금액": "vi_amount",
    "vi금액$": "vi_amount",
    "basetargetmodelcount": "base_target_model_count",
    "base대상모델수": "base_target_model_count",
    "baseappliedmodelcount": "base_applied_model_count",
    "base적용모델수": "base_applied_model_count",
    "basenotappliedmodelcount": "base_not_applied_model_count",
    "base미적용모델수": "base_not_applied_model_count",
    "addedappliedcount": "added_applied_count",
    "추가적용모델수": "added_applied_count",
    "addednotappliedcount": "added_not_applied_count",
    "추가미적용모델수": "added_not_applied_count",
    "addedmixedreviewcount": "added_mixed_review_count",
    "추가혼재검토모델수": "added_mixed_review_count",
}

APPLIED_LABEL = "적용"
NOT_APPLIED_LABEL = "미적용"
BASE_TARGET_STATUSES = {"applied", "not_applied", "mixed_review", "missing_or_model_changed", "target_only"}
CURRENT_TARGET_STATUSES = {"added_applied", "added_not_applied", "added_mixed_review"}

DETAIL_UNIQUENESS_COLUMNS = [
    "subsidiary",
    "model_name",
    "vi_item_id",
    "vi_item_name",
    "parent_assy_part_no",
    "parent_assy_desc_text",
    "base_part_no",
    "new_part_no",
    "application_status",
]

DETAIL_DISPLAY_COLUMNS = {
    "model_name": "모델명",
    "vi_item_name": "VI 아이템",
    "parent_assy_part_no": "상위 ASSY P/N",
    "parent_assy_desc_text": "상위 ASSY Desc",
    "base_part_no": "BASE P/N",
    "new_part_no": "변경 P/N",
    "application_status": "적용 상태",
    "subsidiary": "법인",
    "vi_amount": "VI 금액($)",
}


def log_dashboard_detail_debug(message: str) -> None:
    line = f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {message}"
    print(line, flush=True)
    try:
        root = log_root()
        root.mkdir(parents=True, exist_ok=True)
        with (root / "dashboard_detail_debug.log").open("a", encoding="utf-8") as handle:
            handle.write(line + "\n")
    except OSError:
        pass


def _compact_key(value: object) -> str:
    return "".join(char.lower() for char in str(value).strip() if char.isalnum())


def _ordered_unique(values: list[str]) -> list[str]:
    seen: set[str] = set()
    ordered: list[str] = []
    for value in values:
        item = str(value).strip()
        if not item or item in seen:
            continue
        seen.add(item)
        ordered.append(item)
    return ordered


def _rename_columns(frame: pd.DataFrame) -> pd.DataFrame:
    rename_map: dict[str, str] = {}
    for column in frame.columns:
        alias = COLUMN_ALIASES.get(_compact_key(column))
        if alias:
            rename_map[column] = alias
    return frame.rename(columns=rename_map).copy()


def _normalize_status(value: object) -> str:
    text = str(value).strip().lower()
    if text in {"applied", "added_applied", "적용", "적용완료"}:
        return APPLIED_LABEL
    return NOT_APPLIED_LABEL


def _empty_dashboard_payload() -> dict[str, object]:
    return {
        "detail": pd.DataFrame(
            columns=[
                "vi_item_id",
                "vi_item_name",
                "model_name",
                "parent_assy_part_no",
                "parent_assy_desc_text",
                "base_part_no",
                "new_part_no",
                "application_status",
                "subsidiary",
                "vi_amount",
            ]
        ),
        "item_summary": pd.DataFrame(columns=["vi_item_id", "vi_item_name", "target_model_count", "applied_model_count"]),
        "subsidiary_summary": pd.DataFrame(columns=["subsidiary", "target_model_count", "applied_model_count"]),
        "model_change_summary": pd.DataFrame(
            columns=[
                "vi_item_id",
                "vi_item_name",
                "subsidiary",
                "base_target_model_count",
                "base_applied_model_count",
                "base_not_applied_model_count",
                "added_applied_count",
                "added_not_applied_count",
                "added_mixed_review_count",
            ]
        ),
        "_meta": {
            "detail_source_sheet": "",
            "report_path": "",
            "data_version": DASHBOARD_DATA_VERSION,
        },
    }


def _normalize_item_summary(frame: pd.DataFrame) -> pd.DataFrame:
    data = _rename_columns(frame)
    required = ["vi_item_id", "vi_item_name", "target_model_count", "applied_model_count"]
    for column in required:
        if column not in data.columns:
            data[column] = ""
    data["vi_item_id"] = data["vi_item_id"].fillna("").astype(str).str.strip()
    data["vi_item_name"] = data["vi_item_name"].fillna("").astype(str).str.strip()
    data["target_model_count"] = pd.to_numeric(data["target_model_count"], errors="coerce").fillna(0).astype(int)
    data["applied_model_count"] = pd.to_numeric(data["applied_model_count"], errors="coerce").fillna(0).astype(int)
    return data[required].copy()


def _normalize_subsidiary_summary(frame: pd.DataFrame) -> pd.DataFrame:
    data = _rename_columns(frame)
    required = ["subsidiary", "target_model_count", "applied_model_count"]
    for column in required:
        if column not in data.columns:
            data[column] = ""
    data["subsidiary"] = data["subsidiary"].fillna("").astype(str).str.strip()
    data["target_model_count"] = pd.to_numeric(data["target_model_count"], errors="coerce").fillna(0).astype(int)
    data["applied_model_count"] = pd.to_numeric(data["applied_model_count"], errors="coerce").fillna(0).astype(int)
    return data[required].copy()


def _normalize_model_change_summary(frame: pd.DataFrame) -> pd.DataFrame:
    data = _rename_columns(frame)
    required = [
        "vi_item_id",
        "vi_item_name",
        "subsidiary",
        "base_target_model_count",
        "base_applied_model_count",
        "base_not_applied_model_count",
        "added_applied_count",
        "added_not_applied_count",
        "added_mixed_review_count",
    ]
    for column in required:
        if column not in data.columns:
            data[column] = ""
    data["vi_item_id"] = data["vi_item_id"].fillna("").astype(str).str.strip()
    data["vi_item_name"] = data["vi_item_name"].fillna("").astype(str).str.strip()
    data["subsidiary"] = data["subsidiary"].fillna("").astype(str).str.strip()
    for column in required[3:]:
        data[column] = pd.to_numeric(data[column], errors="coerce").fillna(0).astype(int)
    return data[required].copy()


def _normalize_detail_frame(frame: pd.DataFrame) -> pd.DataFrame:
    data = _rename_columns(frame)
    required = [
        "vi_item_id",
        "vi_item_name",
        "model_name",
        "parent_assy_part_no",
        "parent_assy_desc_text",
        "base_part_no",
        "new_part_no",
        "application_status",
        "raw_application_status",
        "subsidiary",
        "vi_amount",
        "realized_vi_amount",
        "remaining_vi_amount",
    ]
    for column in required:
        if column not in data.columns:
            data[column] = ""

    text_columns = [
        "vi_item_id",
        "vi_item_name",
        "model_name",
        "parent_assy_part_no",
        "parent_assy_desc_text",
        "base_part_no",
        "new_part_no",
        "subsidiary",
    ]
    for column in text_columns:
        data[column] = data[column].fillna("").astype(str).str.strip()

    data["raw_application_status"] = data["application_status"].fillna("").astype(str).str.strip().str.lower()
    data["application_status"] = data["application_status"].map(_normalize_status)
    realized = pd.to_numeric(data["realized_vi_amount"], errors="coerce").fillna(0.0)
    remaining = pd.to_numeric(data["remaining_vi_amount"], errors="coerce").fillna(0.0)
    raw_vi_amount = pd.to_numeric(data["vi_amount"], errors="coerce")
    data["vi_amount"] = raw_vi_amount.fillna(realized.where(realized.ne(0), remaining)).fillna(0.0)

    data.loc[data["vi_item_name"].eq(""), "vi_item_name"] = "미분류 VI 아이템"
    data.loc[data["model_name"].eq(""), "model_name"] = "미분류 모델"
    data.loc[data["subsidiary"].eq(""), "subsidiary"] = "UNMAPPED"
    return data[
        [
            "vi_item_id",
            "vi_item_name",
            "model_name",
            "parent_assy_part_no",
            "parent_assy_desc_text",
            "base_part_no",
            "new_part_no",
            "application_status",
            "raw_application_status",
            "subsidiary",
            "vi_amount",
        ]
    ].copy()


def _normalize_base_target_frame(frame: pd.DataFrame) -> pd.DataFrame:
    data = _rename_columns(frame)
    for column in [
        "vi_item_id",
        "vi_item_name",
        "model_name",
        "parent_assy_part_no",
        "parent_assy_desc_text",
        "base_part_no",
        "new_part_no",
        "subsidiary",
    ]:
        if column not in data.columns:
            data[column] = ""
        data[column] = data[column].fillna("").astype(str).str.strip()

    data["application_status"] = NOT_APPLIED_LABEL
    data["raw_application_status"] = "not_applied"
    data["vi_amount"] = 0.0
    data.loc[data["vi_item_name"].eq(""), "vi_item_name"] = "미분류 VI 아이템"
    data.loc[data["model_name"].eq(""), "model_name"] = "미분류 모델"
    data.loc[data["subsidiary"].eq(""), "subsidiary"] = "UNMAPPED"
    return data[
        [
            "vi_item_id",
            "vi_item_name",
            "model_name",
            "parent_assy_part_no",
            "parent_assy_desc_text",
            "base_part_no",
            "new_part_no",
            "application_status",
            "raw_application_status",
            "subsidiary",
            "vi_amount",
        ]
    ].copy()


def _read_workbook(report_path: Path) -> dict[str, pd.DataFrame]:
    try:
        workbook = pd.ExcelFile(report_path, engine="openpyxl")
        selected_sheet_names: list[str] = []
        for candidate_group in SHEET_CANDIDATES.values():
            for sheet_name in candidate_group:
                if sheet_name in workbook.sheet_names and sheet_name not in selected_sheet_names:
                    selected_sheet_names.append(sheet_name)
                    break
        log_dashboard_detail_debug(
            f"[DashboardDebug] app_root={application_root()} | data_version={DASHBOARD_DATA_VERSION}"
        )
        log_dashboard_detail_debug(f"[DashboardDebug] selected workbook sheets={selected_sheet_names}")
        return {sheet_name: workbook.parse(sheet_name) for sheet_name in selected_sheet_names}
    except Exception as exc:
        log_dashboard_detail_debug(f"[DashboardDebug] failed to read workbook path={report_path} | error={exc}")
        return {}


def _first_sheet_name(workbook: dict[str, pd.DataFrame], candidates: list[str]) -> str:
    for name in candidates:
        if name in workbook:
            return name
    return ""


def _first_sheet(workbook: dict[str, pd.DataFrame], candidates: list[str]) -> pd.DataFrame:
    name = _first_sheet_name(workbook, candidates)
    if not name:
        return pd.DataFrame()
    return workbook[name]


def _build_item_name_map(item_summary: pd.DataFrame) -> dict[str, str]:
    if item_summary.empty:
        return {}
    usable = item_summary[
        item_summary["vi_item_id"].fillna("").astype(str).str.strip().ne("")
        & item_summary["vi_item_name"].fillna("").astype(str).str.strip().ne("")
    ][["vi_item_id", "vi_item_name"]].drop_duplicates()
    return dict(zip(usable["vi_item_id"], usable["vi_item_name"]))


def _fill_missing_item_names(detail: pd.DataFrame, item_summary: pd.DataFrame) -> pd.DataFrame:
    if detail.empty:
        return detail
    item_name_map = _build_item_name_map(item_summary)
    if not item_name_map:
        return detail
    enriched = detail.copy()
    missing_mask = enriched["vi_item_name"].eq("") & enriched["vi_item_id"].ne("")
    enriched.loc[missing_mask, "vi_item_name"] = enriched.loc[missing_mask, "vi_item_id"].map(item_name_map).fillna("")
    enriched.loc[enriched["vi_item_name"].eq(""), "vi_item_name"] = "미분류 VI 아이템"
    return enriched


def _log_column_samples(frame: pd.DataFrame, columns: list[str], prefix: str) -> None:
    for column in columns:
        if column not in frame.columns:
            log_dashboard_detail_debug(f"[DashboardDebug] {prefix} sample {column}=<missing>")
            continue
        values = (
            frame[column]
            .fillna("")
            .astype(str)
            .str.strip()
        )
        samples = [value for value in values.tolist() if value][:3]
        log_dashboard_detail_debug(f"[DashboardDebug] {prefix} sample {column}={samples}")
        for sample in samples:
            if "," in sample:
                log_dashboard_detail_debug(f"[DetailGranularityWarning] column={column} sample={sample}")
                break


def load_dashboard_payload(report_path: str | Path | None = None) -> dict[str, object]:
    if report_path is None:
        log_dashboard_detail_debug("[DashboardDebug] report path is empty")
        return _empty_dashboard_payload()

    path = Path(report_path)
    if not path.exists():
        log_dashboard_detail_debug(f"[DashboardDebug] report path does not exist: {path}")
        return _empty_dashboard_payload()

    workbook = _read_workbook(path)
    if not workbook:
        return _empty_dashboard_payload()

    item_summary = _normalize_item_summary(_first_sheet(workbook, SHEET_CANDIDATES["item_summary"]))
    subsidiary_summary = _normalize_subsidiary_summary(_first_sheet(workbook, SHEET_CANDIDATES["subsidiary_summary"]))
    model_change_summary = _normalize_model_change_summary(_first_sheet(workbook, SHEET_CANDIDATES["model_change_summary"]))

    detail_sheet_name = _first_sheet_name(workbook, SHEET_CANDIDATES["detail"])
    base_target_sheet_name = _first_sheet_name(workbook, SHEET_CANDIDATES["base_target"])
    detail_source_sheet = detail_sheet_name or base_target_sheet_name
    detail_source = workbook.get(detail_source_sheet, pd.DataFrame()) if detail_source_sheet else pd.DataFrame()

    if detail_sheet_name:
        detail = _normalize_detail_frame(detail_source)
    else:
        detail = _normalize_base_target_frame(detail_source)

    detail = _fill_missing_item_names(detail, item_summary)

    log_dashboard_detail_debug(f"[DashboardDebug] report file path={path.resolve()}")
    log_dashboard_detail_debug(f"[DashboardDebug] selected detail source sheet name={detail_source_sheet or '<none>'}")
    log_dashboard_detail_debug(f"[DashboardDebug] normalized detail row count={len(detail.index)}")
    _log_column_samples(detail, ["model_name", "parent_assy_desc_text", "base_part_no", "new_part_no"], "normalized detail")

    return {
        "detail": detail,
        "item_summary": item_summary,
        "subsidiary_summary": subsidiary_summary,
        "model_change_summary": model_change_summary,
        "_meta": {
            "detail_source_sheet": detail_source_sheet,
            "report_path": str(path.resolve()),
            "data_version": DASHBOARD_DATA_VERSION,
        },
    }


def get_filter_options(payload: dict[str, object]) -> dict[str, list[str]]:
    detail = payload["detail"]
    item_summary = payload["item_summary"]
    subsidiary_summary = payload["subsidiary_summary"]
    assert isinstance(detail, pd.DataFrame)
    assert isinstance(item_summary, pd.DataFrame)
    assert isinstance(subsidiary_summary, pd.DataFrame)

    detail_subsidiaries = detail["subsidiary"].dropna().astype(str).tolist() if not detail.empty else []
    summary_subsidiaries = subsidiary_summary["subsidiary"].dropna().astype(str).tolist() if not subsidiary_summary.empty else []
    subsidiaries = DEFAULT_SUBSIDIARY_OPTIONS.copy()
    subsidiaries.extend(
        value for value in _ordered_unique(detail_subsidiaries + summary_subsidiaries) if value not in subsidiaries
    )

    summary_items = item_summary["vi_item_name"].dropna().astype(str).tolist() if not item_summary.empty else []
    detail_items = detail["vi_item_name"].dropna().astype(str).tolist() if not detail.empty else []
    vi_items = _ordered_unique(summary_items + detail_items)
    if not vi_items:
        vi_items = DEFAULT_VI_ITEM_OPTIONS.copy()
    return {
        "subsidiaries": subsidiaries if subsidiaries else DEFAULT_SUBSIDIARY_OPTIONS.copy(),
        "vi_items": vi_items,
    }


def filter_dashboard_payload(
    payload: dict[str, object],
    subsidiaries: list[str] | None = None,
    vi_items: list[str] | None = None,
) -> dict[str, object]:
    detail = payload["detail"].copy()
    item_summary = payload["item_summary"].copy()
    subsidiary_summary = payload["subsidiary_summary"].copy()
    model_change_summary = payload.get("model_change_summary", pd.DataFrame()).copy()
    meta = dict(payload.get("_meta", {}))

    assert isinstance(detail, pd.DataFrame)
    assert isinstance(item_summary, pd.DataFrame)
    assert isinstance(subsidiary_summary, pd.DataFrame)
    assert isinstance(model_change_summary, pd.DataFrame)

    if subsidiaries is not None:
        detail = detail[detail["subsidiary"].isin(subsidiaries)].copy()
        if not subsidiary_summary.empty:
            subsidiary_summary = subsidiary_summary[subsidiary_summary["subsidiary"].isin(subsidiaries)].copy()
        if not model_change_summary.empty:
            model_change_summary = model_change_summary[model_change_summary["subsidiary"].isin(subsidiaries)].copy()

    if vi_items is not None:
        detail = detail[detail["vi_item_name"].isin(vi_items)].copy()
        if not item_summary.empty:
            item_summary = item_summary[item_summary["vi_item_name"].isin(vi_items)].copy()
        if not model_change_summary.empty:
            model_change_summary = model_change_summary[model_change_summary["vi_item_name"].isin(vi_items)].copy()

    return {
        "detail": detail.reset_index(drop=True),
        "item_summary": item_summary.reset_index(drop=True),
        "subsidiary_summary": subsidiary_summary.reset_index(drop=True),
        "model_change_summary": model_change_summary.reset_index(drop=True),
        "_meta": meta,
    }


def _join_unique(series: pd.Series) -> str:
    return ", ".join(_ordered_unique(series.dropna().astype(str).tolist()))


def _aggregate_detail(detail: pd.DataFrame, group_columns: list[str]) -> pd.DataFrame:
    if detail.empty:
        return detail.copy()

    rows: list[dict[str, Any]] = []
    grouped = detail.groupby(group_columns, dropna=False, sort=True)
    for _, frame in grouped:
        working = frame.copy()
        working["vi_amount"] = pd.to_numeric(working["vi_amount"], errors="coerce").fillna(0.0)
        representative = working.sort_values(["vi_amount"], ascending=[False]).iloc[0]
        overall_status = APPLIED_LABEL if working["application_status"].eq(APPLIED_LABEL).all() else NOT_APPLIED_LABEL
        rows.append(
            {
                "vi_item_id": _join_unique(working["vi_item_id"]),
                "vi_item_name": _join_unique(working["vi_item_name"]),
                "model_name": str(representative["model_name"]).strip(),
                "parent_assy_part_no": _join_unique(working["parent_assy_part_no"]),
                "parent_assy_desc_text": _join_unique(working["parent_assy_desc_text"]),
                "base_part_no": _join_unique(working["base_part_no"]),
                "new_part_no": _join_unique(working["new_part_no"]),
                "application_status": overall_status,
                "subsidiary": str(representative["subsidiary"]).strip(),
                "vi_amount": float(working["vi_amount"].sum()),
            }
        )

    return pd.DataFrame(rows)


def deduplicate_detail(detail: pd.DataFrame) -> pd.DataFrame:
    return _aggregate_detail(detail, ["subsidiary", "model_name"])


def build_part_level_detail(detail: pd.DataFrame) -> pd.DataFrame:
    if detail.empty:
        return detail.copy()

    part_level = detail.copy()
    part_level["vi_amount"] = pd.to_numeric(part_level["vi_amount"], errors="coerce").fillna(0.0)
    return part_level.drop_duplicates(subset=DETAIL_UNIQUENESS_COLUMNS, keep="first").reset_index(drop=True)


def build_part_level_table_frame(detail: pd.DataFrame) -> pd.DataFrame:
    if detail.empty:
        return pd.DataFrame(columns=list(DETAIL_DISPLAY_COLUMNS.values()))

    part_level = build_part_level_detail(detail)
    table = part_level[
        [
            "model_name",
            "vi_item_name",
            "parent_assy_part_no",
            "parent_assy_desc_text",
            "base_part_no",
            "new_part_no",
            "application_status",
            "subsidiary",
            "vi_amount",
        ]
    ].rename(columns=DETAIL_DISPLAY_COLUMNS)
    table["VI 금액($)"] = pd.to_numeric(table["VI 금액($)"], errors="coerce").fillna(0.0)
    return table.sort_values(["법인", "모델명", "VI 아이템", "BASE P/N", "변경 P/N"]).reset_index(drop=True)


def build_table_frame(detail: pd.DataFrame) -> pd.DataFrame:
    return build_part_level_table_frame(detail)


def build_kpi_summary(
    item_summary: pd.DataFrame,
    detail: pd.DataFrame,
    model_change_summary: pd.DataFrame | None = None,
) -> dict[str, int]:
    if not detail.empty and "raw_application_status" in detail.columns:
        working = detail.copy()
        working["model_name"] = working["model_name"].fillna("").astype(str).str.strip()
        working["raw_application_status"] = working["raw_application_status"].fillna("").astype(str).str.strip().str.lower()
        working = working[working["model_name"].ne("")].copy()

        if not working.empty:
            model_status = (
                working.groupby("model_name", dropna=False)["raw_application_status"]
                .agg(lambda series: {value for value in series.tolist() if value})
                .reset_index(name="status_set")
            )
            base_target = int(model_status["status_set"].map(lambda values: bool(values & BASE_TARGET_STATUSES)).sum())
            base_applied = int(model_status["status_set"].map(lambda values: "applied" in values).sum())
            base_not_applied = int(model_status["status_set"].map(lambda values: "not_applied" in values).sum())
            current_target = int(model_status["status_set"].map(lambda values: bool(values & CURRENT_TARGET_STATUSES)).sum())
            current_applied = int(model_status["status_set"].map(lambda values: "added_applied" in values).sum())
            current_not_applied = int(model_status["status_set"].map(lambda values: "added_not_applied" in values).sum())
            return {
                "applied_target": base_target + current_target,
                "applied": base_applied + current_applied,
                "not_applied": base_not_applied + current_not_applied,
                "base_target_count": base_target,
                "current_target_count": current_target,
                "base_applied_count": base_applied,
                "current_applied_count": current_applied,
                "base_not_applied_count": base_not_applied,
                "current_not_applied_count": current_not_applied,
            }

    if model_change_summary is not None and not model_change_summary.empty:
        base_target = int(pd.to_numeric(model_change_summary["base_target_model_count"], errors="coerce").fillna(0).sum())
        base_applied = int(pd.to_numeric(model_change_summary["base_applied_model_count"], errors="coerce").fillna(0).sum())
        base_not_applied = int(
            pd.to_numeric(model_change_summary["base_not_applied_model_count"], errors="coerce").fillna(0).sum()
        )
        current_applied = int(pd.to_numeric(model_change_summary["added_applied_count"], errors="coerce").fillna(0).sum())
        current_not_applied = int(
            pd.to_numeric(model_change_summary["added_not_applied_count"], errors="coerce").fillna(0).sum()
        )
        current_mixed = int(
            pd.to_numeric(model_change_summary["added_mixed_review_count"], errors="coerce").fillna(0).sum()
        )
        current_target = current_applied + current_not_applied + current_mixed
        return {
            "applied_target": base_target + current_target,
            "applied": base_applied + current_applied,
            "not_applied": base_not_applied + current_not_applied,
            "base_target_count": base_target,
            "current_target_count": current_target,
            "base_applied_count": base_applied,
            "current_applied_count": current_applied,
            "base_not_applied_count": base_not_applied,
            "current_not_applied_count": current_not_applied,
        }

    if not detail.empty:
        overall = _aggregate_detail(detail, ["model_name"])
        total = int(len(overall.index))
        applied = int(overall["application_status"].eq(APPLIED_LABEL).sum())
        current_compared = int(
            detail["model_name"].fillna("").astype(str).str.strip().replace("", pd.NA).dropna().nunique()
        )
        return {
            "applied_target": total,
            "applied": applied,
            "not_applied": max(total - applied, 0),
            "base_model_count": total,
            "current_model_count": current_compared,
        }

    total = int(pd.to_numeric(item_summary.get("target_model_count", pd.Series(dtype="float64")), errors="coerce").fillna(0).sum())
    applied = int(pd.to_numeric(item_summary.get("applied_model_count", pd.Series(dtype="float64")), errors="coerce").fillna(0).sum())
    return {
        "applied_target": total,
        "applied": applied,
        "not_applied": max(total - applied, 0),
        "base_model_count": total,
        "current_model_count": total,
    }


def build_subsidiary_chart_data(subsidiary_summary: pd.DataFrame, detail: pd.DataFrame) -> pd.DataFrame:
    if not detail.empty:
        deduped = deduplicate_detail(detail)
        grouped = (
            deduped.groupby(["subsidiary", "application_status"], as_index=False)
            .size()
            .pivot(index="subsidiary", columns="application_status", values="size")
            .fillna(0)
            .reset_index()
        )
        if APPLIED_LABEL not in grouped.columns:
            grouped[APPLIED_LABEL] = 0
        if NOT_APPLIED_LABEL not in grouped.columns:
            grouped[NOT_APPLIED_LABEL] = 0
        grouped = grouped.rename(
            columns={
                "subsidiary": "법인",
                APPLIED_LABEL: "적용",
                NOT_APPLIED_LABEL: "미적용",
            }
        )
        grouped["적용 대상"] = grouped["적용"] + grouped["미적용"]
    else:
        grouped = subsidiary_summary.copy()
        if grouped.empty:
            return pd.DataFrame(columns=["법인", "적용 대상", "적용", "미적용", "적용률(%)"])
        grouped["미적용"] = (
            pd.to_numeric(grouped["target_model_count"], errors="coerce").fillna(0)
            - pd.to_numeric(grouped["applied_model_count"], errors="coerce").fillna(0)
        ).clip(lower=0)
        grouped = grouped.rename(
            columns={
                "subsidiary": "법인",
                "target_model_count": "적용 대상",
                "applied_model_count": "적용",
            }
        )[["법인", "적용 대상", "적용", "미적용"]].copy()

    category_order = DEFAULT_SUBSIDIARY_OPTIONS + [
        value for value in grouped["법인"].astype(str).tolist() if value not in DEFAULT_SUBSIDIARY_OPTIONS
    ]
    grouped["법인"] = pd.Categorical(grouped["법인"], categories=category_order, ordered=True)
    grouped = grouped.sort_values("법인").reset_index(drop=True)
    for column in ["적용 대상", "적용", "미적용"]:
        grouped[column] = pd.to_numeric(grouped[column], errors="coerce").fillna(0).astype(int)
    grouped["적용률(%)"] = (grouped["적용"] / grouped["적용 대상"].replace(0, pd.NA) * 100.0).fillna(0.0)
    return grouped
