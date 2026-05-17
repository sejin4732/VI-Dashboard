from __future__ import annotations

import json
import re
import sys
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
from typing import Any

import pandas as pd
from openpyxl.utils import get_column_letter

bootstrap_dir = Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parent
parent_dir = bootstrap_dir.parent
for path in [bootstrap_dir, parent_dir]:
    if str(path) not in sys.path:
        sys.path.insert(0, str(path))

from runtime_paths import application_root, bom_lake_root

try:
    import duckdb
except ImportError as exc:  # pragma: no cover
    duckdb = None
    DUCKDB_IMPORT_ERROR = exc
else:  # pragma: no cover
    DUCKDB_IMPORT_ERROR = None

_DTYPE_DEBUG_PRINTED = False


DEFAULT_LAKE_ROOT_NAME = "bom_lake"
HISTORICAL_DIR_NAME = "historical"
CURRENT_CACHE_DIR_NAME = "current_cache"
METADATA_DIR_NAME = "metadata"
VALIDATION_LOG_FILE_NAME = "validation_log.csv"
PROCESSED_STATUS_FILE_NAME = "processed_months.csv"
DEFAULT_BASELINE_YEAR = 2025

HEADER_SEARCH_START_ROW = 20
HEADER_SEARCH_START_COL = 8
HEADER_REQUIRED_MARKERS = {"modelsuffix", "seq", "partno", "accumqty", "desc"}
PREFERRED_BOM_SHEET_NAMES = [
    "BOM Cost(of all models)(ConsId)",
    "BOM Cost(of all models)(Entity)",
]
BRL_DIVISORS = {
    "USD": 5.6,
    "CNY": 0.789,
    "KRW": 0.004,
}
MANAGE_COLUMN_NAME = "Raw Material Desc."
MANAGE_SUBSIDIARY_COLUMN = "Subsidiary"
MANAGE_REQUIRED_COLUMNS = {MANAGE_COLUMN_NAME, MANAGE_SUBSIDIARY_COLUMN}
MANAGE_HEADER_SCAN_ROWS = 15
SET_MODEL_SHEET_NAME = "SetModelMap"
NEW_MODEL_SHEET_NAME = "NewModelMap"
SET_MODEL_REQUIRED_COLUMNS = {"Model.Suffix", "Set Model Name"}
NEW_MODEL_REQUIRED_COLUMNS = {"New Model", "Base Model", "Start Month"}
SPECIAL_COMPARE_KEYWORDS = [
    "tube,groove",
    "resin",
    "sheet,aluminium",
    "refrigerant",
]
DEFAULT_VALIDATION_COLUMNS = [
    "processed_at",
    "snapshot_kind",
    "bom_year",
    "bom_month",
    "snapshot_month",
    "status",
    "file_name",
    "source_path",
    "sheet_name",
    "row_count",
    "model_count",
    "raw_material_row_count",
    "duplicate_key_count",
    "blank_part_count",
    "blank_model_count",
    "message",
]
BATCH_BOOTSTRAP_LOG_FILE_NAME = "batch_bootstrap_log.csv"
HEADER_ALIASES = {
    "modelsuffix": "modelsuffix",
    "seq": "seq",
    "seq.": "seq",
    "partno": "partno",
    "part no": "partno",
    "part no.": "partno",
    "pn": "partno",
    "accumqty": "accumqty",
    "accum qty": "accumqty",
    "accum. qty": "accumqty",
    "qty": "accumqty",
    "usageqty": "accumqty",
    "unitqty": "accumqty",
    "unit qty": "accumqty",
    "lvl": "lvl",
    "classcode": "classcode",
    "clsscode": "classcode",
    "desc": "desc",
    "desc.": "desc",
    "description": "desc",
    "partdescription": "desc",
    "spec": "spec",
    "specification": "spec",
    "materialspec": "spec",
    "material/spec": "spec",
    "규격": "spec",
    "스펙": "spec",
    "uom": "uom",
    "curr": "curr",
    "unitprice": "unitprice",
    "materialcostloc": "materialcostloc",
    "materialcostusd": "materialcostusd",
    "undfnflag": "undfnflag",
    "rmcflag": "rmcflag",
    "bomapplydate": "bomapplydate",
    "subsidiary": "subsidiary",
    "subsidiaryname": "subsidiary",
    "rawmaterialdesc": "rawmaterialdesc",
    "setmodelname": "setmodelname",
    "newmodel": "newmodel",
    "basemodel": "basemodel",
    "startmonth": "startmonth",
    "sopmonth": "startmonth",
}
COLUMN_NAMES = {
    "subsidiary": "Subsidiary",
    "model_suffix": "Model.Suffix",
    "seq": "seq",
    "level": "Lvl",
    "part_no": "Part No",
    "class_code": "Class Code",
    "quantity": "Accum Qty",
    "description": "Desc.",
    "spec": "Spec",
    "uom": "UOM",
    "currency": "Curr.",
    "unit_price": "Unit Price",
    "material_cost_usd": "Material Cost (USD)",
    "material_cost_loc": "Material cost (LOC)",
    "undfn_flag": "Undfn Flag",
    "rmc_flag": "RMC Flag",
    "bom_apply_date": "BOM Apply Date",
}
VI_COLUMN_ALIASES = {
    "vi item": "vi_item",
    "vi_item": "vi_item",
    "item": "vi_item",
    "change item": "vi_item",
    "model.suffix": "model_suffix",
    "model suffix": "model_suffix",
    "model": "model_suffix",
    "target model": "model_suffix",
    "target_model": "model_suffix",
    "set model": "set_model_name",
    "set model name": "set_model_name",
    "set_model_name": "set_model_name",
    "applied": "applied_flag",
    "applied flag": "applied_flag",
    "apply flag": "applied_flag",
    "status": "applied_flag",
    "saving per model": "saving_per_model",
    "saving_per_model": "saving_per_model",
    "saving": "saving_per_model",
    "current actual saving": "current_actual_saving",
    "current_actual_saving": "current_actual_saving",
    "actual saving": "current_actual_saving",
    "remaining forecast saving": "remaining_forecast_saving",
    "remaining_forecast_saving": "remaining_forecast_saving",
    "forecast saving": "remaining_forecast_saving",
    "opportunity": "total_opportunity",
    "total opportunity": "total_opportunity",
    "total_opportunity": "total_opportunity",
}


def require_duckdb() -> None:
    if duckdb is None:
        raise RuntimeError(
            "duckdb is required for the BOM lake workflow. Install it with: pip install duckdb"
        ) from DUCKDB_IMPORT_ERROR


def script_dir() -> Path:
    return application_root()


def log_progress(message: str) -> None:
    print(f"[Progress] {message}", flush=True)


def default_lake_root(root: Path | None = None) -> Path:
    return Path(root) if root is not None else bom_lake_root()


def historical_root(root: Path | None = None) -> Path:
    return default_lake_root(root) / HISTORICAL_DIR_NAME


def current_cache_root(root: Path | None = None) -> Path:
    return default_lake_root(root) / CURRENT_CACHE_DIR_NAME


def metadata_root(root: Path | None = None) -> Path:
    return default_lake_root(root) / METADATA_DIR_NAME


def validation_log_path(root: Path | None = None) -> Path:
    return metadata_root(root) / VALIDATION_LOG_FILE_NAME


def processed_status_path(root: Path | None = None) -> Path:
    return metadata_root(root) / PROCESSED_STATUS_FILE_NAME


def batch_bootstrap_log_path(root: Path | None = None) -> Path:
    return metadata_root(root) / BATCH_BOOTSTRAP_LOG_FILE_NAME


def ensure_lake_directories(root: Path | None = None) -> Path:
    lake_root = default_lake_root(root)
    historical_root(root).mkdir(parents=True, exist_ok=True)
    current_cache_root(root).mkdir(parents=True, exist_ok=True)
    metadata_root(root).mkdir(parents=True, exist_ok=True)
    return lake_root


def _normalize_text(value: object) -> str:
    text = "" if value is None else str(value)
    text = text.strip().lower()
    text = re.sub(r"\s+", " ", text)
    return text


def _normalize_header(value: object) -> str:
    text = _normalize_text(value)
    text = re.sub(r"\s+", " ", text).strip()
    canonical = text.replace(".", "").replace("_", "").replace(" ", "")
    return HEADER_ALIASES.get(canonical, canonical)


def _normalize_flag(value: object) -> bool:
    return _normalize_text(value) in {"y", "yes", "true", "1", "applied", "done", "complete", "completed"}


def _make_unique_headers(header_values: list[object]) -> list[str]:
    counts: dict[str, int] = {}
    unique_headers: list[str] = []
    for index, value in enumerate(header_values):
        header = str(value).strip() if pd.notna(value) else ""
        if not header:
            header = f"Unnamed_{index}"
        counts[header] = counts.get(header, 0) + 1
        if counts[header] > 1:
            header = f"{header}_{counts[header]}"
        unique_headers.append(header)
    return unique_headers


def _make_unique_column_names(columns: list[object]) -> list[str]:
    counts: dict[str, int] = {}
    unique_columns: list[str] = []
    for index, value in enumerate(columns):
        column = str(value).strip() if value is not None else ""
        if not column:
            column = f"Unnamed_{index}"
        counts[column] = counts.get(column, 0) + 1
        if counts[column] > 1:
            column = f"{column}_{counts[column]}"
        unique_columns.append(column)
    return unique_columns


def _find_header_row(raw_df: pd.DataFrame) -> int:
    for row_idx in range(HEADER_SEARCH_START_ROW - 1, len(raw_df)):
        header_slice = raw_df.iloc[row_idx, HEADER_SEARCH_START_COL - 1 :]
        row_values = {
            _normalize_header(value)
            for value in header_slice.tolist()
            if pd.notna(value) and _normalize_header(value)
        }
        full_row_values = {
            _normalize_header(value)
            for value in raw_df.iloc[row_idx].tolist()
            if pd.notna(value) and _normalize_header(value)
        }
        matched_source_values = row_values if HEADER_REQUIRED_MARKERS.issubset(row_values) else full_row_values
        if HEADER_REQUIRED_MARKERS.issubset(matched_source_values):
            return row_idx
    raise ValueError(
        f"Could not find BOM header row from row {HEADER_SEARCH_START_ROW}. "
        f"Required headers: {', '.join(sorted(HEADER_REQUIRED_MARKERS))}"
    )


def _standardize_columns(data: pd.DataFrame) -> pd.DataFrame:
    rename_map: dict[str, str] = {}
    for column in data.columns:
        base_column = re.sub(r"_\d+$", "", str(column).strip())
        normalized = _normalize_header(base_column)
        if normalized == "modelsuffix":
            rename_map[column] = COLUMN_NAMES["model_suffix"]
        elif normalized == "seq":
            rename_map[column] = COLUMN_NAMES["seq"]
        elif normalized == "partno":
            rename_map[column] = COLUMN_NAMES["part_no"]
        elif normalized == "classcode":
            rename_map[column] = COLUMN_NAMES["class_code"]
        elif normalized == "accumqty":
            rename_map[column] = COLUMN_NAMES["quantity"]
        elif normalized == "desc":
            rename_map[column] = COLUMN_NAMES["description"]
        elif normalized == "spec":
            rename_map[column] = COLUMN_NAMES["spec"]
        elif normalized == "subsidiary":
            rename_map[column] = COLUMN_NAMES["subsidiary"]
        elif normalized == "lvl":
            rename_map[column] = COLUMN_NAMES["level"]
        elif normalized == "uom":
            rename_map[column] = COLUMN_NAMES["uom"]
        elif normalized == "curr":
            rename_map[column] = COLUMN_NAMES["currency"]
        elif normalized == "unitprice":
            rename_map[column] = COLUMN_NAMES["unit_price"]
        elif normalized == "materialcostusd":
            rename_map[column] = COLUMN_NAMES["material_cost_usd"]
        elif normalized == "materialcostloc":
            rename_map[column] = COLUMN_NAMES["material_cost_loc"]
        elif normalized == "undfnflag":
            rename_map[column] = COLUMN_NAMES["undfn_flag"]
        elif normalized == "rmcflag":
            rename_map[column] = COLUMN_NAMES["rmc_flag"]
        elif normalized == "bomapplydate":
            rename_map[column] = COLUMN_NAMES["bom_apply_date"]
    standardized = data.rename(columns=rename_map)
    standardized.columns = _make_unique_column_names(standardized.columns.tolist())
    return standardized


def _select_relevant_bom_columns(data: pd.DataFrame) -> pd.DataFrame:
    relevant_columns = {
        COLUMN_NAMES["subsidiary"],
        COLUMN_NAMES["model_suffix"],
        COLUMN_NAMES["seq"],
        COLUMN_NAMES["level"],
        COLUMN_NAMES["part_no"],
        COLUMN_NAMES["class_code"],
        COLUMN_NAMES["quantity"],
        COLUMN_NAMES["description"],
        COLUMN_NAMES["spec"],
        COLUMN_NAMES["uom"],
        COLUMN_NAMES["currency"],
        COLUMN_NAMES["unit_price"],
        COLUMN_NAMES["material_cost_usd"],
        COLUMN_NAMES["material_cost_loc"],
        COLUMN_NAMES["undfn_flag"],
        COLUMN_NAMES["rmc_flag"],
        COLUMN_NAMES["bom_apply_date"],
    }
    selected = [column for column in data.columns if re.sub(r"_\d+$", "", str(column).strip()) in relevant_columns]
    return data[selected].copy() if selected else data


def _extract_data_from_header_row(raw_df: pd.DataFrame, header_row: int) -> pd.DataFrame:
    header_slice = raw_df.iloc[header_row, HEADER_SEARCH_START_COL - 1 :]
    header_values_in_slice = {
        _normalize_header(value)
        for value in header_slice.tolist()
        if pd.notna(value) and _normalize_header(value)
    }
    start_col = HEADER_SEARCH_START_COL - 1 if HEADER_REQUIRED_MARKERS.issubset(header_values_in_slice) else 0
    sliced = raw_df.iloc[:, start_col:].copy()
    sliced.columns = _make_unique_headers(sliced.iloc[header_row].tolist())
    data = sliced.iloc[header_row + 1 :].copy()
    data = _standardize_columns(data)
    data = _select_relevant_bom_columns(data)
    data = data.dropna(axis=0, how="all")
    return data


def _candidate_excel_engines(excel_engine: str) -> list[str]:
    if excel_engine == "auto":
        return ["calamine", "openpyxl"]
    return [excel_engine]


def _read_excel_frame(
    file_path: Path,
    sheet_name: str | int | None,
    *,
    header: int | None,
    dtype: object | None,
    engine: str,
) -> pd.DataFrame:
    return pd.read_excel(
        file_path,
        sheet_name=sheet_name,
        header=header,
        engine=engine,
        dtype=dtype,
    )


def _get_text_column(frame: pd.DataFrame, column_name: str) -> pd.Series:
    if column_name not in frame.columns:
        return pd.Series([""] * len(frame), index=frame.index, dtype="object")
    return frame[column_name].fillna("").astype(str).str.strip()


def _get_numeric_column(frame: pd.DataFrame, column_name: str) -> pd.Series:
    if column_name not in frame.columns:
        return pd.Series([0.0] * len(frame), index=frame.index, dtype="float64")
    return pd.to_numeric(frame[column_name], errors="coerce").fillna(0.0)


def _get_date_column(frame: pd.DataFrame, column_name: str) -> pd.Series:
    if column_name not in frame.columns:
        return pd.Series([""] * len(frame), index=frame.index, dtype="object")
    parsed = pd.to_datetime(frame[column_name], errors="coerce")
    return parsed.dt.strftime("%Y-%m-%d").fillna("")


def _calculate_material_cost_brl(qty: pd.Series, unit_price: pd.Series, currency: pd.Series) -> pd.Series:
    material_cost = qty * unit_price
    currency_upper = currency.str.upper()
    result = material_cost.copy()
    result = result.where(currency_upper != "USD", material_cost / BRL_DIVISORS["USD"])
    result = result.where(currency_upper != "CNY", material_cost * BRL_DIVISORS["CNY"])
    result = result.where(currency_upper != "KRW", material_cost * BRL_DIVISORS["KRW"])
    return result.fillna(0.0)


def _calculate_fixed_rate_unit_price_brl(unit_price: pd.Series, currency: pd.Series) -> pd.Series:
    currency_upper = currency.str.upper()
    result = unit_price.copy()
    result = result.where(currency_upper != "USD", unit_price / BRL_DIVISORS["USD"])
    result = result.where(currency_upper != "CNY", unit_price * BRL_DIVISORS["CNY"])
    result = result.where(currency_upper != "KRW", unit_price * BRL_DIVISORS["KRW"])
    return result.fillna(0.0)


def load_multibom_sheet(
    file_path: Path,
    preferred_sheet_name: str | None = None,
    excel_engine: str = "auto",
) -> tuple[str, pd.DataFrame, str, float]:
    last_error: Exception | None = None
    for engine in _candidate_excel_engines(excel_engine):
        started_at = time.perf_counter()
        try:
            excel_file = pd.ExcelFile(file_path, engine=engine)
            target_sheet_names: list[str] = []
            if preferred_sheet_name:
                if preferred_sheet_name not in excel_file.sheet_names:
                    raise ValueError(f"Sheet '{preferred_sheet_name}' was not found in {file_path.name}.")
                target_sheet_names.append(preferred_sheet_name)
            target_sheet_names.extend(
                sheet_name
                for sheet_name in PREFERRED_BOM_SHEET_NAMES
                if sheet_name in excel_file.sheet_names and sheet_name not in target_sheet_names
            )
            if not target_sheet_names:
                raise ValueError(
                    f"{file_path.name} must contain one of these sheets: {', '.join(PREFERRED_BOM_SHEET_NAMES)}"
                )

            for index, sheet_name in enumerate(target_sheet_names, start=1):
                log_progress(
                    f"Reading file: {file_path.name} | engine={engine} | scanning sheet {index}/{len(target_sheet_names)}: {sheet_name}"
                )
                raw_df = _read_excel_frame(file_path, sheet_name, header=None, dtype=None, engine=engine)
                if raw_df.empty:
                    continue
                try:
                    header_row = _find_header_row(raw_df)
                    data = _extract_data_from_header_row(raw_df, header_row)
                    required = [
                        COLUMN_NAMES["model_suffix"],
                        COLUMN_NAMES["seq"],
                        COLUMN_NAMES["part_no"],
                        COLUMN_NAMES["quantity"],
                        COLUMN_NAMES["description"],
                    ]
                    if all(column in data.columns for column in required):
                        log_progress(f"Selected sheet: {sheet_name}")
                        return sheet_name, data, engine, time.perf_counter() - started_at
                except ValueError:
                    continue
            raise ValueError(f"Could not find a valid Multi BOM sheet in {file_path.name}.")
        except Exception as exc:
            last_error = exc
            if excel_engine != "auto":
                raise
            continue
    if last_error is not None:
        raise last_error
    raise ValueError(f"Could not read {file_path.name}.")


def _standardize_manage_columns(data: pd.DataFrame) -> pd.DataFrame:
    rename_map: dict[str, str] = {}
    for column in data.columns:
        normalized = _normalize_header(column)
        if normalized == "subsidiary":
            rename_map[column] = MANAGE_SUBSIDIARY_COLUMN
        elif normalized == "rawmaterialdesc":
            rename_map[column] = MANAGE_COLUMN_NAME
    return data.rename(columns=rename_map)


def _load_manage_sheet(manage_path: Path, manage_sheet_name: str) -> pd.DataFrame:
    raw_df = pd.read_excel(manage_path, sheet_name=manage_sheet_name, header=None, engine="openpyxl", dtype=object)
    for header_row in range(min(MANAGE_HEADER_SCAN_ROWS, len(raw_df))):
        candidate = raw_df.iloc[header_row + 1 :].copy()
        candidate.columns = _make_unique_headers(raw_df.iloc[header_row].tolist())
        candidate = _standardize_manage_columns(candidate).dropna(axis=0, how="all").dropna(axis=1, how="all")
        if MANAGE_REQUIRED_COLUMNS.issubset(set(candidate.columns)):
            return candidate
    data = pd.read_excel(manage_path, sheet_name=manage_sheet_name, engine="openpyxl", dtype=object)
    data = _standardize_manage_columns(data).dropna(axis=0, how="all").dropna(axis=1, how="all")
    return data


def load_manage_descriptions(
    manage_path: Path | None,
    manage_sheet_name: str | None,
    subsidiary: str | None,
) -> set[str]:
    if manage_path is None or not manage_path.exists() or not manage_sheet_name or not subsidiary:
        return set()
    manage_df = _load_manage_sheet(manage_path, manage_sheet_name)
    if not MANAGE_REQUIRED_COLUMNS.issubset(set(manage_df.columns)):
        return set()
    filtered = manage_df[
        manage_df[MANAGE_SUBSIDIARY_COLUMN].map(_normalize_text) == _normalize_text(subsidiary)
    ]
    return {
        _normalize_text(value)
        for value in filtered[MANAGE_COLUMN_NAME].dropna().tolist()
        if _normalize_text(value)
    }


def _standardize_set_model_columns(data: pd.DataFrame) -> pd.DataFrame:
    rename_map: dict[str, str] = {}
    for column in data.columns:
        normalized = _normalize_header(column)
        if normalized == "modelsuffix":
            rename_map[column] = "Model.Suffix"
        elif normalized == "setmodelname":
            rename_map[column] = "Set Model Name"
    return data.rename(columns=rename_map)


def _standardize_new_model_columns(data: pd.DataFrame) -> pd.DataFrame:
    rename_map: dict[str, str] = {}
    for column in data.columns:
        normalized = _normalize_header(column)
        if normalized == "newmodel":
            rename_map[column] = "New Model"
        elif normalized == "basemodel":
            rename_map[column] = "Base Model"
        elif normalized == "startmonth":
            rename_map[column] = "Start Month"
    return data.rename(columns=rename_map)


def build_set_model_sheet_frame(data: pd.DataFrame | None = None) -> pd.DataFrame:
    if data is not None and not data.empty:
        return data
    return pd.DataFrame(
        {
            "Model.Suffix": [""],
            "Set Model Name": [""],
            "Note": ["Leave blank to use default rule: replace the 3rd character with '-'."],
        }
    )


def build_new_model_sheet_frame(data: pd.DataFrame | None = None) -> pd.DataFrame:
    if data is not None and not data.empty:
        return data
    return pd.DataFrame(
        {
            "New Model": [""],
            "Base Model": [""],
            "Start Month": [""],
            "Note": ["Fill when a model is new and needs separate monthly tracking."],
        }
    )


def load_optional_management_sheets(
    manage_path: Path | None,
) -> tuple[dict[str, str], dict[str, dict[str, str]], pd.DataFrame, pd.DataFrame]:
    empty_set_sheet = build_set_model_sheet_frame()
    empty_new_sheet = build_new_model_sheet_frame()
    if manage_path is None or not manage_path.exists():
        return {}, {}, empty_set_sheet, empty_new_sheet

    try:
        excel_file = pd.ExcelFile(manage_path, engine="openpyxl")
    except Exception:
        return {}, {}, empty_set_sheet, empty_new_sheet

    set_model_map: dict[str, str] = {}
    new_model_map: dict[str, dict[str, str]] = {}
    set_sheet_frame = empty_set_sheet
    new_sheet_frame = empty_new_sheet

    if SET_MODEL_SHEET_NAME in excel_file.sheet_names:
        set_sheet = pd.read_excel(manage_path, sheet_name=SET_MODEL_SHEET_NAME, engine="openpyxl", dtype=object)
        set_sheet = _standardize_set_model_columns(set_sheet).dropna(axis=0, how="all").dropna(axis=1, how="all")
        set_sheet_frame = build_set_model_sheet_frame(set_sheet)
        if SET_MODEL_REQUIRED_COLUMNS.issubset(set(set_sheet.columns)):
            for row in set_sheet.to_dict("records"):
                model_name = str(row.get("Model.Suffix", "")).strip()
                set_model_name = str(row.get("Set Model Name", "")).strip()
                if model_name and set_model_name:
                    set_model_map[model_name] = set_model_name

    if NEW_MODEL_SHEET_NAME in excel_file.sheet_names:
        new_sheet = pd.read_excel(manage_path, sheet_name=NEW_MODEL_SHEET_NAME, engine="openpyxl", dtype=object)
        new_sheet = _standardize_new_model_columns(new_sheet).dropna(axis=0, how="all").dropna(axis=1, how="all")
        new_sheet_frame = build_new_model_sheet_frame(new_sheet)
        if NEW_MODEL_REQUIRED_COLUMNS.issubset(set(new_sheet.columns)):
            for row in new_sheet.to_dict("records"):
                new_model = str(row.get("New Model", "")).strip()
                if new_model:
                    new_model_map[new_model] = {
                        "base_model": str(row.get("Base Model", "")).strip(),
                        "start_month": str(row.get("Start Month", "")).strip(),
                    }

    return set_model_map, new_model_map, set_sheet_frame, new_sheet_frame


def build_manage_sheet_frame(raw_material_descriptions: set[str], subsidiary: str) -> pd.DataFrame:
    descriptions = sorted(raw_material_descriptions) or [""]
    return pd.DataFrame(
        {
            MANAGE_SUBSIDIARY_COLUMN: [subsidiary] + [""] * (len(descriptions) - 1),
            MANAGE_COLUMN_NAME: descriptions,
            "Note": ["Reference list used to classify raw material rows."] + [""] * (len(descriptions) - 1),
        }
    )


def default_set_model_name(model_name: str) -> str:
    if len(model_name) < 3:
        return model_name
    return f"{model_name[:2]}-{model_name[3:]}"


def resolve_set_model_name(model_name: str, set_model_map: dict[str, str]) -> str:
    if not model_name:
        return ""
    return set_model_map.get(model_name, default_set_model_name(model_name))


def resolve_new_model_note(model_name: str, new_model_map: dict[str, dict[str, str]]) -> str:
    info = new_model_map.get(model_name)
    if not info:
        return ""
    start_month = str(info.get("start_month", "")).strip()
    return f"New model from {start_month}" if start_month else "New model"


def _parse_level_number(value: object) -> int | None:
    text = str(value).strip()
    if not text or text.lower() == "nan":
        return None
    digit_match = re.search(r"(\d+)\s*$", text)
    if digit_match:
        return int(digit_match.group(1))
    dot_count = text.count(".")
    if dot_count > 0:
        return dot_count
    try:
        return int(float(text))
    except Exception:
        return None


def prepare_bom(
    frame: pd.DataFrame,
    source_label: str,
    raw_material_descriptions: set[str],
    include_brl_costs: bool = True,
) -> pd.DataFrame:
    model_suffix = _get_text_column(frame, COLUMN_NAMES["model_suffix"]).replace("nan", "", regex=False)
    model_suffix = model_suffix.mask(model_suffix == "").ffill().fillna("")
    part_no = _get_text_column(frame, COLUMN_NAMES["part_no"])
    prepared = pd.DataFrame(
        {
            "subsidiary": _get_text_column(frame, COLUMN_NAMES["subsidiary"]),
            "model_suffix": model_suffix,
            "seq": _get_numeric_column(frame, COLUMN_NAMES["seq"]),
            "level": _get_text_column(frame, COLUMN_NAMES["level"]),
            "part_no": part_no,
            "class_code": part_no.str[:8],
            "description": _get_text_column(frame, COLUMN_NAMES["description"]),
            "spec_text": _get_text_column(frame, COLUMN_NAMES["spec"]),
            "uom": _get_text_column(frame, COLUMN_NAMES["uom"]),
            "qty": _get_numeric_column(frame, COLUMN_NAMES["quantity"]),
            "currency": _get_text_column(frame, COLUMN_NAMES["currency"]),
            "unit_price": _get_numeric_column(frame, COLUMN_NAMES["unit_price"]),
            "material_cost_usd": _get_numeric_column(frame, COLUMN_NAMES["material_cost_usd"]),
            "material_cost_loc": _get_numeric_column(frame, COLUMN_NAMES["material_cost_loc"]),
            "undfn_flag": _get_text_column(frame, COLUMN_NAMES["undfn_flag"]),
            "rmc_flag": _get_text_column(frame, COLUMN_NAMES["rmc_flag"]),
            "bom_apply_date": _get_date_column(frame, COLUMN_NAMES["bom_apply_date"]),
            "source": source_label,
        }
    )
    prepared["row_order"] = range(len(prepared))
    prepared["level_num"] = prepared["level"].map(_parse_level_number)
    prepared["desc_text_upper"] = prepared["description"].fillna("").astype(str).str.upper()
    prepared["spec_text_upper"] = prepared["spec_text"].fillna("").astype(str).str.upper()
    # These BRL cost columns are derived workflow fields, not original BOM columns.
    if include_brl_costs:
        prepared["fixed_rate_unit_price_brl"] = _calculate_fixed_rate_unit_price_brl(
            prepared["unit_price"], prepared["currency"]
        )
        prepared["material_cost_brl"] = _calculate_material_cost_brl(
            prepared["qty"], prepared["unit_price"], prepared["currency"]
        )
        prepared["fixed_rate_material_cost_brl"] = prepared["qty"] * prepared["fixed_rate_unit_price_brl"]
    else:
        prepared["fixed_rate_unit_price_brl"] = 0.0
        prepared["material_cost_brl"] = 0.0
        prepared["fixed_rate_material_cost_brl"] = 0.0
    prepared["description_key"] = prepared["description"].map(_normalize_text)
    prepared["raw_material_flag"] = (
        prepared["description_key"].isin(raw_material_descriptions) | prepared["rmc_flag"].str.upper().eq("Y")
    )
    prepared["is_raw_material"] = prepared["raw_material_flag"]
    prepared = prepared[
        (prepared["model_suffix"] != "")
        & (prepared["part_no"] != "")
        & (prepared["part_no"].str.lower() != "nan")
    ].copy()
    return prepared.sort_values(["model_suffix", "seq", "row_order", "part_no"]).reset_index(drop=True)


def infer_year_month(
    file_path: Path,
    frame: pd.DataFrame | None = None,
    bom_year: int | None = None,
    bom_month: int | None = None,
) -> tuple[int, int]:
    inferred_year: int | None = int(bom_year) if bom_year is not None else None
    inferred_month: int | None = int(bom_month) if bom_month is not None else None

    if frame is not None and COLUMN_NAMES["bom_apply_date"] in frame.columns:
        parsed = pd.to_datetime(frame[COLUMN_NAMES["bom_apply_date"]], errors="coerce").dropna()
        if not parsed.empty:
            first_date = parsed.iloc[0]
            if inferred_year is None:
                inferred_year = int(first_date.year)
            if inferred_month is None:
                inferred_month = int(first_date.month)

    if inferred_month is None:
        korean_month_match = re.search(r"(?<!\d)(1[0-2]|[1-9])\s*월", file_path.stem)
        if korean_month_match:
            inferred_month = int(korean_month_match.group(1))

    match = re.search(r"((?:20)?\d{2})[.\-_/ ]+(\d{1,2})", file_path.stem)
    if match:
        year_value = int(match.group(1))
        if year_value < 100:
            year_value += 2000
        if inferred_year is None:
            inferred_year = year_value
        if inferred_month is None:
            inferred_month = int(match.group(2))

    if inferred_year is not None and inferred_month is not None:
        return inferred_year, inferred_month
    raise ValueError(f"Could not infer year/month from {file_path.name}.")


def _parquet_snapshot_path(
    root: Path | None,
    snapshot_kind: str,
    bom_year: int,
    bom_month: int,
    file_stem: str,
) -> Path:
    safe_stem = re.sub(r"[^A-Za-z0-9._-]+", "_", file_stem).strip("_") or "snapshot"
    base = historical_root(root) if snapshot_kind == "historical" else current_cache_root(root)
    return base / f"bom_year={bom_year}" / f"bom_month={bom_month:02d}" / f"{safe_stem}.parquet"


def _safe_snapshot_stem(value: str | None, fallback: str = "snapshot") -> str:
    text = "" if value is None else str(value).strip()
    safe_stem = re.sub(r"[^A-Za-z0-9._-]+", "_", text).strip("_")
    return safe_stem or fallback


def _write_parquet(frame: pd.DataFrame, parquet_path: Path) -> None:
    require_duckdb()
    parquet_path.parent.mkdir(parents=True, exist_ok=True)
    if parquet_path.exists():
        parquet_path.unlink()
    con = duckdb.connect()
    try:
        con.register("snapshot_df", frame)
        con.execute("COPY snapshot_df TO ? (FORMAT PARQUET)", [str(parquet_path)])
    finally:
        con.close()


def _delete_existing_month_parquets(root: Path | None, snapshot_kind: str, bom_year: int, bom_month: int) -> None:
    base = historical_root(root) if snapshot_kind == "historical" else current_cache_root(root)
    month_dir = base / f"bom_year={bom_year}" / f"bom_month={bom_month:02d}"
    if not month_dir.exists():
        return
    for parquet_file in month_dir.glob("*.parquet"):
        parquet_file.unlink()


def _read_csv_metadata(path: Path, columns: list[str]) -> pd.DataFrame:
    if not path.exists():
        return pd.DataFrame(columns=columns)
    data = pd.read_csv(path, dtype=object).fillna("")
    for column in columns:
        if column not in data.columns:
            data[column] = ""
    return data[columns]


def _normalize_year_month_dtypes(data: pd.DataFrame) -> pd.DataFrame:
    global _DTYPE_DEBUG_PRINTED
    if data.empty:
        return data

    normalized = data.copy()
    if "bom_year" in normalized.columns:
        normalized["bom_year"] = pd.to_numeric(normalized["bom_year"], errors="coerce").astype("Int64")
    if "bom_month" in normalized.columns:
        normalized["bom_month"] = pd.to_numeric(normalized["bom_month"], errors="coerce").astype("Int64")
    if "snapshot_month" in normalized.columns:
        normalized["snapshot_month"] = normalized["snapshot_month"].fillna("").astype(str)

    year_dtype = normalized["bom_year"].dtype if "bom_year" in normalized.columns else "N/A"
    month_dtype = normalized["bom_month"].dtype if "bom_month" in normalized.columns else "N/A"
    if not _DTYPE_DEBUG_PRINTED:
        print(f"[Debug] [bom_year dtype={year_dtype}, bom_month dtype={month_dtype}]", flush=True)
        _DTYPE_DEBUG_PRINTED = True
    return normalized


def _replace_metadata_rows(
    path: Path,
    frame: pd.DataFrame,
    key_columns: list[str],
    row: dict[str, Any],
    columns: list[str],
) -> None:
    row_frame = pd.DataFrame([{column: row.get(column, "") for column in columns}])
    if not frame.empty:
        mask = pd.Series(True, index=frame.index)
        for key in key_columns:
            mask &= frame[key].astype(str) == str(row.get(key, ""))
        frame = frame[~mask].copy()
    updated = pd.concat([frame, row_frame], ignore_index=True)
    updated = updated.sort_values(key_columns).reset_index(drop=True)
    path.parent.mkdir(parents=True, exist_ok=True)
    updated.to_csv(path, index=False, encoding="utf-8-sig")


def _append_validation_record(root: Path | None, record: dict[str, Any]) -> None:
    path = validation_log_path(root)
    history = _read_csv_metadata(path, DEFAULT_VALIDATION_COLUMNS)
    _replace_metadata_rows(
        path,
        history,
        ["snapshot_kind", "bom_year", "bom_month", "file_name"],
        record,
        DEFAULT_VALIDATION_COLUMNS,
    )


def _upsert_processed_month(root: Path | None, record: dict[str, Any]) -> None:
    columns = [
        "bom_year",
        "bom_month",
        "snapshot_month",
        "status",
        "file_name",
        "sheet_name",
        "parquet_path",
        "processed_at",
        "row_count",
        "model_count",
        "raw_material_row_count",
        "message",
    ]
    path = processed_status_path(root)
    history = _read_csv_metadata(path, columns)
    _replace_metadata_rows(path, history, ["bom_year", "bom_month"], record, columns)


def process_snapshot(
    file_path: Path,
    root: Path | None = None,
    manage_path: Path | None = None,
    manage_sheet_name: str | None = None,
    subsidiary: str | None = None,
    preferred_sheet_name: str | None = None,
    bom_year: int | None = None,
    bom_month: int | None = None,
    snapshot_kind: str = "historical",
    excel_engine: str = "auto",
    include_brl_costs: bool = True,
    update_metadata: bool = True,
    overwrite_existing_month: bool = False,
    snapshot_name: str | None = None,
) -> dict[str, Any]:
    started_at = time.perf_counter()
    ensure_lake_directories(root)
    raw_material_descriptions = load_manage_descriptions(manage_path, manage_sheet_name, subsidiary)
    sheet_name, raw_df, engine_used, read_excel_seconds = load_multibom_sheet(
        file_path,
        preferred_sheet_name,
        excel_engine=excel_engine,
    )
    prepare_started_at = time.perf_counter()
    prepared = prepare_bom(raw_df, snapshot_kind, raw_material_descriptions, include_brl_costs=include_brl_costs)
    prepare_bom_seconds = time.perf_counter() - prepare_started_at
    if prepared.empty or int(len(prepared.index)) == 0:
        raise ValueError("No BOM data rows extracted. Please refresh/save the Excel file first.")
    resolved_year, resolved_month = infer_year_month(file_path, raw_df, bom_year=bom_year, bom_month=bom_month)
    snapshot_month = f"{resolved_year}-{resolved_month:02d}"
    parquet_stem = _safe_snapshot_stem(snapshot_name, file_path.stem)
    parquet_path = _parquet_snapshot_path(root, snapshot_kind, resolved_year, resolved_month, parquet_stem)

    snapshot_frame = prepared.copy()
    snapshot_frame["bom_year"] = int(resolved_year)
    snapshot_frame["bom_month"] = int(resolved_month)
    snapshot_frame["snapshot_month"] = snapshot_month
    snapshot_frame["snapshot_kind"] = snapshot_kind
    snapshot_frame["snapshot_source_file"] = file_path.name
    snapshot_frame["snapshot_source_path"] = str(file_path.resolve())
    snapshot_frame["snapshot_sheet_name"] = sheet_name
    snapshot_frame["snapshot_name"] = parquet_stem
    snapshot_frame["snapshot_loaded_at"] = pd.Timestamp.utcnow().isoformat()
    if overwrite_existing_month:
        _delete_existing_month_parquets(root, snapshot_kind, resolved_year, resolved_month)
    write_started_at = time.perf_counter()
    _write_parquet(snapshot_frame, parquet_path)
    write_parquet_seconds = time.perf_counter() - write_started_at
    total_seconds = time.perf_counter() - started_at
    log_progress(
        f"Loaded {file_path.name} | engine={engine_used} | rows={len(snapshot_frame.index):,} | models={snapshot_frame['model_suffix'].nunique():,} | parquet={parquet_path}"
    )
    log_progress(
        "Timing "
        f"{file_path.name} | read_excel={read_excel_seconds:.2f}s | prepare_bom={prepare_bom_seconds:.2f}s | "
        f"write_parquet={write_parquet_seconds:.2f}s | total={total_seconds:.2f}s"
    )

    validation_record = {
        "processed_at": pd.Timestamp.now().isoformat(timespec="seconds"),
        "snapshot_kind": snapshot_kind,
        "bom_year": int(resolved_year),
        "bom_month": f"{resolved_month:02d}",
        "snapshot_month": snapshot_month,
        "status": "processed",
        "file_name": parquet_stem if snapshot_kind == "current_cache" else file_path.name,
        "source_path": str(file_path.resolve()),
        "sheet_name": sheet_name,
        "row_count": int(len(snapshot_frame.index)),
        "model_count": int(snapshot_frame["model_suffix"].nunique()),
        "raw_material_row_count": int(snapshot_frame["is_raw_material"].fillna(False).sum()),
        "duplicate_key_count": int(
            snapshot_frame.duplicated(subset=["model_suffix", "seq", "part_no"], keep=False).sum()
        ),
        "blank_part_count": int(snapshot_frame["part_no"].fillna("").astype(str).str.strip().eq("").sum()),
        "blank_model_count": int(snapshot_frame["model_suffix"].fillna("").astype(str).str.strip().eq("").sum()),
        "message": "",
    }
    processed_record = {
        "bom_year": int(resolved_year),
        "bom_month": f"{resolved_month:02d}",
        "snapshot_month": snapshot_month,
        "status": "processed",
        "file_name": parquet_stem if snapshot_kind == "current_cache" else file_path.name,
        "sheet_name": sheet_name,
        "parquet_path": str(parquet_path),
        "processed_at": validation_record["processed_at"],
        "row_count": validation_record["row_count"],
        "model_count": validation_record["model_count"],
        "raw_material_row_count": validation_record["raw_material_row_count"],
        "message": "",
    }
    if update_metadata:
        _append_validation_record(root, validation_record)
        if snapshot_kind == "historical":
            _upsert_processed_month(root, processed_record)

    return {
        "sheet_name": sheet_name,
        "engine_used": engine_used,
        "parquet_path": parquet_path,
        "bom_year": int(resolved_year),
        "bom_month": int(resolved_month),
        "snapshot_month": snapshot_month,
        "data": snapshot_frame,
        "validation": validation_record,
        "processed_record": processed_record,
        "timing": {
            "read_excel_seconds": read_excel_seconds,
            "prepare_bom_seconds": prepare_bom_seconds,
            "write_parquet_seconds": write_parquet_seconds,
            "total_seconds": total_seconds,
        },
    }


def bootstrap_historical_year(
    input_dir: Path,
    root: Path | None = None,
    manage_path: Path | None = None,
    manage_sheet_name: str | None = None,
    subsidiary: str | None = None,
    bom_year: int = DEFAULT_BASELINE_YEAR,
    skip_existing: bool = True,
    excel_engine: str = "auto",
    include_brl_costs: bool = True,
    workers: int = 1,
    rebuild_parquet: bool = False,
) -> dict[str, pd.DataFrame]:
    input_files = sorted(path for path in input_dir.iterdir() if path.suffix.lower() in {".xlsx", ".xlsm", ".xls"})
    if not input_files:
        raise ValueError(f"No Excel files were found in {input_dir}")

    processed_rows: list[dict[str, Any]] = []
    validation_rows: list[dict[str, Any]] = []
    batch_bootstrap_rows: list[dict[str, Any]] = []

    existing_months: set[tuple[int, int]] = set()
    existing_status = list_processed_months(root)
    if not existing_status.empty:
        existing_months = {
            (int(row["bom_year"]), int(row["bom_month"]))
            for row in existing_status.to_dict("records")
            if str(row.get("bom_year", "")).strip() and str(row.get("bom_month", "")).strip()
        }

    should_skip_existing = skip_existing and not rebuild_parquet
    worker_count = max(1, min(int(workers), 3))
    scheduled_files: list[tuple[int, Path, int, int | None]] = []

    for index, file_path in enumerate(input_files, start=1):
        try:
            inferred_year, inferred_month = infer_year_month(file_path, bom_year=bom_year)
        except Exception:
            inferred_year, inferred_month = int(bom_year), None

        if should_skip_existing and inferred_month is not None and (inferred_year, inferred_month) in existing_months:
            log_progress(
                f"[{index}/{len(input_files)}] Skipping already processed month: {inferred_year}-{inferred_month:02d} | {file_path.name}"
            )
            batch_bootstrap_rows.append(
                {
                    "file_name": file_path.name,
                    "bom_year": inferred_year,
                    "bom_month": f"{inferred_month:02d}",
                    "status": "skipped_existing",
                    "row_count": "",
                    "model_count": "",
                    "parquet_path": "",
                    "engine_used": "",
                    "read_excel_seconds": "",
                    "prepare_bom_seconds": "",
                    "write_parquet_seconds": "",
                    "total_seconds": "",
                    "error_message": "",
                    "processed_at": pd.Timestamp.now().isoformat(timespec="seconds"),
                }
            )
            continue
        scheduled_files.append((index, file_path, inferred_year, inferred_month))

    def run_single_file(file_path: Path) -> dict[str, Any]:
        return process_snapshot(
            file_path=file_path,
            root=root,
            manage_path=manage_path,
            manage_sheet_name=manage_sheet_name,
            subsidiary=subsidiary,
            bom_year=bom_year,
            snapshot_kind="historical",
            excel_engine=excel_engine,
            include_brl_costs=include_brl_costs,
            update_metadata=False,
            overwrite_existing_month=True,
        )

    if worker_count == 1:
        for index, file_path, inferred_year, inferred_month in scheduled_files:
            log_progress(f"[{index}/{len(input_files)}] Bootstrapping {file_path.name}")
            try:
                result = run_single_file(file_path)
                _append_validation_record(root, result["validation"])
                _upsert_processed_month(root, result["processed_record"])
                processed_rows.append(
                    {
                        "bom_year": result["bom_year"],
                        "bom_month": f"{result['bom_month']:02d}",
                        "snapshot_month": result["snapshot_month"],
                        "file_name": file_path.name,
                        "sheet_name": result["sheet_name"],
                        "parquet_path": str(result["parquet_path"]),
                        "row_count": int(len(result["data"].index)),
                        "model_count": int(result["data"]["model_suffix"].nunique()),
                        "raw_material_row_count": int(result["data"]["is_raw_material"].fillna(False).sum()),
                    }
                )
                validation_rows.append(result["validation"])
                existing_months.add((int(result["bom_year"]), int(result["bom_month"])))
                batch_bootstrap_rows.append(
                    {
                        "file_name": file_path.name,
                        "bom_year": result["bom_year"],
                        "bom_month": f"{result['bom_month']:02d}",
                        "status": "processed",
                        "row_count": int(len(result["data"].index)),
                        "model_count": int(result["data"]["model_suffix"].nunique()),
                        "parquet_path": str(result["parquet_path"]),
                        "engine_used": result["engine_used"],
                        "read_excel_seconds": round(result["timing"]["read_excel_seconds"], 2),
                        "prepare_bom_seconds": round(result["timing"]["prepare_bom_seconds"], 2),
                        "write_parquet_seconds": round(result["timing"]["write_parquet_seconds"], 2),
                        "total_seconds": round(result["timing"]["total_seconds"], 2),
                        "error_message": "",
                        "processed_at": pd.Timestamp.now().isoformat(timespec="seconds"),
                    }
                )
            except Exception as exc:
                log_progress(f"Failed {file_path.name}: {exc}")
                batch_bootstrap_rows.append(
                    {
                        "file_name": file_path.name,
                        "bom_year": inferred_year,
                        "bom_month": f"{inferred_month:02d}" if inferred_month is not None else "",
                        "status": "failed",
                        "row_count": "",
                        "model_count": "",
                        "parquet_path": "",
                        "engine_used": "",
                        "read_excel_seconds": "",
                        "prepare_bom_seconds": "",
                        "write_parquet_seconds": "",
                        "total_seconds": "",
                        "error_message": str(exc),
                        "processed_at": pd.Timestamp.now().isoformat(timespec="seconds"),
                    }
                )
    else:
        log_progress(f"Bootstrapping with {worker_count} workers")
        future_map: dict[Any, tuple[int, Path, int, int | None]] = {}
        with ThreadPoolExecutor(max_workers=worker_count) as executor:
            for index, file_path, inferred_year, inferred_month in scheduled_files:
                log_progress(f"[{index}/{len(input_files)}] Queued {file_path.name}")
                future_map[executor.submit(run_single_file, file_path)] = (index, file_path, inferred_year, inferred_month)
            for future in as_completed(future_map):
                _, file_path, inferred_year, inferred_month = future_map[future]
                try:
                    result = future.result()
                    _append_validation_record(root, result["validation"])
                    _upsert_processed_month(root, result["processed_record"])
                    processed_rows.append(
                        {
                            "bom_year": result["bom_year"],
                            "bom_month": f"{result['bom_month']:02d}",
                            "snapshot_month": result["snapshot_month"],
                            "file_name": file_path.name,
                            "sheet_name": result["sheet_name"],
                            "parquet_path": str(result["parquet_path"]),
                            "row_count": int(len(result["data"].index)),
                            "model_count": int(result["data"]["model_suffix"].nunique()),
                            "raw_material_row_count": int(result["data"]["is_raw_material"].fillna(False).sum()),
                        }
                    )
                    validation_rows.append(result["validation"])
                    existing_months.add((int(result["bom_year"]), int(result["bom_month"])))
                    batch_bootstrap_rows.append(
                        {
                            "file_name": file_path.name,
                            "bom_year": result["bom_year"],
                            "bom_month": f"{result['bom_month']:02d}",
                            "status": "processed",
                            "row_count": int(len(result["data"].index)),
                            "model_count": int(result["data"]["model_suffix"].nunique()),
                            "parquet_path": str(result["parquet_path"]),
                            "engine_used": result["engine_used"],
                            "read_excel_seconds": round(result["timing"]["read_excel_seconds"], 2),
                            "prepare_bom_seconds": round(result["timing"]["prepare_bom_seconds"], 2),
                            "write_parquet_seconds": round(result["timing"]["write_parquet_seconds"], 2),
                            "total_seconds": round(result["timing"]["total_seconds"], 2),
                            "error_message": "",
                            "processed_at": pd.Timestamp.now().isoformat(timespec="seconds"),
                        }
                    )
                except Exception as exc:
                    log_progress(f"Failed {file_path.name}: {exc}")
                    batch_bootstrap_rows.append(
                        {
                            "file_name": file_path.name,
                            "bom_year": inferred_year,
                            "bom_month": f"{inferred_month:02d}" if inferred_month is not None else "",
                            "status": "failed",
                            "row_count": "",
                            "model_count": "",
                            "parquet_path": "",
                            "engine_used": "",
                            "read_excel_seconds": "",
                            "prepare_bom_seconds": "",
                            "write_parquet_seconds": "",
                            "total_seconds": "",
                            "error_message": str(exc),
                            "processed_at": pd.Timestamp.now().isoformat(timespec="seconds"),
                        }
                    )

    batch_bootstrap_log = pd.DataFrame(batch_bootstrap_rows)
    batch_bootstrap_log = _normalize_year_month_dtypes(batch_bootstrap_log)
    batch_log_file = batch_bootstrap_log_path(root)
    batch_log_file.parent.mkdir(parents=True, exist_ok=True)
    batch_bootstrap_log.to_csv(batch_log_file, index=False, encoding="utf-8-sig")

    processed_months_frame = pd.DataFrame(processed_rows)
    if not processed_months_frame.empty:
        processed_months_frame = _normalize_year_month_dtypes(processed_months_frame)
        processed_months_frame = processed_months_frame.sort_values(["bom_year", "bom_month"]).reset_index(drop=True)

    validation_log_frame = pd.DataFrame(validation_rows)
    if not validation_log_frame.empty:
        validation_log_frame = _normalize_year_month_dtypes(validation_log_frame)
        validation_log_frame = validation_log_frame.sort_values(["bom_year", "bom_month"]).reset_index(drop=True)

    return {
        "processed_months": processed_months_frame,
        "validation_log": validation_log_frame,
        "batch_bootstrap_log": batch_bootstrap_log,
    }


def _historical_glob(root: Path | None) -> str:
    return str(historical_root(root) / "bom_year=*" / "bom_month=*" / "*.parquet")


def _current_cache_glob(root: Path | None) -> str:
    return str(current_cache_root(root) / "bom_year=*" / "bom_month=*" / "*.parquet")


def list_processed_months(root: Path | None = None) -> pd.DataFrame:
    status_columns = [
        "bom_year",
        "bom_month",
        "snapshot_month",
        "status",
        "file_name",
        "sheet_name",
        "parquet_path",
        "processed_at",
        "row_count",
        "model_count",
        "raw_material_row_count",
        "message",
    ]
    metadata_status = _read_csv_metadata(processed_status_path(root), status_columns)
    metadata_status = _normalize_year_month_dtypes(metadata_status)
    if not historical_root(root).exists():
        return metadata_status.sort_values(["bom_year", "bom_month"]).reset_index(drop=True) if not metadata_status.empty else metadata_status

    require_duckdb()
    con = duckdb.connect()
    try:
        parquet_status = con.execute(
            """
            SELECT
                bom_year,
                LPAD(CAST(bom_month AS VARCHAR), 2, '0') AS bom_month,
                snapshot_month,
                'processed' AS status,
                MIN(snapshot_source_file) AS file_name,
                MIN(snapshot_sheet_name) AS sheet_name,
                COUNT(*) AS row_count,
                COUNT(DISTINCT model_suffix) AS model_count,
                SUM(CASE WHEN COALESCE(is_raw_material, FALSE) THEN 1 ELSE 0 END) AS raw_material_row_count
            FROM read_parquet(?)
            GROUP BY bom_year, bom_month, snapshot_month
            ORDER BY bom_year, bom_month
            """,
            [_historical_glob(root)],
        ).df()
        parquet_status = _normalize_year_month_dtypes(parquet_status)
    except Exception:
        parquet_status = pd.DataFrame(columns=status_columns)
    finally:
        con.close()

    if parquet_status.empty:
        return metadata_status.sort_values(["bom_year", "bom_month"]).reset_index(drop=True) if not metadata_status.empty else metadata_status

    parquet_status["parquet_path"] = ""
    parquet_status["processed_at"] = ""
    parquet_status["message"] = ""
    if not metadata_status.empty:
        merged = parquet_status.merge(
            metadata_status[["bom_year", "bom_month", "parquet_path", "processed_at", "message"]],
            how="left",
            on=["bom_year", "bom_month"],
            suffixes=("", "_meta"),
        )
        for column in ["parquet_path", "processed_at", "message"]:
            merged[column] = merged[f"{column}_meta"].fillna(merged[column])
            merged = merged.drop(columns=[f"{column}_meta"])
        parquet_status = merged

    for column in status_columns:
        if column not in parquet_status.columns:
            parquet_status[column] = ""
    return parquet_status[status_columns].sort_values(["bom_year", "bom_month"]).reset_index(drop=True)


def read_validation_log(root: Path | None = None) -> pd.DataFrame:
    data = _read_csv_metadata(validation_log_path(root), DEFAULT_VALIDATION_COLUMNS)
    if data.empty:
        return data
    data = _normalize_year_month_dtypes(data)
    return data.sort_values(["snapshot_kind", "bom_year", "bom_month", "file_name"]).reset_index(drop=True)


def query_snapshot(root: Path | None, bom_year: int, bom_month: int, snapshot_kind: str = "historical") -> pd.DataFrame:
    require_duckdb()
    parquet_glob = _historical_glob(root) if snapshot_kind == "historical" else _current_cache_glob(root)
    con = duckdb.connect()
    try:
        return con.execute(
            """
            SELECT *
            FROM read_parquet(?)
            WHERE bom_year = ? AND bom_month = ?
            ORDER BY model_suffix, seq, row_order, part_no
            """,
            [parquet_glob, int(bom_year), int(bom_month)],
        ).df()
    except Exception:
        return pd.DataFrame()
    finally:
        con.close()


def cache_current_month(
    file_path: Path,
    root: Path | None = None,
    manage_path: Path | None = None,
    manage_sheet_name: str | None = None,
    subsidiary: str | None = None,
    preferred_sheet_name: str | None = None,
    bom_year: int | None = None,
    bom_month: int | None = None,
    persist_as_snapshot: bool = False,
    excel_engine: str = "auto",
    include_brl_costs: bool = True,
    overwrite_existing_month: bool = False,
    snapshot_name: str | None = None,
) -> dict[str, Any]:
    snapshot_kind = "historical" if persist_as_snapshot else "current_cache"
    return process_snapshot(
        file_path=file_path,
        root=root,
        manage_path=manage_path,
        manage_sheet_name=manage_sheet_name,
        subsidiary=subsidiary,
        preferred_sheet_name=preferred_sheet_name,
        bom_year=bom_year,
        bom_month=bom_month,
        snapshot_kind=snapshot_kind,
        excel_engine=excel_engine,
        include_brl_costs=include_brl_costs,
        overwrite_existing_month=overwrite_existing_month,
        snapshot_name=snapshot_name,
    )


def save_current_snapshot_from_excel(
    file_path: Path,
    snapshot_name: str,
    bom_year: int | None = None,
    bom_month: int | None = None,
    root: Path | None = None,
    preferred_sheet_name: str | None = None,
    excel_engine: str = "calamine",
    overwrite_existing_month: bool = False,
) -> dict[str, Any]:
    result = process_snapshot(
        file_path=file_path,
        root=root,
        preferred_sheet_name=preferred_sheet_name,
        bom_year=bom_year,
        bom_month=bom_month,
        snapshot_kind="current_cache",
        excel_engine=excel_engine,
        overwrite_existing_month=overwrite_existing_month,
        snapshot_name=snapshot_name,
    )
    print(
        f"Current parquet file updated: {_safe_snapshot_stem(snapshot_name)} "
        f"({int(result['bom_year']):04d}-{int(result['bom_month']):02d})",
        flush=True,
    )
    return result


def save_historical_snapshot_from_excel(
    file_path: Path,
    snapshot_name: str,
    bom_year: int | None = None,
    bom_month: int | None = None,
    root: Path | None = None,
    preferred_sheet_name: str | None = None,
    excel_engine: str = "calamine",
    overwrite_existing_month: bool = False,
) -> dict[str, Any]:
    result = process_snapshot(
        file_path=file_path,
        root=root,
        preferred_sheet_name=preferred_sheet_name,
        bom_year=bom_year,
        bom_month=bom_month,
        snapshot_kind="historical",
        excel_engine=excel_engine,
        overwrite_existing_month=overwrite_existing_month,
        snapshot_name=snapshot_name,
    )
    print(
        f"Historical parquet file updated: {_safe_snapshot_stem(snapshot_name)} "
        f"({int(result['bom_year']):04d}-{int(result['bom_month']):02d})",
        flush=True,
    )
    return result


def inspect_parquet_snapshot(parquet_path: Path) -> dict[str, Any]:
    require_duckdb()
    parquet_path = Path(parquet_path).expanduser().resolve()
    con = duckdb.connect()
    try:
        row = con.execute(
            """
            SELECT
                MAX(CAST(bom_year AS INTEGER)) AS bom_year,
                MAX(CAST(bom_month AS INTEGER)) AS bom_month,
                COUNT(*) AS row_count,
                COUNT(DISTINCT model_suffix) AS model_count
            FROM read_parquet(?)
            """,
            [str(parquet_path)],
        ).fetchone()
    finally:
        con.close()
    bom_year = int((row[0] if row else 0) or 0)
    bom_month = int((row[1] if row else 0) or 0)
    row_count = int((row[2] if row else 0) or 0)
    model_count = int((row[3] if row else 0) or 0)
    return {
        "parquet_path": str(parquet_path),
        "bom_year": bom_year,
        "bom_month": bom_month,
        "row_count": row_count,
        "model_count": model_count,
        "label": f"{bom_year:04d}-{bom_month:02d} | {parquet_path.name} | models={model_count:,} | rows={row_count:,}",
    }


def discover_historical_snapshots(history_root: Path | None = None) -> list[dict[str, Any]]:
    require_duckdb()
    root = historical_root(history_root)
    parquet_files = sorted(root.glob("bom_year=*/bom_month=*/*.parquet"))
    rows: list[dict[str, Any]] = []
    if not parquet_files:
        return rows

    con = duckdb.connect()
    try:
        for parquet_path in parquet_files:
            try:
                columns = {
                    str(row[0])
                    for row in con.execute("DESCRIBE SELECT * FROM read_parquet(?)", [str(parquet_path)]).fetchall()
                }
                snapshot_name_expr = "MAX(CAST(snapshot_name AS VARCHAR))" if "snapshot_name" in columns else "''"
                source_file_expr = "MAX(CAST(snapshot_source_file AS VARCHAR))" if "snapshot_source_file" in columns else "''"
                processed_at_expr = "MAX(CAST(snapshot_loaded_at AS VARCHAR))" if "snapshot_loaded_at" in columns else "''"
                snapshot_month_expr = "MAX(CAST(snapshot_month AS VARCHAR))" if "snapshot_month" in columns else "''"
                stats = con.execute(
                    f"""
                    SELECT
                        MAX(CAST(bom_year AS INTEGER)) AS bom_year,
                        MAX(CAST(bom_month AS INTEGER)) AS bom_month,
                        COALESCE({snapshot_month_expr}, '') AS snapshot_month,
                        COALESCE({snapshot_name_expr}, '') AS snapshot_name,
                        COALESCE({source_file_expr}, '') AS source_file,
                        COALESCE({processed_at_expr}, '') AS processed_at,
                        COUNT(*) AS row_count,
                        COUNT(DISTINCT model_suffix) AS model_count
                    FROM read_parquet(?)
                    """,
                    [str(parquet_path)],
                ).fetchone()
            except Exception:
                continue
            if stats is None:
                continue
            bom_year = int(stats[0] or 0)
            bom_month = int(stats[1] or 0)
            snapshot_month = str(stats[2] or f"{bom_year:04d}-{bom_month:02d}")
            file_name = str(stats[3] or parquet_path.stem)
            row_count = int(stats[6] or 0)
            model_count = int(stats[7] or 0)
            label = f"{snapshot_month} | {file_name} | models={model_count:,} | rows={row_count:,}"
            rows.append(
                {
                    "label": label,
                    "bom_year": bom_year,
                    "bom_month": bom_month,
                    "snapshot_month": snapshot_month,
                    "file_name": file_name,
                    "source_file": str(stats[4] or ""),
                    "parquet_path": str(parquet_path),
                    "row_count": row_count,
                    "model_count": model_count,
                    "processed_at": str(stats[5] or ""),
                }
            )
    finally:
        con.close()

    return sorted(rows, key=lambda row: (row["bom_year"], row["bom_month"], row["file_name"]), reverse=True)


def discover_current_snapshots(history_root: Path | None = None) -> list[dict[str, Any]]:
    require_duckdb()
    root = current_cache_root(history_root)
    parquet_files = sorted(root.glob("bom_year=*/bom_month=*/*.parquet"))
    rows: list[dict[str, Any]] = []
    if not parquet_files:
        return rows

    con = duckdb.connect()
    try:
        for parquet_path in parquet_files:
            try:
                columns = {
                    str(row[0])
                    for row in con.execute("DESCRIBE SELECT * FROM read_parquet(?)", [str(parquet_path)]).fetchall()
                }
                snapshot_name_expr = "MAX(CAST(snapshot_name AS VARCHAR))" if "snapshot_name" in columns else "''"
                source_file_expr = "MAX(CAST(snapshot_source_file AS VARCHAR))" if "snapshot_source_file" in columns else "''"
                processed_at_expr = "MAX(CAST(snapshot_loaded_at AS VARCHAR))" if "snapshot_loaded_at" in columns else "''"
                snapshot_month_expr = "MAX(CAST(snapshot_month AS VARCHAR))" if "snapshot_month" in columns else "''"
                stats = con.execute(
                    f"""
                    SELECT
                        MAX(CAST(bom_year AS INTEGER)) AS bom_year,
                        MAX(CAST(bom_month AS INTEGER)) AS bom_month,
                        COALESCE({snapshot_month_expr}, '') AS snapshot_month,
                        COALESCE({snapshot_name_expr}, '') AS snapshot_name,
                        COALESCE({source_file_expr}, '') AS source_file,
                        COALESCE({processed_at_expr}, '') AS processed_at,
                        COUNT(*) AS row_count,
                        COUNT(DISTINCT model_suffix) AS model_count
                    FROM read_parquet(?)
                    """,
                    [str(parquet_path)],
                ).fetchone()
            except Exception:
                continue
            if stats is None:
                continue
            bom_year = int(stats[0] or 0)
            bom_month = int(stats[1] or 0)
            snapshot_month = str(stats[2] or f"{bom_year:04d}-{bom_month:02d}")
            file_name = str(stats[3] or parquet_path.stem)
            row_count = int(stats[6] or 0)
            model_count = int(stats[7] or 0)
            label = f"{snapshot_month} | {file_name} | models={model_count:,} | rows={row_count:,}"
            rows.append(
                {
                    "label": label,
                    "bom_year": bom_year,
                    "bom_month": bom_month,
                    "snapshot_month": snapshot_month,
                    "file_name": file_name,
                    "source_file": str(stats[4] or ""),
                    "parquet_path": str(parquet_path),
                    "row_count": row_count,
                    "model_count": model_count,
                    "processed_at": str(stats[5] or ""),
                }
            )
    finally:
        con.close()

    return sorted(rows, key=lambda row: (row["bom_year"], row["bom_month"], row["file_name"]), reverse=True)


def standardize_vi_change_list(data: pd.DataFrame) -> pd.DataFrame:
    if data.empty:
        return pd.DataFrame(
            columns=[
                "vi_item",
                "model_suffix",
                "set_model_name",
                "applied_flag",
                "saving_per_model",
                "current_actual_saving",
                "remaining_forecast_saving",
                "total_opportunity",
            ]
        )

    rename_map: dict[str, str] = {}
    for column in data.columns:
        normalized = _normalize_text(column).replace("_", " ")
        rename_map[column] = VI_COLUMN_ALIASES.get(normalized, column)
    standardized = data.rename(columns=rename_map).copy()

    for column in [
        "vi_item",
        "model_suffix",
        "set_model_name",
        "applied_flag",
        "saving_per_model",
        "current_actual_saving",
        "remaining_forecast_saving",
        "total_opportunity",
    ]:
        if column not in standardized.columns:
            standardized[column] = ""

    standardized["vi_item"] = standardized["vi_item"].fillna("").astype(str).str.strip()
    standardized["model_suffix"] = standardized["model_suffix"].fillna("").astype(str).str.strip()
    standardized["set_model_name"] = standardized["set_model_name"].fillna("").astype(str).str.strip()
    standardized["applied_flag"] = standardized["applied_flag"].map(_normalize_flag)
    for numeric_column in [
        "saving_per_model",
        "current_actual_saving",
        "remaining_forecast_saving",
        "total_opportunity",
    ]:
        standardized[numeric_column] = pd.to_numeric(standardized[numeric_column], errors="coerce").fillna(0.0)
    return standardized


def load_vi_change_list(file_path: Path, sheet_name: str | None = None) -> pd.DataFrame:
    data = pd.read_excel(file_path, sheet_name=sheet_name or 0, engine="openpyxl", dtype=object)
    return standardize_vi_change_list(data)


def build_model_material_cost_frame(current_df: pd.DataFrame, baseline_df: pd.DataFrame) -> pd.DataFrame:
    current_grouped = (
        current_df.groupby("model_suffix", as_index=False)
        .agg(
            current_model_material_cost_brl=("material_cost_brl", "sum"),
            current_model_fixed_rate_cost_brl=("fixed_rate_material_cost_brl", "sum"),
        )
    )
    baseline_grouped = (
        baseline_df.groupby("model_suffix", as_index=False)
        .agg(
            baseline_model_material_cost_brl=("material_cost_brl", "sum"),
            baseline_model_fixed_rate_cost_brl=("fixed_rate_material_cost_brl", "sum"),
        )
    )
    merged = current_grouped.merge(baseline_grouped, how="outer", on="model_suffix").fillna(0)
    merged["current_vs_2025_baseline_brl"] = (
        merged["current_model_material_cost_brl"] - merged["baseline_model_material_cost_brl"]
    )
    merged["current_vs_2025_baseline_fixed_rate_brl"] = (
        merged["current_model_fixed_rate_cost_brl"] - merged["baseline_model_fixed_rate_cost_brl"]
    )
    return merged.sort_values("model_suffix").reset_index(drop=True)


def build_vi_metrics(
    vi_change_df: pd.DataFrame,
    current_df: pd.DataFrame,
    baseline_df: pd.DataFrame,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    current_models = set(current_df["model_suffix"].fillna("").astype(str).str.strip())
    vi_data = standardize_vi_change_list(vi_change_df)
    if vi_data.empty:
        overview = pd.DataFrame(
            [
                {"metric": "vi_item_target_model_count", "value": 0},
                {"metric": "applied_model_count", "value": 0},
                {"metric": "application_rate", "value": 0.0},
                {"metric": "current_actual_saving", "value": 0.0},
                {"metric": "remaining_forecast_saving", "value": 0.0},
                {"metric": "total_opportunity", "value": 0.0},
            ]
        )
        return overview, pd.DataFrame(), build_model_material_cost_frame(current_df, baseline_df)

    vi_data["target_flag"] = vi_data["model_suffix"].ne("")
    vi_data["applied_flag"] = vi_data["applied_flag"] | vi_data["model_suffix"].isin(current_models)
    vi_data["current_actual_saving"] = vi_data["current_actual_saving"].where(
        vi_data["current_actual_saving"].ne(0),
        vi_data["saving_per_model"] * vi_data["applied_flag"].astype(int),
    )
    vi_data["remaining_forecast_saving"] = vi_data["remaining_forecast_saving"].where(
        vi_data["remaining_forecast_saving"].ne(0),
        vi_data["saving_per_model"] * (vi_data["target_flag"] & ~vi_data["applied_flag"]).astype(int),
    )
    vi_data["total_opportunity"] = vi_data["total_opportunity"].where(
        vi_data["total_opportunity"].ne(0),
        vi_data["current_actual_saving"] + vi_data["remaining_forecast_saving"],
    )

    item_metrics = (
        vi_data.groupby("vi_item", as_index=False)
        .agg(
            vi_item_target_model_count=("model_suffix", lambda s: s.astype(str).str.strip().replace("", pd.NA).dropna().nunique()),
            applied_model_count=("applied_flag", "sum"),
            current_actual_saving=("current_actual_saving", "sum"),
            remaining_forecast_saving=("remaining_forecast_saving", "sum"),
            total_opportunity=("total_opportunity", "sum"),
        )
        .sort_values("vi_item")
        .reset_index(drop=True)
    )
    item_metrics["application_rate"] = item_metrics["applied_model_count"] / item_metrics["vi_item_target_model_count"].replace(0, pd.NA)
    item_metrics["application_rate"] = item_metrics["application_rate"].fillna(0.0)

    overview = pd.DataFrame(
        [
            {"metric": "vi_item_target_model_count", "value": int(item_metrics["vi_item_target_model_count"].sum())},
            {"metric": "applied_model_count", "value": int(item_metrics["applied_model_count"].sum())},
            {
                "metric": "application_rate",
                "value": float(item_metrics["applied_model_count"].sum() / item_metrics["vi_item_target_model_count"].sum())
                if float(item_metrics["vi_item_target_model_count"].sum()) > 0
                else 0.0,
            },
            {"metric": "current_actual_saving", "value": float(item_metrics["current_actual_saving"].sum())},
            {"metric": "remaining_forecast_saving", "value": float(item_metrics["remaining_forecast_saving"].sum())},
            {"metric": "total_opportunity", "value": float(item_metrics["total_opportunity"].sum())},
        ]
    )
    return overview, item_metrics, build_model_material_cost_frame(current_df, baseline_df)


def _match_current_vs_baseline(
    baseline_df: pd.DataFrame,
    current_df: pd.DataFrame,
    set_model_map: dict[str, str],
    new_model_map: dict[str, dict[str, str]],
) -> tuple[pd.DataFrame, pd.DataFrame]:
    baseline_working = baseline_df.copy()
    current_working = current_df.copy()
    baseline_working["join_key"] = baseline_working["model_suffix"].astype(str) + "|" + baseline_working["part_no"].astype(str)
    current_working["join_key"] = current_working["model_suffix"].astype(str) + "|" + current_working["part_no"].astype(str)
    baseline_working["description_key"] = baseline_working["description"].map(_normalize_text)
    current_working["description_key"] = current_working["description"].map(_normalize_text)

    merged = current_working.merge(
        baseline_working[
            [
                "join_key",
                "model_suffix",
                "part_no",
                "description",
                "currency",
                "qty",
                "unit_price",
                "material_cost_loc",
                "material_cost_brl",
                "fixed_rate_material_cost_brl",
                "row_order",
            ]
        ],
        how="left",
        on="join_key",
        suffixes=("_curr", "_base"),
    )
    merged["baseline_match_type"] = merged["part_no_base"].notna().map(lambda matched: "PART_NO" if matched else "UNMATCHED")
    merged["change_type"] = merged.apply(
        lambda row: (
            "ADDED"
            if pd.isna(row["part_no_base"])
            else "QTY_CHANGED"
            if float(row["qty_curr"]) != float(row["qty_base"])
            else "UNCHANGED"
        ),
        axis=1,
    )

    result = pd.DataFrame(
        {
            "Model.Suffix": merged["model_suffix_curr"],
            "Set Model Name": merged["model_suffix_curr"].map(lambda value: resolve_set_model_name(str(value), set_model_map)),
            "New Model Note": merged["model_suffix_curr"].map(lambda value: resolve_new_model_note(str(value), new_model_map)),
            "Current Part No": merged["part_no_curr"],
            "Baseline Part No": merged["part_no_base"].fillna(""),
            "Current Desc": merged["description_curr"],
            "Baseline Desc": merged["description_base"].fillna(""),
            "Current Qty": merged["qty_curr"],
            "Baseline Qty": merged["qty_base"].fillna(0.0),
            "Current Curr": merged["currency_curr"],
            "Baseline Curr": merged["currency_base"].fillna(""),
            "Current Unit Price": merged["unit_price_curr"],
            "Baseline Unit Price": merged["unit_price_base"].fillna(0.0),
            "Current Material Cost LOC": merged["material_cost_loc_curr"],
            "Baseline Material Cost LOC": merged["material_cost_loc_base"].fillna(0.0),
            "Current Material Cost BRL": merged["material_cost_brl_curr"],
            "Baseline Material Cost BRL": merged["material_cost_brl_base"].fillna(0.0),
            "Current Fixed Rate Material Cost BRL": merged["fixed_rate_material_cost_brl_curr"],
            "Baseline Fixed Rate Material Cost BRL": merged["fixed_rate_material_cost_brl_base"].fillna(0.0),
            "Current vs Baseline BRL Diff": merged["material_cost_brl_curr"] - merged["material_cost_brl_base"].fillna(0.0),
            "Current vs Baseline Fixed BRL Diff": merged["fixed_rate_material_cost_brl_curr"] - merged["fixed_rate_material_cost_brl_base"].fillna(0.0),
            "Is Raw Material": merged["is_raw_material"],
            "RMC Flag": merged["rmc_flag"],
            "BOM Apply Date": merged["bom_apply_date"],
            "Match Type": merged["baseline_match_type"],
            "Change Type": merged["change_type"],
        }
    )
    result = result.sort_values(["Model.Suffix", "Current Part No"]).reset_index(drop=True)

    debug = pd.DataFrame(
        {
            "Model.Suffix": merged["model_suffix_curr"],
            "Current Part No": merged["part_no_curr"],
            "Baseline Part No": merged["part_no_base"].fillna(""),
            "Current Desc": merged["description_curr"],
            "Baseline Desc": merged["description_base"].fillna(""),
            "Matched": merged["part_no_base"].notna().map(lambda matched: "Y" if matched else ""),
            "Matched Stage": merged["baseline_match_type"],
            "Current/Base Change Type": merged["change_type"],
        }
    ).sort_values(["Model.Suffix", "Current Part No"]).reset_index(drop=True)

    return result, debug


def build_set_model_raw_material_summary(result_sheet: pd.DataFrame) -> pd.DataFrame:
    if result_sheet.empty:
        return pd.DataFrame(columns=["Set Model Name", "Raw Material BRL Diff Sum"])
    filtered = result_sheet[result_sheet["Is Raw Material"].fillna(False)].copy()
    filtered["Set Model Name"] = filtered["Set Model Name"].fillna("").astype(str).str.strip()
    filtered = filtered[filtered["Set Model Name"] != ""]
    if filtered.empty:
        return pd.DataFrame(columns=["Set Model Name", "Raw Material BRL Diff Sum"])
    return (
        filtered.groupby("Set Model Name", as_index=False)["Current vs Baseline Fixed BRL Diff"]
        .sum()
        .rename(columns={"Current vs Baseline Fixed BRL Diff": "Raw Material BRL Diff Sum"})
        .sort_values("Set Model Name")
        .reset_index(drop=True)
    )


def build_special_material_compare_sheet(result_sheet: pd.DataFrame) -> pd.DataFrame:
    if result_sheet.empty:
        return pd.DataFrame(columns=["Part No", "Desc", "Baseline Unit Price", "Current Unit Price", "Price Change %"])
    working = result_sheet.copy()

    def contains_special_keyword(row: pd.Series) -> bool:
        current_desc = _normalize_text(row.get("Current Desc", ""))
        baseline_desc = _normalize_text(row.get("Baseline Desc", ""))
        return any(keyword in current_desc or keyword in baseline_desc for keyword in SPECIAL_COMPARE_KEYWORDS)

    working = working[working.apply(contains_special_keyword, axis=1)].copy()
    if working.empty:
        return pd.DataFrame(columns=["Part No", "Desc", "Baseline Unit Price", "Current Unit Price", "Price Change %"])

    working["Part No"] = working["Current Part No"].mask(
        working["Current Part No"].fillna("").astype(str).str.strip() == "",
        working["Baseline Part No"],
    )
    working["Desc"] = working["Current Desc"].mask(
        working["Current Desc"].fillna("").astype(str).str.strip() == "",
        working["Baseline Desc"],
    )
    working["Price Change %"] = (
        (pd.to_numeric(working["Current Unit Price"], errors="coerce") - pd.to_numeric(working["Baseline Unit Price"], errors="coerce"))
        / pd.to_numeric(working["Baseline Unit Price"], errors="coerce").replace(0, pd.NA)
        * 100
    )
    return (
        working[["Part No", "Desc", "Baseline Unit Price", "Current Unit Price", "Price Change %"]]
        .sort_values(["Price Change %", "Part No"], ascending=[False, True], na_position="last")
        .reset_index(drop=True)
    )


def round_financial_columns(data: pd.DataFrame) -> pd.DataFrame:
    rounded = data.copy()
    for column in rounded.columns:
        column_name = str(column).lower()
        if any(token in column_name for token in ["cost", "price", "saving", "opportunity", "rate", "diff", "%", "value"]):
            numeric_values = pd.to_numeric(rounded[column], errors="coerce")
            rounded[column] = rounded[column].where(numeric_values.isna(), numeric_values.round(2))
    return rounded


def build_management_report(root: Path | None = None) -> dict[str, pd.DataFrame]:
    return {
        "processed_months": list_processed_months(root),
        "validation_log": read_validation_log(root),
    }


def run_current_vs_history_report(
    current_file: Path,
    baseline_month: int,
    root: Path | None = None,
    manage_path: Path | None = None,
    manage_sheet_name: str | None = None,
    subsidiary: str | None = None,
    current_sheet_name: str | None = None,
    current_year: int | None = None,
    current_month: int | None = None,
    baseline_year: int = DEFAULT_BASELINE_YEAR,
    persist_current_snapshot: bool = False,
    vi_change_file: Path | None = None,
    vi_change_sheet_name: str | None = None,
    excel_engine: str = "auto",
    include_brl_costs: bool = True,
) -> dict[str, pd.DataFrame]:
    current_result = cache_current_month(
        file_path=current_file,
        root=root,
        manage_path=manage_path,
        manage_sheet_name=manage_sheet_name,
        subsidiary=subsidiary,
        preferred_sheet_name=current_sheet_name,
        bom_year=current_year,
        bom_month=current_month,
        persist_as_snapshot=persist_current_snapshot,
        excel_engine=excel_engine,
        include_brl_costs=include_brl_costs,
    )
    current_bom = current_result["data"].copy()
    baseline_bom = query_snapshot(root, baseline_year, baseline_month, snapshot_kind="historical")
    if baseline_bom.empty:
        raise ValueError(f"No historical snapshot exists for {baseline_year}-{baseline_month:02d}.")

    set_model_map, new_model_map, set_model_sheet, new_model_sheet = load_optional_management_sheets(manage_path)
    result_sheet, debug_sheet = _match_current_vs_baseline(
        baseline_bom,
        current_bom,
        set_model_map,
        new_model_map,
    )
    summary_sheet = build_set_model_raw_material_summary(result_sheet)
    special_compare_sheet = build_special_material_compare_sheet(result_sheet)
    manage_sheet = build_manage_sheet_frame(
        load_manage_descriptions(manage_path, manage_sheet_name, subsidiary),
        subsidiary or "",
    )

    vi_change_df = load_vi_change_list(vi_change_file, vi_change_sheet_name) if vi_change_file else pd.DataFrame()
    vi_overview, vi_item_metrics, model_material_cost = build_vi_metrics(vi_change_df, current_bom, baseline_bom)

    return {
        "CurrentVsBaseline": round_financial_columns(result_sheet),
        "SetModelRawMaterialSummary": round_financial_columns(summary_sheet),
        "SpecialMaterialCompare": round_financial_columns(special_compare_sheet),
        "MatchDebug": round_financial_columns(debug_sheet),
        "Management": manage_sheet,
        "SetModelMap": set_model_sheet,
        "NewModelMap": new_model_sheet,
        "ViOverview": round_financial_columns(vi_overview),
        "ViItemMetrics": round_financial_columns(vi_item_metrics),
        "ModelMaterialCost": round_financial_columns(model_material_cost),
        "ProcessedMonths": list_processed_months(root),
        "ValidationLog": read_validation_log(root),
        "RunInfo": pd.DataFrame(
            [
                {"item": "baseline_month", "value": f"{baseline_year}-{baseline_month:02d}"},
                {"item": "current_snapshot_month", "value": current_result["snapshot_month"]},
                {"item": "current_snapshot_kind", "value": current_result["validation"]["snapshot_kind"]},
                {"item": "current_snapshot_path", "value": str(current_result["parquet_path"])},
                {"item": "historical_root", "value": str(historical_root(root))},
                {"item": "duckdb_query_source", "value": _historical_glob(root)},
            ]
        ),
    }


def _auto_fit_worksheet_columns(worksheet) -> None:
    for column_cells in worksheet.columns:
        max_length = 0
        column_letter = get_column_letter(column_cells[0].column)
        for cell in column_cells:
            value = "" if cell.value is None else str(cell.value)
            if len(value) > max_length:
                max_length = len(value)
        worksheet.column_dimensions[column_letter].width = min(max(max_length + 2, 10), 60)


def save_report_workbook(report: dict[str, pd.DataFrame], output_path: Path) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        for sheet_name, frame in report.items():
            excel_safe = frame.copy()
            excel_safe = excel_safe.replace([float("inf"), float("-inf")], "")
            excel_safe = excel_safe.astype(object).where(pd.notna(excel_safe), "")
            excel_safe.to_excel(writer, sheet_name=str(sheet_name)[:31], index=False)
        for worksheet in writer.book.worksheets:
            worksheet.freeze_panes = "A2"
            worksheet.auto_filter.ref = worksheet.dimensions
            _auto_fit_worksheet_columns(worksheet)


def dump_report_json(report: dict[str, pd.DataFrame]) -> str:
    serialized = {name: frame.fillna("").to_dict(orient="records") for name, frame in report.items()}
    return json.dumps(serialized, ensure_ascii=False, indent=2, default=str)
