from __future__ import annotations

import argparse
import sys
from pathlib import Path

script_dir = Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parent
parent_dir = script_dir.parent
for path in [script_dir, parent_dir]:
    if str(path) not in sys.path:
        sys.path.insert(0, str(path))

from runtime_paths import output_root

try:
    from bom_lake import (
        DEFAULT_BASELINE_YEAR,
        bootstrap_historical_year,
        build_management_report,
        default_lake_root,
        run_current_vs_history_report,
        save_report_workbook,
    )
except ModuleNotFoundError:
    from compat_imports.bom_lake import (  # type: ignore[reportMissingImports]
        DEFAULT_BASELINE_YEAR,
        bootstrap_historical_year,
        build_management_report,
        default_lake_root,
        run_current_vs_history_report,
        save_report_workbook,
    )


def log(message: str) -> None:
    print(f"[Progress] {message}", flush=True)


def parse_bool_text(value: str) -> bool:
    normalized = str(value).strip().lower()
    if normalized in {"true", "1", "yes", "y"}:
        return True
    if normalized in {"false", "0", "no", "n"}:
        return False
    raise argparse.ArgumentTypeError(f"Invalid boolean value: {value}")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Build and query a historical BOM parquet lake for the dashboard workflow."
    )
    subparsers = parser.add_subparsers(dest="command", required=True)

    bootstrap_parser = subparsers.add_parser(
        "bootstrap",
        help="Process a directory of monthly historical Multi BOM Excel files into partitioned Parquet.",
    )
    bootstrap_parser.add_argument("--input-dir", required=True, help="Folder containing monthly BOM Excel files.")
    bootstrap_parser.add_argument("--history-root", help="Optional BOM lake root folder.")
    bootstrap_parser.add_argument("--manage", help="Raw material management workbook path.")
    bootstrap_parser.add_argument("--manage-sheet", help="Management sheet name.")
    bootstrap_parser.add_argument("--subsidiary", help="Subsidiary value for raw material filtering.")
    bootstrap_parser.add_argument("--year", type=int, default=DEFAULT_BASELINE_YEAR, help="Historical BOM year.")
    bootstrap_parser.add_argument("--no-skip-existing", action="store_true", help="Process months even if the same year/month already exists in metadata.")
    bootstrap_parser.add_argument("--rebuild-parquet", action="store_true", help="Rebuild parquet for matching year/month and overwrite existing monthly parquet.")
    bootstrap_parser.add_argument("--excel-engine", choices=["auto", "openpyxl", "calamine"], default="auto", help="Excel reader engine.")
    bootstrap_parser.add_argument("--workers", type=int, default=1, choices=[1, 2, 3], help="Parallel workers for Excel-to-Parquet conversion.")
    bootstrap_parser.add_argument("--include-brl-costs", type=parse_bool_text, default=True, help="Whether to calculate derived BRL cost columns.")
    bootstrap_parser.add_argument("--output", help="Optional output workbook path.")

    status_parser = subparsers.add_parser(
        "status",
        help="Show which historical BOM months are already processed and the validation log.",
    )
    status_parser.add_argument("--history-root", help="Optional BOM lake root folder.")
    status_parser.add_argument("--output", help="Optional output workbook path.")

    report_parser = subparsers.add_parser(
        "report",
        help="Compare a current month BOM against a processed historical snapshot and build dashboard-ready sheets.",
    )
    report_parser.add_argument("--curr", required=True, help="Current month BOM file path.")
    report_parser.add_argument("--curr-sheet", help="Current month BOM sheet name.")
    report_parser.add_argument("--current-year", type=int, help="Current snapshot year. Optional if inferable.")
    report_parser.add_argument("--current-month", type=int, help="Current snapshot month. Optional if inferable.")
    report_parser.add_argument("--baseline-year", type=int, default=DEFAULT_BASELINE_YEAR, help="Historical baseline year.")
    report_parser.add_argument("--baseline-month", required=True, type=int, help="Historical baseline month to compare against.")
    report_parser.add_argument("--history-root", help="Optional BOM lake root folder.")
    report_parser.add_argument("--manage", help="Raw material management workbook path.")
    report_parser.add_argument("--manage-sheet", help="Management sheet name.")
    report_parser.add_argument("--subsidiary", help="Subsidiary value for raw material filtering.")
    report_parser.add_argument("--persist-current-snapshot", action="store_true", help="Store the current month upload as a monthly historical snapshot.")
    report_parser.add_argument("--excel-engine", choices=["auto", "openpyxl", "calamine"], default="auto", help="Excel reader engine.")
    report_parser.add_argument("--include-brl-costs", type=parse_bool_text, default=True, help="Whether to calculate derived BRL cost columns.")
    report_parser.add_argument("--vi-change", help="Optional VI change list workbook path.")
    report_parser.add_argument("--vi-sheet", help="Optional VI change list sheet name.")
    report_parser.add_argument("--output", help="Optional output workbook path.")

    return parser.parse_args()


def default_output_path(command: str) -> Path:
    return output_root() / f"Historical_BOM_{command}.xlsx"


def resolve_history_root(history_root: str | None) -> Path:
    return Path(history_root) if history_root else default_lake_root()


def main() -> int:
    args = parse_args()

    try:
        history_root = resolve_history_root(getattr(args, "history_root", None))

        if args.command == "bootstrap":
            log(f"Bootstrapping historical BOM lake from: {args.input_dir}")
            report = bootstrap_historical_year(
                input_dir=Path(args.input_dir),
                root=history_root,
                manage_path=Path(args.manage) if args.manage else None,
                manage_sheet_name=args.manage_sheet,
                subsidiary=args.subsidiary,
                bom_year=int(args.year),
                skip_existing=not bool(args.no_skip_existing),
                excel_engine=args.excel_engine,
                include_brl_costs=bool(args.include_brl_costs),
                workers=int(args.workers),
                rebuild_parquet=bool(args.rebuild_parquet),
            )
            output_path = Path(args.output) if args.output else default_output_path("Bootstrap")
            save_report_workbook(
                {
                    "ProcessedMonths": report["processed_months"],
                    "ValidationLog": report["validation_log"],
                    "BatchBootstrapLog": report["batch_bootstrap_log"],
                },
                output_path,
            )
            print(f"Historical BOM bootstrap finished: {output_path}")
            print(f"Processed months: {len(report['processed_months'].index)}")
            return 0

        if args.command == "status":
            log(f"Reading BOM lake status from: {history_root}")
            report = build_management_report(history_root)
            output_path = Path(args.output) if args.output else default_output_path("Status")
            save_report_workbook(
                {
                    "ProcessedMonths": report["processed_months"],
                    "ValidationLog": report["validation_log"],
                },
                output_path,
            )
            print(f"BOM lake status workbook created: {output_path}")
            print(f"Processed months: {len(report['processed_months'].index)}")
            return 0

        if args.command == "report":
            log(f"Building current-vs-history report from: {args.curr}")
            report = run_current_vs_history_report(
                current_file=Path(args.curr),
                baseline_month=int(args.baseline_month),
                root=history_root,
                manage_path=Path(args.manage) if args.manage else None,
                manage_sheet_name=args.manage_sheet,
                subsidiary=args.subsidiary,
                current_sheet_name=args.curr_sheet,
                current_year=args.current_year,
                current_month=args.current_month,
                baseline_year=int(args.baseline_year),
                persist_current_snapshot=bool(args.persist_current_snapshot),
                vi_change_file=Path(args.vi_change) if args.vi_change else None,
                vi_change_sheet_name=args.vi_sheet,
                excel_engine=args.excel_engine,
                include_brl_costs=bool(args.include_brl_costs),
            )
            output_path = Path(args.output) if args.output else default_output_path("Report")
            save_report_workbook(report, output_path)
            print(f"Current vs history workbook created: {output_path}")
            print(f"Baseline month: {int(args.baseline_year)}-{int(args.baseline_month):02d}")
            return 0

        raise ValueError(f"Unknown command: {args.command}")
    except Exception as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
