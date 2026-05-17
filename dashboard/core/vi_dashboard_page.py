"""Render the VI application status dashboard page in Streamlit."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

from runtime_paths import application_root

from .vi_dashboard_data import (
    APPLIED_LABEL,
    build_kpi_summary,
    build_part_level_table_frame,
    build_subsidiary_chart_data,
    filter_dashboard_payload,
    get_filter_options,
    load_dashboard_payload,
    log_dashboard_detail_debug,
)


DETAIL_MODE = "적용 진행 상황"


def parse_dashboard_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(add_help=False)
    parser.add_argument("--report-file", dest="report_file", default="")
    args, _ = parser.parse_known_args()
    return args


def inject_page_css() -> None:
    css_path = Path(__file__).resolve().parent.parent / "assets" / "vi_dashboard.css"
    if css_path.exists():
        st.markdown(f"<style>{css_path.read_text(encoding='utf-8')}</style>", unsafe_allow_html=True)
    st.markdown(
        """
        <style>
        .vi-kpi-card {
          position: relative;
          overflow: visible !important;
        }
        .vi-kpi-tooltip {
          position: absolute;
          left: 18px;
          top: calc(100% + 10px);
          min-width: 220px;
          max-width: 320px;
          background: rgba(15, 23, 42, 0.96);
          color: #ffffff;
          border-radius: 12px;
          padding: 10px 12px;
          font-size: 12px;
          line-height: 1.45;
          box-shadow: 0 14px 28px rgba(15, 23, 42, 0.25);
          opacity: 0;
          visibility: hidden;
          transform: translateY(-6px);
          transition: opacity 0.16s ease, transform 0.16s ease, visibility 0.16s ease;
          pointer-events: none;
          z-index: 9999;
          white-space: normal;
        }
        .vi-kpi-tooltip::before {
          content: "";
          position: absolute;
          left: 18px;
          top: -6px;
          width: 12px;
          height: 12px;
          background: rgba(15, 23, 42, 0.96);
          transform: rotate(45deg);
        }
        .vi-kpi-card:hover .vi-kpi-tooltip {
          opacity: 1;
          visibility: visible;
          transform: translateY(0);
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def resolve_report_path(explicit_report_path: str | Path | None = None) -> str | None:
    if explicit_report_path:
        return str(explicit_report_path)
    if "vi_report_path" in st.session_state and st.session_state["vi_report_path"]:
        return str(st.session_state["vi_report_path"])
    return None


def render_title() -> None:
    st.markdown(f'<h1 class="vi-page-title">{DETAIL_MODE}</h1>', unsafe_allow_html=True)


def _render_checkbox_selector(title: str, options: list[str], key_prefix: str) -> list[str]:
    st.markdown(f'<div class="vi-filter-label">{title}</div>', unsafe_allow_html=True)
    if not options:
        return []
    selected: list[str] = []
    columns = st.columns(2)
    for index, option in enumerate(options):
        with columns[index % 2]:
            checked = st.checkbox(str(option), value=True, key=f"{key_prefix}_{index}_{option}")
        if checked:
            selected.append(str(option))
    return selected


def render_sidebar(payload: dict[str, object]) -> tuple[str, list[str], list[str]]:
    options = get_filter_options(payload)
    with st.sidebar:
        st.markdown('<div class="vi-filter-title">필터</div>', unsafe_allow_html=True)
        selected_subsidiaries = _render_checkbox_selector("법인", options["subsidiaries"], "subsidiary")
        selected_vi_items = _render_checkbox_selector("VI 아이템", options["vi_items"], "vi_item")
    return DETAIL_MODE, selected_subsidiaries, selected_vi_items


def render_kpi_card(label: str, value: int, theme_class: str, icon: str, tooltip: str) -> None:
    tooltip_html = tooltip.replace("\n", "<br>")
    st.markdown(
        f"""
        <div class="vi-kpi-card {theme_class}">
          <div>
            <p class="vi-kpi-label">{label}</p>
            <p class="vi-kpi-value">{value:,}</p>
          </div>
          <div class="vi-kpi-icon">{icon}</div>
          <div class="vi-kpi-tooltip">{tooltip_html}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_kpis(summary: dict[str, int]) -> None:
    log_dashboard_detail_debug(f"[DashboardDebug] KPI summary={summary}")
    target_tooltip = (
        f"Base BOM 적용 대상 모델 수: {int(summary.get('base_target_count', summary.get('base_model_count', 0))):,}\n"
        f"Current BOM 적용대상 모델 수: {int(summary.get('current_target_count', summary.get('current_model_count', 0))):,}"
    )
    applied_tooltip = (
        f"Base BOM 적용 모델 수: {int(summary.get('base_applied_count', summary['applied'])):,}\n"
        f"Current BOM 적용 모델 수: {int(summary.get('current_applied_count', 0)):,}"
    )
    not_applied_tooltip = (
        f"Base BOM 미적용 모델 수: {int(summary.get('base_not_applied_count', summary['not_applied'])):,}\n"
        f"Current BOM 미적용 모델 수: {int(summary.get('current_not_applied_count', 0)):,}"
    )
    col1, col2, col3 = st.columns(3)
    with col1:
        render_kpi_card("적용 대상", summary["applied_target"], "vi-kpi-blue", "◎", target_tooltip)
    with col2:
        render_kpi_card("적용", summary["applied"], "vi-kpi-purple", "✓", applied_tooltip)
    with col3:
        render_kpi_card("미적용", summary["not_applied"], "vi-kpi-orange", "!", not_applied_tooltip)


def build_chart_figure(chart_data: pd.DataFrame) -> go.Figure:
    color_map = {
        "LGETA": "#3b82f6",
        "LGESP": "#8b5cf6",
        "LGETH": "#f59e0b",
        "LGESR": "#10b981",
        "LGEIL": "#ef4444",
        "LGEKR": "#06b6d4",
    }
    bar_colors = [color_map.get(str(subsidiary), "#64748b") for subsidiary in chart_data["법인"]]

    figure = go.Figure()
    figure.add_bar(
        x=chart_data["법인"],
        y=chart_data["적용률(%)"],
        marker_color=bar_colors,
        text=chart_data["적용률(%)"].map(lambda value: f"{value:.1f}%"),
        textposition="outside",
        customdata=chart_data[["적용 대상", "적용", "미적용"]].to_numpy(),
        hovertemplate=(
            "법인: %{x}<br>"
            "적용률: %{y:.1f}%<br>"
            "적용 대상: %{customdata[0]:,}<br>"
            "적용: %{customdata[1]:,}<br>"
            "미적용: %{customdata[2]:,}<extra></extra>"
        ),
    )
    figure.update_layout(
        bargap=0.28,
        height=360,
        template="none",
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="#ffffff",
        margin=dict(l=72, r=20, t=12, b=8),
        showlegend=False,
    )
    figure.update_xaxes(
        title=None,
        showgrid=False,
        zeroline=False,
        showline=False,
        showticklabels=True,
        tickfont=dict(size=12, color="#475569"),
        automargin=True,
        fixedrange=True,
    )
    figure.update_yaxes(
        title=None,
        showgrid=True,
        gridcolor="#dbe2ea",
        griddash="dot",
        zeroline=False,
        showline=False,
        range=[0, 100],
        tickmode="array",
        tickvals=[0, 20, 40, 60, 80, 100],
        ticktext=["0%", "20%", "40%", "60%", "80%", "100%"],
        fixedrange=True,
    )
    return figure


def render_chart(chart_data: pd.DataFrame) -> None:
    st.markdown('<div class="vi-card">', unsafe_allow_html=True)
    st.markdown('<h2 class="vi-card-title">법인별 적용률(%)</h2>', unsafe_allow_html=True)
    st.plotly_chart(build_chart_figure(chart_data), use_container_width=True, config={"displayModeBar": False})
    st.markdown("</div>", unsafe_allow_html=True)


def style_status(value: object) -> str:
    if str(value).strip() == APPLIED_LABEL:
        return "background-color: #dcfce7; color: #15803d; font-weight: 700;"
    return "background-color: #fef3c7; color: #b45309; font-weight: 700;"


def style_amount(value: object) -> str:
    amount = pd.to_numeric(pd.Series([value]), errors="coerce").fillna(0.0).iloc[0]
    if amount > 0:
        return "color: #16a34a; font-weight: 700;"
    return "color: #6b7280;"


def _style_map(styler: pd.io.formats.style.Styler, func, subset: list[str]) -> pd.io.formats.style.Styler:
    if hasattr(styler, "map"):
        return styler.map(func, subset=subset)
    return styler.applymap(func, subset=subset)


def build_table_styler(table: pd.DataFrame) -> pd.io.formats.style.Styler:
    styler = table.style.format({"VI 금액($)": "${:,.2f}"})
    styler = _style_map(styler, style_status, subset=["적용 상태"])
    styler = _style_map(styler, style_amount, subset=["VI 금액($)"])
    return styler


def _log_frame_samples(frame: pd.DataFrame, columns: list[str], prefix: str) -> None:
    for column in columns:
        if column not in frame.columns:
            log_dashboard_detail_debug(f"[DashboardDebug] {prefix} sample {column}=<missing>")
            continue
        samples = [value for value in frame[column].fillna("").astype(str).str.strip().tolist() if value][:3]
        log_dashboard_detail_debug(f"[DashboardDebug] {prefix} sample {column}={samples}")
        for sample in samples:
            if "," in sample:
                log_dashboard_detail_debug(f"[DetailGranularityWarning] column={column} sample={sample}")
                break


def render_table(detail: pd.DataFrame, item_summary: pd.DataFrame, view_mode: str, meta: dict[str, object] | None = None) -> None:
    del item_summary, view_mode
    meta = meta or {}

    page_path = Path(__file__).resolve()
    dashboard_script_path = (page_path.parents[2] / "streamlit_vi_dashboard.py").resolve()
    data_module = sys.modules.get("dashboard.core.vi_dashboard_data")
    data_path = Path(getattr(data_module, "__file__", page_path)).resolve()

    log_dashboard_detail_debug(f"[DashboardDebug] app_root={application_root()}")
    log_dashboard_detail_debug(f"[DashboardDebug] dashboard_script path={dashboard_script_path}")
    log_dashboard_detail_debug(f"[DashboardDebug] vi_dashboard_page.py path={page_path}")
    log_dashboard_detail_debug(f"[DashboardDebug] vi_dashboard_data.py path={data_path}")
    log_dashboard_detail_debug(f"[DashboardDebug] report file path={meta.get('report_path', '')}")
    log_dashboard_detail_debug(f"[DashboardDebug] selected detail source sheet name={meta.get('detail_source_sheet', '')}")
    log_dashboard_detail_debug(f"[DashboardDebug] raw detail row count={len(detail.index)}")
    _log_frame_samples(detail, ["model_name", "parent_assy_desc_text", "base_part_no", "new_part_no"], "raw detail")

    table = build_part_level_table_frame(detail)
    log_dashboard_detail_debug(f"[DashboardDebug] final table row count={len(table.index)}")
    log_dashboard_detail_debug(f"[DashboardDebug] final table columns={list(table.columns)}")
    _log_frame_samples(table, ["모델명", "상위 ASSY Desc", "BASE P/N", "변경 P/N"], "final table")

    st.markdown('<div class="vi-card">', unsafe_allow_html=True)
    st.markdown('<h2 class="vi-card-title">모델별 상세 현황</h2>', unsafe_allow_html=True)
    st.dataframe(build_table_styler(table), use_container_width=True, hide_index=True)
    st.markdown("</div>", unsafe_allow_html=True)


def render_dashboard(explicit_report_path: str | Path | None = None) -> None:
    inject_page_css()
    report_path = resolve_report_path(explicit_report_path)
    log_dashboard_detail_debug(f"[DashboardDebug] render_dashboard report_path={report_path}")
    if report_path:
        report_file = Path(report_path)
        if report_file.exists():
            modified_text = pd.Timestamp(report_file.stat().st_mtime, unit="s").strftime("%Y-%m-%d %H:%M:%S")
            st.caption(f"보고서 파일: {report_file} | 수정 시각: {modified_text}")

    payload = load_dashboard_payload(report_path)
    detail = payload["detail"]
    item_summary = payload["item_summary"]
    subsidiary_summary = payload["subsidiary_summary"]
    model_change_summary = payload.get("model_change_summary", pd.DataFrame())
    meta = payload.get("_meta", {})

    assert isinstance(detail, pd.DataFrame)
    assert isinstance(item_summary, pd.DataFrame)
    assert isinstance(subsidiary_summary, pd.DataFrame)
    assert isinstance(model_change_summary, pd.DataFrame)
    assert isinstance(meta, dict)

    if detail.empty and item_summary.empty and subsidiary_summary.empty and model_change_summary.empty:
        st.warning("집계 파일을 읽지 못했거나 표시할 데이터가 없습니다. 결과 파일 경로와 시트 구성을 확인해 주세요.")

    view_mode, selected_subsidiaries, selected_vi_items = render_sidebar(payload)
    filtered_payload = filter_dashboard_payload(payload, selected_subsidiaries, selected_vi_items)

    render_title()
    render_kpis(
        build_kpi_summary(
            filtered_payload["item_summary"],
            filtered_payload["detail"],
            filtered_payload.get("model_change_summary"),
        )
    )
    render_chart(build_subsidiary_chart_data(filtered_payload["subsidiary_summary"], filtered_payload["detail"]))
    render_table(filtered_payload["detail"], filtered_payload["item_summary"], view_mode, meta)
