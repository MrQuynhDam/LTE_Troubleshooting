#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Antigravity Ericsson LTE & VoLTE Telemetry, RCA & AMOS Cockpit
Design System: NOC Hybrid Telemetry Cockpit — Dark OLED (#0B0F19)
"""

from __future__ import annotations

import hashlib
import os
from typing import Any, Dict, List, Optional

import numpy as np
import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import streamlit as st

from pdf_generator import generate_diagnostic_pdf
from rca_engine import (
     compute_executive_metrics,
    convert_prb_row_to_dbm,
    generate_amos_script,
    get_col,
    load_and_process_Counters,
    natural_prb_cols,
    resolve_default_workbook,
    safe_div,
)

# ==============================================================================
# 1. PAGE CONFIGURATION & NOC DARK OLED DESIGN SYSTEM
# ==============================================================================

st.set_page_config(
    page_title="Ericsson LTE & VoLTE RAN Diagnostic Cockpit",
    page_icon="📡",
    layout="wide",
    initial_sidebar_state="expanded",
)

NOC_CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=JetBrains+Mono:wght@400;500;600;700&family=Space+Grotesk:wght@600;700&display=swap');

:root {
    --bg-oled: #0B0F19;
    --bg-card: #151D2E;
    --bg-card-hover: #1B263B;
    --border-subtle: #1E293B;
    --border-bright: #334155;
    --text-primary: #F8FAFC;
    --text-secondary: #94A3B8;
    --text-muted: #64748B;
    --accent-cyan: #38BDF8;
    --sev-critical: #FF4D4F;
    --sev-degraded: #FAAD14;
    --sev-normal: #52C41A;
    --sev-insig: #64748B;
}

html, body, [data-testid="stAppViewContainer"], .stApp {
    background-color: var(--bg-oled) !important;
    color: var(--text-primary) !important;
    font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
}

[data-testid="stSidebar"] {
    background-color: #0E1422 !important;
    border-right: 1px solid var(--border-subtle) !important;
}

[data-testid="stHeader"] {
    background-color: rgba(11, 15, 25, 0.85) !important;
}

.block-container {
    padding-top: 1.1rem !important;
    padding-bottom: 2rem !important;
    max-width: 100% !important;
}

h1, h2, h3, .noc-heading {
    font-family: 'Space Grotesk', sans-serif !important;
    letter-spacing: -0.02em;
}

.mono-num {
    font-family: 'JetBrains Mono', monospace !important;
    font-variant-numeric: tabular-nums !important;
}

/* Top Telemetry Header Bar */
.telemetry-banner {
    background: linear-gradient(180deg, #151D2E 0%, #101726 100%);
    border: 1px solid var(--border-subtle);
    border-left: 4px solid var(--accent-cyan);
    border-radius: 10px;
    padding: 14px 18px;
    margin-bottom: 12px;
}
.telemetry-title {
    font-family: 'Space Grotesk', sans-serif;
    font-size: 1.28rem;
    font-weight: 700;
    color: #F8FAFC;
    margin: 0 0 4px 0;
    display: flex;
    align-items: center;
    gap: 10px;
}
.telemetry-meta {
    font-family: 'JetBrains Mono', monospace;
    font-variant-numeric: tabular-nums;
    font-size: 0.78rem;
    color: #94A3B8;
    display: flex;
    flex-wrap: wrap;
    gap: 14px;
}
.meta-chip {
    background: rgba(15, 23, 42, 0.85);
    border: 1px solid #1E293B;
    border-radius: 5px;
    padding: 2px 8px;
    color: #E2E8F0;
}
.meta-chip b {
    color: #38BDF8;
}

/* Tier 1 Executive KPI Cards */
.kpi-card {
    background: var(--bg-card);
    border: 1px solid var(--border-subtle);
    border-radius: 10px;
    padding: 13px 15px;
    height: 112px;
    display: flex;
    flex-direction: column;
    justify-content: space-between;
    transition: border-color 0.18s ease, transform 0.18s ease;
}
.kpi-card:hover {
    border-color: var(--border-bright);
    transform: translateY(-1px);
}
.kpi-card.crit { border-top: 3px solid var(--sev-critical); }
.kpi-card.deg  { border-top: 3px solid var(--sev-degraded); }
.kpi-card.norm { border-top: 3px solid var(--sev-normal); }
.kpi-card.info { border-top: 3px solid var(--accent-cyan); }

.kpi-label {
    font-size: 0.72rem;
    font-weight: 600;
    text-transform: uppercase;
    letter-spacing: 0.05em;
    color: var(--text-secondary);
}
.kpi-value {
    font-family: 'JetBrains Mono', monospace;
    font-variant-numeric: tabular-nums;
    font-size: 1.55rem;
    font-weight: 700;
    line-height: 1.15;
    color: var(--text-primary);
}
.kpi-sub {
    font-family: 'JetBrains Mono', monospace;
    font-variant-numeric: tabular-nums;
    font-size: 0.72rem;
    display: flex;
    justify-content: space-between;
    align-items: center;
}

/* Semantic Status Badges */
.sev-badge {
    font-family: 'JetBrains Mono', monospace;
    font-size: 0.72rem;
    font-weight: 700;
    padding: 3px 9px;
    border-radius: 5px;
    letter-spacing: 0.03em;
    display: inline-block;
}
.sev-CRITICAL {
    color: #FF4D4F;
    background: rgba(255, 77, 79, 0.14);
    border: 1px solid rgba(255, 77, 79, 0.45);
}
.sev-DEGRADED {
    color: #FAAD14;
    background: rgba(250, 173, 20, 0.14);
    border: 1px solid rgba(250, 173, 20, 0.45);
}
.sev-NORMAL {
    color: #52C41A;
    background: rgba(82, 196, 26, 0.14);
    border: 1px solid rgba(82, 196, 26, 0.45);
}
.sev-STAT_INSIGNIFICANT {
    color: #94A3B8;
    background: rgba(100, 116, 139, 0.18);
    border: 1px solid rgba(100, 116, 139, 0.45);
}

/* Section Panels & Inspector Card */
.panel-hdr {
    font-family: 'Space Grotesk', sans-serif;
    font-size: 0.98rem;
    font-weight: 700;
    color: #F1F5F9;
    padding-bottom: 6px;
    margin-bottom: 8px;
    border-bottom: 1px solid var(--border-subtle);
    display: flex;
    justify-content: space-between;
    align-items: center;
}
.inspector-box {
    background: var(--bg-card);
    border: 1px solid var(--border-subtle);
    border-radius: 10px;
    padding: 14px 16px;
    margin-bottom: 10px;
}
.evidence-grid {
    display: grid;
    grid-template-columns: repeat(3, 1fr);
    gap: 7px;
    margin-top: 10px;
}
.ev-cell {
    background: #0B0F19;
    border: 1px solid #1E293B;
    border-radius: 6px;
    padding: 6px 8px;
}
.ev-lbl {
    font-size: 0.66rem;
    color: #94A3B8;
    text-transform: uppercase;
}
.ev-val {
    font-family: 'JetBrains Mono', monospace;
    font-variant-numeric: tabular-nums;
    font-size: 0.82rem;
    font-weight: 600;
    color: #F8FAFC;
}

/* Button Polish */
div.stDownloadButton > button, div.stButton > button {
    border-radius: 7px !important;
    border: 1px solid #334155 !important;
    font-weight: 600 !important;
    transition: all 0.15s ease !important;
}
div.stDownloadButton > button:hover, div.stButton > button:hover {
    border-color: #38BDF8 !important;
    color: #38BDF8 !important;
}
</style>
"""
st.markdown(NOC_CSS, unsafe_allow_html=True)

PLOTLY_BG = "#151D2E"
PLOTLY_PAPER = "#151D2E"
PLOTLY_GRID = "#1E293B"
PLOTLY_FONT = "JetBrains Mono, monospace"


# ==============================================================================
# 2. CACHED DATA LOADING & PDF GENERATION
# ==============================================================================

@st.cache_data(show_spinner=False)
def cached_load_path(file_path: str, mtime: float) -> Dict[str, Any]:
    return load_and_process_Counters(file_path)


@st.cache_data(show_spinner=False)
def cached_load_bytes(raw_bytes: bytes, filename: str, digest: str) -> Dict[str, Any]:
    import io
    return load_and_process_Counters(io.BytesIO(raw_bytes), source_name=filename)


def base_plotly_layout(fig: go.Figure, height: int = 330, title: str = "") -> go.Figure:
    fig.update_layout(
        title=dict(text=title, font=dict(family="Space Grotesk, sans-serif", size=13, color="#E2E8F0"), x=0.01),
        height=height,
        margin=dict(l=48, r=36, t=38 if title else 18, b=38),
        paper_bgcolor=PLOTLY_PAPER,
        plot_bgcolor=PLOTLY_BG,
        font=dict(family=PLOTLY_FONT, size=11, color="#CBD5E1"),
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="right",
            x=1.0,
            bgcolor="rgba(0,0,0,0)",
            font=dict(size=10.5),
        ),
        hoverlabel=dict(bgcolor="#0B0F19", bordercolor="#38BDF8", font=dict(family=PLOTLY_FONT, size=11, color="#F8FAFC")),
    )
    fig.update_xaxes(gridcolor=PLOTLY_GRID, zerolinecolor=PLOTLY_GRID, linecolor="#334155")
    fig.update_yaxes(gridcolor=PLOTLY_GRID, zerolinecolor=PLOTLY_GRID, linecolor="#334155")
    return fig


# ==============================================================================
# 3. COLLAPSIBLE SIDEBAR: DATA INGESTION & MULTI-LAYER FILTERS
# ==============================================================================

TRIAGE_PRESETS = [
    "Show All (Full Network)",
    "All Degraded / Critical",
    "Drop Call Spike (DCR > 1%)",
    "Accessibility Failure (CSSR < 97%)",
    "Mobility Failure (HOSR < 96%)",
    "VoLTE QCI=1 Degradation",
    "High UL Interference / PIM (> -110 dBm)",
    "IRAT Offload Loop (> 5%)",
]


def _reset_filters() -> None:
    st.session_state["flt_preset"] = "Show All (Full Network)"
    st.session_state["flt_erbs"] = []
    st.session_state["flt_cell"] = []
    st.session_state["flt_earfcn"] = []
    st.session_state["flt_date"] = []
    st.session_state["flt_hour"] = []
    st.session_state["flt_rca"] = []
    st.session_state["selected_cell_id"] = None
    st.session_state["selected_rop_time"] = None


for _k, _default in [
    ("flt_preset", "Show All (Full Network)"),
    ("flt_erbs", []),
    ("flt_cell", []),
    ("flt_earfcn", []),
    ("flt_date", []),
    ("flt_hour", []),
    ("flt_rca", []),
    ("view_mode", "Aggregated Cell View"),
    ("selected_cell_id", None),
    ("selected_rop_time", None),
]:
    if _k not in st.session_state:
        st.session_state[_k] = _default

with st.sidebar:
    st.markdown(
        "<div class='noc-heading' style='font-size:1.08rem;color:#38BDF8;margin-bottom:4px;'>"
        "📡 ANTIGRAVITY RAN v9.0</div>"
        "<div style='font-size:0.75rem;color:#94A3B8;margin-bottom:10px;'>"
        "Ericsson LTE & VoLTE Diagnostics + AMOS Engine</div>",
        unsafe_allow_html=True,
    )

    uploaded = st.file_uploader(
        "Drag & Drop BO / ENIQ Dump (.xlsx, .xls, .csv)",
        type=["xlsx", "xls", "csv"],
        help="Supports raw 4-sheet SAP BusinessObjects Counters.xlsx, 3-sheet Counters_Organized.xlsx, or flat CSV.",
    )

    with st.spinner("Executing Vectorized 4-Sheet ETL & 3GPP/Ericsson RCA Trees..."):
        if uploaded is not None:
            raw_bytes = uploaded.getvalue()
            digest = hashlib.md5(raw_bytes).hexdigest()
            bundle = cached_load_bytes(raw_bytes, uploaded.name, digest)
        else:
            default_path = resolve_default_workbook()
            if not default_path:
                st.error("No `Counters.xlsx` found in workspace or `./Data/`. Please upload a file.")
                st.stop()
            bundle = cached_load_path(default_path, os.path.getmtime(default_path))

    hourly_all: pd.DataFrame = bundle["hourly_df"]
    cell_all: pd.DataFrame = bundle["cell_df"]
    prb_hourly_all: Optional[pd.DataFrame] = bundle["prb_hourly_df"]
    prb_cell_all: Optional[pd.DataFrame] = bundle["prb_cell_df"]
    meta: Dict[str, Any] = bundle["meta"]

    st.divider()
    st.markdown("**⚡ Quick Triage Presets**")
    selected_preset = st.selectbox(
        "Diagnostic Triage Filter",
        options=TRIAGE_PRESETS,
        key="flt_preset",
        label_visibility="collapsed",
    )

    st.markdown("**🎛️ Granular Multi-Select Filters**")
    erbs_opts = sorted(hourly_all["ERBS Id"].dropna().unique().tolist()) if "ERBS Id" in hourly_all.columns else []
    sel_erbs = st.multiselect("ERBS Id (eNodeB)", options=erbs_opts, key="flt_erbs")

    cell_pool = hourly_all[hourly_all["ERBS Id"].isin(sel_erbs)] if sel_erbs else hourly_all
    cell_opts = sorted(cell_pool["EUtranCell Id"].dropna().unique().tolist()) if "EUtranCell Id" in cell_pool.columns else []
    sel_cells = st.multiselect("EUtranCell Id", options=cell_opts, key="flt_cell")

    earfcn_opts = sorted(hourly_all["earfcndl"].dropna().astype(int).unique().tolist()) if "earfcndl" in hourly_all.columns else []
    sel_earfcn = st.multiselect("Carrier Layer (earfcndl)", options=earfcn_opts, key="flt_earfcn")

    date_opts = sorted(hourly_all["Date"].dropna().astype(str).unique().tolist()) if "Date" in hourly_all.columns else []
    sel_dates = st.multiselect("Date", options=date_opts, key="flt_date")

    hour_opts = sorted(hourly_all["Hour"].dropna().astype(int).unique().tolist()) if "Hour" in hourly_all.columns else []
    sel_hours = st.multiselect("Hour (00..23)", options=hour_opts, key="flt_hour")

    rca_opts = sorted(hourly_all["RCA_Category"].dropna().unique().tolist()) if "RCA_Category" in hourly_all.columns else []
    sel_rca = st.multiselect("RCA Category", options=rca_opts, key="flt_rca")

    st.button("🔄 Reset All Filters", on_click=_reset_filters, use_container_width=True)


# ==============================================================================
# 4. APPLY GRANULAR & TRIAGE FILTERS TO HOURLY & CELL DATAFRAMES
# ==============================================================================

def _apply_scope_filters(df: pd.DataFrame, is_hourly: bool) -> pd.DataFrame:
    res = df
    if sel_erbs and "ERBS Id" in res.columns:
        res = res[res["ERBS Id"].isin(sel_erbs)]
    if sel_cells and "EUtranCell Id" in res.columns:
        res = res[res["EUtranCell Id"].isin(sel_cells)]
    if sel_earfcn and "earfcndl" in res.columns:
        res = res[res["earfcndl"].astype(int).isin(sel_earfcn)]
    if is_hourly:
        if sel_dates and "Date" in res.columns:
            res = res[res["Date"].astype(str).isin(sel_dates)]
        if sel_hours and "Hour" in res.columns:
            res = res[res["Hour"].astype(int).isin(sel_hours)]
    return res


def _apply_triage_preset(df: pd.DataFrame, preset: str) -> pd.DataFrame:
    res = df
    if preset == "All Degraded / Critical":
        res = res[res["Severity"].isin(["CRITICAL", "DEGRADED"])]
    elif preset == "Drop Call Spike (DCR > 1%)":
        res = res[(res["DCR_Denom"] >= 30) & (res["DCR_Overall_Pct"] > 1.0)]
    elif preset == "Accessibility Failure (CSSR < 97%)":
        res = res[(res["CSSR_Denom"] >= 30) & (res["CSSR_Pct"] < 97.0)]
    elif preset == "Mobility Failure (HOSR < 96%)":
        res = res[
            ((res["Combined_HO_Att"] >= 20) & (res["Combined_HOSR_Pct"] < 96.0))
            | ((get_col(res, "pmCellHoPrepAttLteIntraF") >= 20) & (res["Intra_Freq_HOSR_Pct"] < 96.0))
            | ((get_col(res, "pmCellHoPrepAttLteInterF") >= 20) & (res["Inter_Freq_HOSR_Pct"] < 94.0))
        ]
    elif preset == "VoLTE QCI=1 Degradation":
        res = res[
            ((res["VoLTE_DCR_Denom"] >= 10) & (res["VoLTE_DCR_Rel_Pct"] > 0.8))
            | ((res["VoLTE_CSSR_Denom"] >= 10) & (res["VoLTE_CSSR_RAN_Pct"] < 97.5))
        ]
    elif preset == "High UL Interference / PIM (> -110 dBm)":
        res = res[
            (res["UL_Interf_Avg_dBm"] > -110.0)
            | ((get_col(res, "PRB_Max_dBm", -121.0) > -105.0) & (res["PRB_Peak_to_Median_dB"] > 6.0))
        ]
    elif preset == "IRAT Offload Loop (> 5%)":
        res = res[(get_col(res, "pmRrcConnEstabSucc") >= 30) & (res["IRAT_Offload_Pct"] > 5.0)]

    if sel_rca and "RCA_Category" in res.columns:
        res = res[res["RCA_Category"].isin(sel_rca)]
    return res.reset_index(drop=True)


# If Date/Hour filter is active in Aggregated Cell View, re-aggregate dynamically from filtered hourly rows!
hourly_scoped = _apply_scope_filters(hourly_all, is_hourly=True)
if sel_dates or sel_hours:
    from rca_engine import aggregate_to_cell_level, compute_kpis_and_rca
    prb_h_sub = _apply_scope_filters(prb_hourly_all, is_hourly=True) if prb_hourly_all is not None else None
    c_raw_sub, prb_cell_scoped = aggregate_to_cell_level(hourly_scoped, prb_h_sub)
    cell_scoped = compute_kpis_and_rca(c_raw_sub, is_aggregated=True)
else:
    cell_scoped = _apply_scope_filters(cell_all, is_hourly=False)
    prb_cell_scoped = prb_cell_all


# ==============================================================================
# 5. TOP TELEMETRY HEADER (VIEW TOGGLE + PDF EXPORT)
# ==============================================================================

hdr_col_left, hdr_col_mid, hdr_col_right = st.columns([6.2, 3.3, 2.5], vertical_alignment="center")

with hdr_col_mid:
    view_mode = st.radio(
        "Telemetry Granularity",
        options=["Aggregated Cell View", "Hourly ROP View"],
        horizontal=True,
        key="view_mode",
        label_visibility="collapsed",
    )

is_cell_view = view_mode == "Aggregated Cell View"
active_scope_df = cell_scoped if is_cell_view else hourly_scoped
active_filtered_df = _apply_triage_preset(active_scope_df, selected_preset)
exec_metrics = compute_executive_metrics(active_scope_df)

with hdr_col_left:
    st.markdown(
        f"""
        <div class="telemetry-banner">
            <div class="telemetry-title">
                <span>ERICSSON LTE &amp; VoLTE RAN DIAGNOSTIC COCKPIT</span>
                <span class="sev-badge sev-NORMAL" style="font-size:0.68rem;">SKILL ENGINE v9.0</span>
            </div>
            <div class="telemetry-meta">
                <span class="meta-chip">📄 File: <b>{meta['filename']}</b></span>
                <span class="meta-chip">📅 Date: <b>{meta['date_range']}</b></span>
                <span class="meta-chip">🗼 ERBS: <b>{meta['total_erbs']}</b> | Cells: <b>{meta['total_cells']}</b></span>
                <span class="meta-chip">⏱️ Cell-Hours: <b>{meta['total_rops']:,} rows</b></span>
                <span class="meta-chip">📡 EARFCN: <b>{', '.join(map(str, meta['carriers']))}</b></span>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

with hdr_col_right:
    # Prepare AMOS lookup for top degraded cells so PDF Export is instantaneous & complete
    pdf_meta = {**meta, "view_mode": view_mode}
    deg_for_pdf = active_filtered_df[active_filtered_df["Severity"].isin(["CRITICAL", "DEGRADED"])].head(15)
    amos_map = {str(r["EUtranCell Id"]): generate_amos_script(r) for _, r in deg_for_pdf.iterrows()}
    if st.session_state["selected_cell_id"]:
        sel_match = active_scope_df[active_scope_df["EUtranCell Id"] == st.session_state["selected_cell_id"]]
        if not sel_match.empty:
            amos_map[st.session_state["selected_cell_id"]] = generate_amos_script(sel_match.iloc[0])

    pdf_bytes = generate_diagnostic_pdf(
        df_display=active_filtered_df if not active_filtered_df.empty else active_scope_df,
        exec_metrics=exec_metrics,
        meta=pdf_meta,
        amos_lookup=amos_map,
        selected_cell_id=st.session_state["selected_cell_id"],
    )
    st.download_button(
        label="📥 Export Diagnostic PDF Report",
        data=pdf_bytes,
        file_name=f"LTE_RAN_Diagnostic_{view_mode.replace(' ', '_')}.pdf",
        mime="application/pdf",
        use_container_width=True,
    )


# ==============================================================================
# 6. TIER 1 — EXECUTIVE METRIC BAR (6 KPI CARDS)
# ==============================================================================

c1, c2, c3, c4, c5, c6 = st.columns(6)

scope_unit = "Cells" if is_cell_view else "Cell-ROPs"
cssr_delta = exec_metrics["sys_cssr_pct"] - 98.5
cssr_cls = "crit" if exec_metrics["sys_cssr_pct"] < 97.0 else ("deg" if exec_metrics["sys_cssr_pct"] < 98.5 else "norm")
dcr_cls = "crit" if exec_metrics["sys_dcr_pct"] > 1.0 else ("deg" if exec_metrics["sys_dcr_pct"] > 0.5 else "norm")
hosr_cls = "crit" if exec_metrics["sys_hosr_pct"] < 95.0 else ("deg" if exec_metrics["sys_hosr_pct"] < 98.0 else "norm")
worst_cls = "crit" if exec_metrics["critical_count"] > 0 else ("deg" if exec_metrics["degraded_count"] > 0 else "norm")
vdcr_cls = "crit" if exec_metrics["sys_volte_dcr_pct"] > 0.8 else ("deg" if exec_metrics["sys_volte_dcr_pct"] > 0.3 else "norm")

with c1:
    st.markdown(
        f"""
        <div class="kpi-card info">
            <div class="kpi-label">Evaluated {scope_unit}</div>
            <div class="kpi-value">{exec_metrics['total_evaluated']:,}</div>
            <div class="kpi-sub">
                <span style="color:#94A3B8;">STAT_INSIG Excluded:</span>
                <span style="color:#CBD5E1;font-weight:700;">{exec_metrics['stat_insig']:,}</span>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

with c2:
    st.markdown(
        f"""
        <div class="kpi-card {worst_cls}">
            <div class="kpi-label">Worst {scope_unit} Detected</div>
            <div class="kpi-value" style="color:{'#FF4D4F' if exec_metrics['worst_count'] > 0 else '#52C41A'};">
                {exec_metrics['worst_count']:,} <span style="font-size:0.95rem;color:#94A3B8;">({exec_metrics['worst_pct']:.1f}%)</span>
            </div>
            <div class="kpi-sub">
                <span style="color:#FF4D4F;">CRIT: {exec_metrics['critical_count']}</span>
                <span style="color:#FAAD14;">DEG: {exec_metrics['degraded_count']}</span>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

with c3:
    st.markdown(
        f"""
        <div class="kpi-card {cssr_cls}">
            <div class="kpi-label">System CSSR (%)</div>
            <div class="kpi-value">{exec_metrics['sys_cssr_pct']:.2f}%</div>
            <div class="kpi-sub">
                <span style="color:#94A3B8;">vs 98.5% Target:</span>
                <span style="color:{'#52C41A' if cssr_delta >= 0 else '#FF4D4F'};font-weight:700;">{cssr_delta:+.2f}%</span>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

with c4:
    dcr_badge_txt = "⚠ ALERT > 1.0%" if exec_metrics["sys_dcr_pct"] > 1.0 else ("✓ <= 0.50% Target" if exec_metrics["sys_dcr_pct"] <= 0.5 else "Watch (> 0.5%)")
    dcr_badge_clr = "#FF4D4F" if exec_metrics["sys_dcr_pct"] > 1.0 else ("#52C41A" if exec_metrics["sys_dcr_pct"] <= 0.5 else "#FAAD14")
    st.markdown(
        f"""
        <div class="kpi-card {dcr_cls}">
            <div class="kpi-label">System Overall DCR (%)</div>
            <div class="kpi-value">{exec_metrics['sys_dcr_pct']:.2f}%</div>
            <div class="kpi-sub">
                <span style="color:#94A3B8;">Baseline &le; 0.5%</span>
                <span style="color:{dcr_badge_clr};font-weight:700;">{dcr_badge_txt}</span>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

with c5:
    st.markdown(
        f"""
        <div class="kpi-card {hosr_cls}">
            <div class="kpi-label">Combined HOSR (%)</div>
            <div class="kpi-value">{exec_metrics['sys_hosr_pct']:.2f}%</div>
            <div class="kpi-sub">
                <span style="color:#94A3B8;">Intra + Inter Mobility</span>
                <span style="color:#38BDF8;">Target &ge; 98.0%</span>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

with c6:
    st.markdown(
        f"""
        <div class="kpi-card {vdcr_cls}">
            <div class="kpi-label">VoLTE QCI=1 DCR &amp; CQI</div>
            <div class="kpi-value">{exec_metrics['sys_volte_dcr_pct']:.2f}%</div>
            <div class="kpi-sub">
                <span style="color:#94A3B8;">VoLTE CSSR: {exec_metrics['sys_volte_cssr_pct']:.2f}%</span>
                <span style="color:#38BDF8;font-weight:700;">CQI: {exec_metrics['sys_avg_cqi']:.2f}</span>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

st.markdown("<div style='height:10px;'></div>", unsafe_allow_html=True)


# ==============================================================================
# 7. TIER 2 — ANALYTICS & DIAGNOSTICS GRID (7 : 5 RATIO)
# ==============================================================================

col_analytics, col_pareto = st.columns([7, 5], gap="medium")

# ------------------------------------------------------------------------------
# RIGHT COLUMN FIRST LOGICALLY? No, render in exact columns; Pareto on_select updates session_state!
# ------------------------------------------------------------------------------
with col_pareto:
    st.markdown(
        "<div class='panel-hdr'>"
        "<span>🎯 Pareto Top 10 Worst Cells (Click Bar to Cross-Filter)</span>"
        "</div>",
        unsafe_allow_html=True,
    )
    pareto_metric_label = st.selectbox(
        "Rank Worst Cells By",
        options=[
            "Overall DCR %",
            "VoLTE DCR %",
            "CSSR %",
            "Intra-Freq HOSR %",
            "Inter-Freq HOSR %",
            "UL Interference dBm",
            "IRAT Offload %",
        ],
        index=0,
        label_visibility="collapsed",
    )

    pareto_cfg = {
        "Overall DCR %": ("DCR_Overall_Pct", False, 1.0, 2.5, "%", cell_scoped[cell_scoped["DCR_Denom"] >= 30]),
        "VoLTE DCR %": ("VoLTE_DCR_Rel_Pct", False, 0.8, 2.0, "%", cell_scoped[cell_scoped["VoLTE_DCR_Denom"] >= 10]),
        "CSSR %": ("CSSR_Pct", True, 97.0, 95.0, "%", cell_scoped[cell_scoped["CSSR_Denom"] >= 30]),
        "Intra-Freq HOSR %": ("Intra_Freq_HOSR_Pct", True, 96.0, 92.0, "%", cell_scoped[get_col(cell_scoped, "pmCellHoPrepAttLteIntraF") >= 20]),
        "Inter-Freq HOSR %": ("Inter_Freq_HOSR_Pct", True, 94.0, 88.0, "%", cell_scoped[get_col(cell_scoped, "pmCellHoPrepAttLteInterF") >= 20]),
        "UL Interference dBm": ("UL_Interf_Avg_dBm", False, -110.0, -108.0, " dBm", cell_scoped),
        "IRAT Offload %": ("IRAT_Offload_Pct", False, 5.0, 10.0, "%", cell_scoped[get_col(cell_scoped, "pmRrcConnEstabSucc") >= 30]),
    }
    p_col, p_asc, p_deg_thr, p_crit_thr, p_unit, p_pool = pareto_cfg[pareto_metric_label]
    p_valid = p_pool.dropna(subset=[p_col]).sort_values(p_col, ascending=p_asc).head(10).iloc[::-1]

    if p_valid.empty:
        st.info("No cells meet minimum sample denominator for the selected Pareto metric.")
    else:
        bar_colors = []
        for v, c_id in zip(p_valid[p_col], p_valid["EUtranCell Id"]):
            if st.session_state["selected_cell_id"] == c_id:
                bar_colors.append("#38BDF8")
            elif (p_asc and v < p_crit_thr) or (not p_asc and v > p_crit_thr):
                bar_colors.append("#FF4D4F")
            elif (p_asc and v < p_deg_thr) or (not p_asc and v > p_deg_thr):
                bar_colors.append("#FAAD14")
            else:
                bar_colors.append("#52C41A")

        fig_pareto = go.Figure(
            go.Bar(
                x=p_valid[p_col],
                y=p_valid["EUtranCell Id"],
                orientation="h",
                marker=dict(color=bar_colors, line=dict(color="#0B0F19", width=1)),
                text=[f"{v:.2f}{p_unit}" for v in p_valid[p_col]],
                textposition="auto",
                customdata=np.stack([p_valid["EUtranCell Id"], p_valid["ERBS Id"], p_valid["Root_Cause"]], axis=-1),
                hovertemplate=(
                    "<b>%{customdata[0]}</b> (eNB: %{customdata[1]})<br>"
                    f"{pareto_metric_label}: <b>%{{x:.2f}}{p_unit}</b><br>"
                    "RCA: %{customdata[2]}<extra></extra>"
                ),
            )
        )
        base_plotly_layout(fig_pareto, height=345)
        if p_asc:
            min_v = float(p_valid[p_col].min())
            fig_pareto.update_xaxes(range=[max(0.0, min_v - 5.0), 100.5], title_text=pareto_metric_label)
        elif pareto_metric_label == "UL Interference dBm":
            # For negative dBm values, anchor base at -122 dBm so louder interference bars extend rightward!
            vals_shift = p_valid[p_col] - (-122.0)
            fig_pareto.data[0].x = vals_shift
            fig_pareto.data[0].base = -122.0
            fig_pareto.update_xaxes(range=[-122.0, max(float(p_valid[p_col].max()) + 2.5, -106.0)], title_text="UL Noise Floor (dBm)")
        else:
            fig_pareto.update_xaxes(title_text=pareto_metric_label)

        pareto_event = st.plotly_chart(
            fig_pareto,
            use_container_width=True,
            on_select="rerun",
            selection_mode="points",
            key="pareto_selector",
        )
        if pareto_event and pareto_event.selection and pareto_event.selection.points:
            pt = pareto_event.selection.points[0]
            clicked_cell = pt.get("y") or (pt.get("customdata")[0] if pt.get("customdata") else None)
            if clicked_cell and clicked_cell != st.session_state.get("_last_pareto_cell"):
                st.session_state["_last_pareto_cell"] = clicked_cell
                st.session_state["selected_cell_id"] = str(clicked_cell)
        elif pareto_event and pareto_event.selection and not pareto_event.selection.points:
            st.session_state["_last_pareto_cell"] = None


with col_analytics:
    st.markdown(
        "<div class='panel-hdr'>"
        "<span>📈 Multi-Domain RAN Telemetry &amp; Radio Distribution</span>"
        f"<span class='mono-num' style='font-size:0.75rem;color:#38BDF8;'>"
        f"{'Cell Focus: ' + st.session_state['selected_cell_id'] if st.session_state['selected_cell_id'] else 'Network Scope'}"
        "</span></div>",
        unsafe_allow_html=True,
    )
    tab_trend, tab_cqi, tab_carrier = st.tabs([
        "⏱️ Hourly Trendline",
        "📶 Unified CQI & MIMO Distribution",
        "📡 Carrier Layer Breakdown",
    ])

    # ---- TAB 1: HOURLY TRENDLINE ----
    with tab_trend:
        t_ctrl1, t_ctrl2 = st.columns([5, 5])
        with t_ctrl1:
            trend_domain = st.radio(
                "Trend Pair",
                options=["MBB: CSSR (%) vs Overall DCR (%)", "VoLTE: QCI=1 CSSR (%) vs VoLTE DCR (%)"],
                horizontal=True,
                label_visibility="collapsed",
            )
        with t_ctrl2:
            use_cell_trend = st.checkbox(
                f"Scope Trend to Selected Cell ({st.session_state['selected_cell_id'] or 'None'})",
                value=bool(st.session_state["selected_cell_id"]),
                disabled=not bool(st.session_state["selected_cell_id"]),
            )

        trend_src = (
            hourly_scoped[hourly_scoped["EUtranCell Id"] == st.session_state["selected_cell_id"]]
            if (use_cell_trend and st.session_state["selected_cell_id"])
            else hourly_scoped
        )
        if trend_src.empty:
            st.info("No hourly records in current filter scope.")
        else:
            # Counter-weighted hourly aggregation
            grp = trend_src.groupby("ROP_Time", as_index=False).agg(
                rrc_suc=("pmRrcConnEstabSucc", "sum"),
                rrc_att=("pmRrcConnEstabAtt", "sum"),
                rrc_reatt=("pmRrcConnEstabAttReatt", "sum"),
                s1_suc=("pmS1SigConnEstabSucc", "sum"),
                s1_att=("pmS1SigConnEstabAtt", "sum"),
                erab_suc=("pmErabEstabSuccInit", "sum"),
                erab_att=("pmErabEstabAttInit", "sum"),
                dcr_num=("Active_Drops", "sum"),
                dcr_den=("DCR_Denom", "sum"),
                v_suc=("pmErabEstabSuccAddedQci", "sum"),
                v_att=("VoLTE_CSSR_Denom", "sum"),
                vd_num=("VoLTE_Active_Drops", "sum"),
                vd_den=("VoLTE_DCR_Denom", "sum"),
            ).sort_values("ROP_Time")

            fig_trend = make_subplots(specs=[[{"secondary_y": True}]])
            if trend_domain.startswith("MBB"):
                y1 = (
                    safe_div(grp["rrc_suc"], grp["rrc_att"] - grp["rrc_reatt"])
                    * safe_div(grp["s1_suc"], grp["s1_att"])
                    * safe_div(grp["erab_suc"], grp["erab_att"])
                    * 100.0
                ).round(2)
                y2 = safe_div(grp["dcr_num"], grp["dcr_den"], 100.0).round(2)
                y1_name, y2_name = "CSSR (%)", "Overall DCR (%)"
                y1_thr, y2_thr = 98.5, 1.0
            else:
                y1 = safe_div(grp["v_suc"], grp["v_att"], 100.0).round(2)
                y2 = safe_div(grp["vd_num"], grp["vd_den"], 100.0).round(2)
                y1_name, y2_name = "VoLTE QCI=1 CSSR (%)", "VoLTE QCI=1 DCR (%)"
                y1_thr, y2_thr = 99.0, 0.8

            fig_trend.add_trace(
                go.Scatter(
                    x=grp["ROP_Time"],
                    y=y1,
                    name=y1_name,
                    mode="lines+markers",
                    line=dict(color="#38BDF8", width=2.4),
                    marker=dict(size=5),
                ),
                secondary_y=False,
            )
            fig_trend.add_trace(
                go.Bar(
                    x=grp["ROP_Time"],
                    y=y2,
                    name=y2_name,
                    marker=dict(color="rgba(255, 77, 79, 0.55)", line=dict(color="#FF4D4F", width=1)),
                ),
                secondary_y=True,
            )
            fig_trend.add_hline(y=y2_thr, line_dash="dot", line_color="#FAAD14", secondary_y=True, annotation_text=f"Degraded ({y2_thr}%)")
            base_plotly_layout(fig_trend, height=295)
            fig_trend.update_yaxes(title_text=y1_name, secondary_y=False)
            fig_trend.update_yaxes(title_text=y2_name, secondary_y=True, showgrid=False)
            st.plotly_chart(fig_trend, use_container_width=True)

    # ---- TAB 2: UNIFIED CQI & MIMO DISTRIBUTION ----
    with tab_cqi:
        cqi_src = (
            active_scope_df[active_scope_df["EUtranCell Id"] == st.session_state["selected_cell_id"]]
            if st.session_state["selected_cell_id"] and not active_scope_df[active_scope_df["EUtranCell Id"] == st.session_state["selected_cell_id"]].empty
            else active_scope_df
        )
        bin_cols = [f"CQI_Bin_{i}" for i in range(16) if f"CQI_Bin_{i}" in cqi_src.columns]
        if len(bin_cols) == 16 and not cqi_src.empty:
            sums = cqi_src[bin_cols].sum()
            tot_samples = max(float(sums.sum()), 1.0)
            pcts = (sums / tot_samples * 100.0).values
            labels = ["Bin 0 (OOR)"] + [f"CQI {i}" for i in range(1, 16)]
            colors_cqi = ["#FF4D4F"] + ["#FAAD14"] * 6 + ["#38BDF8"] * 5 + ["#52C41A"] * 4

            w_avg_cqi = float(get_col(cqi_src, "CQI_Weighted_Sum").sum() / max(get_col(cqi_src, "CQI_Valid_Samples").sum(), 1.0))
            hom_pct = float(get_col(cqi_src, "Hom_256QAM_Samples").sum() / tot_samples * 100.0)
            r1_pct = float(get_col(cqi_src, "Rank1_Samples").sum() / max(get_col(cqi_src, "Rank_Total_Samples").sum(), 1.0) * 100.0)
            oor_pct = float(pcts[0])

            st.markdown(
                f"<div class='telemetry-meta' style='margin-bottom:6px;'>"
                f"<span class='meta-chip'>Avg CQI (1..15): <b>{w_avg_cqi:.2f}</b></span>"
                f"<span class='meta-chip'>CQI OOR (Bin 0): <b style='color:{'#FF4D4F' if oor_pct > 5 else '#52C41A'};'>{oor_pct:.2f}%</b></span>"
                f"<span class='meta-chip'>256-QAM Share: <b>{hom_pct:.1f}%</b></span>"
                f"<span class='meta-chip'>MIMO Rank-1 Ratio: <b style='color:{'#FAAD14' if r1_pct > 75 else '#38BDF8'};'>{r1_pct:.1f}%</b></span>"
                f"</div>",
                unsafe_allow_html=True,
            )

            fig_cqi = go.Figure(
                go.Bar(
                    x=labels,
                    y=pcts,
                    marker=dict(color=colors_cqi),
                    text=[f"{p:.1f}%" for p in pcts],
                    textposition="outside",
                    hovertemplate="<b>%{x}</b>: %{y:.2f}% of reports<extra></extra>",
                )
            )
            base_plotly_layout(fig_cqi, height=285)
            fig_cqi.update_yaxes(title_text="Share of Reports (%)", range=[0, max(float(pcts.max()) * 1.22, 10.0)])
            st.plotly_chart(fig_cqi, use_container_width=True)
        else:
            st.info("CQI Bin 0..15 distribution columns are not available in this dataset.")

    # ---- TAB 3: CARRIER LAYER BREAKDOWN ----
    with tab_carrier:
        if "earfcndl" in active_scope_df.columns and not active_scope_df.empty:
            rows_car = []
            for earfcn, sub in active_scope_df.groupby("earfcndl"):
                em = compute_executive_metrics(sub)
                irat_r = float(safe_div(get_col(sub, "IRAT_Events").sum(), get_col(sub, "pmRrcConnEstabSucc").sum(), 100.0))
                rows_car.append({
                    "Carrier": f"EARFCN {int(earfcn)}",
                    "Cells/ROPs": len(sub),
                    "CSSR %": em["sys_cssr_pct"],
                    "Combined HOSR %": em["sys_hosr_pct"],
                    "VoLTE CSSR %": em["sys_volte_cssr_pct"],
                    "Overall DCR %": em["sys_dcr_pct"],
                    "VoLTE DCR %": em["sys_volte_dcr_pct"],
                    "IRAT Offload %": round(irat_r, 2) if pd.notna(irat_r) else 0.0,
                    "Avg CQI": em["sys_avg_cqi"],
                })
            df_car = pd.DataFrame(rows_car)

            fig_car = make_subplots(specs=[[{"secondary_y": True}]])
            for m_col, clr in [("CSSR %", "#38BDF8"), ("Combined HOSR %", "#52C41A"), ("VoLTE CSSR %", "#818CF8")]:
                fig_car.add_trace(
                    go.Bar(name=m_col, x=df_car["Carrier"], y=df_car[m_col], marker_color=clr),
                    secondary_y=False,
                )
            for d_col, clr in [("Overall DCR %", "#FF4D4F"), ("VoLTE DCR %", "#FAAD14"), ("IRAT Offload %", "#F472B6")]:
                fig_car.add_trace(
                    go.Scatter(name=d_col, x=df_car["Carrier"], y=df_car[d_col], mode="lines+markers", line=dict(color=clr, width=2.5), marker=dict(size=8)),
                    secondary_y=True,
                )
            base_plotly_layout(fig_car, height=295)
            fig_car.update_layout(barmode="group")
            fig_car.update_yaxes(title_text="Success Rate (%)", range=[85, 100.5], secondary_y=False)
            fig_car.update_yaxes(title_text="Drop / Offload Rate (%)", secondary_y=True, showgrid=False)
            st.plotly_chart(fig_car, use_container_width=True)
        else:
            st.info("Carrier `earfcndl` column not present.")


# ==============================================================================
# 8. TIER 3 — MASTER DATA GRID & DEEP-DIVE INSPECTOR / AMOS TERMINAL (7 : 5)
# ==============================================================================

st.markdown("<div style='height:6px;'></div>", unsafe_allow_html=True)
col_grid, col_inspector = st.columns([7, 5], gap="medium")

# Sort table so CRITICAL -> DEGRADED -> NORMAL -> STAT_INSIGNIFICANT appear first,
# and pin Pareto-selected cell to Row 0 if cross-filtered!
sev_rank_map = {"CRITICAL": 0, "DEGRADED": 1, "NORMAL": 2, "STAT_INSIGNIFICANT": 3}
grid_source = active_filtered_df.copy()
if grid_source.empty:
    grid_source = active_scope_df.copy()

grid_source["_sev_order"] = grid_source["Severity"].map(sev_rank_map).fillna(9)
grid_source["_pin"] = np.where(
    grid_source["EUtranCell Id"] == st.session_state["selected_cell_id"]
    if st.session_state["selected_cell_id"]
    else False,
    0,
    1,
)
grid_source = (
    grid_source.sort_values(["_pin", "_sev_order", "DCR_Overall_Pct"], ascending=[True, True, False])
    .drop(columns=["_pin", "_sev_order"])
    .reset_index(drop=True)
)

with col_grid:
    st.markdown(
        f"<div class='panel-hdr'>"
        f"<span>🗂️ Master Diagnostic Matrix ({len(active_filtered_df):,} Matching Records — Select Row to Inspect)</span>"
        f"<span class='mono-num' style='font-size:0.74rem;color:#94A3B8;'>Preset: {selected_preset}</span>"
        f"</div>",
        unsafe_allow_html=True,
    )

    display_cols = [
        "Severity",
        "ERBS Id",
        "EUtranCell Id",
        "earfcndl",
        "ROP_Time",
        "Primary_Metric",
        "Isolated_Phase",
        "Root_Cause",
        "Confidence",
        "CSSR_Pct",
        "DCR_Overall_Pct",
        "Combined_HOSR_Pct",
        "VoLTE_DCR_Rel_Pct",
        "Avg_CQI",
        "UL_Interf_Avg_dBm",
        "PRB_Peak_to_Median_dB",
        "IRAT_Offload_Pct",
    ]
    display_cols = [c for c in display_cols if c in grid_source.columns]

    def _style_severity(val: Any) -> str:
        s = str(val)
        if s == "CRITICAL":
            return "background-color: rgba(255, 77, 79, 0.22); color: #FF4D4F; font-weight: 700;"
        if s == "DEGRADED":
            return "background-color: rgba(250, 173, 20, 0.20); color: #FAAD14; font-weight: 700;"
        if s == "NORMAL":
            return "background-color: rgba(82, 196, 26, 0.16); color: #52C41A; font-weight: 600;"
        return "background-color: rgba(100, 116, 139, 0.18); color: #94A3B8;"

    styled_df = grid_source[display_cols].style.map(_style_severity, subset=["Severity"])

    grid_event = st.dataframe(
        styled_df,
        use_container_width=True,
        height=540,
        hide_index=True,
        on_select="rerun",
        selection_mode="single-row",
        column_config={
            "Confidence": st.column_config.NumberColumn("Conf %", format="%d%%"),
            "CSSR_Pct": st.column_config.NumberColumn("CSSR %", format="%.2f%%"),
            "DCR_Overall_Pct": st.column_config.NumberColumn("DCR %", format="%.2f%%"),
            "Combined_HOSR_Pct": st.column_config.NumberColumn("HOSR %", format="%.2f%%"),
            "VoLTE_DCR_Rel_Pct": st.column_config.NumberColumn("VoLTE DCR", format="%.2f%%"),
            "Avg_CQI": st.column_config.NumberColumn("Avg CQI", format="%.2f"),
            "UL_Interf_Avg_dBm": st.column_config.NumberColumn("UL dBm", format="%.1f"),
            "PRB_Peak_to_Median_dB": st.column_config.NumberColumn("PIM ΔdB", format="%.1f"),
            "IRAT_Offload_Pct": st.column_config.NumberColumn("IRAT %", format="%.2f%%"),
        },
        key="master_grid_table",
    )

    if grid_event and grid_event.selection and grid_event.selection.rows:
        sel_idx = grid_event.selection.rows[0]
        if 0 <= sel_idx < len(grid_source):
            picked = grid_source.iloc[sel_idx]
            st.session_state["selected_cell_id"] = str(picked["EUtranCell Id"])
            st.session_state["selected_rop_time"] = str(picked.get("ROP_Time", ""))

# Resolve active inspected row (default to top worst row if nothing clicked yet)
inspected_row: Optional[pd.Series] = None
if st.session_state["selected_cell_id"]:
    cand = grid_source[grid_source["EUtranCell Id"] == st.session_state["selected_cell_id"]]
    if not is_cell_view and st.session_state["selected_rop_time"] and "ROP_Time" in cand.columns:
        rop_cand = cand[cand["ROP_Time"] == st.session_state["selected_rop_time"]]
        if not rop_cand.empty:
            cand = rop_cand
    if cand.empty:
        cand = active_scope_df[active_scope_df["EUtranCell Id"] == st.session_state["selected_cell_id"]]
    if not cand.empty:
        inspected_row = cand.iloc[0]

if inspected_row is None and not grid_source.empty:
    inspected_row = grid_source.iloc[0]

with col_inspector:
    st.markdown(
        "<div class='panel-hdr'>"
        "<span>🔬 Deep-Dive Cell Inspector &amp; AMOS Terminal</span>"
        "</div>",
        unsafe_allow_html=True,
    )

    if inspected_row is None:
        st.info("Select a cell from the Pareto chart or Master Data Grid to inspect.")
    else:
        ins_cell = str(inspected_row.get("EUtranCell Id", ""))
        ins_erbs = str(inspected_row.get("ERBS Id", ""))
        ins_earfcn = str(inspected_row.get("earfcndl", ""))
        ins_rop = str(inspected_row.get("ROP_Time", ""))
        ins_sev = str(inspected_row.get("Severity", "NORMAL"))
        ins_scen = str(inspected_row.get("Scenario", "NONE"))
        ins_conf = int(inspected_row.get("Confidence", 0) or 0)

        st.markdown(
            f"""
            <div class="inspector-box">
                <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:6px;">
                    <div>
                        <span class="mono-num" style="font-size:1.02rem;font-weight:700;color:#38BDF8;">{ins_cell}</span>
                        <span class="mono-num" style="font-size:0.76rem;color:#94A3B8;margin-left:8px;">({ins_erbs} • EARFCN {ins_earfcn})</span>
                    </div>
                    <span class="sev-badge sev-{ins_sev}">{ins_sev}</span>
                </div>
                <div style="font-size:0.8rem;color:#E2E8F0;margin-bottom:4px;">
                    <b>Isolated Phase:</b> <span class="mono-num" style="color:#FAAD14;">{inspected_row.get('Primary_Metric', '')} — {inspected_row.get('Isolated_Phase', '')}</span>
                    &nbsp;|&nbsp; <b>Scenario:</b> <span class="mono-num">{ins_scen}</span> ({ins_conf}%)
                </div>
                <div style="font-size:0.82rem;color:#F8FAFC;background:#0B0F19;border-left:3px solid #38BDF8;padding:6px 10px;border-radius:4px;">
                    <b>Root Cause:</b> {inspected_row.get('Root_Cause', '')}
                </div>
                <div class="evidence-grid">
                    <div class="ev-cell"><div class="ev-lbl">RACH OOC Disc</div><div class="ev-val">{int(inspected_row.get('pmRaFailCbraMsg1DiscOoc', 0) or 0):,}</div></div>
                    <div class="ev-cell"><div class="ev-lbl">Drops (UE Lost)</div><div class="ev-val">{int(inspected_row.get('pmErabRelAbnormalEnbActUeLost', 0) or 0):,}</div></div>
                    <div class="ev-cell"><div class="ev-lbl">Drops (HO / MME)</div><div class="ev-val">{int(inspected_row.get('pmErabRelAbnormalEnbActHo', 0) or 0)} / {int(inspected_row.get('pmErabRelAbnormalMmeAct', 0) or 0)}</div></div>
                    <div class="ev-cell"><div class="ev-lbl">RLC UL Poll Retx</div><div class="ev-val">{int(inspected_row.get('pmRlcPollRetxUl', 0) or 0):,}</div></div>
                    <div class="ev-cell"><div class="ev-lbl">UL Noise Floor</div><div class="ev-val">{float(inspected_row.get('UL_Interf_Avg_dBm', -121.0)):.1f} dBm</div></div>
                    <div class="ev-cell"><div class="ev-lbl">PIM Peak-Med Δ</div><div class="ev-val" style="color:{'#FF4D4F' if float(inspected_row.get('PRB_Peak_to_Median_dB', 0) or 0) > 6.0 else '#52C41A'};">{float(inspected_row.get('PRB_Peak_to_Median_dB', 0.0)):.2f} dB</div></div>
                    <div class="ev-cell"><div class="ev-lbl">Avg CQI / OOR</div><div class="ev-val">{float(inspected_row.get('Avg_CQI', 0) or 0):.2f} / {float(inspected_row.get('CQI_OOR_Pct', 0) or 0):.1f}%</div></div>
                    <div class="ev-cell"><div class="ev-lbl">MIMO Rank-1</div><div class="ev-val">{float(inspected_row.get('Rank1_Ratio_Pct', 0) or 0):.1f}%</div></div>
                    <div class="ev-cell"><div class="ev-lbl">VoLTE PE / Drops</div><div class="ev-val">{int(inspected_row.get('pmErabRelAbnormalEnbActPeQci', 0) or 0)} / {int(inspected_row.get('VoLTE_Active_Drops', 0) or 0)}</div></div>
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

        # ---- 100-PRB INTERFERENCE SPECTRUM CHART ----
        prb_table = prb_cell_scoped if is_cell_view else prb_hourly_all
        if prb_table is not None and not prb_table.empty:
            prb_cols = natural_prb_cols(list(prb_table.columns))
            p_match = prb_table[prb_table["EUtranCell Id"] == ins_cell]
            if not is_cell_view and ins_rop and "ROP_Time" in p_match.columns:
                p_rop = p_match[p_match["ROP_Time"] == ins_rop]
                if not p_rop.empty:
                    p_match = p_rop

            if prb_cols and not p_match.empty:
                p_row = p_match.iloc[0]
                n_hrs = float(p_row.get("ROP_Count", 1.0)) if is_cell_view else 1.0
                prb_dbm = convert_prb_row_to_dbm(p_row, prb_cols, n_hours=n_hrs)
                prb_idx = list(range(1, len(prb_cols) + 1))
                med_dbm = float(np.median(prb_dbm))
                peak_idx = int(np.argmax(prb_dbm)) + 1
                peak_dbm = float(np.max(prb_dbm))
                pim_delta_v = float(p_row.get("PRB_Peak_to_Median_dB", peak_dbm - med_dbm))

                floor_y = min(float(np.min(prb_dbm)) - 1.5, -122.5)
                fig_prb = go.Figure()
                fig_prb.add_trace(
                    go.Scatter(
                        x=prb_idx,
                        y=[floor_y] * len(prb_idx),
                        mode="lines",
                        line=dict(width=0),
                        showlegend=False,
                        hoverinfo="skip",
                    )
                )
                fig_prb.add_trace(
                    go.Scatter(
                        x=prb_idx,
                        y=prb_dbm,
                        mode="lines",
                        fill="tonexty",
                        name="PRB UL Noise (dBm)",
                        line=dict(color="#FF4D4F" if pim_delta_v > 6.0 else "#38BDF8", width=2.0),
                        fillcolor="rgba(255, 77, 79, 0.18)" if pim_delta_v > 6.0 else "rgba(56, 189, 248, 0.15)",
                        hovertemplate="<b>PRB %{x}</b>: %{y:.2f} dBm<extra></extra>",
                    )
                )
                fig_prb.add_hline(
                    y=med_dbm,
                    line_dash="dash",
                    line_color="#94A3B8",
                    annotation_text=f"Median ({med_dbm:.1f} dBm)",
                    annotation_position="bottom right",
                )
                fig_prb.add_hline(
                    y=med_dbm + 6.0,
                    line_dash="dot",
                    line_color="#FAAD14",
                    annotation_text="+6 dB PIM Gate",
                    annotation_position="top right",
                )
                fig_prb.add_annotation(
                    x=peak_idx,
                    y=peak_dbm,
                    text=f"Peak PRB#{peak_idx}: {peak_dbm:.1f} dBm (Δ={pim_delta_v:.1f}dB)",
                    showarrow=True,
                    arrowhead=2,
                    arrowcolor="#FF4D4F",
                    font=dict(size=10, color="#F8FAFC"),
                    bgcolor="#0B0F19",
                    bordercolor="#FF4D4F",
                )
                base_plotly_layout(fig_prb, height=225, title=f"100-PRB UL Interference Spectrum — {ins_cell}")
                fig_prb.update_xaxes(title_text="Physical Resource Block (PRB 1 .. PRB 100)", range=[1, len(prb_cols)])
                fig_prb.update_yaxes(
                    title_text="dBm / PRB",
                    range=[min(float(np.min(prb_dbm)) - 1.5, -122.0), max(peak_dbm + 3.0, -108.0)],
                )
                st.plotly_chart(fig_prb, use_container_width=True)

        # ---- AMOS / MOSHELL REMEDIATION TERMINAL ----
        amos_script_str = generate_amos_script(inspected_row)
        t_hdr_l, t_hdr_r = st.columns([6.5, 3.5], vertical_alignment="center")
        with t_hdr_l:
            st.markdown(
                "<div style='font-family:Space Grotesk,sans-serif;font-size:0.88rem;font-weight:700;color:#38BDF8;'>"
                "💻 AMOS / Moshell Remediation Terminal (One-Click Copy Top-Right)</div>",
                unsafe_allow_html=True,
            )
        with t_hdr_r:
            st.download_button(
                label="💾 Save .mos",
                data=amos_script_str.encode("utf-8"),
                file_name=f"remediate_{ins_cell}.mos",
                mime="text/plain",
                use_container_width=True,
            )

        st.code(amos_script_str, language="bash")
