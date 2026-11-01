#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Antigravity Ericsson LTE & VoLTE Diagnostic PDF Report Engine
Pure ReportLab Platypus implementation (zero headless browser or GTK/C dependencies).
Preserves NOC semantic severity badges, KPI summary matrix, counter evidence,
and copy-ready AMOS/Moshell remediation scripts.
"""

from __future__ import annotations

import io
import os
from datetime import datetime
from typing import Any, Dict, List, Optional

import pandas as pd
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (
    HRFlowable,
    KeepTogether,
    PageBreak,
    Paragraph,
    Preformatted,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)

# NOC Telemetry Semantic Palette
CLR_BG_DARK = colors.HexColor("#0B0F19")
CLR_CARD_DARK = colors.HexColor("#151D2E")
CLR_BORDER = colors.HexColor("#CBD5E1")
CLR_BORDER_DARK = colors.HexColor("#1E293B")
CLR_ACCENT = colors.HexColor("#0284C7")
CLR_CRITICAL = colors.HexColor("#FF4D4F")
CLR_CRITICAL_BG = colors.HexColor("#FFF1F0")
CLR_DEGRADED = colors.HexColor("#D48806")
CLR_DEGRADED_BG = colors.HexColor("#FFFBE6")
CLR_NORMAL = colors.HexColor("#389E0D")
CLR_NORMAL_BG = colors.HexColor("#F6FFED")
CLR_INSIG = colors.HexColor("#64748B")
CLR_INSIG_BG = colors.HexColor("#F1F5F9")
CLR_TEXT_MAIN = colors.HexColor("#0F172A")
CLR_TEXT_MUTED = colors.HexColor("#475569")
CLR_TERM_TEXT = colors.HexColor("#E2E8F0")


def _register_fonts() -> Dict[str, str]:
    """
    Registers Unicode-capable system fonts (Segoe UI / Arial + Consolas) on Windows/Linux
    when available, gracefully falling back to built-in Helvetica / Courier.
    """
    font_map = {
        "sans": "Helvetica",
        "sans_bold": "Helvetica-Bold",
        "mono": "Courier",
        "mono_bold": "Courier-Bold",
    }
    candidates = [
        ("sans", "AntigravitySans", ["C:/Windows/Fonts/segoeui.ttf", "C:/Windows/Fonts/arial.ttf", "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"]),
        ("sans_bold", "AntigravitySansBold", ["C:/Windows/Fonts/segoeuib.ttf", "C:/Windows/Fonts/arialbd.ttf", "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"]),
        ("mono", "AntigravityMono", ["C:/Windows/Fonts/consola.ttf", "C:/Windows/Fonts/cour.ttf", "/usr/share/fonts/truetype/dejavu/DejaVuSansMono.ttf"]),
        ("mono_bold", "AntigravityMonoBold", ["C:/Windows/Fonts/consolab.ttf", "C:/Windows/Fonts/courbd.ttf", "/usr/share/fonts/truetype/dejavu/DejaVuSansMono-Bold.ttf"]),
    ]
    for key, alias, paths in candidates:
        for p in paths:
            if os.path.exists(p):
                try:
                    pdfmetrics.registerFont(TTFont(alias, p))
                    font_map[key] = alias
                    break
                except Exception:
                    continue
    return font_map


def _severity_colors(severity: str) -> tuple[colors.Color, colors.Color]:
    sev = str(severity).upper()
    if sev == "CRITICAL":
        return CLR_CRITICAL, CLR_CRITICAL_BG
    if sev == "DEGRADED":
        return CLR_DEGRADED, CLR_DEGRADED_BG
    if sev == "NORMAL":
        return CLR_NORMAL, CLR_NORMAL_BG
    return CLR_INSIG, CLR_INSIG_BG


def _fmt_num(val: Any, suffix: str = "", digits: int = 2) -> str:
    if val is None or (isinstance(val, float) and pd.isna(val)) or pd.isna(val):
        return "N/A"
    try:
        return f"{float(val):.{digits}f}{suffix}"
    except (TypeError, ValueError):
        return str(val)


def _xml_escape(text: Any) -> str:
    s = "" if text is None else str(text)
    return (
        s.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
    )


def _draw_page_chrome(canvas, doc, meta: Dict[str, Any], fonts: Dict[str, str]) -> None:
    """Renders top telemetry accent bar and bottom page footer on every PDF page."""
    canvas.saveState()
    page_w, page_h = landscape(A4)

    # Top thin telemetry bar
    canvas.setFillColor(CLR_BG_DARK)
    canvas.rect(0, page_h - 8 * mm, page_w, 8 * mm, stroke=0, fill=1)
    canvas.setFillColor(colors.HexColor("#38BDF8"))
    canvas.setFont(fonts["mono_bold"], 7.5)
    canvas.drawString(
        12 * mm,
        page_h - 5.3 * mm,
        f"ANTIGRAVITY LTE RAN TELEMETRY & RCA ENGINE v9.0  |  DATASET: {meta.get('filename', 'Counters.xlsx')}",
    )
    canvas.setFillColor(colors.HexColor("#94A3B8"))
    canvas.setFont(fonts["mono"], 7.5)
    canvas.drawRightString(
        page_w - 12 * mm,
        page_h - 5.3 * mm,
        f"MODE: {meta.get('view_mode', 'Aggregated Cell View')}  |  GENERATED: {meta.get('generated_at', '')}",
    )

    # Bottom footer
    canvas.setStrokeColor(CLR_BORDER)
    canvas.setLineWidth(0.5)
    canvas.line(12 * mm, 10 * mm, page_w - 12 * mm, 10 * mm)
    canvas.setFillColor(CLR_TEXT_MUTED)
    canvas.setFont(fonts["sans"], 7.5)
    canvas.drawString(
        12 * mm,
        6 * mm,
        "Ericsson 3GPP/MOM Standard KPI & Automated AMOS Remediation Report — Rule of Minimum Denominator Enforced",
    )
    canvas.setFont(fonts["mono_bold"], 7.5)
    canvas.drawRightString(page_w - 12 * mm, 6 * mm, f"Page {doc.page}")
    canvas.restoreState()


def generate_diagnostic_pdf(
    df_display: pd.DataFrame,
    exec_metrics: Dict[str, Any],
    meta: Dict[str, Any],
    amos_lookup: Optional[Dict[str, str]] = None,
    selected_cell_id: Optional[str] = None,
    max_table_rows: int = 80,
    max_playbooks: int = 12,
) -> bytes:
    """
    Builds a complete multi-section A4 Landscape PDF report and returns raw PDF bytes.

    Args:
        df_display: Filtered / evaluated DataFrame (Aggregated Cell or Hourly ROP level).
        exec_metrics: Dictionary of Tier-1 Executive KPI summary metrics.
        meta: Metadata dict (filename, date_range, total_erbs, total_cells, total_rops, view_mode).
        amos_lookup: Mapping of EUtranCell Id (or row key) -> generated AMOS script string.
        selected_cell_id: Optional currently inspected cell to prioritize in Deep-Dive section.
        max_table_rows: Cap on rows rendered in the KPI/RCA matrix table for PDF readability.
        max_playbooks: Cap on full AMOS terminal playbooks included in Section 3.
    """
    fonts = _register_fonts()
    buf = io.BytesIO()
    meta = dict(meta or {})
    meta.setdefault("generated_at", datetime.now().strftime("%Y-%m-%d %H:%M:%S"))

    page_w, page_h = landscape(A4)
    margin_lr = 12 * mm
    usable_w = page_w - 2 * margin_lr

    doc = SimpleDocTemplate(
        buf,
        pagesize=landscape(A4),
        leftMargin=margin_lr,
        rightMargin=margin_lr,
        topMargin=13 * mm,
        bottomMargin=14 * mm,
        title=f"LTE RAN Diagnostic Report - {meta.get('filename', 'Counters.xlsx')}",
        author="Antigravity Principal RAN Performance Engine",
    )

    styles = getSampleStyleSheet()
    st_banner_title = ParagraphStyle(
        "BannerTitle",
        parent=styles["Normal"],
        fontName=fonts["sans_bold"],
        fontSize=15,
        leading=19,
        textColor=colors.white,
    )
    st_banner_sub = ParagraphStyle(
        "BannerSub",
        parent=styles["Normal"],
        fontName=fonts["mono"],
        fontSize=8.5,
        leading=12,
        textColor=colors.HexColor("#94A3B8"),
    )
    st_sec_hdr = ParagraphStyle(
        "SectionHdr",
        parent=styles["Normal"],
        fontName=fonts["sans_bold"],
        fontSize=11.5,
        leading=15,
        textColor=CLR_BG_DARK,
        spaceBefore=8,
        spaceAfter=5,
    )
    st_kpi_lbl = ParagraphStyle(
        "KpiLbl",
        parent=styles["Normal"],
        fontName=fonts["sans_bold"],
        fontSize=7.5,
        leading=10,
        textColor=colors.HexColor("#94A3B8"),
    )
    st_kpi_val = ParagraphStyle(
        "KpiVal",
        parent=styles["Normal"],
        fontName=fonts["mono_bold"],
        fontSize=13,
        leading=16,
        textColor=colors.white,
    )
    st_kpi_sub = ParagraphStyle(
        "KpiSub",
        parent=styles["Normal"],
        fontName=fonts["mono"],
        fontSize=7,
        leading=9.5,
        textColor=colors.HexColor("#38BDF8"),
    )
    st_th = ParagraphStyle(
        "TblHeader",
        parent=styles["Normal"],
        fontName=fonts["sans_bold"],
        fontSize=7.2,
        leading=9.2,
        textColor=colors.white,
        alignment=TA_CENTER,
    )
    st_td_mono = ParagraphStyle(
        "TblMono",
        parent=styles["Normal"],
        fontName=fonts["mono"],
        fontSize=6.8,
        leading=8.8,
        textColor=CLR_TEXT_MAIN,
    )
    st_td_sans = ParagraphStyle(
        "TblSans",
        parent=styles["Normal"],
        fontName=fonts["sans"],
        fontSize=6.8,
        leading=8.8,
        textColor=CLR_TEXT_MAIN,
    )
    st_body = ParagraphStyle(
        "BodySmall",
        parent=styles["Normal"],
        fontName=fonts["sans"],
        fontSize=8.2,
        leading=11.5,
        textColor=CLR_TEXT_MAIN,
    )
    st_amos = ParagraphStyle(
        "AmosCode",
        parent=styles["Code"],
        fontName=fonts["mono"],
        fontSize=6.8,
        leading=8.8,
        textColor=CLR_TERM_TEXT,
    )

    story: List[Any] = []

    # -------------------------------------------------------------------------
    # HERO TELEMETRY HEADER BANNER
    # -------------------------------------------------------------------------
    banner_left = Paragraph(
        "ERICSSON LTE &amp; VoLTE RAN PERFORMANCE DIAGNOSTIC REPORT<br/>"
        f"<font size='8.5' color='#38BDF8'>3GPP / Ericsson MOM v9.0 Vectorized KPI, Multi-Domain RCA &amp; Automated AMOS Playbook</font>",
        st_banner_title,
    )
    banner_right = Paragraph(
        f"<b>Source Workbook:</b> {_xml_escape(meta.get('filename', 'Counters.xlsx'))}<br/>"
        f"<b>Date Window:</b> {_xml_escape(meta.get('date_range', 'N/A'))}<br/>"
        f"<b>Network Scope:</b> {meta.get('total_erbs', 0)} ERBS  |  {meta.get('total_cells', 0)} Cells  |  {meta.get('total_rops', 0):,} Cell-Hours",
        st_banner_sub,
    )
    banner_tbl = Table([[banner_left, banner_right]], colWidths=[usable_w * 0.62, usable_w * 0.38])
    banner_tbl.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, -1), CLR_BG_DARK),
                ("BOX", (0, 0), (-1, -1), 1.0, CLR_BORDER_DARK),
                ("LEFTPADDING", (0, 0), (-1, -1), 12),
                ("RIGHTPADDING", (0, 0), (-1, -1), 12),
                ("TOPPADDING", (0, 0), (-1, -1), 10),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 10),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ]
        )
    )
    story.append(banner_tbl)
    story.append(Spacer(1, 4 * mm))

    # -------------------------------------------------------------------------
    # SECTION 1: EXECUTIVE METRIC BAR (6 KPI CARDS)
    # -------------------------------------------------------------------------
    story.append(Paragraph("1. Executive Network Telemetry Summary", st_sec_hdr))

    def _card_cell(title: str, val: str, sub: str, accent_hex: str = "#38BDF8") -> List[Paragraph]:
        sub_style = ParagraphStyle("CardSubDyn", parent=st_kpi_sub, textColor=colors.HexColor(accent_hex))
        return [
            Paragraph(_xml_escape(title), st_kpi_lbl),
            Spacer(1, 1 * mm),
            Paragraph(_xml_escape(val), st_kpi_val),
            Spacer(1, 0.8 * mm),
            Paragraph(_xml_escape(sub), sub_style),
        ]

    dcr_val = exec_metrics.get("sys_dcr_pct", 0.0)
    cssr_val = exec_metrics.get("sys_cssr_pct", 100.0)
    hosr_val = exec_metrics.get("sys_hosr_pct", 100.0)
    vdcr_val = exec_metrics.get("sys_volte_dcr_pct", 0.0)

    cards_row = [
        _card_cell(
            "EVALUATED SCOPE",
            f"{exec_metrics.get('total_evaluated', 0):,}",
            f"Low-Sample Excluded: {exec_metrics.get('stat_insig', 0):,}",
            "#94A3B8",
        ),
        _card_cell(
            "WORST CELLS / ROPS",
            f"{exec_metrics.get('worst_count', 0):,}",
            f"{exec_metrics.get('worst_pct', 0.0):.1f}% Violation Rate",
            "#FF4D4F" if exec_metrics.get("worst_count", 0) > 0 else "#52C41A",
        ),
        _card_cell(
            "SYSTEM CSSR (%)",
            f"{cssr_val:.2f}%",
            f"Delta vs 98.5%: {cssr_val - 98.5:+.2f}%",
            "#FF4D4F" if cssr_val < 97.0 else ("#FAAD14" if cssr_val < 98.5 else "#52C41A"),
        ),
        _card_cell(
            "SYSTEM OVERALL DCR",
            f"{dcr_val:.2f}%",
            "ALERT > 1.0% Baseline" if dcr_val > 1.0 else "Target <= 0.50%",
            "#FF4D4F" if dcr_val > 1.0 else ("#FAAD14" if dcr_val > 0.5 else "#52C41A"),
        ),
        _card_cell(
            "COMBINED HOSR (%)",
            f"{hosr_val:.2f}%",
            "Intra + Inter Mobility",
            "#FF4D4F" if hosr_val < 95.0 else ("#FAAD14" if hosr_val < 98.0 else "#52C41A"),
        ),
        _card_cell(
            "VoLTE QCI=1 DCR",
            f"{vdcr_val:.2f}%",
            f"Avg CQI: {exec_metrics.get('sys_avg_cqi', 0.0):.2f}",
            "#FF4D4F" if vdcr_val > 0.8 else "#52C41A",
        ),
    ]
    col_w = usable_w / 6.0
    kpi_tbl = Table([cards_row], colWidths=[col_w] * 6)
    kpi_tbl.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, -1), CLR_CARD_DARK),
                ("GRID", (0, 0), (-1, -1), 1.5, CLR_BG_DARK),
                ("BOX", (0, 0), (-1, -1), 1.0, CLR_BORDER_DARK),
                ("LEFTPADDING", (0, 0), (-1, -1), 7),
                ("RIGHTPADDING", (0, 0), (-1, -1), 7),
                ("TOPPADDING", (0, 0), (-1, -1), 7),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ]
        )
    )
    story.append(kpi_tbl)
    story.append(Spacer(1, 4 * mm))

    # -------------------------------------------------------------------------
    # SECTION 2: CELL-BY-CELL KPI & ROOT CAUSE MATRIX
    # -------------------------------------------------------------------------
    sev_order = {"CRITICAL": 0, "DEGRADED": 1, "NORMAL": 2, "STAT_INSIGNIFICANT": 3}
    df_sorted = df_display.copy()
    if "Severity" in df_sorted.columns:
        df_sorted["_sev_rank"] = df_sorted["Severity"].map(sev_order).fillna(9)
        sort_cols = ["_sev_rank"]
        if "DCR_Overall_Pct" in df_sorted.columns:
            sort_cols.append("DCR_Overall_Pct")
        df_sorted = df_sorted.sort_values(
            sort_cols, ascending=[True] + [False] * (len(sort_cols) - 1)
        ).drop(columns=["_sev_rank"])

    df_matrix = df_sorted.head(max_table_rows)
    story.append(
        Paragraph(
            f"2. Cell-by-Cell KPI &amp; Root-Cause Diagnostic Matrix (Showing Top {len(df_matrix)} of {len(df_display):,} Records)",
            st_sec_hdr,
        )
    )

    headers = [
        "Severity",
        "ERBS Id",
        "EUtranCell Id",
        "EARFCN",
        "CSSR %",
        "DCR %",
        "HOSR %",
        "VoLTE DCR",
        "Avg CQI",
        "UL dBm / PIM",
        "Primary Metric & Phase",
        "Isolated Root Cause",
        "Conf.",
    ]
    col_widths = [
        19 * mm,  # Severity
        26 * mm,  # ERBS
        29 * mm,  # Cell
        13 * mm,  # EARFCN
        14 * mm,  # CSSR
        13 * mm,  # DCR
        14 * mm,  # HOSR
        15 * mm,  # VoLTE DCR
        13 * mm,  # CQI
        23 * mm,  # UL / PIM
        32 * mm,  # Metric & Phase
        51 * mm,  # Root Cause
        11 * mm,  # Conf
    ]
    # Scale widths proportionally to exact usable_w
    scale = usable_w / sum(col_widths)
    col_widths = [w * scale for w in col_widths]

    table_data: List[List[Any]] = [[Paragraph(_xml_escape(h), st_th) for h in headers]]
    row_styles: List[tuple] = [
        ("BACKGROUND", (0, 0), (-1, 0), CLR_BG_DARK),
        ("GRID", (0, 0), (-1, -1), 0.4, CLR_BORDER),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 3.5),
        ("RIGHTPADDING", (0, 0), (-1, -1), 3.5),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
    ]

    for idx, (_, r) in enumerate(df_matrix.iterrows(), start=1):
        sev = str(r.get("Severity", "NORMAL"))
        fg_clr, bg_clr = _severity_colors(sev)
        sev_style = ParagraphStyle(
            f"Sev_{idx}",
            parent=st_td_mono,
            fontName=fonts["mono_bold"],
            textColor=fg_clr,
            alignment=TA_CENTER,
        )
        ul_str = f"{_fmt_num(r.get('UL_Interf_Avg_dBm'), '', 1)} / Δ{_fmt_num(r.get('PRB_Peak_to_Median_dB'), 'dB', 1)}"
        phase_str = f"{r.get('Primary_Metric', '-')} ({r.get('Isolated_Phase', '-')})"
        conf_val = r.get("Confidence", "-")
        conf_str = f"{int(conf_val)}%" if isinstance(conf_val, (int, float)) and pd.notna(conf_val) and conf_val > 0 else "-"

        row_cells = [
            Paragraph(_xml_escape(sev[:10]), sev_style),
            Paragraph(_xml_escape(r.get("ERBS Id", "")), st_td_mono),
            Paragraph(_xml_escape(r.get("EUtranCell Id", "")), st_td_mono),
            Paragraph(_xml_escape(str(r.get("earfcndl", ""))), st_td_mono),
            Paragraph(_fmt_num(r.get("CSSR_Pct"), "%"), st_td_mono),
            Paragraph(_fmt_num(r.get("DCR_Overall_Pct"), "%"), st_td_mono),
            Paragraph(_fmt_num(r.get("Combined_HOSR_Pct"), "%"), st_td_mono),
            Paragraph(_fmt_num(r.get("VoLTE_DCR_Rel_Pct"), "%"), st_td_mono),
            Paragraph(_fmt_num(r.get("Avg_CQI"), ""), st_td_mono),
            Paragraph(_xml_escape(ul_str), st_td_mono),
            Paragraph(_xml_escape(phase_str), st_td_sans),
            Paragraph(_xml_escape(r.get("Root_Cause", "Healthy / Within Baseline")), st_td_sans),
            Paragraph(_xml_escape(conf_str), st_td_mono),
        ]
        table_data.append(row_cells)
        row_bg = colors.white if idx % 2 == 1 else colors.HexColor("#F8FAFC")
        row_styles.append(("BACKGROUND", (1, idx), (-1, idx), row_bg))
        row_styles.append(("BACKGROUND", (0, idx), (0, idx), bg_clr))

    matrix_tbl = Table(table_data, colWidths=col_widths, repeatRows=1)
    matrix_tbl.setStyle(TableStyle(row_styles))
    story.append(matrix_tbl)

    # -------------------------------------------------------------------------
    # SECTION 3: DEEP-DIVE EVIDENCE & AUTOMATED AMOS REMEDIATION PLAYBOOKS
    # -------------------------------------------------------------------------
    degraded_df = df_sorted[df_sorted["Severity"].isin(["CRITICAL", "DEGRADED"])].copy() if "Severity" in df_sorted.columns else pd.DataFrame()
    if selected_cell_id and not degraded_df.empty and "EUtranCell Id" in degraded_df.columns:
        sel_rows = degraded_df[degraded_df["EUtranCell Id"] == selected_cell_id]
        other_rows = degraded_df[degraded_df["EUtranCell Id"] != selected_cell_id]
        degraded_df = pd.concat([sel_rows, other_rows], ignore_index=True)

    playbook_df = degraded_df.head(max_playbooks)
    if not playbook_df.empty:
        story.append(PageBreak())
        story.append(
            Paragraph(
                f"3. Deep-Dive Counter Evidence &amp; Automated AMOS / Moshell Playbooks (Top {len(playbook_df)} Worst Cells)",
                st_sec_hdr,
            )
        )
        story.append(
            Paragraph(
                "Operational Safety Protocol: Every playbook includes Pre-Execution Backup (<b>cvms</b>), Online-Safe dynamic tuning commands, and commented Maintenance-Window (<b>bl/deb</b>) commands.",
                st_body,
            )
        )
        story.append(Spacer(1, 3 * mm))

        for _, r in playbook_df.iterrows():
            cell_name = str(r.get("EUtranCell Id", "Unknown_Cell"))
            erbs_name = str(r.get("ERBS Id", "Unknown_eNB"))
            sev = str(r.get("Severity", "DEGRADED"))
            fg_clr, bg_clr = _severity_colors(sev)

            hdr_p = Paragraph(
                f"<b>[{_xml_escape(sev)}] Cell: {_xml_escape(cell_name)}</b>  "
                f"(eNodeB: {_xml_escape(erbs_name)}  |  EARFCN: {_xml_escape(r.get('earfcndl', ''))}  |  "
                f"Scenario: {_xml_escape(r.get('Scenario', 'N/A'))}  |  Confidence: {_xml_escape(r.get('Confidence', 90))}%)",
                ParagraphStyle("CellCardHdr", parent=st_body, fontName=fonts["mono_bold"], textColor=fg_clr),
            )
            ev_p = Paragraph(
                f"<b>Primary Metric / Phase:</b> {_xml_escape(r.get('Primary_Metric', ''))} — {_xml_escape(r.get('Isolated_Phase', ''))}<br/>"
                f"<b>Isolated Root Cause:</b> {_xml_escape(r.get('Root_Cause', ''))}<br/>"
                f"<b>Diagnostic Evidence:</b> {_xml_escape(r.get('Diagnostic_Evidence', ''))}",
                st_body,
            )

            amos_text = ""
            if amos_lookup and cell_name in amos_lookup:
                amos_text = amos_lookup[cell_name]
            elif "AMOS_Script" in r and pd.notna(r["AMOS_Script"]):
                amos_text = str(r["AMOS_Script"])

            amos_box = Table([[Preformatted(amos_text.strip(), st_amos)]], colWidths=[usable_w - 6 * mm])
            amos_box.setStyle(
                TableStyle(
                    [
                        ("BACKGROUND", (0, 0), (-1, -1), CLR_BG_DARK),
                        ("BOX", (0, 0), (-1, -1), 0.8, CLR_BORDER_DARK),
                        ("LEFTPADDING", (0, 0), (-1, -1), 8),
                        ("RIGHTPADDING", (0, 0), (-1, -1), 8),
                        ("TOPPADDING", (0, 0), (-1, -1), 6),
                        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
                    ]
                )
            )

            card_outer = Table(
                [[hdr_p], [ev_p], [amos_box]],
                colWidths=[usable_w],
            )
            card_outer.setStyle(
                TableStyle(
                    [
                        ("BACKGROUND", (0, 0), (0, 0), bg_clr),
                        ("BACKGROUND", (0, 1), (0, 2), colors.white),
                        ("BOX", (0, 0), (-1, -1), 0.8, fg_clr),
                        ("LINEBELOW", (0, 0), (0, 0), 0.6, fg_clr),
                        ("LEFTPADDING", (0, 0), (-1, -1), 8),
                        ("RIGHTPADDING", (0, 0), (-1, -1), 8),
                        ("TOPPADDING", (0, 0), (-1, -1), 5),
                        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
                    ]
                )
            )
            story.append(KeepTogether([card_outer, Spacer(1, 3.5 * mm)]))

    doc.build(
        story,
        onFirstPage=lambda c, d: _draw_page_chrome(c, d, meta, fonts),
        onLaterPages=lambda c, d: _draw_page_chrome(c, d, meta, fonts),
    )
    return buf.getvalue()
