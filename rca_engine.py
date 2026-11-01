#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Ericsson LTE & VoLTE RAN Troubleshooting, ETL & Vectorized RCA Engine (v9.0)
Principal RAN Performance & Full-Stack Python Architecture

Coordinates:
  - SAP BusinessObjects (BO WebI / ENIQ) 4-Sheet ETL (`Data`, `VoLTE`, `CqiDistr`, `InterferencePwrPrb`)
  - Backward compatibility with `Counters_Organized.xlsx` (3-Sheet) & flat `.csv` dumps
  - 3GPP / Ericsson MOM v9.0 Vectorized KPI Formulas & Rule of Minimum Denominator
  - 7 Multi-Domain Root Cause Analysis (RCA) Decision Trees (MBB, Radio/PIM, IRAT, VoLTE QCI=1 & SRVCC)
  - Automated Ericsson Baseband (ERS) / DU (CPP) AMOS / Moshell Playbook Generator
"""

from __future__ import annotations

import argparse
import io
import json
import os
import re
import sys
import warnings
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple, Union

import numpy as np
import pandas as pd

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

# ==============================================================================
# 1. CONSTANTS, KEYS & ERICSSON RF WEIGHTS
# ==============================================================================

KEYS: List[str] = [
    "OSS Id",
    "ERBS Id",
    "EUtranCell Id",
    "Cell Id",
    "earfcndl",
    "Date",
    "Hour",
]

CELL_KEYS: List[str] = [
    "OSS Id",
    "ERBS Id",
    "EUtranCell Id",
    "Cell Id",
    "earfcndl",
]

# Standard Ericsson 16-bin UL Interference Weights in dBm (Vector Index 0..15)
UL_WEIGHTS_DBM: Dict[int, float] = {
    0: -121.0,
    1: -120.0,
    2: -119.0,
    3: -118.0,
    4: -117.0,
    5: -116.0,
    6: -115.0,
    7: -114.0,
    8: -113.0,
    9: -112.0,
    10: -110.0,
    11: -106.0,
    12: -102.0,
    13: -98.0,
    14: -94.0,
    15: -90.0,
}

# Linear milliwatt weights (10^(dBm/10)) per Ericsson MOM RF physics
UL_WEIGHTS_MW: Dict[int, float] = {k: 10.0 ** (v / 10.0) for k, v in UL_WEIGHTS_DBM.items()}

# Rule of Minimum Denominator Thresholds (Skill v9.0)
MIN_SAMPLES: Dict[str, int] = {
    "CSSR": 30,
    "DCR": 30,
    "Intra_HO": 20,
    "Inter_HO": 20,
    "CQI": 100,
    "IRAT": 30,
    "VoLTE_CSSR": 10,
    "VoLTE_DCR": 10,
}

# KPI Baselines & Thresholds
KPI_BASELINES: Dict[str, Dict[str, float]] = {
    "CSSR_Pct": {"target": 98.5, "degraded": 97.0, "critical": 95.0, "dir": "low"},
    "DCR_Overall_Pct": {"target": 0.5, "degraded": 1.0, "critical": 2.5, "dir": "high"},
    "DCR_eNB_Pct": {"target": 0.3, "degraded": 0.7, "critical": 1.5, "dir": "high"},
    "Intra_Freq_HOSR_Pct": {"target": 98.5, "degraded": 96.0, "critical": 92.0, "dir": "low"},
    "Inter_Freq_HOSR_Pct": {"target": 97.0, "degraded": 94.0, "critical": 88.0, "dir": "low"},
    "Combined_HOSR_Pct": {"target": 98.0, "degraded": 95.0, "critical": 90.0, "dir": "low"},
    "Avg_CQI": {"target": 9.0, "degraded": 7.5, "critical": 6.0, "dir": "low"},
    "CQI_OOR_Pct": {"target": 2.0, "degraded": 5.0, "critical": 10.0, "dir": "high"},
    "IRAT_Offload_Pct": {"target": 2.0, "degraded": 5.0, "critical": 10.0, "dir": "high"},
    "VoLTE_CSSR_RAN_Pct": {"target": 99.0, "degraded": 97.5, "critical": 95.0, "dir": "low"},
    "VoLTE_DCR_Rel_Pct": {"target": 0.3, "degraded": 0.8, "critical": 2.0, "dir": "high"},
}


# ==============================================================================
# 2. VECTORIZED HELPERS & SAP BO SHEET PARSER
# ==============================================================================

def safe_div(
    num: Union[pd.Series, np.ndarray, float, int],
    den: Union[pd.Series, np.ndarray, float, int],
    scale: float = 1.0,
    fill_na: Optional[float] = np.nan,
) -> Union[pd.Series, float]:
    """Vectorized division protecting against ZeroDivisionError and +/-Inf.
    Supports both scalar-scalar division (returning float) and vectorized Series/array.
    """
    # 1. Xử lý trường hợp cả hai đầu vào đều là số đơn lẻ (scalar từ .sum(), float, int)
    is_vector = (
        isinstance(num, (pd.Series, pd.Index))
        or isinstance(den, (pd.Series, pd.Index))
        or (isinstance(num, np.ndarray) and num.ndim > 0)
        or (isinstance(den, np.ndarray) and den.ndim > 0)
    )
    if not is_vector:
        try:
            n = float(num)
            d = float(den)
            if d == 0.0 or np.isnan(d) or np.isnan(n):
                return fill_na if fill_na is not None else np.nan
            val = (n / d) * scale
            if np.isinf(val):
                return fill_na if fill_na is not None else np.nan
            return float(val)
        except (ValueError, TypeError, ZeroDivisionError):
            return fill_na if fill_na is not None else np.nan

    # 2. Xử lý trường hợp vector hóa (pd.Series / np.ndarray)
    if isinstance(den, pd.Series):
        den_clean = pd.to_numeric(den, errors="coerce").replace(0, np.nan)
        idx = den.index
    elif isinstance(num, pd.Series):
        den_arr = np.asarray(den, dtype=float)
        den_clean = np.where(den_arr == 0, np.nan, den_arr)
        idx = num.index
    else:
        den_arr = np.asarray(den, dtype=float)
        den_clean = np.where(den_arr == 0, np.nan, den_arr)
        idx = getattr(num, "index", None) or getattr(den, "index", None)

    res = (num / den_clean) * scale
    if not isinstance(res, pd.Series):
        res = pd.Series(res, index=idx)

    res = res.replace([np.inf, -np.inf], np.nan)
    if fill_na is not None and not np.isnan(fill_na):
        res = res.fillna(fill_na)
    return res


def get_col(df: pd.DataFrame, col_name: str, default: float = 0.0) -> pd.Series:
    """Safely fetches a numeric counter column or returns a constant Series."""
    if col_name in df.columns:
        return pd.to_numeric(df[col_name], errors="coerce").fillna(default)
    return pd.Series(default, index=df.index, dtype=float)


def natural_prb_cols(columns: List[str]) -> List[str]:
    """Returns PRB interference columns sorted naturally: Prb1, Prb2, ..., Prb100."""
    prb_map: Dict[int, str] = {}
    for c in columns:
        m = re.match(r"^pmRadioRecInterferencePwrPrb(\d+)$", str(c).strip())
        if m:
            idx = int(m.group(1))
            if 1 <= idx <= 100:
                prb_map[idx] = str(c).strip()
    return [prb_map[k] for k in sorted(prb_map.keys())]


def read_bo_sheet(xls: pd.ExcelFile, sheet_name: str) -> pd.DataFrame:
    """
    Scans the first 10 rows of a SAP BusinessObjects Excel sheet to locate the true
    header row containing 'ERBS Id' or 'EUtranCell Id', then strips all 'Unnamed:' columns.
    """
    raw_preview = pd.read_excel(xls, sheet_name=sheet_name, nrows=10, header=None)
    hdr_idx = 0
    for idx, row in raw_preview.iterrows():
        vals = {str(v).strip() for v in row.dropna().values}
        if "ERBS Id" in vals or "EUtranCell Id" in vals:
            hdr_idx = int(idx)
            break

    df = pd.read_excel(xls, sheet_name=sheet_name, header=hdr_idx)
    df.columns = [str(c).strip() for c in df.columns]
    clean_cols = [c for c in df.columns if not c.startswith("Unnamed:")]
    df = df[clean_cols].copy()

    # Drop completely empty trailing BO footer rows if any
    key_check = [k for k in ["ERBS Id", "EUtranCell Id"] if k in df.columns]
    if key_check:
        df = df.dropna(subset=key_check, how="all").reset_index(drop=True)
    return df


def _standardize_keys(df: pd.DataFrame, keys: List[str]) -> pd.DataFrame:
    """Normalizes key data types so multi-sheet joins on KEYS never fail on dtype mismatch."""
    out = df.copy()
    for k in keys:
        if k not in out.columns:
            continue
        if k in ("OSS Id", "ERBS Id", "EUtranCell Id"):
            out[k] = out[k].astype(str).str.strip()
        elif k in ("Cell Id", "earfcndl", "Hour"):
            out[k] = pd.to_numeric(out[k], errors="coerce").fillna(0).astype(int)
        elif k == "Date":
            dt = pd.to_datetime(out[k], errors="coerce")
            out[k] = dt.dt.strftime("%Y-%m-%d").fillna(out[k].astype(str).str.strip())
    return out


def _make_rop_time(df: pd.DataFrame) -> pd.Series:
    """Creates standardized 'YYYY-MM-DD HH:00' ROP_Time label from Date and Hour."""
    if "Date" in df.columns and "Hour" in df.columns:
        dt_str = (
            pd.to_datetime(df["Date"], errors="coerce")
            .dt.strftime("%Y-%m-%d")
            .fillna(df["Date"].astype(str).str.slice(0, 10))
        )
        hr_str = pd.to_numeric(df["Hour"], errors="coerce").fillna(0).astype(int).astype(str).str.zfill(2)
        return dt_str + " " + hr_str + ":00"
    return pd.Series("Aggregated", index=df.index)


# ==============================================================================
# 3. MULTI-SHEET WORKBOOK ETL PIPELINE
# ==============================================================================

def load_and_process_Counters(
    file_source: Union[str, Path, io.BytesIO, Any],
    source_name: Optional[str] = None,
) -> Dict[str, Any]:
    """
    Master ETL entry point for:
      1) Raw 4-Sheet SAP BO `Counters.xlsx` (`Data`, `VoLTE`, `CqiDistr`, `InterferencePwrPrb`)
      2) Pre-organized 3-Sheet `Counters_Organized.xlsx` (`Cell_Counters`, `Radio_Quality_Pivoted`, `InterferencePwrPrb`)
      3) Single-sheet Excel or CSV counter exports

    Returns a dictionary containing:
      - `hourly_df`: Evaluated Cell-Hour DataFrame with KPIs & RCA
      - `cell_df`: Evaluated Aggregated Cell DataFrame with KPIs & RCA
      - `prb_hourly_df`: Sorted 100-PRB spectrum table at Cell-Hour level
      - `prb_cell_df`: Aggregated 100-PRB spectrum table at Cell level
      - `meta`: Dataset summary metadata
    """
    if source_name is None:
        if isinstance(file_source, (str, Path)):
            source_name = os.path.basename(str(file_source))
        else:
            source_name = getattr(file_source, "name", "Uploaded_Counters.xlsx")

    ext = os.path.splitext(source_name)[1].lower()

    if ext == ".csv":
        if hasattr(file_source, "seek"):
            file_source.seek(0)
        master_raw = pd.read_csv(file_source)
        master_raw.columns = [str(c).strip() for c in master_raw.columns]
        master_raw = master_raw[[c for c in master_raw.columns if not c.startswith("Unnamed:")]]
        keys_present = [k for k in KEYS if k in master_raw.columns]
        master_raw = _standardize_keys(master_raw, keys_present)
        prb_cols = natural_prb_cols(list(master_raw.columns))
        prb_hourly = None
        if prb_cols:
            prb_hourly = _compute_prb_metrics(master_raw[keys_present + prb_cols].copy(), prb_cols)
            if "PRB_Peak_to_Median_dB" not in master_raw.columns:
                master_raw["PRB_Peak_to_Median_dB"] = prb_hourly["PRB_Peak_to_Median_dB"]
                master_raw["PRB_Max_dBm"] = prb_hourly["PRB_Max_dBm"]
    else:
        if hasattr(file_source, "seek"):
            file_source.seek(0)
        xls = pd.ExcelFile(file_source)
        sheet_names = set(xls.sheet_names)

        # ------------------------------------------------------------------
        # CASE A: Raw 4-Sheet SAP BusinessObjects `Counters.xlsx`
        # ------------------------------------------------------------------
        if {"Data", "VoLTE", "CqiDistr"}.issubset(sheet_names):
            df_data = read_bo_sheet(xls, "Data")
            df_volte = read_bo_sheet(xls, "VoLTE")
            df_cqi = read_bo_sheet(xls, "CqiDistr")
            df_prb_raw = read_bo_sheet(xls, "InterferencePwrPrb") if "InterferencePwrPrb" in sheet_names else None

            keys_present = [k for k in KEYS if k in df_data.columns]
            df_data = _standardize_keys(df_data, keys_present)
            df_volte = _standardize_keys(df_volte, [k for k in KEYS if k in df_volte.columns])
            df_cqi = _standardize_keys(df_cqi, [k for k in KEYS if k in df_cqi.columns])

            # Deduplicate safety on Data & VoLTE per KEYS
            df_data = df_data.drop_duplicates(subset=keys_present).reset_index(drop=True)
            df_volte_dedup = df_volte.drop_duplicates(subset=keys_present).reset_index(drop=True)

            # Avoid duplicate non-key columns (e.g. SRVCC counters appearing in Data)
            volte_new_cols = [c for c in df_volte_dedup.columns if c not in df_data.columns]
            master_raw = pd.merge(
                df_data,
                df_volte_dedup[keys_present + volte_new_cols],
                on=keys_present,
                how="left",
            )
            # Coalesce any shared counter columns if Data had NaNs
            shared_volte_cols = [c for c in df_volte_dedup.columns if c in df_data.columns and c not in keys_present]
            if shared_volte_cols:
                v_idx = df_volte_dedup.set_index(keys_present)
                m_idx = master_raw.set_index(keys_present)
                for sc in shared_volte_cols:
                    m_idx[sc] = m_idx[sc].fillna(v_idx[sc])
                master_raw = m_idx.reset_index()

            # Sheet 3: Pivot & Vectorized Radio Quality Aggregation (CqiDistr)
            radio_full = _pivot_cqi_sheet(df_cqi, keys_present)
            master_raw = pd.merge(master_raw, radio_full, on=keys_present, how="left")

            # Sheet 4: Natural Sort PRB1..PRB100 & PIM Signature
            prb_hourly = None
            if df_prb_raw is not None and not df_prb_raw.empty:
                df_prb_raw = _standardize_keys(df_prb_raw, [k for k in KEYS if k in df_prb_raw.columns])
                df_prb_raw = df_prb_raw.drop_duplicates(subset=keys_present).reset_index(drop=True)
                prb_cols = natural_prb_cols(list(df_prb_raw.columns))
                if prb_cols:
                    prb_hourly = _compute_prb_metrics(df_prb_raw[keys_present + prb_cols].copy(), prb_cols, n_hours=1.0)
                    master_raw = pd.merge(
                        master_raw,
                        prb_hourly[keys_present + ["PRB_Peak_to_Median_dB", "PRB_Max_dBm", "PRB_Mean_dBm"]],
                        on=keys_present,
                        how="left",
                    )

        # ------------------------------------------------------------------
        # CASE B: Pre-Organized 3-Sheet `Counters_Organized.xlsx`
        # ------------------------------------------------------------------
        elif {"Cell_Counters", "Radio_Quality_Pivoted"}.issubset(sheet_names):
            df_cell = read_bo_sheet(xls, "Cell_Counters")
            df_radio = read_bo_sheet(xls, "Radio_Quality_Pivoted")
            df_prb_raw = read_bo_sheet(xls, "InterferencePwrPrb") if "InterferencePwrPrb" in sheet_names else None

            keys_present = [k for k in KEYS if k in df_cell.columns]
            df_cell = _standardize_keys(df_cell, keys_present)
            df_radio = _standardize_keys(df_radio, [k for k in KEYS if k in df_radio.columns])

            radio_cols = [c for c in df_radio.columns if c not in df_cell.columns]
            master_raw = pd.merge(df_cell, df_radio[keys_present + radio_cols], on=keys_present, how="left")

            prb_hourly = None
            if df_prb_raw is not None and not df_prb_raw.empty:
                df_prb_raw = _standardize_keys(df_prb_raw, [k for k in KEYS if k in df_prb_raw.columns])
                prb_cols = natural_prb_cols(list(df_prb_raw.columns))
                prb_hourly = _compute_prb_metrics(df_prb_raw, prb_cols, n_hours=1.0) if prb_cols else df_prb_raw
                extra_cols = [c for c in ["PRB_Peak_to_Median_dB", "PRB_Max_dBm", "PRB_Mean_dBm"] if c in prb_hourly.columns and c not in master_raw.columns]
                if extra_cols:
                    master_raw = pd.merge(master_raw, prb_hourly[keys_present + extra_cols], on=keys_present, how="left")

        # ------------------------------------------------------------------
        # CASE C: Single-Sheet Excel Export
        # ------------------------------------------------------------------
        else:
            first_sheet = xls.sheet_names[0]
            master_raw = read_bo_sheet(xls, first_sheet)
            keys_present = [k for k in KEYS if k in master_raw.columns]
            master_raw = _standardize_keys(master_raw, keys_present)
            prb_cols = natural_prb_cols(list(master_raw.columns))
            prb_hourly = _compute_prb_metrics(master_raw[keys_present + prb_cols].copy(), prb_cols, n_hours=1.0) if prb_cols else None

    if keys_present:
        master_raw = master_raw.sort_values(keys_present).reset_index(drop=True)
    master_raw["ROP_Time"] = _make_rop_time(master_raw)
    if prb_hourly is not None:
        if keys_present:
            prb_hourly = prb_hourly.sort_values(keys_present).reset_index(drop=True)
        prb_hourly["ROP_Time"] = _make_rop_time(prb_hourly)

    # Evaluate Hourly ROP DataFrame
    hourly_evaluated = compute_kpis_and_rca(master_raw, is_aggregated=False)

    # Aggregate raw counters to Cell Level & Evaluate Cell DataFrame
    cell_raw, prb_cell = aggregate_to_cell_level(master_raw, prb_hourly)
    cell_evaluated = compute_kpis_and_rca(cell_raw, is_aggregated=True)

    dates = sorted(master_raw["Date"].astype(str).unique()) if "Date" in master_raw.columns else []
    date_range = f"{dates[0]} -> {dates[-1]}" if len(dates) > 1 else (dates[0] if dates else "N/A")

    meta = {
        "filename": source_name,
        "date_range": date_range,
        "dates": dates,
        "total_erbs": int(master_raw["ERBS Id"].nunique()) if "ERBS Id" in master_raw.columns else 0,
        "total_cells": int(master_raw["EUtranCell Id"].nunique()) if "EUtranCell Id" in master_raw.columns else len(cell_evaluated),
        "total_rops": int(len(hourly_evaluated)),
        "carriers": sorted(master_raw["earfcndl"].dropna().astype(int).unique().tolist()) if "earfcndl" in master_raw.columns else [],
    }

    return {
        "hourly_df": hourly_evaluated,
        "cell_df": cell_evaluated,
        "prb_hourly_df": prb_hourly,
        "prb_cell_df": prb_cell,
        "meta": meta,
    }


def _pivot_cqi_sheet(df_cqi: pd.DataFrame, keys: List[str]) -> pd.DataFrame:
    """
    Transforms SAP BO `CqiDistr` (16 Vector Index bins per Cell-Hour) into a single
    Cell-Hour row containing CQI_Bin_0..15, Avg_CQI (bins 1..15), CQI_OOR_Pct (bin 0),
    Low_CQI_Ratio_Pct (bins 1..6), Rank1_Ratio_Pct, Hom_256QAM_Share_Pct, and UL_Interf_Avg_dBm.
    """
    df_c = df_cqi.copy()
    df_c["Vector Index"] = pd.to_numeric(df_c["Vector Index"], errors="coerce").fillna(0).astype(int)

    cqi_std = get_col(df_c, "pmRadioUeRepCqiDistr")
    cqi_hom = get_col(df_c, "pmRadioHomUeRepCqiDistr")
    rank_distr = get_col(df_c, "pmRadioUeRepRankDistr")
    ul_samples = get_col(df_c, "pmRadioRecInterferencePwr")

    df_c["CQI_Unified"] = cqi_std + cqi_hom
    df_c["CQI_Weighted"] = np.where((df_c["Vector Index"] >= 1) & (df_c["Vector Index"] <= 15), df_c["Vector Index"] * df_c["CQI_Unified"], 0.0)
    df_c["CQI_Valid_Samples"] = np.where((df_c["Vector Index"] >= 1) & (df_c["Vector Index"] <= 15), df_c["CQI_Unified"], 0.0)
    df_c["CQI_Low_Samples"] = np.where((df_c["Vector Index"] >= 1) & (df_c["Vector Index"] <= 6), df_c["CQI_Unified"], 0.0)
    df_c["CQI_OOR_Samples"] = np.where(df_c["Vector Index"] == 0, df_c["CQI_Unified"], 0.0)
    df_c["Hom_256QAM_Samples"] = cqi_hom
    df_c["Rank1_Samples"] = np.where(df_c["Vector Index"] == 0, rank_distr, 0.0)
    df_c["Rank_Total_Samples"] = np.where((df_c["Vector Index"] >= 0) & (df_c["Vector Index"] <= 3), rank_distr, 0.0)

    # Both Ericsson dBm weight accumulator & linear mW accumulator
    df_c["UL_Weighted_dBm"] = df_c["Vector Index"].map(UL_WEIGHTS_DBM).fillna(-121.0) * ul_samples
    df_c["UL_Weighted_mW"] = df_c["Vector Index"].map(UL_WEIGHTS_MW).fillna(10.0 ** (-12.1)) * ul_samples
    df_c["UL_Interf_Samples"] = ul_samples

    radio_agg = df_c.groupby(keys, as_index=False).agg(
        CQI_Weighted_Sum=("CQI_Weighted", "sum"),
        CQI_Valid_Samples=("CQI_Valid_Samples", "sum"),
        CQI_Low_Samples=("CQI_Low_Samples", "sum"),
        CQI_OOR_Samples=("CQI_OOR_Samples", "sum"),
        Hom_256QAM_Samples=("Hom_256QAM_Samples", "sum"),
        Rank1_Samples=("Rank1_Samples", "sum"),
        Rank_Total_Samples=("Rank_Total_Samples", "sum"),
        UL_Weighted_dBm_Sum=("UL_Weighted_dBm", "sum"),
        UL_Weighted_mW_Sum=("UL_Weighted_mW", "sum"),
        UL_Interf_Samples=("UL_Interf_Samples", "sum"),
    )

    cqi_piv = (
        df_c[df_c["Vector Index"].between(0, 15)]
        .pivot_table(index=keys, columns="Vector Index", values="CQI_Unified", aggfunc="sum", fill_value=0.0)
        .reset_index()
    )
    for b in range(16):
        if b not in cqi_piv.columns:
            cqi_piv[b] = 0.0
    cqi_piv = cqi_piv[keys + list(range(16))]
    cqi_piv.columns = keys + [f"CQI_Bin_{b}" for b in range(16)]

    radio_full = pd.merge(radio_agg, cqi_piv, on=keys, how="left")
    return _finalize_radio_metrics(radio_full)


def _finalize_radio_metrics(df: pd.DataFrame) -> pd.DataFrame:
    """Computes derived radio quality KPIs from summed bin/accumulator columns."""
    out = df.copy()
    if "CQI_Valid_Samples" not in out.columns and "CQI_Bin_1" in out.columns:
        bins_1_15 = [f"CQI_Bin_{i}" for i in range(1, 16) if f"CQI_Bin_{i}" in out.columns]
        bins_1_6 = [f"CQI_Bin_{i}" for i in range(1, 7) if f"CQI_Bin_{i}" in out.columns]
        out["CQI_Valid_Samples"] = out[bins_1_15].sum(axis=1)
        out["CQI_Low_Samples"] = out[bins_1_6].sum(axis=1)
        out["CQI_OOR_Samples"] = get_col(out, "CQI_Bin_0")
        out["CQI_Weighted_Sum"] = sum(i * get_col(out, f"CQI_Bin_{i}") for i in range(1, 16))

    if "CQI_Valid_Samples" in out.columns:
        total_cqi = get_col(out, "CQI_Valid_Samples") + get_col(out, "CQI_OOR_Samples")
        out["Avg_CQI"] = safe_div(get_col(out, "CQI_Weighted_Sum"), get_col(out, "CQI_Valid_Samples")).round(2)
        out["Low_CQI_Ratio_Pct"] = safe_div(get_col(out, "CQI_Low_Samples"), get_col(out, "CQI_Valid_Samples"), 100.0).round(2)
        out["CQI_OOR_Pct"] = safe_div(get_col(out, "CQI_OOR_Samples"), total_cqi, 100.0).round(2)
        if "Hom_256QAM_Samples" in out.columns:
            out["Hom_256QAM_Share_Pct"] = safe_div(get_col(out, "Hom_256QAM_Samples"), total_cqi, 100.0).round(2)

    if "Rank_Total_Samples" in out.columns:
        out["Rank1_Ratio_Pct"] = safe_div(get_col(out, "Rank1_Samples"), get_col(out, "Rank_Total_Samples"), 100.0).round(2)

    if "UL_Interf_Samples" in out.columns and "UL_Weighted_mW_Sum" in out.columns:
        ul_samp = get_col(out, "UL_Interf_Samples")
        mean_mw = safe_div(get_col(out, "UL_Weighted_mW_Sum"), ul_samp)
        with np.errstate(divide="ignore", invalid="ignore"):
            # True RF power average in linear mW converted back to dBm (Ericsson MOM Reference 1.5)
            ul_linear_dbm = np.where((ul_samp > 0) & (mean_mw > 0), 10.0 * np.log10(mean_mw), -121.0)
        out["UL_Interf_Avg_dBm"] = np.round(ul_linear_dbm, 2)
        out["UL_Interf_Arith_dBm"] = np.where(
            ul_samp > 0,
            np.round(get_col(out, "UL_Weighted_dBm_Sum") / ul_samp.replace(0, np.nan), 2),
            -121.0,
        )
    return out


def _compute_prb_metrics(df_prb: pd.DataFrame, prb_cols: List[str], n_hours: Union[float, pd.Series] = 1.0) -> pd.DataFrame:
    """
    Computes PIM Peak-to-Median (dB) and calibrates PRB1..100 raw Ericsson MOM accumulators
    (sampled every 40ms -> 90,000 samples/hour, scaled by 2^-44 mW) into dBm/PRB.
    """
    out = df_prb.copy()
    vals = out[prb_cols].apply(pd.to_numeric, errors="coerce").fillna(0.0).to_numpy(dtype=float)
    prb_med = np.median(vals, axis=1)
    prb_max = np.max(vals, axis=1)
    prb_mean = np.mean(vals, axis=1)

    if isinstance(n_hours, pd.Series):
        hrs_arr = np.maximum(n_hours.to_numpy(dtype=float), 1.0)
    else:
        hrs_arr = float(max(n_hours, 1.0))

    samples_per_period = 90000.0 * hrs_arr  # 3600s / 0.04s = 90,000 samples per hour
    scale_mw = 2.0 ** -44

    with np.errstate(divide="ignore", invalid="ignore"):
        out["PRB_Peak_to_Median_dB"] = np.round(
            np.where((prb_med > 0) & (prb_max > 0), 10.0 * np.log10(prb_max / prb_med), 0.0),
            2,
        )
        out["PRB_Max_dBm"] = np.round(
            np.where(prb_max > 0, 10.0 * np.log10((prb_max / samples_per_period) * scale_mw), -121.0),
            2,
        )
        out["PRB_Mean_dBm"] = np.round(
            np.where(prb_mean > 0, 10.0 * np.log10((prb_mean / samples_per_period) * scale_mw), -121.0),
            2,
        )
    return out


def convert_prb_row_to_dbm(prb_row: pd.Series, prb_cols: List[str], n_hours: float = 1.0) -> np.ndarray:
    """Converts a single Cell or Cell-Hour row of PRB1..PRB100 raw counters into dBm/PRB."""
    raw = pd.to_numeric(prb_row[prb_cols], errors="coerce").fillna(0.0).to_numpy(dtype=float)
    samples = 90000.0 * max(float(n_hours), 1.0)
    scale_mw = 2.0 ** -44
    with np.errstate(divide="ignore", invalid="ignore"):
        dbm = np.where(raw > 0, 10.0 * np.log10((raw / samples) * scale_mw), -121.0)
    return np.round(dbm, 2)


def aggregate_to_cell_level(
    master_hourly: pd.DataFrame,
    prb_hourly: Optional[pd.DataFrame],
) -> Tuple[pd.DataFrame, Optional[pd.DataFrame]]:
    """
    Aggregates Hourly ROP raw counters to unique Cell level by SUMMING raw counters
    (never averaging percentages!) before computing Cell-level KPIs and RCA.
    """
    cell_keys = [k for k in CELL_KEYS if k in master_hourly.columns]
    if not cell_keys:
        return master_hourly.copy(), prb_hourly

    exclude_non_sum = set(
        cell_keys
        + [
            "Date",
            "Hour",
            "ROP_Time",
            "Avg_CQI",
            "Low_CQI_Ratio_Pct",
            "CQI_OOR_Pct",
            "Hom_256QAM_Share_Pct",
            "Rank1_Ratio_Pct",
            "UL_Interf_Avg_dBm",
            "UL_Interf_Arith_dBm",
            "PRB_Peak_to_Median_dB",
            "PRB_Max_dBm",
            "PRB_Mean_dBm",
        ]
    )
    sum_cols = [
        c
        for c in master_hourly.columns
        if c not in exclude_non_sum and pd.api.types.is_numeric_dtype(master_hourly[c])
    ]

    rop_counts = master_hourly.groupby(cell_keys, as_index=False).size().rename(columns={"size": "ROP_Count"})
    cell_agg = master_hourly.groupby(cell_keys, as_index=False)[sum_cols].sum()
    cell_agg = pd.merge(cell_agg, rop_counts, on=cell_keys, how="left")
    cell_agg["ROP_Time"] = "All Hours (" + cell_agg["ROP_Count"].astype(str) + " ROPs)"
    cell_agg = _finalize_radio_metrics(cell_agg)

    # If Counters_Organized.xlsx didn't have raw UL_Weighted_mW_Sum, preserve mean UL_Interf_Avg_dBm weighted by samples
    if "UL_Interf_Avg_dBm" not in cell_agg.columns and "UL_Interf_Avg_dBm" in master_hourly.columns:
        ul_tmp = master_hourly.copy()
        ul_tmp["_mw"] = 10.0 ** (pd.to_numeric(ul_tmp["UL_Interf_Avg_dBm"], errors="coerce").fillna(-121.0) / 10.0)
        ul_cell = ul_tmp.groupby(cell_keys, as_index=False)["_mw"].mean()
        ul_cell["UL_Interf_Avg_dBm"] = np.round(10.0 * np.log10(ul_cell["_mw"]), 2)
        cell_agg = pd.merge(cell_agg, ul_cell[cell_keys + ["UL_Interf_Avg_dBm"]], on=cell_keys, how="left")

    prb_cell = None
    if prb_hourly is not None and not prb_hourly.empty:
        prb_cols = natural_prb_cols(list(prb_hourly.columns))
        p_keys = [k for k in CELL_KEYS if k in prb_hourly.columns]
        if prb_cols and p_keys:
            p_cnt = prb_hourly.groupby(p_keys, as_index=False).size().rename(columns={"size": "ROP_Count"})
            prb_sum = prb_hourly.groupby(p_keys, as_index=False)[prb_cols].sum()
            prb_sum = pd.merge(prb_sum, p_cnt, on=p_keys, how="left")
            prb_cell = _compute_prb_metrics(prb_sum, prb_cols, n_hours=prb_sum["ROP_Count"])
            cell_agg = pd.merge(
                cell_agg,
                prb_cell[p_keys + ["PRB_Peak_to_Median_dB", "PRB_Max_dBm", "PRB_Mean_dBm"]],
                on=p_keys,
                how="left",
            )
    elif "PRB_Peak_to_Median_dB" in master_hourly.columns:
        p_max = master_hourly.groupby(cell_keys, as_index=False)["PRB_Peak_to_Median_dB"].max()
        cell_agg = pd.merge(cell_agg, p_max, on=cell_keys, how="left")

    return cell_agg.sort_values(cell_keys).reset_index(drop=True), prb_cell


# ==============================================================================
# 4. KPI ENGINE & 7 MULTI-DOMAIN RCA DECISION TREES (SKILL v9.0)
# ==============================================================================

def compute_kpis_and_rca(df: pd.DataFrame, is_aggregated: bool = False) -> pd.DataFrame:
    """
    Computes all 3GPP/Ericsson MOM v9.0 KPIs, enforces the Rule of Minimum Denominator,
    evaluates all 7 RCA decision trees, and attaches Diagnostic Evidence + AMOS Playbooks.
    """
    out = df.copy()

    # Standardized string keys
    out["eNodeB"] = out["ERBS Id"].astype(str) if "ERBS Id" in out.columns else "Unknown_eNB"
    out["Cell_Name"] = out["EUtranCell Id"].astype(str) if "EUtranCell Id" in out.columns else "Unknown_Cell"
    if "ROP_Time" not in out.columns:
        out["ROP_Time"] = _make_rop_time(out)

    # ------------------------------------------------------------------
    # 1. Accessibility: CSSR (%)
    # ------------------------------------------------------------------
    rrc_att = get_col(out, "pmRrcConnEstabAtt")
    rrc_reatt = get_col(out, "pmRrcConnEstabAttReatt")
    rrc_ovl = get_col(out, "pmRrcConnEstabFailMmeOvlMos") + get_col(out, "pmRrcConnEstabFailMmeOvlMod")
    out["CSSR_Denom"] = (rrc_att - rrc_reatt).clip(lower=0)
    rrc_sr = safe_div(get_col(out, "pmRrcConnEstabSucc"), (out["CSSR_Denom"] - rrc_ovl).clip(lower=0))

    s1_ovl = (
        get_col(out, "pmS1SigConnEstabFailMmeOvlMosCe")
        if "pmS1SigConnEstabFailMmeOvlMosCe" in out.columns
        else get_col(out, "pmS1SigConnEstabFailMmeOvlMos")
    )
    s1_sr = safe_div(get_col(out, "pmS1SigConnEstabSucc"), (get_col(out, "pmS1SigConnEstabAtt") - s1_ovl).clip(lower=0))
    erab_sr = safe_div(get_col(out, "pmErabEstabSuccInit"), get_col(out, "pmErabEstabAttInit"))

    out["RRC_SR_Pct"] = (rrc_sr * 100.0).round(2)
    out["S1_SR_Pct"] = (s1_sr * 100.0).round(2)
    out["ERAB_Init_SR_Pct"] = (erab_sr * 100.0).round(2)
    out["CSSR_Pct"] = (rrc_sr * s1_sr * erab_sr * 100.0).round(2)

    # ------------------------------------------------------------------
    # 2. Retainability: DCR_Overall (%), DCR_eNB (%), DCR_Normalized
    # ------------------------------------------------------------------
    abn_enb_act = get_col(out, "pmErabRelAbnormalEnbAct")
    abn_mme_act = get_col(out, "pmErabRelAbnormalMmeAct")
    dcr_denom = get_col(out, "pmErabRelAbnormalEnb") + get_col(out, "pmErabRelNormalEnb") + get_col(out, "pmErabRelMme")
    out["Active_Drops"] = abn_enb_act + abn_mme_act
    out["DCR_Denom"] = dcr_denom
    out["DCR_Overall_Pct"] = safe_div(out["Active_Drops"], dcr_denom, 100.0).round(2)
    out["DCR_eNB_Pct"] = safe_div(abn_enb_act, dcr_denom, 100.0).round(2)
    out["DCR_Normalized"] = safe_div(out["Active_Drops"], get_col(out, "pmSessionTimeUe"), 36_000_000.0).round(2)

    # ------------------------------------------------------------------
    # 3. Mobility: Intra-Freq, Inter-Freq & Combined HOSR (%)
    # ------------------------------------------------------------------
    intra_p_att = get_col(out, "pmCellHoPrepAttLteIntraF")
    intra_p_suc = get_col(out, "pmCellHoPrepSuccLteIntraF")
    intra_e_att = get_col(out, "pmCellHoExeAttLteIntraF")
    intra_e_suc = get_col(out, "pmCellHoExeSuccLteIntraF")
    out["Intra_HO_Prep_SR_Pct"] = safe_div(intra_p_suc, intra_p_att, 100.0).round(2)
    out["Intra_HO_Exe_SR_Pct"] = safe_div(intra_e_suc, intra_e_att, 100.0).round(2)
    out["Intra_Freq_HOSR_Pct"] = ((out["Intra_HO_Prep_SR_Pct"] * out["Intra_HO_Exe_SR_Pct"]) / 100.0).round(2)

    inter_p_att = get_col(out, "pmCellHoPrepAttLteInterF")
    inter_p_suc = get_col(out, "pmCellHoPrepSuccLteInterF")
    inter_e_att = get_col(out, "pmCellHoExeAttLteInterF")
    inter_e_suc = get_col(out, "pmCellHoExeSuccLteInterF")
    rej_ul = get_col(out, "pmHoPrepRejOutUlThres")
    out["Inter_HO_Prep_SR_Pct"] = safe_div(inter_p_suc, inter_p_att + rej_ul, 100.0).round(2)
    out["Inter_HO_Exe_SR_Pct"] = safe_div(inter_e_suc, inter_e_att, 100.0).round(2)
    out["Inter_Freq_HOSR_Pct"] = ((out["Inter_HO_Prep_SR_Pct"] * out["Inter_HO_Exe_SR_Pct"]) / 100.0).round(2)

    comb_prep_sr = safe_div(intra_p_suc + inter_p_suc, intra_p_att + inter_p_att + rej_ul, 100.0)
    comb_exe_sr = safe_div(intra_e_suc + inter_e_suc, intra_e_att + inter_e_att, 100.0)
    out["Combined_HO_Att"] = intra_p_att + inter_p_att
    out["Combined_HOSR_Pct"] = ((comb_prep_sr * comb_exe_sr) / 100.0).round(2)

    # ------------------------------------------------------------------
    # 4. Radio Quality & PIM Signature Defaults
    # ------------------------------------------------------------------
    for col_name, default_val in [
        ("Avg_CQI", np.nan),
        ("Low_CQI_Ratio_Pct", 0.0),
        ("CQI_OOR_Pct", 0.0),
        ("Hom_256QAM_Share_Pct", 0.0),
        ("Rank1_Ratio_Pct", 0.0),
        ("UL_Interf_Avg_dBm", -121.0),
        ("PRB_Peak_to_Median_dB", 0.0),
        ("PRB_Max_dBm", -121.0),
    ]:
        if col_name not in out.columns:
            out[col_name] = default_val

    # ------------------------------------------------------------------
    # 5. 3G/4G IRAT Offload & Instability
    # ------------------------------------------------------------------
    irat_rel_wcdma = get_col(out, "pmUeCtxtRelSCWcdma")
    irat_csfb_wcdma = get_col(out, "pmUeCtxtRelCsfbWcdma")
    irat_ho_succ = get_col(out, "pmHoExeSucc")
    rrc_succ = get_col(out, "pmRrcConnEstabSucc")
    out["IRAT_Events"] = irat_rel_wcdma + irat_csfb_wcdma + irat_ho_succ
    out["IRAT_Offload_Pct"] = safe_div(out["IRAT_Events"], rrc_succ, 100.0).round(2)
    out["B2_Trigger_Ratio_Pct"] = safe_div(get_col(out, "pmB2BestCellEvalReport"), get_col(out, "pmBadCovSearchEvalReport"), 100.0).round(2)

    # ------------------------------------------------------------------
    # 6. VoLTE QCI=1 Accessibility, Retainability & SRVCC
    # ------------------------------------------------------------------
    v_att_add = get_col(out, "pmErabEstabAttAddedQci")
    v_suc_add = get_col(out, "pmErabEstabSuccAddedQci")
    v_att_ini = get_col(out, "pmErabEstabAttInitQci")
    v_suc_ini = get_col(out, "pmErabEstabSuccInitQci")
    v_excl = (
        get_col(out, "pmErabEstabAttAddedHoOngoingQci")
        + get_col(out, "pmErabEstabAttAddedCsfbQci")
        + get_col(out, "pmErabEstabFailAddedErabIdInUseQci")
    )
    out["VoLTE_CSSR_Denom"] = v_att_add + v_att_ini
    v_ran_sr = safe_div(v_suc_add + v_suc_ini, (out["VoLTE_CSSR_Denom"] - v_excl).clip(lower=0))
    out["VoLTE_CSSR_RAN_Pct"] = (v_ran_sr * 100.0).round(2)
    out["VoLTE_CSSR_E2E_Pct"] = (rrc_sr * s1_sr * v_ran_sr * 100.0).round(2)

    v_abn_enb = get_col(out, "pmErabRelAbnormalEnbActQci")
    v_abn_mme = get_col(out, "pmErabRelAbnormalMmeActQci")
    v_rel_den = get_col(out, "pmErabRelAbnormalEnbQci") + get_col(out, "pmErabRelNormalEnbQci") + get_col(out, "pmErabRelMmeQci")
    out["VoLTE_Active_Drops"] = v_abn_enb + v_abn_mme
    out["VoLTE_DCR_Denom"] = v_rel_den
    out["VoLTE_DCR_Rel_Pct"] = safe_div(out["VoLTE_Active_Drops"], v_rel_den, 100.0).round(2)

    rop_sec = 3600.0 * (get_col(out, "ROP_Count", 1.0) if is_aggregated and "ROP_Count" in out.columns else 1.0)
    t_volte_sec = get_col(out, "pmSessionTimeDrbQci")
    if "pmErabQciLevSum" in out.columns:
        lev_samp = get_col(out, "pmErabQciLevSamp", 720.0).replace(0, 720.0)
        fallback_sec = (get_col(out, "pmErabQciLevSum") / lev_samp) * rop_sec
        t_volte_sec = np.where(t_volte_sec > 0, t_volte_sec, fallback_sec)
    out["VoLTE_Drops_10kHrs"] = safe_div(out["VoLTE_Active_Drops"], pd.Series(t_volte_sec, index=out.index), 36_000_000.0).round(2)

    out["SRVCC_Prep_SR_Pct"] = safe_div(get_col(out, "pmHoPrepSuccSrvcc"), get_col(out, "pmHoPrepAttSrvcc"), 100.0).round(2)
    out["SRVCC_Exe_SR_Pct"] = safe_div(get_col(out, "pmHoExeSuccSrvcc"), get_col(out, "pmHoExeAttSrvcc"), 100.0).round(2)

    # ------------------------------------------------------------------
    # 7. Deterministic Multi-Domain RCA Decision Trees (Skill v9.0)
    # ------------------------------------------------------------------
    severities: List[str] = []
    categories: List[str] = []
    primary_metrics: List[str] = []
    isolated_phases: List[str] = []
    root_causes: List[str] = []
    all_issues_list: List[str] = []
    scenarios_list: List[str] = []
    confidences: List[int] = []
    evidences: List[str] = []

    records = out.to_dict("records")
    for r in records:
        findings: List[Dict[str, Any]] = []
        valid_any = False

        ul_dbm = float(r.get("UL_Interf_Avg_dBm") if pd.notna(r.get("UL_Interf_Avg_dBm")) else -121.0)
        pim_db = float(r.get("PRB_Peak_to_Median_dB") if pd.notna(r.get("PRB_Peak_to_Median_dB")) else 0.0)
        prb_max_dbm = float(r.get("PRB_Max_dBm") if pd.notna(r.get("PRB_Max_dBm")) else -121.0)
        avg_cqi = float(r.get("Avg_CQI") if pd.notna(r.get("Avg_CQI")) else 10.0)
        cqi_oor = float(r.get("CQI_OOR_Pct") if pd.notna(r.get("CQI_OOR_Pct")) else 0.0)
        low_cqi = float(r.get("Low_CQI_Ratio_Pct") if pd.notna(r.get("Low_CQI_Ratio_Pct")) else 0.0)
        rank1_pct = float(r.get("Rank1_Ratio_Pct") if pd.notna(r.get("Rank1_Ratio_Pct")) else 0.0)

        # ---- TREE 1: Accessibility (CSSR %) ----
        cssr_den = float(r.get("CSSR_Denom", 0) or 0)
        cssr_val = r.get("CSSR_Pct")
        if cssr_den >= MIN_SAMPLES["CSSR"] and pd.notna(cssr_val):
            valid_any = True
            if cssr_val < 97.0:
                sev = "CRITICAL" if cssr_val < 95.0 else "DEGRADED"
                rrc_sr_v = float(r.get("RRC_SR_Pct", 100.0) or 100.0)
                s1_sr_v = float(r.get("S1_SR_Pct", 100.0) or 100.0)
                erab_sr_v = float(r.get("ERAB_Init_SR_Pct", 100.0) or 100.0)
                ra_att = float(r.get("pmRaAttCbra", 0) or 0)
                ra_ooc = float(r.get("pmRaFailCbraMsg1DiscOoc", 0) or 0)

                if rrc_sr_v <= min(s1_sr_v, erab_sr_v) or rrc_sr_v < 97.5:
                    phase = "RRC / RACH Setup"
                    if (r.get("pmRrcConnEstabFailLic", 0) or 0) > 0:
                        rc = "Connected Users License Limit Exceeded (CapacityState)"
                        conf = 97
                    elif (r.get("pmRrcConnEstabFailMpOverload", 0) or 0) > 0 or (r.get("pmRrcConnEstabFailDuIntens", 0) or 0) > 0:
                        rc = "Baseband MP CPU Overload / DU Intensity Reject"
                        conf = 96
                    elif ra_att > 0 and (ra_ooc / ra_att) > 0.05:
                        rc = f"Overshooting Cell — RACH Msg1 Out-of-Coverage ({ra_ooc/ra_att*100:.1f}% > cellRange)"
                        conf = 93
                    elif ul_dbm > -110.0 or prb_max_dbm > -104.0:
                        rc = f"UL Interference ({ul_dbm:.1f} dBm) Blocking RACH Msg3 Detection"
                        conf = 91
                    else:
                        rc = "PDCCH / Control Channel Element (CCE) Congestion"
                        conf = 87
                elif s1_sr_v < 98.0:
                    phase = "S1-MME Signaling"
                    rc = "S1-MME Overload or SCTP Transport Latency (TermPointToMme)"
                    conf = 90
                else:
                    phase = "Initial E-RAB Setup"
                    if (r.get("pmErabEstabFailInitLic", 0) or 0) > 0:
                        rc = "Missing E-RAB / RLC UM License Capacity"
                        conf = 95
                    elif (r.get("pmErabEstabFailGbrDlEnb", 0) or 0) > 0 or (r.get("pmErabEstabFailGbrUlEnb", 0) or 0) > 0:
                        rc = "GBR Admission Control Bandwidth Rejection (AdmissionControl=1)"
                        conf = 94
                    else:
                        rc = "Radio Bearer Reconfiguration Timeout / UE Unresponsive"
                        conf = 86

                findings.append({
                    "prio": 1 if sev == "CRITICAL" else 3,
                    "sev": sev,
                    "cat": "Accessibility (CSSR)",
                    "metric": f"CSSR ({cssr_val:.2f}%)",
                    "phase": phase,
                    "rc": rc,
                    "scen": ["SCENARIO_A", "SCENARIO_B"] if ul_dbm > -110.0 else ["SCENARIO_B"],
                    "conf": conf,
                })

        # ---- TREE 2: Retainability (DCR %) ----
        dcr_den = float(r.get("DCR_Denom", 0) or 0)
        dcr_val = r.get("DCR_Overall_Pct")
        if dcr_den >= MIN_SAMPLES["DCR"] and pd.notna(dcr_val):
            valid_any = True
            if dcr_val > 1.0:
                sev = "CRITICAL" if dcr_val > 2.5 else "DEGRADED"
                tot_drops = max(float(r.get("Active_Drops", 0) or 0), 1.0)
                ue_lost = float(r.get("pmErabRelAbnormalEnbActUeLost", 0) or 0)
                ho_drop = float(r.get("pmErabRelAbnormalEnbActHo", 0) or 0)
                tn_drop = float(r.get("pmErabRelAbnormalEnbActTnFail", 0) or 0)
                pe_drop = float(r.get("pmErabRelAbnormalEnbActPe", 0) or 0)
                mme_drop = float(r.get("pmErabRelAbnormalMmeAct", 0) or 0)

                scens = ["SCENARIO_C"]
                if (ue_lost / tot_drops) > 0.50:
                    phase = "Radio Link Failure (UE Lost)"
                    if ul_dbm > -108.0 or (prb_max_dbm > -102.0 and pim_db > 6.0):
                        scens = ["SCENARIO_A", "SCENARIO_C"]
                        if pim_db > 6.0:
                            rc = f"Narrowband / PIM UL Interference (UL={ul_dbm:.1f} dBm, PIM Δ={pim_db:.1f} dB)"
                            conf = 96
                        else:
                            rc = f"Broadband External UL Interference / Jammer ({ul_dbm:.1f} dBm)"
                            conf = 94
                    elif avg_cqi < 7.0 or cqi_oor > 5.0:
                        rc = f"Severe DL Coverage Hole / Out-Of-Sync (CQI={avg_cqi:.1f}, OOR={cqi_oor:.1f}%)"
                        conf = 93
                    else:
                        rc = "UL Power Control Limitation (pZeroNominalPusch / alpha Conservative)"
                        conf = 89
                elif (ho_drop / tot_drops) > 0.30:
                    phase = "Handover Execution Drop"
                    rc = "Handover Execution Drop / Late Event A3 or Missing Neighbor"
                    scens = ["SCENARIO_C", "SCENARIO_D"]
                    conf = 92
                elif (tn_drop / tot_drops) > 0.20:
                    phase = "S1/X2 Transport Layer"
                    rc = "S1/X2 Transport Link Flapping (SctpEndpoint / EthernetPort)"
                    conf = 95
                elif (pe_drop / tot_drops) > 0.15:
                    phase = "ARP Pre-emption"
                    rc = "ARP Bearer Pre-emption Drops under High Cell Congestion"
                    conf = 94
                elif (mme_drop / tot_drops) > 0.30:
                    phase = "Core EPC / MME Release"
                    rc = f"EPC / MME Initiated Abnormal Bearer Release ({mme_drop/tot_drops*100:.0f}% Share)"
                    conf = 91
                else:
                    phase = "Multi-Factor E-RAB Drop"
                    rc = "Mixed Radio & Signaling Abnormal E-RAB Releases"
                    conf = 85

                findings.append({
                    "prio": 0 if sev == "CRITICAL" else 2,
                    "sev": sev,
                    "cat": "Retainability (DCR)",
                    "metric": f"Overall DCR ({dcr_val:.2f}%)",
                    "phase": phase,
                    "rc": rc,
                    "scen": scens,
                    "conf": conf,
                })

        # ---- TREE 3: VoLTE QCI=1 & SRVCC ----
        v_cssr_den = float(r.get("VoLTE_CSSR_Denom", 0) or 0)
        v_cssr_val = r.get("VoLTE_CSSR_RAN_Pct")
        if v_cssr_den >= MIN_SAMPLES["VoLTE_CSSR"] and pd.notna(v_cssr_val):
            valid_any = True
            if v_cssr_val < 97.5:
                sev = "CRITICAL" if v_cssr_val < 95.0 else "DEGRADED"
                ho_ong = float(r.get("pmErabEstabAttAddedHoOngoingQci", 0) or 0)
                att_add = max(float(r.get("pmErabEstabAttAddedQci", 0) or 0), 1.0)
                if (r.get("pmErabEstabFailGbrDlEnb", 0) or 0) > 0 or (r.get("pmErabEstabFailGbrUlEnb", 0) or 0) > 0:
                    rc = "VoLTE CAC GBR Bandwidth Exhaustion (AdmissionControl=1)"
                    conf = 95
                elif (r.get("pmErabEstabFailAddedErabIdInUseQci", 0) or 0) > 0:
                    rc = "QCI=1 E-RAB ID Collision / Hung Bearer Context (eNB-MME)"
                    conf = 94
                elif (ho_ong / att_add) > 0.05:
                    rc = "QCI=1 Dedicated Voice Setup Interrupted by Cell-Edge Handover"
                    conf = 91
                elif ul_dbm > -110.0:
                    rc = f"UL Interference ({ul_dbm:.1f} dBm) Blocking QCI=1 Activation"
                    conf = 90
                else:
                    rc = "QCI=1 Dedicated Voice Bearer Setup Timeout / License Limit"
                    conf = 87
                findings.append({
                    "prio": 1 if sev == "CRITICAL" else 3,
                    "sev": sev,
                    "cat": "VoLTE QCI=1 & SRVCC",
                    "metric": f"VoLTE CSSR ({v_cssr_val:.2f}%)",
                    "phase": "QCI=1 Bearer Setup",
                    "rc": rc,
                    "scen": ["SCENARIO_G"],
                    "conf": conf,
                })

        v_dcr_den = float(r.get("VoLTE_DCR_Denom", 0) or 0)
        v_dcr_val = r.get("VoLTE_DCR_Rel_Pct")
        if v_dcr_den >= MIN_SAMPLES["VoLTE_DCR"] and pd.notna(v_dcr_val):
            valid_any = True
            if v_dcr_val > 0.8:
                sev = "CRITICAL" if v_dcr_val > 2.0 else "DEGRADED"
                v_tot = max(float(r.get("VoLTE_Active_Drops", 0) or 0), 1.0)
                v_pe = float(r.get("pmErabRelAbnormalEnbActPeQci", 0) or 0)
                v_uelost = float(r.get("pmErabRelAbnormalEnbActUeLostQci", 0) or 0)
                v_ho = float(r.get("pmErabRelAbnormalEnbActHoQci", 0) or 0)
                v_mme = float(r.get("pmErabRelAbnormalMmeActQci", 0) or 0)
                srvcc_p_att = float(r.get("pmHoPrepAttSrvcc", 0) or 0)
                srvcc_e_att = float(r.get("pmHoExeAttSrvcc", 0) or 0)
                srvcc_p_sr = float(r.get("SRVCC_Prep_SR_Pct", 100.0) if pd.notna(r.get("SRVCC_Prep_SR_Pct")) else 100.0)
                srvcc_e_sr = float(r.get("SRVCC_Exe_SR_Pct", 100.0) if pd.notna(r.get("SRVCC_Exe_SR_Pct")) else 100.0)

                if v_pe > 0:
                    phase = "ARP Pre-emption (QCI=1)"
                    rc = "CRITICAL MISCONFIG: VoLTE QCI=1 Bearer Pre-empted (Set NOT_PREEMPTABLE)"
                    conf = 98
                elif (v_uelost / v_tot) > 0.50:
                    phase = "QCI=1 RTP Radio Link"
                    if ul_dbm > -108.0 or (prb_max_dbm > -102.0 and pim_db > 6.0):
                        rc = f"UL Interference / PIM (Δ={pim_db:.1f} dB) Destroying RTP Voice Packets (Enable ROHC + TTI Bundling)"
                        conf = 95
                    else:
                        rc = "Cell-Edge Voice Coverage Hole / Delayed SRVCC Trigger (Raise b2Threshold1Rsrp)"
                        conf = 92
                elif (v_ho / v_tot) > 0.30:
                    if srvcc_p_att > 0 and srvcc_p_sr < 95.0:
                        phase = "SRVCC Preparation Phase"
                        rc = f"SRVCC Prep Failure ({srvcc_p_sr:.1f}%) — Missing srvccHoAllowed or MME Sv Reject"
                        conf = 94
                    elif srvcc_e_att > 0 and srvcc_e_sr < 92.0:
                        phase = "SRVCC Execution Phase"
                        rc = f"SRVCC Execution Failure ({srvcc_e_sr:.1f}%) — Poor 3G Target RSCP/EcNo Threshold"
                        conf = 93
                    else:
                        phase = "Intra-LTE Voice HO"
                        rc = "VoLTE Active Call Drop During Intra-LTE Handover Execution"
                        conf = 90
                elif (v_mme / v_tot) > 0.40:
                    phase = "IMS / S1-U Media Path"
                    rc = "Core / IMS RTP Inactivity Timeout or DSCP!=46 EF Packet Loss"
                    conf = 90
                else:
                    phase = "QCI=1 Voice Retainability"
                    rc = "VoLTE QCI=1 Abnormal Voice Release (Check RLC Poll Retx & ROHC)"
                    conf = 86

                findings.append({
                    "prio": 0 if sev == "CRITICAL" else 2,
                    "sev": sev,
                    "cat": "VoLTE QCI=1 & SRVCC",
                    "metric": f"VoLTE DCR ({v_dcr_val:.2f}%)",
                    "phase": phase,
                    "rc": rc,
                    "scen": ["SCENARIO_G"],
                    "conf": conf,
                })

        # ---- TREE 4: Intra-Frequency Mobility ----
        intra_p = float(r.get("pmCellHoPrepAttLteIntraF", 0) or 0)
        intra_e = float(r.get("pmCellHoExeAttLteIntraF", 0) or 0)
        intra_hosr = r.get("Intra_Freq_HOSR_Pct")
        if intra_p >= MIN_SAMPLES["Intra_HO"] and intra_e >= MIN_SAMPLES["Intra_HO"] and pd.notna(intra_hosr):
            valid_any = True
            if intra_hosr < 96.0:
                sev = "CRITICAL" if intra_hosr < 92.0 else "DEGRADED"
                prep_sr = float(r.get("Intra_HO_Prep_SR_Pct", 100.0) or 100.0)
                ho_drops = float(r.get("pmErabRelAbnormalEnbActHo", 0) or 0)
                if prep_sr < 98.5:
                    phase = "Preparation Phase"
                    if (r.get("pmHoPrepRejInLicConnUsers", 0) or 0) > 0:
                        rc = "Target Cell Connected Users License Exhausted"
                        conf = 96
                    elif (r.get("pmHoPrepRejOutUlThres", 0) or 0) > 0:
                        rc = "Target Cell Severe UL Interference Rejection"
                        conf = 94
                    else:
                        rc = "X2 Link Down (TermPointToENB) or Neighbor isHoAllowed=false"
                        conf = 91
                else:
                    phase = "Execution Phase"
                    if (ho_drops / max(intra_e, 1.0)) > 0.05 and (avg_cqi < 7.5 or cqi_oor > 5.0):
                        rc = "Late Handover — Event A3 Offset / TimeToTrigger Too High"
                        conf = 93
                    else:
                        rc = "Early / Ping-Pong Handover or PCI Mod-3/Mod-30 Collision"
                        conf = 89
                findings.append({
                    "prio": 1 if sev == "CRITICAL" else 3,
                    "sev": sev,
                    "cat": "Mobility (Intra/Inter HOSR)",
                    "metric": f"Intra-Freq HOSR ({intra_hosr:.2f}%)",
                    "phase": phase,
                    "rc": rc,
                    "scen": ["SCENARIO_D"],
                    "conf": conf,
                })

        # ---- TREE 5: Inter-Frequency Mobility ----
        inter_p = float(r.get("pmCellHoPrepAttLteInterF", 0) or 0)
        inter_e = float(r.get("pmCellHoExeAttLteInterF", 0) or 0)
        inter_hosr = r.get("Inter_Freq_HOSR_Pct")
        if inter_p >= MIN_SAMPLES["Inter_HO"] and inter_e >= MIN_SAMPLES["Inter_HO"] and pd.notna(inter_hosr):
            valid_any = True
            if inter_hosr < 94.0:
                sev = "CRITICAL" if inter_hosr < 88.0 else "DEGRADED"
                prep_sr = float(r.get("Inter_HO_Prep_SR_Pct", 100.0) or 100.0)
                if prep_sr < 98.0:
                    phase = "Preparation Phase"
                    if (r.get("pmHoPrepRejInLicConnUsers", 0) or 0) > 0:
                        rc = "Target Carrier Connected Users License Exhausted"
                        conf = 95
                    else:
                        rc = "Measurement Gap Allocation or EUtranFreqRelation HoAllowed Issue"
                        conf = 90
                else:
                    phase = "Execution Phase"
                    rc = "Late Event A2 Inter-Freq Search or Late Event A5 Execution Threshold"
                    conf = 91
                findings.append({
                    "prio": 1 if sev == "CRITICAL" else 3,
                    "sev": sev,
                    "cat": "Mobility (Intra/Inter HOSR)",
                    "metric": f"Inter-Freq HOSR ({inter_hosr:.2f}%)",
                    "phase": phase,
                    "rc": rc,
                    "scen": ["SCENARIO_E"],
                    "conf": conf,
                })

        # ---- TREE 6: Radio Quality & PIM Signature ----
        cqi_samp = float(r.get("CQI_Valid_Samples", 0) or 0)
        if cqi_samp >= MIN_SAMPLES["CQI"] or ul_dbm > -110.0:
            if cqi_samp >= MIN_SAMPLES["CQI"]:
                valid_any = True
            if avg_cqi < 7.5 or cqi_oor > 5.0 or ul_dbm > -108.0:
                sev = "CRITICAL" if (avg_cqi < 6.0 or cqi_oor > 10.0 or ul_dbm > -102.0) else "DEGRADED"
                if ul_dbm > -108.0:
                    phase = "UL PRB Spectrum"
                    if pim_db > 6.0:
                        rc = f"PIM / Narrowband UL Interference (UL={ul_dbm:.1f} dBm, Peak-to-Med Δ={pim_db:.1f} dB > 6 dB)"
                        conf = 96
                    else:
                        rc = f"Broadband External UL Interference (UL={ul_dbm:.1f} dBm, Flat Spectrum Δ={pim_db:.1f} dB <= 6 dB)"
                        conf = 94
                elif rank1_pct > 75.0:
                    phase = "MIMO Spatial Channel"
                    rc = f"Pilot Pollution / Cross-Pol Imbalance (Rank-1={rank1_pct:.1f}%, Avg CQI={avg_cqi:.2f})"
                    conf = 90
                elif cqi_oor > 5.0:
                    phase = "Cell-Edge DL SINR"
                    rc = f"Excessive CQI Out-of-Range Reports (OOR={cqi_oor:.1f}% > 5% Baseline)"
                    conf = 91
                else:
                    phase = "DL CRS / SINR"
                    rc = f"Weak Downlink SINR / Inter-Cell CRS Interference (Avg CQI={avg_cqi:.2f})"
                    conf = 88
                findings.append({
                    "prio": 2 if sev == "CRITICAL" else 4,
                    "sev": sev,
                    "cat": "Radio Quality / UL PIM",
                    "metric": f"UL {ul_dbm:.1f} dBm" if ul_dbm > -108.0 else (f"CQI OOR ({cqi_oor:.1f}%)" if cqi_oor > 5.0 else f"Avg CQI ({avg_cqi:.2f})"),
                    "phase": phase,
                    "rc": rc,
                    "scen": ["SCENARIO_A"],
                    "conf": conf,
                })

        # ---- TREE 7: 3G/4G IRAT Offload & Instability ----
        rrc_succ_v = float(r.get("pmRrcConnEstabSucc", 0) or 0)
        irat_pct = r.get("IRAT_Offload_Pct")
        if rrc_succ_v >= MIN_SAMPLES["IRAT"] and pd.notna(irat_pct):
            valid_any = True
            if irat_pct > 5.0:
                sev = "CRITICAL" if irat_pct > 10.0 else "DEGRADED"
                bad_cov_s = float(r.get("pmBadCovSearchEvalReport", 0) or 0)
                csfb_rel = float(r.get("pmUeCtxtRelCsfbWcdma", 0) or 0)
                sc_rel = float(r.get("pmUeCtxtRelSCWcdma", 0) or 0)
                b2_rat = float(r.get("B2_Trigger_Ratio_Pct", 0) if pd.notna(r.get("B2_Trigger_Ratio_Pct")) else 0.0)

                if (bad_cov_s / rrc_succ_v) > 0.40:
                    phase = "Event A2 Secondary Search"
                    rc = f"Premature 3G IRAT Search (a2Threshold2RsrpSec Too High, Search Ratio={bad_cov_s/rrc_succ_v*100:.0f}%)"
                    conf = 93
                elif csfb_rel > 0.50 * max(sc_rel + csfb_rel, 1.0):
                    phase = "CSFB Voice Redirection"
                    rc = f"Heavy CSFB 3G Voice Fallback ({int(csfb_rel)} CSFB Rels — Non-VoLTE UEs / IMS Attach Fail)"
                    conf = 92
                elif b2_rat > 50.0:
                    phase = "Event B2 UTRAN Threshold"
                    rc = f"Loose Event B2 3G Target Threshold (B2/Search={b2_rat:.0f}% — Tighten b2Threshold2RscpUtra)"
                    conf = 90
                else:
                    phase = "Session Continuity / Priority"
                    rc = "Excessive 3G Session Continuity Offload / Reselection Priority Mismatch"
                    conf = 88

                findings.append({
                    "prio": 2 if sev == "CRITICAL" else 4,
                    "sev": sev,
                    "cat": "IRAT 3G Offload",
                    "metric": f"IRAT Offload ({irat_pct:.2f}%)",
                    "phase": phase,
                    "rc": rc,
                    "scen": ["SCENARIO_F"],
                    "conf": conf,
                })

        # Format Diagnostic Counter Evidence string
        ev_parts = [
            f"CSSR={r.get('CSSR_Pct', np.nan):.2f}% (Att={int(cssr_den)}, RACH_OOC={int(r.get('pmRaFailCbraMsg1DiscOoc', 0) or 0)})" if pd.notna(r.get("CSSR_Pct")) else "CSSR=N/A",
            f"DCR={r.get('DCR_Overall_Pct', np.nan):.2f}% (UeLost={int(r.get('pmErabRelAbnormalEnbActUeLost', 0) or 0)}, HO={int(r.get('pmErabRelAbnormalEnbActHo', 0) or 0)}, MME={int(r.get('pmErabRelAbnormalMmeAct', 0) or 0)})" if pd.notna(r.get("DCR_Overall_Pct")) else "DCR=N/A",
            f"UL={ul_dbm:.1f}dBm (PIM_Δ={pim_db:.1f}dB, MaxPRB={prb_max_dbm:.1f}dBm)",
            f"CQI={avg_cqi:.2f} (OOR={cqi_oor:.1f}%, Low={low_cqi:.1f}%, Rank1={rank1_pct:.1f}%)",
            f"RLC_PollRetxUl={int(r.get('pmRlcPollRetxUl', 0) or 0)}",
        ]
        evidences.append(" | ".join(ev_parts))

        if not valid_any:
            severities.append("STAT_INSIGNIFICANT")
            categories.append("Low Sample (Excluded)")
            primary_metrics.append("Insufficient Traffic")
            isolated_phases.append("Rule of Min Denominator")
            root_causes.append("STAT_INSIGNIFICANT — Sample Volume Below 3GPP/Ericsson Minimum Threshold")
            all_issues_list.append("Excluded by Rule of Minimum Denominator")
            scenarios_list.append("NONE")
            confidences.append(0)
        elif not findings:
            severities.append("NORMAL")
            categories.append("Healthy / Normal")
            primary_metrics.append("All KPIs Nominal")
            isolated_phases.append("Nominal Operation")
            root_causes.append("All MBB, Radio, Mobility & VoLTE QCI=1 KPIs Within Target Baselines")
            all_issues_list.append("Nominal")
            scenarios_list.append("NONE")
            confidences.append(99)
        else:
            findings.sort(key=lambda x: x["prio"])
            top = findings[0]
            overall_sev = "CRITICAL" if any(f["sev"] == "CRITICAL" for f in findings) else "DEGRADED"
            uniq_scens: List[str] = []
            for f in findings:
                for s in f["scen"]:
                    if s not in uniq_scens:
                        uniq_scens.append(s)

            severities.append(overall_sev)
            categories.append(top["cat"])
            primary_metrics.append(top["metric"])
            isolated_phases.append(top["phase"])
            root_causes.append(top["rc"])
            all_issues_list.append(" ; ".join(f"[{f['sev']}] {f['metric']}: {f['rc']}" for f in findings))
            scenarios_list.append(" + ".join(uniq_scens))
            confidences.append(int(top["conf"]))

    out["Severity"] = severities
    out["RCA_Category"] = categories
    out["Primary_Metric"] = primary_metrics
    out["Isolated_Phase"] = isolated_phases
    out["Root_Cause"] = root_causes
    out["All_Issues"] = all_issues_list
    out["Scenario"] = scenarios_list
    out["Confidence"] = confidences
    out["Diagnostic_Evidence"] = evidences

    return out


# ==============================================================================
# 5. AUTOMATED ERICSSON AMOS / MOSHELL REMEDIATION SCRIPT GENERATOR
# ==============================================================================

def generate_amos_script(row: Union[pd.Series, Dict[str, Any]]) -> str:
    """
    Generates a production-safe Ericsson Baseband (ERS) / DU (CPP) Moshell script
    with (1) Mandatory Safety Gate Backup, (2) Online-Safe Dynamic Commands,
    (3) Commented Maintenance-Window Cell-Lock Commands, and (4) Post-Check Commit.
    """
    r = row if isinstance(row, dict) else row.to_dict()
    cell_name = str(r.get("EUtranCell Id") or r.get("Cell_Name") or "<Cell_Name>")
    erbs_id = str(r.get("ERBS Id") or r.get("eNodeB") or "<ERBS_Id>")
    earfcn = str(int(r["earfcndl"])) if "earfcndl" in r and pd.notna(r["earfcndl"]) else "<EARFCN>"
    scen_str = str(r.get("Scenario", "NONE"))
    rc_str = str(r.get("Root_Cause", ""))
    sev_str = str(r.get("Severity", "NORMAL"))

    lines: List[str] = [
        "# ==============================================================================",
        f"# ANTIGRAVITY ERICSSON AMOS / MOSHELL REMEDIATION PLAYBOOK (v9.0)",
        f"# Node (ERBS): {erbs_id}  |  Cell: {cell_name}  |  EARFCN: {earfcn}",
        f"# Severity   : {sev_str}  |  Scenario: {scen_str}",
        f"# Root Cause : {rc_str}",
        "# ==============================================================================",
        "",
        "# ------------------------------------------------------------------------------",
        "# [1] MANDATORY SAFETY GATE & PRE-EXECUTION BACKUP (BrM / CPP)",
        "# ------------------------------------------------------------------------------",
        "lt all",
        "confb+",
        "gs+",
        "cvls",
        f'cvms Pre_Opt_{cell_name}_v9 "Baseline backup prior to Antigravity RAN optimization"',
        f"st EUtranCellFDD={cell_name}",
        "st SectorCarrier",
        "alt",
        "",
    ]

    if scen_str == "NONE" or sev_str in ("NORMAL", "STAT_INSIGNIFICANT"):
        lines.extend([
            "# [INFO] Cell is operating within 3GPP/Ericsson KPI baselines (or excluded due to low traffic).",
            "# Health-check read-only audit commands:",
            f"pget EUtranCellFDD={cell_name} pmRrcConnEstab|pmErabRelAbnormal|pmRadioRecInterferencePwr$",
            "gs-",
            "confb-",
        ])
        return "\n".join(lines)

    online_cmds: List[str] = [
        "# ------------------------------------------------------------------------------",
        "# [2] ONLINE-SAFE REMEDIATION COMMANDS (ZERO CALL DROP / NO CELL LOCK)",
        "# ------------------------------------------------------------------------------",
    ]
    maint_cmds: List[str] = [
        "# ------------------------------------------------------------------------------",
        "# [3] MAINTENANCE-WINDOW COMMANDS (COMMENTED BY DEFAULT — REQUIRES bl/deb LOCK)",
        "# ------------------------------------------------------------------------------",
    ]

    if "SCENARIO_A" in scen_str:
        online_cmds.extend([
            f"# --- [SCENARIO_A: Radio Quality / UL Interference / PIM Audit & Tuning] ---",
            f"pget EUtranCellFDD={cell_name} pmRadioRecInterferencePwr$",
            f"pget EUtranCellFDD={cell_name} pmRadioRecInterferencePwrPrb",
            f"pget EUtranCellFDD={cell_name} pmRadioUeRepCqiDistr|pmRadioHomUeRepCqiDistr|pmRadioUeRepRankDistr",
            f"get EUtranCellFDD={cell_name} dl256QamEnabled|crsGain|noOfPucchCqiUsers",
            "# Enable Frequency-Selective Scheduling on MBB Data Bearers (QCI 6, 8, 9 - NEVER QCI 1)",
            "set QciProfilePredefined=qci6 dlResourceAllocationStrategy 1",
            "set QciProfilePredefined=qci8 dlResourceAllocationStrategy 1",
            "set QciProfilePredefined=qci9 dlResourceAllocationStrategy 1",
            "# Enable Ericsson Lean Carrier to suppress inter-cell CRS interference",
            "set OptionalFeatureLicense=EricssonLeanCarrier featureState 1",
            "",
        ])
        maint_cmds.extend([
            f"# [SCENARIO_A] Expand PUCCH CQI capacity during approved Maintenance Window:",
            f"# bl EUtranCellFDD={cell_name}",
            f"# set EUtranCellFDD={cell_name} noOfPucchCqiUsers 320",
            f"# deb EUtranCellFDD={cell_name}",
            "",
        ])

    if "SCENARIO_B" in scen_str:
        online_cmds.extend([
            f"# --- [SCENARIO_B: Accessibility / RACH / PDCCH CCE / Load Control] ---",
            f"pget EUtranCellFDD={cell_name} pmRaAttCbra|pmRaFailCbraMsg1DiscOoc|pmRrcConnEstabFailLic|pmRrcConnEstabFailMpOverload",
            f"get EUtranCellFDD={cell_name} cellRange|pdcchCceOptimizationEnabled|pdcchPowerBoostMax",
            "get CapacityState=ConnectedUsers licenseCapacity|grantedCapacityLevel",
            f"set EUtranCellFDD={cell_name} pdcchCceOptimizationEnabled true",
            f"set EUtranCellFDD={cell_name} pdcchPowerBoostMax 3",
            "set ENodeBFunction=1,DynamicLoadControl=1 dlcEnabled true",
            "",
        ])
        maint_cmds.extend([
            f"# [SCENARIO_B] Expand cellRange if RACH Msg1 Out-Of-Coverage > 5% (Requires Cell Lock):",
            f"# bl EUtranCellFDD={cell_name}",
            f"# set EUtranCellFDD={cell_name} cellRange 15",
            f"# deb EUtranCellFDD={cell_name}",
            "",
        ])

    if "SCENARIO_C" in scen_str:
        online_cmds.extend([
            f"# --- [SCENARIO_C: Retainability / UL Power Control / Transport Audit] ---",
            f"pget EUtranCellFDD={cell_name} pmErabRelAbnormalEnbActUeLost|pmErabRelAbnormalEnbUeLostOos|pmErabRelAbnormalEnbActTnFail|pmErabRelAbnormalMmeAct",
            f"get EUtranCellFDD={cell_name} pZeroNominalPusch|alpha|crsGain",
            "st TermPointToMme",
            "st SctpEndpoint",
            f"set EUtranCellFDD={cell_name} pZeroNominalPusch -96",
            f"set EUtranCellFDD={cell_name} alpha 8",
            "",
        ])
        maint_cmds.extend([
            f"# [SCENARIO_C] Adjust CRS Gain only during approved Maintenance Window:",
            f"# bl EUtranCellFDD={cell_name}",
            f"# set EUtranCellFDD={cell_name} crsGain 0",
            f"# deb EUtranCellFDD={cell_name}",
            "",
        ])

    if "SCENARIO_D" in scen_str:
        online_cmds.extend([
            f"# --- [SCENARIO_D: Intra-Frequency Mobility (EARFCN={earfcn}) Optimization] ---",
            f"pget EUtranCellFDD={cell_name} pmCellHoPrepAttLteIntraF|pmCellHoPrepSuccLteIntraF|pmCellHoExeAttLteIntraF|pmCellHoExeSuccLteIntraF",
            f"get EUtranCellFDD={cell_name},UeMeasControl=1,ReportConfigEUtraBestCell=1 a3Offset|hysteresisA3|timeToTriggerA3",
            "st TermPointToENB",
            f"set EUtranCellFDD={cell_name},EUtranFreqRelation={earfcn},EUtranCellRelation=<Target_Cell> isHoAllowed true",
        ])
        if "Late Handover" in rc_str:
            online_cmds.extend([
                f"set EUtranCellFDD={cell_name},UeMeasControl=1,ReportConfigEUtraBestCell=1 a3Offset 20",
                f"set EUtranCellFDD={cell_name},UeMeasControl=1,ReportConfigEUtraBestCell=1 timeToTriggerA3 40",
                "",
            ])
        else:
            online_cmds.extend([
                f"set EUtranCellFDD={cell_name},UeMeasControl=1,ReportConfigEUtraBestCell=1 hysteresisA3 20",
                f"set EUtranCellFDD={cell_name},UeMeasControl=1,ReportConfigEUtraBestCell=1 timeToTriggerA3 160",
                "",
            ])

    if "SCENARIO_E" in scen_str:
        online_cmds.extend([
            f"# --- [SCENARIO_E: Inter-Frequency Mobility & Measurement Gap Tuning] ---",
            f"pget EUtranCellFDD={cell_name} pmCellHoPrepAttLteInterF|pmCellHoPrepSuccLteInterF|pmCellHoExeAttLteInterF|pmCellHoExeSuccLteInterF",
            f"get EUtranCellFDD={cell_name},UeMeasControl=1 measGapPattern",
            f"set EUtranCellFDD={cell_name},UeMeasControl=1,ReportConfigSearch=1 a2Threshold1RsrpPrim -108",
            f"set EUtranCellFDD={cell_name},UeMeasControl=1,ReportConfigA5=1 a5Threshold1Rsrp -112",
            f"set EUtranCellFDD={cell_name},UeMeasControl=1,ReportConfigA5=1 a5Threshold2Rsrp -106",
            f"set EUtranCellFDD={cell_name},UeMeasControl=1,ReportConfigA5=1 timeToTriggerA5 80",
            "",
        ])

    if "SCENARIO_F" in scen_str:
        online_cmds.extend([
            f"# --- [SCENARIO_F: 3G/4G IRAT Offload & Event B2 Stabilization] ---",
            f"pget EUtranCellFDD={cell_name} pmBadCovEvalReport|pmBadCovSearchEvalReport|pmB2BestCellEvalReport|pmUeCtxtRelSCWcdma|pmUeCtxtRelCsfbWcdma",
            f"set EUtranCellFDD={cell_name},UeMeasControl=1,ReportConfigSearch=1 a2Threshold2RsrpSec -115",
            f"set EUtranCellFDD={cell_name},UeMeasControl=1,ReportConfigB2Utra=1 b2Threshold1Rsrp -118",
            f"set EUtranCellFDD={cell_name},UeMeasControl=1,ReportConfigB2Utra=1 b2Threshold2RscpUtra -95",
            f"set EUtranCellFDD={cell_name},UeMeasControl=1,ReportConfigB2Utra=1 b2Threshold2EcNoUtra -12",
            f"set EUtranCellFDD={cell_name},UeMeasControl=1,ReportConfigB2Utra=1 timeToTriggerB2 640",
            f"set EUtranCellFDD={cell_name},UeMeasControl=1,ReportConfigB2Utra=1 hysteresisB2 20",
            "",
        ])

    if "SCENARIO_G" in scen_str:
        online_cmds.extend([
            f"# --- [SCENARIO_G: VoLTE QCI=1 Bearer Protection, ROHC/TTIB & SRVCC] ---",
            f"pget EUtranCellFDD={cell_name} pmErabEstabAttAddedQci|pmErabEstabSuccAddedQci|pmErabRelAbnormalEnbActQci|pmErabRelAbnormalEnbActPeQci",
            f"pget EUtranCellFDD={cell_name} pmHoPrepAttSrvcc|pmHoPrepSuccSrvcc|pmHoExeAttSrvcc|pmHoExeSuccSrvcc",
            "get QciProfilePredefined=qci1 preEmptionVulnerability|rohcEnabled|dscp",
            "set QciProfilePredefined=qci1 preEmptionVulnerability 0",
            "set QciProfilePredefined=qci1 dscp 46",
            "set OptionalFeatureLicense=Rohc featureState 1",
            "set QciProfilePredefined=qci1 rohcEnabled true",
            "set OptionalFeatureLicense=TtiBundling featureState 1",
            f"set EUtranCellFDD={cell_name} ttiBundlingAtSetup 1",
            f"set EUtranCellFDD={cell_name},UeMeasControl=1,ReportConfigB2Utra=1 b2Threshold1Rsrp -116",
            f"set EUtranCellFDD={cell_name},UeMeasControl=1,ReportConfigB2Utra=1 b2Threshold2RscpUtra -98",
            f"set EUtranCellFDD={cell_name},UeMeasControl=1,ReportConfigB2Utra=1 timeToTriggerB2 320",
            "",
        ])

    if len(maint_cmds) == 3:
        maint_cmds.append(f"# No physical cell-lock (bl/deb) required for {scen_str}; all changes above are online-safe.")
        maint_cmds.append("")

    post_cmds = [
        "# ------------------------------------------------------------------------------",
        "# [4] POST-EXECUTION VERIFICATION & BACKUP COMMIT",
        "# ------------------------------------------------------------------------------",
        f"st EUtranCellFDD={cell_name}",
        "alt",
        f'cvms Post_Opt_{cell_name}_v9 "Post-optimization backup after Antigravity tuning"',
        "gs-",
        "confb-",
    ]

    return "\n".join(lines + online_cmds + maint_cmds + post_cmds)


# ==============================================================================
# 6. SYSTEM-WIDE EXECUTIVE SUMMARY METRICS (WEIGHTED BY NETWORK COUNTERS)
# ==============================================================================

def compute_executive_metrics(df_eval: pd.DataFrame) -> Dict[str, Any]:
    """
    Computes Tier-1 Executive KPI summary metrics across the currently filtered DataFrame
    using true counter-weighted network totals (preventing unweighted cell-average distortion).
    """
    if df_eval.empty:
        return {
            "total_evaluated": 0,
            "stat_insig": 0,
            "worst_count": 0,
            "worst_pct": 0.0,
            "critical_count": 0,
            "degraded_count": 0,
            "sys_cssr_pct": 0.0,
            "sys_dcr_pct": 0.0,
            "sys_hosr_pct": 0.0,
            "sys_volte_dcr_pct": 0.0,
            "sys_volte_cssr_pct": 0.0,
            "sys_avg_cqi": 0.0,
        }

    total_rows = int(len(df_eval))
    insig = int((df_eval["Severity"] == "STAT_INSIGNIFICANT").sum())
    crit = int((df_eval["Severity"] == "CRITICAL").sum())
    deg = int((df_eval["Severity"] == "DEGRADED").sum())
    worst = crit + deg
    valid_denom = max(total_rows - insig, 1)
    worst_pct = (worst / valid_denom) * 100.0

    # Network-wide CSSR (%)
    rrc_den = (get_col(df_eval, "pmRrcConnEstabAtt").sum() - get_col(df_eval, "pmRrcConnEstabAttReatt").sum() - get_col(df_eval, "pmRrcConnEstabFailMmeOvlMos").sum() - get_col(df_eval, "pmRrcConnEstabFailMmeOvlMod").sum())
    rrc_sr = (get_col(df_eval, "pmRrcConnEstabSucc").sum() / rrc_den) if rrc_den > 0 else np.nan
    s1_ovl_col = "pmS1SigConnEstabFailMmeOvlMosCe" if "pmS1SigConnEstabFailMmeOvlMosCe" in df_eval.columns else "pmS1SigConnEstabFailMmeOvlMos"
    s1_den = get_col(df_eval, "pmS1SigConnEstabAtt").sum() - get_col(df_eval, s1_ovl_col).sum()
    s1_sr = (get_col(df_eval, "pmS1SigConnEstabSucc").sum() / s1_den) if s1_den > 0 else np.nan
    erab_den = get_col(df_eval, "pmErabEstabAttInit").sum()
    erab_sr = (get_col(df_eval, "pmErabEstabSuccInit").sum() / erab_den) if erab_den > 0 else np.nan

    if pd.notna(rrc_sr) and pd.notna(s1_sr) and pd.notna(erab_sr):
        sys_cssr = float(rrc_sr * s1_sr * erab_sr * 100.0)
    else:
        sys_cssr = float(df_eval["CSSR_Pct"].dropna().mean()) if "CSSR_Pct" in df_eval.columns and df_eval["CSSR_Pct"].notna().any() else 0.0

    # Network-wide Overall DCR (%)
    dcr_num = get_col(df_eval, "pmErabRelAbnormalEnbAct").sum() + get_col(df_eval, "pmErabRelAbnormalMmeAct").sum()
    dcr_den = get_col(df_eval, "pmErabRelAbnormalEnb").sum() + get_col(df_eval, "pmErabRelNormalEnb").sum() + get_col(df_eval, "pmErabRelMme").sum()
    sys_dcr = float((dcr_num / dcr_den) * 100.0) if dcr_den > 0 else float(df_eval["DCR_Overall_Pct"].dropna().mean() or 0.0)

    # Network-wide Combined HOSR (%)
    p_suc = get_col(df_eval, "pmCellHoPrepSuccLteIntraF").sum() + get_col(df_eval, "pmCellHoPrepSuccLteInterF").sum()
    p_att = get_col(df_eval, "pmCellHoPrepAttLteIntraF").sum() + get_col(df_eval, "pmCellHoPrepAttLteInterF").sum() + get_col(df_eval, "pmHoPrepRejOutUlThres").sum()
    e_suc = get_col(df_eval, "pmCellHoExeSuccLteIntraF").sum() + get_col(df_eval, "pmCellHoExeSuccLteInterF").sum()
    e_att = get_col(df_eval, "pmCellHoExeAttLteIntraF").sum() + get_col(df_eval, "pmCellHoExeAttLteInterF").sum()
    sys_hosr = float((p_suc / p_att) * (e_suc / e_att) * 100.0) if (p_att > 0 and e_att > 0) else float(df_eval["Combined_HOSR_Pct"].dropna().mean() or 0.0)

    # Network-wide VoLTE QCI=1 DCR (%) & CSSR (%)
    vd_num = get_col(df_eval, "pmErabRelAbnormalEnbActQci").sum() + get_col(df_eval, "pmErabRelAbnormalMmeActQci").sum()
    vd_den = get_col(df_eval, "pmErabRelAbnormalEnbQci").sum() + get_col(df_eval, "pmErabRelNormalEnbQci").sum() + get_col(df_eval, "pmErabRelMmeQci").sum()
    sys_vdcr = float((vd_num / vd_den) * 100.0) if vd_den > 0 else float(df_eval["VoLTE_DCR_Rel_Pct"].dropna().mean() or 0.0)

    vc_suc = get_col(df_eval, "pmErabEstabSuccAddedQci").sum() + get_col(df_eval, "pmErabEstabSuccInitQci").sum()
    vc_den = (
        get_col(df_eval, "pmErabEstabAttAddedQci").sum()
        + get_col(df_eval, "pmErabEstabAttInitQci").sum()
        - get_col(df_eval, "pmErabEstabAttAddedHoOngoingQci").sum()
        - get_col(df_eval, "pmErabEstabAttAddedCsfbQci").sum()
        - get_col(df_eval, "pmErabEstabFailAddedErabIdInUseQci").sum()
    )
    sys_vcssr = float((vc_suc / vc_den) * 100.0) if vc_den > 0 else float(df_eval["VoLTE_CSSR_RAN_Pct"].dropna().mean() or 0.0)

    # Network-wide Weighted Avg CQI
    cqi_w_sum = get_col(df_eval, "CQI_Weighted_Sum").sum()
    cqi_v_sum = get_col(df_eval, "CQI_Valid_Samples").sum()
    sys_cqi = float(cqi_w_sum / cqi_v_sum) if cqi_v_sum > 0 else float(df_eval["Avg_CQI"].dropna().mean() or 0.0)

    return {
        "total_evaluated": total_rows,
        "stat_insig": insig,
        "worst_count": worst,
        "worst_pct": round(worst_pct, 1),
        "critical_count": crit,
        "degraded_count": deg,
        "sys_cssr_pct": round(sys_cssr, 2),
        "sys_dcr_pct": round(sys_dcr, 2),
        "sys_hosr_pct": round(sys_hosr, 2),
        "sys_volte_dcr_pct": round(sys_vdcr, 2),
        "sys_volte_cssr_pct": round(sys_vcssr, 2),
        "sys_avg_cqi": round(sys_cqi, 2),
    }


def resolve_default_workbook() -> Optional[str]:
    """Finds Counters.xlsx or Counters_Organized.xlsx in cwd or ./Data/ automatically."""
    candidates = [
        "Counters.xlsx",
        os.path.join("Data", "Counters.xlsx"),
        "Counters_Organized.xlsx",
        os.path.join("Data", "Counters_Organized.xlsx"),
    ]
    for p in candidates:
        if os.path.exists(p):
            return p
    return None


# ==============================================================================
# 7. CLI SELF-TEST ENTRY POINT
# ==============================================================================

def main() -> None:
    parser = argparse.ArgumentParser(description="Antigravity Ericsson LTE RAN ETL & RCA Engine v9.0")
    parser.add_argument("input_file", nargs="?", default=None, help="Path to Counters.xlsx / Counters_Organized.xlsx / CSV")
    parser.add_argument("--json-out", default=None, help="Optional JSON summary output path")
    args = parser.parse_args()

    target = args.input_file or resolve_default_workbook()
    if not target or not os.path.exists(target):
        sys.stderr.write("Error: Counters.xlsx not found in working directory or ./Data/.\n")
        sys.exit(1)

    bundle = load_and_process_Counters(target)
    h_df = bundle["hourly_df"]
    c_df = bundle["cell_df"]
    meta = bundle["meta"]
    h_exec = compute_executive_metrics(h_df)
    c_exec = compute_executive_metrics(c_df)

    print("=" * 78)
    print(f"[OK] ANTIGRAVITY LTE RAN ENGINE v9.0 - SELF-TEST PASSED")
    print("=" * 78)
    print(f"  Workbook          : {meta['filename']} ({target})")
    print(f"  Date Window       : {meta['date_range']}")
    print(f"  Network Scope     : {meta['total_erbs']} ERBS | {meta['total_cells']} Cells | {meta['total_rops']:,} Cell-Hour ROPs")
    print(f"  Carriers (EARFCN) : {meta['carriers']}")
    print("-" * 78)
    print(
        f"  [Hourly ROP View] : {h_exec['total_evaluated']:,} ROPs | "
        f"Worst={h_exec['worst_count']:,} (CRIT={h_exec['critical_count']}, DEG={h_exec['degraded_count']}) | "
        f"STAT_INSIG={h_exec['stat_insig']}"
    )
    print(
        f"  [Cell Agg View]   : {c_exec['total_evaluated']:,} Cells | "
        f"Worst={c_exec['worst_count']:,} (CRIT={c_exec['critical_count']}, DEG={c_exec['degraded_count']}) | "
        f"STAT_INSIG={c_exec['stat_insig']}"
    )
    print(
        f"  [System KPIs]     : CSSR={c_exec['sys_cssr_pct']:.2f}% | Overall DCR={c_exec['sys_dcr_pct']:.2f}% | "
        f"Combined HOSR={c_exec['sys_hosr_pct']:.2f}% | VoLTE DCR={c_exec['sys_volte_dcr_pct']:.2f}% | Avg CQI={c_exec['sys_avg_cqi']:.2f}"
    )
    print("-" * 78)

    worst_sample = c_df[c_df["Severity"].isin(["CRITICAL", "DEGRADED"])].head(3)
    for _, r in worst_sample.iterrows():
        print(f"  -> [{r['Severity']}] {r['EUtranCell Id']} (EARFCN {r['earfcndl']}): {r['Primary_Metric']} | {r['Root_Cause']}")

    if args.json_out:
        with open(args.json_out, "w", encoding="utf-8") as f:
            json.dump({"meta": meta, "hourly_summary": h_exec, "cell_summary": c_exec}, f, indent=2, ensure_ascii=False)


if __name__ == "__main__":
    main()
