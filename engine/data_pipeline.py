"""Data pipeline: Excel/CSV ingestion, Main Sheet rebuild, sector mapping, isolation."""

from __future__ import annotations

import json
import math
import os
from typing import Any, Dict, List, Optional, Tuple

import numpy as np
import openpyxl
import pandas as pd

APP_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CONFIG_DIR = os.path.join(APP_DIR, "config")
EXCEL_PATH = os.path.join(APP_DIR, "3. Data sheet.xlsx")
MACC_PATH = os.path.join(APP_DIR, "MACC for cline.xlsx")
CSV_PATH = os.path.join(APP_DIR, "market_data_248_plants.csv")

MAIN_SHEET_COLS = [
    "Company",
    "S_No",
    "Plant",
    "Reg_Number",
    "Baseline_GEI",
    "Target_GEI",
    "Base_Production",
    "Actual_Production",
    "Min_GEI",
    "Mode_GEI",
    "Max_GEI",
]

SECTOR_KEYWORDS = {
    "Cement": ["cement"],
    "Aluminium": ["aluminium", "aluminum", "hindalco", "vedanta", "nalco"],
    "Chlor-Alkali": ["chlor alkali", "chlor-alkali", "soda", "caustic"],
    "Pulp & Paper": ["paper", "pulp"],
}
DEFAULT_SECTOR = "Cement"


def clean_num(val: Any) -> Optional[float]:
    if val is None:
        return None
    if isinstance(val, (int, float)) and not isinstance(val, bool):
        v = float(val)
        return v if math.isfinite(v) else None
    v = str(val).replace(",", "").replace("%", "").strip()
    if v == "" or v.lower() in ("nan", "none", "null", "-", "n/a", "na"):
        return None
    try:
        f = float(v)
        return f if math.isfinite(f) else None
    except ValueError:
        return None


def _load_json(filename: str) -> dict:
    with open(os.path.join(CONFIG_DIR, filename), "r", encoding="utf-8") as f:
        return json.load(f)


def read_sheet_df(wb: openpyxl.workbook.Workbook, sheet_name: str, header_row: int = 1) -> pd.DataFrame:
    if sheet_name not in wb.sheetnames:
        return pd.DataFrame()
    ws = wb[sheet_name]
    data = list(ws.iter_rows(values_only=True))
    if not data or len(data) < header_row:
        return pd.DataFrame()
    header = [str(h).strip() if h is not None else f"Col_{i}" for i, h in enumerate(data[header_row - 1])]
    seen: Dict[str, int] = {}
    uniq: List[str] = []
    for h in header:
        if h in seen:
            seen[h] += 1
            uniq.append(f"{h}_{seen[h]}")
        else:
            seen[h] = 0
            uniq.append(h)
    rows = data[header_row:]
    return pd.DataFrame(rows, columns=uniq)


def load_main_sheet_from_excel(excel_path: str = EXCEL_PATH) -> Tuple[pd.DataFrame, openpyxl.workbook.Workbook]:
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    if "Main Sheet" not in wb.sheetnames:
        return pd.DataFrame(), wb
    ws = wb["Main Sheet"]
    records = []
    for row in ws.iter_rows(min_row=3, values_only=True):
        reg = row[3] if len(row) > 3 else None
        s_no = row[1] if len(row) > 1 else None
        if reg is None or str(reg).strip() == "":
            continue
        if s_no is None or str(s_no).strip() == "":
            continue
        reg_s = str(reg).strip()
        if reg_s.lower() in ("reg. number", "reg number", "registration number", "nan"):
            continue
        company = str(row[0]).strip() if row[0] is not None else ""
        plant = str(row[2]).strip() if len(row) > 2 and row[2] is not None else ""
        base_gei = clean_num(row[4]) if len(row) > 4 else None
        target_gei = clean_num(row[5]) if len(row) > 5 else None
        base_prod = clean_num(row[6]) if len(row) > 6 else 0.0
        actual_prod = clean_num(row[7]) if len(row) > 7 else None
        min_gei = clean_num(row[8]) if len(row) > 8 else None
        mode_gei = clean_num(row[9]) if len(row) > 9 else None
        max_gei = clean_num(row[10]) if len(row) > 10 else None

        if base_gei is None or target_gei is None or actual_prod is None:
            continue
        if min_gei is None or mode_gei is None or max_gei is None:
            continue

        lo, hi = min(min_gei, max_gei), max(min_gei, max_gei)
        mode = max(lo, min(hi, mode_gei))

        records.append({
            "Company": company,
            "S_No": clean_num(s_no),
            "Plant": plant,
            "Reg_Number": reg_s,
            "Baseline_GEI": base_gei,
            "Target_GEI": target_gei,
            "Base_Production": base_prod,
            "Actual_Production": actual_prod,
            "Min_GEI": lo,
            "Mode_GEI": mode,
            "Max_GEI": hi,
        })
    df = pd.DataFrame(records, columns=MAIN_SHEET_COLS)
    return df, wb


def extract_decarb_trends(wb: openpyxl.workbook.Workbook) -> Dict[str, dict]:
    trends: Dict[str, dict] = {}
    sheet = "decarbonisation trend2% adjuste"
    if sheet not in wb.sheetnames:
        for name in wb.sheetnames:
            if "decarbon" in name.lower():
                sheet = name
                break
    if sheet not in wb.sheetnames:
        return trends
    ws = wb[sheet]
    rows = list(ws.iter_rows(values_only=True))
    if len(rows) < 2:
        return trends
    header = [str(c).strip() if c is not None else "" for c in rows[1]]

    def colidx(names: List[str]) -> Optional[int]:
        for i, h in enumerate(header):
            if h in names:
                return i
        return None

    idx_company = colidx(["Company"])
    idx_best = colidx(["Best Case (%) - Minimum YoY change observed"])
    idx_med = colidx(["Median / Most Likely (%) - CAGR trend"])
    idx_worst = colidx(["Worst Case (%) - Maximum YoY jump observed"])
    idx_f26 = colidx(["FY 2025-26 2% adjusted", "FY 2025-26"])
    for r in rows[2:]:
        if idx_company is None:
            continue
        company = r[idx_company] if idx_company is not None and idx_company < len(r) else None
        if company is None or str(company).strip() == "":
            continue
        company = str(company).strip()

        def g(i: Optional[int]) -> Optional[float]:
            if i is None or i >= len(r):
                return None
            return clean_num(r[i])

        trends[company] = {
            "best_yoy": g(idx_best),
            "median_yoy": g(idx_med),
            "worst_yoy": g(idx_worst),
            "f26_gei": g(idx_f26),
        }
    return trends


def extract_sector_medians(wb: openpyxl.workbook.Workbook) -> Dict[str, dict]:
    medians: Dict[str, dict] = {}
    sheet = " medians"
    if sheet not in wb.sheetnames:
        for name in wb.sheetnames:
            if name.strip().lower() == "medians":
                sheet = name
                break
    if sheet not in wb.sheetnames:
        return medians
    ws = wb[sheet]
    rows = list(ws.iter_rows(values_only=True))
    if len(rows) < 2:
        return medians
    for r in rows[1:]:
        sector = r[1] if len(r) > 1 else None
        if sector is None or str(sector).strip() == "":
            continue
        sector = str(sector).strip()

        def g(i: Optional[int]) -> Optional[float]:
            if i is None or i >= len(r):
                return None
            return clean_num(r[i])
        medians[sector] = {
            "cap_2024": g(2),
            "cap_2026": g(3),
            "cuf_2024": g(7),
            "cuf_2026": g(8),
            "best_yoy": g(12),
            "median_yoy": g(13),
            "worst_yoy": g(14),
        }
    return medians


def _detect_sector(reg_number: str = "", company_name: str = "") -> str:
    reg = str(reg_number or "").strip().upper()
    if reg.startswith("CNA"):
        return "Chlor-Alkali"
    if reg.startswith("PNP"):
        return "Pulp & Paper"
    if reg.startswith("AL"):
        return "Aluminium"
    if reg.startswith("CM"):
        return "Cement"
    c = (company_name or "").lower()
    for sector, keywords in SECTOR_KEYWORDS.items():
        for kw in keywords:
            if kw in c:
                return sector
    return DEFAULT_SECTOR


def rebuild_main_sheet_from_sources(
    excel_path: str = EXCEL_PATH, compliance_year: int = 2026
) -> Tuple[pd.DataFrame, openpyxl.workbook.Workbook]:
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    trends = extract_decarb_trends(wb)
    sector_medians = extract_sector_medians(wb)

    raw_df = read_sheet_df(wb, "Raw data, Calculations", header_row=2)
    records = []

    def find_col(df: pd.DataFrame, candidates: List[str]) -> Optional[str]:
        for c in df.columns:
            cs = str(c).strip().lower()
            for cand in candidates:
                if cand in cs:
                    return c
        return None

    col_company = find_col(raw_df, ["obligated entity"])
    col_reg = find_col(raw_df, ["reg. number", "reg number", "registration number"])
    col_base_gei = find_col(raw_df, ["baseline ghg intensity", "baseline gei"])
    col_target = find_col(raw_df, ["target 2025", "target gei"])
    col_base_prod = find_col(raw_df, ["2024 equivalent output", "baseline output"])
    col_actual_prod = find_col(raw_df, ["2026 equivalent product"])

    for _, row in raw_df.iterrows():
        reg = row[col_reg] if col_reg else None
        if reg is None or str(reg).strip() == "":
            continue
        reg_s = str(reg).strip()
        company = str(row[col_company]).strip() if col_company else ""

        base_gei = clean_num(row[col_base_gei]) if col_base_gei else None
        target_gei = clean_num(row[col_target]) if col_target else None
        base_prod = clean_num(row[col_base_prod]) if col_base_prod else 0.0
        actual_prod = clean_num(row[col_actual_prod]) if col_actual_prod else None

        if base_gei is None or target_gei is None or actual_prod is None:
            continue

        tr = trends.get(company, {})
        med = sector_medians.get(DEFAULT_SECTOR, {})
        best_yoy = tr.get("best_yoy")
        med_yoy = tr.get("median_yoy")
        worst_yoy = tr.get("worst_yoy")
        if best_yoy is None:
            best_yoy = med.get("best_yoy", 0.0)
        if med_yoy is None:
            med_yoy = med.get("median_yoy", 0.0)
        if worst_yoy is None:
            worst_yoy = med.get("worst_yoy", 0.0)

        min_gei = base_gei * (1 + best_yoy)
        mode_gei = base_gei * (1 + med_yoy)
        max_gei = base_gei * (1 + worst_yoy)
        lo, hi = min(min_gei, max_gei), max(min_gei, max_gei)
        mode = max(lo, min(hi, mode_gei))

        years_diff = compliance_year - 2026
        escalation = (1 + med_yoy) ** years_diff if med_yoy is not None else 1.0
        target_year = target_gei * escalation

        records.append({
            "Company": company,
            "S_No": clean_num(row[col_reg]) if col_reg else None,
            "Plant": str(row[col_company]).strip() if col_company else "",
            "Reg_Number": reg_s,
            "Baseline_GEI": base_gei,
            "Target_GEI": target_year,
            "Base_Production": base_prod,
            "Actual_Production": actual_prod,
            "Min_GEI": lo,
            "Mode_GEI": mode,
            "Max_GEI": hi,
        })

    df = pd.DataFrame(records, columns=MAIN_SHEET_COLS)
    if df.empty:
        df, _ = load_main_sheet_from_excel(excel_path)
    return df, wb


def apply_sidebar_adjustments(
    df: pd.DataFrame, compliance_year: int, banking_pct: float, cuf_shock_pct: float
) -> pd.DataFrame:
    df = df.copy()
    df["Actual_Production"] = df["Actual_Production"] * (1.0 + cuf_shock_pct / 100.0)
    df["Banking_Pct"] = banking_pct
    df["Compliance_Year"] = compliance_year
    return df


def add_sector_column(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df["Sector"] = [
        _detect_sector(r, c)
        for r, c in zip(df.get("Reg_Number", [""] * len(df)), df.get("Company", [""] * len(df)))
    ]
    return df


def get_sector_emissions_mtco2() -> Dict[str, float]:
    year_rules = _load_json("year_rules.json")
    fallback = dict(year_rules.get("sector_total_2024_emissions_mtco2", {}))
    try:
        wb = openpyxl.load_workbook(EXCEL_PATH, data_only=True)
        if "analysis" in wb.sheetnames:
            ws = wb["analysis"]
            rows = list(ws.iter_rows(values_only=True))
            for r in rows[1:]:
                name = r[1] if len(r) > 1 else None
                if name is None:
                    continue
                name = str(name).strip()
                emissions = clean_num(r[5]) if len(r) > 5 else None
                if emissions is None:
                    continue
                mt = emissions / 1_000_000.0
                key = {"Chlor Alkali": "Chlor-Alkali"}.get(name, name)
                fallback[key] = round(mt, 2)
    except Exception:
        pass
    return fallback


def load_macc(macc_path: str = MACC_PATH) -> Dict[str, pd.DataFrame]:
    wb = openpyxl.load_workbook(macc_path, data_only=True)
    macc: Dict[str, pd.DataFrame] = {}
    tab_to_sector = {
        "cement": "Cement",
        "aluminium": "Aluminium",
        "chlor alkali": "Chlor-Alkali",
        "paper": "Pulp & Paper",
    }
    for ws in wb.worksheets:
        sector = tab_to_sector.get(ws.title.strip().lower(), ws.title.strip())
        rows = list(ws.iter_rows(values_only=True))
        if len(rows) < 6:
            continue
        header = [str(c).strip().replace("\u200b", "") if c is not None else "" for c in rows[4]]
        data_rows = rows[5:]
        records = []
        for r in data_rows:
            lever = r[0]
            if lever is None or str(lever).strip() == "":
                continue
            records.append({
                "Sector": sector,
                "Lever": str(lever).strip(),
                "Category": str(r[1]).strip() if len(r) > 1 and r[1] is not None else "",
                "Cost_USD_tCO2": clean_num(r[2]) if len(r) > 2 else None,
                "Cost_INR_tCO2": clean_num(r[3]) if len(r) > 3 else None,
                "Annual_Abatement_MtCO2": clean_num(r[4]) if len(r) > 4 else None,
                "Cumulative_MtCO2": clean_num(r[5]) if len(r) > 5 else None,
            })
        if records:
            macc[sector] = pd.DataFrame(records)
    return macc


def scale_macc_to_plant(
    macc_df: pd.DataFrame,
    plant_baseline_emissions_tco2: float,
    sector_total_mtco2: float,
) -> pd.DataFrame:
    if macc_df is None or macc_df.empty or sector_total_mtco2 is None or sector_total_mtco2 == 0:
        return macc_df.copy() if macc_df is not None else pd.DataFrame()
    df = macc_df.copy()
    df["Abatement_Share"] = df["Annual_Abatement_MtCO2"] / sector_total_mtco2
    df["Plant_Abatement_Capacity_tCO2"] = df["Abatement_Share"] * plant_baseline_emissions_tco2
    df["Plant_Abatement_Capacity_MtCO2"] = df["Plant_Abatement_Capacity_tCO2"] / 1_000_000.0
    return df


def isolate_for_company(
    df: pd.DataFrame, company: str
) -> Tuple[pd.DataFrame, pd.DataFrame]:
    own = df[df["Company"] == company].copy()
    others = df[df["Company"] != company].copy()
    return own, others


def mask_competitors(df_others: pd.DataFrame) -> pd.DataFrame:
    if df_others.empty:
        return df_others.copy()
    masked = df_others.copy()
    masked["Plant"] = "[Masked - Competitor]"
    masked["Reg_Number"] = "[Masked]"
    masked["Baseline_GEI"] = np.nan
    masked["Actual_Production"] = np.nan
    masked["Min_GEI"] = np.nan
    masked["Mode_GEI"] = np.nan
    masked["Max_GEI"] = np.nan
    return masked






