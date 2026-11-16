"""Executive reporting sheet builders for Market & Sector Analysis and Company-Wise Rankings."""

from __future__ import annotations

from typing import Any, Dict, List, Optional
import numpy as np
import pandas as pd


def _safe_float(v: Any, default: float = 0.0) -> float:
    try:
        if v is None:
            return default
        f = float(v)
        return f if np.isfinite(f) else default
    except (TypeError, ValueError):
        return default


def _pct_change(a: float, b: float) -> float:
    return (b - a) / a * 100.0 if a != 0 else 0.0


SECTOR_BENCHMARKS_2024_2026 = [
    ("Cement", 442_810_296.0, 474_912_638.0, 251_516_215.7, 266_109_498.7),
    ("Aluminium", 16_810_234.0, 17_842_190.0, 74_574_462.23, 78_081_775.42),
    ("Chlor-Alkali", 4_912_450.0, 5_389_120.0, 9_813_231.17, 10_802_829.90),
    ("Pulp & Paper", 4_683_098.0, 5_072_130.0, 9_667_210.91, 8_738_271.19),
]


def build_sector_summary_table(df: pd.DataFrame, res_df: pd.DataFrame) -> pd.DataFrame:
    """Build clean sector-wise 2024 vs 2026 production, emissions, and CCC credit balance table."""
    m = res_df.copy()
    if "Sector" not in m.columns and "Sector" in df.columns:
        m = m.merge(df[["Reg_Number", "Sector"]].drop_duplicates("Reg_Number"), on="Reg_Number", how="left")
    if "Base_Production" not in m.columns and "Base_Production" in df.columns:
        m = m.merge(df[["Reg_Number", "Base_Production"]].drop_duplicates("Reg_Number"), on="Reg_Number", how="left")

    bench_map = {s: (p24, p26, e24, e26) for s, p24, p26, e24, e26 in SECTOR_BENCHMARKS_2024_2026}
    rows = []
    for sec in ["Cement", "Aluminium", "Chlor-Alkali", "Pulp & Paper"]:
        sub = m[m["Sector"].astype(str).str.strip() == sec]
        n_plants = int(len(sub))
        p24_b, p26_b, e24_b, e26_b = bench_map[sec]

        prod_2024 = float(sub["Base_Production"].fillna(0).sum()) if ("Base_Production" in sub.columns and sub["Base_Production"].fillna(0).sum() > 0) else p24_b
        prod_2026 = float(sub["Actual_Production"].fillna(0).sum()) if ("Actual_Production" in sub.columns and sub["Actual_Production"].fillna(0).sum() > 0) else p26_b

        em_2024 = float((sub["Baseline_GEI"].fillna(0) * sub["Base_Production"].fillna(0)).sum()) if ("Base_Production" in sub.columns and (sub["Baseline_GEI"].fillna(0) * sub["Base_Production"].fillna(0)).sum() > 0) else e24_b
        em_2026 = float((sub["Mode_GEI"].fillna(0) * sub["Actual_Production"].fillna(0)).sum()) if ("Mode_GEI" in sub.columns and (sub["Mode_GEI"].fillna(0) * sub["Actual_Production"].fillna(0)).sum() > 0) else e26_b

        purch = float(sub["Expected_Purchases_Demand"].sum()) if len(sub) else 0.0
        cred = float(sub["Expected_Credits_Surplus"].sum()) if len(sub) else 0.0
        net_ccc = cred - purch
        verdict = "Net Seller (Surplus)" if net_ccc > 0 else ("Net Buyer (Deficit)" if net_ccc < 0 else "Balanced")

        rows.append({
            "Sector": sec,
            "Facilities": n_plants,
            "Production_2024": prod_2024,
            "Production_2026": prod_2026,
            "Prod_Change_Pct": _pct_change(prod_2024, prod_2026),
            "Emissions_2024_tCO2": em_2024,
            "Emissions_2026_tCO2": em_2026,
            "Emissions_Change_Pct": _pct_change(em_2024, em_2026),
            "Expected_Purchases_CCC": purch,
            "Expected_Credits_CCC": cred,
            "Net_CCC_Balance": net_ccc,
            "Decarbonisation_Verdict": verdict,
        })
    return pd.DataFrame(rows)


def build_market_analysis_sheet(
    df: pd.DataFrame,
    res_df: pd.DataFrame,
    summary: dict,
) -> pd.DataFrame:
    rows: List[Dict[str, Any]] = []
    n = len(df)
    total_purch = float(res_df["Expected_Purchases_Demand"].sum()) if len(res_df) else 0.0
    total_cred = float(res_df["Expected_Credits_Surplus"].sum()) if len(res_df) else 0.0
    net = summary.get("total_mean_net", total_cred - total_purch) if summary else (total_cred - total_purch)
    supply = summary.get("total_gross_supply", total_cred) if summary else total_cred
    demand = summary.get("total_gross_demand", total_purch) if summary else total_purch
    denom = demand if demand > 0 else 0.01

    sec_tbl = build_sector_summary_table(df, res_df)
    tot_p24 = float(sec_tbl["Production_2024"].sum())
    tot_p26 = float(sec_tbl["Production_2026"].sum())
    tot_e24 = float(sec_tbl["Emissions_2024_tCO2"].sum())
    tot_e26 = float(sec_tbl["Emissions_2026_tCO2"].sum())

    rows.append({"Section": "Macro Market KPIs", "Category": "Market", "Metric": "Total Demand (CCC)", "Value": _safe_float(demand), "Unit": "CCC"})
    rows.append({"Section": "Macro Market KPIs", "Category": "Market", "Metric": "Total Supply (CCC)", "Value": _safe_float(supply), "Unit": "CCC"})
    rows.append({"Section": "Macro Market KPIs", "Category": "Market", "Metric": "Net Supply/Demand (CCC)", "Value": _safe_float(net), "Unit": "CCC"})
    rows.append({"Section": "Macro Market KPIs", "Category": "Market", "Metric": "Net Supply/Demand (%)", "Value": _safe_float(net / denom * 100.0), "Unit": "%"})
    rows.append({"Section": "Macro Market KPIs", "Category": "Market", "Metric": "Liquidity Coverage Ratio (Supply / Demand %)", "Value": _safe_float(supply / denom * 100.0), "Unit": "%"})
    rows.append({"Section": "Macro Market KPIs", "Category": "Activity", "Metric": "2024 to 2026 Change in Production (%)", "Value": _pct_change(tot_p24, tot_p26), "Unit": "%"})
    rows.append({"Section": "Macro Market KPIs", "Category": "Activity", "Metric": "2024 to 2026 Change in Emissions (%)", "Value": _pct_change(tot_e24, tot_e26), "Unit": "%"})

    target_gei = float(df["Target_GEI"].mean()) if n else 0.5
    base_gei = float(df["Baseline_GEI"].mean()) if n else 0.5
    rows.append({"Section": "Macro Market KPIs", "Category": "Counterfactual", "Metric": "No Improvement in Intensity - (Target-GEI - Baseline-GEI) x Production_2026", "Value": _safe_float((target_gei - base_gei) * tot_p26), "Unit": "CCC"})
    rows.append({"Section": "Macro Market KPIs", "Category": "Counterfactual", "Metric": "No Improvement in Intensity and Production - (Target-GEI - Baseline-GEI) x Production_2024", "Value": _safe_float((target_gei - base_gei) * tot_p24), "Unit": "CCC"})

    for r in sec_tbl.itertuples():
        rows.append({"Section": "Sector Analysis", "Category": r.Sector, "Metric": "Prod Change 2024-2026 (%)", "Value": _safe_float(r.Prod_Change_Pct), "Unit": "%"})
        rows.append({"Section": "Sector Analysis", "Category": r.Sector, "Metric": "Emissions Change 2024-2026 (%)", "Value": _safe_float(r.Emissions_Change_Pct), "Unit": "%"})
        rows.append({"Section": "Sector Analysis", "Category": r.Sector, "Metric": "Expected Purchases (CCC)", "Value": _safe_float(r.Expected_Purchases_CCC), "Unit": "CCC"})
        rows.append({"Section": "Sector Analysis", "Category": r.Sector, "Metric": "Expected Credits (CCC)", "Value": _safe_float(r.Expected_Credits_CCC), "Unit": "CCC"})
        rows.append({"Section": "Sector Analysis", "Category": r.Sector, "Metric": "Net CCC Balance", "Value": _safe_float(r.Net_CCC_Balance), "Unit": "CCC"})

    return pd.DataFrame(rows)


def build_company_wise_sheet(
    res_df: pd.DataFrame,
    df: pd.DataFrame,
    selected_sector: Optional[str] = None,
) -> pd.DataFrame:
    """Build Company-Wise & Top 5 Rankings across all 4 CCTS sectors."""
    m = res_df.copy()
    for col in ["Company", "Sector", "Plant"]:
        if col not in m.columns and col in df.columns:
            m = m.merge(df[["Reg_Number", col]].drop_duplicates("Reg_Number"), on="Reg_Number", how="left")
    m["Company"] = m["Company"].fillna("Unknown")
    m["Sector"] = m["Sector"].fillna("Unknown")

    if selected_sector and str(selected_sector).strip().lower() != "all":
        m = m[m["Sector"].astype(str).str.strip() == str(selected_sector).strip()]

    label = f"Company-Wise Analysis (Filtered by Sector: {selected_sector or 'ALL'})"
    rows: List[Dict[str, Any]] = []
    rows.append({"Section": label, "Category": "Header", "Metric": "Total Purchases (CCC)", "Value": float(m["Expected_Purchases_Demand"].sum()), "Unit": "CCC"})
    rows.append({"Section": label, "Category": "Header", "Metric": "Total Credits (CCC)", "Value": float(m["Expected_Credits_Surplus"].sum()), "Unit": "CCC"})

    agg = m.groupby(["Company", "Sector"], as_index=False).agg(
        Total_Purchases_Demand=("Expected_Purchases_Demand", "sum"),
        Total_Credits_Surplus=("Expected_Credits_Surplus", "sum"),
        Count=("Reg_Number", "count"),
    ).sort_values("Total_Purchases_Demand", ascending=False)

    for r in agg.itertuples():
        rows.append({"Section": label, "Category": f"Company Detail ({r.Sector})", "Metric": r.Company, "Value": _safe_float(r.Total_Purchases_Demand), "Unit": "CCC"})
        rows.append({"Section": label, "Category": f"Company Detail ({r.Sector})", "Metric": f"{r.Company} Surplus", "Value": _safe_float(r.Total_Credits_Surplus), "Unit": "CCC"})
        rows.append({"Section": label, "Category": f"Company Detail ({r.Sector})", "Metric": f"{r.Company} Facilities", "Value": int(r.Count), "Unit": "plants"})

    top_buyers = agg.sort_values("Total_Purchases_Demand", ascending=False).head(5)
    for i, r in enumerate(top_buyers.itertuples(), 1):
        rows.append({"Section": label, "Category": "Top 5 Buyers", "Metric": f"Top {i} ({r.Sector})", "Value": r.Company, "Unit": ""})
        rows.append({"Section": label, "Category": "Top 5 Buyers", "Metric": f"Top {i} Purchases", "Value": _safe_float(r.Total_Purchases_Demand), "Unit": "CCC"})

    top_sellers = agg.sort_values("Total_Credits_Surplus", ascending=False).head(5)
    for i, r in enumerate(top_sellers.itertuples(), 1):
        rows.append({"Section": label, "Category": "Top 5 Sellers", "Metric": f"Top {i} ({r.Sector})", "Value": r.Company, "Unit": ""})
        rows.append({"Section": label, "Category": "Top 5 Sellers", "Metric": f"Top {i} Credits", "Value": _safe_float(r.Total_Credits_Surplus), "Unit": "CCC"})

    return pd.DataFrame(rows)
