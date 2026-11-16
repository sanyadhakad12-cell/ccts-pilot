"""MACC parsing, plant-scaled abatement, Abate-vs-Buy decision, and EU/UK CBAM calculator."""

from __future__ import annotations

from typing import Dict, List, Optional, Tuple

import numpy as np
import pandas as pd

from .data_pipeline import load_macc, scale_macc_to_plant, get_sector_emissions_mtco2


SECTOR_ALIASES = {
    "aluminium sector": "Aluminium",
    "chlor alkali": "Chlor-Alkali",
    "chlor-alkali": "Chlor-Alkali",
    "paper": "Pulp & Paper",
    "pulp & paper": "Pulp & Paper",
    "cement": "Cement",
}

def compute_plant_macc_table(
    sector: str,
    plant_baseline_emissions_tco2: float,
    macc_data: Dict[str, pd.DataFrame],
    sector_emissions_mtco2: Dict[str, float],
) -> pd.DataFrame:
    """Build the plant-level MACC table for a given sector.

    % Abatement Share = Annual Abatement (MtCO2) / Sector Total 2024 Emissions (MtCO2)
    Plant abatement capacity (tCO2) = % share * plant baseline emissions (tCO2).
    """
    sector_lower = sector.strip().lower()
    sector_normalized = SECTOR_ALIASES.get(sector_lower, sector.strip())
    sector_total = sector_emissions_mtco2.get(sector_normalized)
    macc_df = macc_data.get(sector_normalized)
    if macc_df is None or macc_df.empty or sector_total is None or sector_total == 0:
        return pd.DataFrame()
    scaled = scale_macc_to_plant(macc_df, plant_baseline_emissions_tco2, sector_total)
    # Ensure Plant_Abatement_Capacity_tCO2 column exists and is float
    if "Plant_Abatement_Capacity_tCO2" not in scaled.columns:
        scaled["Plant_Abatement_Capacity_tCO2"] = 0.0
    else:
        scaled["Plant_Abatement_Capacity_tCO2"] = scaled["Plant_Abatement_Capacity_tCO2"].astype(float).fillna(0.0)
    return scaled


def abate_vs_buy_decision(
    macc_table: pd.DataFrame,
    compliance_deficit_tco2: float,
    expected_mcp_inr: float,
) -> dict:
    """Split the plant's compliance deficit into Abate Internally vs Buy CCCs.

    Levers with Cost_INR_tCO2 <= expected MCP are abated internally (up to their
    capacity); the residual deficit is bought on the market at Expected MCP.
    Returns net ₹ savings vs buying 100% on market.
    """
    if macc_table is None or macc_table.empty:
        return {
            "deficit_tco2": compliance_deficit_tco2,
            "abated_tco2": 0.0,
            "bought_tco2": compliance_deficit_tco2,
            "abatement_cost_inr": 0.0,
            "purchase_cost_inr": compliance_deficit_tco2 * expected_mcp_inr,
            "net_savings_inr": 0.0,
            "macc_table": macc_table,
        }

    df = macc_table.copy().sort_values("Cost_INR_tCO2")
    df["Abate_Amount_tCO2"] = 0.0
    remaining = compliance_deficit_tco2
    total_abatement_cost = 0.0

    for idx, row in df.iterrows():
        if remaining <= 0:
            break
        cost = row["Cost_INR_tCO2"]
        cap = row["Plant_Abatement_Capacity_tCO2"]
        if cost is None or cap is None or cap <= 0:
            continue
        take = min(cap, remaining)
        df.at[idx, "Abate_Amount_tCO2"] = take
        total_abatement_cost += take * cost
        remaining -= take

    abated = compliance_deficit_tco2 - remaining
    bought = remaining
    purchase_cost = bought * expected_mcp_inr
    # Baseline: buy everything at MCP
    baseline_cost = compliance_deficit_tco2 * expected_mcp_inr
    net_savings = baseline_cost - (total_abatement_cost + purchase_cost)

    return {
        "deficit_tco2": compliance_deficit_tco2,
        "abated_tco2": abated,
        "bought_tco2": bought,
        "abatement_cost_inr": total_abatement_cost,
        "purchase_cost_inr": purchase_cost,
        "net_savings_inr": net_savings,
        "macc_table": df,
    }


def compute_cbam_liability(
    plant_row: pd.Series,
    expected_mcp_inr: float,
    year_rules: dict,
    eu_export_pct: float = 0.0,
    uk_export_pct: float = 0.0,
    eu_ets_price_eur: Optional[float] = None,
    uk_cbam_rate_gbp: Optional[float] = None,
) -> dict:
    """Combined domestic + cross-border carbon exposure per installation.

    Domestic: expected MCP * emissions (tCO2) credited against CCTS relief paid.
    EU CBAM: EU ETS price * EU-exported output emissions, minus domestic relief credited.
    UK CBAM: UK CBAM rate * UK-exported output emissions (direct Scope 1), minus domestic relief credited.
    UK CBAM applies to Aluminium, Cement, Fertilisers, Hydrogen, Iron & Steel (effective 1 Jan 2027).
    """
    currency = year_rules.get("currency", {})
    inr_per_eur = currency.get("inr_per_eur", 90.0)
    inr_per_gbp = currency.get("inr_per_gbp", 105.0)
    eu_ets_price_eur = eu_ets_price_eur if eu_ets_price_eur is not None else currency.get("eu_ets_price_eur_per_tco2", 85.0)
    uk_cbam_rate_gbp = uk_cbam_rate_gbp if uk_cbam_rate_gbp is not None else currency.get("uk_cbam_rate_gbp_per_tco2", 45.0)

    sector = plant_row.get("Sector", "")
    actual_prod = plant_row.get("Actual_Production", 0) or 0
    baseline_gei = plant_row.get("Baseline_GEI", 0) or 0
    emissions_tco2 = actual_prod * baseline_gei

    # Domestic CCTS liability / relief
    domestic_cost_inr = emissions_tco2 * expected_mcp_inr

    # EU CBAM
    eu_emissions = emissions_tco2 * (eu_export_pct / 100.0)
    eu_cost_inr = eu_emissions * eu_ets_price_eur * inr_per_eur
    eu_cbam_net_inr = max(0.0, eu_cost_inr - domestic_cost_inr * (eu_export_pct / 100.0))

    # UK CBAM (sector-restricted, effective 2027)
    uk_cbam_sectors = year_rules.get("2027", {}).get("uk_cbam_sectors", [])
    uk_applicable = sector in uk_cbam_sectors
    uk_emissions = emissions_tco2 * (uk_export_pct / 100.0) if uk_applicable else 0.0
    uk_cost_inr = uk_emissions * uk_cbam_rate_gbp * inr_per_gbp
    uk_cbam_net_inr = max(0.0, uk_cost_inr - domestic_cost_inr * (uk_export_pct / 100.0)) if uk_applicable else 0.0

    total_cbam_inr = eu_cbam_net_inr + uk_cbam_net_inr
    total_exposure_inr = domestic_cost_inr + total_cbam_inr

    return {
        "Sector": sector,
        "Emissions_tCO2": emissions_tco2,
        "Domestic_CCTS_Cost_INR": domestic_cost_inr,
        "EU_Export_Pct": eu_export_pct,
        "EU_Emissions_tCO2": eu_emissions,
        "EU_ETS_Price_EUR": eu_ets_price_eur,
        "EU_CBAM_Cost_INR": eu_cost_inr,
        "EU_CBAM_Net_INR": eu_cbam_net_inr,
        "UK_Export_Pct": uk_export_pct,
        "UK_Applicable": uk_applicable,
        "UK_Emissions_tCO2": uk_emissions,
        "UK_CBAM_Rate_GBP": uk_cbam_rate_gbp,
        "UK_CBAM_Cost_INR": uk_cost_inr,
        "UK_CBAM_Net_INR": uk_cbam_net_inr,
        "Total_CBAM_Net_INR": total_cbam_inr,
        "Total_Exposure_INR": total_exposure_inr,
    }


def get_marginal_abatement_cap(
    macc_table: pd.DataFrame,
    deficit_tco2: float,
    price_floor: float = 600.0,
    price_cap: float = 1600.0,
) -> float:
    """Walk up sorted MACC levers until cumulative abatement >= deficit_tco2."""
    if macc_table is None or macc_table.empty or deficit_tco2 <= 0:
        return float(price_cap)
    df = macc_table.dropna(subset=["Cost_INR_tCO2"]).sort_values("Cost_INR_tCO2")
    cum = 0.0
    for _, row in df.iterrows():
        cap = float(row.get("Plant_Abatement_Capacity_tCO2", 0.0) or 0.0)
        cost = float(row["Cost_INR_tCO2"])
        if cap <= 0:
            continue
        cum += cap
        if cum >= deficit_tco2:
            return float(max(price_floor, min(price_cap, cost)))
    return float(price_cap)
