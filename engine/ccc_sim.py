"""10,000-iteration CCC Monte Carlo simulation (existing 2026 logic preserved)."""

from __future__ import annotations

from typing import Dict, Tuple

import numpy as np
import pandas as pd

N_CCC_SIM = 10_000
SEED = 42


def run_ccc_monte_carlo(
    df: pd.DataFrame, n_sim: int = N_CCC_SIM, seed: int = SEED
) -> Tuple[pd.DataFrame, dict, Dict[str, np.ndarray]]:
    """For each plant, sample achieved intensity ~ triangular(min, mode, max) and
    compute CCC = (Target − Achieved) × Production.

    Returns per-plant stats DataFrame, portfolio summary dict, and raw per-plant
    balance matrix (dict reg -> array).
    """
    rng = np.random.default_rng(seed)
    plants = df.reset_index(drop=True)
    n = len(plants)

    target = plants["Target_GEI"].to_numpy(dtype=float)
    prod = plants["Actual_Production"].to_numpy(dtype=float)
    lo = plants["Min_GEI"].to_numpy(dtype=float)
    mode = plants["Mode_GEI"].to_numpy(dtype=float)
    hi = plants["Max_GEI"].to_numpy(dtype=float)

    # Vectorised sampling: shape (n_sim, n_plants).
    # np.random.triangular errors when lo == hi, so handle equal-bound plants
    # by assigning the constant bound directly.
    achieved = np.empty((n_sim, n), dtype=float)
    eq_mask = lo == hi
    if np.any(~eq_mask):
        achieved[:, ~eq_mask] = rng.triangular(
            lo[~eq_mask], mode[~eq_mask], hi[~eq_mask],
            size=(n_sim, int((~eq_mask).sum())))
    if np.any(eq_mask):
        achieved[:, eq_mask] = lo[eq_mask]

    bal = (target[None, :] - achieved) * prod[None, :]  # (n_sim, n_plants)

    gross_shortfall = np.where(bal < 0, -bal, 0.0).sum(axis=1)      # demand
    gross_surplus = np.where(bal > 0, bal, 0.0).sum(axis=1)         # supply
    net_balance = bal.sum(axis=1)

    rows = []
    plant_balances: Dict[str, np.ndarray] = {}
    sorted_net = np.sort(net_balance)
    for i, p in plants.iterrows():
        b = bal[:, i]
        plant_balances[p["Reg_Number"]] = b
        s = np.sort(b)
        mean = float(b.mean())
        p10 = float(s[int(0.10 * n_sim)])
        med = float(s[int(0.50 * n_sim)])
        p90 = float(s[int(0.90 * n_sim)])
        sf_pct = float((b < 0).mean() * 100.0)
        exp_purch = float(np.where(b < 0, -b, 0.0).mean())
        exp_cred = float(np.where(b > 0, b, 0.0).mean())
        rows.append({
            "Reg_Number": p["Reg_Number"],
            "Company": p["Company"],
            "Plant": p["Plant"],
            "Sector": p.get("Sector", ""),
            "Mean_Net_Position_CCC": mean,
            "P10_Net_Position": p10,
            "Median_Net_Position": med,
            "P90_Net_Position": p90,
            "Chance_of_Shortfall": sf_pct / 100.0,
            "Expected_Purchases_Demand": exp_purch,
            "Expected_Credits_Surplus": exp_cred,
            "Baseline_GEI": p["Baseline_GEI"],
            "Target_GEI": p["Target_GEI"],
            "Actual_Production": p["Actual_Production"],
            "Min_GEI": p["Min_GEI"],
            "Mode_GEI": p["Mode_GEI"],
            "Max_GEI": p["Max_GEI"],
        })
    res_df = pd.DataFrame(rows)

    summary = {
        "n_plants": int(n),
        "n_sim": int(n_sim),
        "total_mean_net": float(net_balance.mean()),
        "total_p10": float(sorted_net[int(0.10 * n_sim)]),
        "total_median": float(sorted_net[int(0.50 * n_sim)]),
        "total_p90": float(sorted_net[int(0.90 * n_sim)]),
        "total_gross_demand": float(gross_shortfall.mean()),
        "total_gross_supply": float(gross_surplus.mean()),
    }
    return res_df, summary, plant_balances
