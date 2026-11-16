"""5,000-session market-clearing price discovery with pluggable engines & sealed corporate auction bids."""

from __future__ import annotations

from typing import Dict, List, Optional, Tuple
import numpy as np
import pandas as pd

N_PRICE_SESSIONS = 5_000
SEED = 42
PRICE_STEP = 25

ENGINE_RANDOM = "Random Tick Clearing (Current)"
ENGINE_MACC = "MACC-Anchored Rational Bidding"


def _draw_thresholds(
    n_buyers: int,
    n_sellers: int,
    price_ticks: np.ndarray,
    rng: np.random.Generator,
    macc_inr_per_tco2: Optional[np.ndarray] = None,
    engine: str = ENGINE_RANDOM,
    buyer_bid_overrides: Optional[Dict[int, float]] = None,
    seller_offer_overrides: Optional[Dict[int, float]] = None,
) -> Tuple[np.ndarray, np.ndarray]:
    if engine == ENGINE_MACC and macc_inr_per_tco2 is not None and n_buyers > 0:
        base = rng.choice(price_ticks, size=n_buyers)
        b_thresh = np.minimum(base, macc_inr_per_tco2)
    else:
        b_thresh = rng.choice(price_ticks, size=n_buyers)

    s_thresh = rng.choice(price_ticks, size=n_sellers)

    if buyer_bid_overrides:
        for idx, bid_px in buyer_bid_overrides.items():
            if 0 <= idx < n_buyers:
                b_thresh[idx] = float(bid_px)
    if seller_offer_overrides:
        for idx, ask_px in seller_offer_overrides.items():
            if 0 <= idx < n_sellers:
                s_thresh[idx] = float(ask_px)
    return b_thresh, s_thresh


def _clearing_session(
    demand_vols: np.ndarray,
    surplus_vols: np.ndarray,
    price_ticks: np.ndarray,
    rng: np.random.Generator,
    macc_inr_per_tco2: Optional[np.ndarray] = None,
    engine: str = ENGINE_RANDOM,
    buyer_bid_overrides: Optional[Dict[int, float]] = None,
    seller_offer_overrides: Optional[Dict[int, float]] = None,
) -> Tuple[Optional[float], Optional[float], dict]:
    n_buyers = len(demand_vols)
    n_sellers = len(surplus_vols)

    b_thresh, s_thresh = _draw_thresholds(
        n_buyers, n_sellers, price_ticks, rng,
        macc_inr_per_tco2=macc_inr_per_tco2,
        engine=engine,
        buyer_bid_overrides=buyer_bid_overrides,
        seller_offer_overrides=seller_offer_overrides,
    )

    dem_curve = np.array([demand_vols[b_thresh >= p].sum() for p in price_ticks])
    sup_curve = np.array([surplus_vols[s_thresh <= p].sum() for p in price_ticks])

    trade_vol = np.minimum(dem_curve, sup_curve)
    mismatch = np.abs(dem_curve - sup_curve)

    if not np.any(trade_vol > 0):
        return None, None, {}

    best_idx = int(np.argmin(mismatch))
    mcp = float(price_ticks[best_idx])
    vol = float(trade_vol[best_idx])
    mkt_supply = float(sup_curve[best_idx])
    mkt_demand = float(dem_curve[best_idx])

    meta = {
        "best_idx": best_idx,
        "mcp": mcp,
        "volume": vol,
        "market_demand": mkt_demand,
        "market_supply": mkt_supply,
        "mismatch": float(mismatch[best_idx]),
        "b_thresh": b_thresh,
        "s_thresh": s_thresh,
    }
    return mcp, vol, meta


def run_price_discovery(
    demand_vols: np.ndarray,
    surplus_vols: np.ndarray,
    price_floor: float = 600,
    price_cap: float = 1600,
    price_step: int = PRICE_STEP,
    n_sessions: int = N_PRICE_SESSIONS,
    seed: int = SEED,
    engine: str = ENGINE_RANDOM,
    macc_inr_per_tco2: Optional[np.ndarray] = None,
    buyer_bid_overrides: Optional[Dict[int, float]] = None,
    seller_offer_overrides: Optional[Dict[int, float]] = None,
) -> Tuple[dict, pd.DataFrame, pd.DataFrame, pd.DataFrame, List[float], List[float]]:
    rng = np.random.default_rng(seed)
    price_ticks = np.arange(price_floor, price_cap + price_step, price_step, dtype=float)

    demand_vols = np.asarray(demand_vols, dtype=float)
    surplus_vols = np.asarray(surplus_vols, dtype=float)

    n_buyers = len(demand_vols)
    n_sellers = len(surplus_vols)
    total_base_supply = float(surplus_vols.sum())

    sim_mcps: List[float] = []
    sim_vols: List[float] = []
    sim_records: List[dict] = []
    tick_counts = np.zeros_like(price_ticks, dtype=int)
    tick_vols = np.zeros_like(price_ticks, dtype=float)

    buy_win_sessions = 0
    sell_win_sessions = 0
    buy_win_mcps: List[float] = []
    sell_win_mcps: List[float] = []

    for s in range(1, n_sessions + 1):
        mcp, vol, meta = _clearing_session(
            demand_vols, surplus_vols, price_ticks, rng,
            macc_inr_per_tco2=macc_inr_per_tco2, engine=engine,
            buyer_bid_overrides=buyer_bid_overrides,
            seller_offer_overrides=seller_offer_overrides,
        )
        if mcp is None:
            continue

        pre_market_banked = total_base_supply - meta["market_supply"]
        unmet_seller = max(0.0, meta["market_supply"] - meta["market_demand"])

        sim_mcps.append(mcp)
        sim_vols.append(vol)
        tick_counts[meta["best_idx"]] += 1
        tick_vols[meta["best_idx"]] += vol

        if buyer_bid_overrides:
            if any(bid >= mcp for bid in buyer_bid_overrides.values()):
                buy_win_sessions += 1
                buy_win_mcps.append(mcp)
        if seller_offer_overrides:
            if any(ask <= mcp for ask in seller_offer_overrides.values()):
                sell_win_sessions += 1
                sell_win_mcps.append(mcp)

        sim_records.append({
            "Session_ID": s,
            "Clearing_Price_INR": mcp,
            "Volume_Cleared": round(vol, 2),
            "Market_Demand_at_MCP": round(meta["market_demand"], 2),
            "Market_Supply_at_MCP": round(meta["market_supply"], 2),
            "Pre_Market_Banked_Supply": round(pre_market_banked, 2),
            "Unmet_Seller_Credits": round(unmet_seller, 2),
            "Unmet_Volume_Mismatch": round(meta["mismatch"], 2),
        })

    total_sessions = len(sim_mcps)
    avg_demand_curve = np.zeros_like(price_ticks)
    avg_supply_curve = np.zeros_like(price_ticks)
    for _ in range(max(total_sessions, 1)):
        b_thresh, s_thresh = _draw_thresholds(
            n_buyers, n_sellers, price_ticks, rng,
            macc_inr_per_tco2=macc_inr_per_tco2, engine=engine,
            buyer_bid_overrides=buyer_bid_overrides,
            seller_offer_overrides=seller_offer_overrides,
        )
        avg_demand_curve += np.array([demand_vols[b_thresh >= p].sum() for p in price_ticks])
        avg_supply_curve += np.array([surplus_vols[s_thresh <= p].sum() for p in price_ticks])
    avg_demand_curve /= max(total_sessions, 1)
    avg_supply_curve /= max(total_sessions, 1)

    probabilities = tick_counts / max(total_sessions, 1)
    avg_cleared_vol_per_tick = np.where(
        tick_counts > 0,
        tick_vols / np.maximum(tick_counts, 1),
        0.0,
    )

    if sim_mcps:
        sim_mcps_arr = np.array(sim_mcps)
        expected_mcp = float(sim_mcps_arr.mean())
        nearest_idx = int(np.argmax(tick_counts))
        most_likely_mcp = float(price_ticks[nearest_idx])
        nearest_discrete_mcp = most_likely_mcp
        p50 = float(np.percentile(sim_mcps_arr, 50))
        p90 = float(np.percentile(sim_mcps_arr, 90))
        p10 = float(np.percentile(sim_mcps_arr, 10))
        p25 = float(np.percentile(sim_mcps_arr, 25))
        p75 = float(np.percentile(sim_mcps_arr, 75))
        prob_of_most_likely = float(tick_counts[nearest_idx] / total_sessions)
    else:
        expected_mcp = most_likely_mcp = nearest_discrete_mcp = p50 = p90 = p10 = p25 = p75 = 0.0
        prob_of_most_likely = 0.0

    summary = {
        "n_buyers": int(n_buyers),
        "n_sellers": int(n_sellers),
        "total_sessions": int(total_sessions),
        "price_floor": float(price_floor),
        "price_cap": float(price_cap),
        "price_ticks": int(len(price_ticks)),
        "engine": engine,
        "expected_mcp": expected_mcp,
        "most_likely_mcp": most_likely_mcp,
        "nearest_discrete_mcp": nearest_discrete_mcp,
        "prob_of_nearest_mcp": prob_of_most_likely,
        "p50": p50,
        "p90": p90,
        "p10": p10,
        "p25": p25,
        "p75": p75,
        "avg_cleared_volume": float(np.mean(sim_vols)) if sim_vols else 0.0,
        "total_base_supply": total_base_supply,
        "buy_fill_rate_pct": (buy_win_sessions / total_sessions * 100.0) if (total_sessions and buyer_bid_overrides) else None,
        "avg_buy_cleared_mcp": float(np.mean(buy_win_mcps)) if buy_win_mcps else None,
        "sell_fill_rate_pct": (sell_win_sessions / total_sessions * 100.0) if (total_sessions and seller_offer_overrides) else None,
        "avg_sell_cleared_mcp": float(np.mean(sell_win_mcps)) if sell_win_mcps else None,
    }

    price_tick_df = pd.DataFrame({
        "Price_Tick_INR": price_ticks,
        "Occurrences_Count": tick_counts,
        "Assigned_Probability": np.round(probabilities, 6),
        "Assigned_Probability_Pct": [f"{p * 100:.2f}%" for p in probabilities],
        "Cumulative_Probability_Pct": [
            f"{np.sum(probabilities[:i + 1]) * 100:.2f}%"
            for i in range(len(probabilities))
        ],
        "Avg_Cleared_Volume": np.round(avg_cleared_vol_per_tick, 2),
        "Weighted_Price_Contribution_INR": np.round(price_ticks * probabilities, 4),
    })

    session_df = pd.DataFrame(sim_records)
    curves_df = pd.DataFrame({
        "Price_Tick_INR": price_ticks,
        "Avg_Demand": avg_demand_curve,
        "Avg_Supply": avg_supply_curve,
    })

    return summary, price_tick_df, session_df, curves_df, sim_mcps, sim_vols
