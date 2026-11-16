"""
CCTS — Carbon Compliance Credit (CCC) Monte Carlo & Market Price Discovery
Refactored into engine/ modules. See engine/ for calculation logic.
"""

from __future__ import annotations

import io
import json
import os
from typing import Optional

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots
import streamlit as st
try:
    import pdfplumber
except Exception:
    pdfplumber = None

from engine import (
    load_main_sheet_from_excel,
    rebuild_main_sheet_from_sources,
    apply_sidebar_adjustments,
    clean_num,
    load_macc,
    get_sector_emissions_mtco2,
    add_sector_column,
    isolate_for_company,
    mask_competitors,
    run_ccc_monte_carlo,
    run_price_discovery,
    ENGINE_RANDOM,
    ENGINE_MACC,
    compute_plant_macc_table,
    abate_vs_buy_decision,
    compute_cbam_liability,
)
from engine.macc_cbam import get_marginal_abatement_cap
from engine.reports import build_market_analysis_sheet, build_company_wise_sheet, build_sector_summary_table

APP_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_DIR = os.path.join(APP_DIR, "config")
EXCEL_PATH = os.path.join(APP_DIR, "3. Data sheet.xlsx")

N_CCC_SIM = 10_000
N_PRICE_SESSIONS = 5_000
SEED = 42
DEFAULT_PRICE_FLOOR = 600
DEFAULT_PRICE_CAP = 1600
PRICE_STEP = 25
COMPLIANCE_YEARS = [2026, 2027, 2028, 2029]

MAIN_SHEET_COLS = [
    "Company", "S_No", "Plant", "Reg_Number", "Baseline_GEI", "Target_GEI",
    "Base_Production", "Actual_Production", "Min_GEI", "Mode_GEI", "Max_GEI",
]

st.set_page_config(
    page_title="CCTS — CCC Monte Carlo & Price Discovery",
    page_icon="🏭",
    layout="wide",
)

st.markdown(
    """
    <style>
    .auth-panel {border:1px solid #e0e0e0; border-radius:8px; padding:12px 16px;
                 margin-bottom:16px; background:#fafafa;}
    .auth-title {font-size:1.1rem; font-weight:700; margin-bottom:8px;}
    .isolated-badge {background:#e7f3fe; border-left:4px solid #2196f3;
                     padding:6px 10px; border-radius:4px; font-size:0.85rem; margin-bottom:10px;}
    .admin-badge {background:#e8f5e9; border-left:4px solid #4caf50;
                  padding:6px 10px; border-radius:4px; font-size:0.85rem; margin-bottom:10px;}
    </style>
    """,
    unsafe_allow_html=True,
)

def _ss(key, default=None):
    if key not in st.session_state:
        st.session_state[key] = default
    return st.session_state[key]

_ss("auth_mode", None)
_ss("auth_company", None)
_ss("df_plants", None)
_ss("load_source", None)
_ss("res_df", None)
_ss("summary", None)
_ss("price_summary", None)
_ss("price_tick_df", None)
_ss("session_df", None)
_ss("curves_df", None)
_ss("sim_mcps", None)
_ss("sim_vols", None)
_ss("overrides", {})
_ss("new_facilities", [])
_ss("year_rules", None)
_ss("price_engine", ENGINE_RANDOM)
_ss("macc_abate_overrides", {})
_ss("auction_bid_overrides", {})


# ---------------------------------------------------------------------------
# Google GenAI (Auto-Update tab)
# ---------------------------------------------------------------------------
try:
    from google import genai as _genai_mod
    HAS_GENAI = True
except Exception:
    HAS_GENAI = False
    _genai_mod = None

GEMINI_PROMPT = """
You are a data-extraction assistant for India's Carbon Compliance Credit Scheme (CCTS).
Extract plant/corporate disclosure data from the uploaded BEE target document or BRSR report.

Return ONLY a JSON array of objects, each with these exact fields:
- company (string)
- plant (string)
- reg_number (string)
- baseline_gei (number, GHG intensity 2023-24)
- target_gei (number, target intensity 2025-26)
- base_production (number, 2024 equivalent output)
- actual_production (number, 2026 equivalent product)
- min_gei (number, best-case intensity)
- mode_gei (number, most-likely intensity / CAGR)
- max_gei (number, worst-case intensity)

Rules:
- Use only numbers present in the document; do not invent values.
- If a field is unavailable, set it to null.
- One object per plant/facility.
- Output valid JSON only, no markdown fences, no explanation.
"""


def gemini_parse_pdf(pdf_file, api_key):
    if not HAS_GENAI:
        raise RuntimeError("google-genai is not installed. Add it to requirements.txt.")
    if not api_key:
        raise RuntimeError("GEMINI_API_KEY is not set.")
    from google import genai as genai_mod
    client = genai_mod.Client(api_key=api_key)

    pdf_file.seek(0)
    with pdfplumber.open(pdf_file) as pdf:
        text = "\n".join((p.extract_text() or "") for p in pdf.pages)
    if not text.strip():
        raise ValueError("No text could be extracted from the PDF.")

    response = client.models.generate_content(
        model="gemini-2.0-flash",
        contents=GEMINI_PROMPT + "\n\n--- DOCUMENT TEXT ---\n" + text[:12000],
    )
    raw = response.text or "[]"
    raw = raw.strip()
    if raw.startswith("```"):
        raw = raw.strip("`")
        if "\n" in raw:
            raw = "\n".join(raw.split("\n")[1:])
    import json
    try:
        data = json.loads(raw)
    except json.JSONDecodeError:
        start = raw.find("[")
        end = raw.rfind("]")
        if start != -1 and end != -1 and end > start:
            data = json.loads(raw[start:end + 1])
        else:
            raise
    records = []
    for item in data:
        records.append({
            "Company": item.get("company"),
            "Plant": item.get("plant"),
            "Reg_Number": item.get("reg_number"),
            "Baseline_GEI": clean_num(item.get("baseline_gei")),
            "Target_GEI": clean_num(item.get("target_gei")),
            "Base_Production": clean_num(item.get("base_production")) or 0.0,
            "Actual_Production": clean_num(item.get("actual_production")),
            "Min_GEI": clean_num(item.get("min_gei")),
            "Mode_GEI": clean_num(item.get("mode_gei")),
            "Max_GEI": clean_num(item.get("max_gei")),
        })
    records = [r for r in records if r["Reg_Number"] and r["Baseline_GEI"] is not None
               and r["Target_GEI"] is not None and r["Actual_Production"] is not None
               and r["Min_GEI"] is not None and r["Mode_GEI"] is not None
               and r["Max_GEI"] is not None]
    return records


def load_year_rules() -> dict:
    path = os.path.join(CONFIG_DIR, "year_rules.json")
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def load_company_credentials() -> dict:
    path = os.path.join(CONFIG_DIR, "company_credentials.json")
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


# ---------------------------------------------------------------------------
# Auth panel (Feature 1)
# ---------------------------------------------------------------------------
def render_auth_panel():
    creds = load_company_credentials()
    companies = sorted(creds["companies"].keys())

    st.markdown('<div class="auth-panel">', unsafe_allow_html=True)
    st.markdown('<div class="auth-title">🔐 Corporate Login & Data Isolation</div>', unsafe_allow_html=True)

    login_mode = st.radio(
        "Login mode",
        ["Market Admin (Climactio)", "Corporate Client Login"],
        index=0 if _ss("auth_mode") in (None, "admin") else 1,
        horizontal=True,
        key="login_mode_radio",
    )

    if login_mode == "Market Admin (Climactio)":
        admin_user = st.text_input("Admin username", value="Climactio", key="admin_user")
        admin_pass = st.text_input("Admin passcode", type="password", key="admin_pass")
        if st.button("Login as Admin", key="btn_admin_login"):
            if (admin_user == creds["market_admin"]["username"]
                    and admin_pass == creds["market_admin"]["passcode"]):
                _ss("auth_mode", "admin")
                _ss("auth_company", None)
                st.success("Logged in as Market Admin — full access.")
            else:
                st.error("Invalid admin credentials.")
        st.markdown('<div class="admin-badge">🟢 Market Admin — full access to all facilities & macro settings</div>', unsafe_allow_html=True)
    else:
        company = st.selectbox("Select your company", companies, key="corp_company")
        comp_info = creds["companies"].get(company, {})
        st.caption(f"Sector: {comp_info.get('sector', '—')}")

        col_pass, col_auto = st.columns([3, 1])
        with col_pass:
            passcode = st.text_input("Company passcode", type="password", key="corp_pass")
        with col_auto:
            if st.button("Demo auto-fill", key="btn_autofill"):
                st.session_state["corp_pass"] = comp_info.get("passcode", "")
                st.info(f"Passcode auto-filled for {company}.")
        if st.button("Login as Company", key="btn_corp_login"):
            if passcode == comp_info.get("passcode"):
                _ss("auth_mode", "corporate")
                _ss("auth_company", company)
                st.success(f"Logged in as {company} — isolated to your installations only.")
            else:
                st.error("Invalid company passcode.")

        if _ss("auth_mode") == "corporate" and _ss("auth_company"):
            st.markdown(
                f'<div class="isolated-badge">🔒 Corporate Client: <b>{_ss("auth_company")}</b> — '
                f"isolated to your own installations. Competitors masked in background pool.</div>",
                unsafe_allow_html=True,
            )

    st.markdown("</div>", unsafe_allow_html=True)
    return _ss("auth_mode"), _ss("auth_company")


# ---------------------------------------------------------------------------
# Data loading (respects auth isolation)
# ---------------------------------------------------------------------------
@st.cache_data(show_spinner=False)
def _load_existing_main(excel_path):
    return load_main_sheet_from_excel(excel_path)[0]


@st.cache_data(show_spinner=False)
def _load_rebuilt_main(excel_path, year):
    return rebuild_main_sheet_from_sources(excel_path, compliance_year=year)[0]


def load_facility_data(auth_mode, auth_company, data_mode, uploaded_csv, compliance_year):
    """Load facility data. For corporate clients, isolate to their company."""
    df_plants = None
    load_source = None

    if data_mode == "Existing Excel 'Main Sheet'":
        if os.path.exists(EXCEL_PATH):
            df_plants = _load_existing_main(EXCEL_PATH)
            load_source = "Main Sheet"
    elif data_mode == "Rebuild from source tabs":
        if os.path.exists(EXCEL_PATH):
            with st.spinner("Rebuilding Main Sheet from source tabs…"):
                df_plants = _load_rebuilt_main(EXCEL_PATH, compliance_year)
            load_source = "Rebuilt from sources"
    else:
        if uploaded_csv is not None:
            try:
                df_up = pd.read_csv(uploaded_csv)
                col_map = {}
                for c in df_up.columns:
                    cs = str(c).strip().lower()
                    if "reg" in cs:
                        col_map[c] = "Reg_Number"
                    elif "company" in cs:
                        col_map[c] = "Company"
                    elif "plant" in cs and "reg" not in cs:
                        col_map[c] = "Plant"
                    elif "baseline" in cs:
                        col_map[c] = "Baseline_GEI"
                    elif "target" in cs:
                        col_map[c] = "Target_GEI"
                    elif "base" in cs and "prod" in cs:
                        col_map[c] = "Base_Production"
                    elif "actual" in cs or "2026" in cs:
                        col_map[c] = "Actual_Production"
                    elif "min" in cs:
                        col_map[c] = "Min_GEI"
                    elif "mode" in cs or "most" in cs or "median" in cs:
                        col_map[c] = "Mode_GEI"
                    elif "max" in cs:
                        col_map[c] = "Max_GEI"
                df_up = df_up.rename(columns=col_map)
                for col in MAIN_SHEET_COLS:
                    if col not in df_up.columns:
                        df_up[col] = np.nan
                df_plants = df_up[MAIN_SHEET_COLS]
                load_source = "Uploaded CSV"
            except Exception as e:
                st.error(f"Failed to read CSV: {e}")

    if df_plants is None or df_plants.empty:
        return None, None, None, None

    # Add sector column
    df_plants = add_sector_column(df_plants)

    # Apply overrides from session state (Feature 3)
    overrides = _ss("overrides")
    for reg, ov in overrides.items():
        mask = df_plants["Reg_Number"] == reg
        for k, v in ov.items():
            if k in df_plants.columns:
                df_plants.loc[mask, k] = v
    # Append new facilities
    for nf in _ss("new_facilities"):
        df_plants = pd.concat([df_plants, pd.DataFrame([nf])], ignore_index=True)

    # Auth isolation
    own_df = None
    others_df = None
    masked_others = None
    if auth_mode == "corporate" and auth_company:
        own_df, others_df = isolate_for_company(df_plants, auth_company)
        masked_others = mask_competitors(others_df)
        # For simulation, use full df (own editable + others original)
        sim_df = df_plants
    else:
        sim_df = df_plants

    sim_df = apply_sidebar_adjustments(sim_df, compliance_year,
                                       _ss("banking_pct", 0), _ss("cuf_shock", 0))
    return sim_df, own_df, masked_others, load_source


# ===========================================================================
# RENDER
# ===========================================================================
st.title("🏭 CCTS — CCC Monte Carlo & Market Price Discovery")
st.caption("End-to-end: Excel data replication → 10,000-iteration CCC Monte Carlo → "
           "5,000-session Price Discovery → MACC Abate-vs-Buy → EU/UK CBAM → Excel downloads.")

auth_mode, auth_company = render_auth_panel()

# ---------------------------------------------------------------------------
# Sidebar
# ---------------------------------------------------------------------------
with st.sidebar:
    st.header("⚙️ Model Controls")
    compliance_year = st.selectbox(
        "Compliance Year", COMPLIANCE_YEARS, index=0,
        help="Target year for compliance. Targets are escalated using each company's decarbonisation CAGR.")
    banking_pct = st.slider(
        "Banking Carry-Forward %", 0, 100, 0, step=5,
        help="Fraction of surplus credits held back (reduces market supply) and "
        "deficit covered by banked credits (reduces market demand).")
    cuf_shock = st.slider(
        "Sector CUF Shock (%)", -20, 20, 0, step=1,
        help="Scaling applied to 2026 equivalent production (capacity utilisation shock).")
    _ss("banking_pct", banking_pct)
    _ss("cuf_shock", cuf_shock)

    st.divider()
    st.subheader("Price Grid")
    price_floor = st.number_input("Floor (₹)", min_value=100, max_value=5000,
                                  value=DEFAULT_PRICE_FLOOR, step=25)
    price_cap = st.number_input("Cap (₹)", min_value=100, max_value=5000,
                                value=DEFAULT_PRICE_CAP, step=25)
    if price_cap <= price_floor:
        st.error("Price cap must be above floor.")
    _ss("price_floor", price_floor)
    _ss("price_cap", price_cap)

    st.divider()
    st.subheader("Price Discovery Engine (Feature 5)")
    price_engine = st.selectbox(
        "Clearing engine",
        [ENGINE_RANDOM, ENGINE_MACC],
        index=0 if _ss("price_engine") == ENGINE_RANDOM else 1,
        key="price_engine_select",
        help="MACC-Anchored Rational Bidding caps deficit plants' willingness-to-pay at their marginal abatement cost.")
    _ss("price_engine", price_engine)

    st.divider()
    st.subheader("Data Source")
    data_mode = st.radio(
        "Load facility data from",
        ["Existing Excel 'Main Sheet'", "Rebuild from source tabs", "Upload CSV"],
        index=0,
        help="Existing Main Sheet = fastest. Rebuild replicates Excel calcs from source tabs.")
    uploaded_csv = None
    if data_mode == "Upload CSV":
        uploaded_csv = st.file_uploader("Upload facility CSV", type=["csv"])
    run_button = st.button("🚀 Run Full Model", type="primary", use_container_width=True)

# ---------------------------------------------------------------------------
# Load data
# ---------------------------------------------------------------------------
sim_df, own_df, masked_others, load_source = load_facility_data(
    auth_mode, auth_company, data_mode, uploaded_csv, compliance_year)

# ---------------------------------------------------------------------------
# Run model
# ---------------------------------------------------------------------------
res_df = None
summary = None
price_summary = None
price_tick_df = None
session_df = None
curves_df = None
sim_mcps = None
sim_vols = None

if run_button:
    base_df = sim_df if sim_df is not None else None
    if base_df is None or base_df.empty:
        st.error("No facility data loaded. Please load data first.")
    else:
        st.session_state["df_plants"] = base_df
        st.session_state["price_floor"] = price_floor
        st.session_state["price_cap"] = price_cap
        st.session_state["compliance_year"] = compliance_year
        st.session_state["banking_pct"] = banking_pct
        st.session_state["cuf_shock"] = cuf_shock

        with st.spinner(f"Running {N_CCC_SIM:,}-iteration CCC Monte Carlo…"):
            res_df, summary, plant_balances = run_ccc_monte_carlo(base_df)
        st.session_state["res_df"] = res_df
        st.session_state["summary"] = summary

        demand_vols = res_df["Expected_Purchases_Demand"].to_numpy(dtype=float).copy()
        surplus_vols = res_df["Expected_Credits_Surplus"].to_numpy(dtype=float).copy()

        macc_abate_ov = _ss("macc_abate_overrides", {})
        auction_bid_ov = _ss("auction_bid_overrides", {})
        buyer_bid_overrides = {}
        seller_offer_overrides = {}

        for idx_p, r in res_df.iterrows():
            reg_id = r["Reg_Number"]
            abated_vol = float(macc_abate_ov.get(reg_id, 0.0) or 0.0)
            if abated_vol > 0:
                if demand_vols[idx_p] > 0:
                    rem_def = demand_vols[idx_p] - abated_vol
                    if rem_def >= 0:
                        demand_vols[idx_p] = rem_def
                    else:
                        demand_vols[idx_p] = 0.0
                        surplus_vols[idx_p] += abs(rem_def)
                else:
                    surplus_vols[idx_p] += abated_vol

            if reg_id in auction_bid_ov:
                ov_entry = auction_bid_ov[reg_id]
                if ov_entry.get("side") == "buy" and demand_vols[idx_p] > 0:
                    buyer_bid_overrides[int(idx_p)] = float(ov_entry["price"])
                elif ov_entry.get("side") == "sell" and surplus_vols[idx_p] > 0:
                    seller_offer_overrides[int(idx_p)] = float(ov_entry["price"])

        macc_caps = None
        if price_engine == ENGINE_MACC:
            macc_data = load_macc()
            sector_emissions = get_sector_emissions_mtco2()
            macc_caps = []
            for idx_p, r in res_df.iterrows():
                sector = r.get("Sector", "Cement")
                base_emissions = (r["Actual_Production"] or 0) * (r["Baseline_GEI"] or 0)
                macc_table = compute_plant_macc_table(
                    sector, base_emissions, macc_data, sector_emissions)
                cap_val = get_marginal_abatement_cap(
                    macc_table, float(demand_vols[idx_p]),
                    price_floor=float(price_floor), price_cap=float(price_cap),
                )
                macc_caps.append(cap_val)
            macc_caps = np.array(macc_caps, dtype=float)

        with st.spinner(f"Running {N_PRICE_SESSIONS:,}-session price discovery ({price_engine})…"):
            (price_summary, price_tick_df, session_df, curves_df,
             sim_mcps, sim_vols) = run_price_discovery(
                demand_vols, surplus_vols,
                price_floor=price_floor, price_cap=price_cap,
                n_sessions=N_PRICE_SESSIONS, seed=SEED,
                engine=price_engine,
                macc_inr_per_tco2=macc_caps,
                buyer_bid_overrides=buyer_bid_overrides or None,
                seller_offer_overrides=seller_offer_overrides or None)
        st.session_state["price_summary"] = price_summary
        st.session_state["price_tick_df"] = price_tick_df
        st.session_state["session_df"] = session_df
        st.session_state["curves_df"] = curves_df
        st.session_state["sim_mcps"] = sim_mcps
        st.session_state["sim_vols"] = sim_vols

        st.success(f"Model complete: {len(base_df)} plants | "
                   f"{N_CCC_SIM:,} CCC iterations | {N_PRICE_SESSIONS:,} price sessions "
                   f"[{price_engine}].")
else:
    res_df = st.session_state.get("res_df")
    summary = st.session_state.get("summary")
    price_summary = st.session_state.get("price_summary")
    price_tick_df = st.session_state.get("price_tick_df")
    session_df = st.session_state.get("session_df")
    curves_df = st.session_state.get("curves_df")
    sim_mcps = st.session_state.get("sim_mcps")
    sim_vols = st.session_state.get("sim_vols")


# ===========================================================================
# TABS
sel_reg = None
# ===========================================================================
tab_data, tab_ccc, tab_price, tab_install, tab_cbam, tab_rules, tab_downloads, tab_autoupdate, tab_market, tab_company = (
    st.tabs([
        "📊 Facility Data", "⚡ CCC Monte Carlo", "💰 Price Discovery",
        "🏭 Installation Decision Studio", "🌍 EU/UK CBAM Calculator",
        "⚙️ Multi-Year Rules", "💾 Downloads", "🤖 Auto-Update (2027–2029)",
        "📈 Market & Sector Analysis", "📊 Company-Wise & Top 5 Rankings"
    ]))

# -------------------- Tab 1: Facility Data --------------------
with tab_data:
    st.subheader("Facility Input Table")
    if sim_df is not None and not sim_df.empty:
        show_df = sim_df
        if auth_mode == "corporate" and own_df is not None:
            st.markdown(
                f'<div class="isolated-badge">🔒 Showing only <b>{auth_company}</b> installations '
                f"({len(own_df)} of {len(sim_df)} total). Competitors masked in background pool.</div>",
                unsafe_allow_html=True,
            )
            show_df = own_df
        st.write(f"**Source:** {load_source}  |  **Plants shown:** {len(show_df)}  |  "
                 f"**Compliance Year:** {compliance_year}  |  **Auth:** {auth_mode or '—'}")
        st.dataframe(
            show_df[MAIN_SHEET_COLS + (["Sector"] if "Sector" in show_df.columns else [])]
            .style.format({
                "Baseline_GEI": "{:.4f}", "Target_GEI": "{:.4f}",
                "Base_Production": "{:,.0f}", "Actual_Production": "{:,.0f}",
                "Min_GEI": "{:.4f}", "Mode_GEI": "{:.4f}", "Max_GEI": "{:.4f}",
            }),
            use_container_width=True, hide_index=True,
        )
        csv_bytes = show_df.to_csv(index=False).encode("utf-8")
        st.download_button("⬇️ Download Facility CSV", csv_bytes,
                           "facility_input.csv", "text/csv")
        if auth_mode == "corporate" and masked_others is not None and not masked_others.empty:
            with st.expander("Masked competitor pool (background simulation only)"):
                st.caption("Competitor plant-level data is masked. These installations "
                           "remain in the macro simulation pool but are not visible to you.")
                st.dataframe(masked_others[["Company", "Sector"]].drop_duplicates(),
                             use_container_width=True, hide_index=True)
    else:
        st.info("Load facility data using the sidebar controls, then click Run Full Model.")

# -------------------- Tab 2: CCC Monte Carlo --------------------
with tab_ccc:
    st.subheader("CCC Monte Carlo Results (10,000 iterations)")
    if res_df is not None and summary is not None:
        view_df = res_df
        if auth_mode == "corporate" and own_df is not None:
            own_regs = set(own_df["Reg_Number"])
            view_df = res_df[res_df["Reg_Number"].isin(own_regs)]
            st.markdown(
                f'<div class="isolated-badge">🔒 Showing CCC results for <b>{auth_company}</b> '
                f"only ({len(view_df)} installations).</div>", unsafe_allow_html=True)
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Plants shown", f"{len(view_df):,}")
        c2.metric("Gross Demand (avg)", f"{summary['total_gross_demand']:,.0f} CCC")
        c3.metric("Gross Supply (avg)", f"{summary['total_gross_supply']:,.0f} CCC")
        c4.metric("Net Balance (avg)", f"{summary['total_mean_net']:,.0f} CCC")
        st.dataframe(
            view_df.style.format({
                "Mean_Net_Position_CCC": "{:,.0f}",
                "P10_Net_Position": "{:,.0f}",
                "Median_Net_Position": "{:,.0f}",
                "P90_Net_Position": "{:,.0f}",
                "Chance_of_Shortfall": "{:.1%}",
                "Expected_Purchases_Demand": "{:,.0f}",
                "Expected_Credits_Surplus": "{:,.0f}",
                "Baseline_GEI": "{:.4f}", "Target_GEI": "{:.4f}",
                "Min_GEI": "{:.4f}", "Mode_GEI": "{:.4f}", "Max_GEI": "{:.4f}",
            }),
            use_container_width=True, hide_index=True,
        )
    else:
        st.info("Click Run Full Model to generate CCC results.")


# -------------------- Tab 3: Price Discovery --------------------
with tab_price:
    st.subheader("Market-Clearing Price Discovery (5,000 sessions)")
    if price_summary is not None:
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Expected MCP", f"₹{price_summary['expected_mcp']:.0f}")
        c2.metric("Most Likely MCP", f"₹{price_summary['most_likely_mcp']:.0f}")
        c3.metric("P50 Median", f"₹{price_summary['p50']:.0f}")
        c4.metric("P90 Risk Cap", f"₹{price_summary['p90']:.0f}")
        c5, c6 = st.columns(2)
        c5.metric("Buyers", f"{price_summary['n_buyers']:,}")
        c6.metric("Sellers", f"{price_summary['n_sellers']:,}")
        st.caption(f"Engine: {price_summary.get('engine', ENGINE_RANDOM)}")
        if price_summary.get("buy_fill_rate_pct") is not None or price_summary.get("sell_fill_rate_pct") is not None:
            st.markdown("#### 🎯 Corporate Sealed Auction Execution Scorecard (5,000 Sessions)")
            sc1, sc2, sc3, sc4 = st.columns(4)
            if price_summary.get("buy_fill_rate_pct") is not None:
                sc1.metric("Buy Bid Fill Rate", f"{price_summary['buy_fill_rate_pct']:.1f}%")
                sc2.metric("Avg MCP Paid When Cleared", f"₹{price_summary['avg_buy_cleared_mcp']:,.0f}" if price_summary.get("avg_buy_cleared_mcp") else "Unfilled")
            if price_summary.get("sell_fill_rate_pct") is not None:
                sc3.metric("Sell Offer Fill Rate", f"{price_summary['sell_fill_rate_pct']:.1f}%")
                sc4.metric("Avg MCP Earned When Cleared", f"₹{price_summary['avg_sell_cleared_mcp']:,.0f}" if price_summary.get("avg_sell_cleared_mcp") else "Unfilled")

        col_a, col_b = st.columns(2)
        # Merit order
        fig_mo = go.Figure()
        fig_mo.add_trace(go.Scatter(
            x=curves_df["Avg_Demand"], y=curves_df["Price_Tick_INR"],
            mode="lines", name="Aggregate Demand", line=dict(color="#1f77b4", width=3),
            hovertemplate="Price: ₹%{y:.0f}<br>Demand: %{x:,.0f}<extra></extra>"))
        fig_mo.add_trace(go.Scatter(
            x=curves_df["Avg_Supply"], y=curves_df["Price_Tick_INR"],
            mode="lines", name="Aggregate Supply", line=dict(color="#d62728", width=3),
            hovertemplate="Price: ₹%{y:.0f}<br>Supply: %{x:,.0f}<extra></extra>"))
        fig_mo.add_trace(go.Scatter(
            x=[price_summary["avg_cleared_volume"]], y=[price_summary["expected_mcp"]],
            mode="markers", name="Expected MCP",
            marker=dict(color="#2ca02c", size=14, symbol="diamond",
                        line=dict(width=2, color="black")),
            hovertemplate="MCP: ₹%{y:.0f}<br>Volume: %{x:,.0f}<extra></extra>"))
        fig_mo.update_layout(title="Merit-Order X-Curve", xaxis_title="Volume (CCCs)",
                             yaxis_title="Price (₹)", yaxis=dict(autorange="reversed"),
                             template="plotly_white", height=420,
                             margin=dict(l=10, r=10, t=40, b=10))
        col_a.plotly_chart(fig_mo, use_container_width=True)

        # MCP probability
        fig_mp = make_subplots(specs=[[{"secondary_y": True}]])
        fig_mp.add_trace(go.Bar(
            x=price_tick_df["Price_Tick_INR"],
            y=price_tick_df["Assigned_Probability"] * 100,
            name="Clearing Probability (%)", marker_color="#1f77b4",
            hovertemplate="Price: ₹%{x:.0f}<br>Probability: %{y:.2f}%<extra></extra>"),
            secondary_y=False)
        fig_mp.add_trace(go.Scatter(
            x=price_tick_df["Price_Tick_INR"],
            y=price_tick_df["Avg_Cleared_Volume"],
            mode="lines+markers", name="Avg Cleared Volume",
            line=dict(color="#ff7f0e", width=3),
            hovertemplate="Price: ₹%{x:.0f}<br>Volume: %{y:,.0f}<extra></extra>"),
            secondary_y=True)
        fig_mp.add_vline(x=price_summary["expected_mcp"], line=dict(color="green", width=2, dash="dash"),
                         annotation_text=f"Expected MCP ₹{price_summary['expected_mcp']:.0f}")
        fig_mp.add_vline(x=price_summary["p90"], line=dict(color="red", width=2, dash="dot"),
                         annotation_text=f"P90 Risk Cap ₹{price_summary['p90']:.0f}")
        fig_mp.update_layout(title="MCP Probability & Liquidity Depth",
                             legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
                             template="plotly_white", height=420, margin=dict(l=10, r=10, t=40, b=10))
        col_b.plotly_chart(fig_mp, use_container_width=True)

        # CDF
        fig_cdf = go.Figure()
        sorted_mcp = np.sort(sim_mcps) if sim_mcps else np.array([price_summary["expected_mcp"]])
        cdf = np.arange(1, len(sorted_mcp) + 1) / len(sorted_mcp)
        fig_cdf.add_trace(go.Scatter(
            x=sorted_mcp, y=cdf * 100, mode="lines", name="CDF",
            line=dict(color="#6a3d9a", width=3),
            hovertemplate="Price: ₹%{x:.0f}<br>Cumulative: %{y:.1f}%<extra></extra>"))
        fig_cdf.add_vline(x=price_summary["p50"], line=dict(color="green", width=2, dash="dash"),
                          annotation_text=f"P50 ₹{price_summary['p50']:.0f}")
        fig_cdf.add_vline(x=price_summary["p90"], line=dict(color="red", width=2, dash="dot"),
                          annotation_text=f"P90 ₹{price_summary['p90']:.0f}")
        fig_cdf.update_layout(title="Cumulative Risk CDF", xaxis_title="Clearing Price (₹)",
                              yaxis_title="Cumulative Probability (%)", template="plotly_white",
                              height=420, margin=dict(l=10, r=10, t=40, b=10))
        st.plotly_chart(fig_cdf, use_container_width=True)

        st.dataframe(price_tick_df, use_container_width=True, hide_index=True)
    else:
        st.info("Click Run Full Model to generate price-discovery results.")


# -------------------- Tab 4: Installation Decision Studio (Features 2,3) --------------------
with tab_install:
    st.subheader("Installation Decision Studio — MACC Step Curve & Sealed Auction Order")
    st.write("Step 1: Check MACC internal abatement levers (X = Cumulative tCO₂, Y = Cost ₹/tCO₂) → Step 2: Enter your sealed buy/sell price (₹/CCC) → Step 3: Click Run Full Model.")

    if auth_mode == "corporate" and auth_company:
        avail_companies = [auth_company]
    elif sim_df is not None and not sim_df.empty:
        avail_companies = sorted(sim_df["Company"].dropna().unique())
    else:
        avail_companies = sorted(load_company_credentials()["companies"].keys())

    sel_company = st.selectbox("Select company", avail_companies, key="sel_company_studio")
    if auth_mode == "corporate" and own_df is not None:
        studio_df = own_df.copy()
        studio_df["Plant"] = studio_df["Plant"].str.replace(r"\[Masked.*?\]", "", regex=True)
    elif sim_df is not None and not sim_df.empty:
        studio_df = sim_df[sim_df["Company"] == sel_company].copy()
    else:
        studio_df = pd.DataFrame(columns=MAIN_SHEET_COLS + ["Sector"])

    sector_emissions = get_sector_emissions_mtco2()
    macc_data = load_macc()

    if not studio_df.empty:
        sel_plant = st.selectbox("Select installation", sorted(studio_df["Plant"].unique()), key="sel_inst_studio")
        plant_rows = studio_df[studio_df["Plant"] == sel_plant]
    else:
        sel_plant, plant_rows = "", pd.DataFrame()

    if not plant_rows.empty and sel_plant:
        sel_reg = st.selectbox("Select facility registration", plant_rows["Reg_Number"].tolist(), key="sel_reg_studio")
        plant = plant_rows[plant_rows["Reg_Number"] == sel_reg].iloc[0]
        sector = plant.get("Sector", "Cement")
        plant_name = plant["Plant"]

        base_em = float((plant["Actual_Production"] or 0) * (plant["Baseline_GEI"] or 0))
        sec_mt = float(sector_emissions.get(sector, 0) or 0)
        scale_pct = (base_em / (sec_mt * 1_000_000.0) * 100.0) if sec_mt > 0 else 0.0

        plant_purch = plant_cred = 0.0
        if res_df is not None and sel_reg in res_df["Reg_Number"].values:
            rm = res_df[res_df["Reg_Number"] == sel_reg].iloc[0]
            plant_purch, plant_cred = float(rm["Expected_Purchases_Demand"]), float(rm["Expected_Credits_Surplus"])
        else:
            db = ((plant["Target_GEI"] or 0) - (plant["Mode_GEI"] or plant["Baseline_GEI"] or 0)) * (plant["Actual_Production"] or 0)
            if db < 0:
                plant_purch = abs(db)
            else:
                plant_cred = db

        k1, k2, k3, k4, k5 = st.columns(5)
        k1.metric("Sector (Reg Prefix)", sector)
        k2.metric("Baseline Emissions", f"{base_em:,.0f} tCO₂")
        k3.metric("Sector 2024 Total", f"{sec_mt:,.2f} MtCO₂")
        k4.metric("Plant Scale Share", f"{scale_pct:.3f}%")
        k5.metric("Pre-MACC Position", f"-{plant_purch:,.0f} Deficit" if plant_purch > 0 else f"+{plant_cred:,.0f} Surplus")

        st.divider()
        st.subheader("Step 1: Choose Internal MACC Abatement Levers")
        macc_table = compute_plant_macc_table(sector, base_em, macc_data, sector_emissions)
        mcp_display = float(price_summary["expected_mcp"]) if price_summary else float(DEFAULT_PRICE_CAP)

        if not macc_table.empty:
            ms = macc_table.sort_values("Cost_INR_tCO2").reset_index(drop=True).copy()
            ms["Adopt"] = ms["Cost_INR_tCO2"].fillna(999999) <= mcp_display
            dcols = ["Adopt", "Lever", "Category", "Annual_Abatement_MtCO2", "Plant_Abatement_Capacity_tCO2", "Cost_USD_tCO2", "Cost_INR_tCO2"]
            for dc in dcols:
                if dc not in ms.columns:
                    ms[dc] = 0.0

            edited_macc = st.data_editor(
                ms[dcols],
                column_config={
                    "Adopt": st.column_config.CheckboxColumn("Adopt Internally?", default=False),
                    "Lever": st.column_config.TextColumn("Abatement Lever", disabled=True),
                    "Category": st.column_config.TextColumn("Category", disabled=True),
                    "Annual_Abatement_MtCO2": st.column_config.NumberColumn("Sector Potential (MtCO₂)", format="%.3f", disabled=True),
                    "Plant_Abatement_Capacity_tCO2": st.column_config.NumberColumn("Plant Scaled Capacity (tCO₂)", format="%,.0f"),
                    "Cost_USD_tCO2": st.column_config.NumberColumn("Cost ($/tCO₂)", format="$%.1f", disabled=True),
                    "Cost_INR_tCO2": st.column_config.NumberColumn("Cost (₹/tCO₂)", format="₹%,.0f"),
                },
                use_container_width=True, hide_index=True, key=f"macc_ed_{sel_reg}",
            )

            cdf = edited_macc.sort_values("Cost_INR_tCO2").reset_index(drop=True).copy()
            widths = cdf["Plant_Abatement_Capacity_tCO2"].fillna(0).clip(lower=0).to_numpy(dtype=float)
            costs = cdf["Cost_INR_tCO2"].fillna(0).to_numpy(dtype=float)
            cum_r = np.cumsum(widths)
            cum_l = np.concatenate([[0.0], cum_r[:-1]])
            centers = (cum_l + cum_r) / 2.0
            colors = ["#2e7d32" if a else ("#1565c0" if c <= mcp_display else "#ef6c00") for a, c in zip(cdf["Adopt"], costs)]

            fig_macc = go.Figure()
            fig_macc.add_trace(go.Bar(
                x=centers, y=costs, width=widths, marker_color=colors, marker_line=dict(color="white", width=1),
                customdata=np.stack([cdf["Lever"], cdf["Category"], widths, cum_r, np.where(cdf["Adopt"], "Adopted", "Not Adopted")], axis=-1),
                hovertemplate="<b>%{customdata[0]}</b> (%{customdata[1]})<br>Cost (Y): ₹%{y:,.0f}/tCO₂<br>Capacity Width: %{customdata[2]:,.0f} tCO₂<br>Cumulative Abatement (X): %{customdata[3]:,.0f} tCO₂<br>Status: %{customdata[4]}<extra></extra>",
                name="MACC Levers",
            ))
            fig_macc.add_hline(y=mcp_display, line=dict(color="#d32f2f", width=2, dash="dash"), annotation_text=f"Expected MCP: ₹{mcp_display:,.0f}/CCC")
            if plant_purch > 0:
                fig_macc.add_vline(x=plant_purch, line=dict(color="#6a1b9a", width=2, dash="dot"), annotation_text=f"Deficit: {plant_purch:,.0f} tCO₂")
            fig_macc.update_layout(
                title=f"Marginal Abatement Cost Curve — {plant_name} ({sector})",
                xaxis_title="Cumulative Plant Abatement Potential (tCO₂)",
                yaxis_title="Marginal Abatement Cost (₹ / tCO₂)",
                template="plotly_white", height=440, bargap=0, margin=dict(l=10, r=10, t=45, b=10),
            )
            st.plotly_chart(fig_macc, use_container_width=True)

            adopted = edited_macc[edited_macc["Adopt"] == True]
            sel_abated = float(adopted["Plant_Abatement_Capacity_tCO2"].fillna(0).sum())
            abate_cost = float((adopted["Plant_Abatement_Capacity_tCO2"].fillna(0) * adopted["Cost_INR_tCO2"].fillna(0)).sum())
            net_def = max(0.0, plant_purch - sel_abated) if plant_purch > 0 else 0.0
            net_sur = max(0.0, sel_abated - plant_purch) if plant_purch > 0 else (plant_cred + sel_abated)
            marg_px = get_marginal_abatement_cap(edited_macc, max(net_def, 1.0), float(price_floor), float(price_cap))

            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Selected Internal Abatement", f"{sel_abated:,.0f} tCO₂")
            m2.metric("Internal Abatement Cost", f"₹{abate_cost:,.0f}")
            m3.metric("Post-MACC Auction Volume", f"Buy {net_def:,.0f} CCC" if net_def > 0 else f"Sell {net_sur:,.0f} CCC")
            m4.metric("Marginal MACC Cost", f"₹{marg_px:,.0f}/CCC")
        else:
            sel_abated, net_def, net_sur, marg_px = 0.0, plant_purch, plant_cred, float(mcp_display)

        st.divider()
        st.subheader("Step 2: Enter Sealed Buy / Sell Price & Lock Order")
        auc_side = st.radio("Order Side", ["Buy Credits (Deficit)", "Sell Credits (Surplus)"], index=0 if net_def > 0 else 1, horizontal=True, key=f"auc_side_{sel_reg}")
        def_px = float(min(max(round(marg_px / 25) * 25, price_floor), price_cap))
        sealed_px = st.number_input("Your Sealed Buy Bid Price (₹/CCC)" if "Buy" in auc_side else "Your Sealed Sell Offer Price (₹/CCC)", min_value=float(price_floor), max_value=float(price_cap), value=def_px, step=25.0, key=f"sealed_px_{sel_reg}")

        b1, b2 = st.columns(2)
        with b1:
            if st.button("🚀 Lock MACC & Sealed Price (Then Click Run Full Model)", type="primary", use_container_width=True, key=f"lock_auc_{sel_reg}"):
                ab_ov = dict(_ss("macc_abate_overrides", {}))
                ab_ov[sel_reg] = sel_abated
                st.session_state["macc_abate_overrides"] = ab_ov
                bid_ov = dict(_ss("auction_bid_overrides", {}))
                bid_ov[sel_reg] = {"side": "buy" if "Buy" in auc_side else "sell", "price": float(sealed_px), "abated_tco2": sel_abated}
                st.session_state["auction_bid_overrides"] = bid_ov
                st.success(f"Locked: Internal Abatement = {sel_abated:,.0f} tCO₂ | Sealed {'Buy' if 'Buy' in auc_side else 'Sell'} = ₹{sealed_px:,.0f}/CCC. Click '🚀 Run Full Model' in the sidebar!")
        with b2:
            if st.button("🧹 Clear Custom MACC & Auction Overrides", use_container_width=True, key=f"clear_auc_{sel_reg}"):
                st.session_state["macc_abate_overrides"] = {}
                st.session_state["auction_bid_overrides"] = {}
                st.info("Cleared custom overrides.")
    else:
        st.info("Select a company and installation to open the MACC Step Curve & Sealed Auction Builder.")

# -------------------- Tab 5: EU/UK CBAM Calculator (Feature 4) --------------------
with tab_cbam:
    st.subheader("EU / UK CBAM Export Liability Calculator")
    st.write("For each installation, set EU/UK export share, EU ETS price (€/tCO2), and UK CBAM "
             "quarterly rate (£/tCO2, effective 1 Jan 2027 for Aluminium, Cement, Fertilisers, "
             "Hydrogen, Iron & Steel). Combined domestic + cross-border carbon exposure.")

    yr_rules = _ss("year_rules") or load_year_rules()
    yr = COMPLIANCE_YEARS[_ss("compliance_year") - 2026] if _ss("compliance_year") in COMPLIANCE_YEARS else 2026
    yr_rules = yr_rules.get(str(yr), yr_rules)
    uk_sectors = yr_rules.get("uk_cbam_sectors", [])

    eu_ets_price = st.number_input("EU ETS carbon price (€/tCO2)", value=85.0, step=1.0,
                                   key="cbam_eu_ets")
    uk_cbam_rate = st.number_input("UK CBAM rate (£/tCO2)", value=45.0, step=0.5,
                                   key="cbam_uk_rate")
    c1, c2 = st.columns(2)
    with c1:
        eu_export = st.slider("EU Export %", 0, 100, 0, step=5, key="cbam_eu_export")
    with c2:
        uk_export = st.slider("UK Export %", 0, 100, 0, step=5, key="cbam_uk_export")

    st.subheader("CBAM Liability per Installation")
    if res_df is not None and not res_df.empty:
        view_df = res_df.copy()
        if auth_mode == "corporate" and own_df is not None:
            own_regs = set(own_df["Reg_Number"])
            view_df = res_df[res_df["Reg_Number"].isin(own_regs)]
            st.caption(f"Showing {auth_company} installations only.")

        cbam_rows = []
        for _, r in view_df.head(20).iterrows():
            mcp_display = price_summary["expected_mcp"] if price_summary else DEFAULT_PRICE_CAP
            cb = compute_cbam_liability(
                r, mcp_display, yr_rules, eu_export_pct=eu_export,
                uk_export_pct=uk_export, eu_ets_price_eur=eu_ets_price,
                uk_cbam_rate_gbp=uk_cbam_rate)
            cbam_rows.append({
                "Company": r["Company"], "Plant": r["Plant"],
                "Sector": r.get("Sector", ""),
                "Emissions (tCO2)": f"{cb['Emissions_tCO2']:,.0f}",
                "Domestic CCTS (₹)": f"{cb['Domestic_CCTS_Cost_INR']:,.0f}",
                "EU Export (%)": f"{cb['EU_Export_Pct']:.0f}%",
                "EU CBAM (₹)": f"{cb['EU_CBAM_Cost_INR']:,.0f}",
                "EU CBAM Net (₹)": f"{cb['EU_CBAM_Net_INR']:,.0f}",
                "UK Export (%)": f"{cb['UK_Export_Pct']:.0f}%",
                "UK Applicable": "Yes" if cb["UK_Applicable"] else "No",
                "UK CBAM Net (₹)": f"{cb['UK_CBAM_Net_INR']:,.0f}",
                "Total CBAM (₹)": f"{cb['Total_CBAM_Net_INR']:,.0f}",
                "Total Exposure (₹)": f"{cb['Total_Exposure_INR']:,.0f}",
            })
        st.dataframe(pd.DataFrame(cbam_rows), use_container_width=True, hide_index=True)

        if cbam_rows:
            tot_exp = sum(float(str(next((v for k, v in r.items() if "exposure" in k.lower() and "inr" in k.lower()), 0)).replace(",", "")) for r in cbam_rows)
            st.metric("Combined Carbon Exposure (₹)",
                      f"{tot_exp:,.0f}")

            if sel_reg and res_df is not None:
                with st.expander("Adjust EU/UK export share for this plant"):
                    eu_e = st.slider("EU Export %", 0, 100, eu_export, step=5, key="cbam_eu_export_sel")
                    uk_e = st.slider("UK Export %", 0, 100, uk_export, step=5, key="cbam_uk_export_sel")
                    if st.button("Apply CBAM export override", key="btn_apply_cbam_ovr"):
                        _ss("overrides", {sel_reg: {"EU_Export_Pct": eu_e, "UK_Export_Pct": uk_e}})
                        st.success("CBAM export override queued.")
    else:
        st.info("Run the model first to compute CBAM liabilities.")


# -------------------- Tab 6: Multi-Year Rules (Feature 5) --------------------
with tab_rules:
    st.subheader("Multi-Year Rules (2026–2029) & Pluggable Engine")
    st.write("Edit year-specific structural assumptions and switch the clearing engine.")

    yr_rules = load_year_rules()
    if _ss("compliance_year") in COMPLIANCE_YEARS:
        current_idx = COMPLIANCE_YEARS.index(_ss("compliance_year"))
        selected_yr = COMPLIANCE_YEARS[current_idx]
    else:
        selected_yr = COMPLIANCE_YEARS[0]
        _ss("compliance_year", selected_yr)

    c1, c2 = st.columns(2)
    with c1:
        st.subheader("Structural Assumptions by Year")
        for y in COMPLIANCE_YEARS:
            st.dataframe(pd.DataFrame([{"Parameter": k, "Value": str(v)} for k, v in yr_rules.get(str(y), yr_rules.get(y, {})).items()]), use_container_width=True, hide_index=True)
            st.divider()
        st.caption("Corporate office deduction is the 2% deduction applied in 2026 "
                   "(editable per year).")

    with c2:
        st.subheader("Clearing Engine (pluggable)")
        eng = st.radio(
            "Engine", [ENGINE_RANDOM, ENGINE_MACC],
            index=0 if _ss("price_engine") == ENGINE_RANDOM else 1,
            key="yr_engine_radio")
        _ss("price_engine", eng)
        st.caption("MACC-Anchored Rational Bidding caps deficit plants' WTP at marginal abatement cost.")

        st.subheader("Edit Year Rules")
        yr_rules_edit = dict(yr_rules[str(selected_yr)])
        for k in list(yr_rules_edit.keys()):
            if isinstance(yr_rules_edit[k], bool):
                del yr_rules_edit[k]
        for k, v in yr_rules_edit.items():
            if isinstance(v, (int, float)):
                yr_rules_edit[k] = st.number_input(f"{k}", value=v, key=f"yr_{selected_yr}_{k}")
        yr_rules[str(selected_yr)] = yr_rules_edit
        with open(os.path.join(CONFIG_DIR, "year_rules.json"), "w") as f:
            json.dump(yr_rules, f, indent=2)
        st.success("Year rules saved to config/year_rules.json.")


def df_to_excel_bytes(dfs):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        for sheet, d in dfs.items():
            d.to_excel(writer, sheet_name=sheet, index=False)
    buf.seek(0)
    return buf.read()


# -------------------- Tab 7: Downloads --------------------
with tab_downloads:
    st.subheader("1-Click Excel Downloads")
    if res_df is not None and price_summary is not None:
        sum_rows = {k: v for k, v in summary.items()
                    if not isinstance(v, np.ndarray) and not k.endswith("_series")}
        fac_bytes = df_to_excel_bytes({
            "Facility_Results": res_df,
            "Portfolio_Summary": pd.DataFrame([sum_rows]),
        })
        # Build the two new sheet DataFrames
        market_df = build_market_analysis_sheet(sim_df, res_df, summary) if sim_df is not None and res_df is not None and summary is not None else pd.DataFrame()
        company_df = build_company_wise_sheet(res_df, sim_df, locals().get("selected_sector", None)) if sim_df is not None and res_df is not None else pd.DataFrame()

        mkt_bytes = df_to_excel_bytes({
            "Dashboard_Summary": pd.DataFrame([price_summary]),
            "Price_Tick_Distribution": price_tick_df,
            "All_Sessions_Raw_Data": (session_df
                                      if session_df is not None
                                      else pd.DataFrame()),
            "Avg_Demand_Supply_Curves": (curves_df
                                         if curves_df is not None
                                         else pd.DataFrame()),
            "Market & Sector Analysis": market_df,
            "Company-Wise & Top 5 Rankings": company_df,
        })
        dl1, dl2 = st.columns(2)
        dl1.download_button(
            "⬇️ Facility Results (CCC) Excel", fac_bytes,
            "CCTS_Facility_CCC_Results.xlsx",
            ("application/vnd.openxmlformats-"
             "officedocument.spreadsheetml.sheet"),
            use_container_width=True,
        )
        dl2.download_button(
            "⬇️ Market Clearing Report Excel", mkt_bytes,
            "CCTS_Market_Clearing_Report.xlsx",
            ("application/vnd.openxmlformats-"
             "officedocument.spreadsheetml.sheet"),
            use_container_width=True,
        )
    else:
        st.info("Run the model first to enable downloads.")


# -------------------- Tab 8: Auto-Update (2027-2029) --------------------
with tab_autoupdate:
    st.subheader("Auto-Update Data (2027–2029)")
    st.write("Upload a BEE target PDF or BRSR report. Gemini parses it into the exact schema "
             "of 'CUF and BRSR data' / 'decarbonisation trend2% adjuste', appends new entities/plants, "
             "and re-runs the full model.")

    if not HAS_GENAI:
        st.warning("google-genai is not installed. Add `google-genai` to requirements.txt and restart.")
    api_key = st.text_input("GEMINI_API_KEY", type="password",
                            value=os.getenv("GEMINI_API_KEY", ""),
                            help="Set the GEMINI_API_KEY environment variable or paste it here.")
    up_pdf = st.file_uploader("Upload BEE / BRSR PDF", type=["pdf"])
    if st.button("🔍 Parse PDF & Append to Model", use_container_width=True):
        if not api_key:
            st.error("Please provide a GEMINI_API_KEY.")
        elif up_pdf is None:
            st.error("Please upload a PDF.")
        else:
            try:
                new_records = gemini_parse_pdf(up_pdf, api_key)
                if not new_records:
                    st.warning("No complete plant records were extracted from the PDF.")
                else:
                    new_df = pd.DataFrame(new_records, columns=MAIN_SHEET_COLS)
                    new_df = apply_sidebar_adjustments(
                        new_df, compliance_year, banking_pct, cuf_shock)
                    combined = pd.concat([sim_df, new_df], ignore_index=True) if sim_df is not None else new_df
                    st.session_state["df_plants"] = combined
                    st.success(f"Appended {len(new_df)} new plant(s). Total: {len(combined)}. "
                               "Click 'Run Full Model' to re-run.")
                    st.dataframe(new_df, use_container_width=True, hide_index=True)
            except Exception as e:
                st.error(f"Auto-update failed: {e}")

st.divider()
st.caption("Built for the CCTS Carbon Compliance Credit Monte Carlo & Price Discovery workflow. "
           "Random seed locked (42) for reproducible results.")


# -------------------- Tab 9: Market & Sector Analysis --------------------
with tab_market:
    st.subheader("Market & Sector Analysis")
    if sim_df is not None and res_df is not None and summary is not None:
        market_df = build_market_analysis_sheet(sim_df, res_df, summary)
        sec_table = build_sector_summary_table(sim_df, res_df)

        st.markdown("### Macro Market KPIs")
        k1, k2, k3, k4 = st.columns(4)
        k1.metric("Total Demand (CCC)", f"{market_df.iloc[0]['Value']:,.0f}")
        k2.metric("Total Supply (CCC)", f"{market_df.iloc[1]['Value']:,.0f}")
        k3.metric("Net Supply/Demand (CCC)", f"{market_df.iloc[2]['Value']:,.0f}")
        k4.metric("Liquidity Coverage Ratio", f"{market_df.iloc[4]['Value']:.1f}%")

        k5, k6, k7, k8 = st.columns(4)
        k5.metric("2024 → 2026 Production Change", f"{market_df.iloc[5]['Value']:+.2f}%")
        k6.metric("2024 → 2026 Emissions Change", f"{market_df.iloc[6]['Value']:+.2f}%")
        k7.metric("No Intensity Improvement (CCC)", f"{market_df.iloc[7]['Value']:,.0f}")
        k8.metric("No Intensity/Prod Improvement", f"{market_df.iloc[8]['Value']:,.0f}")

        st.markdown("### Sector-Wise Production, Emissions & Credit Balance (All 4 CCTS Sectors)")
        st.dataframe(
            sec_table.style.format({
                "Production_2024": "{:,.0f}", "Production_2026": "{:,.0f}", "Prod_Change_Pct": "{:+.2f}%",
                "Emissions_2024_tCO2": "{:,.0f}", "Emissions_2026_tCO2": "{:,.0f}", "Emissions_Change_Pct": "{:+.2f}%",
                "Expected_Purchases_CCC": "{:,.0f}", "Expected_Credits_CCC": "{:,.0f}", "Net_CCC_Balance": "{:+,.0f}",
            }),
            use_container_width=True, hide_index=True,
        )
    else:
        st.info("Run the model first to display Market & Sector Analysis.")

# -------------------- Tab 10: Company-Wise & Top 5 Rankings --------------------
with tab_company:
    st.subheader("Company-Wise & Top 5 Rankings")
    if sim_df is not None and res_df is not None:
        m_comp = res_df.copy()
        for col in ["Company", "Sector", "Plant"]:
            if col not in m_comp.columns and col in sim_df.columns:
                m_comp = m_comp.merge(sim_df[["Reg_Number", col]].drop_duplicates("Reg_Number"), on="Reg_Number", how="left")

        available_sectors = sorted(m_comp["Sector"].dropna().unique())
        selected_sector = st.selectbox("Filter by sector", ["All"] + available_sectors, key="sec_filter_company")
        m_filtered = m_comp[m_comp["Sector"] == selected_sector].copy() if selected_sector != "All" else m_comp.copy()

        c1, c2, c3 = st.columns(3)
        c1.metric("Companies Shown", f"{m_filtered['Company'].nunique():,}")
        c2.metric("Total Purchases / Demand", f"{m_filtered['Expected_Purchases_Demand'].sum():,.0f} CCC")
        c3.metric("Total Credits / Surplus", f"{m_filtered['Expected_Credits_Surplus'].sum():,.0f} CCC")

        agg_comp = m_filtered.groupby(["Company", "Sector"], as_index=False).agg(
            Facilities=("Reg_Number", "count"),
            Purchases_Demand=("Expected_Purchases_Demand", "sum"),
            Credits_Surplus=("Expected_Credits_Surplus", "sum"),
        )
        agg_comp["Net_CCC_Balance"] = agg_comp["Credits_Surplus"] - agg_comp["Purchases_Demand"]

        cb, cs = st.columns(2)
        with cb:
            st.markdown("### Top 5 Buyers (Expected Purchases)")
            st.dataframe(agg_comp.sort_values("Purchases_Demand", ascending=False).head(5)[["Company", "Sector", "Facilities", "Purchases_Demand"]].style.format({"Purchases_Demand": "{:,.0f}"}), use_container_width=True, hide_index=True)
        with cs:
            st.markdown("### Top 5 Sellers (Expected Credits / Surplus)")
            st.dataframe(agg_comp.sort_values("Credits_Surplus", ascending=False).head(5)[["Company", "Sector", "Facilities", "Credits_Surplus"]].style.format({"Credits_Surplus": "{:,.0f}"}), use_container_width=True, hide_index=True)

        st.markdown("### Full Company-Wise Breakdown")
        st.dataframe(agg_comp.sort_values("Purchases_Demand", ascending=False).style.format({"Purchases_Demand": "{:,.0f}", "Credits_Surplus": "{:,.0f}", "Net_CCC_Balance": "{:+,.0f}"}), use_container_width=True, hide_index=True)
    else:
        st.info("Run the model first to display Company-Wise rankings.")
