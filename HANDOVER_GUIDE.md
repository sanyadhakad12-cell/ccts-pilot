# Climactio — CCTS Institutional Platform & Simulation Engine
### Complete Product Handover & Deployment Guide
**Organization:** Climactio Private Services Limited  
**Version:** Institutional Edition 2.0  
**Compliance Cycle:** CCTS 2026 (Projections 2024–2030)

---

## 1. Executive Summary

This package contains the complete, production-ready institutional software platform for India's **Carbon Credit Trading Scheme (CCTS)**, developed for **Climactio Private Services Limited**.

The platform is designed for two distinct audiences:
1. **Market Authority & Policy Simulation Engine (Admin Portal)**: Used by regulators, power exchanges, and market policymakers to simulate national market liquidity, model double-auction price discovery across 248 facilities, manage regulatory gazette banners, and review corporate multi-year intake filings.
2. **Company-Specific Client Portal**: A dedicated executive advisory dashboard for corporate clients (e.g., UltraTech, Tata Steel, ACC, Ambuja, JSW, Shree Cement, etc.) to evaluate installation-level deficits/surpluses, model Marginal Abatement Cost Curves (MACC), assess EU/UK CBAM border tax liabilities, submit 2027–2029 decarbonization intake forms, and generate publication-quality 5-page PDF advisory reports.

---

## 2. Directory & File Manifest

```
ccts-app/
├── app.py                     # Main Streamlit web application & UI orchestrator
├── requirements.txt           # Python dependencies (NumPy, Pandas, Plotly, Streamlit, etc.)
├── HANDOVER_GUIDE.md          # Complete handover & setup instructions (this file)
│
├── 1. Results.xlsx            # Reconstructed simulation results & company benchmarks
├── 3. Data sheet.xlsx         # Master national facility registry (248 plants, GEI, baselines)
├── MACC for cline.xlsx        # National sector MACC levers (Cement, Steel, Aluminium, etc.)
├── market_data_248_plants.csv # Flat dataset of regulated facilities
│
├── assets/
│   └── climactio_logo.png     # Official high-resolution Climactio brand logo
│
├── config/
│   ├── company_credentials.json # Client authentication & company list (pilot mode)
│   ├── market_news.json         # Live regulatory announcements & ticker database
│   ├── year_rules.json          # Statutory trajectory rules, FX rates, CBAM sectors
│   └── intake_submissions/      # JSON repository of corporate multi-year filings (2027–2029)
│
├── engine/
│   ├── ccc_sim.py             # 10,000-iteration triangular Monte Carlo demand/supply engine
│   ├── price_discovery.py     # 5,000-session sealed double-auction price equilibrium engine
│   ├── macc_cbam.py           # MACC project curves, EU/UK CBAM border tax calculations
│   ├── data_pipeline.py       # Excel loaders, data normalizers, sector classifiers
│   └── reports.py             # Macro analytics, peer tables, ranking extractors
│
├── 4. Methodology - Logic and Assumptions.pdf  # Technical methodology reference document
├── price discovery logic.pdf                   # Mathematical clearing price specification
├── 5. ccts_monte_carlo.py                      # Original standalone prototype script
└── run_simulation.py                           # CLI batch runner for automated simulations
```

---

## 3. Quick-Start Guide (Run on Any Laptop/Server)

The platform runs on **macOS, Windows, and Linux** using Python 3.10, 3.11, 3.12, 3.13, or 3.14.

### Step 1: Open Terminal / Command Prompt
* **Mac/Linux**: Open Terminal.
* **Windows**: Open PowerShell or Command Prompt.

Navigate to the unzipped `ccts-app` folder:
```bash
cd /path/to/ccts-app
```

### Step 2: Create a Clean Python Virtual Environment (Recommended)
```bash
# Mac / Linux:
python3 -m venv .venv
source .venv/bin/activate

# Windows:
python -m venv .venv
.venv\Scripts\activate
```

### Step 3: Install Required Dependencies
```bash
pip install -r requirements.txt
```
*(Takes under 60 seconds. Installs `streamlit`, `pandas`, `numpy`, `openpyxl`, `plotly`, `matplotlib`, `pdfplumber`)*

### Step 4: Launch the Platform
```bash
streamlit run app.py
```

Streamlit will automatically open your web browser to:
👉 **`http://localhost:8501`**

---

## 4. Key Functional Modules & Usage

### A. Dual Login Selection
* **🏛️ Market Login (Admin Authority)**:
  * Pilot Testing Mode: Password block is present for presentation, but pressing **Enter Market Authority Portal** grants direct access with no password required.
* **🏢 Company Specific Client Portal**:
  * Select any client (e.g. *UltraTech Cement Limited*, *Tata Steel*, *ACC Limited*).
  * Password block is present for institutional feel, but entering with blank/any password works directly in pilot testing mode.

### B. Company Specific Client Portal (Key Tabs)
1. **Header Layout**: Features the client's name prominently centered in large typography, with the official Climactio logo and brand badge elegantly positioned.
2. **Tab 1 — Portfolio Asset Registry**: Interactive grid where users can double-click any cell to tweak their actual installation GEIs or output.
3. **Tab 2 — Monte Carlo Compliance Position**: Simulates 10,000 iterations to determine facility-by-facility probability of shortfall.
4. **Tab 3 — Price Discovery & Financial Exposure**:
   * Evaluates clearing price (E[MCP], Most Likely, P50, P90).
   * **Strict Facility-Level Accounting**: Enforces the statutory CCTS rule that **no inter-netting across separate installations is permitted**.
   * Shows exact purchase outlays and surplus values in **₹ Crores** alongside exact rupee values.
5. **Tab 4 — MACC Abatement Studio**: Lets companies toggle internal abatement projects (WHR, alternative fuels, energy efficiency) and lock sealed auction bids.
6. **Tab 6 — EU / UK CBAM Exposure**: Calculates cross-border border tax liabilities under EU ETS (€85/tCO₂) and UK CBAM (£45/tCO₂).
7. **Tab 8 — Commercial Advisory Report**: Displays an executive briefing on-screen and offers a 1-click **Download Publication-Quality 5-Page PDF Report** branded with Climactio.
8. **Tab 9 — Multi-Year Decarbonization Form (2027–2029)**: Official intake protocol for production CAGR, TSR, WHRS, and RE targets. Supports partial and draft filings with zero blocking errors; immediately syncs to Admin.

### C. Market Authority Admin Portal (Key Tabs)
1. **Tab 3 — Price Discovery (All 3 Models)**: Merit-Order X-Curve, Probability & Traded Liquidity Depth, and Cumulative Risk Distribution (CDF).
2. **Tab 7 — Regulatory News Manager**: 
   * Post new regulatory alerts and compliance notices.
   * **Interactive Checkbox Control**: Tick or untick the **Active Alert (Show in Banner)** checkbox directly in the table to instantly control which announcements appear on the top banner across all client dashboards.
   * Delete or clear announcements with one click.
3. **Tab 8 — Multi-Year Data Intake Submissions**: Live registry of all submissions completed by companies in Tab 9, with full JSON inspection.
4. **Sidebar Quick-Tweak**: Regulators can select any company and installation in India, tweak its achieved production and intensity, and re-run simulations on demand.

---

## 5. Important Statutory & Methodological Rules Codified

1. **Baseline-and-Credit Scheme**: India's CCTS assesses liability on the **compliance shortfall** ($\text{CCCs} = (\text{Target GEI} - \text{Achieved GEI}) \times \text{Production}$), **not** on gross facility emissions.
2. **No Inter-Netting**: Under Bureau of Energy Efficiency (BEE) rules, designated consumers (DCs) cannot cancel out deficits between separate facilities off-market. Each shortfall facility must procure and surrender CCCs independently.
3. **Penalty Multiplier**: Shortfall CCCs not surrendered by the statutory verification deadline face an environmental penalty calculated at **$2.0\times$ the average market clearing price**.

---

## 6. How to Share with Team Members

1. **Email / Slack / Google Drive / Teams**: Send the generated **`ccts_app_complete_package.zip`** file.
2. Tell your team member:
   > *"Unzip the archive, open terminal in that folder, run `pip install -r requirements.txt`, then `streamlit run app.py`."*
3. All data sheets, brand logos, pre-calculated results, and simulation engines are bundled inside the zip file — no external database or cloud connection is required.

---
**Climactio Private Services Limited**  
*Confidential — For Internal & Institutional Client Use Only*
