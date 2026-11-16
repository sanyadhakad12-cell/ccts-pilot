import csv
from datetime import datetime
import os
import random
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

print("\n" + "=" * 65)
print("    CCTS MONTE CARLO SIMULATION ENGINE (NET DEMAND & SURPLUS AUDIT)")
print("=" * 65 + "\n")

# 1. Configuration
N_SIMULATIONS = 10000
RANDOM_SEED = 42
INPUT_CSV = "ccts_plants_data.csv"
OUTPUT_XLSX = "CCTS_Monte_Carlo_Net_Demand_Surplus.xlsx"

random.seed(RANDOM_SEED)

if not os.path.exists(INPUT_CSV):
    print(f"ERROR: Cannot find '{INPUT_CSV}' in {os.getcwd()}")
    print("Please ensure your CSV is in the same folder and named correctly.")
    exit()

plants = []
skipped_rows = []


def clean_num(val):
    if val is None:
        return None
    v = str(val).replace(",", "").replace("%", "").strip()
    if not v:
        return None
    try:
        return float(v)
    except ValueError:
        return None


# 2. Ingest and Parse Data
current_company = ""

with open(INPUT_CSV, mode="r", encoding="utf-8-sig") as f:
    reader = csv.reader(f)

    for row_idx, row in enumerate(reader, start=1):
        # Must have sufficient columns to reach Column K (index 10)
        if not row or len(row) < 8:
            continue

        company_raw = row[0].strip() if len(row) > 0 else ""
        if company_raw:
            current_company = company_raw

        plant_name = row[2].strip() if len(row) > 2 else ""
        reg_num = row[3].strip() if len(row) > 3 else ""

        # --- STRICT FILTER: MUST HAVE A VALID REGISTRATION NUMBER IN COLUMN D ---
        # Skips summary rows, company group headers (e.g. rows 3 & 10), and empty spacer rows
        if not reg_num:
            continue

        # Skip table header row if repeated
        if reg_num.lower() in ["reg. number", "reg number", "registration number"]:
            continue

        # --- COLUMN MAPPING ACCORDING TO SCHEMA ---
        # Col A (0): Company Name
        # Col B (1): S. No.
        # Col C (2): Obligated Entity / Plant Name
        # Col D (3): Reg. Number
        # Col E (4): Baseline GHG Intensity 2023-24 (base_gei)
        # Col F (5): Target 2025-26 Intensity (target_gei)
        # Col G (6): 2024 equivalent Output / Baseline Output (base_prod)
        # Col H (7): 2026 equivalent product / Production (actual_prod)
        # Col I (8): Best Case / Minimum GEI
        # Col J (9): Median / Most Likely GEI
        # Col K (10): Worst Case / Maximum GEI

        base_gei = clean_num(row[4]) if len(row) > 4 else None
        target_gei = clean_num(row[5]) if len(row) > 5 else None
        base_prod = clean_num(row[6]) if len(row) > 6 else 0.0

        # Read 2026 Equivalent Product directly from Column H (no default multipliers)
        actual_prod = clean_num(row[7]) if len(row) > 7 else None

        # Read direct intensity potential range from Columns I, J, K (no default fallbacks)
        val_i = clean_num(row[8]) if len(row) > 8 else None   # Best Case / Min
        val_j = clean_num(row[9]) if len(row) > 9 else None   # Median / Mode
        val_k = clean_num(row[10]) if len(row) > 10 else None # Worst Case / Max

        # Validate that all required values are present
        if (
            base_gei is None
            or target_gei is None
            or actual_prod is None
            or val_i is None
            or val_j is None
            or val_k is None
        ):
            skipped_rows.append(
                (
                    row_idx,
                    reg_num,
                    plant_name,
                    "Missing Target/Baseline/Production or Min/Mode/Max Intensity values",
                )
            )
            continue

        # Establish bounds directly from the sheet columns without default assumptions
        min_gei = min(val_i, val_k)
        max_gei = max(val_i, val_k)
        # Clamp mode between min and max to ensure valid triangular distribution bounds
        mode_gei = max(min_gei, min(max_gei, val_j))

        unique_id = f"{reg_num} | {current_company} | {plant_name.split(',')[0]}"

        plants.append({
            "id": unique_id,
            "company": current_company,
            "plant": plant_name,
            "reg_num": reg_num,
            "base_prod": base_prod,
            "base_gei": base_gei,
            "target_gei": target_gei,
            "actual_prod": actual_prod,
            "min_gei": min_gei,
            "mode_gei": mode_gei,
            "max_gei": max_gei,
            "source_row": row_idx,
        })

print(f"--> Successfully loaded {len(plants)} operational facilities with Registration Numbers.")
if skipped_rows:
    print(f"--> Notice: Skipped {len(skipped_rows)} rows due to missing data:")
    for r_num, r_id, p_name, reason in skipped_rows[:10]:
        print(f"    Row {r_num} [{r_id} - {p_name}]: {reason}")

# 3. Fast Vectorized Monte Carlo Loop
print(
    f"\n--> Running {N_SIMULATIONS:,} joint iterations across all {len(plants)} plants..."
)

plant_balances = {p["id"]: [0.0] * N_SIMULATIONS for p in plants}
gross_shortfalls = [0.0] * N_SIMULATIONS
gross_surpluses = [0.0] * N_SIMULATIONS
net_balances = [0.0] * N_SIMULATIONS

for k in range(N_SIMULATIONS):
    run_net = 0.0
    run_shortfall = 0.0
    run_surplus = 0.0

    for p in plants:
        low, high, mode = p["min_gei"], p["max_gei"], p["mode_gei"]
        sim_gei = random.triangular(low, high, mode) if low != high else low

        # CCTS Statutory Balance: (Target - Actual) * Production
        bal = (p["target_gei"] - sim_gei) * p["actual_prod"]

        plant_balances[p["id"]][k] = bal
        run_net += bal

        if bal > 0:
            run_surplus += bal
        else:
            run_shortfall += abs(bal)

    gross_shortfalls[k] = run_shortfall
    gross_surpluses[k] = run_surplus
    net_balances[k] = run_net


def calc_stats(arr):
    s = sorted(arr)
    p10 = s[int(0.10 * N_SIMULATIONS)]
    median = s[int(0.50 * N_SIMULATIONS)]
    p90 = s[int(0.90 * N_SIMULATIONS)]
    mean = sum(arr) / N_SIMULATIONS
    shortfall_pct = (sum(1 for x in arr if x < 0) / N_SIMULATIONS) * 100
    expected_purchases = sum(abs(x) for x in arr if x < 0) / N_SIMULATIONS
    expected_credits = sum(x for x in arr if x > 0) / N_SIMULATIONS
    return (
        mean,
        p10,
        median,
        p90,
        shortfall_pct,
        expected_purchases,
        expected_credits,
    )


# 4. Generate Professional 2-Sheet Excel Workbook (Results & Inputs)
print("--> Generating streamlined 2-sheet audit workbook...")
wb = openpyxl.Workbook()

font_title = Font(name="Arial", size=13, bold=True, color="1F497D")
font_sub = Font(name="Arial", size=9, italic=True, color="595959")
font_header = Font(name="Arial", size=9, bold=True, color="FFFFFF")
font_bold = Font(name="Arial", size=9, bold=True)
font_regular = Font(name="Arial", size=9)
fill_header = PatternFill(
    start_color="1F497D", end_color="1F497D", fill_type="solid"
)
fill_total = PatternFill(
    start_color="D9E1F2", end_color="D9E1F2", fill_type="solid"
)
border_thin = Border(
    left=Side(style="thin", color="D9D9D9"),
    right=Side(style="thin", color="D9D9D9"),
    top=Side(style="thin", color="D9D9D9"),
    bottom=Side(style="thin", color="D9D9D9"),
)

# ----------------- Sheet 1: Results (Net Demand & Surplus) -----------------
ws_res = wb.active
ws_res.title = "Net Demand & Surplus"

ws_res["A1"] = "CCTS SECTOR-WIDE NET DEMAND & COMPLIANCE SURPLUS (2025–26)"
ws_res["A1"].font = font_title
ws_res["A2"] = (
    f"Evaluated: {len(plants)} Plants | {N_SIMULATIONS:,} Iterations | Random Seed {RANDOM_SEED}"
)
ws_res["A2"].font = font_sub
ws_res["A3"] = (
    "Positive Net Position = Compliance Surplus (CCCs generated); Negative Net Position = Compliance Deficit / Demand."
)
ws_res["A3"].font = font_sub
ws_res["A4"] = (
    "P10/P90 demarcate the central 80% confidence interval under empirical peer operational variance."
)
ws_res["A4"].font = font_sub

headers_res = [
    "Obligated Facility",
    "Mean Net Position (CCCs)",
    "P10 Net Position",
    "Median Net Position",
    "P90 Net Position",
    "Chance of Shortfall",
    "Expected Purchases / Demand",
    "Expected Credits / Surplus",
]

for col_idx, h in enumerate(headers_res, 1):
    c = ws_res.cell(row=5, column=col_idx, value=h)
    c.font = font_header
    c.fill = fill_header
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

curr_row = 6
for p in plants:
    mean, p10, med, p90, sf_pct, exp_purch, exp_cred = calc_stats(
        plant_balances[p["id"]]
    )

    row_data = [
        p["id"],
        mean,
        p10,
        med,
        p90,
        sf_pct / 100,
        exp_purch,
        exp_cred,
    ]
    for col_idx, val in enumerate(row_data, 1):
        c = ws_res.cell(row=curr_row, column=col_idx, value=val)
        c.font = font_regular
        c.border = border_thin
        if col_idx == 1:
            c.alignment = Alignment(horizontal="left", vertical="center")
        elif col_idx == 6:
            c.number_format = "0.0%"
            c.alignment = Alignment(horizontal="right", vertical="center")
        else:
            c.number_format = "#,##0;(#,##0);0"
            c.alignment = Alignment(horizontal="right", vertical="center")
    curr_row += 1

# Total Sector Portfolio Row
t_mean, t_p10, t_med, t_p90, t_sf_pct, _, _ = calc_stats(net_balances)
total_gross_purch = sum(gross_shortfalls) / N_SIMULATIONS
total_gross_cred = sum(gross_surpluses) / N_SIMULATIONS

tot_row_data = [
    f"TOTAL SECTOR POSITION ({len(plants)} PLANTS)",
    t_mean,
    t_p10,
    t_med,
    t_p90,
    t_sf_pct / 100,
    total_gross_purch,
    total_gross_cred,
]
for col_idx, val in enumerate(tot_row_data, 1):
    c = ws_res.cell(row=curr_row, column=col_idx, value=val)
    c.font = font_bold
    c.fill = fill_total
    c.border = border_thin
    if col_idx == 1:
        c.alignment = Alignment(horizontal="left", vertical="center")
    elif col_idx == 6:
        c.number_format = "0.0%"
        c.alignment = Alignment(horizontal="right", vertical="center")
    else:
        c.number_format = "#,##0;(#,##0);0"
        c.alignment = Alignment(horizontal="right", vertical="center")

# ----------------- Sheet 2: Inputs Used -----------------
print("--> Archiving inputs into Sheet 2...")
ws_inp = wb.create_sheet(title="Inputs used")
ws_inp["A1"] = "CCTS MODEL AUDIT LOG & RECONSTRUCTED BASELINES"
ws_inp["A1"].font = font_bold

meta = [
    ("Simulation Runs", N_SIMULATIONS),
    ("Random Seed", RANDOM_SEED),
    ("Compliance Period", "2025–26"),
    ("Source Data File", INPUT_CSV),
    ("Execution Timestamp", datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
    ("Total Plants Evaluated", len(plants)),
]
for idx, (lbl, val) in enumerate(meta, 3):
    ws_inp.cell(row=idx, column=1, value=lbl).font = font_bold
    ws_inp.cell(row=idx, column=2, value=str(val)).font = font_regular

inp_headers = [
    "Reg. Number",
    "Facility Name",
    "Baseline Output (t)",
    "Baseline GEI",
    "Target GEI",
    "Actual FY26 Output (t)",
    "Min GEI Used",
    "Mode GEI Used",
    "Max GEI Used",
]
for col_idx, h in enumerate(inp_headers, 1):
    c = ws_inp.cell(row=10, column=col_idx, value=h)
    c.font = font_header
    c.fill = fill_header
    c.alignment = Alignment(horizontal="center", vertical="center")

for r_idx, p in enumerate(plants, 11):
    vals = [
        p["reg_num"],
        p["plant"],
        p["base_prod"],
        p["base_gei"],
        p["target_gei"],
        p["actual_prod"],
        p["min_gei"],
        p["mode_gei"],
        p["max_gei"],
    ]
    for col_idx, v in enumerate(vals, 1):
        c = ws_inp.cell(row=r_idx, column=col_idx, value=v)
        c.font = font_regular
        c.border = border_thin
        if col_idx in [3, 6]:
            c.number_format = "#,##0"
        elif col_idx in [4, 5, 7, 8, 9]:
            c.number_format = "0.0000"

for ws in [ws_res, ws_inp]:
    for col in ws.columns:
        max_len = max(len(str(c.value or "")) for c in col[:50])
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 3, 14)

wb.save(OUTPUT_XLSX)
print(f"\n--> SUCCESS! Evaluated {len(plants)} plants. Saved to '{OUTPUT_XLSX}'.")
print("=" * 65 + "\n")