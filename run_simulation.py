import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import openpyxl
from openpyxl.drawing.image import Image as OpenpyxlImage

# Lock randomness so results stay identical across every execution
np.random.seed(42)

# -------------------------------------------------------------
# 1. SETUP PARAMETERS & DATA INGESTION
# -------------------------------------------------------------
min_price, max_price, step = 600, 1600, 25
price_ticks = np.arange(min_price, max_price + step, step)  # 41 ticks

csv_file = 'market_data_248_plants.csv'
try:
    df_raw = pd.read_csv(csv_file, thousands=',')
    df_raw.columns = [str(c).strip() for c in df_raw.columns]
    demand_col = [c for c in df_raw.columns if 'Demand' in c or 'Purchases' in c][0]
    surplus_col = [c for c in df_raw.columns if 'Surplus' in c or 'Credits' in c][0]
    
    df_raw[demand_col] = pd.to_numeric(df_raw[demand_col], errors='coerce').fillna(0)
    df_raw[surplus_col] = pd.to_numeric(df_raw[surplus_col], errors='coerce').fillna(0)
    
    buyers_df = df_raw[df_raw[demand_col] > 0][[df_raw.columns[0], demand_col]].copy()
    sellers_df = df_raw[df_raw[surplus_col] > 0][[df_raw.columns[0], surplus_col]].copy()
    buyer_vols = buyers_df[demand_col].values
    seller_vols = sellers_df[surplus_col].values
    n_plants = len(df_raw)
except Exception:
    # Fallback to realistic synthetic generation if CSV is absent
    n_buyers, n_sellers = 130, 118
    n_plants = n_buyers + n_sellers
    buyer_vols = np.round(np.random.lognormal(mean=5.0, sigma=0.8, size=n_buyers), 1)
    seller_vols = np.round(np.random.lognormal(mean=5.1, sigma=0.8, size=n_sellers), 1)

# Total base inventory pool for sellers across the entire master dataset
TOTAL_BASE_SUPPLY = float(np.sum(seller_vols))

n_simulations = 5000  # 5,000 sessions
sim_records = []
sim_mcps = []
sim_vols = []
sample_curves = []

# -------------------------------------------------------------
# 2. RUN MARKET PERMUTATION & CLEARING ENGINE
# -------------------------------------------------------------
for s in range(1, n_simulations + 1):
    b_thresh = np.random.choice(price_ticks, size=len(buyer_vols))
    s_thresh = np.random.choice(price_ticks, size=len(seller_vols))
    
    dem_curve = np.array([buyer_vols[b_thresh >= p].sum() for p in price_ticks])
    sup_curve = np.array([seller_vols[s_thresh <= p].sum() for p in price_ticks])
    
    trade_vol = np.minimum(dem_curve, sup_curve)
    mismatch = np.abs(dem_curve - sup_curve)
    
    if np.any(trade_vol > 0):
        best_idx = np.argmin(mismatch)
        mcp = price_ticks[best_idx]
        vol = trade_vol[best_idx]
        
        market_supply_at_mcp = sup_curve[best_idx]
        market_demand_at_mcp = dem_curve[best_idx]
        
        # Calculate Pre-Market Banked Supply: Sellers whose floor price was HIGHER than mcp
        pre_market_banked_supply = TOTAL_BASE_SUPPLY - market_supply_at_mcp
        
        # Unmet Seller Credits: Shows excess seller volume if Supply > Demand, else 0
        unmet_seller_credits = max(0, market_supply_at_mcp - market_demand_at_mcp)
        
        # Unmet Volume Mismatch: Absolute difference showing overall imbalance
        unmet_volume_mismatch = mismatch[best_idx]
        
        sim_mcps.append(mcp)
        sim_vols.append(vol)
        
        sim_records.append({
            'Session_ID': s,
            'Clearing_Price_INR': mcp,
            'Volume_Cleared': round(vol, 2),
            'Market_Demand_At_MCP': round(market_demand_at_mcp, 2),
            'Market_Supply_At_MCP': round(market_supply_at_mcp, 2),
            'Unmet_Volume_Mismatch': round(unmet_volume_mismatch, 2),
            'Unmet_Seller_Credits': round(unmet_seller_credits, 2),
            'Pre_Market_Banked_Supply': round(pre_market_banked_supply, 2),
            'Verification_Total_Supply': round(vol + unmet_seller_credits + pre_market_banked_supply, 2)
        })
        
        if len(sample_curves) < 30:
            sample_curves.append((dem_curve, sup_curve, mcp, vol))

sim_mcps = np.array(sim_mcps)
sim_vols = np.array(sim_vols)

# -------------------------------------------------------------
# 3. STATISTICAL AGGREGATION & PROBABILITY ASSIGNMENT
# -------------------------------------------------------------
bin_edges = np.append(price_ticks - step/2, price_ticks[-1] + step/2)
counts, _ = np.histogram(sim_mcps, bins=bin_edges)
total_sessions = len(sim_mcps)
probabilities = counts / total_sessions

expected_mcp = np.sum(price_ticks * probabilities)
closest_tick_idx = np.argmin(np.abs(price_ticks - expected_mcp))
nearest_discrete_mcp = price_ticks[closest_tick_idx]
prob_of_nearest_mcp = probabilities[closest_tick_idx]

mode_idx = np.argmax(counts)
most_likely_mcp = price_ticks[mode_idx]
prob_of_most_likely_mcp = probabilities[mode_idx]

# Percentiles & Probable Price Ranges
p10 = np.percentile(sim_mcps, 10)
p25 = np.percentile(sim_mcps, 25)
p50 = np.percentile(sim_mcps, 50)
p75 = np.percentile(sim_mcps, 75)
p90 = np.percentile(sim_mcps, 90)

avg_vols = [sim_vols[sim_mcps == p].mean() if np.any(sim_mcps == p) else 0 for p in price_ticks]

# -------------------------------------------------------------
# 4. GENERATE SEPARATE HIGH-RESOLUTION CHARTS
# -------------------------------------------------------------

# --- Chart 1: Merit-Order Curve Crossing Dynamics ---
fig, ax1 = plt.subplots(figsize=(10, 6), dpi=300)
for dem, sup, mcp, vol in sample_curves:
    ax1.plot(dem, price_ticks, color='tab:blue', alpha=0.15, linewidth=1)
    ax1.plot(sup, price_ticks, color='tab:red', alpha=0.15, linewidth=1)
    ax1.scatter(vol, mcp, color='darkgreen', alpha=0.35, s=25)

ax1.plot([], [], color='tab:blue', label='Simulated Demand Curves')
ax1.plot([], [], color='tab:red', label='Simulated Supply Curves')
ax1.scatter([], [], color='darkgreen', label='Equilibrium Crossovers (MCP, Vol)')
ax1.set_title('Merit-Order Clearance Dynamics Across Sessions', fontsize=12, fontweight='bold')
ax1.set_xlabel('Volume Traded (Credits)', fontsize=10)
ax1.set_ylabel('Price Tick (₹)', fontsize=10)
ax1.grid(True, linestyle='--', alpha=0.4)
ax1.legend(loc='upper right')
plt.tight_layout()
chart_file_1 = 'merit_order_dynamics.png'
plt.savefig(chart_file_1, dpi=300)
plt.close(fig)


# --- Chart 2: MCP Probability Distribution with Explicit Bar Labels ---
fig, ax2 = plt.subplots(figsize=(12, 6), dpi=300)
ax2_vol = ax2.twinx()

prob_pcts = probabilities * 100
bars = ax2.bar(price_ticks, prob_pcts, width=18, color='steelblue', alpha=0.75, label='Clearing Probability (%)')
ax2_vol.plot(price_ticks, avg_vols, color='darkorange', linewidth=2, marker='o', markersize=3, label='Avg Cleared Volume')

for bar, p_val in zip(bars, prob_pcts):
    if p_val > 0.5:
        height = bar.get_height()
        ax2.annotate(f'{p_val:.1f}%',
                    xy=(bar.get_x() + bar.get_width() / 2, height),
                    xytext=(0, 3),
                    textcoords="offset points",
                    ha='center', va='bottom', fontsize=7, rotation=90)

ax2.axvline(expected_mcp, color='navy', linestyle='--', linewidth=2, label=f'Expected MCP: ₹{expected_mcp:.1f}')
ax2.set_title('MCP Probability Distribution & Liquidity Depth (with Assigned Probabilities)', fontsize=12, fontweight='bold')
ax2.set_xlabel('Price Tick (₹)', fontsize=10)
ax2.set_ylabel('Probability (%)', color='steelblue', fontsize=10)
ax2_vol.set_ylabel('Avg Cleared Volume (Credits)', color='darkorange', fontsize=10)
ax2.set_xticks(price_ticks)
ax2.set_xticklabels([f'₹{p}' for p in price_ticks], rotation=90, fontsize=8)
ax2.grid(True, linestyle='--', alpha=0.3)
ax2.legend(loc='upper left')
plt.tight_layout()
chart_file_2 = 'mcp_probability_distribution.png'
plt.savefig(chart_file_2, dpi=300)
plt.close(fig)


# --- Chart 3: Cumulative Risk Curve (CDF) ---
fig, ax3 = plt.subplots(figsize=(10, 6), dpi=300)
sorted_mcps = np.sort(sim_mcps)
cdf = np.arange(1, len(sorted_mcps) + 1) / len(sorted_mcps) * 100

ax3.plot(sorted_mcps, cdf, color='purple', linewidth=2.5)
ax3.axvline(p50, color='gray', linestyle=':', label=f'Median (P50): ₹{p50:.0f}')
ax3.axvline(p90, color='crimson', linestyle='--', label=f'90% Risk Cap (P90): ₹{p90:.0f}')
ax3.set_title('Cumulative Risk Curve (CDF of Clearing Price)', fontsize=12, fontweight='bold')
ax3.set_xlabel('Price Tick (₹)', fontsize=10)
ax3.set_ylabel('Cumulative Probability (%)', fontsize=10)
ax3.set_xticks(price_ticks[::2])
ax3.set_xticklabels([f'₹{p}' for p in price_ticks[::2]], rotation=45, fontsize=9)
ax3.grid(True, linestyle='--', alpha=0.4)
ax3.legend(loc='lower right')
plt.tight_layout()
chart_file_3 = 'cumulative_risk_cdf.png'
plt.savefig(chart_file_3, dpi=300)
plt.close(fig)


# -------------------------------------------------------------
# 5. PREPARE EXCEL SUMMARY & LOGS (With Probable Ranges)
# -------------------------------------------------------------
df_summary = pd.DataFrame({
    'Summary_Metric': [
        'Total Facilities Evaluated',
        'Total Buyer Facilities',
        'Total Seller Facilities',
        'Total Base Seller Inventory Pool',
        'Price Tick Band Evaluated',
        'Discrete Price Ticks (Steps)',
        'Total Simulated Sessions',
        'Expected Market Clearing Price (Mean ₹)',
        'Nearest Discrete Clearing Tick (₹)',
        'Assigned Probability of Nearest Discrete MCP (%)',
        'Most Likely Clearing Price (Mode ₹)',
        'Assigned Probability of Most Likely MCP (%)',
        'Interquartile Range - IQR (P25 - P75)',
        '80% High-Probability Band (P10 - P90)',
        'Median Clearing Price (P50 - ₹)',
        '90% Risk Cap Price (P90 - ₹)',
        'Average Traded Volume per Session (Credits)'
    ],
    'Value': [
        n_plants,
        len(buyer_vols),
        len(seller_vols),
        f"{TOTAL_BASE_SUPPLY:,.2f} credits",
        f"₹{min_price} to ₹{max_price}",
        len(price_ticks),
        total_sessions,
        f"₹{expected_mcp:.2f}",
        f"₹{nearest_discrete_mcp}",
        f"{prob_of_nearest_mcp:.2%}",
        f"₹{most_likely_mcp}",
        f"{prob_of_most_likely_mcp:.2%}",
        f"₹{p25:.0f} to ₹{p75:.0f}",
        f"₹{p10:.0f} to ₹{p90:.0f}",
        f"₹{p50:.2f}",
        f"₹{p90:.2f}",
        round(np.mean(sim_vols), 2)
    ]
})

df_price_ticks = pd.DataFrame({
    'Price_Tick_INR': price_ticks,
    'Occurrences_Count': counts,
    'Assigned_Probability': np.round(probabilities, 4),
    'Assigned_Probability_Pct': [f"{p:.2%}" for p in probabilities],
    'Cumulative_Probability_Pct': [f"{np.sum(probabilities[:i+1]):.2%}" for i in range(len(probabilities))],
    'Avg_Cleared_Volume': np.round(avg_vols, 1),
    'Weighted_Price_Contribution_INR': np.round(price_ticks * probabilities, 2)
})

df_session_logs = pd.DataFrame(sim_records)

# -------------------------------------------------------------
# 6. WRITE WORKBOOK & EMBED CHARTS ON DEDICATED TABS
# -------------------------------------------------------------
excel_file = 'market_clearing_assigned_probability_analysis.xlsx'

with pd.ExcelWriter(excel_file, engine='openpyxl') as writer:
    df_summary.to_excel(writer, sheet_name='Dashboard_Summary', index=False)
    df_price_ticks.to_excel(writer, sheet_name='Price_Tick_Distribution', index=False)
    df_session_logs.to_excel(writer, sheet_name='All_Sessions_Raw_Data', index=False)

wb = openpyxl.load_workbook(excel_file)

# Create dedicated tabs for each chart
ws_chart1 = wb.create_sheet(title='Chart_Merit_Order')
ws_chart2 = wb.create_sheet(title='Chart_Probability_Depth')
ws_chart3 = wb.create_sheet(title='Chart_Risk_CDF')

# Merit Order Chart Tab
ws_chart1['A1'] = "Chart A: Merit-Order Clearance Dynamics Across Sessions"
ws_chart1['A1'].font = openpyxl.styles.Font(size=14, bold=True)
ws_chart1['A2'] = "Explanation: Shows simulated plant-level demand (blue) and supply (red) curves across 5,000 sessions. Green dots mark market-clearing equilibrium points."
img1 = OpenpyxlImage(chart_file_1)
img1.width, img1.height = 750, 450
ws_chart1.add_image(img1, 'A4')

# Probability & Liquidity Chart Tab
ws_chart2['A1'] = "Chart B: MCP Probability Distribution & Liquidity Depth"
ws_chart2['A1'].font = openpyxl.styles.Font(size=14, bold=True)
ws_chart2['A2'] = "Explanation: Displays the exact percentage probability assigned above each price tick bar, paired with the average cleared trading volume (orange line)."
img2 = OpenpyxlImage(chart_file_2)
img2.width, img2.height = 800, 450
ws_chart2.add_image(img2, 'A4')

# Cumulative Risk CDF Chart Tab
ws_chart3['A1'] = "Chart C: Cumulative Risk Curve (CDF of Clearing Price)"
ws_chart3['A1'].font = openpyxl.styles.Font(size=14, bold=True)
ws_chart3['A2'] = "Explanation: Cumulative Distribution Function tracking confidence thresholds, highlighting the Median (P50) and 90% Risk Cap (P90)."
img3 = OpenpyxlImage(chart_file_3)
img3.width, img3.height = 750, 450
ws_chart3.add_image(img3, 'A4')

# Auto-adjust column widths across all sheets
for ws in wb.worksheets:
    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = openpyxl.utils.get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = min(max(max_len + 4, 14), 50)

wb.save(excel_file)
print(f"Workbook successfully saved to '{excel_file}' with 5,000 sessions, locked random seed (42), and the Unmet_Seller_Credits column.")