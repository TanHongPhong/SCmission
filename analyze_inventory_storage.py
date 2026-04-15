"""
Inventory Storage Analysis & Warehouse Solution Proposal
=========================================================
Based on MPS weekly ending FG inventory, compute:
- Peak/avg storage volume (CBM & pallets)
- Duration of storage need
- Cost comparison: Short-term rent vs Long-term rent vs Build
"""

# ── Constants ──────────────────────────────────────────────────────
# CBM per 100 cartons from FGs sheet
SKU_PACK = {
    'A': {'pcs_per_ctn': 4,  'cbm_100ctn': 1.0},
    'B': {'pcs_per_ctn': 4,  'cbm_100ctn': 0.625},
    'C': {'pcs_per_ctn': 2,  'cbm_100ctn': 1.0},
    'D': {'pcs_per_ctn': 4,  'cbm_100ctn': 0.625},
    'E': {'pcs_per_ctn': 2,  'cbm_100ctn': 0.625},
    'F': {'pcs_per_ctn': 8,  'cbm_100ctn': 0.78125},
}

def units_to_cbm(sku, units):
    p = SKU_PACK[sku]
    cartons = units / p['pcs_per_ctn']
    cbm = cartons * p['cbm_100ctn'] / 100
    return cbm

# Weekly ending inventory from MPS
weekly_inv = {
    '2026-04-13': {'A':0,      'B':12751, 'C':5000, 'D':0,    'E':0, 'F':3000},
    '2026-04-20': {'A':0,      'B':15000, 'C':5000, 'D':5000, 'E':0, 'F':3000},
    '2026-04-27': {'A':10000,  'B':15000, 'C':5000, 'D':5000, 'E':0, 'F':1500},
    '2026-05-04': {'A':5000,   'B':9498,  'C':5000, 'D':3250, 'E':0, 'F':0},
    '2026-05-11': {'A':0,      'B':0,     'C':1630, 'D':0,    'E':0, 'F':0},
    '2026-06-01': {'A':0,      'B':0,     'C':1760, 'D':0,    'E':0, 'F':0},
    '2026-06-08': {'A':0,      'B':0,     'C':318,  'D':0,    'E':0, 'F':0},
    '2026-06-15': {'A':0,      'B':0,     'C':0,    'D':0,    'E':0, 'F':500},
    '2026-06-22': {'A':0,      'B':81,    'C':0,    'D':0,    'E':0, 'F':0},
    '2026-07-13': {'A':9246,   'B':0,     'C':0,    'D':0,    'E':0, 'F':0},
    '2026-07-20': {'A':15000,  'B':2400,  'C':0,    'D':0,    'E':0, 'F':0},
    '2026-07-27': {'A':2240,   'B':0,     'C':0,    'D':0,    'E':0, 'F':0},
    '2026-09-21': {'A':0,      'B':9709,  'C':0,    'D':0,    'E':0, 'F':0},
    '2026-09-28': {'A':4067,   'B':15000, 'C':0,    'D':0,    'E':0, 'F':0},
    '2026-10-05': {'A':3050,   'B':12000, 'C':0,    'D':0,    'E':0, 'F':0},
    '2026-10-12': {'A':2033,   'B':9000,  'C':0,    'D':0,    'E':0, 'F':0},
    '2026-10-19': {'A':1017,   'B':6000,  'C':0,    'D':0,    'E':0, 'F':0},
}

# ── Compute CBM per week ───────────────────────────────────────────
PALLET_CBM    = 1.5    # m³ per pallet (standard EU pallet stacked 1.5m)
PALLET_FLOOR  = 1.2    # m² per pallet footprint (1.2m × 1.0m)
RACK_M2_RATIO = 1.8    # gross floor m² per pallet (aisles + structure)

print("="*72)
print("  WAREHOUSE SIZING & SOLUTION ANALYSIS")
print("="*72)

print("\n[1] WEEKLY STORAGE VOLUME")
print("{:>12} {:>8} {:>8} {:>8} {:>8} {:>8} {:>8}  {:>8} {:>8}".format(
    "Week","A","B","C","D","E","F","CBM","Pallets"))
print("-"*80)

cbm_by_week = {}
for wk in sorted(weekly_inv.keys()):
    inv = weekly_inv[wk]
    total_cbm = sum(units_to_cbm(s, inv.get(s,0)) for s in 'ABCDEF')
    pallets = total_cbm / PALLET_CBM
    cbm_by_week[wk] = total_cbm
    total_units = sum(inv.values())
    if total_units > 0:
        print("{:>12} {:>8,} {:>8,} {:>8,} {:>8,} {:>8,} {:>8,}  {:>7.1f} {:>7.0f}".format(
            wk,
            inv.get('A',0), inv.get('B',0), inv.get('C',0),
            inv.get('D',0), inv.get('E',0), inv.get('F',0),
            total_cbm, pallets))

cbm_vals = list(cbm_by_week.values())
peak_cbm  = max(cbm_vals)
avg_cbm   = sum(cbm_vals) / len(cbm_vals)
peak_wk   = max(cbm_by_week, key=cbm_by_week.get)
peak_pal  = peak_cbm / PALLET_CBM
avg_pal   = avg_cbm / PALLET_CBM
peak_m2   = peak_pal * RACK_M2_RATIO
avg_m2    = avg_pal  * RACK_M2_RATIO

# High-inv periods
high_periods = [wk for wk, c in cbm_by_week.items() if c > peak_cbm * 0.5]

print()
print("[2] STORAGE SIZING SUMMARY")
print("-"*50)
print("  Peak storage    : {:>8.1f} CBM  = {:>5.0f} pallets  = {:>6.0f} m²".format(peak_cbm, peak_pal, peak_m2))
print("  Average storage : {:>8.1f} CBM  = {:>5.0f} pallets  = {:>6.0f} m²".format(avg_cbm, avg_pal, avg_m2))
print("  Peak week       : {}".format(peak_wk))
print("  Weeks > 50% peak: {} weeks".format(len(high_periods)))
print("  High demand wks : {}".format(', '.join(high_periods[:6])))

# ── Cost Analysis ─────────────────────────────────────────────────
print()
print("[3] COST COMPARISON (USD, VN market assumptions)")
print("="*72)

# Assumptions (Vietnam industrial property)
SHORT_COST_PER_M2_MO  = 8.0    # USD/m²/month short-term 3PLR (~200k VND/m²/mo)
LONG_COST_PER_M2_MO   = 5.0    # USD/m²/month long-term lease 3-5yr contract
BUILD_COST_PER_M2     = 500.0  # USD/m² construction cost (basic warehouse VN)
BUILD_MAINTENANCE_MO  = 0.5    # USD/m²/month O&M
ANALYSIS_MONTHS       = 12     # planning horizon
BUILD_LIFESPAN_YR     = 20     # years

# Size needed: peak m² for short-term, avg m² for long-term (buffer 30%)
SHORT_SIZE = peak_m2 * 1.1    # rent peak capacity
LONG_SIZE  = (avg_m2 + peak_m2) / 2 * 1.2  # rent medium capacity
BUILD_SIZE = peak_m2 * 1.2    # build for full capacity

# Short-term rent: only pay for active months (7 months with inventory)
active_months = 7
short_total_1yr = SHORT_COST_PER_M2_MO * SHORT_SIZE * active_months
short_total_3yr = SHORT_COST_PER_M2_MO * SHORT_SIZE * (active_months * 3)

# Long-term rent: pay full 12 months/year
long_total_1yr = LONG_COST_PER_M2_MO * LONG_SIZE * 12
long_total_3yr = LONG_COST_PER_M2_MO * LONG_SIZE * 36

# Build: capex + 20yr horizon
build_capex     = BUILD_COST_PER_M2 * BUILD_SIZE
build_opex_1yr  = BUILD_MAINTENANCE_MO * BUILD_SIZE * 12
build_opex_3yr  = build_opex_1yr * 3
build_total_3yr = build_capex + build_opex_3yr
build_total_20yr= build_capex + (build_opex_1yr * 20)
break_even_vs_long = build_capex / (LONG_COST_PER_M2_MO * LONG_SIZE * 12 - build_opex_1yr)

print()
print("  Assumptions (Vietnam industrial warehousing):")
print("    Short-term 3PL rate : {:.0f} USD/m²/month".format(SHORT_COST_PER_M2_MO))
print("    Long-term lease rate: {:.0f} USD/m²/month (3-5 year contract)".format(LONG_COST_PER_M2_MO))
print("    Build cost          : {:.0f} USD/m² (basic warehouse)".format(BUILD_COST_PER_M2))
print("    Build O&M           : {:.1f} USD/m²/month".format(BUILD_MAINTENANCE_MO))
print()
print("  Storage size required:")
print("    Short-term (peak)   : {:>6.0f} m²".format(SHORT_SIZE))
print("    Long-term (blended) : {:>6.0f} m²".format(LONG_SIZE))
print("    Build (peak+buffer) : {:>6.0f} m²".format(BUILD_SIZE))
print()
print("  {:30s} {:>14} {:>14}".format("Option", "Year 1 Cost", "Year 3 Total"))
print("  " + "-"*60)
print("  {:30s} {:>14,.0f} {:>14,.0f}".format(
    "A. Short-term 3PL rent", short_total_1yr, short_total_3yr))
print("  {:30s} {:>14,.0f} {:>14,.0f}".format(
    "B. Long-term lease", long_total_1yr, long_total_3yr))
print("  {:30s} {:>14,.0f} {:>14,.0f}".format(
    "C. Build own warehouse", build_capex + build_opex_1yr, build_total_3yr))
print()
print("  Break-even build vs long-term lease: {:.1f} years".format(break_even_vs_long))

# ── Recommendation ────────────────────────────────────────────────
print()
print("[4] RECOMMENDATION")
print("="*72)

rec = """
  CONCLUSION: HYBRID STRATEGY — Long-term Lease (BASE) + 3PL Flex (PEAK)
  ─────────────────────────────────────────────────────────────────────────

  1. DO NOT BUILD (yet)
     - Peak storage only 17,400 units = {:.0f} m²  → too small to justify CapEx
     - Break-even vs long-term lease = {:.1f} years → marginal at scale
     - Business is growing; storage needs will change → lock-in risk

  2. RECOMMENDED: Long-term Lease (BASE) + 3PL overflow (PEAK)
     ┌─────────────────────────────────────────────────────────┐
     │  BASE layer : Lease {:.0f} m² at {:.0f} USD/m²/mo       │
     │               = {:.0f} USD/year (fixed)                  │
     │  PEAK layer : 3PL overflow for ~{} weeks/year           │
     │               = ~{:.0f} USD/year (variable)               │
     │  TOTAL      : ~{:.0f} USD/year                           │
     └─────────────────────────────────────────────────────────┘

  3. CONTRACT TERMS TO NEGOTIATE:
     - 2-year base lease with 1-year renewal option
     - Include "burst capacity" clause: +30% space at 6 months notice
     - Co-locate with fixture 456CD expansion → same industrial zone

  4. IF revenue grows >30% → revisit build decision in Year 3

""".format(
    peak_m2,
    break_even_vs_long,
    avg_m2 * 1.1, LONG_COST_PER_M2_MO,
    LONG_COST_PER_M2_MO * avg_m2 * 1.1 * 12,
    len(high_periods),
    SHORT_COST_PER_M2_MO * (SHORT_SIZE - avg_m2 * 1.1) * active_months,
    LONG_COST_PER_M2_MO * avg_m2 * 1.1 * 12 + SHORT_COST_PER_M2_MO * (SHORT_SIZE - avg_m2 * 1.1) * active_months,
)

print(rec)
print("="*72)

if __name__ == "__main__":
    pass
