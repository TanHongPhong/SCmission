"""
Build scenario logistics solver input + Monthly Cost Report
============================================================
1. Creates solver_scenario.xlsx from BOM_demand_Scenario (SKU A & B only)
2. Runs the MILP transport solver
3. Outputs monthly_logistics_cost.xlsx combining scenario vs baseline
"""

import math
import pandas as pd
from pathlib import Path
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime

BASE        = Path(__file__).parent
MPS_MRP     = BASE / "SCM_round2.1_scenario_MPS_MRP.xlsx"
ORIG_SOLVER = BASE / "solver.xlsx" if (BASE / "solver.xlsx").exists() else None
OUTPUT_XLSX = BASE / "transport_output_scenario_v2.xlsx"
MONTHLY_OUT = BASE / "monthly_logistics_cost.xlsx"
SCENARIO_SOLVER = BASE / "solver_scenario.xlsx"

# ── Styles ──────────────────────────────────────────────────
NAVY   = PatternFill("solid", fgColor="1F4E78")
ORANGE = PatternFill("solid", fgColor="E65100")
TEAL   = PatternFill("solid", fgColor="00897B")
GREEN  = PatternFill("solid", fgColor="E2F0D9")
YELLOW = PatternFill("solid", fgColor="FFF2CC")
RED    = PatternFill("solid", fgColor="FDE9D9")
WH_B   = Font(color="FFFFFF", bold=True)
BLD    = Font(bold=True)
THIN   = Side(style="thin", color="BFBFBF")
BDR    = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def styled_hdr(ws, row, headers, fill=ORANGE):
    for c, h in enumerate(headers, 1):
        cl = ws.cell(row, c, h)
        cl.fill, cl.font = fill, WH_B
        cl.alignment = Alignment(horizontal="center", wrap_text=True)
        cl.border = BDR


def auto_width(ws, max_w=22):
    for col in ws.columns:
        letter = get_column_letter(col[0].column)
        w = max((len(str(c.value or "")) for c in col), default=0)
        ws.column_dimensions[letter].width = min(max(w + 3, 10), max_w)


def build_scenario_solver_xlsx():
    """
    Build solver_scenario.xlsx that transport_exact_global_milp.py can read.
    Structure: one sheet with SKU master + Lane params + Decision table (A & B demand).
    """
    print("[1/4] Reading BOM_demand_Scenario from MPS_MRP file ...")

    # Read BOM_demand_Scenario – weekly production by SKU (A, B)
    bom_df = pd.read_excel(MPS_MRP, sheet_name="BOM_demand_Scenario", header=None)
    # Row 0 = header
    header = [str(bom_df.iloc[0, c]) for c in range(bom_df.shape[1])]
    week_col   = header.index("Week")
    a_col      = header.index("A")
    b_col      = header.index("B")

    from datetime import timedelta
    SHIP_MIN = pd.Timestamp("2026-04-27")   # first valid ship date
    weekly_rows = []
    for r in range(1, bom_df.shape[0]):
        wk = bom_df.iloc[r, week_col]
        if pd.isna(wk): break
        # Shipment date = production week + 7 days
        wk_ship = pd.Timestamp(wk) + timedelta(days=7)
        if wk_ship < SHIP_MIN:
            continue   # skip — too early to ship
        a_qty = int(float(bom_df.iloc[r, a_col] or 0))
        b_qty = int(float(bom_df.iloc[r, b_col] or 0))
        weekly_rows.append({"Week": wk_ship, "A": a_qty, "B": b_qty})

    print(f"    Found {len(weekly_rows)} weekly rows")
    print("    Ship date range: {} -> {} (+7d offset)".format(
        weekly_rows[0]['Week'].date(), weekly_rows[-1]['Week'].date()))



    # Read FGs info (SKU master) from MPS_MRP
    fgs_df = pd.read_excel(MPS_MRP, sheet_name="FGs & Log information ", header=None)
    # Find header row (row 1)
    sku_master = []
    for r in range(2, 8):
        item = str(fgs_df.iloc[r, 1]).strip()
        if item in ("A", "B"):  # Only A and B for this scenario
            desc      = str(fgs_df.iloc[r, 2])
            pack_size = int(float(fgs_df.iloc[r, 3]))
            cbm100    = float(fgs_df.iloc[r, 4])
            price     = float(fgs_df.iloc[r, 5])
            market    = str(fgs_df.iloc[r, 6]).strip()
            sku_master.append({
                "Item name": item, "Des": desc,
                "Packing size (pcs/ carton)": pack_size,
                " CBM (100 cartons) ": cbm100,
                "Ex. Work price (Usd/pcs)": price,
                "Market ": market
            })

    print(f"    SKU master: {[s['Item name'] for s in sku_master]}")

    # Lane parameters (from original file or hardcoded from FGs sheet)
    # US: 40ft=5200, 20ft=3000, LCL=200/cbm  |  UK: 40ft=4200, 20ft=2500, LCL=70/cbm (but A&B both US)
    lane_params = [
        {"Market": "US", "Mode": "40",  "Cost": 5200, "Cap_CBM": 65.0},
        {"Market": "US", "Mode": "20",  "Cost": 3000, "Cap_CBM": 28.0},
        {"Market": "US", "Mode": "LCL", "Cost": 200,  "Cap_CBM": 99999},
        {"Market": "UK", "Mode": "40",  "Cost": 4200, "Cap_CBM": 65.0},
        {"Market": "UK", "Mode": "20",  "Cost": 2500, "Cap_CBM": 28.0},
        {"Market": "UK", "Mode": "LCL", "Cost": 70,   "Cap_CBM": 99999},
        {"Market": "AU", "Mode": "40",  "Cost": 2000, "Cap_CBM": 65.0},
        {"Market": "AU", "Mode": "20",  "Cost": 1100, "Cap_CBM": 28.0},
        {"Market": "AU", "Mode": "LCL", "Cost": 35,   "Cap_CBM": 99999},
    ]

    # Write solver_scenario.xlsx
    wb = Workbook()
    ws = wb.active
    ws.title = "Scenario_Solver"

    r = 1
    # ── Section 1: SKU Master ───────────────────────────────
    ws.cell(r, 1, "SKU Master").font = BLD
    r += 1
    master_hdrs = list(sku_master[0].keys())
    for c, h in enumerate(master_hdrs, 1):
        ws.cell(r, c, h).font = BLD
    r += 1
    for sm in sku_master:
        for c, k in enumerate(master_hdrs, 1):
            ws.cell(r, c, sm[k])
        r += 1

    r += 1  # blank
    # ── Section 2: Lane Parameters ──────────────────────────
    ws.cell(r, 1, "Lane Parameters").font = BLD
    r += 1
    lane_hdrs = ["Market", "Mode", "Cost", "Cap_CBM"]
    for c, h in enumerate(lane_hdrs, 1):
        ws.cell(r, c, h).font = BLD
    r += 1
    for lp in lane_params:
        for c, k in enumerate(lane_hdrs, 1):
            ws.cell(r, c, lp[k])
        r += 1

    r += 1  # blank
    # ── Section 3: Decision Table (weekly demand) ───────────
    ws.cell(r, 1, "Decision Table (Weekly Demand by SKU)").font = BLD
    r += 1
    ws.cell(r, 1, "Week"); ws.cell(r, 1).font = BLD
    ws.cell(r, 2, "A");    ws.cell(r, 2).font = BLD
    ws.cell(r, 3, "B");    ws.cell(r, 3).font = BLD
    r += 1
    for wr in weekly_rows:
        ws.cell(r, 1, wr["Week"])
        ws.cell(r, 1).number_format = "YYYY-MM-DD"
        ws.cell(r, 2, wr["A"])
        ws.cell(r, 3, wr["B"])
        r += 1

    wb.save(SCENARIO_SOLVER)
    print(f"    Saved: {SCENARIO_SOLVER.name}")
    return weekly_rows, sku_master, lane_params


def run_logistics(weekly_rows, sku_master, lane_params):
    """
    Run the same MILP logic from transport_exact_global_milp.py
    directly here for A & B.
    """
    print("[2/4] Running MILP transport solver for A & B ...")

    # Import the solver
    import sys
    sys.path.insert(0, str(BASE))
    from transport_exact_global_milp import solve_exact_market_week

    items = [s["Item name"] for s in sku_master]
    master_df = pd.DataFrame([{
        "Item": s["Item name"],
        "PackSize": s["Packing size (pcs/ carton)"],
        "CBMPerBox": s[" CBM (100 cartons) "] / 100.0,
        "Market": s["Market "].strip().upper().replace("AUSTRALIA", "AU"),
    } for s in sku_master])

    lanes_df = pd.DataFrame(lane_params)
    lanes_df["Mode"] = lanes_df["Mode"].astype(str)

    weekly_market_rows = []
    detail_rows_all = []

    for wi, wr in enumerate(weekly_rows, 1):
        for market in ["US", "UK", "AU"]:
            master_mkt = master_df[master_df["Market"] == market].copy()
            if master_mkt.empty: continue
            lanes_mkt = lanes_df[lanes_df["Market"] == market].copy()
            demand = {item: float(wr.get(item, 0)) for item in master_mkt["Item"].tolist()}

            weekly, detail_df = solve_exact_market_week(
                week_idx=wi, week_raw=wr["Week"],
                market=market,
                item_demands_units=demand,
                master_mkt=master_mkt,
                lane_mkt=lanes_mkt,
            )
            weekly_market_rows.append(weekly)
            if not detail_df.empty:
                detail_rows_all.append(detail_df)

    weekly_all = pd.DataFrame(weekly_market_rows)
    detail_all = pd.concat(detail_rows_all, ignore_index=True) if detail_rows_all else pd.DataFrame()

    print(f"    Solved {len(weekly_rows)} weeks × {len(set(master_df['Market']))} markets")
    return weekly_all, detail_all


def build_monthly_cost_report(weekly_all):
    """Build monthly logistics cost report with scenario vs baseline comparison."""
    print("[3/4] Building monthly logistics cost report ...")

    wb = Workbook()
    ws = wb.active
    ws.title = "Monthly_Logistics_Cost"

    # Add month column
    def get_month(x):
        if hasattr(x, "month"): return x.strftime("%Y-%m")
        try: return pd.to_datetime(x).strftime("%Y-%m")
        except: return "Unknown"

    weekly_all["Month"] = weekly_all["Week_Date"].apply(get_month)
    weekly_all["Week_Date_dt"] = pd.to_datetime(weekly_all["Week_Date"], errors="coerce")

    # Monthly summary
    monthly = (
        weekly_all.groupby(["Month", "Market"], as_index=False)
        .agg(
            Weeks=("Week", "count"),
            Total_CBM=("Total_Demand_CBM", "sum"),
            n40=("n40", "sum"),
            n20=("n20", "sum"),
            LCL_CBM=("LCL_CBM", "sum"),
            Cost_40=("Cost_40", "sum"),
            Cost_20=("Cost_20", "sum"),
            Cost_LCL=("Cost_LCL", "sum"),
            Total_Cost=("Cost", "sum"),
        )
    ).sort_values(["Month", "Market"])

    # Sheet 1: Monthly by Market
    styled_hdr(ws, 1, [
        "Month", "Market", "Weeks", "Total CBM",
        "40ft Ctnr", "20ft Ctnr", "LCL CBM",
        "Cost 40ft", "Cost 20ft", "Cost LCL", "Total Cost"
    ], TEAL)

    for ri, row in enumerate(monthly.itertuples(index=False), 2):
        vals = [row.Month, row.Market, row.Weeks, round(row.Total_CBM, 2),
                row.n40, row.n20, round(row.LCL_CBM, 2),
                round(row.Cost_40, 2), round(row.Cost_20, 2),
                round(row.Cost_LCL, 2), round(row.Total_Cost, 2)]
        for ci, v in enumerate(vals, 1):
            cl = ws.cell(ri, ci, v)
            cl.border = BDR
            if ci >= 8:  # cost columns
                cl.number_format = "$#,##0.00"
            elif ci == 4 or ci == 7:
                cl.number_format = "0.00"

    # Grand total row
    ri_total = len(monthly) + 2
    ws.cell(ri_total, 1, "GRAND TOTAL").font = BLD
    ws.cell(ri_total, 1).fill = ORANGE
    ws.cell(ri_total, 1).font = WH_B
    for ci, val in enumerate([
        "", "", monthly["Weeks"].sum(), round(monthly["Total_CBM"].sum(), 2),
        monthly["n40"].sum(), monthly["n20"].sum(), round(monthly["LCL_CBM"].sum(), 2),
        round(monthly["Cost_40"].sum(), 2), round(monthly["Cost_20"].sum(), 2),
        round(monthly["Cost_LCL"].sum(), 2), round(monthly["Total_Cost"].sum(), 2)
    ], 2):
        cl = ws.cell(ri_total, ci, val)
        cl.font = BLD
        cl.border = BDR
        if ci >= 8:
            cl.number_format = "$#,##0.00"

    auto_width(ws)
    ws.freeze_panes = "A2"

    # Sheet 2: Weekly detail
    ws2 = wb.create_sheet("Weekly_Detail")
    weekly_display = weekly_all[["Month", "Market", "Week", "Week_Date",
                                  "Total_Demand_CBM", "n40", "n20", "LCL_CBM",
                                  "Cost_40", "Cost_20", "Cost_LCL", "Cost"]].copy()
    styled_hdr(ws2, 1, weekly_display.columns.tolist(), NAVY)
    for ri, row in enumerate(weekly_display.itertuples(index=False), 2):
        for ci, v in enumerate(row, 1):
            cl = ws2.cell(ri, ci, v)
            cl.border = BDR
            hdr = weekly_display.columns[ci - 1]
            if "Cost" in hdr:
                cl.number_format = "$#,##0.00"
            elif "CBM" in hdr:
                cl.number_format = "0.0000"
    auto_width(ws2)
    ws2.freeze_panes = "A2"

    # Sheet 3: Summary by Market
    ws3 = wb.create_sheet("Summary_by_Market")
    mkt_summary = monthly.groupby("Market", as_index=False).agg(
        Total_CBM=("Total_CBM", "sum"),
        Total_n40=("n40", "sum"),
        Total_n20=("n20", "sum"),
        Total_LCL_CBM=("LCL_CBM", "sum"),
        Total_Cost_40=("Cost_40", "sum"),
        Total_Cost_20=("Cost_20", "sum"),
        Total_Cost_LCL=("Cost_LCL", "sum"),
        Total_Cost=("Total_Cost", "sum"),
    )
    styled_hdr(ws3, 1, mkt_summary.columns.tolist(), ORANGE)
    for ri, row in enumerate(mkt_summary.itertuples(index=False), 2):
        for ci, v in enumerate(row, 1):
            cl = ws3.cell(ri, ci, v)
            cl.border = BDR
            if "Cost" in mkt_summary.columns[ci - 1]:
                cl.number_format = "$#,##0.00"
    auto_width(ws3)

    # Sheet 4: Scenario Cost Breakdown (include PCBA Air freight)
    ws4 = wb.create_sheet("Scenario_Cost_Breakdown")
    styled_hdr(ws4, 1, ["Cost Category", "Amount (USD)", "Notes"], TEAL)
    ocean_cost = round(monthly["Total_Cost"].sum(), 2)
    pcba_air   = 70_040.00      # from scenario analysis
    total_scen = ocean_cost + pcba_air

    cost_breakdown = [
        ["--- Ocean Freight (FG Shipment) ---", "", ""],
        ["  40ft Container cost", round(monthly["Cost_40"].sum(), 2), "A+B US market"],
        ["  20ft Container cost", round(monthly["Cost_20"].sum(), 2), "A+B US market"],
        ["  LCL cost", round(monthly["LCL_CBM"].sum() * 0, 2), "If any overflow"],
        ["  Total Ocean Freight", ocean_cost, "Sum of all containers"],
        ["", "", ""],
        ["--- PCBA Inbound Air Freight (Scenario) ---", "", ""],
        ["  PCBA units shipped by air", 35_020, "= 60% recovery"],
        ["  Air premium per PCBA", 2.0, "$2.00 / unit"],
        ["  Total PCBA Air Freight", pcba_air, "= 35,020 × $2"],
        ["", "", ""],
        ["--- Total Scenario Logistics Cost ---", "", ""],
        ["  Ocean Freight", ocean_cost, "FG delivery to customers"],
        ["  PCBA Air Freight", pcba_air, "Component inbound"],
        ["  GRAND TOTAL", total_scen, "Full scenario logistics cost"],
        ["", "", ""],
        ["--- Baseline Comparison ---", "", ""],
        ["  Baseline ocean freight", ocean_cost, "Same (no change in shipping lanes)"],
        ["  Baseline PCBA air", 0, "No air freight in baseline"],
        ["  Baseline total", ocean_cost, ""],
        ["  ADDITIONAL cost vs baseline", pcba_air, "= Scenario impact"],
        ["  Additional cost %", "{:.1f}%".format(100 * pcba_air / ocean_cost) if ocean_cost > 0 else "N/A", ""],
    ]

    for ri, row in enumerate(cost_breakdown, 2):
        bold = row[0].startswith("---") or row[0].strip().startswith("GRAND") or row[0].strip().startswith("ADDITIONAL")
        for ci, v in enumerate(row, 1):
            cl = ws4.cell(ri, ci, v)
            cl.border = BDR
            if bold:
                cl.font = BLD
                cl.fill = YELLOW
            if ci == 2 and isinstance(v, (int, float)) and v != 0:
                cl.number_format = "$#,##0.00"
    auto_width(ws4, 45)

    wb.save(MONTHLY_OUT)
    print(f"    Saved: {MONTHLY_OUT.name}")
    return monthly


def main():
    print("=" * 60)
    print("  SCENARIO LOGISTICS ANALYSIS")
    print("  Source: BOM_demand_Scenario (A & B)")
    print("=" * 60)

    # Build solver input
    weekly_rows, sku_master, lane_params = build_scenario_solver_xlsx()

    # Run MILP
    weekly_all, detail_all = run_logistics(weekly_rows, sku_master, lane_params)

    # Monthly cost report
    monthly = build_monthly_cost_report(weekly_all)

    # Also save full transport output
    print("[4/4] Saving transport output ...")
    wb2 = Workbook()
    wb2.remove(wb2.active)
    from transport_exact_global_milp import write_df_sheet
    for name, df in [("Weekly_All", weekly_all), ("All_Detail", detail_all)]:
        write_df_sheet(wb2, name, df)

    # Summary
    summary = (weekly_all.groupby("Market", as_index=False)
               .agg(Weeks=("Week","count"),
                    Total_CBM=("Total_Demand_CBM","sum"),
                    n40=("n40","sum"), n20=("n20","sum"),
                    LCL_CBM=("LCL_CBM","sum"),
                    Total_Cost=("Cost","sum")))
    write_df_sheet(wb2, "Summary_Market", summary)
    wb2.save(OUTPUT_XLSX)

    print()
    print("=" * 60)
    print("  RESULTS")
    print("=" * 60)
    print()
    print("  Market summary:")
    for _, row in summary.iterrows():
        print("    {:>2}: CBM={:>8.1f}  40ft={:>3}  20ft={:>3}  LCL={:>7.2f}  Cost=${:>10,.2f}".format(
            row["Market"], row["Total_CBM"],
            int(row["n40"]), int(row["n20"]),
            row["LCL_CBM"], row["Total_Cost"]))
    print()
    print("  Total ocean freight : ${:>10,.2f}".format(summary["Total_Cost"].sum()))
    print("  PCBA air freight    : $    70,040.00")
    print("  GRAND TOTAL         : ${:>10,.2f}".format(summary["Total_Cost"].sum() + 70040))
    print()
    print("  Output files:")
    print("    -", OUTPUT_XLSX.name)
    print("    -", MONTHLY_OUT.name)
    print("    -", SCENARIO_SOLVER.name)
    print("=" * 60)


if __name__ == "__main__":
    main()
