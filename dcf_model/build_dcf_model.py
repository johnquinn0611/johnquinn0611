"""Build the DCF valuation workbook (Assumptions, Segment_Revenue, Working_Capital, Capex, WACC, DCF,
Scenarios, Sensitivity tabs).

Run:  python dcf_model/build_dcf_model.py   ->   dcf_model/DCF_Model.xlsx
Requires LibreOffice (soffice) on PATH: the workbook is recalculated headless once per scenario to
fill the static Scenario Summary table, and once more for the final file.

Layout conventions shared by every time-series tab (keep new tabs on these):
  * Column B = row labels
  * Column C = base year (last actual fiscal year, period t = 0)
  * Columns D:H = projection years 1-5 (periods t = 1..5)
"""
import shutil
import subprocess
import sys
import tempfile
from datetime import date
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation

OUT = Path(__file__).with_name("DCF_Model.xlsx")

FONT = "Arial"
BLUE, BLACK, GREEN, WHITE = "0000FF", "000000", "008000", "FFFFFF"
INPUT_FILL = PatternFill("solid", fgColor="FFFF00")
SECTION_FILL = PatternFill("solid", fgColor="1F3864")
HEADER_FILL = PatternFill("solid", fgColor="D9E1F2")
TOP_BORDER = Border(top=Side(style="thin", color="000000"))
CENTER_FILL = PatternFill("solid", fgColor="FCE4D6")

FMT_USD = '$#,##0.0_);($#,##0.0);"-"_)'
FMT_PCT = '0.0%_);(0.0%);"-"_)'
FMT_YEAR = "0"
FMT_DATE = "mmm d, yyyy"
FMT_DAYS = '0.0_);(0.0);"-"_)'
FMT_INTENSITY = '$0.00_);($0.00);"-"_)'
FMT_MULT = '0.0"x"'
FMT_PRICE = '$#,##0.00_);($#,##0.00);"-"_)'
FMT_BETA = "0.00"
FMT_FACTOR = "0.0000"
FMT_PERIOD = "0.0"
FMT_SHARES = "#,##0.0"

N_SEGMENTS = 8
YEAR_COLS = "CDEFGH"  # C = base year, D:H = projection years 1-5
PROJ_COLS = YEAR_COLS[1:]

# Example inputs so the model shows numbers out of the box -- replace with real data.
EXAMPLE_SEGMENTS = [
    # name, current revenue ($mm), growth / yr, operating margin
    ("Products", 1200.0, 0.06, 0.22),
    ("Services", 650.0, 0.10, 0.30),
    ("Licensing & Other", 150.0, 0.03, 0.08),
]

# Example working capital / capex drivers (base year + 5 projected years) -- replace.
EXAMPLE_COGS_PCT = [0.55] * 6
EXAMPLE_DSO = [45, 45, 44, 44, 43, 43]
EXAMPLE_DIO = [60, 60, 58, 57, 56, 55]
EXAMPLE_DPO = [40] * 6
EXAMPLE_BALANCES = (245.0, 180.0, 120.0)  # base-year AR, inventory, AP ($mm)
EXAMPLE_MAINT_CAPEX = 0.03
EXAMPLE_CAPEX_INTENSITY = [0.25, 0.15, 0.10]  # $ capex per $1 of incremental revenue

# Example WACC / DCF inputs -- replace.
EXAMPLE_RF, EXAMPLE_ERP, EXAMPLE_BETA, EXAMPLE_KD = 0.0425, 0.055, 1.10, 0.06
EXAMPLE_MCAP, EXAMPLE_DEBT = 6000.0, 1000.0  # $mm
EXAMPLE_DA_PCT, EXAMPLE_DCF_TGR, EXAMPLE_EXIT_MULT = 0.035, 0.025, 12.0
EXAMPLE_NET_DEBT, EXAMPLE_SHARES, EXAMPLE_MIDYEAR = 800.0, 150.0, 1  # $mm, mm shares, 1 = mid-year

# Scenarios: Base = the example inputs above; Bear / Bull = Base shifted by these placeholder amounts.
SCENARIOS = ("Bear", "Base", "Bull")
SCENARIO_SHIFTS = {  # driver: (bear shift, bull shift)
    "growth": (-0.02, 0.02), "margin": (-0.02, 0.02),
    "dso": (5, -5), "dio": (5, -5), "dpo": (-5, 5),
    "maint": (0.005, -0.005), "beta": (0.10, -0.10), "tgr": (-0.005, 0.005), "exit": (-2.0, 2.0),
}
SENS_STEPS = (0.005, 0.005, 1.0)  # WACC, terminal growth, exit multiple

# ---- Assumptions tab cell map --------------------------------------------
A = "Assumptions"
A_COMPANY, A_TICKER, A_UNITS, A_BASE_YEAR, A_VAL_DATE = 7, 8, 9, 10, 11
A_TAX, A_WACC, A_TGR, A_TGR_CHECK = 14, 15, 16, 17
A_SEG_HDR = 20
A_SEG_FIRST = A_SEG_HDR + 1
A_SEG_LAST = A_SEG_FIRST + N_SEGMENTS - 1
A_SEG_TOTAL = A_SEG_LAST + 1
A_NOTES = A_SEG_TOTAL + 2  # notes occupy A_NOTES+1 .. A_NOTES+5

# Sections 4-5 (working capital & capex drivers), added below the Model Notes block
A_WC_HDR = A_NOTES + 7
A_WC_FY = A_WC_HDR + 1
A_COGS_PCT, A_DSO, A_DIO, A_DPO, A_DAYS = range(A_WC_FY + 1, A_WC_FY + 6)
A_BAL_HDR = A_DAYS + 1
A_AR0, A_INV0, A_AP0 = range(A_BAL_HDR + 1, A_BAL_HDR + 4)
A_CPX_HDR = A_AP0 + 2
A_MAINT = A_CPX_HDR + 1
A_INT_HDR = A_MAINT + 1
A_INT_FIRST = A_INT_HDR + 1
A_INT_LAST = A_INT_FIRST + N_SEGMENTS - 1
A_DRV_NOTES = A_INT_LAST + 2  # driver notes occupy A_DRV_NOTES+1 .. A_DRV_NOTES+4

# Sections 6-7 (WACC & DCF inputs), added below the Driver Notes block
A_WACC_HDR = A_DRV_NOTES + 6
A_RF, A_ERP, A_BETA, A_KD, A_MCAP, A_DEBT = range(A_WACC_HDR + 1, A_WACC_HDR + 7)
A_DCF_HDR = A_DEBT + 2
A_DA_PCT, A_DCF_TGR, A_EXIT_MULT, A_NET_DEBT, A_SHARES, A_MIDYEAR = range(A_DCF_HDR + 1, A_DCF_HDR + 7)

# ---- Segment_Revenue tab row map ------------------------------------------
S = "Segment_Revenue"
S_FY, S_T = 4, 5
S_REV_HDR = 7
S_REV_FIRST = S_REV_HDR + 1
S_REV_TOTAL = S_REV_FIRST + N_SEGMENTS
S_REV_GROWTH = S_REV_TOTAL + 1
S_OI_HDR = S_REV_GROWTH + 2
S_OI_FIRST = S_OI_HDR + 1
S_OI_TOTAL = S_OI_FIRST + N_SEGMENTS
S_OI_MARGIN = S_OI_TOTAL + 1
S_SUM_HDR = S_OI_MARGIN + 2
S_SUM_REV, S_SUM_GROWTH, S_SUM_EBIT, S_SUM_MARGIN = range(S_SUM_HDR + 1, S_SUM_HDR + 5)
S_CHECK = S_SUM_MARGIN + 2

# ---- Working_Capital tab row map ------------------------------------------
W = "Working_Capital"
W_DRV_HDR = 7
W_REV, W_COGS_PCT, W_COGS, W_DSO, W_DIO, W_DPO = range(W_DRV_HDR + 1, W_DRV_HDR + 7)
W_BAL_HDR = W_DPO + 2
W_AR, W_INV, W_AP, W_NWC, W_NWC_PCT = range(W_BAL_HDR + 1, W_BAL_HDR + 6)
W_CHG_HDR = W_NWC_PCT + 2
W_CHG = W_CHG_HDR + 1

# ---- Capex tab row map ----------------------------------------------------
X = "Capex"
X_MAINT_HDR = 7
X_REV, X_MAINT_PCT, X_MAINT = range(X_MAINT_HDR + 1, X_MAINT_HDR + 4)
X_INC_HDR = X_MAINT + 2
X_INC_FIRST = X_INC_HDR + 1
X_INC_TOTAL = X_INC_FIRST + N_SEGMENTS
X_GC_HDR = X_INC_TOTAL + 2
X_GC_FIRST = X_GC_HDR + 1
X_GC_TOTAL = X_GC_FIRST + N_SEGMENTS
X_SUM_HDR = X_GC_TOTAL + 2
X_SUM_MAINT, X_SUM_GROWTH, X_SUM_TOTAL, X_SUM_PCT = range(X_SUM_HDR + 1, X_SUM_HDR + 5)

# ---- WACC tab row map (label in B, value in C) ----------------------------
K = "WACC"
K_IN_HDR = 4
K_RF, K_ERP, K_BETA, K_KD, K_TAX, K_MCAP, K_DEBT = range(K_IN_HDR + 1, K_IN_HDR + 8)
K_COC_HDR = K_DEBT + 2
K_KE, K_KD_AT = K_COC_HDR + 1, K_COC_HDR + 2
K_CAP_HDR = K_KD_AT + 2
K_CAP, K_WE, K_WD = range(K_CAP_HDR + 1, K_CAP_HDR + 4)
K_WACC_HDR = K_WD + 2
K_WACC = K_WACC_HDR + 1

# ---- DCF tab row map (projections in D:H; valuation in C = perpetuity, D = exit multiple)
D = "DCF"
D_FCF_HDR = 7
D_EBIT, D_TAX, D_DA, D_CAPEX, D_NWC, D_UFCF = range(D_FCF_HDR + 1, D_FCF_HDR + 7)
D_DISC_HDR = D_UFCF + 2
D_RATE, D_TOGGLE, D_PERIOD, D_FACTOR, D_PV, D_SUM_PV = range(D_DISC_HDR + 1, D_DISC_HDR + 7)
D_VAL_HDR = D_SUM_PV + 2
(D_METHOD, D_TV_FCF, D_TV_G, D_TV_EBITDA, D_TV_MULT, D_TV, D_TV_PERIOD, D_TV_FACTOR, D_TV_PV,
 D_EV_PV_FCF, D_EV, D_NET_DEBT, D_EQUITY, D_SHARES, D_PRICE) = range(D_VAL_HDR + 1, D_VAL_HDR + 16)
D_CHK_HDR = D_PRICE + 2
D_TV_PCT, D_IMPL_MULT, D_IMPL_G = range(D_CHK_HDR + 1, D_CHK_HDR + 4)

# ---- Scenario selector on Assumptions (empty row 3) -------------------------
A_SCEN = 3  # B = label, C = Bear/Base/Bull dropdown, D = label, E = 1/2/3 helper

# ---- Scenarios tab row map (C = Bear, D = Base, E = Bull, F = Live) ---------
SC = "Scenarios"
SC_HDR = 4
SC_GROWTH_HDR = SC_HDR + 2
SC_GROWTH_FIRST = SC_GROWTH_HDR + 1
SC_MARGIN_HDR = SC_GROWTH_FIRST + N_SEGMENTS + 1
SC_MARGIN_FIRST = SC_MARGIN_HDR + 1
SC_DAYS_HDR = SC_MARGIN_FIRST + N_SEGMENTS + 1
SC_DSO_FIRST = SC_DAYS_HDR + 1
SC_DIO_FIRST = SC_DSO_FIRST + len(PROJ_COLS)
SC_DPO_FIRST = SC_DIO_FIRST + len(PROJ_COLS)
SC_OTHER_HDR = SC_DPO_FIRST + len(PROJ_COLS) + 1
SC_MAINT, SC_BETA, SC_TGR, SC_EXIT = range(SC_OTHER_HDR + 1, SC_OTHER_HDR + 5)
SC_SUM_HDR = SC_EXIT + 2
SC_SUM_COLS = SC_SUM_HDR + 1
SC_EV_PGM, SC_EV_EXIT, SC_PX_PGM, SC_PX_EXIT = range(SC_SUM_COLS + 1, SC_SUM_COLS + 5)

# ---- Sensitivity tab row map (5 x 5 grids in C:G, WACC axis in B) -----------
SN = "Sensitivity"
SN_STEP_HDR = 4
SN_STEP_WACC, SN_STEP_TGR, SN_STEP_MULT = range(SN_STEP_HDR + 1, SN_STEP_HDR + 4)
SN_PGM_HDR = SN_STEP_MULT + 2
SN_PGM_AXIS = SN_PGM_HDR + 1
SN_PGM_FIRST = SN_PGM_AXIS + 1
SN_EXIT_HDR = SN_PGM_FIRST + 6
SN_EXIT_AXIS = SN_EXIT_HDR + 1
SN_EXIT_FIRST = SN_EXIT_AXIS + 1
GRID_COLS = "CDEFG"
GRID_MID = 2  # index of the center (live case) row / column


def put(ws, ref, value, color=BLACK, fmt=None, bold=False, italic=False, fill=None, size=10):
    c = ws[ref]
    c.value = value
    c.font = Font(name=FONT, color=color, bold=bold, italic=italic, size=size)
    if fmt:
        c.number_format = fmt
    if fill:
        c.fill = fill
    return c


def section(ws, row, text, first="B", last="F"):
    for col in range(ord(first), ord(last) + 1):
        ws[f"{chr(col)}{row}"].fill = SECTION_FILL
    put(ws, f"{first}{row}", text, color=WHITE, bold=True)


def total_style(ws, row, first="B", last="H"):
    for col in range(ord(first), ord(last) + 1):
        c = ws[f"{chr(col)}{row}"]
        c.border = TOP_BORDER
        c.font = Font(name=FONT, color=c.font.color, bold=True, size=10)


def add_name(wb, name, ref):
    wb.defined_names[name] = DefinedName(name, attr_text=ref)


def pct_validation(ws, rng, lo, hi):
    dv = DataValidation(
        type="decimal", operator="between", formula1=str(lo), formula2=str(hi),
        allow_blank=True, showErrorMessage=True, errorTitle="Invalid percentage",
        error=f"Enter a percentage between {lo:.0%} and {hi:.0%} (e.g. 25% or 0.25).",
    )
    ws.add_data_validation(dv)
    dv.add(rng)


def num_validation(ws, rng, lo, hi, what):
    dv = DataValidation(
        type="decimal", operator="between", formula1=str(lo), formula2=str(hi),
        allow_blank=True, showErrorMessage=True, errorTitle=f"Invalid {what}",
        error=f"Enter a {what} between {lo} and {hi}.",
    )
    ws.add_data_validation(dv)
    dv.add(rng)


def seg_label(a_row):
    return f'=IF({A}!$B{a_row}="","",{A}!$B{a_row})'


def build_assumptions(wb):
    ws = wb.active
    ws.title = A
    ws.sheet_view.showGridLines = False
    for col, width in zip("ABCDEF", (2, 40, 16, 18, 18, 50)):
        ws.column_dimensions[col].width = width

    put(ws, "B1", "DCF Valuation Model - Assumptions & Config", bold=True, size=14)
    put(ws, "B2", "All currency figures in $ millions unless noted. "
                  "Edit only blue cells with yellow fill; everything else is formula-driven.", italic=True)
    put(ws, "B4", "Legend:", bold=True)
    put(ws, "C4", "Input", color=BLUE, fill=INPUT_FILL)
    put(ws, "D4", "Formula", color=BLACK)
    put(ws, "E4", "Link from other tab", color=GREEN)

    # 1. General / config
    section(ws, 6, "1. General / Config")
    rows = [
        (A_COMPANY, "Company Name", "Example Co.", None, "Example placeholder - replace"),
        (A_TICKER, "Ticker", "EXCO", None, "Example placeholder - replace"),
        (A_UNITS, "Units", "$ in millions", None, "Label only; shown in tab headers"),
        (A_BASE_YEAR, "Last Actual Fiscal Year (Base Year)", 2025, FMT_YEAR,
         "Projections run Base Year +1 through +5"),
        (A_VAL_DATE, "Valuation Date", date(2026, 10, 13), FMT_DATE, "Example placeholder - replace"),
    ]
    for r, label, val, fmt, note in rows:
        put(ws, f"B{r}", label)
        put(ws, f"C{r}", val, color=BLUE, fmt=fmt, fill=INPUT_FILL)
        put(ws, f"D{r}", note, italic=True)

    # 2. Company-wide valuation inputs
    section(ws, 13, "2. Company-Wide Valuation Inputs")
    rows = [
        (A_TAX, "Tax Rate", 0.25, "Example placeholder - effective/marginal cash tax rate"),
        (A_WACC, "Discount Rate (WACC)", 0.09, "Example placeholder - replace with WACC build"),
        (A_TGR, "Terminal Growth Rate", 0.025, "Example placeholder - long-run nominal growth"),
    ]
    for r, label, val, note in rows:
        put(ws, f"B{r}", label)
        put(ws, f"C{r}", val, color=BLUE, fmt=FMT_PCT, fill=INPUT_FILL)
        put(ws, f"D{r}", note, italic=True)
    put(ws, f"B{A_TGR_CHECK}", "Check: Terminal Growth < WACC", italic=True)
    put(ws, f"C{A_TGR_CHECK}", f'=IF(C{A_TGR}<C{A_WACC},"OK","ERROR")', bold=True)
    pct_validation(ws, f"C{A_TAX}:C{A_WACC}", 0, 1)
    pct_validation(ws, f"C{A_TGR}", -0.1, 0.2)

    # 3. Segment inputs
    section(ws, A_SEG_HDR - 1, f"3. Segment Inputs (up to {N_SEGMENTS} segments)")
    headers = ["Segment Name", "Current Revenue ($mm)", "Projected Growth Rate (% / yr)",
               "Operating Margin (%)", "Notes / Source"]
    for col, text in zip("BCDEF", headers):
        c = put(ws, f"{col}{A_SEG_HDR}", text, bold=True, fill=HEADER_FILL)
        c.alignment = Alignment(wrap_text=True, horizontal="left" if col in "BF" else "center",
                                vertical="center")
    ws.row_dimensions[A_SEG_HDR].height = 30

    for i in range(N_SEGMENTS):
        r = A_SEG_FIRST + i
        name, rev, growth, margin = EXAMPLE_SEGMENTS[i] if i < len(EXAMPLE_SEGMENTS) else (None,) * 4
        put(ws, f"B{r}", name, color=BLUE, fill=INPUT_FILL)
        put(ws, f"C{r}", rev, color=BLUE, fmt=FMT_USD, fill=INPUT_FILL)
        put(ws, f"D{r}", growth, color=BLUE, fmt=FMT_PCT, fill=INPUT_FILL)
        put(ws, f"E{r}", margin, color=BLUE, fmt=FMT_PCT, fill=INPUT_FILL)
        put(ws, f"F{r}", "Example placeholder - replace with company data" if name else None,
            color=BLUE, italic=True)
    pct_validation(ws, f"D{A_SEG_FIRST}:D{A_SEG_LAST}", -1, 1)
    pct_validation(ws, f"E{A_SEG_FIRST}:E{A_SEG_LAST}", -1, 1)

    f, l, t = A_SEG_FIRST, A_SEG_LAST, A_SEG_TOTAL
    put(ws, f"B{t}", "Total / Revenue-Weighted Average")
    put(ws, f"C{t}", f"=SUM(C{f}:C{l})", fmt=FMT_USD)
    put(ws, f"D{t}", f"=IF(C{t}=0,0,SUMPRODUCT(C{f}:C{l},D{f}:D{l})/C{t})", fmt=FMT_PCT)
    put(ws, f"E{t}", f"=IF(C{t}=0,0,SUMPRODUCT(C{f}:C{l},E{f}:E{l})/C{t})", fmt=FMT_PCT)
    total_style(ws, t, last="F")

    # Notes
    put(ws, f"B{A_NOTES}", "Model Notes", bold=True)
    notes = [
        f"To add a segment, fill the next empty slot (rows {A_SEG_FIRST}-{A_SEG_LAST}); "
        f"{S} picks it up automatically. Blank slots calculate as zero.",
        "Growth rate and operating margin are applied uniformly to all 5 projection years; "
        "base-year operating income = current revenue x operating margin.",
        f"Time-series tabs share one column layout: C = base year (actual), D:H = years 1-5. "
        f"Link downstream tabs to the Total Company Summary block on {S}.",
        "Named ranges: CompanyName, Ticker, Units, BaseYear, ValuationDate, TaxRate, WACC, "
        "TerminalGrowth, Seg_Names, Seg_Revenue, Seg_Growth, Seg_Margin, FY_Labels, Period, "
        "Total_Revenue, Total_EBIT.",
        "Add new driver sections (e.g. 4. Working Capital, 5. Capex & D&A) below this notes block.",
    ]
    for i, text in enumerate(notes, start=1):
        put(ws, f"B{A_NOTES + i}", f"{i}. {text}", italic=True)

    for name, row in [("CompanyName", A_COMPANY), ("Ticker", A_TICKER), ("Units", A_UNITS),
                      ("BaseYear", A_BASE_YEAR), ("ValuationDate", A_VAL_DATE),
                      ("TaxRate", A_TAX), ("WACC", A_WACC), ("TerminalGrowth", A_TGR)]:
        add_name(wb, name, f"{A}!$C${row}")
    for name, col in [("Seg_Names", "B"), ("Seg_Revenue", "C"), ("Seg_Growth", "D"), ("Seg_Margin", "E")]:
        add_name(wb, name, f"{A}!${col}${A_SEG_FIRST}:${col}${A_SEG_LAST}")


def build_segment_revenue(wb):
    ws = wb.create_sheet(S)
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 2
    ws.column_dimensions["B"].width = 38
    for col in YEAR_COLS + "I":
        ws.column_dimensions[col].width = 13

    put(ws, "B1", '=CompanyName&" - Segment Revenue & Operating Income Build"', bold=True, size=14)
    put(ws, "B2", '="("&Units&"; all drivers link to the Assumptions tab)"', italic=True)

    # Timeline header
    put(ws, f"B{S_FY}", "Fiscal Year", bold=True)
    put(ws, f"C{S_FY}", '="FY"&BaseYear&"A"', bold=True)
    for col in PROJ_COLS:
        put(ws, f"{col}{S_FY}", f'="FY"&(BaseYear+{col}${S_T})&"E"', bold=True)
    put(ws, f"I{S_FY}", "5-Yr CAGR", bold=True)
    put(ws, f"B{S_T}", "Projection Period (t)", italic=True)
    put(ws, f"C{S_T}", 0, color=BLUE, italic=True)
    for prev, col in zip(YEAR_COLS, PROJ_COLS):
        put(ws, f"{col}{S_T}", f"={prev}{S_T}+1", italic=True)
    for col in "BCDEFGHI":
        ws[f"{col}{S_FY}"].fill = HEADER_FILL
        ws[f"{col}{S_FY}"].alignment = Alignment(horizontal="left" if col == "B" else "right")
        ws[f"{col}{S_T}"].alignment = Alignment(horizontal="left" if col == "B" else "right")

    def cagr(row):
        return f"=IF(C{row}<=0,0,(H{row}/C{row})^(1/$H${S_T})-1)"

    # Revenue by segment
    section(ws, S_REV_HDR, "Revenue by Segment", last="I")
    for i in range(N_SEGMENTS):
        r, a = S_REV_FIRST + i, A_SEG_FIRST + i
        put(ws, f"B{r}", seg_label(a), color=GREEN)
        put(ws, f"C{r}", f"={A}!$C{a}", color=GREEN, fmt=FMT_USD)
        for prev, col in zip(YEAR_COLS, PROJ_COLS):
            put(ws, f"{col}{r}", f"={prev}{r}*(1+{A}!$D{a})", fmt=FMT_USD)
        put(ws, f"I{r}", cagr(r), fmt=FMT_PCT)
    put(ws, f"B{S_REV_TOTAL}", "Total Revenue")
    for col in YEAR_COLS:
        put(ws, f"{col}{S_REV_TOTAL}", f"=SUM({col}{S_REV_FIRST}:{col}{S_REV_TOTAL - 1})", fmt=FMT_USD)
    put(ws, f"I{S_REV_TOTAL}", cagr(S_REV_TOTAL), fmt=FMT_PCT)
    total_style(ws, S_REV_TOTAL, last="I")
    put(ws, f"B{S_REV_GROWTH}", "  % Growth", italic=True)
    for prev, col in zip(YEAR_COLS, PROJ_COLS):
        put(ws, f"{col}{S_REV_GROWTH}",
            f"=IF({prev}{S_REV_TOTAL}=0,0,{col}{S_REV_TOTAL}/{prev}{S_REV_TOTAL}-1)",
            fmt=FMT_PCT, italic=True)

    # Operating income by segment
    section(ws, S_OI_HDR, "Operating Income by Segment", last="I")
    for i in range(N_SEGMENTS):
        r, a, rev = S_OI_FIRST + i, A_SEG_FIRST + i, S_REV_FIRST + i
        put(ws, f"B{r}", seg_label(a), color=GREEN)
        for col in YEAR_COLS:
            put(ws, f"{col}{r}", f"={col}{rev}*{A}!$E{a}", fmt=FMT_USD)
        put(ws, f"I{r}", cagr(r), fmt=FMT_PCT)
    put(ws, f"B{S_OI_TOTAL}", "Total Operating Income")
    for col in YEAR_COLS:
        put(ws, f"{col}{S_OI_TOTAL}", f"=SUM({col}{S_OI_FIRST}:{col}{S_OI_TOTAL - 1})", fmt=FMT_USD)
    put(ws, f"I{S_OI_TOTAL}", cagr(S_OI_TOTAL), fmt=FMT_PCT)
    total_style(ws, S_OI_TOTAL, last="I")
    put(ws, f"B{S_OI_MARGIN}", "  % Operating Margin", italic=True)
    for col in YEAR_COLS:
        put(ws, f"{col}{S_OI_MARGIN}",
            f"=IF({col}{S_REV_TOTAL}=0,0,{col}{S_OI_TOTAL}/{col}{S_REV_TOTAL})", fmt=FMT_PCT, italic=True)

    # Total company summary -- the link point for downstream tabs
    section(ws, S_SUM_HDR, "Total Company Summary (link downstream tabs to these rows)", last="I")
    put(ws, f"B{S_SUM_REV}", "Total Revenue", bold=True)
    put(ws, f"B{S_SUM_GROWTH}", "  % Growth", italic=True)
    put(ws, f"B{S_SUM_EBIT}", "Total Operating Income (EBIT)", bold=True)
    put(ws, f"B{S_SUM_MARGIN}", "  % Operating Margin", italic=True)
    for col in YEAR_COLS:
        put(ws, f"{col}{S_SUM_REV}", f"={col}{S_REV_TOTAL}", fmt=FMT_USD, bold=True)
        put(ws, f"{col}{S_SUM_EBIT}", f"={col}{S_OI_TOTAL}", fmt=FMT_USD, bold=True)
        put(ws, f"{col}{S_SUM_MARGIN}", f"={col}{S_OI_MARGIN}", fmt=FMT_PCT, italic=True)
    for col in PROJ_COLS:
        put(ws, f"{col}{S_SUM_GROWTH}", f"={col}{S_REV_GROWTH}", fmt=FMT_PCT, italic=True)
    put(ws, f"I{S_SUM_REV}", cagr(S_SUM_REV), fmt=FMT_PCT, bold=True)
    put(ws, f"I{S_SUM_EBIT}", cagr(S_SUM_EBIT), fmt=FMT_PCT, bold=True)

    put(ws, f"B{S_CHECK}", "Check: base-year revenue ties to Assumptions", italic=True)
    put(ws, f"C{S_CHECK}", f'=IF(ABS(C{S_SUM_REV}-{A}!$C${A_SEG_TOTAL})<0.001,"OK","ERROR")', bold=True)

    ws.freeze_panes = f"C{S_T + 1}"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    add_name(wb, "FY_Labels", f"{S}!$C${S_FY}:$H${S_FY}")
    add_name(wb, "Period", f"{S}!$C${S_T}:$H${S_T}")
    add_name(wb, "Total_Revenue", f"{S}!$C${S_SUM_REV}:$H${S_SUM_REV}")
    add_name(wb, "Total_EBIT", f"{S}!$C${S_SUM_EBIT}:$H${S_SUM_EBIT}")


def add_driver_inputs(wb):
    """Sections 4-5 on Assumptions: new input cells only, below the existing Model Notes."""
    ws = wb[A]
    for col, width in zip("GHI", (18, 18, 48)):
        ws.column_dimensions[col].width = width

    # 4. Working capital drivers -- year columns follow the C:H convention
    section(ws, A_WC_HDR, "4. Working Capital Drivers", last="I")
    put(ws, f"B{A_WC_FY}", "Fiscal Year", bold=True, fill=HEADER_FILL)
    for col in YEAR_COLS:
        c = put(ws, f"{col}{A_WC_FY}", f"={S}!{col}${S_FY}", color=GREEN, bold=True, fill=HEADER_FILL)
        c.alignment = Alignment(horizontal="right")
    put(ws, f"I{A_WC_FY}", "Notes / Source", bold=True, fill=HEADER_FILL)
    rows = [
        (A_COGS_PCT, "COGS (% of Revenue)", EXAMPLE_COGS_PCT, FMT_PCT,
         "Drives COGS for DIO / DPO; example placeholder"),
        (A_DSO, "DSO - Days Sales Outstanding", EXAMPLE_DSO, FMT_DAYS,
         "Receivables = DSO x Revenue / Days per Year"),
        (A_DIO, "DIO - Days Inventory Outstanding", EXAMPLE_DIO, FMT_DAYS,
         "Inventory = DIO x COGS / Days per Year"),
        (A_DPO, "DPO - Days Payables Outstanding", EXAMPLE_DPO, FMT_DAYS,
         "Payables = DPO x COGS / Days per Year"),
    ]
    for r, label, vals, fmt, note in rows:
        put(ws, f"B{r}", label)
        for col, v in zip(YEAR_COLS, vals):
            put(ws, f"{col}{r}", v, color=BLUE, fmt=fmt, fill=INPUT_FILL)
        put(ws, f"I{r}", note, italic=True)
    put(ws, f"B{A_DAYS}", "Days per Year")
    put(ws, f"C{A_DAYS}", 365, color=BLUE, fmt="0", fill=INPUT_FILL)
    put(ws, f"I{A_DAYS}", "Day-count basis for DSO / DIO / DPO", italic=True)
    pct_validation(ws, f"C{A_COGS_PCT}:H{A_COGS_PCT}", 0, 1)
    num_validation(ws, f"C{A_DSO}:H{A_DPO}", 0, 365, "number of days")

    put(ws, f"B{A_BAL_HDR}", "Base-Year Balances ($mm)", bold=True)
    put(ws, f"C{A_BAL_HDR}", f"={S}!C${S_FY}", color=GREEN, bold=True)
    ws[f"C{A_BAL_HDR}"].alignment = Alignment(horizontal="right")
    for r, label, val in zip((A_AR0, A_INV0, A_AP0),
                             ("Accounts Receivable", "Inventory", "Accounts Payable"), EXAMPLE_BALANCES):
        put(ws, f"B{r}", label)
        put(ws, f"C{r}", val, color=BLUE, fmt=FMT_USD, fill=INPUT_FILL)
        put(ws, f"I{r}", "Actual balance sheet value; example placeholder - replace", italic=True)

    # 5. Capex drivers
    section(ws, A_CPX_HDR, "5. Capex Drivers", last="I")
    put(ws, f"B{A_MAINT}", "Maintenance Capex (% of Total Revenue)")
    put(ws, f"C{A_MAINT}", EXAMPLE_MAINT_CAPEX, color=BLUE, fmt=FMT_PCT, fill=INPUT_FILL)
    put(ws, f"I{A_MAINT}", "Applied to total revenue each projection year; example placeholder", italic=True)
    pct_validation(ws, f"C{A_MAINT}", 0, 1)

    for col, text in (("B", "Segment"), ("C", "Growth Capex Intensity ($ per $1 incr. revenue)"),
                      ("I", "Notes / Source")):
        c = put(ws, f"{col}{A_INT_HDR}", text, bold=True, fill=HEADER_FILL)
        c.alignment = Alignment(wrap_text=True, horizontal="center" if col == "C" else "left",
                                vertical="center")
    ws.row_dimensions[A_INT_HDR].height = 42
    for i in range(N_SEGMENTS):
        r, a = A_INT_FIRST + i, A_SEG_FIRST + i
        put(ws, f"B{r}", f'=IF($B{a}="","",$B{a})')
        val = EXAMPLE_CAPEX_INTENSITY[i] if i < len(EXAMPLE_CAPEX_INTENSITY) else None
        put(ws, f"C{r}", val, color=BLUE, fmt=FMT_INTENSITY, fill=INPUT_FILL)
        put(ws, f"I{r}", "Example placeholder - replace" if val is not None else None, color=BLUE, italic=True)
    num_validation(ws, f"C{A_INT_FIRST}:C{A_INT_LAST}", 0, 10, "capex intensity")

    put(ws, f"B{A_DRV_NOTES}", "Driver Notes (sections 4-5)", bold=True)
    notes = [
        "Base-year days (column C) are for reference; base-year NWC uses the actual balances above.",
        "Growth capex = intensity x MAX(0, segment revenue increase); a shrinking segment adds no growth capex.",
        "Sign convention: Change in NWC and Total Capex are positive when they use cash - "
        "subtract both on the FCF tab.",
        "Named ranges: COGS_Pct, DSO_Days, DIO_Days, DPO_Days, DaysPerYear, MaintCapexPct, "
        "Seg_CapexIntensity, NWC, Change_NWC, Maint_Capex, Growth_Capex, Total_Capex.",
    ]
    for i, text in enumerate(notes, start=1):
        put(ws, f"B{A_DRV_NOTES + i}", f"{i}. {text}", italic=True)

    for name, row in [("COGS_Pct", A_COGS_PCT), ("DSO_Days", A_DSO), ("DIO_Days", A_DIO), ("DPO_Days", A_DPO)]:
        add_name(wb, name, f"{A}!$C${row}:$H${row}")
    add_name(wb, "DaysPerYear", f"{A}!$C${A_DAYS}")
    add_name(wb, "MaintCapexPct", f"{A}!$C${A_MAINT}")
    add_name(wb, "Seg_CapexIntensity", f"{A}!$C${A_INT_FIRST}:$C${A_INT_LAST}")


def timeseries_sheet(wb, name, title):
    """New tab with the shared header: title, units, FY labels and periods linked from Segment_Revenue."""
    ws = wb.create_sheet(name)
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 2
    ws.column_dimensions["B"].width = 44
    for col in YEAR_COLS:
        ws.column_dimensions[col].width = 13
    put(ws, "B1", f'=CompanyName&" - {title}"', bold=True, size=14)
    put(ws, "B2", '="("&Units&"; drivers link to the Assumptions and Segment_Revenue tabs)"', italic=True)
    put(ws, f"B{S_FY}", "Fiscal Year", bold=True, fill=HEADER_FILL)
    put(ws, f"B{S_T}", "Projection Period (t)", italic=True)
    for col in YEAR_COLS:
        put(ws, f"{col}{S_FY}", f"={S}!{col}${S_FY}", color=GREEN, bold=True, fill=HEADER_FILL)
        put(ws, f"{col}{S_T}", f"={S}!{col}${S_T}", color=GREEN, italic=True)
        for r in (S_FY, S_T):
            ws[f"{col}{r}"].alignment = Alignment(horizontal="right")
    ws.freeze_panes = f"C{S_T + 1}"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    return ws


def build_working_capital(wb):
    ws = timeseries_sheet(wb, W, "Working Capital Schedule")

    section(ws, W_DRV_HDR, "Drivers", last="H")
    put(ws, f"B{W_REV}", "Total Revenue")
    put(ws, f"B{W_COGS_PCT}", "  COGS (% of Revenue)", italic=True)
    put(ws, f"B{W_COGS}", "COGS")
    put(ws, f"B{W_DSO}", "DSO (days)")
    put(ws, f"B{W_DIO}", "DIO (days)")
    put(ws, f"B{W_DPO}", "DPO (days)")
    for col in YEAR_COLS:
        put(ws, f"{col}{W_REV}", f"={S}!{col}{S_SUM_REV}", color=GREEN, fmt=FMT_USD)
        put(ws, f"{col}{W_COGS_PCT}", f"={A}!{col}{A_COGS_PCT}", color=GREEN, fmt=FMT_PCT, italic=True)
        put(ws, f"{col}{W_COGS}", f"={col}{W_REV}*{col}{W_COGS_PCT}", fmt=FMT_USD)
        for r, a in ((W_DSO, A_DSO), (W_DIO, A_DIO), (W_DPO, A_DPO)):
            put(ws, f"{col}{r}", f"={A}!{col}{a}", color=GREEN, fmt=FMT_DAYS)

    section(ws, W_BAL_HDR, "Working Capital Balances", last="H")
    days = f"{A}!$C${A_DAYS}"
    for r, label, a0, drv, base in ((W_AR, "Accounts Receivable", A_AR0, W_DSO, W_REV),
                                    (W_INV, "Inventory", A_INV0, W_DIO, W_COGS),
                                    (W_AP, "Accounts Payable", A_AP0, W_DPO, W_COGS)):
        put(ws, f"B{r}", label)
        put(ws, f"C{r}", f"={A}!$C${a0}", color=GREEN, fmt=FMT_USD)
        for col in PROJ_COLS:
            put(ws, f"{col}{r}", f"={col}{drv}*{col}{base}/{days}", fmt=FMT_USD)
    put(ws, f"B{W_NWC}", "Net Working Capital (AR + Inventory - AP)")
    put(ws, f"B{W_NWC_PCT}", "  NWC (% of Revenue)", italic=True)
    for col in YEAR_COLS:
        put(ws, f"{col}{W_NWC}", f"={col}{W_AR}+{col}{W_INV}-{col}{W_AP}", fmt=FMT_USD)
        put(ws, f"{col}{W_NWC_PCT}", f"=IF({col}{W_REV}=0,0,{col}{W_NWC}/{col}{W_REV})", fmt=FMT_PCT, italic=True)
    total_style(ws, W_NWC)

    section(ws, W_CHG_HDR, "Change in Net Working Capital (link FCF tab to this row)", last="H")
    put(ws, f"B{W_CHG}", "Change in NWC (increase = cash outflow)")
    for prev, col in zip(YEAR_COLS, PROJ_COLS):
        put(ws, f"{col}{W_CHG}", f"={col}{W_NWC}-{prev}{W_NWC}", fmt=FMT_USD)
    total_style(ws, W_CHG)

    add_name(wb, "NWC", f"{W}!$C${W_NWC}:$H${W_NWC}")
    add_name(wb, "Change_NWC", f"{W}!$C${W_CHG}:$H${W_CHG}")


def build_capex(wb):
    ws = timeseries_sheet(wb, X, "Capital Expenditure Schedule")

    section(ws, X_MAINT_HDR, "Maintenance Capex", last="H")
    put(ws, f"B{X_REV}", "Total Revenue")
    put(ws, f"B{X_MAINT_PCT}", "  Maintenance Capex (% of Revenue)", italic=True)
    put(ws, f"B{X_MAINT}", "Maintenance Capex")
    for col in YEAR_COLS:
        put(ws, f"{col}{X_REV}", f"={S}!{col}{S_SUM_REV}", color=GREEN, fmt=FMT_USD)
    for col in PROJ_COLS:
        put(ws, f"{col}{X_MAINT_PCT}", f"={A}!$C${A_MAINT}", color=GREEN, fmt=FMT_PCT, italic=True)
        put(ws, f"{col}{X_MAINT}", f"={col}{X_REV}*{col}{X_MAINT_PCT}", fmt=FMT_USD)

    section(ws, X_INC_HDR, "Incremental Revenue by Segment (year-over-year increase)", last="H")
    section(ws, X_GC_HDR, "Growth Capex by Segment (intensity x MAX(0, incremental revenue))", last="H")
    for i in range(N_SEGMENTS):
        inc, gc, a, sr = X_INC_FIRST + i, X_GC_FIRST + i, A_SEG_FIRST + i, S_REV_FIRST + i
        put(ws, f"B{inc}", seg_label(a), color=GREEN)
        put(ws, f"B{gc}", seg_label(a), color=GREEN)
        for prev, col in zip(YEAR_COLS, PROJ_COLS):
            put(ws, f"{col}{inc}", f"={S}!{col}{sr}-{S}!{prev}{sr}", fmt=FMT_USD)
            put(ws, f"{col}{gc}", f"=MAX(0,{col}{inc})*{A}!$C${A_INT_FIRST + i}", fmt=FMT_USD)
    put(ws, f"B{X_INC_TOTAL}", "Total Incremental Revenue")
    put(ws, f"B{X_GC_TOTAL}", "Total Growth Capex")
    for col in PROJ_COLS:
        put(ws, f"{col}{X_INC_TOTAL}", f"=SUM({col}{X_INC_FIRST}:{col}{X_INC_TOTAL - 1})", fmt=FMT_USD)
        put(ws, f"{col}{X_GC_TOTAL}", f"=SUM({col}{X_GC_FIRST}:{col}{X_GC_TOTAL - 1})", fmt=FMT_USD)
    total_style(ws, X_INC_TOTAL)
    total_style(ws, X_GC_TOTAL)

    section(ws, X_SUM_HDR, "Total Capex (link FCF tab to the Total Capex row)", last="H")
    put(ws, f"B{X_SUM_MAINT}", "Maintenance Capex")
    put(ws, f"B{X_SUM_GROWTH}", "Growth Capex")
    put(ws, f"B{X_SUM_TOTAL}", "Total Capex (positive = cash outflow)")
    put(ws, f"B{X_SUM_PCT}", "  Total Capex (% of Revenue)", italic=True)
    for col in PROJ_COLS:
        put(ws, f"{col}{X_SUM_MAINT}", f"={col}{X_MAINT}", fmt=FMT_USD)
        put(ws, f"{col}{X_SUM_GROWTH}", f"={col}{X_GC_TOTAL}", fmt=FMT_USD)
        put(ws, f"{col}{X_SUM_TOTAL}", f"={col}{X_SUM_MAINT}+{col}{X_SUM_GROWTH}", fmt=FMT_USD)
        put(ws, f"{col}{X_SUM_PCT}", f"=IF({col}{X_REV}=0,0,{col}{X_SUM_TOTAL}/{col}{X_REV})",
            fmt=FMT_PCT, italic=True)
    total_style(ws, X_SUM_TOTAL)

    add_name(wb, "Maint_Capex", f"{X}!$C${X_SUM_MAINT}:$H${X_SUM_MAINT}")
    add_name(wb, "Growth_Capex", f"{X}!$C${X_SUM_GROWTH}:$H${X_SUM_GROWTH}")
    add_name(wb, "Total_Capex", f"{X}!$C${X_SUM_TOTAL}:$H${X_SUM_TOTAL}")


def add_wacc_dcf_inputs(wb):
    """Sections 6-7 on Assumptions: new input cells, plus linking Discount Rate (C15) to the WACC tab."""
    ws = wb[A]
    note = "Example placeholder - replace"

    section(ws, A_WACC_HDR, "6. WACC Inputs", last="I")
    section(ws, A_DCF_HDR, "7. DCF Inputs", last="I")
    rows = [
        (A_RF, "Risk-Free Rate", EXAMPLE_RF, FMT_PCT),
        (A_ERP, "Equity Risk Premium", EXAMPLE_ERP, FMT_PCT),
        (A_BETA, "Levered Beta", EXAMPLE_BETA, FMT_BETA),
        (A_KD, "Pre-Tax Cost of Debt", EXAMPLE_KD, FMT_PCT),
        (A_MCAP, "Market Capitalization ($mm)", EXAMPLE_MCAP, FMT_USD),
        (A_DEBT, "Total Debt ($mm)", EXAMPLE_DEBT, FMT_USD),
        (A_DA_PCT, "D&A (% of Revenue)", EXAMPLE_DA_PCT, FMT_PCT),
        (A_DCF_TGR, "Terminal Growth Rate", EXAMPLE_DCF_TGR, FMT_PCT),
        (A_EXIT_MULT, "Exit EV/EBITDA Multiple", EXAMPLE_EXIT_MULT, FMT_MULT),
        (A_NET_DEBT, "Net Debt ($mm)", EXAMPLE_NET_DEBT, FMT_USD),
        (A_SHARES, "Diluted Shares Outstanding (mm)", EXAMPLE_SHARES, FMT_SHARES),
        (A_MIDYEAR, "Mid-Year Convention (1 = mid-year, 0 = end-of-year)", EXAMPLE_MIDYEAR, "0"),
    ]
    for r, label, val, fmt in rows:
        put(ws, f"B{r}", label)
        put(ws, f"C{r}", val, color=BLUE, fmt=fmt, fill=INPUT_FILL)
        put(ws, f"I{r}", note, italic=True)

    # The model's discount rate now comes from the WACC tab
    put(ws, f"C{A_WACC}", f"={K}!$C${K_WACC}", color=GREEN, fmt=FMT_PCT, fill=PatternFill(fill_type=None))


def build_wacc(wb):
    ws = wb.create_sheet(K)
    ws.sheet_view.showGridLines = False
    for col, width in zip("ABC", (2, 40, 14)):
        ws.column_dimensions[col].width = width
    put(ws, "B1", '=CompanyName&" - WACC Build"', bold=True, size=14)
    put(ws, "B2", '="("&Units&"; inputs link to the Assumptions tab)"', italic=True)

    section(ws, K_IN_HDR, "Inputs (linked from Assumptions)", last="C")
    for r, label, a, fmt in ((K_RF, "Risk-Free Rate", A_RF, FMT_PCT),
                             (K_ERP, "Equity Risk Premium", A_ERP, FMT_PCT),
                             (K_BETA, "Levered Beta", A_BETA, FMT_BETA),
                             (K_KD, "Pre-Tax Cost of Debt", A_KD, FMT_PCT),
                             (K_TAX, "Tax Rate", A_TAX, FMT_PCT),
                             (K_MCAP, "Market Capitalization ($mm)", A_MCAP, FMT_USD),
                             (K_DEBT, "Total Debt ($mm)", A_DEBT, FMT_USD)):
        put(ws, f"B{r}", label)
        put(ws, f"C{r}", f"={A}!$C${a}", color=GREEN, fmt=fmt)

    section(ws, K_COC_HDR, "Cost of Capital", last="C")
    put(ws, f"B{K_KE}", "Cost of Equity (CAPM: Rf + Beta x ERP)")
    put(ws, f"C{K_KE}", f"=C{K_RF}+C{K_BETA}*C{K_ERP}", fmt=FMT_PCT)
    put(ws, f"B{K_KD_AT}", "After-Tax Cost of Debt (Kd x (1 - Tax))")
    put(ws, f"C{K_KD_AT}", f"=C{K_KD}*(1-C{K_TAX})", fmt=FMT_PCT)

    section(ws, K_CAP_HDR, "Capital Structure", last="C")
    put(ws, f"B{K_CAP}", "Total Capitalization (Market Cap + Debt)")
    put(ws, f"C{K_CAP}", f"=C{K_MCAP}+C{K_DEBT}", fmt=FMT_USD)
    put(ws, f"B{K_WE}", "Equity Weight (E / (D + E))")
    put(ws, f"C{K_WE}", f"=IF(C{K_CAP}=0,0,C{K_MCAP}/C{K_CAP})", fmt=FMT_PCT)
    put(ws, f"B{K_WD}", "Debt Weight (D / (D + E))")
    put(ws, f"C{K_WD}", f"=IF(C{K_CAP}=0,0,C{K_DEBT}/C{K_CAP})", fmt=FMT_PCT)

    section(ws, K_WACC_HDR, "Weighted Average Cost of Capital", last="C")
    put(ws, f"B{K_WACC}", "WACC (discount rate used by the model)")
    put(ws, f"C{K_WACC}", f"=C{K_WE}*C{K_KE}+C{K_WD}*C{K_KD_AT}", fmt=FMT_PCT)
    total_style(ws, K_WACC, last="C")


def build_dcf(wb):
    ws = timeseries_sheet(wb, D, "Discounted Cash Flow Valuation")

    # Unlevered free cash flow -- deductions shown as negatives
    section(ws, D_FCF_HDR, "Unlevered Free Cash Flow", last="H")
    put(ws, f"B{D_EBIT}", "EBIT")
    put(ws, f"B{D_TAX}", "Less: Taxes (EBIT x Tax Rate)")
    put(ws, f"B{D_DA}", "Plus: D&A (Revenue x D&A %)")
    put(ws, f"B{D_CAPEX}", "Less: Total Capex")
    put(ws, f"B{D_NWC}", "Less: Change in NWC")
    put(ws, f"B{D_UFCF}", "Unlevered Free Cash Flow")
    for col in PROJ_COLS:
        put(ws, f"{col}{D_EBIT}", f"={S}!{col}{S_SUM_EBIT}", color=GREEN, fmt=FMT_USD)
        put(ws, f"{col}{D_TAX}", f"=-{col}{D_EBIT}*{A}!$C${A_TAX}", fmt=FMT_USD)
        put(ws, f"{col}{D_DA}", f"={S}!{col}{S_SUM_REV}*{A}!$C${A_DA_PCT}", fmt=FMT_USD)
        put(ws, f"{col}{D_CAPEX}", f"=-{X}!{col}{X_SUM_TOTAL}", color=GREEN, fmt=FMT_USD)
        put(ws, f"{col}{D_NWC}", f"=-{W}!{col}{W_CHG}", color=GREEN, fmt=FMT_USD)
        put(ws, f"{col}{D_UFCF}", f"=SUM({col}{D_EBIT}:{col}{D_NWC})", fmt=FMT_USD)
    total_style(ws, D_UFCF)

    # Discounting
    section(ws, D_DISC_HDR, "Discounting", last="H")
    put(ws, f"B{D_RATE}", "Discount Rate (WACC)")
    put(ws, f"C{D_RATE}", f"={A}!$C${A_WACC}", color=GREEN, fmt=FMT_PCT)
    put(ws, f"B{D_TOGGLE}", "Mid-Year Convention (1 = mid-year, 0 = end-of-year)")
    put(ws, f"C{D_TOGGLE}", f"={A}!$C${A_MIDYEAR}", color=GREEN, fmt="0")
    put(ws, f"B{D_PERIOD}", "Discount Period (t - 0.5 x toggle)")
    put(ws, f"B{D_FACTOR}", "Discount Factor (1 / (1 + WACC)^period)")
    put(ws, f"B{D_PV}", "PV of Unlevered FCF")
    for col in PROJ_COLS:
        put(ws, f"{col}{D_PERIOD}", f"={col}${S_T}-0.5*$C${D_TOGGLE}", fmt=FMT_PERIOD)
        put(ws, f"{col}{D_FACTOR}", f"=1/(1+$C${D_RATE})^{col}{D_PERIOD}", fmt=FMT_FACTOR)
        put(ws, f"{col}{D_PV}", f"={col}{D_UFCF}*{col}{D_FACTOR}", fmt=FMT_USD)
    put(ws, f"B{D_SUM_PV}", "Sum of PV of Unlevered FCF")
    put(ws, f"C{D_SUM_PV}", f"=SUM(D{D_PV}:H{D_PV})", fmt=FMT_USD)
    total_style(ws, D_SUM_PV, last="C")

    # Terminal value & valuation: C = perpetuity growth method, D = exit multiple method
    section(ws, D_VAL_HDR, "Terminal Value & Valuation", last="H")
    for col, text in (("C", "Perpetuity Growth"), ("D", "Exit Multiple")):
        c = put(ws, f"{col}{D_METHOD}", text, bold=True, fill=HEADER_FILL)
        c.alignment = Alignment(wrap_text=True, horizontal="center", vertical="center")
    ws[f"B{D_METHOD}"].fill = HEADER_FILL
    ws.row_dimensions[D_METHOD].height = 28
    rate, last = f"$C${D_RATE}", PROJ_COLS[-1]

    put(ws, f"B{D_TV_FCF}", "Terminal-Year Unlevered FCF")
    put(ws, f"C{D_TV_FCF}", f"={last}{D_UFCF}", fmt=FMT_USD)
    put(ws, f"B{D_TV_G}", "Terminal Growth Rate")
    put(ws, f"C{D_TV_G}", f"={A}!$C${A_DCF_TGR}", color=GREEN, fmt=FMT_PCT)
    put(ws, f"B{D_TV_EBITDA}", "Terminal-Year EBITDA (EBIT + D&A)")
    put(ws, f"D{D_TV_EBITDA}", f"={last}{D_EBIT}+{last}{D_DA}", fmt=FMT_USD)
    put(ws, f"B{D_TV_MULT}", "Exit EV/EBITDA Multiple")
    put(ws, f"D{D_TV_MULT}", f"={A}!$C${A_EXIT_MULT}", color=GREEN, fmt=FMT_MULT)
    put(ws, f"B{D_TV}", "Terminal Value")
    put(ws, f"C{D_TV}", f"=IF({rate}<=C{D_TV_G},0,C{D_TV_FCF}*(1+C{D_TV_G})/({rate}-C{D_TV_G}))", fmt=FMT_USD)
    put(ws, f"D{D_TV}", f"=D{D_TV_EBITDA}*D{D_TV_MULT}", fmt=FMT_USD)

    put(ws, f"B{D_TV_PERIOD}", "Terminal Discount Period (end of final year)")
    put(ws, f"B{D_TV_FACTOR}", "Terminal Discount Factor")
    put(ws, f"B{D_TV_PV}", "PV of Terminal Value")
    put(ws, f"B{D_EV_PV_FCF}", "Plus: Sum of PV of Unlevered FCF")
    put(ws, f"B{D_EV}", "Enterprise Value")
    put(ws, f"B{D_NET_DEBT}", "Less: Net Debt")
    put(ws, f"B{D_EQUITY}", "Equity Value")
    put(ws, f"B{D_SHARES}", "Diluted Shares Outstanding (mm)")
    put(ws, f"B{D_PRICE}", "Implied Share Price ($ per share) - OUTPUT")
    for col in "CD":
        put(ws, f"{col}{D_TV_PERIOD}", f"=${last}${S_T}", fmt=FMT_PERIOD)
        put(ws, f"{col}{D_TV_FACTOR}", f"=1/(1+{rate})^{col}{D_TV_PERIOD}", fmt=FMT_FACTOR)
        put(ws, f"{col}{D_TV_PV}", f"={col}{D_TV}*{col}{D_TV_FACTOR}", fmt=FMT_USD)
        put(ws, f"{col}{D_EV_PV_FCF}", f"=$C${D_SUM_PV}", fmt=FMT_USD)
        put(ws, f"{col}{D_EV}", f"={col}{D_TV_PV}+{col}{D_EV_PV_FCF}", fmt=FMT_USD)
        put(ws, f"{col}{D_NET_DEBT}", f"=-{A}!$C${A_NET_DEBT}", color=GREEN, fmt=FMT_USD)
        put(ws, f"{col}{D_EQUITY}", f"={col}{D_EV}+{col}{D_NET_DEBT}", fmt=FMT_USD)
        put(ws, f"{col}{D_SHARES}", f"={A}!$C${A_SHARES}", color=GREEN, fmt=FMT_SHARES)
        put(ws, f"{col}{D_PRICE}", f"=IF({col}{D_SHARES}=0,0,{col}{D_EQUITY}/{col}{D_SHARES})", fmt=FMT_PRICE)
    for r in (D_EV, D_EQUITY, D_PRICE):
        total_style(ws, r, last="D")

    # Cross-checks
    section(ws, D_CHK_HDR, "Cross-Checks", last="H")
    put(ws, f"B{D_TV_PCT}", "PV of Terminal Value (% of Enterprise Value)")
    for col in "CD":
        put(ws, f"{col}{D_TV_PCT}", f"=IF({col}{D_EV}=0,0,{col}{D_TV_PV}/{col}{D_EV})", fmt=FMT_PCT)
    put(ws, f"B{D_IMPL_MULT}", "Implied Exit EV/EBITDA (from perpetuity method)")
    put(ws, f"C{D_IMPL_MULT}", f"=IF(D{D_TV_EBITDA}=0,0,C{D_TV}/D{D_TV_EBITDA})", fmt=FMT_MULT)
    put(ws, f"B{D_IMPL_G}", "Implied Perpetual Growth (from exit multiple method)")
    put(ws, f"D{D_IMPL_G}",
        f"=IF(D{D_TV}+C{D_TV_FCF}=0,0,(D{D_TV}*{rate}-C{D_TV_FCF})/(D{D_TV}+C{D_TV_FCF}))", fmt=FMT_PCT)


def build_scenarios(wb, active, snapshot):
    """Scenarios tab, the selector on Assumptions row 3, and repointing the Assumptions drivers to Live."""
    a = wb[A]
    put(a, f"B{A_SCEN}", "Active Scenario (Bear / Base / Bull)", bold=True)
    put(a, f"C{A_SCEN}", active, color=BLUE, fill=INPUT_FILL)
    dv = DataValidation(type="list", formula1='"' + ",".join(SCENARIOS) + '"', allow_blank=False,
                        showErrorMessage=True, errorTitle="Invalid scenario", error="Choose Bear, Base or Bull.")
    a.add_data_validation(dv)
    dv.add(f"C{A_SCEN}")
    put(a, f"D{A_SCEN}", "Scenario # (1/2/3)")
    put(a, f"E{A_SCEN}", f"=MATCH($C${A_SCEN},{SC}!$C${SC_HDR}:$E${SC_HDR},0)")
    case = f"{A}!$E${A_SCEN}"

    ws = wb.create_sheet(SC)
    ws.sheet_view.showGridLines = False
    for col, width in zip("ABCDEF", (2, 40, 12, 12, 12, 12)):
        ws.column_dimensions[col].width = width
    put(ws, "B1", '=CompanyName&" - Scenarios"', bold=True, size=14)
    put(ws, "B2", f'="(Edit Bear / Base / Bull inputs; Live = active case: "&{A}!$C${A_SCEN}&")"', italic=True)
    for col, text in zip("BCDEF", ("Driver",) + SCENARIOS + ("Live",)):
        c = put(ws, f"{col}{SC_HDR}", text, bold=True, fill=HEADER_FILL)
        c.alignment = Alignment(horizontal="left" if col == "B" else "right")
    ws.freeze_panes = f"C{SC_HDR + 1}"

    def driver_row(r, label, base, shift, fmt, label_color=BLACK):
        put(ws, f"B{r}", label, color=label_color)
        bear, bull = SCENARIO_SHIFTS[shift]
        vals = (None, None, None) if base is None else (round(base + bear, 6), base, round(base + bull, 6))
        for col, v in zip("CDE", vals):
            put(ws, f"{col}{r}", v, color=BLUE, fmt=fmt, fill=INPUT_FILL)
        put(ws, f"F{r}", f"=CHOOSE({case},C{r},D{r},E{r})", fmt=fmt, bold=True)

    def repoint(ref, sc_row, fmt):
        put(a, ref, f"={SC}!$F${sc_row}", color=GREEN, fmt=fmt, fill=PatternFill(fill_type=None))

    section(ws, SC_GROWTH_HDR, "Segment Growth Rate (% / yr)")
    section(ws, SC_MARGIN_HDR, "Segment Operating Margin (%)")
    for i in range(N_SEGMENTS):
        seg = EXAMPLE_SEGMENTS[i] if i < len(EXAMPLE_SEGMENTS) else (None,) * 4
        g, m, ar = SC_GROWTH_FIRST + i, SC_MARGIN_FIRST + i, A_SEG_FIRST + i
        driver_row(g, seg_label(ar), seg[2], "growth", FMT_PCT, GREEN)
        driver_row(m, seg_label(ar), seg[3], "margin", FMT_PCT, GREEN)
        repoint(f"D{ar}", g, FMT_PCT)
        repoint(f"E{ar}", m, FMT_PCT)

    section(ws, SC_DAYS_HDR, "Working Capital Days (projection years)")
    for first, name, a_row, example, shift in ((SC_DSO_FIRST, "DSO", A_DSO, EXAMPLE_DSO, "dso"),
                                               (SC_DIO_FIRST, "DIO", A_DIO, EXAMPLE_DIO, "dio"),
                                               (SC_DPO_FIRST, "DPO", A_DPO, EXAMPLE_DPO, "dpo")):
        for j, col in enumerate(PROJ_COLS):
            r = first + j
            driver_row(r, f'="{name} - "&{S}!{col}${S_FY}', example[j + 1], shift, FMT_DAYS, GREEN)
            repoint(f"{col}{a_row}", r, FMT_DAYS)

    section(ws, SC_OTHER_HDR, "Capex, WACC & Terminal Value")
    for r, label, a_row, base, shift, fmt in (
            (SC_MAINT, "Maintenance Capex (% of Revenue)", A_MAINT, EXAMPLE_MAINT_CAPEX, "maint", FMT_PCT),
            (SC_BETA, "Levered Beta", A_BETA, EXAMPLE_BETA, "beta", FMT_BETA),
            (SC_TGR, "Terminal Growth Rate", A_DCF_TGR, EXAMPLE_DCF_TGR, "tgr", FMT_PCT),
            (SC_EXIT, "Exit EV/EBITDA Multiple", A_EXIT_MULT, EXAMPLE_EXIT_MULT, "exit", FMT_MULT)):
        driver_row(r, label, base, shift, fmt)
        repoint(f"C{a_row}", r, fmt)

    # Static snapshot written by main() after recalculating each scenario in LibreOffice
    section(ws, SC_SUM_HDR, "Scenario Summary (snapshot values, rerun script to refresh)")
    for col, text in zip("BCDE", ("Output",) + SCENARIOS):
        c = put(ws, f"{col}{SC_SUM_COLS}", text, bold=True, fill=HEADER_FILL)
        c.alignment = Alignment(horizontal="left" if col == "B" else "right")
    for k, (r, label, fmt) in enumerate(((SC_EV_PGM, "Enterprise Value - Perpetuity Growth ($mm)", FMT_USD),
                                         (SC_EV_EXIT, "Enterprise Value - Exit Multiple ($mm)", FMT_USD),
                                         (SC_PX_PGM, "Implied Share Price - Perpetuity Growth ($)", FMT_PRICE),
                                         (SC_PX_EXIT, "Implied Share Price - Exit Multiple ($)", FMT_PRICE))):
        put(ws, f"B{r}", label)
        for col, name in zip("CDE", SCENARIOS):
            put(ws, f"{col}{r}", snapshot[name][k] if snapshot else None, fmt=fmt)


def build_sensitivity(wb):
    ws = wb.create_sheet(SN)
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 2
    ws.column_dimensions["B"].width = 30
    for col in GRID_COLS:
        ws.column_dimensions[col].width = 12
    put(ws, "B1", '=CompanyName&" - Sensitivity Analysis"', bold=True, size=14)
    put(ws, "B2", f'="(Implied share price, $ per share; centered on the live "&{A}!$C${A_SCEN}&" case)"',
        italic=True)

    section(ws, SN_STEP_HDR, "Grid Steps", last="G")
    for r, label, val, fmt in ((SN_STEP_WACC, "WACC Step", SENS_STEPS[0], FMT_PCT),
                               (SN_STEP_TGR, "Terminal Growth Step", SENS_STEPS[1], FMT_PCT),
                               (SN_STEP_MULT, "Exit Multiple Step", SENS_STEPS[2], FMT_MULT)):
        put(ws, f"B{r}", label)
        put(ws, f"C{r}", val, color=BLUE, fmt=fmt, fill=INPUT_FILL)

    # Shared pieces of each grid cell's formula (all from the DCF tab)
    pv_fcf = f"SUMPRODUCT({D}!$D${D_UFCF}:$H${D_UFCF}/(1+{{w}})^{D}!$D${D_PERIOD}:$H${D_PERIOD})"
    shares = {"C": f"{D}!$C${D_SHARES}", "D": f"{D}!$D${D_SHARES}"}
    grids = (
        (SN_PGM_HDR, SN_PGM_AXIS, SN_PGM_FIRST, "Perpetuity Growth Method: WACC vs. Terminal Growth",
         "WACC \\ Terminal Growth", f"{D}!$C${D_TV_G}", SN_STEP_TGR, FMT_PCT,
         lambda w, x: (f"=IF(OR({w}<={x},{shares['C']}=0),0,({pv_fcf.format(w=w)}"
                       f"+{D}!$C${D_TV_FCF}*(1+{x})/({w}-{x})/(1+{w})^{D}!$C${D_TV_PERIOD}"
                       f"+{D}!$C${D_NET_DEBT})/{shares['C']})")),
        (SN_EXIT_HDR, SN_EXIT_AXIS, SN_EXIT_FIRST, "Exit Multiple Method: WACC vs. Exit EV/EBITDA Multiple",
         "WACC \\ Exit Multiple", f"{D}!$D${D_TV_MULT}", SN_STEP_MULT, FMT_MULT,
         lambda w, x: (f"=IF({shares['D']}=0,0,({pv_fcf.format(w=w)}"
                       f"+{D}!$D${D_TV_EBITDA}*{x}/(1+{w})^{D}!$D${D_TV_PERIOD}"
                       f"+{D}!$D${D_NET_DEBT})/{shares['D']})")),
    )
    mid_col = GRID_COLS[GRID_MID]
    for hdr, axis, first, title, corner, live_x, step_row, x_fmt, cell_formula in grids:
        section(ws, hdr, f"Implied Share Price - {title}", last="G")
        put(ws, f"B{axis}", corner, bold=True, fill=HEADER_FILL)
        mid_row = first + GRID_MID
        for k, col in enumerate(GRID_COLS):
            off = k - GRID_MID
            if off == 0:
                put(ws, f"{col}{axis}", f"={live_x}", color=GREEN, fmt=x_fmt, bold=True, fill=HEADER_FILL)
            else:
                put(ws, f"{col}{axis}", f"=${mid_col}${axis}{off:+d}*$C${step_row}", fmt=x_fmt, bold=True,
                    fill=HEADER_FILL)
        for k in range(len(GRID_COLS)):
            r, off = first + k, k - GRID_MID
            if off == 0:
                put(ws, f"B{r}", f"={D}!$C${D_RATE}", color=GREEN, fmt=FMT_PCT, bold=True, fill=HEADER_FILL)
            else:
                put(ws, f"B{r}", f"=$B${mid_row}{off:+d}*$C${SN_STEP_WACC}", fmt=FMT_PCT, bold=True,
                    fill=HEADER_FILL)
            for col in GRID_COLS:
                put(ws, f"{col}{r}", cell_formula(f"$B{r}", f"{col}${axis}"), fmt=FMT_PRICE)
        put(ws, f"{mid_col}{mid_row}", ws[f"{mid_col}{mid_row}"].value, fmt=FMT_PRICE, bold=True,
            fill=CENTER_FILL)


def build_workbook(active, snapshot=None):
    wb = Workbook()
    build_assumptions(wb)
    build_segment_revenue(wb)
    add_driver_inputs(wb)
    build_working_capital(wb)
    build_capex(wb)
    add_wacc_dcf_inputs(wb)
    build_wacc(wb)
    build_dcf(wb)
    build_scenarios(wb, active, snapshot)
    build_sensitivity(wb)
    return wb


def find_soffice():
    for cand in ("soffice", "libreoffice", "/Applications/LibreOffice.app/Contents/MacOS/soffice"):
        path = shutil.which(cand)
        if path:
            return path
    sys.exit("LibreOffice (soffice) not found on PATH; it is required to recalculate the scenarios.")


def recalc(src, out_dir, soffice, profile):
    """Recalculate every formula via a headless LibreOffice convert; returns the recalculated copy."""
    subprocess.run([soffice, "--headless", "--calc", f"-env:UserInstallation={profile.as_uri()}",
                    "--convert-to", "xlsx", "--outdir", str(out_dir), str(src)],
                   check=True, capture_output=True, timeout=300)
    return out_dir / src.name


def main():
    soffice = find_soffice()
    with tempfile.TemporaryDirectory() as tmp:
        tmp = Path(tmp)
        src_dir, out_dir, profile = tmp / "src", tmp / "out", tmp / "lo_profile"
        src_dir.mkdir()
        snapshot = {}
        for name in SCENARIOS:
            src = src_dir / f"{name}.xlsx"
            build_workbook(name).save(src)
            dcf = load_workbook(recalc(src, out_dir, soffice, profile), data_only=True)[D]
            snapshot[name] = [dcf[f"C{D_EV}"].value, dcf[f"D{D_EV}"].value,
                              dcf[f"C{D_PRICE}"].value, dcf[f"D{D_PRICE}"].value]
            print(f"{name}: share price {snapshot[name][2]:.2f} (perpetuity) / {snapshot[name][3]:.2f} (exit)")
        src = src_dir / OUT.name
        build_workbook("Base", snapshot).save(src)
        shutil.copyfile(recalc(src, out_dir, soffice, profile), OUT)
    print(f"Saved {OUT}")


if __name__ == "__main__":
    main()
