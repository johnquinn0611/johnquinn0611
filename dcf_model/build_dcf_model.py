"""Build the DCF valuation workbook (Assumptions + Segment_Revenue tabs).

Run:  python dcf_model/build_dcf_model.py   ->   dcf_model/DCF_Model.xlsx

Layout conventions shared by every time-series tab (keep new tabs on these):
  * Column B = row labels
  * Column C = base year (last actual fiscal year, period t = 0)
  * Columns D:H = projection years 1-5 (periods t = 1..5)
"""
from datetime import date
from pathlib import Path

from openpyxl import Workbook
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

FMT_USD = '$#,##0.0_);($#,##0.0);"-"_)'
FMT_PCT = '0.0%_);(0.0%);"-"_)'
FMT_YEAR = "0"
FMT_DATE = "mmm d, yyyy"

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

# ---- Assumptions tab cell map --------------------------------------------
A = "Assumptions"
A_COMPANY, A_TICKER, A_UNITS, A_BASE_YEAR, A_VAL_DATE = 7, 8, 9, 10, 11
A_TAX, A_WACC, A_TGR, A_TGR_CHECK = 14, 15, 16, 17
A_SEG_HDR = 20
A_SEG_FIRST = A_SEG_HDR + 1
A_SEG_LAST = A_SEG_FIRST + N_SEGMENTS - 1
A_SEG_TOTAL = A_SEG_LAST + 1
A_NOTES = A_SEG_TOTAL + 2

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

    def seg_label(a_row):
        return f'=IF({A}!$B{a_row}="","",{A}!$B{a_row})'

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


def main():
    wb = Workbook()
    build_assumptions(wb)
    build_segment_revenue(wb)
    wb.save(OUT)
    print(f"Saved {OUT}")


if __name__ == "__main__":
    main()
