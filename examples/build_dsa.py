from fastpyxl import Workbook
from fastpyxl.comments import Comment
from fastpyxl.utils import get_column_letter

YEARS = list(range(2021, 2036))  # 2021-2035
N_YEARS = len(YEARS)  # 15
FIRST_COL = 2  # column B
LAST_COL = FIRST_COL + N_YEARS - 1  # column P
HIST_YEARS = [2021, 2022, 2023, 2024]  # actual
EST_YEAR = 2025  # estimate
PROJ_YEARS = list(range(2026, 2036))  # projection 2026-2035


def col_letter(i):
    return get_column_letter(i)


FIRST_COL_L = col_letter(FIRST_COL)
LAST_COL_L = col_letter(LAST_COL)
PROJ_START_COL = FIRST_COL + YEARS.index(PROJ_YEARS[0])
PROJ_START_L = col_letter(PROJ_START_COL)

wb = Workbook()
wb.remove(wb.active)


def new_sheet(name):
    return wb.create_sheet(name)


def set_col_widths(ws, first_width=42, year_width=10, extra=None):
    del ws, first_width, year_width, extra


def title_block(ws, title, subtitle, last_col_letter=None):
    lcl = last_col_letter or LAST_COL_L
    ws.merge_cells(f"A1:{lcl}1")
    ws["A1"] = title
    ws.merge_cells(f"A2:{lcl}2")
    ws["A2"] = subtitle
    ws.merge_cells(f"A3:{lcl}3")


def year_header_row(ws, row, start_col=FIRST_COL, label="Year"):
    ws.cell(row=row, column=1, value=label)
    for i, y in enumerate(YEARS):
        ws.cell(row=row, column=start_col + i, value=str(y))
    tag_row = row + 1
    ws.cell(row=tag_row, column=1, value="Status")
    for i, y in enumerate(YEARS):
        if y in HIST_YEARS:
            tag = "Actual"
        elif y == EST_YEAR:
            tag = "Estimate"
        else:
            tag = "Projection"
        ws.cell(row=tag_row, column=start_col + i, value=tag)
    return tag_row


def section_row(ws, row, text, last_col_letter=None):
    lcl = last_col_letter or LAST_COL_L
    ws.merge_cells(f"A{row}:{lcl}{row}")
    ws.cell(row=row, column=1, value=text)


def label_cell(ws, row, text, indent=1, bold=False, italic=False, note=None):
    del indent, bold, italic
    c = ws.cell(row=row, column=1, value=text)
    if note:
        c.comment = Comment(note, "DSA Model")
    return c


def data_row(ws, row, kind="input", fmt=None, start_col=FIRST_COL, end_col=LAST_COL):
    del kind, fmt  # kept for call-site compatibility
    for col in range(start_col, end_col + 1):
        ws.cell(row=row, column=col)


def fill_input_row(ws, row, values, fmt=None):
    del fmt
    data_row(ws, row, kind="input")
    for i, v in enumerate(values):
        if v is not None:
            ws.cell(row=row, column=FIRST_COL + i, value=v)


def fill_formula_row(ws, row, formula_fn, fmt=None, kind="formula", start_i=0):
    del fmt, kind
    data_row(ws, row, kind="formula")
    for i in range(start_i, len(YEARS)):
        cl = col_letter(FIRST_COL + i)
        pcl = col_letter(FIRST_COL + i - 1) if i > 0 else None
        f = formula_fn(i, cl, pcl)
        if f is not None:
            ws.cell(row=row, column=FIRST_COL + i, value=f)


def note_box(ws, row, col_start, col_end, lines, height_per_line=14, fill=None):
    del height_per_line, fill
    top = row
    bottom = row + len(lines) - 1
    ws.merge_cells(start_row=top, start_column=col_start, end_row=bottom, end_column=col_end)
    ws.cell(row=top, column=col_start, value="\n".join(lines))
    return bottom


print("Setup complete. Years:", YEARS)
print("First col:", FIRST_COL_L, "Last col:", LAST_COL_L, "Proj start:", PROJ_START_L)

# =================================================================
# SHEET 1: COVER
# =================================================================
ws = new_sheet("Cover")
set_col_widths(ws, first_width=4)

ws.merge_cells("B2:O3")
ws["B2"] = "Sovereign Debt Sustainability Analysis (DSA)"

ws.merge_cells("B4:O4")
ws["B4"] = (
    "Illustrative Template — Country: [Enter Country Name]  |  "
    "Base Year: 2025  |  Projection Horizon: 2026-2035"
)

ws.merge_cells("B6:O6")

overview_lines = [
    "PURPOSE",
    "This workbook provides a full public/sovereign Debt Sustainability Analysis (DSA) framework, broadly consistent with the approach used by the IMF/World Bank",
    "Debt Sustainability Framework. It projects the evolution of public debt, gross financing needs, and key sustainability indicators under a baseline scenario and",
    "a set of standardized macro-fiscal stress tests, and flags risks against common sustainability thresholds.",
    "",
    "HOW TO USE THIS WORKBOOK",
    "1. Go to the 'Assumptions' tab and replace all BLUE cells with your country's actual/estimated macro-fiscal data (historical 2021-2024, estimate 2025).",
    "2. Enter your projection assumptions for 2026-2035 (also blue cells) — growth, inflation, interest rates, primary balance, exchange rate, etc.",
    "3. The 'Debt Dynamics' tab automatically computes the baseline debt path and decomposes the change in debt into its economic drivers.",
    "4. The 'Gross Financing Needs' tab computes annual financing requirements (amortization + interest + primary deficit).",
    "5. The 'Scenario Analysis' tab stress-tests the baseline against growth, interest rate, primary balance, combined, and exchange rate shocks (edit the yellow shock cells).",
    "6. The 'Sustainability Indicators' tab benchmarks results against standard thresholds and flags risk levels automatically.",
    "7. The 'Dashboard' tab summarizes results visually for presentation.",
    "",
    "INPUT CONVENTIONS",
]
r = 8
for line in overview_lines:
    ws.merge_cells(f"B{r}:O{r}")
    ws.cell(row=r, column=2, value=line)
    r += 1

r += 1
legend = [
    ("Hardcoded input", "User should review and edit"),
    ("Formula / calculation", "Do not overwrite"),
    ("Cross-sheet link", "Pulls from another worksheet in this workbook"),
    ("Key assumption or shock", "Requires attention"),
]
for label, desc in legend:
    ws.cell(row=r, column=2, value=label)
    ws.merge_cells(f"C{r}:O{r}")
    ws.cell(row=r, column=3, value=desc)
    r += 1

r += 1
ws.merge_cells(f"B{r}:O{r}")
ws.cell(row=r, column=2, value="WORKBOOK CONTENTS")
r += 1
contents = [
    ("Assumptions", "All macroeconomic, fiscal, financing, and external sector inputs (2021-2035)"),
    ("Debt Dynamics", "Baseline public debt/GDP path and decomposition of debt drivers"),
    (
        "Gross Financing Needs",
        "Annual financing requirement: amortization + interest + primary deficit",
    ),
    (
        "Scenario Analysis",
        "Standardized stress tests vs. baseline: growth, interest rate, primary balance, combined, FX shocks",
    ),
    ("Sustainability Indicators", "Key ratios vs. thresholds, automated risk flags"),
    ("Dashboard", "Summary charts for presentation"),
]
for name, desc in contents:
    ws.cell(row=r, column=2, value=name)
    ws.merge_cells(f"D{r}:O{r}")
    ws.cell(row=r, column=4, value=desc)
    r += 1

r += 1
ws.merge_cells(f"B{r}:O{r + 2}")
disclaimer = (
    "DISCLAIMER: This is an analytical template with illustrative placeholder figures. It is not country-specific advice and does not constitute an "
    "official IMF/World Bank DSA. Verify all formulas and replace all illustrative inputs before use in any decision-making context."
)
ws.cell(row=r, column=2, value=disclaimer)
print("Cover sheet built. Rows used:", r)

# =================================================================
# SHEET 2: ASSUMPTIONS
# =================================================================
# Row map (referenced by every other sheet):
A_G = 9  # real GDP growth (fraction)
A_PI = 10  # GDP deflator inflation (fraction)
A_NGDP_G = 11  # nominal GDP growth (fraction, formula)
A_NGDP = 12  # nominal GDP level (LCU bn)
A_REV = 15  # revenue and grants (% GDP, fraction)
A_EXP = 16  # primary expenditure (% GDP, fraction)
A_PB = 17  # primary balance (% GDP, fraction, formula)
A_IR = 20  # effective nominal interest rate on debt (fraction)
A_FX_SHARE = 21  # FX-denominated share of debt (fraction)
A_DEP = 22  # exchange rate depreciation (+ = depreciation, fraction)
A_AMORT = 23  # amortization rate (% of previous-year debt stock, fraction)
A_D0_ROW = 26  # initial public debt (% GDP, end-2020) — single cell in column B

INIT_DEBT_REF = f"Assumptions!$B${A_D0_ROW}"

ws = new_sheet("Assumptions")
set_col_widths(ws)
title_block(ws, "Assumptions", "All values are decimal fractions unless noted (e.g. 0.035 = 3.5%)")
year_header_row(ws, 5)

section_row(ws, 8, "MACROECONOMIC FRAMEWORK")
label_cell(ws, A_G, "Real GDP growth")
fill_input_row(
    ws,
    A_G,
    [
        0.035,
        0.042,
        0.030,
        0.028,
        0.026,
        0.028,
        0.030,
        0.032,
        0.033,
        0.034,
        0.035,
        0.035,
        0.035,
        0.035,
        0.035,
    ],
)
label_cell(ws, A_PI, "GDP deflator inflation")
fill_input_row(
    ws,
    A_PI,
    [
        0.050,
        0.075,
        0.060,
        0.045,
        0.040,
        0.040,
        0.038,
        0.036,
        0.035,
        0.035,
        0.035,
        0.035,
        0.035,
        0.035,
        0.035,
    ],
)
label_cell(ws, A_NGDP_G, "Nominal GDP growth")
fill_formula_row(
    ws,
    A_NGDP_G,
    lambda i, cl, pcl: f"=(1+{cl}{A_G})*(1+{cl}{A_PI})-1",
)
label_cell(ws, A_NGDP, "Nominal GDP (LCU bn)")
fill_input_row(ws, A_NGDP, [1000.0] + [None] * (N_YEARS - 1))
fill_formula_row(
    ws,
    A_NGDP,
    lambda i, cl, pcl: f"={pcl}{A_NGDP}*(1+{cl}{A_NGDP_G})",
    start_i=1,
)

section_row(ws, 14, "FISCAL FRAMEWORK (% OF GDP)")
label_cell(ws, A_REV, "Revenue and grants")
fill_input_row(
    ws,
    A_REV,
    [
        0.235,
        0.238,
        0.240,
        0.242,
        0.243,
        0.244,
        0.245,
        0.246,
        0.247,
        0.248,
        0.248,
        0.248,
        0.248,
        0.248,
        0.248,
    ],
)
label_cell(ws, A_EXP, "Primary expenditure")
fill_input_row(
    ws,
    A_EXP,
    [
        0.265,
        0.272,
        0.268,
        0.264,
        0.262,
        0.258,
        0.255,
        0.252,
        0.250,
        0.249,
        0.248,
        0.247,
        0.246,
        0.246,
        0.246,
    ],
)
label_cell(ws, A_PB, "Primary balance")
fill_formula_row(ws, A_PB, lambda i, cl, pcl: f"={cl}{A_REV}-{cl}{A_EXP}")

section_row(ws, 19, "FINANCING AND DEBT STRUCTURE")
label_cell(ws, A_IR, "Effective nominal interest rate on debt")
fill_input_row(
    ws,
    A_IR,
    [
        0.055,
        0.058,
        0.062,
        0.066,
        0.068,
        0.068,
        0.067,
        0.066,
        0.065,
        0.065,
        0.064,
        0.064,
        0.063,
        0.063,
        0.062,
    ],
)
label_cell(ws, A_FX_SHARE, "FX-denominated share of public debt")
fill_input_row(
    ws,
    A_FX_SHARE,
    [0.42, 0.41, 0.41, 0.40, 0.40, 0.39, 0.39, 0.38, 0.38, 0.37, 0.37, 0.36, 0.36, 0.35, 0.35],
)
label_cell(ws, A_DEP, "Exchange rate depreciation (+ = depreciation)")
fill_input_row(
    ws,
    A_DEP,
    [0.030, 0.080, 0.050, 0.030, 0.025] + [0.020] * len(PROJ_YEARS),
)
label_cell(ws, A_AMORT, "Amortization rate (share of previous-year debt)")
fill_input_row(ws, A_AMORT, [0.12] * N_YEARS)

section_row(ws, 25, "INITIAL CONDITIONS")
label_cell(ws, A_D0_ROW, "Public debt (% GDP), end-2020")
ws.cell(row=A_D0_ROW, column=FIRST_COL, value=0.620)

print("Assumptions sheet built.")

# =================================================================
# SHEET 3: DEBT DYNAMICS
# =================================================================
# d_t = d_{t-1} * (1 + i_t + eps_t*alpha_t) / ((1+g_t)*(1+pi_t)) - pb_t
DD = "'Debt Dynamics'"
DD_DEBT = 9  # public debt (% GDP)
DD_CHANGE = 10  # change in debt (pp of GDP)
DD_PB_C = 12  # primary balance contribution
DD_IR_C = 13  # real interest rate contribution
DD_G_C = 14  # real GDP growth contribution
DD_FX_C = 15  # exchange rate depreciation contribution
DD_RES = 16  # residual
DD_INT = 19  # memo: interest payments (% GDP)
DD_AMO = 20  # memo: amortization (% GDP)

ws = new_sheet("Debt Dynamics")
set_col_widths(ws)
title_block(ws, "Debt Dynamics", "Baseline public debt path and decomposition of annual changes")
year_header_row(ws, 5)


def _prev_debt(i, pcl):
    return INIT_DEBT_REF if i == 0 else f"{pcl}{DD_DEBT}"


section_row(ws, 8, "BASELINE DEBT PATH (% OF GDP)")
label_cell(ws, DD_DEBT, "Public debt")
fill_formula_row(
    ws,
    DD_DEBT,
    lambda i, cl, pcl: (
        f"={_prev_debt(i, pcl)}"
        f"*(1+Assumptions!{cl}{A_IR}+Assumptions!{cl}{A_DEP}*Assumptions!{cl}{A_FX_SHARE})"
        f"/((1+Assumptions!{cl}{A_G})*(1+Assumptions!{cl}{A_PI}))"
        f"-Assumptions!{cl}{A_PB}"
    ),
    kind="bold_formula",
)
label_cell(ws, DD_CHANGE, "Change in public debt")
fill_formula_row(
    ws,
    DD_CHANGE,
    lambda i, cl, pcl: f"={cl}{DD_DEBT}-{_prev_debt(i, pcl)}",
)

section_row(ws, 11, "CONTRIBUTIONS TO CHANGE IN DEBT (PP OF GDP)")
label_cell(ws, DD_PB_C, "Primary balance (- = debt-reducing surplus)")
fill_formula_row(ws, DD_PB_C, lambda i, cl, pcl: f"=-Assumptions!{cl}{A_PB}")
label_cell(ws, DD_IR_C, "Real interest rate")
fill_formula_row(
    ws,
    DD_IR_C,
    lambda i, cl, pcl: (
        f"=(Assumptions!{cl}{A_IR}-Assumptions!{cl}{A_PI}*(1+Assumptions!{cl}{A_G}))"
        f"/(1+Assumptions!{cl}{A_NGDP_G})*{_prev_debt(i, pcl)}"
    ),
)
label_cell(ws, DD_G_C, "Real GDP growth")
fill_formula_row(
    ws,
    DD_G_C,
    lambda i, cl, pcl: (
        f"=-Assumptions!{cl}{A_G}/(1+Assumptions!{cl}{A_NGDP_G})*{_prev_debt(i, pcl)}"
    ),
)
label_cell(ws, DD_FX_C, "Exchange rate depreciation")
fill_formula_row(
    ws,
    DD_FX_C,
    lambda i, cl, pcl: (
        f"=Assumptions!{cl}{A_DEP}*Assumptions!{cl}{A_FX_SHARE}"
        f"/(1+Assumptions!{cl}{A_NGDP_G})*{_prev_debt(i, pcl)}"
    ),
)
label_cell(ws, DD_RES, "Residual")
fill_formula_row(
    ws,
    DD_RES,
    lambda i, cl, pcl: f"={cl}{DD_CHANGE}-({cl}{DD_PB_C}+{cl}{DD_IR_C}+{cl}{DD_G_C}+{cl}{DD_FX_C})",
)

section_row(ws, 18, "MEMO ITEMS (% OF GDP)")
label_cell(ws, DD_INT, "Interest payments")
fill_formula_row(
    ws,
    DD_INT,
    lambda i, cl, pcl: (
        f"=Assumptions!{cl}{A_IR}*{_prev_debt(i, pcl)}/(1+Assumptions!{cl}{A_NGDP_G})"
    ),
)
label_cell(ws, DD_AMO, "Amortization")
fill_formula_row(
    ws,
    DD_AMO,
    lambda i, cl, pcl: (
        f"=Assumptions!{cl}{A_AMORT}*{_prev_debt(i, pcl)}/(1+Assumptions!{cl}{A_NGDP_G})"
    ),
)

print("Debt Dynamics sheet built.")

# =================================================================
# SHEET 4: GROSS FINANCING NEEDS
# =================================================================
GFN = "'Gross Financing Needs'"
GFN_PD = 9  # primary deficit (% GDP)
GFN_INT = 10  # interest payments (% GDP)
GFN_AMO = 11  # amortization (% GDP)
GFN_TOT = 12  # gross financing needs (% GDP)
GFN_LCU = 14  # memo: GFN in LCU bn

ws = new_sheet("Gross Financing Needs")
set_col_widths(ws)
title_block(
    ws,
    "Gross Financing Needs",
    "Annual financing requirement: primary deficit + interest + amortization",
)
year_header_row(ws, 5)

section_row(ws, 8, "GROSS FINANCING NEEDS (% OF GDP)")
label_cell(ws, GFN_PD, "Primary deficit (- = surplus)")
fill_formula_row(ws, GFN_PD, lambda i, cl, pcl: f"=-Assumptions!{cl}{A_PB}", kind="link")
label_cell(ws, GFN_INT, "Interest payments")
fill_formula_row(ws, GFN_INT, lambda i, cl, pcl: f"={DD}!{cl}{DD_INT}", kind="link")
label_cell(ws, GFN_AMO, "Amortization of maturing debt")
fill_formula_row(ws, GFN_AMO, lambda i, cl, pcl: f"={DD}!{cl}{DD_AMO}", kind="link")
label_cell(ws, GFN_TOT, "Gross financing needs")
fill_formula_row(
    ws,
    GFN_TOT,
    lambda i, cl, pcl: f"=SUM({cl}{GFN_PD}:{cl}{GFN_AMO})",
    kind="bold_formula",
)

section_row(ws, 13, "MEMO ITEMS")
label_cell(ws, GFN_LCU, "Gross financing needs (LCU bn)")
fill_formula_row(
    ws,
    GFN_LCU,
    lambda i, cl, pcl: f"={cl}{GFN_TOT}*Assumptions!{cl}{A_NGDP}",
)

print("Gross Financing Needs sheet built.")

# =================================================================
# SHEET 5: SCENARIO ANALYSIS
# =================================================================
SC = "'Scenario Analysis'"
SC_G_SHOCK = 9  # growth shock (pp reduction, fraction)
SC_IR_SHOCK = 10  # interest rate shock (pp increase, fraction)
SC_PB_SHOCK = 11  # primary balance shock (pp deterioration, fraction)
SC_FX_SHOCK = 12  # one-time extra depreciation in first projection year
SC_BASE = 15
SC_GROWTH = 16
SC_IR = 17
SC_PB = 18
SC_COMBINED = 19
SC_FX = 20
PROJ_START_I = YEARS.index(PROJ_YEARS[0])

ws = new_sheet("Scenario Analysis")
set_col_widths(ws)
title_block(
    ws, "Scenario Analysis", "Standardized stress tests applied to projection years (2026-2035)"
)
year_header_row(ws, 5)

section_row(ws, 8, "SHOCK PARAMETERS (APPLIED 2026-2035)")
label_cell(ws, SC_G_SHOCK, "Real GDP growth shock (reduction, pp)")
ws.cell(row=SC_G_SHOCK, column=FIRST_COL, value=0.010)
label_cell(ws, SC_IR_SHOCK, "Interest rate shock (increase, pp)")
ws.cell(row=SC_IR_SHOCK, column=FIRST_COL, value=0.020)
label_cell(ws, SC_PB_SHOCK, "Primary balance shock (deterioration, pp)")
ws.cell(row=SC_PB_SHOCK, column=FIRST_COL, value=0.010)
label_cell(ws, SC_FX_SHOCK, "One-time extra depreciation in 2026")
ws.cell(row=SC_FX_SHOCK, column=FIRST_COL, value=0.150)


def scenario_formula(i, cl, pcl, row, *, d_g="", d_ir="", d_pb="", d_fx_2026=""):
    """Debt recursion mirroring the baseline, with parameter shocks in projection years."""
    if i < PROJ_START_I:
        return f"={DD}!{cl}{DD_DEBT}"
    prev = f"{pcl}{row}"
    g = f"(Assumptions!{cl}{A_G}{d_g})"
    ir = f"(Assumptions!{cl}{A_IR}{d_ir})"
    pb = f"(Assumptions!{cl}{A_PB}{d_pb})"
    dep = f"Assumptions!{cl}{A_DEP}"
    if d_fx_2026 and YEARS[i] == PROJ_YEARS[0]:
        dep = f"(Assumptions!{cl}{A_DEP}{d_fx_2026})"
    return (
        f"={prev}*(1+{ir}+{dep}*Assumptions!{cl}{A_FX_SHARE})"
        f"/((1+{g})*(1+Assumptions!{cl}{A_PI}))-{pb}"
    )


section_row(ws, 14, "PUBLIC DEBT PATHS UNDER SCENARIOS (% OF GDP)")
label_cell(ws, SC_BASE, "Baseline")
fill_formula_row(ws, SC_BASE, lambda i, cl, pcl: f"={DD}!{cl}{DD_DEBT}", kind="link")
label_cell(ws, SC_GROWTH, "Growth shock")
fill_formula_row(
    ws,
    SC_GROWTH,
    lambda i, cl, pcl: scenario_formula(i, cl, pcl, SC_GROWTH, d_g=f"-$B${SC_G_SHOCK}"),
)
label_cell(ws, SC_IR, "Interest rate shock")
fill_formula_row(
    ws,
    SC_IR,
    lambda i, cl, pcl: scenario_formula(i, cl, pcl, SC_IR, d_ir=f"+$B${SC_IR_SHOCK}"),
)
label_cell(ws, SC_PB, "Primary balance shock")
fill_formula_row(
    ws,
    SC_PB,
    lambda i, cl, pcl: scenario_formula(i, cl, pcl, SC_PB, d_pb=f"-$B${SC_PB_SHOCK}"),
)
label_cell(ws, SC_COMBINED, "Combined shock")
fill_formula_row(
    ws,
    SC_COMBINED,
    lambda i, cl, pcl: scenario_formula(
        i,
        cl,
        pcl,
        SC_COMBINED,
        d_g=f"-$B${SC_G_SHOCK}",
        d_ir=f"+$B${SC_IR_SHOCK}",
        d_pb=f"-$B${SC_PB_SHOCK}",
    ),
)
label_cell(ws, SC_FX, "Exchange rate shock")
fill_formula_row(
    ws,
    SC_FX,
    lambda i, cl, pcl: scenario_formula(i, cl, pcl, SC_FX, d_fx_2026=f"+$B${SC_FX_SHOCK}"),
)

print("Scenario Analysis sheet built.")

# =================================================================
# SHEET 6: SUSTAINABILITY INDICATORS
# =================================================================
SI = "'Sustainability Indicators'"
SI_DEBT_THRESH = 9  # debt threshold (single cell, column B)
SI_GFN_THRESH = 10  # GFN threshold (single cell, column B)
SI_DEBT = 13
SI_DEBT_T_ROW = 14
SI_DEBT_FLAG = 15
SI_GFN = 16
SI_GFN_FLAG = 17
SI_STRESS_MAX = 20
SI_STRESS_FLAG = 21

ws = new_sheet("Sustainability Indicators")
set_col_widths(ws)
title_block(
    ws, "Sustainability Indicators", "Baseline and stressed indicators vs. standard thresholds"
)
year_header_row(ws, 5)

section_row(ws, 8, "THRESHOLDS (EDITABLE)")
label_cell(ws, SI_DEBT_THRESH, "Public debt threshold (% GDP)")
ws.cell(row=SI_DEBT_THRESH, column=FIRST_COL, value=0.70)
label_cell(ws, SI_GFN_THRESH, "Gross financing needs threshold (% GDP)")
ws.cell(row=SI_GFN_THRESH, column=FIRST_COL, value=0.15)

section_row(ws, 12, "BASELINE INDICATORS")
label_cell(ws, SI_DEBT, "Public debt (% GDP)")
fill_formula_row(ws, SI_DEBT, lambda i, cl, pcl: f"={DD}!{cl}{DD_DEBT}", kind="link")
label_cell(ws, SI_DEBT_T_ROW, "Debt threshold")
fill_formula_row(ws, SI_DEBT_T_ROW, lambda i, cl, pcl: f"=$B${SI_DEBT_THRESH}")
label_cell(ws, SI_DEBT_FLAG, "Debt flag")
fill_formula_row(
    ws,
    SI_DEBT_FLAG,
    lambda i, cl, pcl: f'=IF({cl}{SI_DEBT}>$B${SI_DEBT_THRESH},"BREACH","OK")',
)
label_cell(ws, SI_GFN, "Gross financing needs (% GDP)")
fill_formula_row(ws, SI_GFN, lambda i, cl, pcl: f"={GFN}!{cl}{GFN_TOT}", kind="link")
label_cell(ws, SI_GFN_FLAG, "GFN flag")
fill_formula_row(
    ws,
    SI_GFN_FLAG,
    lambda i, cl, pcl: f'=IF({cl}{SI_GFN}>$B${SI_GFN_THRESH},"BREACH","OK")',
)

section_row(ws, 19, "STRESS TESTS")
label_cell(ws, SI_STRESS_MAX, "Maximum debt across shock scenarios (% GDP)")
fill_formula_row(
    ws,
    SI_STRESS_MAX,
    lambda i, cl, pcl: f"=MAX({SC}!{cl}{SC_GROWTH}:{cl}{SC_FX})",
)
label_cell(ws, SI_STRESS_FLAG, "Stress flag")
fill_formula_row(
    ws,
    SI_STRESS_FLAG,
    lambda i, cl, pcl: f'=IF({cl}{SI_STRESS_MAX}>$B${SI_DEBT_THRESH},"BREACH","OK")',
)

print("Sustainability Indicators sheet built.")

# =================================================================
# SHEET 7: DASHBOARD
# =================================================================
EST_COL_L = col_letter(FIRST_COL + YEARS.index(EST_YEAR))

ws = new_sheet("Dashboard")
set_col_widths(ws, extra={"B": 16})
title_block(
    ws, "Dashboard", "Headline results (see source tabs for full detail)", last_col_letter="D"
)

section_row(ws, 5, "HEADLINE RESULTS", last_col_letter="D")
dashboard_rows = [
    (f"Public debt {EST_YEAR}, baseline (% GDP)", f"={DD}!{EST_COL_L}{DD_DEBT}"),
    (f"Public debt {YEARS[-1]}, baseline (% GDP)", f"={DD}!{LAST_COL_L}{DD_DEBT}"),
    (
        "Peak debt, baseline (% GDP)",
        f"=MAX({DD}!{FIRST_COL_L}{DD_DEBT}:{LAST_COL_L}{DD_DEBT})",
    ),
    (
        "Peak debt under stress (% GDP)",
        f"=MAX({SI}!{FIRST_COL_L}{SI_STRESS_MAX}:{LAST_COL_L}{SI_STRESS_MAX})",
    ),
    (
        f"Average GFN {PROJ_YEARS[0]}-{PROJ_YEARS[-1]} (% GDP)",
        f"=AVERAGE({GFN}!{PROJ_START_L}{GFN_TOT}:{LAST_COL_L}{GFN_TOT})",
    ),
    (
        "Years baseline debt above threshold",
        f'=COUNTIF({SI}!{FIRST_COL_L}{SI_DEBT_FLAG}:{LAST_COL_L}{SI_DEBT_FLAG},"BREACH")',
    ),
    (
        "Years GFN above threshold",
        f'=COUNTIF({SI}!{FIRST_COL_L}{SI_GFN_FLAG}:{LAST_COL_L}{SI_GFN_FLAG},"BREACH")',
    ),
]
r = 6
for label, formula in dashboard_rows:
    ws.cell(row=r, column=1, value=label)
    ws.cell(row=r, column=2, value=formula)
    r += 1

DASH_DEBT_BREACHES = 11  # row of "Years baseline debt above threshold"
DASH_PEAK_STRESS = 9  # row of "Peak debt under stress"
ws.cell(row=r, column=1, value="Overall risk signal")
ws.cell(
    row=r,
    column=2,
    value=(
        f'=IF(B{DASH_DEBT_BREACHES}>0,"HIGH",'
        f'IF(B{DASH_PEAK_STRESS}>{SI}!$B${SI_DEBT_THRESH},"MODERATE","LOW"))'
    ),
)
r += 1

print("Dashboard sheet built.")

wb.save("dsa_model.xlsx")
print("Workbook saved to dsa_model.xlsx with sheets:", wb.sheetnames)
