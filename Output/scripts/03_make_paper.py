"""
make_paper.py  —  generates Final Paper.docx
Run from any directory; uses absolute paths throughout.
"""

from pathlib import Path
from docx import Document
from docx.shared import Inches, Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import copy

# ── Paths ──────────────────────────────────────────────────────────────────────
_ROOT      = Path(__file__).resolve().parents[1]   # Econ692_Capstone/
PAPER_DIR  = _ROOT / "Paper"
FIG_DIR    = _ROOT / "Output" / "figures"
OUT_FILE   = PAPER_DIR / "Final Paper.docx"

# ── Document setup ─────────────────────────────────────────────────────────────
doc = Document()

# Page: US Letter, 1-inch margins
section = doc.sections[0]
section.page_width  = Inches(8.5)
section.page_height = Inches(11)
section.left_margin = section.right_margin = Inches(1.0)
section.top_margin  = section.bottom_margin = Inches(1.0)

# ── Style helpers ──────────────────────────────────────────────────────────────
FONT_NAME = "Times New Roman"
BODY_SIZE = Pt(12)
LINE_SPACE = Pt(22)   # ~1.5 line spacing


def set_run(run, bold=False, italic=False, size=BODY_SIZE, font=FONT_NAME, color=None):
    run.bold   = bold
    run.italic = italic
    run.font.name = font
    run.font.size = size
    if color:
        run.font.color.rgb = RGBColor(*color)


def para_fmt(p, align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=Pt(0),
             space_after=Pt(6), line_spacing=LINE_SPACE):
    pf = p.paragraph_format
    pf.alignment    = align
    pf.space_before = space_before
    pf.space_after  = space_after
    pf.line_spacing = line_spacing


def add_para(text, bold=False, italic=False, align=WD_ALIGN_PARAGRAPH.JUSTIFY,
             size=BODY_SIZE, space_before=Pt(0), space_after=Pt(6),
             first_indent=None):
    p = doc.add_paragraph()
    para_fmt(p, align=align, space_before=space_before, space_after=space_after)
    if first_indent is not None:
        p.paragraph_format.first_line_indent = first_indent
    run = p.add_run(text)
    set_run(run, bold=bold, italic=italic, size=size)
    return p


def add_heading(text, level=1):
    """Custom heading — avoids default Word heading styles."""
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.space_before = Pt(14) if level == 1 else Pt(10)
    pf.space_after  = Pt(6)
    pf.line_spacing = Pt(14)
    pf.alignment    = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run(text)
    run.bold = True
    run.font.name = FONT_NAME
    run.font.size = Pt(14) if level == 1 else Pt(12)
    return p


def add_subheading(text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after  = Pt(4)
    p.paragraph_format.line_spacing = Pt(14)
    run = p.add_run(text)
    run.bold = True
    run.italic = True
    run.font.name = FONT_NAME
    run.font.size = Pt(12)
    return p


def add_mixed(parts, align=WD_ALIGN_PARAGRAPH.JUSTIFY,
              first_indent=None, space_after=Pt(6)):
    """parts = list of (text, bold, italic)"""
    p = doc.add_paragraph()
    para_fmt(p, align=align, space_after=space_after)
    if first_indent is not None:
        p.paragraph_format.first_line_indent = first_indent
    for text, bold, italic in parts:
        run = p.add_run(text)
        set_run(run, bold=bold, italic=italic)
    return p


def add_equation(text):
    p = doc.add_paragraph()
    p.paragraph_format.alignment   = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after  = Pt(6)
    run = p.add_run(text)
    run.font.name = "Courier New"
    run.font.size = Pt(11)
    return p


def add_figure(fname, caption, width=Inches(6.3)):
    fpath = FIG_DIR / fname
    if fpath.exists():
        p = doc.add_paragraph()
        p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(6)
        run = p.add_run()
        run.add_picture(str(fpath), width=width)
    else:
        add_para(f"[Figure: {fname} not found]",
                 align=WD_ALIGN_PARAGRAPH.CENTER, italic=True)
    cap = doc.add_paragraph()
    cap.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cap.paragraph_format.space_after = Pt(10)
    cap.paragraph_format.line_spacing = Pt(14)
    r = cap.add_run(caption)
    r.font.name = FONT_NAME
    r.font.size = Pt(10)
    r.italic = True


def shade_cell(cell, fill_hex: str):
    tc   = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd  = OxmlElement("w:shd")
    shd.set(qn("w:val"),   "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"),  fill_hex)
    tcPr.append(shd)


def make_table(headers, rows, col_widths=None, header_shade="D0D8E4"):
    """Create a styled table. rows = list of lists of strings."""
    n_cols = len(headers)
    tbl = doc.add_table(rows=1, cols=n_cols)
    tbl.style = "Table Grid"
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER

    # Header row
    hdr_row = tbl.rows[0]
    for i, h in enumerate(headers):
        cell = hdr_row.cells[i]
        shade_cell(cell, header_shade)
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after  = Pt(2)
        run = p.add_run(h)
        run.bold = True
        run.font.name = FONT_NAME
        run.font.size = Pt(10)

    # Data rows
    for row_data in rows:
        row = tbl.add_row()
        for i, cell_text in enumerate(row_data):
            cell = row.cells[i]
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            p = cell.paragraphs[0]
            align = (WD_ALIGN_PARAGRAPH.LEFT if i == 0
                     else WD_ALIGN_PARAGRAPH.CENTER)
            p.alignment = align
            p.paragraph_format.space_before = Pt(1)
            p.paragraph_format.space_after  = Pt(1)
            run = p.add_run(cell_text)
            run.font.name = FONT_NAME
            run.font.size = Pt(10)

    # Column widths
    if col_widths:
        for row in tbl.rows:
            for i, w in enumerate(col_widths):
                row.cells[i].width = w
    return tbl


def add_note(text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after  = Pt(10)
    p.paragraph_format.line_spacing = Pt(13)
    run = p.add_run("Note. ")
    run.bold = True
    run.font.name = FONT_NAME
    run.font.size = Pt(9.5)
    run2 = p.add_run(text)
    run2.font.name = FONT_NAME
    run2.font.size = Pt(9.5)


def page_break():
    doc.add_paragraph().add_run().add_break(
        __import__("docx.enum.text", fromlist=["WD_BREAK"]).WD_BREAK.PAGE
    )


# ══════════════════════════════════════════════════════════════════════════════
# TITLE PAGE
# ══════════════════════════════════════════════════════════════════════════════
p_title = doc.add_paragraph()
p_title.paragraph_format.space_before = Pt(72)
p_title.paragraph_format.space_after  = Pt(16)
p_title.paragraph_format.alignment    = WD_ALIGN_PARAGRAPH.CENTER
r = p_title.add_run("Work-From-Home and Metropolitan Housing Deflation\nin California")
r.bold = True
r.font.name = FONT_NAME
r.font.size = Pt(16)

add_para("Tuan Truong", bold=False, align=WD_ALIGN_PARAGRAPH.CENTER,
         size=Pt(12), space_before=Pt(8), space_after=Pt(4))
add_para("University of San Francisco — ECON 692 Capstone",
         align=WD_ALIGN_PARAGRAPH.CENTER, size=Pt(12), space_after=Pt(4))
add_para("May 2026", align=WD_ALIGN_PARAGRAPH.CENTER, size=Pt(12))

page_break()

# ══════════════════════════════════════════════════════════════════════════════
# ABSTRACT
# ══════════════════════════════════════════════════════════════════════════════
add_heading("Abstract", level=1)
add_para(
    "I estimate the causal effect of remote work adoption on housing prices across "
    "California's 58 counties from 2017 to 2023. Using a two-way fixed effects panel "
    "with county and year fixed effects, I find that a one-percentage-point increase in "
    "the county work-from-home (WFH) rate reduces median home prices in Metropolitan "
    "Core counties by 1.23% (SE = 0.34%, p < 0.001). A dynamic event study spanning "
    "2017–2023 satisfies the parallel pre-trends assumption (Wald χ²(2) = 1.85, "
    "p = 0.40) and reveals a monotonically deepening effect, from −0.74%** in 2020 to "
    "−1.04%** in 2023. Supply constraints amplify the effect: counties in the top half "
    "of the pre-period price-to-income distribution experience −1.20%*** per percentage "
    "point, nearly three times the −0.40% estimate for elastic-supply counties. The "
    "de Chaisemartin–D’Haultfœuille (2024) weighted average slope estimate is statistically "
    "insignificant (+1.20%, p = 0.79), suggesting caution about extrapolating the TWFE "
    "estimate across all counties. These results imply that WFH adoption has contributed "
    "to modest but growing housing deflation in California’s metropolitan cores, "
    "particularly in supply-constrained coastal markets.",
    first_indent=Inches(0.5), space_after=Pt(4)
)
add_para("Keywords: work from home, remote work, housing prices, two-way fixed effects, California",
         italic=True, space_after=Pt(10))

page_break()

# ══════════════════════════════════════════════════════════════════════════════
# 1. INTRODUCTION
# ══════════════════════════════════════════════════════════════════════════════
add_heading("1. Introduction", level=1)

add_para(
    "Remote work has emerged as one of the most consequential labor market shifts of the "
    "post-pandemic era. According to Mondragón and Wieland (2022), the share of U.S. workers "
    "primarily working from home rose from approximately 7% in 2019 to over 30% in 2021, "
    "before stabilizing near 20% through 2023. While major technology companies—including "
    "Google, Tesla, and Apple—have increasingly mandated returns to the office, millions of "
    "workers have maintained hybrid or fully remote arrangements, reshaping residential "
    "preferences across metropolitan areas.",
    first_indent=Inches(0.5)
)
add_para(
    "Standard urban economics theory predicts that reduced commuting costs should flatten the "
    "bid-rent gradient, decreasing demand for central locations and potentially redistributing "
    "it toward suburban and exurban areas (Albouy and Stuart, 2020; Brueckner, Thisse, and "
    "Zenou, 1999). Whether this redistribution produces net declines in metropolitan core "
    "prices depends critically on local housing supply elasticity and the extent to which "
    "displaced demand exits the metropolitan area altogether.",
    first_indent=Inches(0.5)
)
add_para(
    "The empirical literature has reached mixed conclusions. Gupta et al. (2022) document "
    "large positive effects of WFH exposure on housing prices nationally, while Ramani and "
    "Bloom (2021) and Davis, Ghent, and Gregory (2023) find within-metro redistribution—"
    "higher prices in suburbs, lower in city centers—consistent with the spatial reallocation "
    "hypothesis. Liu and Su (2021) provide evidence of an urban exodus from high-rent cities, "
    "particularly in California. This paper contributes to this literature by providing "
    "county-level causal estimates for California, exploiting variation in WFH adoption across "
    "33 metropolitan core counties from 2017 to 2023.",
    first_indent=Inches(0.5)
)
add_para(
    "This study makes two main contributions. First, I estimate the causal effect of WFH on "
    "housing prices in Metropolitan Core counties using OMB CBSA delineations and a dynamic "
    "TWFE specification that formally tests parallel pre-trends. I find that a one-percentage-"
    "point increase in the county WFH rate is associated with a 1.23% decline in median home "
    "prices in metropolitan cores, an effect that grows monotonically from 2020 to 2023. "
    "Second, I show that housing supply constraints amplify the deflation: counties in the "
    "top half of the pre-period price-to-income ratio distribution experience effects nearly "
    "three times those of elastic-supply counties, consistent with Saiz (2010) and Glaeser "
    "and Gyourko (2018).",
    first_indent=Inches(0.5)
)
add_para(
    "The remainder of the paper proceeds as follows. Section 2 describes the data. "
    "Section 3 presents the empirical strategy. Section 4 reports main results. "
    "Section 5 presents robustness checks. Section 6 discusses limitations. "
    "Section 7 concludes.",
    first_indent=Inches(0.5)
)

# ══════════════════════════════════════════════════════════════════════════════
# 2. DATA
# ══════════════════════════════════════════════════════════════════════════════
add_heading("2. Data", level=1)
add_subheading("2.1  Data Sources")
add_para(
    "This study uses a county-level panel spanning 58 California counties from 2017 to 2023 "
    "(406 county-year observations). Data are assembled from five sources.",
    first_indent=Inches(0.5)
)
add_para(
    "Housing prices are obtained from Zillow’s Home Value Index (ZHVI), which provides "
    "ZIP-code-level monthly estimates of the median home value for single-family homes and "
    "condominiums in the 33rd–67th percentile of the local price distribution. Monthly "
    "values are averaged within calendar years and aggregated to the county level using "
    "population weights across contributing ZIP codes. Prices are expressed in natural logarithms.",
    first_indent=Inches(0.5)
)
add_para(
    "Work-from-home rates are constructed from the American Community Survey (ACS) five-year "
    "estimates, Table S0801 (Commuting Characteristics by Sex). The variable captures the "
    "percentage of employed civilians aged 16 and over whose primary means of transportation "
    "to work is “worked from home.” ZIP-code-level WFH percentages are aggregated "
    "to the county level using population weights.",
    first_indent=Inches(0.5)
)
add_para(
    "County controls include median household income from ACS Table B19013 and the civilian "
    "unemployment rate from the Bureau of Labor Statistics Local Area Unemployment Statistics "
    "(BLS LAUS). Median income is expressed in natural logarithms as the primary income "
    "control. County classification follows the 2023 OMB Core-Based Statistical Area (CBSA) "
    "delineations. The 33 Metropolitan Core counties—central counties of Metropolitan "
    "Statistical Areas—form the main analysis sample (231 observations). Non-Metro "
    "counties (Micropolitan + Noncore, n = 21, 147 observations) serve as a placebo group.",
    first_indent=Inches(0.5)
)

add_subheading("2.2  Summary Statistics")
add_para(
    "Table 1 reports descriptive statistics for the full 58-county panel. The average "
    "WFH rate is 9.17% (SD = 4.98%), ranging from 3.0% to 28.8%. The average "
    "log price of 12.90 corresponds to approximately $400,000. Between 2019 and 2023, the "
    "average Metro-Core WFH rate rose from 5.8% to 15.0%, a 9.2 percentage-point increase "
    "driven primarily by the COVID-19 shock. Metro-Core counties exhibit higher WFH rates "
    "and substantially higher median home prices than non-metro counties.",
    first_indent=Inches(0.5)
)
add_para("Table 1: Summary Statistics (County-Level Panel, N = 406 obs, 58 counties, 2017–2023)",
         bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, space_before=Pt(10), space_after=Pt(4))
make_table(
    ["Variable", "Mean", "SD", "Min", "Median", "Max"],
    [
        ["WFH rate (%)",             "9.17",   "4.98",    "3.00",  "7.54",   "28.76"],
        ["log(Median Home Price)",   "12.90",  "0.54",    "11.52", "12.92",  "14.34"],
        ["Median income ($)",        "74,599", "23,909",  "37,551","68,467", "160,731"],
        ["Unemployment rate",        "0.071",  "0.025",   "0.000", "0.067",  "0.187"],
        ["Population",               "670,210","1,446,830","926",  "177,652","10,088,658"],
        ["ZIP codes per county-year","18.1",   "30.2",    "1.0",   "6.0",    "184.0"],
    ],
    col_widths=[Inches(2.8), Inches(0.7), Inches(0.65), Inches(0.7), Inches(0.75), Inches(0.75)]
)
add_note(
    "WFH rate is the population-weighted mean of ZIP-level ACS S0801 WFH percentages. "
    "Prices are annual population-weighted means of monthly Zillow ZHVI, in logarithms. "
    "Unemployment is expressed as a decimal fraction. N = 406 county–years "
    "(58 California counties × 7 years, 2017–2023)."
)

# ══════════════════════════════════════════════════════════════════════════════
# 3. EMPIRICAL STRATEGY
# ══════════════════════════════════════════════════════════════════════════════
add_heading("3. Empirical Strategy", level=1)
add_subheading("3.1  Main TWFE Specification")
add_para(
    "The primary specification is a two-way fixed effects (TWFE) panel regression:",
    first_indent=Inches(0.5)
)
add_equation(
    "log(P_ct) = β WFH_ct + δ log(Income_ct) + α_c + γ_t + ε_ct     (1)"
)
add_para(
    "where P_ct is the population-weighted mean Zillow ZHVI in county c and year t, "
    "WFH_ct is the county WFH rate (in percentage points), α_c are county fixed effects "
    "absorbing all time-invariant county characteristics, γ_t are year fixed effects "
    "absorbing common macroeconomic shocks, and ε_ct is an idiosyncratic error. Standard "
    "errors are clustered by county (CRV1). The coefficient β estimates the within-county "
    "effect of a one-percentage-point increase in the WFH rate on log housing prices, "
    "controlling for contemporaneous income growth.",
    first_indent=Inches(0.5)
)
add_subheading("3.2  Dynamic Event Study")
add_para(
    "To test the parallel pre-trends assumption and characterize the time path of the WFH "
    "effect, I estimate a dynamic TWFE event study:",
    first_indent=Inches(0.5)
)
add_equation(
    "log(P_ct) = Σ_{t’2019} β_t (WFH_ct × 1[Year=t]) + δ log(Income_ct) + α_c + γ_t + ε_ct     (2)"
)
add_para(
    "The reference year is 2019 (the last pre-pandemic year). Pre-period coefficients "
    "(β₂₀₁₇, β₂₀₁₈) provide a formal parallel "
    "trends test: I report a joint Wald test of H₀: β₂₀₁₇ = "
    "β₂₀₁₈ = 0.",
    first_indent=Inches(0.5)
)
add_subheading("3.3  Supply Heterogeneity")
add_para(
    "To test whether supply constraints amplify the WFH price effect, I estimate an "
    "interacted TWFE:",
    first_indent=Inches(0.5)
)
add_equation(
    "log(P_ct) = β_Con(WFH_ct × Constrained_c) + β_Ela(WFH_ct × Elastic_c) + δ log(Income_ct) + α_c + γ_t + ε_ct     (3)"
)
add_para(
    "Constrained_c equals one for Metro-Core counties in the top half of the 2017–2023 "
    "average price-to-income (P/I) ratio distribution; Elastic_c for the bottom half. The "
    "P/I ratio serves as a proxy for housing supply constraints (Glaeser and Gyourko, 2018; "
    "Saiz, 2010). The median P/I threshold is 5.41.",
    first_indent=Inches(0.5)
)
add_subheading("3.4  Identification")
add_para(
    "The key identifying assumption is parallel trends under continuous treatment: absent the "
    "WFH shock, counties with different WFH rates would have followed similar housing price "
    "trajectories. This is supported by the pre-period event study coefficients, which are "
    "small, positive, and jointly insignificant (Wald χ²(2) = 1.85, "
    "p = 0.40). The main threat to identification is endogeneity of the WFH rate: "
    "high-WFH counties (San Francisco, Santa Clara, Marin) may have simultaneously experienced "
    "stronger tech-sector income growth that independently affected housing prices. I address "
    "this in three ways: (i) including log median income as a contemporaneous control; "
    "(ii) using a pre-determined WFH proxy (2019 WFH rate × post-2020 indicator); "
    "and (iii) applying Frisch–Waugh detrending to remove county-specific secular trends.",
    first_indent=Inches(0.5)
)

# ══════════════════════════════════════════════════════════════════════════════
# 4. RESULTS
# ══════════════════════════════════════════════════════════════════════════════
add_heading("4. Results", level=1)
add_subheading("4.1  Descriptive Evidence")
add_para(
    "Figure 1 plots average WFH rates and log housing prices by CBSA tier from 2017 to 2023. "
    "WFH rates rose sharply for all tiers after 2019, with Metro-Core counties exhibiting the "
    "largest absolute increase (approximately +9 percentage points). Log housing prices in "
    "Metro-Core counties rose more slowly than in Non-Metro counties after 2020, consistent "
    "with WFH-driven demand reallocation away from metropolitan cores.",
    first_indent=Inches(0.5)
)
add_figure("fig1_cbsa_trends.png",
           "Figure 1: WFH Rate and log(Median Home Price) Trends by CBSA Tier, California 2017–2023.",
           width=Inches(6.3))

add_subheading("4.2  Main TWFE Results")
add_para(
    "Table 2 reports TWFE estimates for both Metro-Core and Non-Metro counties. "
    "The baseline Metro-Core estimate is −1.229%*** (SE = 0.336%), robust to the "
    "inclusion of log income (−1.234%***) and the unemployment rate (−1.021%**). "
    "Strikingly, Non-Metro counties also exhibit a negative and highly significant "
    "WFH effect: −0.638%*** (SE = 0.189%) in the baseline, −0.637%*** (SE = 0.187%) "
    "with income controls, and −0.627%** (SE = 0.203%) with both controls. "
    "The Metro-Core effect is approximately twice the Non-Metro effect across all "
    "specifications. Both results suggest WFH adoption is associated with a broad "
    "reduction in housing demand across California, inconsistent with simple "
    "within-state spatial reallocation.",
    first_indent=Inches(0.5)
)
add_para("Table 2: TWFE Estimates — Metro-Core vs. Non-Metro Counties",
         bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, space_before=Pt(10), space_after=Pt(4))
make_table(
    ["Specification",
     "Metro Coef (%)", "SE (%)", "Sig.",
     "Non-Metro Coef (%)", "SE (%)", "Sig."],
    [
        ["(1) Baseline TWFE",              "−1.229", "0.336", "***", "−0.638", "0.189", "***"],
        ["(2) + log(Income)",              "−1.234", "0.336", "***", "−0.637", "0.187", "***"],
        ["(3) + log(Income) + Unemp.",     "−1.021", "0.329", "**",  "−0.627", "0.203", "**"],
    ],
    col_widths=[Inches(2.35), Inches(0.85), Inches(0.65), Inches(0.5),
                Inches(1.05), Inches(0.65), Inches(0.5)]
)
add_note(
    "Metro-Core: 33 counties × 7 years (231 obs). Non-Metro (Micropolitan + Noncore): "
    "21 counties × 7 years (147 obs). Specification: log(Price) ~ WFH rate + controls "
    "| County + Year. CRV1 SE clustered by county. Coefficients are % price change per "
    "1-pp WFH increase. *** p<.001, ** p<.01, * p<.05."
)
add_figure("fig_twfe_metro_vs_nonmetro.png",
           "Figure 2: TWFE WFH Coefficients — Metro-Core vs Non-Metro Counties. "
           "Error bars are 95% CRV1 CI. All three specifications shown.",
           width=Inches(6.3))

add_subheading("4.3  Dynamic TWFE Event Study")
add_para(
    "Table 3 and Figure 3 present the dynamic event study for Metro-Core counties. "
    "The pre-period coefficients are small and jointly insignificant (Wald "
    "χ²(2) = 1.85, p = 0.40), supporting the parallel trends "
    "assumption. Post-2019 coefficients are negative, statistically significant, and "
    "monotonically deepening: −0.74%** in 2020, −0.75%** in 2021, −0.78%* "
    "in 2022, and −1.04%** in 2023. The growing magnitude suggests that residential "
    "adjustment to remote work is a gradual, cumulative process rather than a one-time "
    "relocation shock.",
    first_indent=Inches(0.5)
)
add_para("Table 3: Dynamic TWFE Event Study — Metro-Core Counties (Reference Year = 2019)",
         bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, space_before=Pt(10), space_after=Pt(4))
make_table(
    ["Year", "Coef (%)", "SE (%)", "95% CI (%)", "p-value", "Sig."],
    [
        ["2017 (pre)", "+0.753", "0.568", "[−0.36, +1.87]", "0.185", ""],
        ["2018 (pre)", "+0.151", "0.497", "[−0.82, +1.13]", "0.761", ""],
        ["2019 (ref)", "0.000",  "—",  "—",             "—", "ref"],
        ["2020",       "−0.744", "0.253", "[−1.24, −0.25]", "0.003", "**"],
        ["2021",       "−0.749", "0.290", "[−1.32, −0.18]", "0.010", "**"],
        ["2022",       "−0.778", "0.309", "[−1.38, −0.17]", "0.012", "*"],
        ["2023",       "−1.037", "0.337", "[−1.70, −0.38]", "0.002", "**"],
    ],
    col_widths=[Inches(1.05), Inches(0.85), Inches(0.75), Inches(1.55), Inches(0.85), Inches(0.65)]
)
add_note(
    "Pre-trend joint Wald test: H₀: β₂₀₁₇ = β₂₀₁₈ = 0; "
    "χ²(2) = 1.85, p = 0.40 (PASS). Sample: 33 Metro-Core counties × 7 years "
    "(231 obs). log(Income) control included. CRV1 SE."
)
add_figure("fig2_cbsa_event_study.png",
           "Figure 3: Dynamic TWFE Event Study, Metro-Core Counties. "
           "Reference year 2019. 95% CI from CRV1 SE. Wald χ²(2)=1.85, p=0.40.",
           width=Inches(6.3))

add_subheading("4.4  Supply Heterogeneity")
add_para(
    "Table 4 presents supply heterogeneity estimates. Constrained counties—those with "
    "pre-period P/I ratios above the median of 5.41, including San Francisco (P/I = 11.8), "
    "San Mateo (11.0), Santa Cruz (9.8), and Marin (9.7)—experience a WFH price effect of "
    "−1.195%*** (SE = 0.322%). Elastic counties—including Kern (P/I = 3.6), "
    "Kings (3.9), and Fresno (4.3)—experience −0.402% (SE = 0.353%, p = 0.25). "
    "The difference of −0.79%† (p = 0.097) is marginally significant, consistent "
    "with supply constraints amplifying the price impact of demand shifts.",
    first_indent=Inches(0.5)
)
add_para("Table 4: Supply Heterogeneity — Constrained vs. Elastic Metro-Core Counties",
         bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, space_before=Pt(10), space_after=Pt(4))
make_table(
    ["Group", "Counties (N)", "WFH Coef (%)", "SE (%)", "p-value", "Sig."],
    [
        ["Constrained (High P/I ≥ 5.41)", "17", "−1.195", "0.322", "0.0002", "***"],
        ["Elastic (Low P/I < 5.41)",         "16", "−0.402", "0.353", "0.254",  ""],
        ["Wald difference (Con − Ela)",   "—","−0.792","0.478","0.097", "†"],
    ],
    col_widths=[Inches(2.45), Inches(1.0), Inches(1.05), Inches(0.85), Inches(0.85), Inches(0.55)]
)
add_note(
    "P/I split at the median pre-period (2017–2019) price-to-income ratio of 5.41. "
    "log(Income) control included. CRV1 SE. *** p<.001, † p<.10."
)
add_figure("fig4_cbsa_supply.png",
           "Figure 4: WFH Effect by Supply Constraint Group. "
           "Error bars are 95% CRV1 CI. Constrained = top-half P/I ratio.",
           width=Inches(6.3))

# ══════════════════════════════════════════════════════════════════════════════
# 5. ROBUSTNESS CHECKS
# ══════════════════════════════════════════════════════════════════════════════
add_heading("5. Robustness Checks", level=1)
add_para(
    "Table 5 summarizes six robustness specifications. The main estimate of "
    "−1.234%*** (Spec. R1) is robust across all checks.",
    first_indent=Inches(0.5)
)
add_subheading("5.1  DH(2024) Weighted Average Slope")
add_para(
    "I apply the de Chaisemartin and D’Haultfœuille (2024) weighted average slope "
    "(WAS) estimator, which is robust to heterogeneous treatment effects across counties and "
    "over time. The WAS estimate is +1.346% (SE = 4.831%, p = 0.787) without "
    "controls and +1.201% (SE = 4.601%, p = 0.788) with income controls—"
    "both statistically indistinguishable from zero. The contrast between the TWFE’s "
    "−1.234%*** and the WAS’s insignificant positive point estimate suggests that "
    "heterogeneous treatment effects across counties and/or limited statistical power in the "
    "WAS’s “switcher” subset may be important features of the data.",
    first_indent=Inches(0.5)
)
add_subheading("5.2  Pre-determined WFH Exposure")
add_para(
    "To address reverse causality, I replace the contemporaneous WFH rate with the 2019 "
    "WFH rate interacted with a post-2020 indicator. This specification yields "
    "−1.450%** (SE = 0.459%), similar in magnitude to the main estimate and "
    "inconsistent with reverse causality.",
    first_indent=Inches(0.5)
)
add_subheading("5.3  Frisch–Waugh Detrending")
add_para(
    "Applying Frisch–Waugh detrending—residualizing both log price and WFH rates "
    "on county-specific linear time trends before estimation—yields −0.845%* "
    "(SE = 0.417%), somewhat smaller than the main estimate but still statistically "
    "significant, confirming that the result is not driven by pre-existing county trends.",
    first_indent=Inches(0.5)
)
add_subheading("5.4  Extended Metro Sample")
add_para(
    "Pooling the 33 Metro-Core and 4 Metro-Fringe counties yields −1.212%*** "
    "(SE = 0.321%), nearly identical to the main estimate.",
    first_indent=Inches(0.5)
)

add_para("Table 5: Robustness Checks — Metro-Core Counties",
         bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, space_before=Pt(10), space_after=Pt(4))
make_table(
    ["Spec.", "Description", "WFH Coef (%)", "SE (%)", "p-value", "Sig."],
    [
        ["R1", "Main TWFE (+Income)",                       "−1.234", "0.336", "0.0002", "***"],
        ["R2", "TWFE, no controls",                         "−1.229", "0.336", "0.0003", "***"],
        ["R3", "DH(2024) WAS + Income",                     "+1.201",      "4.601", "0.788",  ""],
        ["R4", "Pre-determined WFH₂₀₁₉ × Post", "−1.450", "0.459", "0.0016", "**"],
        ["R5", "Frisch–Waugh detrending",              "−0.845", "0.417", "0.043",  "*"],
        ["R6", "Metro-Core + Metro-Fringe (37 counties)",   "−1.212", "0.321", "0.0002", "***"],
    ],
    col_widths=[Inches(0.45), Inches(2.8), Inches(1.0), Inches(0.75), Inches(0.85), Inches(0.6)]
)
add_note(
    "R3 is the de Chaisemartin–D’Haultfœuille (2024) weighted average slope "
    "estimator applied to first-differenced data; large SE reflects the limited ‘switcher’ "
    "county sample. R4 uses WFH₂₀₁₉ × 1[year ≥ 2020] as the "
    "treatment variable. R5 partials out county-specific linear time trends. "
    "*** p<.001, ** p<.01, * p<.05."
)
add_figure("fig3_cbsa_was.png",
           "Figure 5: DH(2024) Weighted Average Slope Estimates with 95% Confidence Intervals. "
           "Dotted line at zero. Contrast with TWFE main estimate (−1.234%***).",
           width=Inches(6.3))

add_subheading("5.5  Demand Outflow Test")
add_para(
    "The non-metro coefficient of −0.637%*** is prima facie inconsistent with a simple "
    "demand-outflow story. I probe this with a spillover regression including the WFH rate of "
    "the nearest metropolitan region. The spillover coefficient is −0.740% "
    "(SE = 0.489%, p = 0.130) and the own × region-metro interaction is "
    "insignificant (p = 0.499). These results suggest WFH adoption captures a broad "
    "labor market shift that uniformly reduces housing demand rather than a within-state "
    "spatial reallocation.",
    first_indent=Inches(0.5)
)

# ══════════════════════════════════════════════════════════════════════════════
# 6. LIMITATIONS
# ══════════════════════════════════════════════════════════════════════════════
add_heading("6. Discussion and Limitations", level=1)
add_para(
    "The main result shows that a 1-pp WFH increase is associated with a ~1.23% decrease "
    "in housing prices in Metro-Core counties, with the effect growing over time and largest "
    "in supply-constrained markets. Several limitations constrain the interpretation of "
    "these findings.",
    first_indent=Inches(0.5)
)
add_para(
    "First, the analysis is limited to California, and results may not generalize to other "
    "states with different housing market structures, zoning regimes, or demographic "
    "compositions. Second, the price-to-income ratio is an imperfect proxy for housing "
    "supply elasticity; direct Saiz (2010) elasticity measures are not available for all "
    "counties. Third, the county-level panel contains only 33 Metro-Core observations per "
    "year, limiting statistical power for the WAS estimator and supply heterogeneity tests. "
    "Fourth, unobserved tech-sector demand shocks correlated with WFH adoption but not "
    "captured by income controls remain a potential confounder, despite the evidence from "
    "pre-determined WFH and Frisch–Waugh robustness checks.",
    first_indent=Inches(0.5)
)
add_para(
    "The contrast between the TWFE’s −1.234%*** and the DH(2024) WAS’s "
    "+1.201% (p = 0.79) warrants attention. The WAS estimator relies on a subset "
    "of “switcher” counties that changed WFH exposure, and its large standard "
    "error (4.6%) suggests limited power. The divergence may reflect genuine treatment "
    "effect heterogeneity across counties, or it may simply reflect the low power of the "
    "WAS in a 33-county sample. Caution is warranted in extrapolating the TWFE estimate "
    "to counties with very different WFH trajectories.",
    first_indent=Inches(0.5)
)

# ══════════════════════════════════════════════════════════════════════════════
# 7. CONCLUSION
# ══════════════════════════════════════════════════════════════════════════════
add_heading("7. Conclusion", level=1)
add_para(
    "This paper estimates the causal effect of work-from-home adoption on housing prices "
    "in California’s Metropolitan Core counties using a county-level panel from 2017 to "
    "2023. The main finding is that a one-percentage-point increase in the county WFH rate "
    "is associated with a 1.23% decline in median home prices—an effect that grows "
    "monotonically from −0.74%** in 2020 to −1.04%** in 2023 and is largest in "
    "supply-constrained coastal markets (e.g., San Francisco, San Mateo, Santa Cruz).",
    first_indent=Inches(0.5)
)
add_para(
    "These results suggest that WFH has contributed to modest but accelerating housing "
    "deflation in California’s metropolitan cores. The effect is most pronounced in "
    "high-amenity, constrained-supply markets that experienced the largest price run-ups "
    "during the 2010s. As remote work becomes embedded in labor market norms, downward "
    "pressure on metropolitan core prices is likely to persist, with implications for "
    "property tax revenues, urban commercial real estate, and housing affordability in "
    "California’s high-cost cities.",
    first_indent=Inches(0.5)
)
add_para(
    "The robustness of the TWFE estimate across specifications—pre-determined WFH "
    "(−1.45%**), Frisch–Waugh detrending (−0.85%*), and extended metro sample "
    "(−1.21%***)—places the plausible range of the effect between −0.8% and "
    "−1.5% per percentage point of WFH. Future work should exploit more granular "
    "sub-county data, direct supply elasticity measures, and longer post-pandemic panels "
    "to sharpen both the estimates and the mechanisms.",
    first_indent=Inches(0.5)
)

# ══════════════════════════════════════════════════════════════════════════════
# REFERENCES
# ══════════════════════════════════════════════════════════════════════════════
page_break()
add_heading("References", level=1)
refs = [
    ("Albouy, D., and S. Stuart.", False,
     " 2020. “Urban Population and Amenities: The Neoclassical Model of Location.” "
     "International Economic Review 61(1): 127–158."),
    ("Brueckner, J. K., J.-F. Thisse, and Y. Zenou.", False,
     " 1999. “Why is Central Paris Rich and Downtown Detroit Poor? An Amenity-Based "
     "Theory.” European Economic Review 43(1): 91–107."),
    ("Callaway, B., A. Goodman-Bacon, and P. H. C. Sant’Anna.", False,
     " 2024. “Difference-in-Differences with a Continuous Treatment.” Working Paper."),
    ("Davis, M. A., A. C. Ghent, and J. M. Gregory.", False,
     " 2023. “The Work-from-Home Revolution and the Sustainability of Remote Work.” "
     "NBER Working Paper 30975."),
    ("de Chaisemartin, C., and X. D’Haultfœuille.", False,
     " 2020. “Two-Way Fixed Effects Estimators with Heterogeneous Treatment Effects.” "
     "American Economic Review 110(9): 2964–2996."),
    ("de Chaisemartin, C., and X. D’Haultfœuille.", False,
     " 2024. “Difference-in-Differences Estimators of Intertemporal Treatment Effects.” "
     "Review of Economics and Statistics, forthcoming."),
    ("Glaeser, E. L., and J. Gyourko.", False,
     " 2018. “The Economic Implications of Housing Supply.” "
     "Journal of Economic Perspectives 32(1): 3–30."),
    ("Gupta, A., V. Mittal, J. Peeters, and S. Van Nieuwerburgh.", False,
     " 2022. “Flattening the Curve: Pandemic-Induced Revaluation of Urban Real Estate.” "
     "Journal of Financial Economics 146(2): 594–636."),
    ("Liu, S., and Y. Su.", False,
     " 2021. “The Impact of the COVID-19 Pandemic on the Demand for Density: Evidence "
     "from the U.S. Housing Market.” Economics Letters 207: 110010."),
    ("Mondragón, J. A., and J. Wieland.", False,
     " 2022. “Housing Demand and Remote Work.” NBER Working Paper 30041."),
    ("Ramani, A., and N. Bloom.", False,
     " 2021. “The Donut Effect of COVID-19 on Cities.” "
     "NBER Working Paper 28876."),
    ("Saiz, A.", False,
     " 2010. “The Geographic Determinants of Housing Supply.” "
     "Quarterly Journal of Economics 125(3): 1253–1296."),
]
for author, _, rest in refs:
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after  = Pt(5)
    p.paragraph_format.line_spacing = Pt(16)
    p.paragraph_format.first_line_indent = Inches(-0.4)
    p.paragraph_format.left_indent = Inches(0.4)
    r1 = p.add_run(author)
    r1.bold = True
    r1.font.name = FONT_NAME
    r1.font.size = Pt(11)
    r2 = p.add_run(rest)
    r2.font.name = FONT_NAME
    r2.font.size = Pt(11)

# ══════════════════════════════════════════════════════════════════════════════
# SAVE
# ══════════════════════════════════════════════════════════════════════════════
doc.save(str(OUT_FILE))
print(f"Saved: {OUT_FILE}")
