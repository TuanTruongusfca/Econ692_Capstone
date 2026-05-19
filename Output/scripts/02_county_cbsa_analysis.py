"""
county_cbsa_analysis.py
========================
Full county-level analysis using OMB CBSA classifications (gold standard).

Classification (2023 OMB delineations):
  Metro-Core    : Central county of a Metropolitan Statistical Area  (33 CA counties)
  Metro-Fringe  : Outlying county of a Metropolitan Statistical Area (4 CA counties)
  Micropolitan  : Central/outlying county of a Micropolitan MSA      (10 CA counties)
  Noncore       : Not in any CBSA — fully rural                      (11 CA counties)

Analyses (all at county-year level):
  1. Main TWFE          — Metro-Core counties, own WFH rate
  2. Dynamic TWFE       — Metro-Core event study, pre-trend test
  3. DH(2024) WAS       — County-level WAS with wild bootstrap SE
  4. Supply Heterogen.  — Metro-Core split by pre-period P/I ratio
  5. Demand Outflow     — Non-metro (Micro + Noncore) null test
  6. Robustness         — R7 pre-determined WFH proxy, R8 detrended

Outputs:
  figures/fig1_cbsa_trends.png
  figures/fig2_cbsa_event_study.png
  figures/fig3_cbsa_was.png
  figures/fig4_cbsa_supply.png
  figures/fig5_cbsa_demand.png
  figures/fig6_cbsa_robustness.png
  results/cbsa_main_results.csv
  results/cbsa_was_results.csv
  results/cbsa_supply_results.csv
  results/cbsa_demand_results.csv
  results/cbsa_county_classification.csv
"""

import warnings
warnings.filterwarnings("ignore")

import numpy as np
import pandas as pd
import pyfixest as pf
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.gridspec as gridspec
import matplotlib.patches as mpatches
from pathlib import Path
from scipy import stats

# ── Paths & constants ─────────────────────────────────────────────────────────
ROOT      = Path(__file__).resolve().parents[1]   # Econ692_Capstone/Output/
FIG       = ROOT / "figures"
RES       = ROOT / "results"
DATA      = ROOT / "data"
CBSA_FILE = ROOT.parent / "Data" / "raw" / "cbsa" / "cbsa_delineation_2023.xlsx"
SEED     = 42
N_BOOT   = 1000
STAYER_PCT = 20.0
YEARS    = list(range(2017, 2024))
REF_YEAR = 2019
RED      = "#d62728"
BLUE     = "#1f77b4"
GREEN    = "#2ca02c"
ORANGE   = "#ff7f0e"
PURPLE   = "#9467bd"
GRAY     = "#7f7f7f"

np.random.seed(SEED)

plt.rcParams.update({
    "font.family": "serif", "font.size": 10,
    "axes.spines.top": False, "axes.spines.right": False,
})


def stars(p: float) -> str:
    if pd.isna(p): return ""
    if p < 0.001: return "***"
    if p < 0.01:  return "**"
    if p < 0.05:  return "*"
    if p < 0.10:  return "†"
    return ""


def pval_t(coef: float, se: float) -> float:
    if pd.isna(se) or se <= 0: return np.nan
    return 2 * (1 - stats.norm.cdf(abs(coef / se)))


def print_sep(title: str = "") -> None:
    print(f"\n{'='*65}")
    if title: print(f"  {title}")
    print("=" * 65)


# ══════════════════════════════════════════════════════════════════════════════
# 1. LOAD & CLASSIFY
# ══════════════════════════════════════════════════════════════════════════════
print_sep("1. Load data & CBSA classification")

county = pd.read_csv(DATA / "county_dataset.csv")
county = county[county["year"].between(2017, 2023)].copy()
county["price"]      = np.exp(county["log_price"])
county["pi_ratio"]   = county["price"] / county["median_income"].clip(lower=1)
county["log_income"] = np.log(county["median_income"].clip(lower=1))

# Load CBSA delineation
xl = pd.read_excel(CBSA_FILE, sheet_name=0, header=2)
ca = xl[xl["State Name"].str.strip() == "California"][[
    "CBSA Code", "CBSA Title",
    "Metropolitan/Micropolitan Statistical Area",
    "Central/Outlying County",
    "County/County Equivalent",
]].copy()
ca.columns = ["cbsa_code", "cbsa_title", "metro_micro", "central_outlying", "CountyName"]
ca["CountyName"] = ca["CountyName"].str.strip()

# Merge onto county panel
county = county.merge(
    ca[["CountyName", "cbsa_code", "cbsa_title", "metro_micro", "central_outlying"]],
    on="CountyName", how="left"
)
county["metro_micro"]     = county["metro_micro"].fillna("Noncore")
county["central_outlying"]= county["central_outlying"].fillna("—")


def cbsa_type(row: pd.Series) -> str:
    if row["metro_micro"] == "Metropolitan Statistical Area":
        return "Metro-Core" if row["central_outlying"] == "Central" else "Metro-Fringe"
    if row["metro_micro"] == "Micropolitan Statistical Area":
        return "Micropolitan"
    return "Noncore"


county["cbsa_type"] = county.apply(cbsa_type, axis=1)

# Summary
type_counts = county.groupby("cbsa_type")["CountyName"].nunique()
print(f"\n  County panel: {len(county):,} obs  |  {county['CountyName'].nunique()} counties  "
      f"|  {county['year'].nunique()} years")
print("\n  CBSA classification:")
for t, n in type_counts.items():
    print(f"    {t:<18} {n:>3} counties")

# Save classification table
cls_tbl = (county[county["year"] == 2019]
           [["CountyName", "cbsa_type", "cbsa_title", "metro_micro", "central_outlying"]]
           .drop_duplicates()
           .sort_values(["cbsa_type", "CountyName"]))
cls_tbl.to_csv(RES / "cbsa_county_classification.csv", index=False)
print(f"\n  Saved: cbsa_county_classification.csv")

# Working subsets
metro   = county[county["cbsa_type"] == "Metro-Core"].copy()
fringe  = county[county["cbsa_type"] == "Metro-Fringe"].copy()
nonmeta = county[county["cbsa_type"].isin(["Micropolitan", "Noncore"])].copy()

print(f"\n  Metro-Core obs:  {len(metro):,}  ({metro['CountyName'].nunique()} counties)")
print(f"  Metro-Fringe:    {len(fringe):,}  ({fringe['CountyName'].nunique()} counties)")
print(f"  Non-Metro obs:   {len(nonmeta):,}  ({nonmeta['CountyName'].nunique()} counties)")

# ══════════════════════════════════════════════════════════════════════════════
# 2. DESCRIPTIVE TRENDS
# ══════════════════════════════════════════════════════════════════════════════
print_sep("2. Descriptive trends")

fig, axes = plt.subplots(1, 2, figsize=(13, 5))
type_colors = {
    "Metro-Core": RED, "Metro-Fringe": ORANGE,
    "Micropolitan": BLUE, "Noncore": GRAY,
}
for t, col in type_colors.items():
    sub = county[county["cbsa_type"] == t]
    for ax, var, ylabel in zip(
        axes,
        ["wfh_rate", "log_price"],
        ["Mean WFH Rate (%)", "Mean log(Median House Price)"]
    ):
        mn = sub.groupby("year")[var].mean()
        # wfh_rate in county_dataset.csv is already expressed as a percent
        # (mean ~9.17, max 28.76), so no further *100 is needed.
        ax.plot(mn.index, mn.values, color=col, lw=2, marker="o", ms=5, label=t)

for ax, title in zip(axes, ["(A) WFH Rate by CBSA Type", "(B) log(Price) by CBSA Type"]):
    ax.axvline(2019.5, color="gray", ls="--", lw=1)
    ax.set(title=title, xlabel="Year")
    ax.legend(fontsize=8, frameon=False)
    ax.grid(axis="y", ls="--", alpha=0.3)
axes[0].set_ylabel("Mean WFH Rate (%)")
axes[1].set_ylabel("Mean log(Median House Price)")

fig.suptitle("Figure 1: Trends by CBSA Type  |  California 2017–2023",
             fontsize=12, fontweight="bold")
plt.tight_layout()
fig.savefig(FIG / "fig1_cbsa_trends.png", dpi=150, bbox_inches="tight")
plt.close()
print("  Saved: fig1_cbsa_trends.png")

# ══════════════════════════════════════════════════════════════════════════════
# 3. MAIN TWFE + DYNAMIC EVENT STUDY  (Metro-Core)
# ══════════════════════════════════════════════════════════════════════════════
print_sep("3. TWFE + Dynamic Event Study — Metro-Core")

# ── 3a. Baseline TWFE specs
twfe_specs = [
    ("(1) Baseline",     "log_price ~ wfh_rate | CountyName + year"),
    ("(2) +Income",      "log_price ~ wfh_rate + log_income | CountyName + year"),
    ("(3) +Unemp",       "log_price ~ wfh_rate + log_income + unemployment_rate | CountyName + year"),
]
twfe_rows = []
print(f"\n  {'Model':<22} {'Coef (%)':>9} {'SE (%)':>8} {'p':>7}  Stars")
print("  " + "-" * 55)
for label, fml in twfe_specs:
    fit = pf.feols(fml, data=metro, vcov={"CRV1": "CountyName"})
    c = float(fit.coef()["wfh_rate"])
    s = float(fit.se()["wfh_rate"])
    p = pval_t(c, s)
    twfe_rows.append({"model": label, "coef": c, "se": s, "p": p, "stars": stars(p)})
    print(f"  {label:<22} {c*100:>+9.3f}  {s*100:>8.3f}  {p:>7.4f}  {stars(p)}")

main_c = twfe_rows[1]["coef"]  # +Income spec as headline
main_s = twfe_rows[1]["se"]
main_p = twfe_rows[1]["p"]
print(f"\n  Headline: +Income TWFE  coef={main_c*100:+.3f}%  p={main_p:.4f}  {stars(main_p)}")

# ── 3b. Dynamic TWFE event study
print("\n  Dynamic TWFE event study:")
for y in YEARS:
    if y != REF_YEAR:
        metro[f"wfh_x_{y}"] = metro["wfh_rate"] * (metro["year"] == y).astype(int)

es_terms = " + ".join(f"wfh_x_{y}" for y in YEARS if y != REF_YEAR)
fit_dyn = pf.feols(
    f"log_price ~ {es_terms} + log_income | CountyName + year",
    data=metro, vcov={"CRV1": "CountyName"}
)
coefs_e = fit_dyn.coef().reset_index(); coefs_e.columns = ["term", "coef"]
ses_e   = fit_dyn.se().reset_index();   ses_e.columns   = ["term", "se"]
es = (coefs_e.merge(ses_e)
      .query('term.str.startswith("wfh_x_")', engine="python")
      .copy())
es["year"] = es["term"].str.extract(r"(\d{4})").astype(int)
ref_row = pd.DataFrame({"term": ["ref"], "coef": [0.], "se": [0.], "year": [REF_YEAR]})
es = pd.concat([es, ref_row], ignore_index=True).sort_values("year").reset_index(drop=True)

pre = es[es["year"] < REF_YEAR]
chi2_pre = sum((r["coef"] / r["se"])**2 for _, r in pre.iterrows() if r["se"] > 0)
p_pre    = 1 - stats.chi2.cdf(chi2_pre, df=len(pre))
pre_str  = f"chi2({len(pre)})={chi2_pre:.3f}  p={p_pre:.4f}  {'PASS' if p_pre > 0.1 else 'CONCERN'}"
print(f"  Pre-trend test: {pre_str}")
print(f"  {'Year':<6} {'Coef (%)':>9} {'SE (%)':>8} {'p':>7}")
print("  " + "-" * 36)
for _, r in es.iterrows():
    ref = " (ref)" if r["year"] == REF_YEAR else ""
    pv  = pval_t(r["coef"], r["se"]) if r["se"] > 0 else 1.0
    print(f"  {int(r['year']):<6} {r['coef']*100:>+9.3f}  {r['se']*100:>8.3f}  "
          f"{pv:>7.4f} {stars(pv):3}{ref}")

# ── Figure 2: Event study
fig, ax = plt.subplots(figsize=(10, 5))
ci = es["se"] * 1.96
ax.axhline(0, color="black", lw=0.8, ls="--")
ax.axvline(REF_YEAR - 0.5, color="gray", lw=0.8, ls=":", label="Pre/Post cutoff")
ax.fill_between(es["year"], (es["coef"] - ci) * 100,
                (es["coef"] + ci) * 100, alpha=0.18, color=RED)
ax.plot(es["year"], es["coef"] * 100, "o-", color=RED, lw=2, ms=8,
        label=f"Metro-Core TWFE  (N={metro['CountyName'].nunique()} counties)")
ax.axvspan(2019.5, 2023.5, alpha=0.05, color=RED)

# Annotate pre-trend
ax.annotate(f"Pre-trend: {pre_str}",
            xy=(0.03, 0.06), xycoords="axes fraction",
            fontsize=9, color="darkgreen" if p_pre > 0.1 else "red")

# Annotate post-period coefficients
for _, r in es[es["year"] >= 2020].iterrows():
    pv = pval_t(r["coef"], r["se"])
    if stars(pv):
        ax.annotate(f"{r['coef']*100:+.2f}%{stars(pv)}",
                    xy=(r["year"], (r["coef"] - r["se"] * 1.96) * 100 - 0.1),
                    ha="center", fontsize=8, color=RED)

ax.set(title="Figure 2: Dynamic TWFE Event Study — Metro-Core Counties\n"
             "California 2017–2023  |  County FE + Year FE  |  CRV1 SE by County",
       xlabel="Year", ylabel="% price change per 1pp WFH rate")
ax.set_xticks(YEARS)
ax.legend(fontsize=10, frameon=False)
ax.grid(axis="y", alpha=0.3)
plt.tight_layout()
fig.savefig(FIG / "fig2_cbsa_event_study.png", dpi=150, bbox_inches="tight")
plt.close()
print("\n  Saved: fig2_cbsa_event_study.png")

# ══════════════════════════════════════════════════════════════════════════════
# 3b. TWFE — Metro-Core vs Non-Metro side-by-side
# ══════════════════════════════════════════════════════════════════════════════
print_sep("3b. TWFE — Metro-Core vs Non-Metro (side-by-side)")

twfe_specs_both = [
    ("(1) Baseline",  "log_price ~ wfh_rate | CountyName + year"),
    ("(2) +Income",   "log_price ~ wfh_rate + log_income | CountyName + year"),
    ("(3) +Unemp",    "log_price ~ wfh_rate + log_income + unemployment_rate | CountyName + year"),
]

rows_both = []
header = f"  {'Model':<22} {'Metro Coef':>11} {'Metro p':>9}  {'NonMetro Coef':>14} {'NonMetro p':>11}"
print(header)
print("  " + "-" * 72)

for label, fml in twfe_specs_both:
    # Metro-Core
    fit_m = pf.feols(fml, data=metro, vcov={"CRV1": "CountyName"})
    cm = float(fit_m.coef()["wfh_rate"])
    sm = float(fit_m.se()["wfh_rate"])
    pm = pval_t(cm, sm)
    # Non-Metro
    fit_n = pf.feols(fml, data=nonmeta, vcov={"CRV1": "CountyName"})
    cn = float(fit_n.coef()["wfh_rate"])
    sn = float(fit_n.se()["wfh_rate"])
    pn = pval_t(cn, sn)

    rows_both.append({
        "model": label,
        "metro_coef": cm, "metro_se": sm, "metro_p": pm, "metro_stars": stars(pm),
        "nm_coef":    cn, "nm_se":    sn, "nm_p":    pn, "nm_stars":    stars(pn),
    })
    print(f"  {label:<22} {cm*100:>+9.3f}%{stars(pm):<3}  {pm:>7.4f}   "
          f"{cn*100:>+12.3f}%{stars(pn):<3}  {pn:>9.4f}")

# Save table
df_both = pd.DataFrame(rows_both)
df_both.to_csv(RES / "cbsa_metro_nonmetro_twfe.csv", index=False)
print("\n  Saved: cbsa_metro_nonmetro_twfe.csv")

# ── Figure: Metro vs Non-Metro TWFE coefficient comparison ────────────────
fig, ax = plt.subplots(figsize=(9, 5))

n_specs  = len(rows_both)
x        = np.arange(n_specs)
width    = 0.35
labels_x = [r["model"] for r in rows_both]

# Metro bars
m_coefs = [r["metro_coef"] * 100 for r in rows_both]
m_ses   = [r["metro_se"]   * 100 for r in rows_both]
m_stars = [r["metro_stars"]       for r in rows_both]
bars_m  = ax.bar(x - width / 2, m_coefs, width, label="Metro-Core (33 counties)",
                 color=RED, alpha=0.80, edgecolor="white")
ax.errorbar(x - width / 2, m_coefs, yerr=[1.96 * s for s in m_ses],
            fmt="none", color="black", capsize=5, lw=1.5)
for xi, coef, se, st in zip(x - width / 2, m_coefs, m_ses, m_stars):
    if st:
        ax.text(xi, coef - 1.96 * se - 0.06, st, ha="center", fontsize=11,
                fontweight="bold", color=RED)

# Non-Metro bars
n_coefs = [r["nm_coef"] * 100 for r in rows_both]
n_ses   = [r["nm_se"]   * 100 for r in rows_both]
n_stars = [r["nm_stars"]       for r in rows_both]
bars_n  = ax.bar(x + width / 2, n_coefs, width, label="Non-Metro (21 counties)",
                 color=BLUE, alpha=0.80, edgecolor="white")
ax.errorbar(x + width / 2, n_coefs, yerr=[1.96 * s for s in n_ses],
            fmt="none", color="black", capsize=5, lw=1.5)
for xi, coef, se, st in zip(x + width / 2, n_coefs, n_ses, n_stars):
    if st:
        ax.text(xi, coef - 1.96 * se - 0.06, st, ha="center", fontsize=11,
                fontweight="bold", color=BLUE)

ax.axhline(0, color="black", lw=0.8, ls="--")
ax.set_xticks(x)
ax.set_xticklabels(labels_x, fontsize=10)
ax.set_ylabel("% price change per 1pp WFH rate")
ax.set_title(
    "TWFE: Metro-Core vs Non-Metro Counties\n"
    "California 2017–2023  |  County FE + Year FE  |  CRV1 SE by County",
    fontsize=11, fontweight="bold"
)
ax.legend(fontsize=10, frameon=False)
ax.grid(axis="y", alpha=0.3)
plt.tight_layout()
fig.savefig(FIG / "fig_twfe_metro_vs_nonmetro.png", dpi=150, bbox_inches="tight")
plt.close()
print("  Saved: fig_twfe_metro_vs_nonmetro.png")

# ══════════════════════════════════════════════════════════════════════════════
# 3c. DYNAMIC TWFE EVENT STUDY — Non-Metro Counties
# ══════════════════════════════════════════════════════════════════════════════
print_sep("3c. Dynamic TWFE Event Study — Non-Metro Counties")

# Build interaction terms (uses full nonmeta before any Section 6 filtering)
for y in YEARS:
    if y != REF_YEAR:
        nonmeta[f"wfh_x_{y}"] = nonmeta["wfh_rate"] * (nonmeta["year"] == y).astype(int)

nm_es_terms_sa = " + ".join(f"wfh_x_{y}" for y in YEARS if y != REF_YEAR)
fit_nm_dyn_sa = pf.feols(
    f"log_price ~ {nm_es_terms_sa} + log_income | CountyName + year",
    data=nonmeta, vcov={"CRV1": "CountyName"}
)
c_nm_d_sa = fit_nm_dyn_sa.coef().reset_index(); c_nm_d_sa.columns = ["term", "coef"]
s_nm_d_sa = fit_nm_dyn_sa.se().reset_index();   s_nm_d_sa.columns = ["term", "se"]
es_nm_sa = (c_nm_d_sa.merge(s_nm_d_sa)
            .query('term.str.startswith("wfh_x_")', engine="python")
            .copy())
es_nm_sa["year"] = es_nm_sa["term"].str.extract(r"(\d{4})").astype(int)
ref_row_nm = pd.DataFrame({"term": ["ref"], "coef": [0.], "se": [0.], "year": [REF_YEAR]})
es_nm_sa = (pd.concat([es_nm_sa, ref_row_nm], ignore_index=True)
            .sort_values("year").reset_index(drop=True))

# Pre-trend Wald test
pre_nm_sa = es_nm_sa[es_nm_sa["year"] < REF_YEAR]
chi2_pre_nm_sa = sum(
    (r["coef"] / r["se"]) ** 2 for _, r in pre_nm_sa.iterrows() if r["se"] > 0
)
p_pre_nm_sa = 1 - stats.chi2.cdf(chi2_pre_nm_sa, df=len(pre_nm_sa))
pre_str_nm = (f"chi2({len(pre_nm_sa)})={chi2_pre_nm_sa:.3f}  p={p_pre_nm_sa:.4f}  "
              f"{'PASS' if p_pre_nm_sa > 0.1 else 'CONCERN'}")
print(f"\n  Pre-trend test: {pre_str_nm}")

# Print year-by-year table
print(f"  {'Year':<6} {'Coef (%)':>9} {'SE (%)':>8} {'p':>7}")
print("  " + "-" * 36)
for _, r in es_nm_sa.iterrows():
    ref_tag = " (ref)" if r["year"] == REF_YEAR else ""
    pv = pval_t(r["coef"], r["se"]) if r["se"] > 0 else 1.0
    print(f"  {int(r['year']):<6} {r['coef']*100:>+9.3f}  {r['se']*100:>8.3f}  "
          f"{pv:>7.4f} {stars(pv):3}{ref_tag}")

# Save CSV
es_nm_save = es_nm_sa.copy()
es_nm_save["p"] = es_nm_save.apply(
    lambda r: pval_t(r["coef"], r["se"]) if r["se"] > 0 else 1.0, axis=1
)
es_nm_save["stars"] = es_nm_save["p"].apply(stars)
es_nm_save["ref_year"] = (es_nm_save["year"] == REF_YEAR)
es_nm_save.to_csv(RES / "cbsa_nm_event_study.csv", index=False)
print(f"\n  Saved: cbsa_nm_event_study.csv")

# ── Figure: Non-Metro Event Study (standalone) ────────────────────────────
fig, ax = plt.subplots(figsize=(10, 5))
ci_nm_sa = es_nm_sa["se"] * 1.96
ax.axhline(0, color="black", lw=0.8, ls="--")
ax.axvline(REF_YEAR - 0.5, color="gray", lw=0.8, ls=":", label="Pre/Post cutoff")
ax.fill_between(es_nm_sa["year"],
                (es_nm_sa["coef"] - ci_nm_sa) * 100,
                (es_nm_sa["coef"] + ci_nm_sa) * 100,
                alpha=0.18, color=BLUE)
ax.plot(es_nm_sa["year"], es_nm_sa["coef"] * 100, "s-", color=BLUE, lw=2, ms=8,
        label=f"Non-Metro TWFE  (N={nonmeta['CountyName'].nunique()} counties)")
ax.axvspan(2019.5, 2023.5, alpha=0.05, color=BLUE)

# Annotate pre-trend
ax.annotate(f"Pre-trend: {pre_str_nm}",
            xy=(0.03, 0.06), xycoords="axes fraction",
            fontsize=9, color="darkgreen" if p_pre_nm_sa > 0.1 else "red")

# Annotate significant post-period coefficients
for _, r in es_nm_sa[es_nm_sa["year"] >= 2020].iterrows():
    pv = pval_t(r["coef"], r["se"])
    if stars(pv):
        ax.annotate(f"{r['coef']*100:+.2f}%{stars(pv)}",
                    xy=(r["year"], (r["coef"] - r["se"] * 1.96) * 100 - 0.08),
                    ha="center", fontsize=8, color=BLUE)

ax.set(title="Dynamic TWFE Event Study — Non-Metro Counties\n"
             "California 2017-2023  |  County FE + Year FE  |  CRV1 SE by County",
       xlabel="Year", ylabel="% price change per 1pp WFH rate")
ax.set_xticks(YEARS)
ax.legend(fontsize=10, frameon=False)
ax.grid(axis="y", alpha=0.3)
plt.tight_layout()
fig.savefig(FIG / "fig_nm_event_study.png", dpi=150, bbox_inches="tight")
plt.close()
print("  Saved: fig_nm_event_study.png")

# ══════════════════════════════════════════════════════════════════════════════
# 4. DH(2024) WAS — Metro-Core Counties
# ══════════════════════════════════════════════════════════════════════════════
print_sep("4. DH(2024) WAS — Metro-Core Counties")

# Build first-difference dataset (Metro-Core only)
fd_metro = metro.sort_values(["CountyName", "year"]).copy()
for col in ["log_price", "wfh_rate", "log_income", "unemployment_rate"]:
    fd_metro[f"d_{col}"] = fd_metro.groupby("CountyName")[col].diff()
fd_metro["d_lag"] = fd_metro.groupby("CountyName")["wfh_rate"].shift(1)
fd = fd_metro.dropna(subset=["d_log_price", "d_wfh_rate", "d_log_income"]).copy()
fd = fd.rename(columns={"d_log_price": "dy", "d_wfh_rate": "dd"})

print(f"  First-diff obs: {len(fd):,}  |  Counties: {fd['CountyName'].nunique()}")


def estimate_was_period(sub: pd.DataFrame, stayer_pct: float = 20.0,
                        controls: list | None = None) -> dict | None:
    controls = controls or []
    if len(sub) < 4:
        return None
    thresh = max(np.percentile(sub["dd"].abs(), stayer_pct), 1e-5)
    stay   = sub[sub["dd"].abs() <= thresh]
    swit   = sub[sub["dd"].abs() >  thresh]
    if len(swit) < 2 or len(stay) < 2:
        return None
    # Residualise switchers on stayer trends (order-1 polynomial of lag)
    X_w = stay[["d_lag"]].values
    if controls:
        X_w = np.hstack([X_w, stay[controls].values])
    X_w = np.hstack([np.ones((len(X_w), 1)), X_w])
    try:
        coef, *_ = np.linalg.lstsq(X_w, stay["dy"].values, rcond=None)
    except np.linalg.LinAlgError:
        return None
    X_sw = swit[["d_lag"]].values
    if controls:
        X_sw = np.hstack([X_sw, swit[controls].values])
    X_sw = np.hstack([np.ones((len(X_sw), 1)), X_sw])
    slopes   = (swit["dy"].values - X_sw @ coef) / swit["dd"].values
    abs_dd   = swit["dd"].abs().values
    return {
        "as_est":      slopes.mean(),
        "was_est":     np.average(slopes, weights=abs_dd),
        "n_switch":    len(swit),
        "n_stay":      len(stay),
        "sum_abs_dd":  abs_dd.sum(),
    }


def run_was(fd_in: pd.DataFrame, controls: list | None = None) -> dict | None:
    periods = {}
    for t in sorted(fd_in["year"].unique()):
        r = estimate_was_period(fd_in[fd_in["year"] == t].copy(), STAYER_PCT, controls)
        if r:
            periods[t] = r
    if not periods:
        return None
    pts   = sorted(periods)
    w_n   = np.array([periods[t]["n_switch"]  for t in pts], dtype=float)
    w_dd  = np.array([periods[t]["sum_abs_dd"] for t in pts], dtype=float)
    as_ov  = np.average([periods[t]["as_est"]  for t in pts], weights=w_n)
    was_ov = np.average([periods[t]["was_est"] for t in pts], weights=w_dd)
    return {
        "as_overall":  as_ov,
        "was_overall": was_ov,
        "periods": pts,
        "as_t":  [periods[t]["as_est"]  for t in pts],
        "was_t": [periods[t]["was_est"] for t in pts],
        "n_sw_t":[periods[t]["n_switch"] for t in pts],
    }


def bootstrap_was(fd_in: pd.DataFrame, controls: list | None = None,
                  n_boot: int = N_BOOT) -> dict | None:
    rng   = np.random.default_rng(SEED)
    base  = run_was(fd_in, controls)
    if base is None:
        return None
    cnts  = fd_in["CountyName"].unique()
    was_b, as_b = [], []
    for _ in range(n_boot):
        signs = pd.Series(rng.choice([-1., 1.], len(cnts)), index=cnts)
        bd    = fd_in.copy()
        bd["dy"] *= bd["CountyName"].map(signs)
        r = run_was(bd, controls)
        if r:
            was_b.append(r["was_overall"])
            as_b.append(r["as_overall"])
    was_se = np.std(was_b)
    was_p  = float(np.mean(np.abs(was_b) >= abs(base["was_overall"])))
    as_se  = np.std(as_b)
    as_p   = float(np.mean(np.abs(as_b)  >= abs(base["as_overall"])))
    return {**base,
            "was_overall_se": was_se, "was_p": was_p,
            "as_overall_se":  as_se,  "as_p":  as_p}


ctrl_specs = [
    ("No controls",   None),
    ("+d.Income",     ["d_log_income"]),
    ("+d.Unemp",      ["d_unemployment_rate" if "d_unemployment_rate" in fd.columns else "d_log_income"]),
    ("+All controls", ["d_log_income"]),
]

was_results = {}
print(f"\n  {'Specification':<22} {'WAS (%)':>9} {'SE (%)':>8} {'p':>7}  Stars")
print("  " + "-" * 56)
for label, ctrl in ctrl_specs:
    print(f"  {label} ...", end="", flush=True)
    r = bootstrap_was(fd, ctrl, N_BOOT)
    was_results[label] = r
    if r:
        print(f"\r  {label:<22} {r['was_overall']*100:>+9.3f}  "
              f"{r['was_overall_se']*100:>8.3f}  {r['was_p']:>7.4f}  "
              f"{stars(r['was_p'])}")
    else:
        print("  FAILED")

was_base = was_results["No controls"]
was_inc  = was_results["+d.Income"]

# Save WAS results
was_rows = []
for label, r in was_results.items():
    if r:
        was_rows.append({
            "spec": label, "was": r["was_overall"], "se": r["was_overall_se"],
            "p": r["was_p"], "stars": stars(r["was_p"]),
            "as": r["as_overall"], "as_se": r["as_overall_se"], "as_p": r["as_p"],
        })
pd.DataFrame(was_rows).to_csv(RES / "cbsa_was_results.csv", index=False)

# ── Figure 3: WAS period-by-period
fig, ax = plt.subplots(figsize=(10, 5))
if was_base:
    ax.axhline(0, color="black", lw=0.8, ls="--")
    ax.axvline(REF_YEAR - 0.5, color="gray", lw=0.8, ls=":")
    ax.plot(was_base["periods"], [x * 100 for x in was_base["was_t"]],
            "o-", color=GREEN, lw=2, ms=8, label="WAS (weighted by |ΔD|)")
    ax.plot(was_base["periods"], [x * 100 for x in was_base["as_t"]],
            "s--", color=GRAY, lw=1.5, ms=6, label="AS (unweighted)")
    ax.axhline(was_base["was_overall"] * 100, color=GREEN, ls="-.", lw=1.5, alpha=0.8,
               label=f"Overall WAS = {was_base['was_overall']*100:+.3f}%"
                     f" (p={was_base['was_p']:.3f})")
    ax.axvspan(2019.5, 2023.5, alpha=0.05, color=GREEN)
    ax.set(title="Figure 3: DH(2024) WAS — Metro-Core Counties\n"
                 "Wild Bootstrap SE (1,000 draws)  |  No Controls",
           xlabel="Year", ylabel="Slope estimate (% per 1pp WFH)")
    ax.set_xticks(was_base["periods"])
    ax.legend(fontsize=10, frameon=False)
    ax.grid(axis="y", alpha=0.3)
plt.tight_layout()
fig.savefig(FIG / "fig3_cbsa_was.png", dpi=150, bbox_inches="tight")
plt.close()
print("\n  Saved: fig3_cbsa_was.png")

# ══════════════════════════════════════════════════════════════════════════════
# 5. SUPPLY HETEROGENEITY — Metro-Core split by pre-period P/I ratio
# ══════════════════════════════════════════════════════════════════════════════
print_sep("5. Supply Heterogeneity — Metro-Core P/I Ratio Split")

# Compute pre-period (2017-2019) average P/I ratio per county
pre_pi = (metro[metro["year"] <= 2019]
          .groupby("CountyName")["pi_ratio"]
          .mean()
          .reset_index()
          .rename(columns={"pi_ratio": "pre_pi_ratio"}))

median_pi = pre_pi["pre_pi_ratio"].median()
pre_pi["supply_group"] = np.where(
    pre_pi["pre_pi_ratio"] >= median_pi,
    "Constrained (High P/I)",
    "Elastic (Low P/I)"
)
metro = metro.merge(pre_pi[["CountyName", "supply_group", "pre_pi_ratio"]],
                    on="CountyName", how="left")

print(f"\n  P/I ratio split threshold (median): {median_pi:.2f}")
print(f"  Constrained counties: {(pre_pi['supply_group'] == 'Constrained (High P/I)').sum()}")
print(f"  Elastic counties:     {(pre_pi['supply_group'] == 'Elastic (Low P/I)').sum()}")
print()

# Show top/bottom counties
hi = pre_pi.nlargest(5, "pre_pi_ratio")[["CountyName", "pre_pi_ratio", "supply_group"]]
lo = pre_pi.nsmallest(5, "pre_pi_ratio")[["CountyName", "pre_pi_ratio", "supply_group"]]
print("  Most constrained (highest P/I):")
for _, r in hi.iterrows():
    print(f"    {r['CountyName']:<30} P/I = {r['pre_pi_ratio']:.1f}")
print("  Most elastic (lowest P/I):")
for _, r in lo.iterrows():
    print(f"    {r['CountyName']:<30} P/I = {r['pre_pi_ratio']:.1f}")

# ── TWFE interaction: wfh_rate × supply_group
metro["wfh_constrained"] = (metro["wfh_rate"]
                             * (metro["supply_group"] == "Constrained (High P/I)").astype(int))
metro["wfh_elastic"]     = (metro["wfh_rate"]
                             * (metro["supply_group"] == "Elastic (Low P/I)").astype(int))

fit_int = pf.feols(
    "log_price ~ wfh_constrained + wfh_elastic + log_income | CountyName + year",
    data=metro, vcov={"CRV1": "CountyName"}
)
c_con = float(fit_int.coef()["wfh_constrained"])
s_con = float(fit_int.se()["wfh_constrained"])
p_con = pval_t(c_con, s_con)
c_ela = float(fit_int.coef()["wfh_elastic"])
s_ela = float(fit_int.se()["wfh_elastic"])
p_ela = pval_t(c_ela, s_ela)
diff_se = np.sqrt(s_con**2 + s_ela**2)
diff_p  = pval_t(c_con - c_ela, diff_se)

print(f"\n  Interaction TWFE:")
print(f"  {'Group':<30} {'Coef (%)':>9} {'SE (%)':>8} {'p':>7}  Stars")
print("  " + "-" * 60)
print(f"  {'Constrained (High P/I)':<30} {c_con*100:>+9.3f}  {s_con*100:>8.3f}  "
      f"{p_con:>7.4f}  {stars(p_con)}")
print(f"  {'Elastic (Low P/I)':<30} {c_ela*100:>+9.3f}  {s_ela*100:>8.3f}  "
      f"{p_ela:>7.4f}  {stars(p_ela)}")
print(f"  {'Difference (Con - Ela)':<30} {(c_con-c_ela)*100:>+9.3f}  "
      f"{diff_se*100:>8.3f}  {diff_p:>7.4f}  {stars(diff_p)}")

# ── Dynamic TWFE by supply group
supply_es = {}
for grp, col in [("Constrained (High P/I)", RED), ("Elastic (Low P/I)", BLUE)]:
    sub = metro[metro["supply_group"] == grp].copy()
    for y in YEARS:
        if y != REF_YEAR:
            sub[f"wfh_x_{y}"] = sub["wfh_rate"] * (sub["year"] == y).astype(int)
    es_t = " + ".join(f"wfh_x_{y}" for y in YEARS if y != REF_YEAR)
    try:
        fit_s = pf.feols(
            f"log_price ~ {es_t} + log_income | CountyName + year",
            data=sub, vcov={"CRV1": "CountyName"}
        )
        c_s = fit_s.coef().reset_index(); c_s.columns = ["term", "coef"]
        s_s = fit_s.se().reset_index();   s_s.columns = ["term", "se"]
        es_g = (c_s.merge(s_s)
                .query('term.str.startswith("wfh_x_")', engine="python")
                .copy())
        es_g["year"] = es_g["term"].str.extract(r"(\d{4})").astype(int)
        ref_row = pd.DataFrame({"term": ["ref"], "coef": [0.], "se": [0.], "year": [REF_YEAR]})
        es_g = pd.concat([es_g, ref_row], ignore_index=True).sort_values("year")
        supply_es[grp] = es_g
    except Exception as e:
        print(f"  [{grp}] event study failed: {e}")

# ── Figure 4: Supply heterogeneity
fig = plt.figure(figsize=(13, 6.5))
gs  = gridspec.GridSpec(1, 2, figure=fig, wspace=0.38,
                        top=0.82, bottom=0.12, left=0.08, right=0.97)

# Panel A: event study by supply group
ax1 = fig.add_subplot(gs[0, 0])
ax1.axhline(0, color="black", lw=0.8, ls="--")
ax1.axvline(REF_YEAR - 0.5, color="gray", lw=0.8, ls=":")
for grp, col, mk in [("Constrained (High P/I)", RED, "o"),
                      ("Elastic (Low P/I)", BLUE, "s")]:
    if grp in supply_es:
        eg = supply_es[grp]
        ci = eg["se"] * 1.96
        ax1.fill_between(eg["year"], (eg["coef"] - ci)*100,
                          (eg["coef"] + ci)*100, alpha=0.12, color=col)
        ax1.plot(eg["year"], eg["coef"]*100, f"{mk}-", color=col,
                  lw=2, ms=7, label=grp.split(" (")[0])
ax1.set_title("Panel A: Event Study by Supply Group", fontsize=11)
ax1.set_xlabel("Year"); ax1.set_ylabel("% price change per 1pp WFH")
ax1.set_xticks(YEARS); ax1.tick_params(axis="x", rotation=30)
ax1.legend(fontsize=9, frameon=False); ax1.grid(axis="y", alpha=0.3)

# Panel B: coefficient comparison
ax2 = fig.add_subplot(gs[0, 1])
groups   = ["Constrained\n(High P/I)", "Elastic\n(Low P/I)"]
coefs_s  = [c_con * 100, c_ela * 100]
ses_s    = [s_con * 100, s_ela * 100]
stars_s  = [stars(p_con), stars(p_ela)]
bar_cols = [RED, BLUE]
bars = ax2.bar(groups, coefs_s, color=bar_cols, alpha=0.75, width=0.45, edgecolor="white")
ax2.errorbar(groups, coefs_s, yerr=[1.96 * s for s in ses_s],
             fmt="none", color="black", capsize=6, lw=1.8)
ax2.axhline(0, color="black", lw=0.8)
for bar, st, cv, sv in zip(bars, stars_s, coefs_s, ses_s):
    ypos = cv - 1.96 * sv - 0.08
    ax2.text(bar.get_x() + bar.get_width() / 2, ypos, st,
             ha="center", fontsize=14, fontweight="bold", color=bar.get_facecolor())
ax2.set_title("Panel B: Supply Group Coefficients\n(TWFE Interaction)", fontsize=11)
ax2.set_ylabel("% price change per 1pp WFH"); ax2.grid(axis="y", alpha=0.3)
ax2.annotate(f"Diff = {(c_con-c_ela)*100:+.3f}%  p={diff_p:.3f}  {stars(diff_p)}",
             xy=(0.5, 0.05), xycoords="axes fraction", ha="center",
             fontsize=9, color="black",
             bbox=dict(boxstyle="round,pad=0.3", facecolor="lightyellow", alpha=0.8))

fig.suptitle("Figure 4: Supply Heterogeneity — Metro-Core Counties\n"
             "P/I Ratio Split  |  County FE + Year FE  |  CRV1 SE",
             fontsize=12, fontweight="bold", y=0.98)
fig.savefig(FIG / "fig4_cbsa_supply.png", dpi=150, bbox_inches="tight")
plt.close()
print("  Saved: fig4_cbsa_supply.png")

# Save supply results
supply_rows = [
    {"group": "Constrained (High P/I)", "N_counties": (metro["supply_group"] == "Constrained (High P/I)").sum() // len(YEARS),
     "coef": c_con, "se": s_con, "p": p_con, "stars": stars(p_con)},
    {"group": "Elastic (Low P/I)", "N_counties": (metro["supply_group"] == "Elastic (Low P/I)").sum() // len(YEARS),
     "coef": c_ela, "se": s_ela, "p": p_ela, "stars": stars(p_ela)},
    {"group": "Difference", "N_counties": np.nan,
     "coef": c_con - c_ela, "se": diff_se, "p": diff_p, "stars": stars(diff_p)},
]
pd.DataFrame(supply_rows).to_csv(RES / "cbsa_supply_results.csv", index=False)
print("  Saved: cbsa_supply_results.csv")

# ══════════════════════════════════════════════════════════════════════════════
# 6. DEMAND OUTFLOW NULL — Non-Metro Counties
# ══════════════════════════════════════════════════════════════════════════════
print_sep("6. Demand Outflow Null — Non-Metro (Micro + Noncore)")

# Non-metro own WFH effect (should be null if demand exits CA)
fit_nm = pf.feols(
    "log_price ~ wfh_rate + log_income | CountyName + year",
    data=nonmeta, vcov={"CRV1": "CountyName"}
)
c_nm = float(fit_nm.coef()["wfh_rate"])
s_nm = float(fit_nm.se()["wfh_rate"])
p_nm = pval_t(c_nm, s_nm)
print(f"\n  Non-Metro TWFE:  coef={c_nm*100:+.3f}%  SE={s_nm*100:.3f}  "
      f"p={p_nm:.4f}  {stars(p_nm)}")
print(f"  N counties: {nonmeta['CountyName'].nunique()}  |  N obs: {len(nonmeta)}")

# Metro-Core regional WFH → non-metro price test
# Build region-level urban WFH and merge onto non-metro counties
REGION_MAP = {
    "Alameda County": "Bay Area", "Contra Costa County": "Bay Area",
    "Marin County": "Bay Area", "San Francisco County": "Bay Area",
    "San Mateo County": "Bay Area", "Santa Clara County": "Bay Area",
    "Santa Cruz County": "Bay Area", "Sonoma County": "Bay Area",
    "Napa County": "Bay Area", "Solano County": "Bay Area",
    "Sacramento County": "Sacramento", "Placer County": "Sacramento",
    "El Dorado County": "Sacramento", "Yolo County": "Sacramento",
    "Los Angeles County": "SoCal", "Orange County": "SoCal",
    "San Diego County": "SoCal", "Riverside County": "SoCal",
    "San Bernardino County": "SoCal", "Ventura County": "SoCal",
    "Fresno County": "Central Valley", "Kern County": "Central Valley",
    "Stanislaus County": "Central Valley", "San Joaquin County": "Central Valley",
    "Merced County": "Central Valley", "Kings County": "Central Valley",
    "Tulare County": "Central Valley", "Madera County": "Central Valley",
    "Santa Barbara County": "Central Coast", "San Luis Obispo County": "Central Coast",
    "Monterey County": "Central Coast",
    "Shasta County": "North", "Sutter County": "North", "Yuba County": "North",
    "Butte County": "North", "Imperial County": "Desert",
}
NON_METRO_REGION = {
    "Amador County": "Sacramento", "Calaveras County": "Central Valley",
    "Colusa County": "Sacramento", "Glenn County": "North",
    "Mariposa County": "Central Valley", "Modoc County": "North",
    "Mono County": "Desert", "Plumas County": "North",
    "Sierra County": "North", "Siskiyou County": "North",
    "Trinity County": "North", "Alpine County": "Sacramento",
    "Del Norte County": "North", "Humboldt County": "North",
    "Inyo County": "Desert", "Lake County": "Bay Area",
    "Lassen County": "North", "Mendocino County": "Bay Area",
    "Nevada County": "Sacramento", "Tehama County": "North",
    "Tuolumne County": "Central Valley",
}

metro["region"]   = metro["CountyName"].map(REGION_MAP).fillna("Other")
nonmeta["region"] = nonmeta["CountyName"].map({**REGION_MAP, **NON_METRO_REGION}).fillna("Other")

region_metro_wfh = (metro.groupby(["region", "year"])["wfh_rate"]
                    .mean().reset_index()
                    .rename(columns={"wfh_rate": "region_metro_wfh"}))

nonmeta = nonmeta.merge(region_metro_wfh, on=["region", "year"], how="left")
miss = nonmeta["region_metro_wfh"].isna().sum()
nonmeta = nonmeta.dropna(subset=["region_metro_wfh"])
print(f"  Non-metro obs with region metro WFH: {len(nonmeta)} (dropped {miss} missing)")

if nonmeta["CountyName"].nunique() >= 5:
    fit_spill = pf.feols(
        "log_price ~ wfh_rate + region_metro_wfh + log_income | CountyName + year",
        data=nonmeta, vcov={"CRV1": "CountyName"}
    )
    coef_names = fit_spill.coef().index.tolist()
    c_sp = float(fit_spill.coef()["region_metro_wfh"]) if "region_metro_wfh" in coef_names else np.nan
    s_sp = float(fit_spill.se()["region_metro_wfh"])   if "region_metro_wfh" in coef_names else np.nan
    p_sp = pval_t(c_sp, s_sp)
    print(f"  Non-Metro spillover: coef={c_sp*100:+.3f}%  SE={s_sp*100:.3f}  "
          f"p={p_sp:.4f}  {stars(p_sp)}")
    if not np.isnan(c_sp):
        if c_sp > 0 and p_sp < 0.1:
            print("  -> POSITIVE: non-metro counties gain when nearby metros have high WFH")
        elif c_sp < 0 and p_sp < 0.1:
            print("  -> NEGATIVE: WFH shock depresses all counties, demand exits CA")
        else:
            print(f"  -> NULL: no significant spillover (p={p_sp:.3f}), demand exits CA")

    # ── Metro × Non-Metro WFH interaction ─────────────────────────────────
    # Tests whether non-metro own-WFH effect depends on nearby metro WFH.
    # Positive interaction: complementarity (non-metro WFH + metro WFH reinforce each other)
    # Negative interaction: substitution (non-metro WFH effect dampened when metro WFH is high)
    print("\n  -- Metro x Non-Metro WFH Interaction --")
    # Demean to reduce collinearity and improve interpretability of main effects
    wfh_mean    = nonmeta["wfh_rate"].mean()
    rmetro_mean = nonmeta["region_metro_wfh"].mean()
    nonmeta["wfh_c"]        = nonmeta["wfh_rate"] - wfh_mean
    nonmeta["region_metro_c"] = nonmeta["region_metro_wfh"] - rmetro_mean
    nonmeta["wfh_x_rmetro"] = nonmeta["wfh_c"] * nonmeta["region_metro_c"]

    fit_int = pf.feols(
        "log_price ~ wfh_c + region_metro_c + wfh_x_rmetro + log_income "
        "| CountyName + year",
        data=nonmeta, vcov={"CRV1": "CountyName"}
    )
    c_own_i = float(fit_int.coef()["wfh_c"])
    s_own_i = float(fit_int.se()["wfh_c"])
    p_own_i = pval_t(c_own_i, s_own_i)
    c_met_i = float(fit_int.coef()["region_metro_c"])
    s_met_i = float(fit_int.se()["region_metro_c"])
    p_met_i = pval_t(c_met_i, s_met_i)
    c_int   = float(fit_int.coef()["wfh_x_rmetro"])
    s_int   = float(fit_int.se()["wfh_x_rmetro"])
    p_int   = pval_t(c_int, s_int)

    print(f"  Own WFH (at mean region metro):     "
          f"coef={c_own_i*100:+.3f}%  SE={s_own_i*100:.3f}  p={p_own_i:.4f}  {stars(p_own_i)}")
    print(f"  Region metro WFH (at mean own WFH): "
          f"coef={c_met_i*100:+.3f}%  SE={s_met_i*100:.3f}  p={p_met_i:.4f}  {stars(p_met_i)}")
    print(f"  Interaction (own x region metro):   "
          f"coef={c_int*100:+.4f}%  SE={s_int*100:.4f}  p={p_int:.4f}  {stars(p_int)}")

    # Marginal effect of own WFH at low vs high regional metro WFH (25th / 75th pct)
    q25 = nonmeta["region_metro_wfh"].quantile(0.25)
    q75 = nonmeta["region_metro_wfh"].quantile(0.75)
    me_low  = c_own_i + c_int * (q25 - rmetro_mean)
    me_high = c_own_i + c_int * (q75 - rmetro_mean)
    print(f"\n  Marginal effect of own WFH:")
    print(f"    At low regional metro WFH (p25={q25:.2f}%):  "
          f"{me_low*100:+.3f}%")
    print(f"    At high regional metro WFH (p75={q75:.2f}%): "
          f"{me_high*100:+.3f}%")

    if p_int < 0.10:
        if c_int > 0:
            print("  -> COMPLEMENTARITY: Non-metro WFH more beneficial near high-WFH metros")
            print("     (consistent with receiving displaced metro workers)")
        else:
            print("  -> SUBSTITUTION: Non-metro WFH effect dampened by high metro WFH")
            print("     (demand leakage already depresses non-metro prices)")
    else:
        print(f"  -> NULL interaction (p={p_int:.3f}): own and metro WFH act additively")
else:
    c_sp, s_sp, p_sp = np.nan, np.nan, np.nan
    c_own_i = s_own_i = p_own_i = np.nan
    c_met_i = s_met_i = p_met_i = np.nan
    c_int = s_int = p_int = np.nan
    me_low = me_high = np.nan
    print("  Insufficient non-metro counties for spillover regression.")

# Dynamic event study — Non-Metro
for y in YEARS:
    if y != REF_YEAR:
        nonmeta[f"wfh_x_{y}"] = nonmeta["wfh_rate"] * (nonmeta["year"] == y).astype(int)
nm_es_terms = " + ".join(f"wfh_x_{y}" for y in YEARS if y != REF_YEAR)
fit_nm_dyn = pf.feols(
    f"log_price ~ {nm_es_terms} + log_income | CountyName + year",
    data=nonmeta, vcov={"CRV1": "CountyName"}
)
c_nm_d = fit_nm_dyn.coef().reset_index(); c_nm_d.columns = ["term", "coef"]
s_nm_d = fit_nm_dyn.se().reset_index();   s_nm_d.columns = ["term", "se"]
es_nm  = (c_nm_d.merge(s_nm_d)
          .query('term.str.startswith("wfh_x_")', engine="python")
          .copy())
es_nm["year"] = es_nm["term"].str.extract(r"(\d{4})").astype(int)
es_nm = pd.concat([es_nm, ref_row.copy()], ignore_index=True).sort_values("year")

pre_nm = es_nm[es_nm["year"] < REF_YEAR]
chi2_nm = sum((r["coef"] / r["se"])**2 for _, r in pre_nm.iterrows() if r["se"] > 0)
p_pre_nm = 1 - stats.chi2.cdf(chi2_nm, df=len(pre_nm)) if len(pre_nm) > 0 else 1.0

# ── Figure 5: Demand outflow
fig, axes = plt.subplots(1, 2, figsize=(13, 6))

# Panel A: Metro-Core vs Non-Metro event studies
ax = axes[0]
ax.axhline(0, color="black", lw=0.8, ls="--")
ax.axvline(REF_YEAR - 0.5, color="gray", lw=0.8, ls=":")
for es_plot, col, lbl, mk in [
    (es,    RED,  f"Metro-Core ({metro['CountyName'].nunique()} counties)", "o"),
    (es_nm, BLUE, f"Non-Metro ({nonmeta['CountyName'].nunique()} counties)", "s"),
]:
    ci_p = es_plot["se"] * 1.96
    ax.fill_between(es_plot["year"], (es_plot["coef"] - ci_p)*100,
                     (es_plot["coef"] + ci_p)*100, alpha=0.12, color=col)
    ax.plot(es_plot["year"], es_plot["coef"]*100, f"{mk}-",
             color=col, lw=2, ms=7, label=lbl)
ax.set_title("Panel A: Metro-Core vs Non-Metro\nDynamic TWFE Event Study", fontsize=11)
ax.set_xlabel("Year"); ax.set_ylabel("% price change per 1pp WFH")
ax.set_xticks(YEARS); ax.tick_params(axis="x", rotation=30)
ax.legend(fontsize=9, frameon=False); ax.grid(axis="y", alpha=0.3)
ax.annotate(f"Metro pre-trend: p={p_pre:.3f}  {'PASS' if p_pre > 0.1 else 'FAIL'}",
            xy=(0.03, 0.10), xycoords="axes fraction", fontsize=8,
            color="darkgreen" if p_pre > 0.1 else "red")
ax.annotate(f"Non-Metro pre-trend: p={p_pre_nm:.3f}  {'PASS' if p_pre_nm > 0.1 else 'FAIL'}",
            xy=(0.03, 0.04), xycoords="axes fraction", fontsize=8,
            color="darkgreen" if p_pre_nm > 0.1 else "red")

# Panel B: coefficient comparison bar
ax = axes[1]
labels_d = ["Metro-Core\n(own WFH)", "Non-Metro\n(own WFH)"]
coefs_d  = [main_c * 100, c_nm * 100]
ses_d    = [main_s * 100, s_nm * 100]
cols_d   = [RED, BLUE]
sts_d    = [stars(main_p), stars(p_nm)]
bars = ax.bar(labels_d, coefs_d, color=cols_d, alpha=0.75, width=0.45, edgecolor="white")
ax.errorbar(labels_d, coefs_d, yerr=[1.96 * s for s in ses_d],
            fmt="none", color="black", capsize=6, lw=1.8)
ax.axhline(0, color="black", lw=0.8)
for bar, st, cv, sv in zip(bars, sts_d, coefs_d, ses_d):
    ypos = cv - 1.96 * sv - 0.05
    ax.text(bar.get_x() + bar.get_width() / 2, ypos, st,
            ha="center", fontsize=14, fontweight="bold", color=bar.get_facecolor())
ax.set_title("Panel B: WFH Effect — Metro vs Non-Metro\n(TWFE +Income)", fontsize=11)
ax.set_ylabel("% price change per 1pp WFH"); ax.grid(axis="y", alpha=0.3)

fig.suptitle(
    "Figure 5: Demand Outflow Test — Metro-Core vs Non-Metro Counties\n"
    "No positive non-metro response → demand exits California",
    fontsize=12, fontweight="bold", linespacing=1.6
)
fig.subplots_adjust(top=0.84, wspace=0.35)
fig.savefig(FIG / "fig5_cbsa_demand.png", dpi=150, bbox_inches="tight")
plt.close()
print("  Saved: fig5_cbsa_demand.png")

demand_rows = [
    {"group": "Metro-Core (own WFH)", "N": metro["CountyName"].nunique(),
     "coef": main_c, "se": main_s, "p": main_p, "stars": stars(main_p)},
    {"group": "Non-Metro (own WFH)", "N": nonmeta["CountyName"].nunique(),
     "coef": c_nm, "se": s_nm, "p": p_nm, "stars": stars(p_nm)},
]
if not np.isnan(c_sp):
    demand_rows.append({
        "group": "Non-Metro (region metro WFH spillover)", "N": nonmeta["CountyName"].nunique(),
        "coef": c_sp, "se": s_sp, "p": p_sp, "stars": stars(p_sp),
    })
if not np.isnan(c_int):
    demand_rows.append({
        "group": "Own WFH (interaction model, at mean region metro)",
        "N": nonmeta["CountyName"].nunique(),
        "coef": c_own_i, "se": s_own_i, "p": p_own_i, "stars": stars(p_own_i),
    })
    demand_rows.append({
        "group": "Region metro WFH (interaction model, at mean own)",
        "N": nonmeta["CountyName"].nunique(),
        "coef": c_met_i, "se": s_met_i, "p": p_met_i, "stars": stars(p_met_i),
    })
    demand_rows.append({
        "group": "Interaction (own WFH x region metro WFH)",
        "N": nonmeta["CountyName"].nunique(),
        "coef": c_int, "se": s_int, "p": p_int, "stars": stars(p_int),
    })
    demand_rows.append({
        "group": "Own WFH marginal effect at LOW region metro (p25)",
        "N": nonmeta["CountyName"].nunique(),
        "coef": me_low, "se": np.nan, "p": np.nan, "stars": "",
    })
    demand_rows.append({
        "group": "Own WFH marginal effect at HIGH region metro (p75)",
        "N": nonmeta["CountyName"].nunique(),
        "coef": me_high, "se": np.nan, "p": np.nan, "stars": "",
    })
pd.DataFrame(demand_rows).to_csv(RES / "cbsa_demand_results.csv", index=False)
print("  Saved: cbsa_demand_results.csv")

# ══════════════════════════════════════════════════════════════════════════════
# 7. ROBUSTNESS CHECKS
# ══════════════════════════════════════════════════════════════════════════════
print_sep("7. Robustness Checks — Metro-Core")

rob_rows = []

# R1 — Main +Income (already done)
rob_rows.append({"spec": "R1: Main TWFE (+Income)", "coef": main_c,
                 "se": main_s, "p": main_p, "stars": stars(main_p)})

# R2 — No controls
c_r2 = twfe_rows[0]["coef"]; s_r2 = twfe_rows[0]["se"]; p_r2 = twfe_rows[0]["p"]
rob_rows.append({"spec": "R2: TWFE (no controls)", "coef": c_r2,
                 "se": s_r2, "p": p_r2, "stars": stars(p_r2)})

# R3 — DH WAS (already done)
if was_inc:
    rob_rows.append({"spec": "R3: DH(2024) WAS +Income",
                     "coef": was_inc["was_overall"], "se": was_inc["was_overall_se"],
                     "p": was_inc["was_p"], "stars": stars(was_inc["was_p"])})

# R4 — Pre-determined WFH proxy (2019 WFH rate × post)
metro_r4 = metro[metro["year"] == 2019][["CountyName", "wfh_rate"]].rename(
    columns={"wfh_rate": "wfh_2019"})
metro_r4_panel = metro.merge(metro_r4, on="CountyName", how="left")
metro_r4_panel["post"] = (metro_r4_panel["year"] >= 2020).astype(int)
metro_r4_panel["wfh_2019_x_post"] = metro_r4_panel["wfh_2019"] * metro_r4_panel["post"]

fit_r4 = pf.feols(
    "log_price ~ wfh_2019_x_post + log_income | CountyName + year",
    data=metro_r4_panel, vcov={"CRV1": "CountyName"}
)
c_r4 = float(fit_r4.coef()["wfh_2019_x_post"])
s_r4 = float(fit_r4.se()["wfh_2019_x_post"])
p_r4 = pval_t(c_r4, s_r4)
rob_rows.append({"spec": "R4: Pre-determined WFH (2019×post)",
                 "coef": c_r4, "se": s_r4, "p": p_r4, "stars": stars(p_r4)})
print(f"  R4 pre-determined WFH: {c_r4*100:+.3f}%  p={p_r4:.4f}  {stars(p_r4)}")

# R5 — Frisch-Waugh: remove county-specific linear trends
metro_r5 = metro.copy()
metro_r5["t"] = metro_r5["year"] - REF_YEAR
# Partial out county-specific linear time trend from log_price and wfh_rate
for col in ["log_price", "wfh_rate"]:
    metro_r5[f"{col}_dm"] = (metro_r5
        .groupby("CountyName")
        .apply(lambda g: g[col] - np.polyval(
            np.polyfit(g["t"], g[col], 1), g["t"]))
        .reset_index(level=0, drop=True))

fit_r5 = pf.feols(
    "log_price_dm ~ wfh_rate_dm + log_income | CountyName + year",
    data=metro_r5, vcov={"CRV1": "CountyName"}
)
c_r5 = float(fit_r5.coef()["wfh_rate_dm"])
s_r5 = float(fit_r5.se()["wfh_rate_dm"])
p_r5 = pval_t(c_r5, s_r5)
rob_rows.append({"spec": "R5: Frisch-Waugh (detrended)",
                 "coef": c_r5, "se": s_r5, "p": p_r5, "stars": stars(p_r5)})
print(f"  R5 Frisch-Waugh detrended: {c_r5*100:+.3f}%  p={p_r5:.4f}  {stars(p_r5)}")

# R6 — Metro + Fringe (pooled metropolitan)
metro_fringe = county[county["cbsa_type"].isin(["Metro-Core", "Metro-Fringe"])].copy()
fit_r6 = pf.feols(
    "log_price ~ wfh_rate + log_income | CountyName + year",
    data=metro_fringe, vcov={"CRV1": "CountyName"}
)
c_r6 = float(fit_r6.coef()["wfh_rate"])
s_r6 = float(fit_r6.se()["wfh_rate"])
p_r6 = pval_t(c_r6, s_r6)
rob_rows.append({"spec": "R6: Metro-Core + Metro-Fringe",
                 "coef": c_r6, "se": s_r6, "p": p_r6, "stars": stars(p_r6)})
print(f"  R6 pooled metro+fringe: {c_r6*100:+.3f}%  p={p_r6:.4f}  {stars(p_r6)}")

rob_df = pd.DataFrame(rob_rows)
print(f"\n  {'Specification':<40} {'Coef (%)':>9} {'SE (%)':>8} {'p':>7}  Stars")
print("  " + "-" * 72)
for _, r in rob_df.iterrows():
    print(f"  {r['spec']:<40} {r['coef']*100:>+9.3f}  {r['se']*100:>8.3f}  "
          f"{r['p']:>7.4f}  {r['stars']}")

# ── Figure 6: Robustness forest plot
fig, ax = plt.subplots(figsize=(10, 6))
y_pos = list(range(len(rob_df)))
for i, (_, r) in enumerate(rob_df.iterrows()):
    col = RED if r["spec"].startswith("R1") else (GREEN if "WAS" in r["spec"] else GRAY)
    ax.errorbar(r["coef"] * 100, i,
                xerr=1.96 * r["se"] * 100,
                fmt="D" if r["spec"].startswith("R1") else "o",
                color=col, ms=8, capsize=4, lw=1.5)
    ax.text(r["coef"] * 100 - 0.02,
            i + 0.25,
            f"{r['coef']*100:+.3f}%{r['stars']}",
            va="bottom", ha="right", fontsize=8.5, color=col)

ax.axvline(0, color="black", lw=0.8, ls="--")
ax.axvline(main_c * 100, color=RED, lw=1.2, ls=":", alpha=0.6, label="Main estimate")
ax.set_yticks(y_pos)
ax.set_yticklabels(rob_df["spec"], fontsize=9)
ax.set_xlabel("Coefficient (% price change per 1pp WFH)", fontsize=10)
ax.set_title("Figure 6: Robustness — Metro-Core County TWFE\n"
             "All specifications show negative WFH effect on prices",
             fontsize=11)
ax.legend(fontsize=9, frameon=False)
ax.grid(axis="x", alpha=0.3)
plt.tight_layout()
fig.savefig(FIG / "fig6_cbsa_robustness.png", dpi=150, bbox_inches="tight")
plt.close()
print("\n  Saved: fig6_cbsa_robustness.png")

# ══════════════════════════════════════════════════════════════════════════════
# 8. MASTER RESULTS TABLE
# ══════════════════════════════════════════════════════════════════════════════
print_sep("8. Master Results Table")

main_rows = [
    {"col": "(1)", "method": "TWFE Baseline",
     "sample": "Metro-Core", "level": "County",
     "coef": twfe_rows[0]["coef"], "se": twfe_rows[0]["se"], "p": twfe_rows[0]["p"]},
    {"col": "(2)", "method": "TWFE +Income",
     "sample": "Metro-Core", "level": "County",
     "coef": main_c, "se": main_s, "p": main_p},
    {"col": "(3)", "method": "TWFE +Income +Unemp",
     "sample": "Metro-Core", "level": "County",
     "coef": twfe_rows[2]["coef"], "se": twfe_rows[2]["se"], "p": twfe_rows[2]["p"]},
    {"col": "(4)", "method": "DH(2024) WAS (no ctrl)",
     "sample": "Metro-Core", "level": "County",
     "coef": was_base["was_overall"] if was_base else np.nan,
     "se":   was_base["was_overall_se"] if was_base else np.nan,
     "p":    was_base["was_p"] if was_base else np.nan},
    {"col": "(5)", "method": "DH(2024) WAS +Income",
     "sample": "Metro-Core", "level": "County",
     "coef": was_inc["was_overall"] if was_inc else np.nan,
     "se":   was_inc["was_overall_se"] if was_inc else np.nan,
     "p":    was_inc["was_p"] if was_inc else np.nan},
    {"col": "(6)", "method": "Supply: Constrained",
     "sample": "Metro-Core High P/I", "level": "County",
     "coef": c_con, "se": s_con, "p": p_con},
    {"col": "(7)", "method": "Supply: Elastic",
     "sample": "Metro-Core Low P/I", "level": "County",
     "coef": c_ela, "se": s_ela, "p": p_ela},
    {"col": "(8)", "method": "Non-Metro (demand null)",
     "sample": "Non-Metro", "level": "County",
     "coef": c_nm, "se": s_nm, "p": p_nm},
]
for r in main_rows:
    r["stars"] = stars(r["p"])
res_df = pd.DataFrame(main_rows)

print(f"\n  {'Col':<4} {'Method':<28} {'Sample':<22} {'Coef (%)':>9} {'SE (%)':>8} {'p':>7}  Stars")
print("  " + "-" * 86)
for _, r in res_df.iterrows():
    pv = r['p'] if not pd.isna(r['p']) else float('nan')
    print(f"  {r['col']:<4} {r['method']:<28} {r['sample']:<22} "
          f"{r['coef']*100:>+9.3f}  {r['se']*100:>8.3f}  {pv:>7.4f}  {r['stars']}")

res_df.to_csv(RES / "cbsa_main_results.csv", index=False)
rob_df.to_csv(RES / "cbsa_robustness_results.csv", index=False)
print("\n  Saved: cbsa_main_results.csv, cbsa_robustness_results.csv")

print_sep("DONE")
sep = "-" * 61
print(f"\n  Key Results:")
print(f"  {sep}")
print(f"  Main TWFE (+Income):   {main_c*100:+.3f}%   p={main_p:.4f}  {stars(main_p)}")
was_c_str  = f"{was_base['was_overall']*100:+.3f}" if was_base else "  n/a "
was_p_str  = f"{was_base['was_p']:.4f}"            if was_base else "  n/a "
was_st_str = stars(was_base['was_p'])               if was_base else ""
print(f"  DH(2024) WAS:          {was_c_str}%   p={was_p_str}  {was_st_str}")
print(f"  Pre-trend test:        chi2={chi2_pre:.3f}  p={p_pre:.4f}  {'PASS' if p_pre > 0.1 else 'CONCERN'}")
print(f"  Supply: Constrained:   {c_con*100:+.3f}%   p={p_con:.4f}  {stars(p_con)}")
print(f"  Supply: Elastic:       {c_ela*100:+.3f}%   p={p_ela:.4f}  {stars(p_ela)}")
print(f"  Supply: Difference:    {(c_con-c_ela)*100:+.3f}%   p={diff_p:.4f}  {stars(diff_p)}")
print(f"  Non-Metro (null):      {c_nm*100:+.3f}%   p={p_nm:.4f}  {stars(p_nm)}")
print(f"  {sep}")
