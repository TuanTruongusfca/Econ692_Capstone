"""
build_dataset.py
================
Builds county_dataset.csv (county × year panel) from source files.

DATA SOURCES
------------
Primary — processed ZIP-level panel:
  Data/Dataset_final.xlsx
    Columns used: County, Year, Median Home Price ($), WFH Rate (%),
                  Population, Median Income ($), Workers 16+, WFH Count (ACS)
    Built from:
      - Zillow ZHVI  (Data/raw/zillow/)   → home prices by ZIP × month
      - ACS S0801    (Data/raw/acs_commute/)    → WFH rate by ZIP × year
      - ACS B01003   (Data/raw/acs_population/) → population by ZIP × year
      - ACS B19013   (Data/raw/acs_income/)     → median income by ZIP × year
      - BLS LAUS     (Data/raw/bls_laus/)       → county unemployment rate
      - OMB CBSA     (Data/raw/cbsa/)           → county classification

OUTPUT
------
  Output/data/county_dataset.csv
  Columns:
    CountyName        : CA county name (e.g. "Alameda County")
    year              : Calendar year (2017–2023)
    log_price         : log(median home price), pop-weighted county mean
    wfh_rate          : WFH rate (%), pop-weighted county mean of ZIP rates
    median_income     : Median household income ($), pop-weighted county mean
    unemployment_rate : Civilian unemployment rate (fraction), from BLS LAUS
    population        : Total county population (sum across ZIPs)
    n_zips            : Number of ZIPs with valid Zillow price data

AGGREGATION NOTES
-----------------
  wfh_rate / median_income / population:
    Aggregated from ALL ZIPs in the county, including those without a
    Zillow price observation. This matches the original county_dataset.csv.

  log_price / n_zips:
    Aggregated from only ZIPs that have a valid Zillow price.

  Weighting: population-weighted average (ZIP population as weight).
"""

import warnings
from pathlib import Path

import numpy as np
import pandas as pd

warnings.filterwarnings("ignore")

# ── Paths ──────────────────────────────────────────────────────────────────────
ROOT = Path(__file__).resolve().parents[2]       # Econ692_Capstone/
DATA = ROOT / "Data"
OUT  = ROOT / "Output" / "data"

ZIP_PANEL = DATA / "Dataset_final.xlsx"
RAW_BLS   = DATA / "raw" / "bls_laus" / "bls_laus_ca_counties_2013_2023.csv"

STUDY_YEARS = list(range(2017, 2024))


def print_step(n: int, label: str) -> None:
    print(f"\n{'='*65}")
    print(f"  Step {n}: {label}")
    print("=" * 65)


def wavg(grp: pd.DataFrame, col: str, weight: str) -> float:
    """Population-weighted average; skips NaN values in col."""
    v = grp[col].dropna()
    w = grp.loc[v.index, weight]
    return float((v * w).sum() / w.sum()) if w.sum() > 0 else np.nan


# ══════════════════════════════════════════════════════════════════════════════
# Step 1: Load zip-level panel from Dataset_final.xlsx
# ══════════════════════════════════════════════════════════════════════════════
print_step(1, "Load Dataset_final.xlsx → ZIP-level panel")

df_raw = pd.read_excel(ZIP_PANEL, sheet_name="Sheet1")
df = df_raw.rename(columns={
    "County"               : "CountyName",
    "Year"                 : "year",
    "Median Home Price ($)": "price",
    "WFH Rate (%)"         : "wfh_pct",
    "Population"           : "population",
    "Median Income ($)"    : "median_income",
})

# Filter: California, study period
df = df[(df["State"] == "CA") & (df["year"].isin(STUDY_YEARS))].copy()
df["w"] = df["population"].fillna(1.0).clip(lower=1.0)    # population weight

print(f"  Rows: {len(df):,}  |  ZIPs: {df['Zip Code'].nunique():,}  "
      f"|  Counties: {df['CountyName'].nunique()}  |  Years: {df['year'].nunique()}")


# ══════════════════════════════════════════════════════════════════════════════
# Step 2: Aggregate to county level
# ══════════════════════════════════════════════════════════════════════════════
print_step(2, "Aggregate ZIP → county level")

# ── 2a. WFH rate, income, population — use ALL ZIPs (even those without price)
socio = (
    df.groupby(["CountyName", "year"])
    .apply(lambda g: pd.Series({
        "wfh_rate"     : wavg(g, "wfh_pct",      "w"),
        "median_income": wavg(g, "median_income", "w"),
        "population"   : float(g["w"].sum()),
    }), include_groups=False)
    .reset_index()
)

# ── 2b. log(price), n_zips — use only ZIPs with valid Zillow price
priced = df.dropna(subset=["price"]).copy()
priced["log_price"] = np.log(priced["price"].clip(lower=1))

prices = (
    priced.groupby(["CountyName", "year"])
    .apply(lambda g: pd.Series({
        "log_price": wavg(g, "log_price", "w"),
        "n_zips"   : int(g["Zip Code"].nunique()),
    }), include_groups=False)
    .reset_index()
)

# ── Merge
county_df = socio.merge(prices, on=["CountyName", "year"], how="inner")
county_df = county_df.sort_values(["CountyName", "year"]).reset_index(drop=True)

print(f"  County panel: {len(county_df):,} obs  "
      f"({county_df['CountyName'].nunique()} counties × {county_df['year'].nunique()} years)")
miss = county_df[["log_price", "wfh_rate", "median_income", "population"]].isna().mean()
print("  Missing rates:")
for col, rate in miss.items():
    print(f"    {col:<20} {rate:.1%}")


# ══════════════════════════════════════════════════════════════════════════════
# Step 3: Merge BLS LAUS unemployment (county × year)
# ══════════════════════════════════════════════════════════════════════════════
print_step(3, "Merge BLS LAUS unemployment rate")

if RAW_BLS.exists():
    bls = pd.read_csv(RAW_BLS)
    col_map: dict[str, str] = {}
    for c in bls.columns:
        cl = c.lower().replace(" ", "_")
        if "county" in cl or "area" in cl:
            col_map[c] = "CountyName"
        elif cl == "year":
            col_map[c] = "year"
        elif "unemp" in cl and "rate" in cl:
            col_map[c] = "unemployment_rate"
    bls = bls.rename(columns=col_map)[["CountyName", "year", "unemployment_rate"]].drop_duplicates()
    bls["CountyName"] = bls["CountyName"].str.strip()
    county_df = county_df.merge(bls, on=["CountyName", "year"], how="left")
    print(f"  Merged. Missing: {county_df['unemployment_rate'].isna().mean():.1%}")
else:
    existing = OUT / "county_dataset.csv"
    if existing.exists():
        unemp = pd.read_csv(existing, usecols=["CountyName", "year", "unemployment_rate"])
        county_df = county_df.merge(unemp, on=["CountyName", "year"], how="left")
        print("  BLS LAUS file not found — carried unemployment_rate from existing county_dataset.csv.")
        print("  To refresh: see Data/raw/bls_laus/README.txt")
    else:
        county_df["unemployment_rate"] = np.nan
        print("  WARNING: unemployment_rate set to NaN. See Data/raw/bls_laus/README.txt")


# ══════════════════════════════════════════════════════════════════════════════
# Step 4: Save
# ══════════════════════════════════════════════════════════════════════════════
print_step(4, "Save county_dataset.csv")

col_order = ["CountyName", "year",
             "log_price", "wfh_rate", "median_income",
             "unemployment_rate", "population", "n_zips"]
county_df = county_df[col_order]

out_path = OUT / "county_dataset.csv"
county_df.to_csv(out_path, index=False)
print(f"  Saved : {out_path}")
print(f"  Shape : {county_df.shape}")
print(f"\n  Summary:")
print(county_df[["log_price", "wfh_rate", "median_income",
                  "unemployment_rate"]].describe().round(4).to_string())

print("\n" + "=" * 65)
print("  Done. Next steps:")
print("    python Output/scripts/county_cbsa_analysis.py")
print("    python Output/scripts/dynamic_twfe_ca_nonmetro.py")
print("=" * 65)
