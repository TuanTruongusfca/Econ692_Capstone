# Data Directory — WFH & California Housing Prices

## Overview

This directory contains all raw and processed data for the Econ692 Capstone project
analyzing the effect of remote work on California county housing prices (2017–2023).

---

## Directory Structure

```
Data/
├── README.md                          ← this file
├── raw/                               ← original downloaded source files
│   ├── zillow/
│   │   └── Zip_zhvi_sfr_condo_tier_mid_2000_2024.zip
│   ├── acs_commute/
│   │   └── acs_s0801_commute_wfh_zip_2013_2023.zip
│   ├── acs_population/
│   │   └── acs_b01003_population_zip_2013_2023.zip
│   ├── acs_income/
│   │   └── acs_b19013_median_income_zip_2013_2023.zip
│   ├── bls_laus/
│   │   └── README.txt                 ← download instructions for BLS LAUS
│   └── cbsa/
│       └── cbsa_delineation_2023.xlsx
├── processed/                         ← intermediate and final processed files
│   ├── zip_panel_ca_2013_2023.csv     ← original zip-level panel (all years, reference)
│   └── zip_panel_ca_2017_2023.csv     ← zip-level panel for study period (built by build_dataset.py)
├── CA_Housing_WFH_Dataset.xlsx        ← reference workbook (codebook, summary stats, regression)
└── Dataset_final.xlsx                 ← zip-level panel with corrected WFH rates (reference)
```

The **final analysis dataset** is `Output/data/county_dataset.csv`.
It is built from these raw sources by `Output/scripts/build_dataset.py`.

---

## Data Sources

### 1. Zillow ZHVI — Home Values by ZIP Code
| Field        | Value |
|-------------|-------|
| File        | `raw/zillow/Zip_zhvi_sfr_condo_tier_mid_2000_2024.zip` |
| Source      | Zillow Research — https://www.zillow.com/research/data/ |
| Table       | Zip_zhvi_uc_sfrcondo_tier_0.33_0.67_sm_sa_month |
| Coverage    | All U.S. ZIP codes, monthly, January 2000 – present |
| Price tier  | Middle tier (33rd–67th percentile), SFR + condo, smoothed & seasonally adjusted |
| Format      | Wide CSV: one column per month (`2000-01-31`, …); annual price = mean of 12 months |

### 2. ACS S0801 — Commuting Characteristics by Sex (ZIP/ZCTA level)
| Field        | Value |
|-------------|-------|
| File        | `raw/acs_commute/acs_s0801_commute_wfh_zip_2013_2023.zip` |
| Source      | U.S. Census Bureau ACS 5-Year Estimates — https://data.census.gov |
| Table       | S0801 "Commuting Characteristics by Sex" |
| Geography   | ZIP Code Tabulation Areas (ZCTAs) |
| Years       | 2013–2023 (one CSV per ACS vintage, e.g. `ACSST5Y2021.S0801-Data.csv`) |
| Key columns | `S0801_C01_001E` = Total workers 16+ (count) |
|             | `S0801_C01_013E` = Worked from home (**percentage**, 0–100) |
| Note        | In ACS Subject Tables (S-tables), transportation mode rows report **percentages**. |
|             | WFH count = (S0801_C01_013E / 100) × S0801_C01_001E |

### 3. ACS B01003 — Total Population (ZIP/ZCTA level)
| Field        | Value |
|-------------|-------|
| File        | `raw/acs_population/acs_b01003_population_zip_2013_2023.zip` |
| Source      | U.S. Census Bureau ACS 5-Year Estimates |
| Table       | B01003 "Total Population" |
| Key column  | `B01003_001E` = Total population (count) |

### 4. ACS B19013 — Median Household Income (ZIP/ZCTA level)
| Field        | Value |
|-------------|-------|
| File        | `raw/acs_income/acs_b19013_median_income_zip_2013_2023.zip` |
| Source      | U.S. Census Bureau ACS 5-Year Estimates |
| Table       | B19013 "Median Household Income in the Past 12 Months" |
| Key column  | `B19013_001E` = Median household income in dollars (count) |

### 5. BLS LAUS — Local Area Unemployment Statistics (county level)
| Field        | Value |
|-------------|-------|
| File        | `raw/bls_laus/bls_laus_ca_counties_2013_2023.csv` (**download separately**) |
| Source      | U.S. Bureau of Labor Statistics — https://www.bls.gov/lau/tables.htm |
| Table       | Table 14: "Counties, Annual Averages, Not Seasonally Adjusted" |
| Note        | See `raw/bls_laus/README.txt` for download instructions |

### 6. OMB CBSA Delineation File — County Classification
| Field        | Value |
|-------------|-------|
| File        | `raw/cbsa/cbsa_delineation_2023.xlsx` |
| Source      | U.S. Census Bureau OMB Delineations (2023) |
| URL         | https://www.census.gov/geographies/reference-files/time-series/demo/metro-micro/delineation-files.html |
| Used for    | Classifying all 58 CA counties as Metro-Core / Metro-Fringe / Micropolitan / Noncore |

---

## Final Dataset — `Output/data/county_dataset.csv`

**406 observations: 58 California counties × 7 years (2017–2023)**

| Column             | Type    | Description |
|--------------------|---------|-------------|
| `CountyName`       | string  | California county name (e.g. "Alameda County") |
| `year`             | int     | Calendar year (2017–2023) |
| `log_price`        | float   | Log of population-weighted median home price ($) across ZIPs |
| `wfh_rate`         | float   | WFH workers / total workers × 100 (county-level aggregate, %) |
| `median_income`    | float   | Population-weighted median household income ($) across ZIPs |
| `unemployment_rate`| float   | Civilian unemployment rate (fraction), from BLS LAUS |
| `population`       | float   | Sum of ZIP-level population within county |
| `n_zips`           | int     | Number of ZIPs with valid Zillow price data in county × year |

---

## How to Rebuild

Run `Output/scripts/build_dataset.py` from the project root:

```bash
cd Econ692_Capstone/Output
python scripts/build_dataset.py
```

This reads from `Data/raw/` and writes:
- `Data/processed/zip_panel_ca_2017_2023.csv` (zip-level intermediate)
- `Output/data/county_dataset.csv` (county-level final dataset)

**Prerequisite**: Download BLS LAUS data first — see `Data/raw/bls_laus/README.txt`.

---

## ACS 5-Year Estimate Notes

All ACS variables are **5-year rolling estimates**, not single-year snapshots:
- The "2021" ACS estimate covers 2017–2021
- The "2023" ACS estimate covers 2019–2023

This means post-2020 WFH rates are smoothed (partially diluting the COVID spike),
and the treatment timing in the TWFE model is inherently attenuated toward zero.
