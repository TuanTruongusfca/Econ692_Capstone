# Replication Package
## Work-From-Home and Metropolitan Housing Deflation in California

**Author:** Tuan Truong  
**Course:** ECON 692 Capstone — University of San Francisco  
**Date:** May 2026

---

## Overview

This package replicates all tables, figures, and the final paper in:

> Truong, T. (2026). *Work-From-Home and Metropolitan Housing Deflation in California.* ECON 692 Capstone, University of San Francisco.

The analysis uses a county-level panel of 58 California counties from 2017 to 2023 to estimate the causal effect of remote work adoption on housing prices via two-way fixed effects (TWFE).

---

## Folder Structure

```
Econ692_Capstone/
├── README.md                           <- this file
├── requirements.txt                    <- Python dependencies
│
├── Data/
│   ├── Dataset_final.xlsx              <- PRIMARY: ZIP x year panel (CA, 2013-2023)
│   ├── README.md                       <- data source documentation
│   └── raw/
│       ├── zillow/                     <- Zillow ZHVI ZIP-level monthly prices
│       ├── acs_commute/                <- ACS S0801 WFH rates by ZIP x year
│       ├── acs_income/                 <- ACS B19013 median income by ZIP x year
│       ├── acs_population/             <- ACS B01003 population by ZIP x year
│       ├── bls_laus/README.txt         <- BLS LAUS (download instructions)
│       └── cbsa/                       <- OMB CBSA 2023 county delineations
│
├── Output/
│   ├── scripts/
│   │   ├── 01_build_dataset.py         <- Step 1: build county_dataset.csv
│   │   ├── 02_county_cbsa_analysis.py  <- Step 2: regressions + figures
│   │   └── 03_make_paper.py            <- Step 3: generate Final Paper.docx
│   ├── data/
│   │   └── county_dataset.csv          <- pre-built county panel (406 obs)
│   ├── figures/                        <- all figures (pre-generated)
│   └── results/                        <- all result CSVs (pre-generated)
│
└── Paper/
    └── Final Paper.docx                <- final submission paper
```

---

## How to Reproduce

### 0. Install dependencies

```bash
pip install -r requirements.txt
```

Python 3.10+ required.

### 1. Build the county dataset

```bash
python Output/scripts/01_build_dataset.py
```

- Input:  `Data/Dataset_final.xlsx`
- Output: `Output/data/county_dataset.csv` (406 obs, 58 counties x 7 years)
- Note: BLS LAUS unemployment is not included due to file size. The script
  carries existing values. To refresh, follow `Data/raw/bls_laus/README.txt`.

### 2. Run all regressions and generate figures

```bash
python Output/scripts/02_county_cbsa_analysis.py
```

- Input:  `Output/data/county_dataset.csv`
- Output: `Output/figures/fig[1-6]_*.png`, `fig_twfe_metro_vs_nonmetro.png`
          `Output/results/cbsa_*.csv`

### 3. Generate the final paper

```bash
python Output/scripts/03_make_paper.py
```

- Input:  figures from `Output/figures/`
- Output: `Paper/Final Paper.docx`

---

## Data Sources

| Variable | Source | Location |
|---|---|---|
| Housing prices | Zillow ZHVI (ZIP, monthly) | `Data/raw/zillow/` |
| WFH rate | ACS S0801 (5-year estimates) | `Data/raw/acs_commute/` |
| Median income | ACS B19013 (5-year estimates) | `Data/raw/acs_income/` |
| Population | ACS B01003 (5-year estimates) | `Data/raw/acs_population/` |
| Unemployment | BLS LAUS (county, annual) | see `Data/raw/bls_laus/README.txt` |
| County classification | OMB CBSA 2023 Delineations | `Data/raw/cbsa/` |

All ZIP-level variables are pre-aggregated to county level in `Dataset_final.xlsx`
using population-weighted averages. Script 01 reads this file and applies
county-level aggregation with population weights.

---

## Key Results

| Specification | Sample | WFH Coef (%) | SE (%) | p-value |
|---|---|---|---|---|
| TWFE + Income | Metro-Core (33 counties) | -1.234 | 0.336 | 0.0002 *** |
| TWFE + Income | Non-Metro (21 counties) | -0.637 | 0.187 | 0.0007 *** |
| Supply: Constrained | Metro-Core High P/I | -1.195 | 0.322 | 0.0002 *** |
| Supply: Elastic | Metro-Core Low P/I | -0.402 | 0.353 | 0.254 |
| DH(2024) WAS + Income | Metro-Core | +1.201 | 4.601 | 0.788 |

Pre-trend Wald test: chi2(2) = 1.853, p = 0.396 (PASS).

---

## Software

- Python 3.10+
- Key packages: `pyfixest`, `pandas`, `numpy`, `matplotlib`, `scipy`,
  `python-docx`, `openpyxl`
- See `requirements.txt` for full list with version constraints.
