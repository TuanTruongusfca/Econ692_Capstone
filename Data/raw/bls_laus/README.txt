BLS Local Area Unemployment Statistics (LAUS) — County-Level Annual Data
==========================================================================

Source : U.S. Bureau of Labor Statistics
URL    : https://www.bls.gov/lau/tables.htm
Table  : Table 14 – "Counties, Annual Averages, Not Seasonally Adjusted"
Years  : 2013–2023 (download all years needed)

Steps to download
-----------------
1. Go to https://www.bls.gov/lau/tables.htm
2. Click "Table 14. Counties, annual averages, not seasonally adjusted"
3. Download as .xlsx or copy the data
4. Save as: bls_laus_ca_counties_2013_2023.csv

Expected CSV format (after filtering California):
  area_fips, area_title, year, labor_force, employed, unemployed, unemployment_rate
  06001, Alameda County CA, 2017, 845900, 795700, 50200, 0.059
  ...

Notes
-----
- Filter to state_fips = 06 (California) and county-level records
- unemployment_rate = unemployed / labor_force (already computed by BLS)
- The `unemployment_rate` column in county_dataset.csv uses this data
- Not included in repo due to file size; download directly from BLS
