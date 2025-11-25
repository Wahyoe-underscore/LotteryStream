# Sistem Undian Move & Groove

## Overview
A Streamlit-based lottery system for selecting 900 winners from participant data. The application securely randomizes participants and assigns prizes across 9 tiers.

## Features
- CSV file upload for participant data (requires "Nomor Undian" column)
- Secure randomization using Python's `secrets` module (cryptographically secure)
- 9 prize tiers with varying winners (total 775 winners)
- Prize summary and detailed winners table
- CSV download functionality for results

## Prize Tiers (Total 775 Winners)
| Rank | Prize | Winners |
|------|-------|---------|
| 1-75 | Bensin Rp.100.000,- | 75 |
| 76-175 | Top100 Rp.100.000,- | 100 |
| 176-250 | SNL Rp.100.000,- | 75 |
| 251-325 | Bensin Rp.150.000,- | 75 |
| 326-400 | Top100 Rp.150.000,- | 75 |
| 401-475 | SNL Rp.150.000,- | 75 |
| 476-575 | Bensin Rp.200.000,- | 100 |
| 576-675 | Top100 Rp.200.000,- | 100 |
| 676-775 | SNL Rp.200.000,- | 100 |

## Project Structure
- `app.py` - Main Streamlit application
- `.streamlit/config.toml` - Streamlit server configuration

## Running the Application
```bash
streamlit run app.py --server.port 5000
```

## CSV Format
The uploaded CSV file must contain a column named "Nomor Undian" (4-digit format with leading zeros preserved):
```csv
Nomor Undian
0001
0002
0003
...
```

## Output Format
Results are exported as Excel file (.xlsx) with properly separated columns.

## Dependencies
- streamlit
- pandas
