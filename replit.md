# Sistem Undian Move & Groove

## Overview
A Streamlit-based lottery system for selecting 900 winners from participant data. The application securely randomizes participants and assigns prizes across 9 tiers.

## Features
- CSV file upload for participant data (requires "Nomor Undian" column)
- Secure randomization using Python's `secrets` module (cryptographically secure)
- 9 prize tiers with 100 winners each (total 900 winners)
- Prize summary and detailed winners table
- CSV download functionality for results

## Prize Tiers
| Rank | Prize |
|------|-------|
| 1-100 | Hadiah 1 (Kulkas) |
| 101-200 | Hadiah 2 (TV) |
| 201-300 | Hadiah 3 (Mesin Cuci) |
| 301-400 | Hadiah 4 (Microwave) |
| 401-500 | Hadiah 5 (Blender) |
| 501-600 | Hadiah 6 (Rice Cooker) |
| 601-700 | Hadiah 7 (Setrika) |
| 701-800 | Hadiah 8 (Kipas Angin) |
| 801-900 | Hadiah 9 (Voucher Belanja) |

## Project Structure
- `app.py` - Main Streamlit application
- `.streamlit/config.toml` - Streamlit server configuration

## Running the Application
```bash
streamlit run app.py --server.port 5000
```

## CSV Format
The uploaded CSV file must contain a column named "Nomor Undian":
```csv
Nomor Undian
001
002
003
...
```

## Dependencies
- streamlit
- pandas
