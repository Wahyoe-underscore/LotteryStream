# Sistem Undian Move & Groove

## Overview
A Streamlit-based lottery system for the December 7th Move & Groove event with 3 distinct lottery modes. The application uses a state machine architecture for clear navigation between different lottery stages.

## Lottery Flow (State Machine)

### 1. Homepage (`home`)
- Upload data via CSV file or Google Sheets URL
- Shows 3 lottery mode panels (E-Voucher, Shuffle, Wheel)
- E-Voucher always enabled
- Shuffle & Wheel disabled until E-Voucher is completed

### 2. E-Voucher Mode (`evoucher_preview` → `evoucher_results`)
- **Preview**: Shows 4 prize categories (Tokopedia, Indomaret, Bensin, SNL)
- **Run Lottery**: 700 winners selected from eligible participants
- **Results**: View winners by category, download Excel & PPT
- **"Sisa Nomor"**: Returns to main menu, enables Shuffle/Wheel modes

### 3. Shuffle Mode (`shuffle_page`)
- 3 sessions of 30 winners each
- Each session requires prize name input
- Download Excel & PPT per session
- "Sisa Nomor" returns to main menu

### 4. Wheel Mode (`wheel_page`)
- 10 Grand Prizes with structured table configuration (like E-Voucher/Shuffle)
- Editable data table with "No", "Nama Hadiah", "Keterangan" columns
- Preview cards showing configured prizes
- Spinning wheel animation for each prize
- Download results when complete

## Prize Tiers (E-Voucher - 700 Total)
| Category | Winners |
|----------|---------|
| Tokopedia Rp.100.000,- | 175 |
| Indomaret Rp.100.000,- | 175 |
| Bensin Rp.100.000,- | 175 |
| SNL Rp.100.000,- | 175 |

## Features
- CSV upload or Google Sheets URL for participant data
- VIP/F participants automatically excluded
- Secure randomization using `secrets` module
- Phone numbers masked on display (****1234)
- Excel export with full details (all pages)
- PowerPoint export for presentation (all pages)
- Remaining participants tracked across all lottery stages
- "Nomor yang Belum Diundi" expander on each result page
- 8 winner result buttons on main page (4 E-Voucher + 3 Shuffle + 1 Wheel)
- MD5 hash-based content change detection for reliable data source tracking

## Project Structure
- `app.py` - Main Streamlit application with state machine
- `prize_config.json` - Saved prize configuration
- `.streamlit/config.toml` - Streamlit server configuration
- `attached_assets/` - Banner images

## State Management
Key session state variables:
- `current_page`: Controls which page to display
- `participant_data`: Full DataFrame of all participants
- `eligible_participants`: List of eligible lottery numbers
- `remaining_pool`: DataFrame of remaining participants (persisted across draws)
- `evoucher_done`: Flag to enable Shuffle/Wheel modes
- `evoucher_results`: E-Voucher lottery results DataFrame
- `shuffle_results`: Dictionary of shuffle batch results
- `wheel_winners`: List of wheel prize winners

## Running the Application
```bash
streamlit run app.py --server.port 5000
```

## CSV Format
```csv
Nomor Undian,Nama,No HP
0001,John Doe,081234567890
0002,Jane Smith,082345678901
```

## Dependencies
- streamlit
- pandas
- openpyxl
- python-pptx
