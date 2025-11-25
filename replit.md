# Sistem Undian Move & Groove

## Overview
A Streamlit-based lottery system for selecting 800 winners from participant data. The application securely randomizes participants and assigns prizes across 9 tiers. Optimized for large screen presentation at events.

## Features
- CSV file upload for participant data (requires "Nomor Undian" and "No HP" columns)
- Secure randomization using Python's `secrets` module (cryptographically secure)
- 9 prize tiers with varying winners (total 800 winners)
- Winner display in 10-column grid format (optimized for large screens)
- Phone numbers matched to winners automatically
- Privacy: Phone numbers are masked on display (shows ****1234 format)
- Full phone numbers available in Excel export for admin use
- Excel and PowerPoint export functionality
- Must download both Excel and PowerPoint before starting new lottery

## Prize Tiers (Total 800 Winners)
| Rank | Prize | Winners |
|------|-------|---------|
| 1-75 | Bensin Rp.100.000,- | 75 |
| 76-175 | Top100 Rp.100.000,- | 100 |
| 176-250 | SNL Rp.100.000,- | 75 |
| 251-325 | Bensin Rp.150.000,- | 75 |
| 326-400 | Top100 Rp.150.000,- | 75 |
| 401-500 | SNL Rp.150.000,- | 100 |
| 501-600 | Bensin Rp.200.000,- | 100 |
| 601-700 | Top100 Rp.200.000,- | 100 |
| 701-800 | SNL Rp.200.000,- | 100 |

## Project Structure
- `app.py` - Main Streamlit application
- `.streamlit/config.toml` - Streamlit server configuration
- `attached_assets/` - Banner images for the event

## Running the Application
```bash
streamlit run app.py --server.port 5000
```

## CSV Format
The uploaded CSV file must contain two columns:
```csv
Nomor Undian,No HP
0001,081234567890
0002,082345678901
0003,083456789012
...
```
- **Nomor Undian**: 4-digit lottery number (leading zeros preserved)
- **No HP**: Phone number for voucher delivery

## Output Formats
1. **Excel (.xlsx)**: Complete winner list with full phone numbers (for admin/voucher distribution)
2. **PowerPoint (.pptx)**: Presentation slides with gradient design for event display (no phone numbers)

## Privacy & Security
- Phone numbers are masked on the large screen display (shows ****1234)
- Full phone numbers are only available in the Excel export
- This prevents accidental exposure of personal data during public presentation

## Dependencies
- streamlit
- pandas
- openpyxl
- python-pptx
