# Digital Lottery System

A professional digital lottery application built with Python and Streamlit, designed for corporate events. Successfully deployed and used at a client's "Move & Groove" event in December 2025.

## Features

### Three Integrated Lottery Modes

**1. E-Voucher Draw**
- Handles 700 prizes across 4 categories
- Batch winner selection with animated display
- Automatic prize tier distribution

**2. Shuffle Mode**
- Slot machine-style cascade animation
- 3 sessions with 30 winners each
- Real-time visual feedback for audience engagement

**3. Spinning Wheel**
- Grand prize selection with rotating wheel animation
- 10 customizable prize slots
- "Void" and "Redraw" features for live event flexibility

### Additional Features

- **Quick Draw (Undian Cepat)** - Single winner selection for spontaneous draws
- **Backup Draw (Undian Cadangan)** - Reserve winners for each batch
- **Full Transparency** - All participant numbers displayed during animations
- **Smart Filtering** - Automatically excludes invalid entries
- **Export Options** - One-click export to Excel and PowerPoint
- **Data Sources** - Support for CSV upload and Google Sheets integration

## Tech Stack

- **Python 3.11**
- **Streamlit** - Web application framework
- **Pandas** - Data manipulation and analysis
- **OpenPyXL** - Excel file generation
- **python-pptx** - PowerPoint presentation generation

## Installation

1. Clone the repository:
```bash
git clone https://github.com/yourusername/digital-lottery-system.git
cd digital-lottery-system
```

2. Install dependencies:
```bash
pip install streamlit pandas openpyxl python-pptx requests
```

3. Run the application:
```bash
streamlit run app.py --server.port 5000
```

## Usage

### Data Format

Prepare your participant data in CSV format with the following columns:

| Column | Description |
|--------|-------------|
| Nomor Undian | Unique lottery number |
| Nama | Participant name |
| No HP | Phone number |

Example:
```csv
Nomor Undian,Nama,No HP
0001,John Doe,081234567890
0002,Jane Smith,082345678901
```

### Exclusion Rules

The system automatically excludes entries with:
- Lottery numbers containing "D"
- Name or phone fields with exactly "F", "D", or "VIP"
- Empty name AND phone fields

### Running a Lottery

1. Upload your participant data (CSV or Google Sheets URL)
2. Select lottery mode (E-Voucher, Shuffle, or Wheel)
3. Configure prizes as needed
4. Run the lottery and watch the animation
5. Export results to Excel or PowerPoint

## Screenshots

*Add screenshots of your application here*

## Live Demo

Try the live demo on Replit: [Digital Lottery System](https://your-replit-url.replit.app)

## Benefits

- **Transparency** - All participants visible during draws, building trust
- **Time-Efficient** - Automated selection and instant results
- **Professional** - Ready-to-use PowerPoint exports for presentations
- **Flexible** - Void and redraw options for live event situations
- **Reliable** - Smart data validation and error handling

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Author

Built with passion for creating engaging event experiences.

---

*Successfully deployed at Move & Groove corporate event, December 2025*
