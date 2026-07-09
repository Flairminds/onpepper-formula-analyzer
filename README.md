# Excel Formula Analyzer

This repository contains a Python project for extracting Excel workbook metadata and formulas, storing them in PostgreSQL, and preparing for further analysis.

## Project Structure
```
Excel_Formula_Analyser/
├─ README.md               # Project overview
├─ requirements.txt        # Python dependencies
├─ init_schema.sql         # Initial PostgreSQL schema
├─ .gitignore              # Ignored files
├─ config/
│   └─ .env               # Environment variables (DB connection etc.)
├─ src/
│   ├─ __init__.py
│   └─ extractor.py       # Functions to read workbooks using openpyxl
└─ tracker.py              # Entry‑point script
```

## Getting Started
1. Install dependencies:
   ```
   pip install -r requirements.txt
   ```
2. Create a PostgreSQL database and run `init_schema.sql` to set up tables.
3. Copy `.env.example` to `.env` and fill in your database credentials.
4. Run the tracker script:
   ```
   python tracker.py <path-to-workbook.xlsx>
   ```

The script will print basic workbook information and extracted formulas.
