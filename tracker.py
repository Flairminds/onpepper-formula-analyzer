# tracker.py
"""Entry point script for the Excel Formula Analyzer.
It loads a workbook, lists sheets, prints basic metadata, and stores a single record per workbook.
"""

import os
import sys
import re
import datetime
from dotenv import load_dotenv

# Load environment variables from the .env file located in the config directory
load_dotenv(dotenv_path=os.path.join(os.path.dirname(__file__), 'config', '.env'))

# Internal modules
from src.reader import load_excel_workbook, list_sheets, get_sheet_metadata
from src.mapping import get_business_column_mapping, build_mapping_metadata
from src.formula_extractor import extract_formulas
from src.persist import save_workbook_record, get_record_count
from src.dependency_extractor import enrich_payload_with_dependencies
from src.formula_normalizer import normalize_all_formulas
from src.ast_builder import build_all_asts


def main():
    if len(sys.argv) != 2:
        print("Usage: python tracker.py <path-to-workbook.xlsx>")
        sys.exit(1)
    workbook_path = sys.argv[1]

    # Load workbook (data_only=False so formulas are retained)
    try:
        wb = load_excel_workbook(workbook_path)
    except Exception as e:
        print(f"Error loading workbook: {e}")
        sys.exit(1)

    print(f"Workbook loaded: {workbook_path}")
    sheets = list_sheets(wb)
    print(f"Sheets ({len(sheets)}): {', '.join(sheets)}")

    # Gather column mapping once – applies to the whole workbook
    column_mapping = get_business_column_mapping(wb)

    # Show metadata for each sheet (no DB writes yet)
    for sheet_name in sheets:
        meta = get_sheet_metadata(wb, sheet_name)
        sheet_map = column_mapping.get(sheet_name, {})
        print(f"\nSheet: {sheet_name}")
        print(f"  Max Row: {meta['max_row']}, Max Column: {meta['max_column']}")
        print(f"  Headers: {meta['headers']}")
        print(f"  Column Mapping: {sheet_map}")

    # After all sheets processed, extract formulas and persist to DB
    filename = os.path.basename(workbook_path)
    version = ""
    formulas = extract_formulas(wb)
    # Add a human‑readable upload timestamp (local time)
    uploaded_at = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    # Build the full metadata structure (includes column mapping, formulas, and resolved references)
    payload = build_mapping_metadata(
        file_name=filename,
        version=version,
        mapping=column_mapping,
        formulas=formulas,
    )

    # Phase 3: attach dependency graph and summary
    payload = enrich_payload_with_dependencies(payload, wb, formulas)

    # Phase 4: attach normalized formulas (cell refs replaced with business names)
    payload["normalized_formulas"] = normalize_all_formulas(formulas, column_mapping)

    # Phase 5: attach ASTs (formula structure trees + complexity stats)
    payload["formula_asts"] = build_all_asts(formulas)
    # Attach the upload timestamp
    payload["uploaded_at"] = uploaded_at
    save_workbook_record(filename, version, payload)
    print(f"[DB] Record saved for {filename} version '{version}'")
    print(f"Total records now: {get_record_count()}")

if __name__ == "__main__":
    main()
