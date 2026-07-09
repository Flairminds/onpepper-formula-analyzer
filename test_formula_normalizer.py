# test_formula_normalizer.py
"""
Simple test script to verify the formula normalizer.
Run with:
    python test_formula_normalizer.py [workbook_path]
If no path is given, it defaults to the same workbook used in other tests.
"""

import json
import sys
from pathlib import Path

from openpyxl import load_workbook

# Local imports
from src.formula_extractor import extract_formulas
from src.mapping import get_business_column_mapping
from src.formula_normalizer import normalize_all_formulas

DEFAULT_WORKBOOK = Path(r"E:\\Excel_Formula_Analyser\\PFLT Sub - Borrowing Base - 07.31.25 ME Truist.xlsx")


def main(workbook_path: Path) -> None:
    if not workbook_path.is_file():
        raise FileNotFoundError(f"Workbook not found at {workbook_path}")

    wb = load_workbook(workbook_path, data_only=False)
    formulas = extract_formulas(wb)
    column_mapping = get_business_column_mapping(wb)
    normalized = normalize_all_formulas(formulas, column_mapping)

    # Print a tiny preview
    print("\n=== Normalized Formulas (preview) ===")
    for sheet, cells in list(normalized.items())[:2]:
        print(f"\nSheet: {sheet}")
        for cell, formula in list(cells.items())[:3]:
            print(f"  {cell}: {formula}")

    out_file = Path("normalized_formulas_preview.json")
    out_file.write_text(json.dumps(normalized, indent=2))
    print(f"\nFull normalized data written to {out_file.resolve()}")


if __name__ == "__main__":
    arg_path = Path(sys.argv[1]) if len(sys.argv) > 1 else DEFAULT_WORKBOOK
    main(arg_path)
