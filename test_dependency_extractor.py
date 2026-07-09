# test_dependency_extractor.py
"""
Simple sanity‑check for the dependency‑extraction stage.
Run with:
    python test_dependency_extractor.py
"""

import json
from pathlib import Path

from openpyxl import load_workbook

# Local project imports – adjust if your package layout changes
from src.mapping import get_business_column_mapping, build_mapping_metadata
from src.formula_extractor import extract_formulas
from src.dependency_extractor import enrich_payload_with_dependencies

# ---------------------------------------------------------------------------
# Configuration – point to an existing workbook that contains a few formulas
# ---------------------------------------------------------------------------
WORKBOOK_PATH = Path(r"E:\Excel_Formula_Analyser\PFLT Sub - Borrowing Base - 07.31.25 ME Truist.xlsx")

# Verify the workbook file exists before proceeding
if not WORKBOOK_PATH.is_file():
    raise FileNotFoundError(
        f"Workbook not found at {WORKBOOK_PATH!s}.\n"
        "Please place a .xlsx file at this location or update WORKBOOK_PATH "
        "in test_dependency_extractor.py to point to an existing workbook."
    )

def main():
    # 1️⃣ Load workbook (data_only=False to keep the formula strings)
    wb = load_workbook(WORKBOOK_PATH, data_only=False)

    # 2️⃣ Extract formulas from every sheet
    formulas = extract_formulas(wb)

    # 3️⃣ Build the column‑header mapping (header row defaults to 9)
    mapping = get_business_column_mapping(wb)

    # 4️⃣ Assemble base payload – version left blank per user request
    payload = build_mapping_metadata(
        file_name=WORKBOOK_PATH.name,
        version="",
        mapping=mapping,
        formulas=formulas,
    )

    # 5️⃣ Enrich payload with dependency graph & summary
    payload = enrich_payload_with_dependencies(payload, wb, formulas)

    # --------------------------------------------------------------------
    # Show a concise view – enough to confirm the extractor works
    # --------------------------------------------------------------------
    print("\n=== Dependency Graph (sample) ===")
    graph = payload["dependency_graph"]
    for sheet, cells in list(graph.items())[:3]:  # up to 3 sheets
        print(f"\nSheet: {sheet}")
        for cell, node in list(cells.items())[:5]:  # up to 5 cells per sheet
            print(f"  {cell}: {json.dumps(node, indent=2)}")

    print("\n=== Dependency Summary ===")
    print(json.dumps(payload["dependency_summary"], indent=2))

    # Optional: dump full payload for manual inspection
    out_file = Path("dependency_test_output.json")
    out_file.write_text(json.dumps(payload, indent=2))
    print(f"\nFull payload written to {out_file.resolve()}")

if __name__ == "__main__":
    main()
