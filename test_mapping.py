# test_mapping.py  (place next to the project root)
import os
from openpyxl import load_workbook
from src.mapping import (
    get_business_column_mapping,
    resolve_formula_references,
    validate_mappings,
    build_mapping_metadata,
)
from src.formula_extractor import extract_formulas   # existing extractor


def load_wb(path):
    return load_workbook(path, data_only=False)


if __name__ == "__main__":
    # 1️⃣ Load two versions of the same workbook (or any two files)
    wb_v1 = load_wb(r"E:\Excel_Formula_Analyser\PFLT Sub - Borrowing Base - 07.31.25 ME Truist.xlsx")
    wb_v2 = load_wb(r"E:\Excel_Formula_Analyser\PFLT Sub - Borrowing Base - 08.27.2025 vBorrow_4_TruistSend.xlsx")

    # 2️⃣ Build column mappings
    map_v1 = get_business_column_mapping(wb_v1)
    map_v2 = get_business_column_mapping(wb_v2)

    # 3️⃣ Validate differences
    diff = validate_mappings(map_v1, map_v2)
    print("\n=== Mapping Differences ===")
    print(diff)

    # 4️⃣ Extract formulas from one workbook (optional)
    formulas = extract_formulas(wb_v1)

    # 5️⃣ Build the full JSON‑ready metadata (includes resolved refs)
    meta = build_mapping_metadata(
        file_name=os.path.basename(r"E:\Excel_Formula_Analyser\PFLT Sub - Borrowing Base - 07.31.25 ME Truist.xlsx"),
        version="v1",
        mapping=map_v1,
        formulas=formulas,
    )
    print("\n=== Sample Metadata (truncated) ===")
    import json
    print(json.dumps(meta, indent=2)[:500] + "…")
