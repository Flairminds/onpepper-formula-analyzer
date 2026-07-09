"""Reusable comparison logic shared by the compare_versions.py CLI and the Flask backend."""

import os
import datetime

from src.reader import load_excel_workbook
from src.formula_extractor import extract_formulas
from src.mapping import get_business_column_mapping
from src.formula_normalizer import normalize_all_formulas
from src.formula_comparator import compare_workbooks, generate_summary
from src.semantic_comparator import batch_compare, generate_semantic_summary


def run_comparison(old_path: str, new_path: str) -> dict:
    """Load two workbooks, diff their formulas, and return the full report dict."""
    wb_old = load_excel_workbook(old_path)
    wb_new = load_excel_workbook(new_path)

    formulas_old = extract_formulas(wb_old)
    formulas_new = extract_formulas(wb_new)

    mapping_old = get_business_column_mapping(wb_old)
    mapping_new = get_business_column_mapping(wb_new)

    norm_old = normalize_all_formulas(formulas_old, mapping_old)
    norm_new = normalize_all_formulas(formulas_new, mapping_new)

    diff = compare_workbooks(formulas_old, formulas_new, norm_old, norm_new, mapping_old, mapping_new)
    summary = generate_summary(diff)

    semantic_results = batch_compare(formulas_old, formulas_new, mapping_old, mapping_new)
    semantic_summary = generate_semantic_summary(semantic_results)

    return {
        "generated_at": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "old_file": os.path.basename(old_path),
        "new_file": os.path.basename(new_path),
        "summary": summary,
        "column_mapping": {"old": mapping_old, "new": mapping_new},
        "diff": diff,
        "semantic_summary": semantic_summary,
        "semantic_diff": semantic_results,
    }
