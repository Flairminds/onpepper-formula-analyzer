# src/formula_comparator.py
"""
Formula Comparison Engine – Compare formulas between two workbook versions.

Identifies every formula cell as one of four states:
    ADDED     – exists in new version, not in old
    REMOVED   – existed in old version, not in new
    MODIFIED  – exists in both but the formula text changed
    UNCHANGED – exists in both with the same formula text

For MODIFIED cells, a change_type label is also set:
    "semantic_change"    – normalized formula also changed  (logic changed)
    "reference_shift"    – raw formula changed but normalized is the same
                           (e.g. row number moved, but still the same column/business name)

Public functions:
    compare_sheet()         – diff one sheet
    compare_workbooks()     – diff all sheets across two versions
    generate_summary()      – aggregate counts and per-sheet breakdown
"""

from typing import Any, Dict, List, Optional


# ─────────────────────────────────────────────────────────────
# Core comparison – one sheet at a time
# ─────────────────────────────────────────────────────────────

def compare_sheet(
    old_formulas:   Dict[str, str],
    new_formulas:   Dict[str, str],
    old_normalized: Optional[Dict[str, str]] = None,
    new_normalized: Optional[Dict[str, str]] = None,
) -> Dict[str, Any]:
    """Compare all formula cells for one sheet between two versions.

    Parameters
    ----------
    old_formulas   : {cell: formula}           from the OLD workbook
    new_formulas   : {cell: formula}           from the NEW workbook
    old_normalized : {cell: normalized_formula} from normalize_all_formulas()
    new_normalized : {cell: normalized_formula}

    Returns
    -------
    dict with four keys:
        added     – {cell: new_formula}
        removed   – {cell: old_formula}
        modified  – {cell: {old_formula, new_formula, old_normalized,
                             new_normalized, change_type}}
        unchanged – list of cell addresses identical in both versions
    """
    old_norm = old_normalized or {}
    new_norm = new_normalized or {}

    old_keys = set(old_formulas)
    new_keys = set(new_formulas)

    added   = {c: new_formulas[c] for c in new_keys - old_keys}
    removed = {c: old_formulas[c] for c in old_keys - new_keys}

    modified: Dict[str, Any] = {}
    unchanged: List[str] = []

    for cell in sorted(old_keys & new_keys):
        old_f = old_formulas[cell]
        new_f = new_formulas[cell]

        if old_f == new_f:
            unchanged.append(cell)
            continue

        old_n = old_norm.get(cell, "")
        new_n = new_norm.get(cell, "")
        change_type = "reference_shift" if (old_n and new_n and old_n == new_n) else "semantic_change"

        modified[cell] = {
            "old_formula":    old_f,
            "new_formula":    new_f,
            "old_normalized": old_n,
            "new_normalized": new_n,
            "change_type":    change_type,
        }

    return {
        "added":     added,
        "removed":   removed,
        "modified":  modified,
        "unchanged": sorted(unchanged),
    }


# ─────────────────────────────────────────────────────────────
# Compare across all sheets
# ─────────────────────────────────────────────────────────────

def compare_workbooks(
    formulas_old: Dict[str, Dict[str, str]],
    formulas_new: Dict[str, Dict[str, str]],
    norm_old:  Optional[Dict[str, Dict[str, str]]] = None,
    norm_new:  Optional[Dict[str, Dict[str, str]]] = None,
) -> Dict[str, Any]:
    """Compare formulas across two full workbooks.

    Parameters
    ----------
    formulas_old : output of extract_formulas() for the OLD workbook
    formulas_new : output of extract_formulas() for the NEW workbook
    norm_old     : output of normalize_all_formulas() for OLD (optional)
    norm_new     : output of normalize_all_formulas() for NEW (optional)

    Returns
    -------
    dict with:
        sheets_added    – sheet names that exist only in the new workbook
        sheets_removed  – sheet names that exist only in the old workbook
        sheet_diffs     – {sheet_name: compare_sheet() result}
                          Only sheets that have at least one change are included.
    """
    sheets_old = set(formulas_old.keys())
    sheets_new = set(formulas_new.keys())

    result: Dict[str, Any] = {
        "sheets_added":   sorted(sheets_new - sheets_old),
        "sheets_removed": sorted(sheets_old - sheets_new),
        "sheet_diffs":    {},
    }

    for sheet in sorted(sheets_old & sheets_new):
        diff = compare_sheet(
            old_formulas   = formulas_old[sheet],
            new_formulas   = formulas_new[sheet],
            old_normalized = (norm_old or {}).get(sheet),
            new_normalized = (norm_new or {}).get(sheet),
        )

        if diff["added"] or diff["removed"] or diff["modified"]:
            result["sheet_diffs"][sheet] = diff

    return result


# ─────────────────────────────────────────────────────────────
# Summary report
# ─────────────────────────────────────────────────────────────

def generate_summary(diff: Dict[str, Any]) -> Dict[str, Any]:
    """Produce a concise summary of a compare_workbooks() result.

    Returns
    -------
    dict with:
        sheets_added        – list of new sheet names
        sheets_removed      – list of removed sheet names
        sheets_with_changes – list of sheet names that had formula changes
        total_added         – total formula cells added across all sheets
        total_removed       – total formula cells removed
        total_modified      – total formula cells modified
        total_unchanged     – total formula cells unchanged
        total_changes       – added + removed + modified combined
        semantic_changes    – modified cells where business logic changed
        reference_shifts    – modified cells where only row/address shifted
        per_sheet           – {sheet: {added, removed, modified, unchanged counts}}
    """
    total_added = total_removed = total_modified = total_unchanged = 0
    semantic_changes = reference_shifts = 0
    per_sheet: Dict[str, Any] = {}

    for sheet, s_diff in diff["sheet_diffs"].items():
        a = len(s_diff["added"])
        r = len(s_diff["removed"])
        m = len(s_diff["modified"])
        u = len(s_diff["unchanged"])

        total_added     += a
        total_removed   += r
        total_modified  += m
        total_unchanged += u

        # Count change types within modified cells
        sem  = sum(1 for v in s_diff["modified"].values() if v["change_type"] == "semantic_change")
        refs = sum(1 for v in s_diff["modified"].values() if v["change_type"] == "reference_shift")
        semantic_changes  += sem
        reference_shifts  += refs

        per_sheet[sheet] = {
            "added":     a,
            "removed":   r,
            "modified":  m,
            "unchanged": u,
            "semantic_changes":  sem,
            "reference_shifts":  refs,
        }

    return {
        "sheets_added":        diff["sheets_added"],
        "sheets_removed":      diff["sheets_removed"],
        "sheets_with_changes": sorted(diff["sheet_diffs"].keys()),
        "total_added":         total_added,
        "total_removed":       total_removed,
        "total_modified":      total_modified,
        "total_unchanged":     total_unchanged,
        "total_changes":       total_added + total_removed + total_modified,
        "semantic_changes":    semantic_changes,
        "reference_shifts":    reference_shifts,
        "per_sheet":           per_sheet,
    }
