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

Every ADDED / REMOVED cell and every MODIFIED entry also gets a "reason"
attributing the change to a likely structural cause, so a user can tell
"this is a new column" apart from "someone rewrote this formula's logic":
    row_added         – this row's formula count jumped from near-empty to
                         near-full (a new record, e.g. a new loan)
    row_removed       – the mirror: a row collapsed from full to near-empty
    column_added      – this column's header name is new (didn't exist before)
    column_removed    – this column's header name no longer exists
    column_shifted    – header name exists in both, but at a different column letter
    None              – no structural explanation; a targeted, standalone change

Public functions:
    compare_sheet()         – diff one sheet
    compare_workbooks()     – diff all sheets across two versions
    generate_summary()      – aggregate counts and per-sheet breakdown
"""

import re
from collections import Counter
from typing import Any, Dict, List, Optional

_ROW_RE = re.compile(r"(\d+)$")


def _row_of(cell: str) -> int:
    """'AB123' -> 123"""
    m = _ROW_RE.search(cell)
    return int(m.group(1)) if m else -1


def _col_of(cell: str) -> str:
    """'AB123' -> 'AB'"""
    return "".join(ch for ch in cell if ch.isalpha())


def _classify_columns(
    old_map: Optional[Dict[str, str]],
    new_map: Optional[Dict[str, str]],
) -> Dict[str, Any]:
    """Classify columns as added / removed / shifted by comparing header
    NAMES (not raw letters) between two column mappings — a column that
    kept its name but shifted letter (because another column was inserted
    or removed earlier in the row) is 'shifted', not 'added' + 'removed'.
    """
    old_map = old_map or {}
    new_map = new_map or {}
    old_names = {n for n in old_map.values() if n}
    new_names = {n for n in new_map.values() if n}
    old_letter_by_name = {n: c for c, n in old_map.items() if n}

    added   = {c for c, n in new_map.items() if n and n not in old_names}
    removed = {c for c, n in old_map.items() if n and n not in new_names}
    shifted = {
        c: {"from": old_letter_by_name[n], "name": n}
        for c, n in new_map.items()
        if n and old_letter_by_name.get(n) and old_letter_by_name[n] != c
    }
    return {"added": added, "removed": removed, "shifted": shifted}


# ─────────────────────────────────────────────────────────────
# Core comparison – one sheet at a time
# ─────────────────────────────────────────────────────────────

def compare_sheet(
    old_formulas:   Dict[str, str],
    new_formulas:   Dict[str, str],
    old_normalized: Optional[Dict[str, str]] = None,
    new_normalized: Optional[Dict[str, str]] = None,
    old_col_map:    Optional[Dict[str, str]] = None,
    new_col_map:    Optional[Dict[str, str]] = None,
) -> Dict[str, Any]:
    """Compare all formula cells for one sheet between two versions.

    Parameters
    ----------
    old_formulas   : {cell: formula}           from the OLD workbook
    new_formulas   : {cell: formula}           from the NEW workbook
    old_normalized : {cell: normalized_formula} from normalize_all_formulas()
    new_normalized : {cell: normalized_formula}
    old_col_map    : {col_letter: business_name} for this sheet, OLD workbook
    new_col_map    : {col_letter: business_name} for this sheet, NEW workbook
                     (used only to attribute a "reason" to each change)

    Returns
    -------
    dict with:
        added          – {cell: new_formula}
        added_reasons  – {cell: {type, label}}
        removed        – {cell: old_formula}
        removed_reasons– {cell: {type, label}}
        modified       – {cell: {old_formula, new_formula, old_normalized,
                                  new_normalized, change_type, reason}}
        unchanged      – list of cell addresses identical in both versions
    """
    old_norm = old_normalized or {}
    new_norm = new_normalized or {}

    old_keys = set(old_formulas)
    new_keys = set(new_formulas)

    added   = {c: new_formulas[c] for c in new_keys - old_keys}
    removed = {c: old_formulas[c] for c in old_keys - new_keys}

    # ── Structural context for reason attribution ──────────────────────────
    col_info = _classify_columns(old_col_map, new_col_map)
    # Reverse lookup for removed cells: old letter -> where its content moved to
    shifted_from = {v["from"]: {"to": k, "name": v["name"]} for k, v in col_info["shifted"].items()}

    # Row "density" (how many columns have a formula in that row) rather than
    # mere existence — templated sheets pre-fill a handful of skeleton
    # formulas (row numbers, borders, …) down every row regardless of
    # whether real data occupies it, so a genuinely-new row rarely has ZERO
    # formulas in the old file; it has a sparse skeleton that suddenly fills
    # out. A row is "added"/"removed" when its formula count moves from a
    # small fraction of a "fully loaded" row to a large fraction of one, or
    # vice versa. "Fully loaded" is calibrated per sheet (the 75th percentile
    # of non-empty row counts) since a Prices row (1 formula) and a Loan List
    # row (150+ formulas) mean completely different things by "loaded".
    old_row_density = Counter(_row_of(c) for c in old_formulas)
    new_row_density = Counter(_row_of(c) for c in new_formulas)

    def _typical_row_size() -> float:
        values = sorted(v for v in [*old_row_density.values(), *new_row_density.values()] if v > 0)
        if not values:
            return 0
        return values[max(0, int(len(values) * 0.75) - 1)]

    _typical = _typical_row_size()

    def _row_added(row: int) -> bool:
        if _typical <= 0:
            return False
        o, n = old_row_density.get(row, 0), new_row_density.get(row, 0)
        return n > o and o <= 0.2 * _typical

    def _row_removed(row: int) -> bool:
        if _typical <= 0:
            return False
        o, n = old_row_density.get(row, 0), new_row_density.get(row, 0)
        return o > n and n <= 0.2 * _typical

    def _reason_for_added(cell: str) -> Dict[str, str]:
        if _row_added(_row_of(cell)):
            return {"type": "row_added", "label": "New row added"}
        col = _col_of(cell)
        if col in col_info["shifted"]:
            shift = col_info["shifted"][col]
            return {"type": "column_shifted", "label": f"Column shifted ({shift['from']}→{col})"}
        if col in col_info["added"]:
            return {"type": "column_added", "label": "New column added"}
        return {"type": "content_added", "label": "Formula added"}

    def _reason_for_removed(cell: str) -> Dict[str, str]:
        if _row_removed(_row_of(cell)):
            return {"type": "row_removed", "label": "Row removed"}
        col = _col_of(cell)
        if col in shifted_from:
            shift = shifted_from[col]
            return {"type": "column_shifted", "label": f"Column shifted ({col}→{shift['to']})"}
        if col in col_info["removed"]:
            return {"type": "column_removed", "label": "Column removed"}
        return {"type": "content_removed", "label": "Formula removed"}

    def _aligned_old_cell(cell: str) -> str:
        """The OLD-file address to compare THIS new-file cell against.

        Normally that's the same address. But if this column shifted letters,
        comparing the same letter in both files compares two DIFFERENT
        business columns (whatever used to be at this letter vs whatever
        moved into it) — that falsely looks like a logic change even when the
        actual business column's formula never changed. Align on the
        column's PREVIOUS letter instead.
        """
        shift = col_info["shifted"].get(_col_of(cell))
        return f"{shift['from']}{_row_of(cell)}" if shift else cell

    def _reason_for_modified(cell: str, change_type: str) -> Optional[Dict[str, str]]:
        shift = col_info["shifted"].get(_col_of(cell))
        if not shift:
            return None
        arrow = f"{shift['from']}→{_col_of(cell)}"
        if change_type == "semantic_change":
            return {
                "type":  "column_shifted_and_logic_changed",
                "label": f"Column shifted ({arrow}) — formula logic also changed",
            }
        return {"type": "column_shifted", "label": f"Column shifted ({arrow}), logic unchanged"}

    added_reasons   = {c: _reason_for_added(c)   for c in added}
    removed_reasons = {c: _reason_for_removed(c) for c in removed}

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

        aligned_cell = _aligned_old_cell(cell)
        aligned_old_n = old_norm.get(aligned_cell, "") if aligned_cell != cell else old_n
        change_type = "reference_shift" if (aligned_old_n and new_n and aligned_old_n == new_n) else "semantic_change"

        entry = {
            "old_formula":    old_f,
            "new_formula":    new_f,
            "old_normalized": old_n,
            "new_normalized": new_n,
            "change_type":    change_type,
            "reason":         _reason_for_modified(cell, change_type),
        }
        if aligned_cell != cell:
            # Extra context for the UI: the formula that occupied this
            # business column's PREVIOUS position, so a "logic unchanged"
            # claim on a shifted column can be visually verified.
            entry["aligned_old_cell"]       = aligned_cell
            entry["aligned_old_formula"]    = old_formulas.get(aligned_cell)
            entry["aligned_old_normalized"] = aligned_old_n

        modified[cell] = entry

    return {
        "added":           added,
        "added_reasons":   added_reasons,
        "removed":         removed,
        "removed_reasons": removed_reasons,
        "modified":        modified,
        "unchanged":       sorted(unchanged),
    }


# ─────────────────────────────────────────────────────────────
# Compare across all sheets
# ─────────────────────────────────────────────────────────────

def compare_workbooks(
    formulas_old: Dict[str, Dict[str, str]],
    formulas_new: Dict[str, Dict[str, str]],
    norm_old:  Optional[Dict[str, Dict[str, str]]] = None,
    norm_new:  Optional[Dict[str, Dict[str, str]]] = None,
    mapping_old: Optional[Dict[str, Dict[str, str]]] = None,
    mapping_new: Optional[Dict[str, Dict[str, str]]] = None,
) -> Dict[str, Any]:
    """Compare formulas across two full workbooks.

    Parameters
    ----------
    formulas_old : output of extract_formulas() for the OLD workbook
    formulas_new : output of extract_formulas() for the NEW workbook
    norm_old     : output of normalize_all_formulas() for OLD (optional)
    norm_new     : output of normalize_all_formulas() for NEW (optional)
    mapping_old  : output of get_business_column_mapping() for OLD (optional)
    mapping_new  : output of get_business_column_mapping() for NEW (optional)
                   (used only to attribute a "reason" to each change)

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
            old_col_map    = (mapping_old or {}).get(sheet),
            new_col_map    = (mapping_new or {}).get(sheet),
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
