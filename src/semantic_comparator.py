# src/semantic_comparator.py
"""
Semantic / Business-Meaning Formula Comparison

Compares formulas by their BUSINESS MEANING rather than raw text.

The difference from formula_comparator.py:
  - formula_comparator  → raw formula text diff  (=SUM(B10:C10) vs =SUM(B10:D10))
  - semantic_comparator → business-name diff     (=SUM([Revenue]:[Cost]) vs =SUM([Revenue]:[Cost]:[Margin]))

This also catches a case formula_comparator misses:
  - Raw formula is IDENTICAL  (=SUM(B10:C10) in both versions)
  - Column C was renamed from "Cost" to "Expenses" between versions
  - → Business meaning CHANGED even though the raw formula did not

How it works:
  1. Build an AST for each formula (ast_builder.build_ast).
  2. Walk the AST and replace CellRef / Range nodes with business names
     from the column mapping  (e.g. "B10" → "[Revenue]").
  3. Compare the two business-name ASTs to find what changed.
  4. Return a structured result with change_type and a plain-English description.

Public functions:
    compare_meaning(...)          – compare one formula pair
    batch_compare(...)            – compare all shared cells across two workbooks
    generate_semantic_summary()   – aggregate counts and change-type breakdown
"""

import copy
from typing import Any, Dict, List, Optional

from src.ast_builder import build_ast


# ─────────────────────────────────────────────────────────────
# Step 1 – Normalize an AST with business names
# ─────────────────────────────────────────────────────────────

def _col_from_ref(ref: str) -> str:
    """'B10' → 'B',  '$AB$5' → 'AB'  (extract column letters only)"""
    return "".join(ch for ch in ref.replace("$", "").upper() if ch.isalpha())


def _replace_ref(ref: str, col_map: Dict[str, str]) -> str:
    """Return '[BusinessName]' if the column is mapped, else keep original."""
    col = _col_from_ref(ref)
    return f"[{col_map[col]}]" if col in col_map else ref


def _normalize_range(ref: str, col_map: Dict[str, str]) -> str:
    """'B10:C10' → '[Revenue]:[Cost]'"""
    if ":" in ref:
        start, end = ref.split(":", 1)
        return f"{_replace_ref(start, col_map)}:{_replace_ref(end, col_map)}"
    return _replace_ref(ref, col_map)


def normalize_ast(
    node: Dict[str, Any],
    col_map: Dict[str, str],
    all_mappings: Optional[Dict[str, Dict[str, str]]] = None,
) -> Dict[str, Any]:
    """Return a deep copy of the AST with CellRef / Range refs replaced by business names.

    Parameters
    ----------
    node        : AST dict from ast_builder.build_ast()
    col_map     : {col_letter: business_name} for the sheet that owns this formula
    all_mappings: full workbook mapping for resolving cross-sheet references

    Example
    -------
    CellRef {ref: "B10"}    → CellRef {ref: "[Revenue]"}
    Range   {ref: "B10:C10"}→ Range   {ref: "[Revenue]:[Cost]"}
    """
    node = copy.deepcopy(node)
    t = node.get("type", "")

    if t == "CellRef":
        m = (all_mappings or {}).get(node.get("sheet", ""), col_map)
        node["ref"] = _replace_ref(node["ref"], m)

    elif t == "Range":
        m = (all_mappings or {}).get(node.get("sheet", ""), col_map)
        node["ref"] = _normalize_range(node["ref"], m)

    elif t == "BinaryOp":
        node["left"]  = normalize_ast(node["left"],  col_map, all_mappings)
        node["right"] = normalize_ast(node["right"], col_map, all_mappings)

    elif t == "UnaryOp":
        node["operand"] = normalize_ast(node["operand"], col_map, all_mappings)

    elif t == "Function":
        node["args"] = [normalize_ast(a, col_map, all_mappings) for a in node["args"]]

    elif t == "Group":
        node["expr"] = normalize_ast(node["expr"], col_map, all_mappings)

    return node


# ─────────────────────────────────────────────────────────────
# Step 2 – Reconstruct a readable formula string from a normalized AST
# ─────────────────────────────────────────────────────────────

def _ast_to_str(node: Dict[str, Any]) -> str:
    """Convert a normalized AST back to a compact formula string for display."""
    t = node.get("type", "")

    if t == "BinaryOp":
        return f"{_ast_to_str(node['left'])}{node['op']}{_ast_to_str(node['right'])}"

    if t == "UnaryOp":
        op, inner = node["op"], _ast_to_str(node["operand"])
        return f"{inner}{op}" if op == "%" else f"{op}{inner}"

    if t == "Function":
        args = ",".join(_ast_to_str(a) for a in node["args"])
        return f"{node['name']}({args})"

    if t == "Group":
        return f"({_ast_to_str(node['expr'])})"

    if t in ("CellRef", "Range"):
        prefix = f"{node['sheet']}!" if node.get("sheet") else ""
        return f"{prefix}{node['ref']}"

    if t == "Number":
        return str(node["value"])

    if t == "String":
        return f'"{node["value"]}"'

    if t == "Boolean":
        return "TRUE" if node["value"] else "FALSE"

    return node.get("value", "?")


# ─────────────────────────────────────────────────────────────
# Step 3 – Walk two normalized ASTs and describe the first difference
# ─────────────────────────────────────────────────────────────

def _diff_asts(
    a: Dict[str, Any],
    b: Dict[str, Any],
    path: str = "root",
) -> Optional[Dict[str, str]]:
    """Recursively compare two normalized AST nodes.

    Returns the FIRST difference as {change_type, description, location},
    or None when the nodes are identical.
    """
    if a == b:
        return None

    ta, tb = a.get("type"), b.get("type")

    # Root types differ → structural change
    if ta != tb:
        return {
            "change_type": "structural_change",
            "description": f"Formula structure changed from {ta} to {tb}",
            "location":    path,
        }

    if ta == "Function":
        if a["name"] != b["name"]:
            return {
                "change_type": "function_changed",
                "description": f"Function changed: {a['name']} → {b['name']}",
                "location":    path,
            }
        if len(a["args"]) != len(b["args"]):
            delta     = len(b["args"]) - len(a["args"])
            direction = "added" if delta > 0 else "removed"
            return {
                "change_type": "argument_count_changed",
                "description": (
                    f"{abs(delta)} argument(s) {direction} in {a['name']} "
                    f"({len(a['args'])} → {len(b['args'])})"
                ),
                "location":    path,
            }
        for i, (arg_a, arg_b) in enumerate(zip(a["args"], b["args"])):
            diff = _diff_asts(arg_a, arg_b, path=f"{path}.arg{i + 1}")
            if diff:
                return diff

    elif ta == "BinaryOp":
        if a["op"] != b["op"]:
            return {
                "change_type": "operator_changed",
                "description": f"Operator changed: '{a['op']}' → '{b['op']}'",
                "location":    path,
            }
        diff = _diff_asts(a["left"],  b["left"],  path=f"{path}.left")
        if diff:
            return diff
        diff = _diff_asts(a["right"], b["right"], path=f"{path}.right")
        if diff:
            return diff

    elif ta == "UnaryOp":
        if a["op"] != b["op"]:
            return {
                "change_type": "operator_changed",
                "description": f"Unary operator changed: '{a['op']}' → '{b['op']}'",
                "location":    path,
            }
        return _diff_asts(a["operand"], b["operand"], path=f"{path}.operand")

    elif ta in ("CellRef", "Range"):
        if a.get("sheet") != b.get("sheet"):
            return {
                "change_type": "sheet_changed",
                "description": f"Sheet reference changed: '{a.get('sheet')}' → '{b.get('sheet')}'",
                "location":    path,
            }
        if a["ref"] != b["ref"]:
            return {
                "change_type": "column_changed",
                "description": f"Business column changed: {a['ref']} → {b['ref']}",
                "location":    path,
            }

    elif ta == "Group":
        return _diff_asts(a["expr"], b["expr"], path=f"{path}.group")

    elif ta in ("Number", "String", "Boolean"):
        if a.get("value") != b.get("value"):
            return {
                "change_type": "constant_changed",
                "description": f"{ta} value changed: {a.get('value')} → {b.get('value')}",
                "location":    path,
            }

    return None


# ─────────────────────────────────────────────────────────────
# Step 4 – Single-formula semantic comparison (public)
# ─────────────────────────────────────────────────────────────

def compare_meaning(
    formula_old: str,
    formula_new: str,
    col_map_old: Dict[str, str],
    col_map_new: Dict[str, str],
    all_mappings_old: Optional[Dict[str, Dict[str, str]]] = None,
    all_mappings_new: Optional[Dict[str, Dict[str, str]]] = None,
) -> Dict[str, Any]:
    """Compare the business meaning of two formula strings.

    Normalizes each formula with its OWN column mapping before comparing,
    so a renamed column is detected even when raw formula text is unchanged.

    Parameters
    ----------
    formula_old / formula_new     raw formula strings
    col_map_old / col_map_new     {col_letter: business_name} for the owning sheet
    all_mappings_old / _new       full workbook mapping (for cross-sheet refs)

    Returns
    -------
    dict:
        result          "identical" | "same_meaning" | "different_meaning"
        change_type     fine-grained label (only when result == "different_meaning")
        description     plain-English explanation of what changed
        old_normalized  normalized formula string (old)
        new_normalized  normalized formula string (new)
    """
    ast_old = normalize_ast(build_ast(formula_old), col_map_old, all_mappings_old)
    ast_new = normalize_ast(build_ast(formula_new), col_map_new, all_mappings_new)

    old_norm = "=" + _ast_to_str(ast_old)
    new_norm = "=" + _ast_to_str(ast_new)

    # Case 1: raw text AND normalized form are both the same → truly identical
    if formula_old == formula_new and old_norm == new_norm:
        return {
            "result":         "identical",
            "change_type":    None,
            "description":    "Formula and business meaning are identical",
            "old_normalized": old_norm,
            "new_normalized": new_norm,
        }

    # Case 2: normalized forms match → row/address shifted, business logic unchanged
    if old_norm == new_norm:
        return {
            "result":         "same_meaning",
            "change_type":    "reference_shift",
            "description":    "Raw formula changed but business meaning is the same (row/address shift only)",
            "old_normalized": old_norm,
            "new_normalized": new_norm,
        }

    # Case 3: normalized forms differ → walk ASTs for a precise description
    diff = _diff_asts(ast_old, ast_new) or {
        "change_type": "unknown_change",
        "description": "Business meaning changed",
        "location":    "root",
    }

    return {
        "result":         "different_meaning",
        "change_type":    diff["change_type"],
        "description":    diff["description"],
        "location":       diff.get("location"),
        "old_normalized": old_norm,
        "new_normalized": new_norm,
    }


# ─────────────────────────────────────────────────────────────
# Step 5 – Batch comparison across two workbooks (public)
# ─────────────────────────────────────────────────────────────

def batch_compare(
    formulas_old: Dict[str, Dict[str, str]],
    formulas_new: Dict[str, Dict[str, str]],
    mapping_old:  Dict[str, Dict[str, str]],
    mapping_new:  Dict[str, Dict[str, str]],
) -> Dict[str, Any]:
    """Run semantic comparison for every shared formula cell across two workbooks.

    Parameters
    ----------
    formulas_old / new  {sheet: {cell: raw_formula}} from extract_formulas()
    mapping_old / new   {sheet: {col_letter: name}} from get_business_column_mapping()
    norm_old / new      pre-built normalized formula dicts (avoids re-normalizing)

    Returns
    -------
    {sheet_name: {cell_address: compare_meaning() result}}
    Cells with result == "identical" are omitted to keep output clean.
    """
    results: Dict[str, Any] = {}

    for sheet in sorted(set(formulas_old) & set(formulas_new)):
        col_map_old = mapping_old.get(sheet, {})
        col_map_new = mapping_new.get(sheet, {})

        shared = set(formulas_old[sheet]) & set(formulas_new[sheet])
        sheet_results: Dict[str, Any] = {}
        for cell in sorted(shared):
            result = compare_meaning(
                formula_old      = formulas_old[sheet][cell],
                formula_new      = formulas_new[sheet][cell],
                col_map_old      = col_map_old,
                col_map_new      = col_map_new,
                all_mappings_old = mapping_old,
                all_mappings_new = mapping_new,
            )
            if result["result"] != "identical":
                sheet_results[cell] = result

        if sheet_results:
            results[sheet] = sheet_results

    return results


# ─────────────────────────────────────────────────────────────
# Step 6 – Summary (public)
# ─────────────────────────────────────────────────────────────

def generate_semantic_summary(results: Dict[str, Any]) -> Dict[str, Any]:
    """Count outcomes and rank change types by frequency.

    Parameters
    ----------
    results : output of batch_compare()

    Returns
    -------
    dict:
        same_meaning_count        formulas that shifted address but kept logic
        different_meaning_count   formulas with real business logic changes
        change_type_breakdown     {change_type: count} sorted by frequency
        sheets_with_changes       sheet names that have at least one meaning change
        cells_with_changes        "Sheet!Cell" list for every meaning change
    """
    same_meaning = different_meaning = 0
    change_type_counts: Dict[str, int] = {}
    sheets_with_changes: List[str] = []
    cells_with_changes:  List[str] = []

    for sheet, sheet_results in results.items():
        has_change = False
        for cell, info in sheet_results.items():
            if info["result"] == "same_meaning":
                same_meaning += 1
            else:
                different_meaning += 1
                ct = info.get("change_type", "unknown")
                change_type_counts[ct] = change_type_counts.get(ct, 0) + 1
                cells_with_changes.append(f"{sheet}!{cell}")
                has_change = True
        if has_change:
            sheets_with_changes.append(sheet)

    sorted_breakdown = dict(
        sorted(change_type_counts.items(), key=lambda x: x[1], reverse=True)
    )

    return {
        "same_meaning_count":      same_meaning,
        "different_meaning_count": different_meaning,
        "change_type_breakdown":   sorted_breakdown,
        "sheets_with_changes":     sorted(sheets_with_changes),
        "cells_with_changes":      sorted(cells_with_changes),
    }
