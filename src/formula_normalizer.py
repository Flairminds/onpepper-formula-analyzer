# src/formula_normalizer.py
"""
Formula Normalization – Replace cell references with business names.

How it works:
  - Uses openpyxl's built-in Tokenizer (no regex needed).
  - For every cell/range token, the column letter (e.g. B) is looked up in
    the business-name mapping (e.g. B → Revenue).
  - The address is replaced with [BusinessName] (e.g. B10 → [Revenue]).
  - If a column has no business name, the original address is kept unchanged.

Public functions:
  normalize_formula(formula, col_map, all_mappings)  → normalized formula string
  normalize_all_formulas(formulas, column_mapping)   → {sheet: {cell: normalized}}
"""

from typing import Dict, Optional
from openpyxl.formula import Tokenizer


# ─────────────────────────────────────────────────────────────
# Private helpers
# ─────────────────────────────────────────────────────────────

def _get_col(cell_ref: str) -> str:
    """Extract the column letters from a cell address.

    Examples:
        'B10'    → 'B'
        '$AB$99' → 'AB'
        'D5'     → 'D'
    """
    # Strip $ then take only the leading letters
    clean = cell_ref.replace('$', '').upper()
    return ''.join(ch for ch in clean if ch.isalpha())


def _replace_cell(cell_ref: str, col_map: Dict[str, str]) -> str:
    """Replace one cell address with its business name if available.

    Examples (col_map = {'B': 'Revenue', 'C': 'Cost'}):
        'B10'  → '[Revenue]'
        'C5'   → '[Cost]'
        'Z99'  → 'Z99'   ← not in col_map, kept as-is
    """
    col = _get_col(cell_ref)
    if col in col_map:
        return f"[{col_map[col]}]"
    return cell_ref                          # keep original when no mapping exists


def _normalize_token(token_value: str,
                     col_map: Dict[str, str],
                     all_mappings: Optional[Dict[str, Dict[str, str]]]) -> str:
    """Normalize one RANGE token value (cell, range, or cross-sheet ref).

    Handles four cases:
        Local cell       B10          →  [Revenue]
        Local range      B10:C10      →  [Revenue]:[Cost]
        Cross-sheet cell Sheet1!B10   →  Sheet1![Revenue]
        Cross-sheet range Sheet1!B10:C10 → Sheet1![Revenue]:[Cost]
    """
    # ── Cross-sheet reference (contains '!') ────────────────────────────────
    if '!' in token_value:
        sheet_part, cell_part = token_value.rsplit('!', 1)

        # Identify the sheet name (strip surrounding quotes if present)
        sheet_name = sheet_part.strip("'").replace("''", "'")

        # Pick the col_map for that sheet (fall back to empty dict if unknown)
        sheet_col_map = (all_mappings or {}).get(sheet_name, {})

        # Normalize the cell/range part, then reassemble
        normalized_cell = _normalize_local_ref(cell_part, sheet_col_map)
        return f"{sheet_part}!{normalized_cell}"

    # ── Local reference ──────────────────────────────────────────────────────
    return _normalize_local_ref(token_value, col_map)


def _normalize_local_ref(ref: str, col_map: Dict[str, str]) -> str:
    """Normalize a local cell or range reference (no sheet prefix).

    Examples:
        'B10'      → '[Revenue]'
        'B10:C10'  → '[Revenue]:[Cost]'
        '$B$10'    → '[Revenue]'
    """
    clean = ref.replace('$', '').upper()

    if ':' in clean:
        # Range: normalize each end separately
        start, end = clean.split(':', 1)
        return f"{_replace_cell(start, col_map)}:{_replace_cell(end, col_map)}"

    # Single cell
    return _replace_cell(clean, col_map)


# ─────────────────────────────────────────────────────────────
# Public API
# ─────────────────────────────────────────────────────────────

def normalize_formula(
    formula: str,
    col_map: Dict[str, str],
    all_mappings: Optional[Dict[str, Dict[str, str]]] = None,
) -> str:
    """Return a human-readable version of a formula with business names.

    Cell references are replaced with [BusinessName] using col_map.
    All other parts (functions, operators, constants) are kept unchanged.

    Parameters
    ----------
    formula : str
        Raw formula string, e.g. ``"=SUM(B10:C10)+D5"``
    col_map : dict
        ``{column_letter: business_name}`` for the sheet that owns this formula.
        Comes from ``get_business_column_mapping()[sheet_name]``.
    all_mappings : dict, optional
        Full workbook mapping ``{sheet_name: {col_letter: business_name}}``.
        Pass ``get_business_column_mapping()`` here to resolve cross-sheet refs.

    Returns
    -------
    str
        Normalized formula, e.g. ``"=SUM([Revenue]:[Cost])+[TotalAssets]"``

    Examples
    --------
    col_map = {'B': 'Revenue', 'C': 'Cost', 'D': 'TotalAssets'}

    normalize_formula("=SUM(B10:C10)+D5", col_map)
    → "=SUM([Revenue]:[Cost])+[TotalAssets]"

    normalize_formula("=IF(B10>0, B10, 0)", col_map)
    → "=IF([Revenue]>0,[Revenue],0)"
    """
    tok = Tokenizer(formula)
    parts = []

    for token in tok.items:
        # Only cell/range tokens get replaced; everything else passes through
        if token.type == "OPERAND" and token.subtype == "RANGE":
            parts.append(_normalize_token(token.value, col_map, all_mappings))
        else:
            parts.append(token.value)

    return '=' + ''.join(parts)


def normalize_all_formulas(
    formulas: Dict[str, Dict[str, str]],
    column_mapping: Dict[str, Dict[str, str]],
) -> Dict[str, Dict[str, str]]:
    """Normalize every formula in the workbook.

    Parameters
    ----------
    formulas : dict
        ``{sheet_name: {cell_address: formula}}``
        — output of ``formula_extractor.extract_formulas()``
    column_mapping : dict
        ``{sheet_name: {col_letter: business_name}}``
        — output of ``mapping.get_business_column_mapping()``

    Returns
    -------
    dict
        ``{sheet_name: {cell_address: normalized_formula}}``

    Example output
    --------------
    {
      "Sheet1": {
        "B15": "=SUM([Revenue]:[Cost])+[TotalAssets]",
        "C20": "=IF([Revenue]>0,[Revenue],0)"
      }
    }
    """
    result = {}

    for sheet_name, sheet_formulas in formulas.items():
        # col_map for this sheet; empty dict if sheet has no header mapping
        col_map = column_mapping.get(sheet_name, {})

        result[sheet_name] = {
            cell: normalize_formula(formula, col_map, column_mapping)
            for cell, formula in sheet_formulas.items()
        }

    return result
