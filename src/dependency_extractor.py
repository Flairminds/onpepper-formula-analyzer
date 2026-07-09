# src/dependency_extractor.py
"""
Phase 3 – Dependency Extraction

Uses openpyxl's built-in formula tokenizer so we do NOT need to write
any custom regex. Each formula is broken into tokens by openpyxl, and
we just inspect the type/subtype of each token.

Steps:
  1. parse_formula()                       – extract refs/functions from one formula
  2. validate_refs()                       – check cross-sheet refs exist in workbook
  3. build_dependency_graph()              – run steps 1-2 for every formula cell
  4. generate_dependency_summary()         – totals, top cells, top functions
  5. enrich_payload_with_dependencies()    – attach graph + summary to the payload

Phase 4 extension points (no code changes needed, just extend):
  - Formula comparison  : diff two graphs returned by build_dependency_graph()
  - Impact analysis     : build a reverse map (which cells depend on cell X?)
  - Named ranges        : openpyxl's workbook.defined_names can be queried here
"""

from collections import defaultdict
from typing import Any, Dict, List

from openpyxl.formula import Tokenizer   # openpyxl built-in tokenizer


# ─────────────────────────────────────────────────────────────
# Step 1 – Parse a single formula using openpyxl's tokenizer
# ─────────────────────────────────────────────────────────────

def parse_formula(formula: str) -> Dict[str, Any]:
    """Extract all dependencies and functions from a single formula string.

    openpyxl Token types we use:
      token.type == 'OPERAND'  AND  token.subtype == 'RANGE'
          → a cell or range reference (local or cross-sheet)
      token.type == 'FUNC'     AND  token.subtype == 'OPEN'
          → a function call, e.g. value = "SUM("

    Parameters
    ----------
    formula : str
        Raw formula string such as ``"=SUM(Sheet1!B10:C10)+D5"``

    Returns
    -------
    dict with keys:
        local_cells   – cell addresses on the same sheet (e.g. ["A1", "D5"])
        local_ranges  – range addresses on the same sheet (e.g. ["B10:C10"])
        cross_sheet   – {sheet_name: [refs…]}  for cross-sheet references
        functions     – list of unique function names (e.g. ["SUM", "IF"])
        constants     – numeric and boolean literals found in the formula
    """
    tok = Tokenizer(formula)

    local_cells  = set()
    local_ranges = set()
    cross_sheet: Dict[str, list] = defaultdict(list)
    functions    = set()
    constants    = []

    for token in tok.items:

        # ── Cell / Range reference ──────────────────────────────────────────
        if token.type == "OPERAND" and token.subtype == "RANGE":
            # Normalise: strip $ signs, make uppercase
            ref = token.value.replace("$", "").upper()

            if "!" in ref:
                # Cross-sheet reference – split on the last '!'
                sheet_part, cell_part = ref.rsplit("!", 1)
                sheet_name = sheet_part.strip("'")          # remove quotes
                if cell_part not in cross_sheet[sheet_name]:
                    cross_sheet[sheet_name].append(cell_part)

            elif ":" in ref:
                local_ranges.add(ref)

            else:
                local_cells.add(ref)

        # ── Function name ───────────────────────────────────────────────────
        elif token.type == "FUNC" and token.subtype == "OPEN":
            # openpyxl gives "SUM(" — strip the trailing paren
            functions.add(token.value.rstrip("(").upper())

        # ── Numeric / Boolean constants ─────────────────────────────────────
        elif token.type == "OPERAND" and token.subtype in ("NUMBER", "LOGICAL"):
            constants.append(token.value)

    return {
        "local_cells":   sorted(local_cells),
        "local_ranges":  sorted(local_ranges),
        "cross_sheet":   {s: sorted(r) for s, r in sorted(cross_sheet.items())},
        "functions":     sorted(functions),
        "constants":     constants,
    }


# ─────────────────────────────────────────────────────────────
# Step 2 – Validate references
# ─────────────────────────────────────────────────────────────

def validate_refs(deps: Dict[str, Any], workbook: Any) -> Dict[str, Any]:
    """Check that every cross-sheet reference points to an existing sheet.

    Returns
    -------
    {"valid": bool, "issues": [list of problem strings]}
    """
    # Build a set of existing sheet names in uppercase for case‑insensitive comparison
    existing_upper = {name.upper() for name in workbook.sheetnames}
    issues = [
        f"Sheet '{sheet}' does not exist in this workbook"
        for sheet in deps["cross_sheet"]
        if sheet.upper() not in existing_upper
    ]
    return {"valid": not issues, "issues": issues}


# ─────────────────────────────────────────────────────────────
# Step 3 – Build dependency graph
# ─────────────────────────────────────────────────────────────

def build_dependency_graph(
    formulas: Dict[str, Dict[str, str]],
    workbook: Any,
) -> Dict[str, Any]:
    """Build one node per formula cell describing what it depends on.

    Parameters
    ----------
    formulas : dict
        ``{sheet_name: {cell_address: formula_string}}``
        — output of formula_extractor.extract_formulas()
    workbook : openpyxl.Workbook
        Used only for sheet-existence validation.

    Returns
    -------
    Nested dict: ``{sheet_name: {cell_address: node}}``

    Each node looks like::

        {
          "formula"     : "=SUM(Sheet2!B10:C10)+D5",
          "depends_on"  : {
              "local_cells"  : ["D5"],
              "local_ranges" : [],
              "cross_sheet"  : {"Sheet2": ["B10:C10"]}
          },
          "token_summary": {"functions": ["SUM"], "constants": []},
          "validation"  : {"valid": true, "issues": []}
        }
    """
    graph: Dict[str, Any] = {}

    for sheet_name, sheet_formulas in formulas.items():
        graph[sheet_name] = {}

        for cell, formula in sheet_formulas.items():
            deps  = parse_formula(formula)
            valid = validate_refs(deps, workbook)

            graph[sheet_name][cell] = {
                "formula": formula,
                "depends_on": {
                    "local_cells":  deps["local_cells"],
                    "local_ranges": deps["local_ranges"],
                    "cross_sheet":  deps["cross_sheet"],
                },
                "token_summary": {
                    "functions": deps["functions"],
                    "constants": deps["constants"],
                },
                "validation": valid,
            }

    return graph


# ─────────────────────────────────────────────────────────────
# Step 4 – Generate summary statistics
# ─────────────────────────────────────────────────────────────

def generate_dependency_summary(graph: Dict[str, Any]) -> Dict[str, Any]:
    """Count totals, find the most‑referenced cells, and collect validation issues.

    Parameters
    ----------
    graph: dict – output of `build_dependency_graph`.

    Returns
    -------
    dict with aggregate statistics.
    """
    total_cells = total_local = total_cross = 0
    ref_counts: Dict[str, int] = defaultdict(int)
    fn_counts: Dict[str, int] = defaultdict(int)
    ext_sheets = set()
    issues: List[Dict] = []

    for sheet, sheet_graph in graph.items():
        for cell, info in sheet_graph.items():
            total_cells += 1
            deps = info["depends_on"]
            # Local cell and range references
            for ref in deps["local_cells"] + deps["local_ranges"]:
                total_local += 1
                ref_counts[f"{sheet}!{ref}"] += 1
            # Cross‑sheet references
            for ref_sheet, refs in deps["cross_sheet"].items():
                ext_sheets.add(ref_sheet)
                for ref in refs:
                    total_cross += 1
                    ref_counts[f"{ref_sheet}!{ref}"] += 1
            # Function usage
            for fn in info["token_summary"]["functions"]:
                fn_counts[fn] += 1
            # Validation problems
            if not info["validation"]["valid"]:
                for issue in info["validation"]["issues"]:
                    issues.append({"sheet": sheet, "cell": cell, "issue": issue})

    top_refs = sorted(ref_counts.items(), key=lambda x: x[1], reverse=True)[:10]
    top_fns = sorted(fn_counts.items(), key=lambda x: x[1], reverse=True)[:10]

    return {
        "total_formula_cells": total_cells,
        "total_local_dependencies": total_local,
        "total_cross_sheet_dependencies": total_cross,
        "sheets_externally_referenced": sorted(ext_sheets),
        "most_referenced_cells": [{"cell": c, "count": n} for c, n in top_refs],
        "top_functions": [{"function": f, "count": n} for f, n in top_fns],
        "validation_issues_count": len(issues),
        "validation_issues": issues,
    }


# ─────────────────────────────────────────────────────────────
# Step 5 – Integrate into the existing JSON payload
# ─────────────────────────────────────────────────────────────

def enrich_payload_with_dependencies(
    payload: Dict[str, Any],
    workbook: Any,
    formulas: Dict[str, Dict[str, str]],
) -> Dict[str, Any]:
    """Attach dependency_graph and dependency_summary to the existing payload.

    This is the only function tracker.py needs to import.

    Usage::

        payload = build_mapping_metadata(file_name, version, mapping, formulas)
        payload = enrich_payload_with_dependencies(payload, wb, formulas)
        save_workbook_record(file_name, version, payload)

    Parameters
    ----------
    payload  : dict  — output of mapping.build_mapping_metadata()
    workbook : openpyxl.Workbook
    formulas : dict  — output of formula_extractor.extract_formulas()

    Returns
    -------
    The same payload dict with two new keys added:
        "dependency_graph"    – per-cell dependency nodes
        "dependency_summary"  – aggregate statistics
    """
    graph = build_dependency_graph(formulas, workbook)

    payload["dependency_graph"]   = graph
    payload["dependency_summary"] = generate_dependency_summary(graph)

    return payload
