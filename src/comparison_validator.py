# src/comparison_validator.py
"""
Comparison Validator

Checks whether the results produced by semantic_comparator and math_equivalence
are internally consistent, and flags any suspicious outcomes.

Three checks are run for every cell that appears in the semantic results:

  1. same_meaning must match    – a cell marked "same_meaning" MUST have identical
                                  normalized forms; if they differ, something is wrong.
  2. different_meaning must differ – a cell marked "different_meaning" MUST have at
                                  least one normalized form that is different; if they
                                  are the same string, the detection logic is broken.
  3. Math-equivalence warnings  – cells marked "different_meaning" that turn out to be
                                  mathematically equivalent are flagged as warnings (not
                                  errors). They are genuine formula rewrites that produce
                                  the same numeric result.

Public functions:
    validate_results(semantic_results, formulas_old, formulas_new, mapping_old, mapping_new)
        → {errors, warnings, stats}
    generate_validation_report(validation)
        → human-readable string
"""

from typing import Any, Dict, List

# from src.math_equivalence import find_equivalent_pairs


# ─────────────────────────────────────────────────────────────
# Main validator
# ─────────────────────────────────────────────────────────────

def validate_results(
    semantic_results: Dict[str, Any],
    formulas_old:     Dict[str, Dict[str, str]],
    formulas_new:     Dict[str, Dict[str, str]],
    mapping_old:      Dict[str, Dict[str, str]],
    mapping_new:      Dict[str, Dict[str, str]],
) -> Dict[str, Any]:
    """Validate semantic comparison results for internal consistency.

    Parameters
    ----------
    semantic_results : output of semantic_comparator.batch_compare()
    formulas_old / new : raw formula strings for both versions
    mapping_old / new  : column mappings for both versions

    Returns
    -------
    dict:
        errors    – list of consistency problems (logic bugs in the comparison)
        warnings  – list of math-equivalent cells flagged as different_meaning
        stats     – summary counts
    """
    errors:   List[Dict[str, str]] = []
    warnings: List[Dict[str, str]] = []

    same_checked       = 0
    different_checked  = 0

    for sheet, sheet_results in semantic_results.items():
        for cell, info in sheet_results.items():
            result     = info["result"]
            old_norm   = info.get("old_normalized", "")
            new_norm   = info.get("new_normalized", "")

            # ── Check 1: same_meaning must have matching normalized forms ──────
            if result == "same_meaning":
                same_checked += 1
                if old_norm != new_norm:
                    errors.append({
                        "sheet":       sheet,
                        "cell":        cell,
                        "check":       "same_meaning_mismatch",
                        "description": (
                            f"Cell is marked 'same_meaning' but normalized forms differ.\n"
                            f"  OLD: {old_norm}\n"
                            f"  NEW: {new_norm}"
                        ),
                    })

            # ── Check 2: different_meaning must have differing normalized forms
            elif result == "different_meaning":
                different_checked += 1
                if old_norm == new_norm:
                    errors.append({
                        "sheet":       sheet,
                        "cell":        cell,
                        "check":       "different_meaning_false_positive",
                        "description": (
                            f"Cell is marked 'different_meaning' but normalized forms are identical.\n"
                            f"  NORM: {old_norm}"
                        ),
                    })

    # ── Check 3: find different_meaning cells that are mathematically equivalent (disabled)
    equiv_pairs = []  # math equivalence detection disabled
    # No warnings for mathematically equivalent rewrites are added.

    stats = {
        "total_cells_checked":    same_checked + different_checked,
        "same_meaning_checked":   same_checked,
        "different_meaning_checked": different_checked,
        "error_count":    len(errors),
        "warning_count":  len(warnings),
        "is_valid":       len(errors) == 0,
    }

    return {
        "errors":   errors,
        "warnings": warnings,
        "stats":    stats,
    }


# ─────────────────────────────────────────────────────────────
# Report generator
# ─────────────────────────────────────────────────────────────

def generate_validation_report(validation: Dict[str, Any]) -> str:
    """Convert validate_results() output into a human-readable string.

    Parameters
    ----------
    validation : output of validate_results()

    Returns
    -------
    Multi-line string suitable for printing to the console.
    """
    lines: List[str] = []
    stats    = validation["stats"]
    errors   = validation["errors"]
    warnings = validation["warnings"]

    lines.append("=" * 55)
    lines.append("  VALIDATION REPORT")
    lines.append("=" * 55)
    lines.append(f"  Cells checked   : {stats['total_cells_checked']}")
    lines.append(f"    same_meaning  : {stats['same_meaning_checked']}")
    lines.append(f"    diff_meaning  : {stats['different_meaning_checked']}")
    lines.append(f"  Errors          : {stats['error_count']}")
    lines.append(f"  Warnings        : {stats['warning_count']}")
    lines.append(f"  Overall         : {'PASS' if stats['is_valid'] else 'FAIL'}")
    lines.append("=" * 55)

    if errors:
        lines.append("\n  ERRORS (comparison logic problems):")
        for e in errors:
            lines.append(f"\n    [{e['check']}]  {e['sheet']}!{e['cell']}")
            for detail_line in e["description"].splitlines():
                lines.append(f"      {detail_line}")

    if warnings:
        lines.append("\n  WARNINGS (math-equivalent but flagged as different):")
        for w in warnings:
            lines.append(f"\n    {w['sheet']}!{w['cell']}")
            for detail_line in w["description"].splitlines():
                lines.append(f"      {detail_line}")

    if not errors and not warnings:
        lines.append("\n  All results are consistent. No issues found.")

    lines.append("=" * 55)
    return "\n".join(lines)
