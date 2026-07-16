# compare_versions.py
"""
Entry point: Compare two Excel workbook versions and print a diff report.

Usage:
    python compare_versions.py <old.xlsx> <new.xlsx> [options]

Options:
    --save     <file>   Save full JSON report to <file>.
    --export   <file>   Export Excel audit report to <file> (e.g. report.xlsx).
    --max-diff <N>      Max formula changes to print per sheet (default 50; 0 = all).

Examples:
    python compare_versions.py v1.xlsx v2.xlsx
    python compare_versions.py v1.xlsx v2.xlsx --save report.json
    python compare_versions.py v1.xlsx v2.xlsx --export audit.xlsx
    python compare_versions.py v1.xlsx v2.xlsx --save report.json --export audit.xlsx --max-diff 20
"""

import os
import sys
import json
import datetime
from dotenv import load_dotenv

# Windows consoles default to cp1252 which can't encode box-drawing characters.
# Reconfigure stdout to UTF-8 so Unicode print statements work on all platforms.
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="backslashreplace")

load_dotenv(dotenv_path=os.path.join(os.path.dirname(__file__), 'config', '.env'))

from src.compare_service import run_comparison
from src.excel_exporter import export_report
# from src.math_equivalence import find_equivalent_pairs
# from src.comparison_validator import validate_results, generate_validation_report


def _print_summary(summary: dict):
    """Print the comparison summary in a readable way."""
    print("\n" + "=" * 55)
    print("  FORMULA COMPARISON SUMMARY")
    print("=" * 55)

    if summary["sheets_added"]:
        print(f"  New sheets added  : {', '.join(summary['sheets_added'])}")
    if summary["sheets_removed"]:
        print(f"  Sheets removed    : {', '.join(summary['sheets_removed'])}")
    if summary["sheets_with_changes"]:
        print(f"  Sheets changed    : {', '.join(summary['sheets_with_changes'])}")

    print(f"\n  Formulas added    : {summary['total_added']}")
    print(f"  Formulas removed  : {summary['total_removed']}")
    print(f"  Formulas modified : {summary['total_modified']}")
    print(f"    -> Semantic changes (logic changed) : {summary['semantic_changes']}")
    print(f"    -> Reference shifts (address only)  : {summary['reference_shifts']}")
    print(f"  Formulas unchanged: {summary['total_unchanged']}")
    print(f"  Total changes     : {summary['total_changes']}")
    print("=" * 55)


def _print_sheet_diff(sheet_name: str, diff: dict, max_diff: int = 50):
    """Print detailed changes for one sheet.

    max_diff  – max entries to print per section (0 = unlimited).
    """
    print(f"\n── Sheet: {sheet_name} " + "─" * 30)
    def _show(items, label, fmt):
        if not items:
            return
        total = len(items)
        print(f"\n  [{label}] {total} formula(s)")
        shown = list(items)[:max_diff] if max_diff else list(items)
        for entry in shown:
            fmt(entry)
        if max_diff and total > max_diff:
            print(f"    … {total - max_diff} more (use --max-diff 0 to show all)")

    added_reasons   = diff.get("added_reasons", {})
    removed_reasons = diff.get("removed_reasons", {})

    # Added
    _show(
        sorted(diff["added"].items()),
        "ADDED",
        lambda kv: print(f"    {kv[0]:>6}  +  {kv[1]}  [{added_reasons.get(kv[0], {}).get('label', '')}]"),
    )

    # Removed
    _show(
        sorted(diff["removed"].items()),
        "REMOVED",
        lambda kv: print(f"    {kv[0]:>6}  -  {kv[1]}  [{removed_reasons.get(kv[0], {}).get('label', '')}]"),
    )

    # Modified
    def _fmt_modified(kv):
        cell, info = kv
        label = "SEMANTIC" if info["change_type"] == "semantic_change" else "REF-SHIFT"
        reason = info.get("reason")
        tag = f"  [{reason['label']}]" if reason else ""
        print(f"    {cell:>6}  [{label}]{tag}")
        print(f"           OLD: {info['old_formula']}")
        print(f"           NEW: {info['new_formula']}")
        sem_desc = info.get("semantic_description")
        if sem_desc:
            print(f"           WHY: {sem_desc}")
        elif info["old_normalized"] and info["old_normalized"] != info["old_formula"]:
            print(f"           OLD (normalized): {info['old_normalized']}")
            print(f"           NEW (normalized): {info['new_normalized']}")

    _show(sorted(diff["modified"].items()), "MODIFIED", _fmt_modified)


def _parse_args():
    """Parse sys.argv and return (old_path, new_path, save_path, export_path, max_diff)."""
    args = sys.argv[1:]
    save_path   = None
    export_path = None
    max_diff    = 50

    def _pop_flag(flag):
        if flag not in args:
            return None
        idx = args.index(flag)
        val = args[idx + 1] if idx + 1 < len(args) else None
        del args[idx:idx + 2]
        return val

    val = _pop_flag("--save")
    if val is not None:
        save_path = val

    val = _pop_flag("--export")
    if val is not None:
        export_path = val

    val = _pop_flag("--max-diff")
    if val is not None:
        try:
            max_diff = int(val)
        except ValueError:
            print(f"Warning: --max-diff must be an integer, got '{val}'. Using default 50.")

    if len(args) != 2:
        print("Usage: python compare_versions.py <old.xlsx> <new.xlsx>")
        print("       [--save output.json]")
        print("       [--export audit.xlsx]")
        print("       [--max-diff N]   (default 50; 0 = show all)")
        sys.exit(1)

    return args[0], args[1], save_path, export_path, max_diff


def main():
    old_path, new_path, save_path, export_path, max_diff = _parse_args()

    print(f"Loading OLD: {os.path.basename(old_path)}")
    print(f"Loading NEW: {os.path.basename(new_path)}")
    print("Extracting formulas, building mappings, and comparing...")
    report = run_comparison(old_path, new_path)

    diff              = report["diff"]
    summary           = report["summary"]
    semantic_results  = report["semantic_diff"]
    semantic_summary  = report["semantic_summary"]

    # ── Print raw diff ────────────────────────────────────────────────────────
    _print_summary(summary)
    for sheet_name, sheet_diff in diff["sheet_diffs"].items():
        _print_sheet_diff(sheet_name, sheet_diff)

    # ── Print semantic summary ────────────────────────────────────────────────
    print("\n" + "=" * 55)
    print("  BUSINESS-MEANING COMPARISON")
    print("=" * 55)
    print(f"  Same meaning (ref shift only) : {semantic_summary['same_meaning_count']}")
    print(f"  Different meaning             : {semantic_summary['different_meaning_count']}")
    if semantic_summary["change_type_breakdown"]:
        print("\n  Change-type breakdown:")
        for ct, count in semantic_summary["change_type_breakdown"].items():
            print(f"    {ct:<30} {count}")
    if semantic_summary["cells_with_changes"]:
        print("\n  Cells with meaning changes:")
        for entry in semantic_results.items():
            sheet, sheet_res = entry
            for cell, info in sorted(sheet_res.items()):
                if info["result"] == "different_meaning":
                    print(f"    {sheet}!{cell:<8}  [{info['change_type']}]")
                    print(f"      {info['description']}")
                    print(f"      OLD: {info['old_normalized']}")
                    print(f"      NEW: {info['new_normalized']}")
    print("=" * 55)

#    # ── Math equivalence: cells that look different but compute the same value ──
#    # equiv_pairs = find_equivalent_pairs(
#    #     semantic_results, formulas_old, formulas_new, mapping_old, mapping_new
#    # )
#    # if equiv_pairs:
#    #     print("\n" + "=" * 55)
#    #     print("  MATHEMATICALLY EQUIVALENT REWRITES")
#    #     print("=" * 55)
#    #     print(f"  {len(equiv_pairs)} cell(s) are flagged 'different_meaning' but")
#    #     print("  are mathematically equivalent (e.g. A+B rewritten as B+A):")
#    #     for p in equiv_pairs:
#    #         print(f"\n    {p['sheet']}!{p['cell']}")
#    #         print(f"      OLD: {p['old_formula']}")
#    #         print(f"      NEW: {p['new_formula']}")
#    #         print(f"      Why: {p['reason']}")
#    #     print("=" * 55)

#    # ── Validation: consistency checks across all comparison results ──────────
#    # print("\nValidating comparison results...")
#    # validation = validate_results(
#    #     semantic_results, formulas_old, formulas_new, mapping_old, mapping_new
#    # )
#    # print(generate_validation_report(validation))

    # ── Optionally save full report as JSON ───────────────────────────────────
    if save_path:
        with open(save_path, "w", encoding="utf-8") as f:
            json.dump(report, f, indent=2)
        print(f"\nFull report saved to: {save_path}")

    # ── Optionally export Excel audit report ──────────────────────────────────
    if export_path:
        print(f"\nGenerating Excel audit report...")
        out = export_report(report, export_path)
        print(f"Excel audit report saved to: {out}")


if __name__ == "__main__":
    main()
