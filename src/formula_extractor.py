# src/formula_extractor.py
"""Utilities to extract formulas from an Excel workbook.

The functions operate on an openpyxl Workbook object. They return a nested dict:
    {
        "SheetName": {"A1": "=SUM(B1:C1)", "B2": "=IF(...)", ...},
        ...
    }
"""

from typing import Any, Dict


def extract_formulas(workbook: Any) -> Dict[str, Dict[str, str]]:
    """Extract all formulas from every sheet in the workbook.

    Parameters
    ----------
    workbook: openpyxl.Workbook
        The workbook loaded with ``data_only=False`` so that ``cell.value`` retains the formula string.

    Returns
    -------
    dict
        Mapping of sheet name → {cell_coordinate: formula_string} for all cells that contain a formula.
    """
    formulas: Dict[str, Dict[str, str]] = {}
    for sheet_name in workbook.sheetnames:
        sheet = workbook[sheet_name]
        sheet_formulas: Dict[str, str] = {}
        for row in sheet.iter_rows(values_only=False):
            for cell in row:
                if isinstance(cell.value, str) and cell.value.startswith('='):
                    sheet_formulas[cell.coordinate] = cell.value
        formulas[sheet_name] = sheet_formulas
    return formulas
