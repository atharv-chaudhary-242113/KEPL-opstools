"""
Post-processing cryptograph-style numeric verification.
Calculates dynamic bounds strictly on active line items to evaluate
apples-to-apples numeric fidelity without footer interference.
"""

import logging
from contextlib import suppress
from decimal import Decimal, InvalidOperation

from bom_dic.bom_split.config_split import (
    END_MARKER,
    HEADER_ROW,
    PANEL_ROW,
    PANEL_START_COL,
)
from openpyxl.workbook.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet


def _is_empty(val: object) -> bool:
    if val is None:
        return True
    return str(val).strip().upper() in ("", "0", "0.0", "NONE", "NULL", "-")


def _is_numeric(val: object) -> bool:
    if val is None:
        return False
    if isinstance(val, (int, float)):
        return True
    try:
        float(str(val))
        return True
    except ValueError:
        return False


def _extract_sum(ws: Worksheet, col_idx: int) -> Decimal:
    """
    Aggregates absolute columnar data utilizing exact arithmetic.
    Strictly limits extraction to valid item rows (ignoring dynamic footers)
    to ensure apples-to-apples fidelity.
    """
    total = Decimal("0.0")

    sno_col = 1
    cat_col = -1

    # Locate identifying columns
    for c in range(1, ws.max_column + 1):
        val = str(ws.cell(row=HEADER_ROW, column=c).value or "").strip().upper()
        if val == "SNO":
            sno_col = c
        elif val == "CAT NO.":
            cat_col = c

    for r in range(HEADER_ROW + 1, ws.max_row + 1):
        sno_val = ws.cell(row=r, column=sno_col).value
        cat_val = ws.cell(row=r, column=cat_col).value if cat_col != -1 else None

        # Core fix: Only sum if the row represents actual inventory data
        is_data_row = _is_numeric(sno_val) or not _is_empty(cat_val)

        if is_data_row:
            cell_val = ws.cell(row=r, column=col_idx).value
            if cell_val is not None:
                if isinstance(cell_val, str) and str(cell_val).startswith("="):
                    continue
                with suppress(InvalidOperation):
                    total += Decimal(str(cell_val).strip())

    return total


def validate_outputs(
    wb_master: Workbook, wb_out: Workbook, logger: logging.Logger
) -> bool:
    """
    Enforces a strict equivalence check between the sanitized source numeric footprint
    and the compiled output to guarantee structural and financial fidelity.
    """
    logger.info("Initiating precision numeric validation protocol.")
    passed: bool = True

    master_sums: dict[str, Decimal] = {}

    for sheet_name in wb_master.sheetnames:
        ws = wb_master[sheet_name]

        # Master sheet scans from PANEL_START_COL
        for col in range(PANEL_START_COL, ws.max_column + 1):
            val = str(ws.cell(row=PANEL_ROW, column=col).value or "").strip().upper()
            if END_MARKER.upper() in val:
                break

            panel_name = str(ws.cell(row=PANEL_ROW, column=col).value or "").strip()
            if panel_name:
                master_sums[panel_name] = master_sums.get(
                    panel_name, Decimal("0.0")
                ) + _extract_sum(ws, col)

    output_sums: dict[str, Decimal] = {}
    for sheet_name in wb_out.sheetnames:
        ws = wb_out[sheet_name]

        # Split sheets MUST scan from Col 1 because
        # the Qty column shifted left during extraction
        for col in range(1, ws.max_column + 1):
            val = str(ws.cell(row=PANEL_ROW, column=col).value or "").strip().upper()
            if END_MARKER.upper() in val:
                break

            panel_name = str(ws.cell(row=PANEL_ROW, column=col).value or "").strip()
            if panel_name in master_sums:
                output_sums[panel_name] = output_sums.get(
                    panel_name, Decimal("0.0")
                ) + _extract_sum(ws, col)

    for panel_name, expected in master_sums.items():
        actual = output_sums.get(panel_name, Decimal("0.0"))

        # Rounding precision for safety in float comparisons
        if round(expected, 4) != round(actual, 4):
            logger.error(
                f"Fidelity failure on [{panel_name}]. "
                f"Baseline={expected}, Compiled={actual}."
            )
            passed = False

    if passed:
        logger.info(
            "Validation complete: 100% equivalence achieved against "
            "sanitized baseline matrix."
        )

    return passed
