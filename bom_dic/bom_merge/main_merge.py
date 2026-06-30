"""
Entry point.

Usage via CLI:
    bom --merge <path>
"""

import logging
import os
import sys

from .config_merge import INPUT_DIR, LOG_FILE, LOGS_DIR, OUTPUT_DIR
from .utils.consolidator_merge import consolidate
from .utils.excel_reader_merge import read_workbook
from .utils.exporter_merge import export_bom, export_exceptions, safety_check


def _setup_logger() -> logging.Logger:
    os.makedirs(LOGS_DIR, exist_ok=True)
    log_path = os.path.join(LOGS_DIR, LOG_FILE)

    fmt = "%(asctime)s [%(levelname)s] %(message)s"
    logging.basicConfig(
        level=logging.INFO,
        format=fmt,
        handlers=[
            logging.FileHandler(log_path, encoding="utf-8"),
            logging.StreamHandler(sys.stdout),
        ],
    )
    return logging.getLogger("bom_consolidator")


def run_consolidator(target_path: str | None = None) -> None:
    logger = _setup_logger()
    logger.info("═══ BoM Consolidation run started ═══════════════════════════")

    for d in (INPUT_DIR, OUTPUT_DIR, LOGS_DIR):
        os.makedirs(d, exist_ok=True)

    candidates: list[str] = []
    if target_path and target_path != "DEFAULT":
        if os.path.isfile(target_path):
            candidates = [target_path]
        elif os.path.isdir(target_path):
            candidates = [
                os.path.join(target_path, f)
                for f in os.listdir(target_path)
                if f.lower().endswith((".xlsx", ".xls")) and not f.startswith("~$")
            ]
        else:
            logger.error(f"Provided path does not exist: {target_path}")
            return
    else:
        candidates = [
            os.path.join(INPUT_DIR, f)
            for f in os.listdir(INPUT_DIR)
            if f.lower().endswith((".xlsx", ".xls")) and not f.startswith("~$")
        ]

    if not candidates:
        logger.error("No Excel workbook(s) found. Add your .xlsx file(s) and re-run.")
        return

    all_sheets_data = []
    all_exceptions = []

    # Loop through ALL candidates and treat them as one giant file
    for input_file in candidates:
        logger.info(f"Ingesting file: {input_file}")
        sheets_data, exceptions = read_workbook(input_file, logger)
        all_sheets_data.extend(sheets_data)
        all_exceptions.extend(exceptions)

    total_sheets = len(all_sheets_data)
    total_rows = sum(len(s["data_rows"]) for s in all_sheets_data)
    logger.info(f"Total Sheets with valid BoM content : {total_sheets}")
    logger.info(f"Total valid rows extracted          : {total_rows}")

    if total_sheets == 0:
        logger.error("No usable sheets found in any files. Check structures.")
        export_exceptions(all_exceptions, [], logger)
        return

    all_panels, sorted_rows = consolidate(all_sheets_data)
    logger.info(f"Unique panels collected             : {len(all_panels)}")
    logger.info(f"Rows after sort                     : {len(sorted_rows)}")

    export_bom(all_panels, sorted_rows, logger)
    export_exceptions(all_exceptions, all_panels, logger)

    ok = safety_check(all_panels, sorted_rows, all_sheets_data, logger)

    logger.info(f"Exceptions flagged                  : {len(all_exceptions)}")
    logger.info("═══ Run complete ════════════════════════════════════════════")

    if not ok:
        logger.warning(
            "Safety verification failed! Please check output logs for discrepancies."
        )


if __name__ == "__main__":
    run_consolidator()
