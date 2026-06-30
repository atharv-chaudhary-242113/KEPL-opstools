"""
Execution orchestrator. Initializes the environment, enforces path
security, and triggers the processing pipeline.
"""

import logging
import sys
from pathlib import Path

from .config_split import INPUT_DIR, LOG_FILE, LOGS_DIR, OUTPUT_DIR
from .utils.excel_reader_split import process_workbooks


def _setup_logger() -> logging.Logger:
    LOGS_DIR.mkdir(parents=True, exist_ok=True)
    log_path: Path = LOGS_DIR / LOG_FILE

    fmt: str = "%(asctime)s [%(levelname)s] %(message)s"
    logging.basicConfig(
        level=logging.INFO,
        format=fmt,
        handlers=[
            logging.FileHandler(log_path, encoding="utf-8"),
            logging.StreamHandler(sys.stdout),
        ],
    )
    return logging.getLogger("panel_classifier")


def run_classifier(target_path: str | None = None, split_sub: bool = False) -> None:
    logger: logging.Logger = _setup_logger()

    for directory in (INPUT_DIR, OUTPUT_DIR, LOGS_DIR):
        directory.mkdir(parents=True, exist_ok=True)

    candidates: list[Path] = []

    if target_path and target_path != "DEFAULT":
        target_file = Path(target_path).resolve()
        if not target_file.exists():
            logger.error(f"Provided path does not exist: {target_path}")
            return
        if target_file.is_file():
            candidates = [target_file]
        elif target_file.is_dir():
            candidates = [
                p
                for p in target_file.iterdir()
                if p.is_file()
                and p.suffix.lower() in (".xlsx", ".xls")
                and not p.name.startswith("~$")
            ]
    else:
        candidates = [
            p
            for p in INPUT_DIR.iterdir()
            if p.is_file()
            and p.suffix.lower() in (".xlsx", ".xls")
            and not p.name.startswith("~$")
        ]

    if not candidates:
        logger.error(
            "No valid Excel workbook(s) found in the input directory. Terminating."
        )
        return

    logger.info(f"Targeting {len(candidates)} workbook(s)...")
    success: bool = process_workbooks(candidates, split_sub, logger)

    if not success:
        logger.error("Classification encountered errors.")


if __name__ == "__main__":
    run_classifier()
