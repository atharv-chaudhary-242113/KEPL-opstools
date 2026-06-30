# KEPL OpsTools

A unified toolkit for BoM consolidation and panel classification.

KEPL OpsTools is a high-performance, type-safe Python CLI package for engineering operations. It streamlines the processing, splitting, and consolidation of complex Bill of Materials (BoM) Excel workbooks into single, cleanly formatted outputs. The project is built with modern enterprise-grade Python tooling to ensure speed, stability, and ease of use.

## Features

* **Unified Architecture**: Consolidates the BoM Splitter (Panel Classifier) and BoM Merger (Consolidator) into a single, cohesive toolset.
* **Professional CLI**: Run complex data pipelines instantly using standard terminal flags (`--split`, `--merge`), or fall back to an interactive wizard.
* **Enterprise Type Safety**: Powered by Meta's Pyrefly in Strict Mode to guarantee zero implicit `Any` types and robust, crash-free data parsing.
* **Blazing Fast**: Managed entirely via `uv` and built on Hatchling for instant dependency resolution and millisecond build times.
* **Smart Excel Parsing**: Dynamically detects column headers, bypasses `MergedCell` reading errors, and standardizes numeric types directly from OpenPyXL.

---

## Installation

Because this repository uses `pyproject.toml` and modern build tools, you can install the CLI command globally in your virtual environment.

### Option 1: Direct Cloud Install (For End-Users)

Users with access to the GitHub repository can install the tool directly without downloading the source code manually:

```bash
uv pip install git+https://github.com/atharv-chaudhary-242113/KEPL-opstools.git

```

### Option 2: Local Editable Install (For Developers)

If you are actively developing the tool, clone the repository and install it in editable mode so your changes reflect instantly:

```bash
git clone https://github.com/atharv-chaudhary-242113/KEPL-opstools.git
cd KEPL-opstools
uv pip install -e .

```

---

## Usage

Installing the package automatically binds the `bom` command to your terminal.

### 1. Interactive Mode

If you run the command with no arguments, it will launch a user-friendly terminal wizard.

```bash
bom

```

### 2. BoM Splitter (Panel Classifier)

Splits a master BoM file into individual panel worksheets based on column detection.

* Process all workbooks in the default `input/` directory:

```bash
bom --split

```

* Target a specific file or folder:

```bash
bom --split "path/to/your/master_bom.xlsx"

```

* Force splitting of sub-categories (e.g. -1, -2, -A):

```bash
bom --split "path/to/your/master_bom.xlsx" --split-sub

```

### 3. BoM Merger (Consolidator)

Ingests multiple individual Excel files or sheets and consolidates them into a single `final_bom.xlsx` output with a dynamically mapped `CATEGORY` column.

* Process all workbooks in the default `input/` directory:

```bash
bom --merge

```

* Target a specific file or folder:

```bash
bom --merge "C:/path/to/panel/files/"

```

---

## Development & Architecture

### Core Stack

* **Language**: Python 3.12+
* **Package Manager**: `uv`
* **Build Backend**: Hatchling
* **Type Checker**: Pyrefly (Strict Mode)
* **Linter/Formatter**: Ruff
* **Core Libraries**: `openpyxl`, `lxml`, `pandas`, `numpy`, `defusedxml`

### Repository Structure

```text
kepl_opstools/
├── bom_dic/                 # Core source code
│   ├── bom.py               # Central CLI router (argparse)
│   ├── bom_merge/           # Consolidator logic
│   └── bom_split/           # Panel Classifier logic
├── pyproject.toml           # Modernized dependency & tool config
└── README.md

```

### Running Checks

Before pushing new code, ensure the codebase maintains its strict quality standards:

* Run the linter:

```bash
uv run ruff check .

```

* Run the strict-mode type checker:

```bash
uv run pyrefly check

```

---

*Note: The `inventory-forecast` pipeline is currently under active development and will be integrated into this namespace soon.*
