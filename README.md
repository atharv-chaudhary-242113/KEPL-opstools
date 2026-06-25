# KEPL OpsTools

KEPL OpsTools is a unified toolkit designed to automate and simplify how we handle Bill of Materials (BoM) and panel schedules. Transitioning from standalone scripts into this scalable package significantly reduces technical debt and makes the tools much easier for the company to adopt.

* **Purpose:** This is a single, consolidated repository for the operational tools built for the company.


* **Data Consolidation:** It features an automated Python pipeline designed to extract, clean, and consolidate BoM data spread across multiple Excel worksheets.


* **Data Classification:** It includes a Python-based utility designed to parse master BoMs or panel schedules in Excel format.


* **Format Preservation:** The classifier categorizes panel series into individual worksheets while strictly preserving existing structural elements, static columns, formulas, and formatting.


* **Centralized Access:** Combining the BoM Consolidator and Panel Classifier under the kepl_opstools namespace creates a streamlined, highly accessible workspace.


---

## What Does It Do?

This package currently features two main operations that can be run independently or together:

### 1. BoM Merge (Consolidator)

* **Functionality:** This tool scans standard BoM templates, dynamically identifies target panels and item categories, merges duplicate components based on Catalog Numbers or Descriptions, aggregates their quantities, and exports a unified master BoM along with an exception log.


### 2. BoM Split (Panel Classifier)

* **Functionality:** It consolidates and categorizes panel series into individual worksheets while strictly preserving existing structural elements, static columns, formulas, and formatting.

* **Dynamic Series Detection:** Automatically identifies panel columns starting from Column F (Row 1) up to the explicit "EXISTING COST" boundary marker.

* **Structural Integrity Retention:** Preserves all master formulas, cell colors (e.g., green panel headers), and static category rows (Rows 1–4).

* **Automated Row/Column Pruning:** Deletes rows with zero or null values for the active panel series and strips unrelated panel columns without mutating the source schema.

* **Output Generation:** Generates `classified_panels.xlsx`, `exceptions.xlsx`, and `run_log.log`.

---

## Getting Started

Follow these steps to set up the environment on your machine.

### Step 1: Install Python

* Ensure you have Python 3.12 or higher installed.

* You can download it from python.org.

### Step 2: Set Up Your Virtual Environment

A virtual environment keeps the tool's dependencies safely isolated from the rest of your computer. You can use standard pip or the lightning-fast uv package manager.

**Option A: Standard Setup (Using pip)**

1. Create the virtual environment:
* Windows: `python -m venv venv`.
* Mac/Linux: `python3 -m venv venv`.


2. Activate the environment:
* Windows: `venv\Scripts\activate`.
* Mac/Linux: `source venv/bin/activate`.

3. Install the required packages:
* `pip install -r requirements.txt`.


**Option B: Fast Setup (Using uv)**
If you have uv installed, setting up is practically instant.

1. Create the virtual environment:
* `uv venv`.


2. Activate the environment:
* Windows: `venv\Scripts\activate`.

* Mac/Linux: `source venv/bin/activate`.


3. Install the required packages:
* `uv pip install -r requirements.txt`.



---

## How to Use the Tool

Once your environment is set up and activated, you use a single command line interface (`bom.py`) to run either tool. Neither `--split` nor `--merge` is a required command but they can be used together.

To execute the commands, simply open your terminal and run the following, replacing the path placeholders with your actual file paths:

`python bom.py --split "path/to/the/file" --merge "path/to/the/file"`.

---

## Customization

If your input Excel files have a different layout, you don't need to rewrite the code. You can simply adjust the variables in `config.py` to match your spreadsheet's structure. Open the configuration file in a text editor and modify the following sections based on your data:

* **PANEL_ROW = 1:** Change this if the names of your panels (e.g., "Main Board", "Sub Panel A") are not in Row 1.


* **HEADER_ROW = 4:** Change this if the static column headers (like SNo, DESCRIPTION, MAKE) are located on a different row.


* **PANEL_START_COL = 7:** Change this if your panel quantities don't start at Column G (which is the 7th letter of the alphabet).


* **STATIC_COLS = ["SNo", "DESCRIPTION", "SPEC", "MAKE", "UNIT", "CAT NO."]:** Change this if your input columns have slightly different names.
