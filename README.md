# Excel Budget Generator

A modular Python project that generates an enhanced personal budgeting workbook in Excel using `openpyxl`.

## What the workbook supports

- Budget vs Actual tracking for each month
- automatic Variance calculations using `Actual - Budget`
- conditional formatting for overspend / under-budget visibility
- a structured grocery detail table that rolls into the main Groceries row
- an expandable custom variable-expense table
- formula-driven summary totals, dashboards, and trend analysis
- bonus tracking, emergency fund tracking, and a five-year projection

## Folder structure

```text
excel-budget-generator/
├── budget_workbook/
│   ├── __init__.py
│   ├── config.py
│   ├── generator.py
│   ├── rows.py
│   ├── styles.py
│   ├── versioning.py
│   └── builders/
│       ├── base.py
│       ├── bonus_tracker.py
│       ├── cover.py
│       ├── emergency_fund.py
│       ├── five_year_projection.py
│       ├── monthly_entry.py
│       ├── summary_dashboard.py
│       └── trend_analysis.py
├── workbooks/
│   └── .gitkeep
├── budget_workbook_generator.py
├── instructions.md
├── main.py
├── requirements.txt
└── README.md
```

## Workbook output location

All generated Excel files are stored in the root-level `workbooks/` folder.

Default output naming now follows semantic-style versioning:

```text
workbooks/Personal_Budget_Workbook_v0.0.0.xlsx
```

## Versioning strategy

Workbook files use a semantic-style version in the filename:

- `v0.0.0` → initial versioned workbook
- `v0.1.0` → minor workbook enhancement
- `v1.0.0` → stable major version

The current default is controlled by `WorkbookConfig.workbook_version` in `budget_workbook/config.py`.

## Naming convention

Generated workbook filenames follow this pattern:

```text
<workbook_name_prefix>_<version>.xlsx
```

Default example:

```text
Personal_Budget_Workbook_v0.0.0.xlsx
```

## Rules for new workbook files

- New files are written to `workbooks/`
- Existing versioned files are **not overwritten**
- If the target filename already exists, generation raises an error so you can bump the version first
- Older files stay in place for rollback and historical tracking

## What should trigger a version bump

### Patch bump (`v0.0.1`)

Use when changing documentation or implementation details that do not alter the workbook structure.

### Minor bump (`v0.1.0`)

Use when adding or enhancing workbook features without redefining the whole model, for example:

- adding a new dashboard metric
- adding a new dynamic table section
- extending formulas or charts

### Major bump (`v1.0.0`)

Use when making a stable release or introducing a significant workbook structure change.

## Key workbook enhancements already supported

### Budget vs Actual columns

The Monthly Entry sheet uses Budget, Actual, and Variance triplets for every month and year total block.

### Expandable detail tables

Two Excel tables keep the model extendable:

- `GroceryDetailTable`
- `VariableExpenseTable`

New rows added inside those tables automatically feed the summary Groceries and Additional Variable Items rows.

### Grocery starter structure

The grocery section includes starter rows for:

- Grains / Staples
- Protein / Meat
- Vegetables
- Spices / Condiments
- Cereals / Breakfast
- Fruit
- Dairy / Fridge Items
- Household / Cooking Items
- Snacks

Extra blank rows are included for future custom grocery items.

### Dynamic rollups

The workbook keeps calculations formula-driven:

- Groceries summary rows pull from `GroceryDetailTable`
- Additional variable totals pull from `VariableExpenseTable`
- dashboard metrics reference Monthly Entry summary rows
- charts reference summary tables so row additions in the detail tables flow through automatically

## Installation

### Python

Ensure you have python installed on your machine, if not:

[Download the latest version for Windows👈](https://www.python.org/downloads/)

[Download the latest version for Linux/Unix👇]

```bash
sudo apt update
sudo apt install python3 python3-pip
python3 --version
```

### Windows / PowerShell

```bash
python -m venv .venv
.\.venv\Scripts\Activate.ps1
```

### Linux

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

To confirm if you are within the python virtual environment, run:

```bash
python -c "import sys; print(sys.executable)"
```

response should show you a relative path to, e.g:

```bash
C:\path\to\project\.venv\Scripts\python.exe
```

## How to run

```bash
python main.py
```

Backward-compatible wrapper:

```bash
python budget_workbook_generator.py
```

## Manual input guidance

Enter values manually in:

- Budget cells for planned amounts
- Actual cells for real spent / received amounts

Do **not** edit:

- Variance cells
- Year total cells
- rollup formulas

## Safe extension guidance

### Add grocery items

Add rows inside `GroceryDetailTable`. The Groceries summary row, dashboards, and charts will continue to use the table totals.

### Add custom variable expenses

Add rows inside `VariableExpenseTable`. The Additional Variable Items total row and downstream summaries will update automatically.

### Add more months

Update `WorkbookConfig.months` in `budget_workbook/config.py` and regenerate the workbook. All grouped month columns, formulas, and summary sheets are generated from that list.
