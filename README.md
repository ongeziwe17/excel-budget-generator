# Excel Budget Generator

A modular Python project that generates personal or two-person household budgeting workbooks in Excel using `openpyxl`.

## What the workbook supports

- Budget vs Actual tracking for each month
- automatic Variance calculations using `Actual - Budget`
- conditional formatting for overspend / under-budget visibility
- a structured grocery detail table that rolls into the main Groceries row
- an expandable custom variable-expense table
- formula-driven summary totals, dashboards, and trend analysis
- bonus tracking, emergency fund tracking, and a five-year projection
- an optional couple mode with two personal budgets, a shared-household budget, configurable contribution allocation, and a consolidated dashboard

## Budget modes

### Single-person mode

The default remains the original personal workbook, so existing commands and integrations remain compatible.

```bash
python main.py
```

### Couple mode

Couple mode creates separate personal budgets for both partners, a shared-household budget, and consolidated household reporting.

```bash
python main.py \
  --mode couple \
  --partner-one-name "Partner 1" \
  --partner-two-name "Partner 2" \
  --partner-one-gross-income 36000 \
  --partner-two-gross-income 24000 \
  --partner-one-net-income 30000 \
  --partner-two-net-income 20000
```

The generated workbook starts at:

```text
workbooks/Household_Budget_Workbook_v0.1.0.xlsx
```

Couple workbooks contain:

- `Cover`
- `Household Dashboard`
- `Household Setup`
- `Partner 1 Budget`
- `Partner 2 Budget`
- `Shared Household`
- `Trend Analysis`
- `Savings Goals`
- `Bonus Tracker`
- `5-Year Projection`

Personal and shared expenses are deliberately separate. Partner contributions to the shared household are internal transfers and are excluded from consolidated income and expense totals.

## Shared-cost allocation

The `Household Setup` sheet supports three methods:

- `Income proportional` (default): each partner's share follows their configured net-income share
- `Equal`: each partner contributes 50%
- `Custom`: use explicit partner percentages that add up to 100%

The method can be selected when generating the workbook:

```bash
python main.py --mode couple --contribution-method Equal
```

For a custom 60% / 40% split:

```bash
python main.py \
  --mode couple \
  --contribution-method Custom \
  --partner-one-share 0.60 \
  --partner-two-share 0.40
```

The allocation method and custom percentages remain editable on `Household Setup` after generation.

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
│       ├── couple.py
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
- Repeated runs auto-create the next patch version (for example `v0.0.0`, `v0.0.1`, `v0.0.2`)
- Older files stay in place for rollback and historical tracking
- Invalid or unrelated filenames in `workbooks/` are ignored during version detection


## Automatic version progression

Running `python main.py` repeatedly does not fail on existing files.

Example sequence:

```text
workbooks/Personal_Budget_Workbook_v0.0.0.xlsx
workbooks/Personal_Budget_Workbook_v0.0.1.xlsx
workbooks/Personal_Budget_Workbook_v0.0.2.xlsx
```

By default, generation increments the **patch** number using the highest existing patch for the configured major/minor pair.

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

In couple mode:

- review names, income assumptions, allocation method and opening savings on `Household Setup`
- enter personal actual income, expenses, savings and household contributions on each partner sheet
- enter shared expense budgets/actuals and joint savings on `Shared Household`
- do not enter partner contributions again on `Shared Household`; they are linked from the partner sheets
- use `Household Dashboard` for combined results

## Safe extension guidance

### Add grocery items
Add rows inside `GroceryDetailTable`. The Groceries summary row, dashboards, and charts will continue to use the table totals.

### Add custom variable expenses
Add rows inside `VariableExpenseTable`. The Additional Variable Items total row and downstream summaries will update automatically.

### Add more months
Update `WorkbookConfig.months` in `budget_workbook/config.py` and regenerate the workbook. All grouped month columns, formulas, and summary sheets are generated from that list.
