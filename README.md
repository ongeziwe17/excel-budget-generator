# Excel Budget Generator

A modular Python project that generates an enhanced personal budgeting workbook in Excel using `openpyxl`.

## What the workbook now supports

- Budget vs Actual tracking for each month
- automatic Variance calculations using `Actual - Budget`
- conditional formatting for overspend / under-budget visibility
- a structured grocery detail table that rolls into the main Groceries row
- an expandable custom variable-expense table
- formula-driven summary totals, dashboards, and trend analysis
- bonus tracking, emergency fund tracking, and a five-year projection

## Project structure

```text
excel-budget-generator/
├── budget_workbook/
│   ├── __init__.py
│   ├── config.py
│   ├── generator.py
│   ├── rows.py
│   ├── styles.py
│   └── builders/
│       ├── base.py
│       ├── bonus_tracker.py
│       ├── cover.py
│       ├── emergency_fund.py
│       ├── five_year_projection.py
│       ├── monthly_entry.py
│       ├── summary_dashboard.py
│       └── trend_analysis.py
├── budget_workbook_generator.py
├── instructions.md
├── main.py
├── requirements.txt
└── README.md
```

## Key workbook enhancements

### 1. Budget vs Actual columns
The Monthly Entry sheet now uses Budget, Actual, and Variance triplets for every month and year total block.

### 2. Expandable detail tables
Two Excel tables keep the model extendable:

- `GroceryDetailTable`
- `VariableExpenseTable`

New rows added inside those tables automatically feed the summary Groceries and Additional Variable Items rows.

### 3. Grocery starter structure
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

### 4. Dynamic rollups
The workbook keeps calculations formula-driven:

- Groceries summary rows pull from `GroceryDetailTable`
- Additional variable totals pull from `VariableExpenseTable`
- dashboard metrics reference Monthly Entry summary rows
- charts reference summary tables so row additions in the detail tables flow through automatically

## Installation

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
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
