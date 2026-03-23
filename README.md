# Excel Budget Generator

A maintainable, modular Python project that generates a personal budgeting workbook in Excel using `openpyxl`.

## What it generates

The workbook includes the following sheets:

- Cover
- Monthly Entry
- Summary Dashboard
- Trend Analysis
- Bonus Tracker
- Emergency Fund
- 5-Year Projection

The generated workbook supports:

- manual monthly entry for income and expenses
- automatic yearly totals
- calculated savings metrics
- spending and savings trend analysis
- bonus tracking
- emergency fund progress tracking
- five-year savings projections

## Recommended project structure

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
├── main.py
├── requirements.txt
├── instructions.md
└── README.md
```

## Module overview

### `budget_workbook/config.py`
Holds workbook defaults and business configuration in a small immutable dataclass.

### `budget_workbook/styles.py`
Centralizes palette and reusable style factories so formatting is consistent and easy to update.

### `budget_workbook/rows.py`
Stores Monthly Entry row references in one place to keep formulas readable and maintainable.

### `budget_workbook/builders/base.py`
Provides shared helper behavior for sheet builders, including reusable row and section creation logic.

### `budget_workbook/builders/*.py`
Each builder owns one worksheet and keeps worksheet-specific logic isolated.

### `budget_workbook/generator.py`
Coordinates the full workbook build and saves the final `.xlsx` file.

### `main.py`
Simple executable entry point for normal use.

### `budget_workbook_generator.py`
Backward-compatible wrapper that preserves the original function name.

## Improvements made in this refactor

- Split the single-file script into focused modules.
- Introduced class-based sheet builders with clear ownership.
- Centralized styles, config, and row references.
- Preserved the workbook formulas and layout behavior.
- Added a stable entry point and retained backward compatibility.
- Added lightweight project documentation and installation guidance.

## Installation

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## How to run

Generate the workbook with the standard entry point:

```bash
python main.py
```

Or use the compatibility wrapper:

```bash
python budget_workbook_generator.py
```

The workbook will be saved as:

```text
Personal_Budget_Workbook.xlsx
```

## Notes

- The workbook is designed around manual monthly entry.
- Most calculations happen through Excel formulas embedded by `openpyxl`.
- The default currency formatting is South African Rand (ZAR).
