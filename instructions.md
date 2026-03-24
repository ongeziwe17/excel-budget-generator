# Instructions

## Prerequisites

- Python 3.10+
- `pip`
- Excel or another `.xlsx`-compatible spreadsheet application

## Setup

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## Generate the workbook

```bash
python main.py
```

## Backward-compatible entry point

```bash
python budget_workbook_generator.py
```

## Where files are stored

All generated workbooks must be stored in:

```text
/workbooks/
```

The default generated file is:

```text
workbooks/Personal_Budget_Workbook_v0.0.0.xlsx
```

## Versioning approach

Workbook filenames follow semantic-style versioning:

- `v0.0.0` for the initial versioned workbook
- `v0.1.0` for minor workbook enhancements
- `v1.0.0` for a stable major release

Update `WorkbookConfig.workbook_version` before generating a new structural version.

## Important workbook handling rules

- Do not overwrite older workbook versions
- Keep previous files in `workbooks/` for rollback and tracking
- If generation fails because the file already exists, bump the version and rerun

## Where to enter values manually

Enter values on the **Monthly Entry** sheet in:

- Budget cells for planned amounts
- Actual cells for real money flow

Leave formula-driven cells alone:

- Variance
- Year totals
- Groceries rollups
- Additional variable rollups
- dashboard metrics

## Expandable sections

### Grocery detail
Use `GroceryDetailTable` for starter grocery rows and future additions.

### Custom variable items
Use `VariableExpenseTable` for new variable-expense rows that do not fit the main fixed categories.

## Extending months safely

To add months beyond the current configuration, update `WorkbookConfig.months` and regenerate the workbook so every sheet, formula block, and chart uses the same month layout.
