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

This preserves the original single-person workbook.

## Generate a couple workbook

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

Couple mode keeps each person's finances separate from shared household costs. The consolidated dashboard excludes household contributions because they are internal transfers.

Available shared-cost allocation options:

```bash
python main.py --mode couple --contribution-method Equal
```

```bash
python main.py \
  --mode couple \
  --contribution-method Custom \
  --partner-one-share 0.60 \
  --partner-two-share 0.40
```

Custom shares must be decimals between `0` and `1` and must add up to `1`.

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

The initial couple workbook is:

```text
workbooks/Household_Budget_Workbook_v0.1.0.xlsx
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
- Repeated runs auto-generate the next patch versioned filename
- Invalid or unrelated files in `workbooks/` are ignored by version detection

Example repeated runs:

```text
workbooks/Personal_Budget_Workbook_v0.0.0.xlsx
workbooks/Personal_Budget_Workbook_v0.0.1.xlsx
workbooks/Personal_Budget_Workbook_v0.0.2.xlsx
```

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

For couple workbooks:

- edit household assumptions on `Household Setup`
- enter personal actuals on `Partner 1 Budget` and `Partner 2 Budget`
- enter shared expense and joint-savings values on `Shared Household`
- review combined performance on `Household Dashboard`
- never enter household contributions a second time on the shared sheet

## Expandable sections

### Grocery detail
Use `GroceryDetailTable` for starter grocery rows and future additions.

### Custom variable items
Use `VariableExpenseTable` for new variable-expense rows that do not fit the main fixed categories.

## Extending months safely

To add months beyond the current configuration, update `WorkbookConfig.months` and regenerate the workbook so every sheet, formula block, and chart uses the same month layout.
