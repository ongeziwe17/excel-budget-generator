# Instructions

## Prerequisites

- Python 3.11+ recommended
- `pip`
- Microsoft Excel, LibreOffice Calc, or another spreadsheet application capable of opening `.xlsx` files

## Setup

1. Create and activate a virtual environment.
2. Install project dependencies.

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## Run the workbook generator

```bash
python main.py
```

## Alternate backward-compatible entry point

```bash
python budget_workbook_generator.py
```

## Output

The default output file is:

```text
Personal_Budget_Workbook.xlsx
```

## Customization guidance

If you need to change workbook defaults, update `WorkbookConfig` in `budget_workbook/config.py`.

Common examples:

- change default rent amount
- change growth rate assumption
- change output file name
- change salary and bonus assumptions

## Developer notes

- Add shared sheet helpers in `budget_workbook/builders/base.py`.
- Add new worksheets as builder classes under `budget_workbook/builders/`.
- Register new sheet builders in `budget_workbook/generator.py`.
