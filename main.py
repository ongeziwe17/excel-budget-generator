"""CLI entry point for generating the budget workbook."""

from budget_workbook import BudgetWorkbookGenerator


def main() -> None:
    """Generate the default budget workbook."""
    output_path = BudgetWorkbookGenerator().create_workbook()
    print(f"Budget workbook created successfully: {output_path}")


if __name__ == "__main__":
    main()
