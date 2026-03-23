"""Backward-compatible wrapper for the refactored workbook generator."""

from pathlib import Path

from budget_workbook import BudgetWorkbookGenerator, WorkbookConfig


def create_budget_workbook(output_path: str | None = None) -> Path:
    """Generate the budget workbook using the refactored package."""
    generator = BudgetWorkbookGenerator(WorkbookConfig())
    result = generator.create_workbook(output_path)
    print(f"Budget workbook created successfully: {result}")
    return result


if __name__ == "__main__":
    create_budget_workbook()
