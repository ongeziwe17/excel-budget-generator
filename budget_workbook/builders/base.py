"""Base classes and helper utilities for worksheet builders."""

from __future__ import annotations

from abc import ABC, abstractmethod

from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from budget_workbook.config import WorkbookConfig
from budget_workbook.rows import MonthlyEntryRows
from budget_workbook.styles import WorkbookStyles


class BaseSheetBuilder(ABC):
    """Base class for every worksheet builder."""

    def __init__(
        self,
        config: WorkbookConfig,
        styles: WorkbookStyles,
        rows: MonthlyEntryRows,
    ) -> None:
        self.config = config
        self.styles = styles
        self.rows = rows

    @abstractmethod
    def build(self, workbook):
        """Create and return the worksheet."""

    def apply_border_range(
        self,
        worksheet: Worksheet,
        start_row: int,
        end_row: int,
        start_col: int,
        end_col: int,
    ) -> None:
        for row in range(start_row, end_row + 1):
            for col in range(start_col, end_col + 1):
                worksheet.cell(row=row, column=col).border = self.styles.thin_border

    def create_section_header(
        self,
        worksheet: Worksheet,
        row: int,
        title: str,
        start_col: int = 2,
        end_col: int = 15,
    ) -> None:
        worksheet.merge_cells(
            start_row=row,
            start_column=start_col,
            end_row=row,
            end_column=end_col,
        )
        cell = worksheet.cell(row=row, column=start_col)
        cell.value = title
        cell.font = self.styles.section_font()
        cell.fill = self.styles.section_fill()
        cell.alignment = self.styles.left_alignment()
        self.apply_border_range(worksheet, row, row, start_col, end_col)
        worksheet.row_dimensions[row].height = 22

    def create_input_row(
        self,
        worksheet: Worksheet,
        row: int,
        label: str,
        *,
        indent: bool = False,
        start_col: int = 2,
        num_months: int = 12,
    ) -> None:
        label_cell = worksheet.cell(row=row, column=start_col)
        label_cell.value = f"  {label}" if indent else label
        label_cell.font = self._label_font(indent=indent)
        label_cell.alignment = self.styles.left_alignment()
        label_cell.border = self.styles.thin_border

        for index in range(num_months):
            col = start_col + 1 + index
            cell = worksheet.cell(row=row, column=col)
            cell.border = self.styles.thin_border
            cell.alignment = self.styles.right_alignment()
            cell.number_format = self.config.currency_format
            cell.font = self.styles.input_font()

        total_col = start_col + 1 + num_months
        total_cell = worksheet.cell(row=row, column=total_col)
        total_cell.border = self.styles.thin_border
        total_cell.alignment = self.styles.right_alignment()
        total_cell.number_format = self.config.currency_format
        total_cell.font = self.styles.calc_font()
        start_month_col = get_column_letter(start_col + 1)
        end_month_col = get_column_letter(start_col + num_months)
        total_cell.value = f"=SUM({start_month_col}{row}:{end_month_col}{row})"

        self.apply_alternating_fill(worksheet, row, start_col, total_col)

    def create_formula_row(
        self,
        worksheet: Worksheet,
        row: int,
        label: str,
        formula_template: str,
        *,
        indent: bool = False,
        use_percent: bool = False,
        start_col: int = 2,
        num_months: int = 12,
    ) -> None:
        label_cell = worksheet.cell(row=row, column=start_col)
        label_cell.value = f"  {label}" if indent else label
        label_cell.font = self._label_font(indent=indent)
        label_cell.alignment = self.styles.left_alignment()
        label_cell.border = self.styles.thin_border

        for index in range(num_months):
            col = start_col + 1 + index
            col_letter = get_column_letter(col)
            cell = worksheet.cell(row=row, column=col)
            cell.value = formula_template.replace("{col}", col_letter)
            cell.font = self.styles.calc_font()
            cell.border = self.styles.thin_border
            cell.alignment = self.styles.right_alignment()
            cell.number_format = (
                self.config.percent_format if use_percent else self.config.currency_format
            )

        total_col = start_col + 1 + num_months
        total_cell = worksheet.cell(row=row, column=total_col)
        start_month_col = get_column_letter(start_col + 1)
        end_month_col = get_column_letter(start_col + num_months)
        total_cell.value = f"=SUM({start_month_col}{row}:{end_month_col}{row})"
        total_cell.font = self.styles.calc_font()
        total_cell.border = self.styles.thin_border
        total_cell.alignment = self.styles.right_alignment()
        total_cell.number_format = (
            self.config.percent_format if use_percent else self.config.currency_format
        )

        self.apply_alternating_fill(worksheet, row, start_col, total_col)

    def apply_alternating_fill(
        self,
        worksheet: Worksheet,
        row: int,
        start_col: int,
        end_col: int,
    ) -> None:
        if row % 2 == 0:
            for col in range(start_col, end_col + 1):
                worksheet.cell(row=row, column=col).fill = self.styles.alt_row_fill()

    def set_standard_header(self, cell) -> None:
        cell.font = self.styles.header_font()
        cell.fill = self.styles.header_fill()
        cell.border = self.styles.thin_border
        cell.alignment = self.styles.center_alignment()

    def _label_font(self, *, indent: bool):
        from openpyxl.styles import Font

        return Font(size=10, bold=not indent, color=self.styles.palette.text_dark)
