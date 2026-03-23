"""Trend analysis worksheet builder."""

from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

from .base import BaseSheetBuilder


class TrendAnalysisSheetBuilder(BaseSheetBuilder):
    """Build the Trend Analysis worksheet."""

    def build(self, workbook):
        worksheet = workbook.create_sheet("Trend Analysis")
        worksheet.sheet_view.showGridLines = False

        worksheet.column_dimensions["A"].width = 3
        worksheet.column_dimensions["B"].width = 28
        for index in range(11):
            worksheet.column_dimensions[get_column_letter(3 + index)].width = 14

        worksheet.merge_cells("B2:N2")
        worksheet["B2"] = "TREND ANALYSIS"
        worksheet["B2"].font = Font(size=18, bold=True, color=self.styles.palette.header_dark)
        worksheet["B2"].alignment = self.styles.left_alignment()
        worksheet.row_dimensions[2].height = 30

        worksheet.merge_cells("B3:N3")
        worksheet["B3"] = (
            "Month-to-month changes - GREEN means improvement, RED means increase in spending"
        )
        worksheet["B3"].font = Font(size=10, italic=True, color="666666")

        worksheet["B4"] = "▼ = Spending Decreased (Good) | ▲ = Spending Increased (Warning)"
        worksheet["B4"].font = Font(size=9, color="666666")

        worksheet["B6"] = "METRIC"
        worksheet["B6"].font = self.styles.header_font()
        worksheet["B6"].fill = self.styles.header_fill()
        worksheet["B6"].border = self.styles.thin_border
        worksheet["B6"].alignment = self.styles.left_alignment()

        change_labels = [
            "Feb-Jan",
            "Mar-Feb",
            "Apr-Mar",
            "May-Apr",
            "Jun-May",
            "Jul-Jun",
            "Aug-Jul",
            "Sep-Aug",
            "Oct-Sep",
            "Nov-Oct",
            "Dec-Nov",
        ]
        for index, label in enumerate(change_labels):
            col = get_column_letter(3 + index)
            worksheet[f"{col}6"] = label
            self.set_standard_header(worksheet[f"{col}6"])
        worksheet.row_dimensions[6].height = 25

        row = 8
        row = self._create_section_banner(
            worksheet,
            row,
            "SPENDING TRENDS (RED = Increased, GREEN = Decreased)",
            self.styles.negative_fill(),
        )

        for label, monthly_row in [
            ("Total Spending Change", self.rows.total_expenses),
            ("Fixed Expenses Change", self.rows.total_fixed),
            ("Variable Expenses Change", self.rows.total_variable),
            ("  Data", self.rows.data),
            ("  Transport", self.rows.transport),
            ("  Subscriptions", self.rows.subscriptions),
            ("  TFG Debit", self.rows.tfg_debit),
            ("  Family Support", self.rows.family_support),
            ("  Tithe", self.rows.tithe),
            ("  Groceries", self.rows.groceries),
            ("  Eating Out", self.rows.eating_out),
            ("  Lunch", self.rows.lunch),
            ("  Haircuts", self.rows.haircuts),
            ("  Clothing", self.rows.clothing),
            ("  Random Spending", self.rows.random_spending),
        ]:
            self._create_trend_row(worksheet, row, label, monthly_row)
            row += 1

        row += 1
        row = self._create_section_banner(
            worksheet,
            row,
            "SAVINGS TRENDS (GREEN = Increased, RED = Decreased)",
            self.styles.positive_fill(),
        )
        self._create_trend_row(worksheet, row, "Savings Transfer Change", self.rows.savings_transfer)
        row += 1
        self._create_trend_row(
            worksheet,
            row,
            "Savings Rate % Change",
            self.rows.savings_rate,
            use_percent=True,
        )
        row += 2

        row = self._create_section_banner(
            worksheet,
            row,
            "INCOME TRENDS",
            self.styles.section_fill(),
        )
        self._create_trend_row(worksheet, row, "Net Income Change", self.rows.net_income)
        row += 1
        self._create_trend_row(worksheet, row, "Bonus Received Change", self.rows.bonus)
        row += 1
        self._create_trend_row(worksheet, row, "Total Income Change", self.rows.total_income)

        self._apply_conditional_formatting(worksheet)
        return worksheet

    def _create_section_banner(self, worksheet, row: int, title: str, fill) -> int:
        worksheet.merge_cells(f"B{row}:N{row}")
        worksheet[f"B{row}"] = title
        worksheet[f"B{row}"].font = self.styles.section_font()
        worksheet[f"B{row}"].fill = fill
        for col in range(2, 15):
            worksheet.cell(row=row, column=col).border = self.styles.thin_border
        return row + 1

    def _create_trend_row(
        self,
        worksheet,
        row: int,
        label: str,
        monthly_row: int,
        *,
        use_percent: bool = False,
    ) -> None:
        worksheet[f"B{row}"] = label
        worksheet[f"B{row}"].font = Font(size=10, color=self.styles.palette.text_dark)
        worksheet[f"B{row}"].border = self.styles.thin_border
        worksheet[f"B{row}"].alignment = self.styles.left_alignment()

        month_pairs = [
            ("D", "C"),
            ("E", "D"),
            ("F", "E"),
            ("G", "F"),
            ("H", "G"),
            ("I", "H"),
            ("J", "I"),
            ("K", "J"),
            ("L", "K"),
            ("M", "L"),
            ("N", "M"),
        ]
        for current_col, previous_col in month_pairs:
            cell = worksheet[f"{current_col}{row}"]
            cell.value = (
                f"='Monthly Entry'!{current_col}{monthly_row}-'Monthly Entry'!{previous_col}{monthly_row}"
            )
            cell.font = self.styles.calc_font()
            cell.border = self.styles.thin_border
            cell.alignment = self.styles.right_alignment()
            cell.number_format = self.config.percent_format if use_percent else self.config.currency_format

        if row % 2 == 0:
            worksheet[f"B{row}"].fill = self.styles.alt_row_fill()

    def _apply_conditional_formatting(self, worksheet) -> None:
        green_fill = PatternFill(
            start_color=self.styles.palette.green_fill,
            end_color=self.styles.palette.green_fill,
            fill_type="solid",
        )
        red_fill = PatternFill(
            start_color=self.styles.palette.red_fill,
            end_color=self.styles.palette.red_fill,
            fill_type="solid",
        )

        for row in range(9, 24):
            for col_letter in ["D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N"]:
                worksheet.conditional_formatting.add(
                    f"{col_letter}{row}",
                    CellIsRule(operator="greaterThan", formula=["0"], fill=red_fill),
                )
                worksheet.conditional_formatting.add(
                    f"{col_letter}{row}",
                    CellIsRule(operator="lessThan", formula=["0"], fill=green_fill),
                )

        for row in [25, 26]:
            for col_letter in ["D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N"]:
                worksheet.conditional_formatting.add(
                    f"{col_letter}{row}",
                    CellIsRule(operator="greaterThan", formula=["0"], fill=green_fill),
                )
                worksheet.conditional_formatting.add(
                    f"{col_letter}{row}",
                    CellIsRule(operator="lessThan", formula=["0"], fill=red_fill),
                )
