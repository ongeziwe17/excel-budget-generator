"""Monthly entry worksheet builder."""

from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

from .base import BaseSheetBuilder


class MonthlyEntrySheetBuilder(BaseSheetBuilder):
    """Build the Monthly Entry worksheet."""

    def build(self, workbook):
        worksheet = workbook.create_sheet("Monthly Entry")
        worksheet.sheet_view.showGridLines = False

        worksheet.column_dimensions["A"].width = 3
        worksheet.column_dimensions["B"].width = 28
        for index in range(12):
            worksheet.column_dimensions[get_column_letter(3 + index)].width = 14
        worksheet.column_dimensions["O"].width = 16
        worksheet.column_dimensions["Q"].hidden = True
        worksheet.column_dimensions["R"].hidden = True

        worksheet.merge_cells("B2:O2")
        worksheet["B2"] = "MONTHLY DATA ENTRY"
        worksheet["B2"].font = Font(size=18, bold=True, color=self.styles.palette.header_dark)
        worksheet["B2"].alignment = self.styles.left_alignment()
        worksheet.row_dimensions[2].height = 30

        worksheet.merge_cells("B3:O3")
        worksheet["B3"] = "Enter your actual income and expenses in the BLUE cells below."
        worksheet["B3"].font = Font(size=10, italic=True, color="666666")

        worksheet["B5"] = "CATEGORY"
        worksheet["B5"].font = self.styles.header_font()
        worksheet["B5"].fill = self.styles.header_fill()
        worksheet["B5"].alignment = self.styles.left_alignment()
        worksheet["B5"].border = self.styles.thin_border

        for index, month in enumerate(self.config.months):
            col = get_column_letter(3 + index)
            worksheet[f"{col}5"] = month
            self.set_standard_header(worksheet[f"{col}5"])

        worksheet["O5"] = "YEAR TOTAL"
        self.set_standard_header(worksheet["O5"])
        worksheet.row_dimensions[5].height = 25

        row = 7
        self.create_section_header(worksheet, row, "INCOME")
        row += 1
        self.create_input_row(worksheet, row, "Net Income Received")
        net_income_row = row
        row += 2

        self.create_section_header(worksheet, row, "FIXED EXPENSES")
        row += 1
        for expense in ["Data", "Transport", "Subscriptions", "TFG Debit", "Family Support"]:
            self.create_input_row(worksheet, row, expense, indent=True)
            row += 1

        self.create_formula_row(
            worksheet,
            row,
            "Tithe (10% of Net Income)",
            f"={{col}}{net_income_row}*0.1",
            indent=True,
        )
        row += 2

        self.create_section_header(worksheet, row, "VARIABLE EXPENSES")
        row += 1
        for expense in [
            "Groceries",
            "Eating Out",
            "Lunch",
            "Haircuts",
            "Clothing",
            "Random Spending",
        ]:
            self.create_input_row(worksheet, row, expense, indent=True)
            row += 1

        row += 1
        self.create_section_header(worksheet, row, "FUTURE RENT (Toggle On/Off)")
        row += 1
        self._create_rent_rows(worksheet, row)
        row += 4

        self.create_section_header(worksheet, row, "SAVINGS & TRANSFERS")
        row += 1
        self.create_input_row(worksheet, row, "Transfer to Savings")
        savings_row = row
        row += 1
        self.create_input_row(worksheet, row, "Starting Savings Balance")
        starting_savings_row = row
        row += 2

        self.create_section_header(worksheet, row, "BONUSES RECEIVED")
        row += 1
        self.create_input_row(worksheet, row, "Bonus Received")
        bonus_row = row
        row += 2

        self.create_section_header(worksheet, row, "UNEXPECTED EXPENSES")
        row += 1
        self.create_input_row(worksheet, row, "Unexpected Expenses")
        unexpected_row = row
        row += 2

        self.create_section_header(worksheet, row, "SUMMARY TOTALS")
        row += 1
        row = self._create_summary_rows(
            worksheet,
            row,
            net_income_row=net_income_row,
            bonus_row=bonus_row,
            savings_row=savings_row,
            starting_savings_row=starting_savings_row,
            unexpected_row=unexpected_row,
        )

        self._store_hidden_references(worksheet)
        return worksheet

    def _create_rent_rows(self, worksheet, row: int) -> None:
        worksheet[f"B{row}"] = "  Rent Active? (1=Yes, 0=No)"
        worksheet[f"B{row}"].font = Font(size=10, color=self.styles.palette.text_dark)
        worksheet[f"B{row}"].alignment = self.styles.left_alignment()
        worksheet[f"B{row}"].border = self.styles.thin_border

        rent_toggle_row = row
        for index in range(12):
            col = get_column_letter(3 + index)
            cell = worksheet[f"{col}{row}"]
            cell.value = 0
            cell.font = self.styles.input_font()
            cell.border = self.styles.thin_border
            cell.alignment = self.styles.center_alignment()
            cell.number_format = "0"
        row += 1

        worksheet[f"B{row}"] = f"  Rent Amount (R{self.config.default_rent_amount:,})"
        worksheet[f"B{row}"].font = Font(size=10, color=self.styles.palette.text_dark)
        worksheet[f"B{row}"].alignment = self.styles.left_alignment()
        worksheet[f"B{row}"].border = self.styles.thin_border

        rent_amount_row = row
        for index in range(12):
            col = get_column_letter(3 + index)
            cell = worksheet[f"{col}{row}"]
            cell.value = self.config.default_rent_amount
            cell.font = self.styles.input_font()
            cell.border = self.styles.thin_border
            cell.alignment = self.styles.right_alignment()
            cell.number_format = self.config.currency_format
        row += 1

        worksheet[f"B{row}"] = "  Rent Paid (calculated)"
        worksheet[f"B{row}"].font = Font(size=10, color=self.styles.palette.text_dark)
        worksheet[f"B{row}"].alignment = self.styles.left_alignment()
        worksheet[f"B{row}"].border = self.styles.thin_border

        for index in range(12):
            col = get_column_letter(3 + index)
            cell = worksheet[f"{col}{row}"]
            cell.value = f"=IF({col}{rent_toggle_row}=1,{col}{rent_amount_row},0)"
            cell.font = self.styles.calc_font()
            cell.border = self.styles.thin_border
            cell.alignment = self.styles.right_alignment()
            cell.number_format = self.config.currency_format

        worksheet[f"O{row}"] = f"=SUM(C{row}:N{row})"
        worksheet[f"O{row}"].font = self.styles.calc_font()
        worksheet[f"O{row}"].border = self.styles.thin_border
        worksheet[f"O{row}"].alignment = self.styles.right_alignment()
        worksheet[f"O{row}"].number_format = self.config.currency_format

    def _create_summary_rows(
        self,
        worksheet,
        row: int,
        *,
        net_income_row: int,
        bonus_row: int,
        savings_row: int,
        starting_savings_row: int,
        unexpected_row: int,
    ) -> int:
        row = self._create_total_fixed_row(worksheet, row)
        row = self._create_total_variable_row(worksheet, row)
        row = self._create_total_expenses_row(worksheet, row, unexpected_row=unexpected_row)
        row = self._create_total_income_row(worksheet, row, net_income_row=net_income_row, bonus_row=bonus_row)
        row = self._create_surplus_row(worksheet, row, savings_row=savings_row)
        row = self._create_savings_rate_row(worksheet, row, savings_row=savings_row)
        row = self._create_running_savings_row(
            worksheet,
            row,
            starting_savings_row=starting_savings_row,
            savings_row=savings_row,
        )
        return row

    def _create_total_fixed_row(self, worksheet, row: int) -> int:
        worksheet[f"B{row}"] = "Total Fixed Expenses"
        worksheet[f"B{row}"].font = Font(size=10, bold=True, color=self.styles.palette.text_dark)
        worksheet[f"B{row}"].alignment = self.styles.left_alignment()
        worksheet[f"B{row}"].border = self.styles.thin_border

        for index in range(12):
            col = get_column_letter(3 + index)
            cell = worksheet[f"{col}{row}"]
            cell.value = (
                f"={col}{self.rows.data}+{col}{self.rows.transport}+{col}{self.rows.subscriptions}"
                f"+{col}{self.rows.tfg_debit}+{col}{self.rows.family_support}+{col}{self.rows.tithe}"
            )
            cell.font = self.styles.calc_font()
            cell.border = self.styles.thin_border
            cell.alignment = self.styles.right_alignment()
            cell.number_format = self.config.currency_format

        worksheet[f"O{row}"] = f"=SUM(C{row}:N{row})"
        worksheet[f"O{row}"].font = self.styles.calc_font()
        worksheet[f"O{row}"].border = self.styles.thin_border
        worksheet[f"O{row}"].alignment = self.styles.right_alignment()
        worksheet[f"O{row}"].number_format = self.config.currency_format
        return row + 1

    def _create_total_variable_row(self, worksheet, row: int) -> int:
        worksheet[f"B{row}"] = "Total Variable Expenses"
        worksheet[f"B{row}"].font = Font(size=10, bold=True, color=self.styles.palette.text_dark)
        worksheet[f"B{row}"].alignment = self.styles.left_alignment()
        worksheet[f"B{row}"].border = self.styles.thin_border

        for index in range(12):
            col = get_column_letter(3 + index)
            cell = worksheet[f"{col}{row}"]
            cell.value = (
                f"={col}{self.rows.groceries}+{col}{self.rows.eating_out}+{col}{self.rows.lunch}"
                f"+{col}{self.rows.haircuts}+{col}{self.rows.clothing}+{col}{self.rows.random_spending}"
            )
            cell.font = self.styles.calc_font()
            cell.border = self.styles.thin_border
            cell.alignment = self.styles.right_alignment()
            cell.number_format = self.config.currency_format

        worksheet[f"O{row}"] = f"=SUM(C{row}:N{row})"
        worksheet[f"O{row}"].font = self.styles.calc_font()
        worksheet[f"O{row}"].border = self.styles.thin_border
        worksheet[f"O{row}"].alignment = self.styles.right_alignment()
        worksheet[f"O{row}"].number_format = self.config.currency_format
        return row + 1

    def _create_total_expenses_row(self, worksheet, row: int, *, unexpected_row: int) -> int:
        header_fill = self.styles.header_fill()
        worksheet[f"B{row}"] = "TOTAL ALL EXPENSES"
        worksheet[f"B{row}"].font = Font(size=11, bold=True, color=self.styles.palette.white)
        worksheet[f"B{row}"].fill = header_fill
        worksheet[f"B{row}"].alignment = self.styles.left_alignment()
        worksheet[f"B{row}"].border = self.styles.thin_border

        for index in range(12):
            col = get_column_letter(3 + index)
            cell = worksheet[f"{col}{row}"]
            cell.value = f"={col}{row-2}+{col}{row-1}+{col}{row-13}+{col}{unexpected_row}"
            cell.font = Font(bold=True, size=11, color=self.styles.palette.white)
            cell.fill = header_fill
            cell.border = self.styles.thin_border
            cell.alignment = self.styles.right_alignment()
            cell.number_format = self.config.currency_format

        worksheet[f"O{row}"] = f"=SUM(C{row}:N{row})"
        worksheet[f"O{row}"].font = Font(bold=True, size=11, color=self.styles.palette.white)
        worksheet[f"O{row}"].fill = header_fill
        worksheet[f"O{row}"].border = self.styles.thin_border
        worksheet[f"O{row}"].alignment = self.styles.right_alignment()
        worksheet[f"O{row}"].number_format = self.config.currency_format
        return row + 1

    def _create_total_income_row(self, worksheet, row: int, *, net_income_row: int, bonus_row: int) -> int:
        positive_fill = self.styles.positive_fill()
        worksheet[f"B{row}"] = "TOTAL INCOME"
        worksheet[f"B{row}"].font = Font(size=11, bold=True, color=self.styles.palette.white)
        worksheet[f"B{row}"].fill = positive_fill
        worksheet[f"B{row}"].alignment = self.styles.left_alignment()
        worksheet[f"B{row}"].border = self.styles.thin_border

        for index in range(12):
            col = get_column_letter(3 + index)
            cell = worksheet[f"{col}{row}"]
            cell.value = f"={col}{net_income_row}+{col}{bonus_row}"
            cell.font = Font(bold=True, size=11, color=self.styles.palette.white)
            cell.fill = positive_fill
            cell.border = self.styles.thin_border
            cell.alignment = self.styles.right_alignment()
            cell.number_format = self.config.currency_format

        worksheet[f"O{row}"] = f"=SUM(C{row}:N{row})"
        worksheet[f"O{row}"].font = Font(bold=True, size=11, color=self.styles.palette.white)
        worksheet[f"O{row}"].fill = positive_fill
        worksheet[f"O{row}"].border = self.styles.thin_border
        worksheet[f"O{row}"].alignment = self.styles.right_alignment()
        worksheet[f"O{row}"].number_format = self.config.currency_format
        return row + 1

    def _create_surplus_row(self, worksheet, row: int, *, savings_row: int) -> int:
        accent_fill = PatternFill(
            start_color=self.styles.palette.accent_blue,
            end_color=self.styles.palette.accent_blue,
            fill_type="solid",
        )
        worksheet[f"B{row}"] = "SURPLUS / (DEFICIT)"
        worksheet[f"B{row}"].font = Font(size=11, bold=True, color=self.styles.palette.white)
        worksheet[f"B{row}"].fill = accent_fill
        worksheet[f"B{row}"].alignment = self.styles.left_alignment()
        worksheet[f"B{row}"].border = self.styles.thin_border

        for index in range(12):
            col = get_column_letter(3 + index)
            cell = worksheet[f"{col}{row}"]
            cell.value = f"={col}{row-1}-{col}{row-2}-{col}{savings_row}"
            cell.font = Font(bold=True, size=11, color=self.styles.palette.white)
            cell.fill = accent_fill
            cell.border = self.styles.thin_border
            cell.alignment = self.styles.right_alignment()
            cell.number_format = self.config.currency_format

        worksheet[f"O{row}"] = f"=SUM(C{row}:N{row})"
        worksheet[f"O{row}"].font = Font(bold=True, size=11, color=self.styles.palette.white)
        worksheet[f"O{row}"].fill = accent_fill
        worksheet[f"O{row}"].border = self.styles.thin_border
        worksheet[f"O{row}"].alignment = self.styles.right_alignment()
        worksheet[f"O{row}"].number_format = self.config.currency_format
        return row + 1

    def _create_savings_rate_row(self, worksheet, row: int, *, savings_row: int) -> int:
        worksheet[f"B{row}"] = "Savings Rate %"
        worksheet[f"B{row}"].font = Font(size=10, bold=True, color=self.styles.palette.text_dark)
        worksheet[f"B{row}"].alignment = self.styles.left_alignment()
        worksheet[f"B{row}"].border = self.styles.thin_border

        for index in range(12):
            col = get_column_letter(3 + index)
            cell = worksheet[f"{col}{row}"]
            cell.value = f"=IF({col}{row-2}=0,0,{col}{savings_row}/{col}{row-2})"
            cell.font = self.styles.calc_font()
            cell.border = self.styles.thin_border
            cell.alignment = self.styles.right_alignment()
            cell.number_format = self.config.percent_format

        worksheet[f"O{row}"] = f"=IF(O{row-2}=0,0,O{savings_row}/O{row-2})"
        worksheet[f"O{row}"].font = self.styles.calc_font()
        worksheet[f"O{row}"].border = self.styles.thin_border
        worksheet[f"O{row}"].alignment = self.styles.right_alignment()
        worksheet[f"O{row}"].number_format = self.config.percent_format
        return row + 1

    def _create_running_savings_row(
        self,
        worksheet,
        row: int,
        *,
        starting_savings_row: int,
        savings_row: int,
    ) -> int:
        worksheet[f"B{row}"] = "Running Savings Balance"
        worksheet[f"B{row}"].font = Font(size=10, bold=True, color=self.styles.palette.text_dark)
        worksheet[f"B{row}"].alignment = self.styles.left_alignment()
        worksheet[f"B{row}"].border = self.styles.thin_border

        worksheet[f"C{row}"] = f"=C{starting_savings_row}+C{savings_row}"
        worksheet[f"C{row}"].font = self.styles.calc_font()
        worksheet[f"C{row}"].border = self.styles.thin_border
        worksheet[f"C{row}"].alignment = self.styles.right_alignment()
        worksheet[f"C{row}"].number_format = self.config.currency_format

        for index in range(1, 12):
            col = get_column_letter(3 + index)
            previous_col = get_column_letter(2 + index)
            cell = worksheet[f"{col}{row}"]
            cell.value = f"={previous_col}{row}+{col}{savings_row}"
            cell.font = self.styles.calc_font()
            cell.border = self.styles.thin_border
            cell.alignment = self.styles.right_alignment()
            cell.number_format = self.config.currency_format

        worksheet[f"O{row}"] = f"=N{row}"
        worksheet[f"O{row}"].font = self.styles.calc_font()
        worksheet[f"O{row}"].border = self.styles.thin_border
        worksheet[f"O{row}"].alignment = self.styles.right_alignment()
        worksheet[f"O{row}"].number_format = self.config.currency_format
        return row + 1

    def _store_hidden_references(self, worksheet) -> None:
        worksheet["Q5"] = "KEY ROW REFERENCE"
        worksheet["Q5"].font = Font(bold=True, color=self.styles.palette.header_dark)

        references = [
            ("Net Income Row:", self.rows.net_income),
            ("Total Expenses Row:", self.rows.total_expenses),
            ("Total Income Row:", self.rows.total_income),
            ("Savings Row:", self.rows.savings_transfer),
            ("Bonus Row:", self.rows.bonus),
            ("Surplus Row:", self.rows.surplus_deficit),
            ("Savings Rate Row:", self.rows.savings_rate),
            ("Running Savings Row:", self.rows.running_savings),
        ]

        for index, (label, row_number) in enumerate(references):
            worksheet[f"Q{7 + index}"] = label
            worksheet[f"R{7 + index}"] = row_number
