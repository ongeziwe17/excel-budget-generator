"""Bonus tracker worksheet builder."""

from openpyxl.styles import Font

from .base import BaseSheetBuilder


class BonusTrackerSheetBuilder(BaseSheetBuilder):
    """Build the Bonus Tracker worksheet."""

    def build(self, workbook):
        worksheet = workbook.create_sheet("Bonus Tracker")
        worksheet.sheet_view.showGridLines = False
        worksheet.column_dimensions["A"].width = 3
        worksheet.column_dimensions["B"].width = 20
        worksheet.column_dimensions["C"].width = 16
        worksheet.column_dimensions["D"].width = 16
        worksheet.column_dimensions["E"].width = 16

        worksheet.merge_cells("B2:E2")
        worksheet["B2"] = "BONUS TRACKER"
        worksheet["B2"].font = Font(size=18, bold=True, color=self.styles.palette.header_dark)
        worksheet["B2"].alignment = self.styles.left_alignment()
        worksheet.row_dimensions[2].height = 30

        worksheet.merge_cells("B3:E3")
        worksheet["B3"] = "Monthly and quarterly bonus budget vs actual rollups from the Monthly Entry sheet."
        worksheet["B3"].font = Font(size=10, italic=True, color="666666")

        self._create_monthly_table(worksheet)
        self._create_quarterly_table(worksheet)
        return worksheet

    def _create_monthly_table(self, worksheet) -> None:
        worksheet["B5"] = "MONTHLY BONUS TRACKING"
        worksheet["B5"].font = Font(size=12, bold=True, color=self.styles.palette.header_dark)

        headers = ["Month", "Budget", "Actual", "Variance"]
        for index, header in enumerate(headers, start=2):
            cell = worksheet.cell(row=7, column=index)
            cell.value = header
            cell.font = self.styles.header_font()
            cell.fill = self.styles.header_fill()
            cell.border = self.styles.thin_border
            cell.alignment = self.styles.left_alignment() if index == 2 else self.styles.right_alignment()

        row = 8
        for index, month in enumerate(self.config.months):
            worksheet[f"B{row}"] = month
            worksheet[f"B{row}"].font = Font(size=10, color=self.styles.palette.text_dark)
            worksheet[f"B{row}"].border = self.styles.thin_border
            worksheet[f"B{row}"].alignment = self.styles.left_alignment()

            worksheet[f"C{row}"] = f"='Monthly Entry'!{self.month_budget_letter(index)}{self.rows.bonus}"
            worksheet[f"D{row}"] = f"='Monthly Entry'!{self.month_actual_letter(index)}{self.rows.bonus}"
            worksheet[f"E{row}"] = f"='Monthly Entry'!{self.month_variance_letter(index)}{self.rows.bonus}"
            for col in ["C", "D", "E"]:
                worksheet[f"{col}{row}"].font = self.styles.calc_font()
                worksheet[f"{col}{row}"].border = self.styles.thin_border
                worksheet[f"{col}{row}"].alignment = self.styles.right_alignment()
                worksheet[f"{col}{row}"].number_format = self.config.currency_format
            row += 1

        worksheet[f"B{row}"] = "Year Total"
        worksheet[f"C{row}"] = f"='Monthly Entry'!{self.year_budget_letter()}{self.rows.bonus}"
        worksheet[f"D{row}"] = f"='Monthly Entry'!{self.year_actual_letter()}{self.rows.bonus}"
        worksheet[f"E{row}"] = f"='Monthly Entry'!{self.year_variance_letter()}{self.rows.bonus}"
        for col in ["B", "C", "D", "E"]:
            worksheet[f"{col}{row}"].font = Font(bold=True, color=self.styles.palette.white if col != "B" else self.styles.palette.white)
            worksheet[f"{col}{row}"].fill = self.styles.header_fill()
            worksheet[f"{col}{row}"].border = self.styles.thin_border
            worksheet[f"{col}{row}"].alignment = self.styles.left_alignment() if col == "B" else self.styles.right_alignment()
            worksheet[f"{col}{row}"].number_format = self.config.currency_format if col != "B" else "General"

        self.add_variance_conditional_formatting(worksheet, 8, row, columns=[5])

    def _create_quarterly_table(self, worksheet) -> None:
        start_row = 23
        worksheet[f"B{start_row}"] = "QUARTERLY BONUS ROLLUP"
        worksheet[f"B{start_row}"].font = Font(size=12, bold=True, color=self.styles.palette.header_dark)
        header_row = start_row + 2
        for index, header in enumerate(["Quarter", "Budget", "Actual", "Variance"], start=2):
            cell = worksheet.cell(row=header_row, column=index)
            cell.value = header
            cell.font = self.styles.header_font()
            cell.fill = self.styles.header_fill()
            cell.border = self.styles.thin_border
            cell.alignment = self.styles.left_alignment() if index == 2 else self.styles.right_alignment()

        quarter_slices = [
            ("Q1", 0, 3),
            ("Q2", 3, 6),
            ("Q3", 6, 9),
            ("Q4", 9, 12),
        ]
        row = header_row + 1
        for quarter, start_index, end_index in quarter_slices:
            worksheet[f"B{row}"] = quarter
            worksheet[f"B{row}"].font = Font(size=10, color=self.styles.palette.text_dark)
            worksheet[f"B{row}"].border = self.styles.thin_border
            worksheet[f"B{row}"].alignment = self.styles.left_alignment()
            budget_cells = [f"C{8 + idx}" for idx in range(start_index, min(end_index, len(self.config.months)))]
            actual_cells = [f"D{8 + idx}" for idx in range(start_index, min(end_index, len(self.config.months)))]
            worksheet[f"C{row}"] = f"=SUM({','.join(budget_cells)})" if budget_cells else 0
            worksheet[f"D{row}"] = f"=SUM({','.join(actual_cells)})" if actual_cells else 0
            worksheet[f"E{row}"] = f"=D{row}-C{row}"
            for col in ["C", "D", "E"]:
                worksheet[f"{col}{row}"].font = self.styles.calc_font()
                worksheet[f"{col}{row}"].border = self.styles.thin_border
                worksheet[f"{col}{row}"].alignment = self.styles.right_alignment()
                worksheet[f"{col}{row}"].number_format = self.config.currency_format
            row += 1

        self.add_variance_conditional_formatting(worksheet, header_row + 1, row - 1, columns=[5])
