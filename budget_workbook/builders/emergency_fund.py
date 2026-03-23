"""Emergency fund worksheet builder."""

from openpyxl.styles import Font

from .base import BaseSheetBuilder


class EmergencyFundSheetBuilder(BaseSheetBuilder):
    """Build the Emergency Fund worksheet."""

    def build(self, workbook):
        worksheet = workbook.create_sheet("Emergency Fund")
        worksheet.sheet_view.showGridLines = False
        worksheet.column_dimensions["A"].width = 3
        worksheet.column_dimensions["B"].width = 34
        worksheet.column_dimensions["C"].width = 18
        worksheet.column_dimensions["D"].width = 18
        worksheet.column_dimensions["E"].width = 18

        worksheet.merge_cells("B2:E2")
        worksheet["B2"] = "EMERGENCY FUND TRACKER"
        worksheet["B2"].font = Font(size=18, bold=True, color=self.styles.palette.header_dark)
        worksheet["B2"].alignment = self.styles.left_alignment()
        worksheet.row_dimensions[2].height = 30

        worksheet.merge_cells("B3:E3")
        worksheet["B3"] = "Budget vs Actual emergency fund readiness using Monthly Entry totals."
        worksheet["B3"].font = Font(size=10, italic=True, color="666666")

        self._create_status_metrics(worksheet)
        self._create_progress_table(worksheet)
        return worksheet

    def _create_status_metrics(self, worksheet) -> None:
        worksheet["B5"] = "CURRENT STATUS"
        worksheet["B5"].font = Font(size=14, bold=True, color=self.styles.palette.header_dark)
        headers = ["Metric", "Budget", "Actual", "Variance"]
        for index, header in enumerate(headers, start=2):
            cell = worksheet.cell(row=7, column=index)
            cell.value = header
            cell.font = self.styles.header_font()
            cell.fill = self.styles.header_fill()
            cell.border = self.styles.thin_border
            cell.alignment = self.styles.left_alignment() if index == 2 else self.styles.right_alignment()

        metrics = [
            ("Average Monthly Expenses", f"='Monthly Entry'!{self.year_budget_letter()}{self.rows.total_expenses}/12", f"='Monthly Entry'!{self.year_actual_letter()}{self.rows.total_expenses}/12", self.config.currency_format),
            ("6-Month Target", "=C8*6", "=D8*6", self.config.currency_format),
            ("Running Savings Balance", f"='Monthly Entry'!{self.year_budget_letter()}{self.rows.running_savings}", f"='Monthly Entry'!{self.year_actual_letter()}{self.rows.running_savings}", self.config.currency_format),
            ("Progress Toward Target", "=IF(C9=0,0,C10/C9)", "=IF(D9=0,0,D10/D9)", self.config.percent_format),
            ("Amount Still Needed", "=MAX(0,C9-C10)", "=MAX(0,D9-D10)", self.config.currency_format),
        ]
        row = 8
        for label, budget_formula, actual_formula, number_format in metrics:
            worksheet[f"B{row}"] = label
            worksheet[f"B{row}"].font = Font(size=11, color=self.styles.palette.text_dark)
            worksheet[f"B{row}"].border = self.styles.thin_border
            worksheet[f"B{row}"].alignment = self.styles.left_alignment()
            worksheet[f"C{row}"] = budget_formula
            worksheet[f"D{row}"] = actual_formula
            worksheet[f"E{row}"] = f"=D{row}-C{row}"
            for col in ["C", "D", "E"]:
                worksheet[f"{col}{row}"].font = self.styles.calc_font()
                worksheet[f"{col}{row}"].border = self.styles.thin_border
                worksheet[f"{col}{row}"].alignment = self.styles.right_alignment()
                worksheet[f"{col}{row}"].number_format = number_format
            row += 1
        self.add_variance_conditional_formatting(worksheet, 8, row - 1, columns=[5])

    def _create_progress_table(self, worksheet) -> None:
        start_row = 16
        worksheet[f"B{start_row}"] = "MONTHLY SAVINGS PROGRESS"
        worksheet[f"B{start_row}"].font = Font(size=14, bold=True, color=self.styles.palette.header_dark)
        header_row = start_row + 2
        headers = ["Month", "Budget Savings", "Actual Savings", "Actual Cumulative", "% of Actual Target"]
        for index, header in enumerate(headers, start=2):
            cell = worksheet.cell(row=header_row, column=index)
            cell.value = header
            cell.font = self.styles.header_font()
            cell.fill = self.styles.header_fill()
            cell.border = self.styles.thin_border
            cell.alignment = self.styles.left_alignment() if index == 2 else self.styles.right_alignment()

        row = header_row + 1
        for index, month in enumerate(self.config.months):
            worksheet[f"B{row}"] = month
            worksheet[f"B{row}"].font = Font(size=10, color=self.styles.palette.text_dark)
            worksheet[f"B{row}"].border = self.styles.thin_border
            worksheet[f"B{row}"].alignment = self.styles.left_alignment()
            worksheet[f"C{row}"] = f"='Monthly Entry'!{self.month_budget_letter(index)}{self.rows.savings_transfer}"
            worksheet[f"D{row}"] = f"='Monthly Entry'!{self.month_actual_letter(index)}{self.rows.savings_transfer}"
            worksheet[f"E{row}"] = f"='Monthly Entry'!{self.month_actual_letter(index)}{self.rows.running_savings}"
            worksheet[f"F{row}"] = f"=IF($D$9=0,0,E{row}/$D$9)"
            for col in ["C", "D", "E"]:
                worksheet[f"{col}{row}"].font = self.styles.calc_font()
                worksheet[f"{col}{row}"].border = self.styles.thin_border
                worksheet[f"{col}{row}"].alignment = self.styles.right_alignment()
                worksheet[f"{col}{row}"].number_format = self.config.currency_format
            worksheet[f"F{row}"].font = self.styles.calc_font()
            worksheet[f"F{row}"].border = self.styles.thin_border
            worksheet[f"F{row}"].alignment = self.styles.right_alignment()
            worksheet[f"F{row}"].number_format = self.config.percent_format
            row += 1
