"""Summary dashboard worksheet builder."""

from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.styles import Font

from .base import BaseSheetBuilder


class SummaryDashboardSheetBuilder(BaseSheetBuilder):
    """Build the Summary Dashboard worksheet with budget vs actual visibility."""

    def build(self, workbook):
        worksheet = workbook.create_sheet("Summary Dashboard")
        worksheet.sheet_view.showGridLines = False
        self._configure_columns(worksheet)
        self._create_header(worksheet)
        self._create_key_metrics(worksheet)
        expense_breakdown_end = self._create_expense_breakdown(worksheet, start_row=18)
        monthly_table_end = self._create_monthly_comparison(worksheet, start_row=expense_breakdown_end + 3)
        self._add_charts(worksheet, expense_breakdown_end, monthly_table_end)
        return worksheet

    def _configure_columns(self, worksheet) -> None:
        worksheet.column_dimensions["A"].width = 3
        worksheet.column_dimensions["B"].width = 30
        worksheet.column_dimensions["C"].width = 18
        worksheet.column_dimensions["D"].width = 18
        worksheet.column_dimensions["E"].width = 18
        worksheet.column_dimensions["F"].width = 18
        worksheet.column_dimensions["G"].width = 18

    def _create_header(self, worksheet) -> None:
        worksheet.merge_cells("B2:F2")
        worksheet["B2"] = "MONTHLY SUMMARY DASHBOARD"
        worksheet["B2"].font = Font(size=18, bold=True, color=self.styles.palette.header_dark)
        worksheet["B2"].alignment = self.styles.left_alignment()
        worksheet.row_dimensions[2].height = 30

        worksheet.merge_cells("B3:F3")
        worksheet["B3"] = "Budget vs Actual overview driven entirely by formulas from Monthly Entry."
        worksheet["B3"].font = Font(size=10, italic=True, color="666666")

    def _create_key_metrics(self, worksheet) -> None:
        worksheet["B5"] = "KEY METRICS (YEAR TOTALS)"
        worksheet["B5"].font = Font(size=14, bold=True, color=self.styles.palette.header_dark)
        worksheet.row_dimensions[5].height = 25

        for cell_ref, label in [("B7", "Metric"), ("C7", "Budget"), ("D7", "Actual"), ("E7", "Variance")]:
            worksheet[cell_ref] = label
            worksheet[cell_ref].font = self.styles.header_font()
            worksheet[cell_ref].fill = self.styles.header_fill()
            worksheet[cell_ref].border = self.styles.thin_border
            worksheet[cell_ref].alignment = self.styles.left_alignment() if cell_ref == "B7" else self.styles.right_alignment()

        metrics = [
            ("Total Income", self.rows.total_income),
            ("Total Expenses", self.rows.total_expenses),
            ("Transfer to Savings", self.rows.savings_transfer),
            ("Bonuses Received", self.rows.bonus),
            ("Unexpected Expenses", self.rows.unexpected_expenses),
            ("Surplus / (Deficit)", self.rows.surplus_deficit),
            ("Savings Rate", self.rows.savings_rate),
            ("Running Savings Balance", self.rows.running_savings),
        ]
        row = 8
        for label, source_row in metrics:
            worksheet[f"B{row}"] = label
            worksheet[f"B{row}"].font = Font(size=11, color=self.styles.palette.text_dark)
            worksheet[f"B{row}"].border = self.styles.thin_border
            worksheet[f"B{row}"].alignment = self.styles.left_alignment()

            worksheet[f"C{row}"] = f"='Monthly Entry'!{self.year_budget_letter()}{source_row}"
            worksheet[f"D{row}"] = f"='Monthly Entry'!{self.year_actual_letter()}{source_row}"
            worksheet[f"E{row}"] = f"='Monthly Entry'!{self.year_variance_letter()}{source_row}"

            for col in ["C", "D", "E"]:
                worksheet[f"{col}{row}"].font = Font(size=11, bold=col != "E", color=self.styles.palette.header_light if col != "E" else self.styles.palette.text_dark)
                worksheet[f"{col}{row}"].border = self.styles.thin_border
                worksheet[f"{col}{row}"].alignment = self.styles.right_alignment()
                worksheet[f"{col}{row}"].number_format = self.config.percent_format if source_row == self.rows.savings_rate else self.config.currency_format

            if row % 2 == 0:
                for col in ["B", "C", "D", "E"]:
                    worksheet[f"{col}{row}"].fill = self.styles.alt_row_fill()
            row += 1

        self.add_variance_conditional_formatting(worksheet, 8, row - 1, columns=[5])

    def _create_expense_breakdown(self, worksheet, start_row: int) -> int:
        worksheet[f"B{start_row}"] = "EXPENSE BREAKDOWN"
        worksheet[f"B{start_row}"].font = Font(size=14, bold=True, color=self.styles.palette.header_dark)

        header_row = start_row + 2
        headers = ["Category", "Budget", "Actual", "Variance", "% of Actual"]
        for offset, header in enumerate(headers):
            cell = worksheet.cell(row=header_row, column=2 + offset)
            cell.value = header
            cell.font = self.styles.header_font()
            cell.fill = self.styles.header_fill()
            cell.border = self.styles.thin_border
            cell.alignment = self.styles.left_alignment() if offset == 0 else self.styles.right_alignment()

        categories = [
            ("Data", self.rows.data),
            ("Transport", self.rows.transport),
            ("Subscriptions", self.rows.subscriptions),
            ("TFG Debit", self.rows.tfg_debit),
            ("Family Support", self.rows.family_support),
            ("Tithe", self.rows.tithe),
            ("Rent / Housing", self.rows.rent_housing),
            ("Groceries", self.rows.groceries),
            ("Eating Out", self.rows.eating_out),
            ("Lunch", self.rows.lunch),
            ("Haircuts", self.rows.haircuts),
            ("Clothing", self.rows.clothing),
            ("Random Spending", self.rows.random_spending),
            ("Additional Variable Items", self.rows.custom_variable_total),
            ("Unexpected Expenses", self.rows.unexpected_expenses),
        ]

        row = header_row + 1
        total_row = row + len(categories)
        for label, source_row in categories:
            worksheet[f"B{row}"] = label
            worksheet[f"B{row}"].font = Font(size=10, color=self.styles.palette.text_dark)
            worksheet[f"B{row}"].border = self.styles.thin_border
            worksheet[f"B{row}"].alignment = self.styles.left_alignment()

            worksheet[f"C{row}"] = f"='Monthly Entry'!{self.year_budget_letter()}{source_row}"
            worksheet[f"D{row}"] = f"='Monthly Entry'!{self.year_actual_letter()}{source_row}"
            worksheet[f"E{row}"] = f"='Monthly Entry'!{self.year_variance_letter()}{source_row}"
            worksheet[f"F{row}"] = f"=IF($D${total_row}=0,0,D{row}/$D${total_row})"
            for col in ["C", "D", "E", "F"]:
                worksheet[f"{col}{row}"].font = self.styles.calc_font()
                worksheet[f"{col}{row}"].border = self.styles.thin_border
                worksheet[f"{col}{row}"].alignment = self.styles.right_alignment()
                worksheet[f"{col}{row}"].number_format = self.config.percent_format if col == "F" else self.config.currency_format

            if row % 2 == 0:
                for col in ["B", "C", "D", "E", "F"]:
                    worksheet[f"{col}{row}"].fill = self.styles.alt_row_fill()
            row += 1

        worksheet[f"B{total_row}"] = "TOTAL EXPENSES"
        worksheet[f"C{total_row}"] = f"='Monthly Entry'!{self.year_budget_letter()}{self.rows.total_expenses}"
        worksheet[f"D{total_row}"] = f"='Monthly Entry'!{self.year_actual_letter()}{self.rows.total_expenses}"
        worksheet[f"E{total_row}"] = f"='Monthly Entry'!{self.year_variance_letter()}{self.rows.total_expenses}"
        worksheet[f"F{total_row}"] = "=IF(D{0}=0,0,D{0}/D{0})".format(total_row)
        for col in ["B", "C", "D", "E", "F"]:
            worksheet[f"{col}{total_row}"].font = Font(bold=True, color=self.styles.palette.white)
            worksheet[f"{col}{total_row}"].fill = self.styles.header_fill()
            worksheet[f"{col}{total_row}"].border = self.styles.thin_border
            worksheet[f"{col}{total_row}"].alignment = self.styles.left_alignment() if col == "B" else self.styles.right_alignment()
            worksheet[f"{col}{total_row}"].number_format = self.config.percent_format if col == "F" else self.config.currency_format

        self.add_variance_conditional_formatting(worksheet, header_row + 1, total_row, columns=[5])
        return total_row

    def _create_monthly_comparison(self, worksheet, start_row: int) -> int:
        worksheet[f"B{start_row}"] = "MONTHLY BUDGET VS ACTUAL COMPARISON"
        worksheet[f"B{start_row}"].font = Font(size=14, bold=True, color=self.styles.palette.header_dark)

        header_row = start_row + 2
        for offset, header in enumerate(["Month", "Expense Budget", "Expense Actual", "Expense Variance", "Income Actual", "Savings Rate Actual"]):
            cell = worksheet.cell(row=header_row, column=2 + offset)
            cell.value = header
            cell.font = self.styles.header_font()
            cell.fill = self.styles.header_fill()
            cell.border = self.styles.thin_border
            cell.alignment = self.styles.left_alignment() if offset == 0 else self.styles.right_alignment()

        row = header_row + 1
        for index, month in enumerate(self.config.months):
            worksheet[f"B{row}"] = month
            worksheet[f"B{row}"].font = Font(size=10, color=self.styles.palette.text_dark)
            worksheet[f"B{row}"].border = self.styles.thin_border
            worksheet[f"B{row}"].alignment = self.styles.left_alignment()

            worksheet[f"C{row}"] = f"='Monthly Entry'!{self.month_budget_letter(index)}{self.rows.total_expenses}"
            worksheet[f"D{row}"] = f"='Monthly Entry'!{self.month_actual_letter(index)}{self.rows.total_expenses}"
            worksheet[f"E{row}"] = f"='Monthly Entry'!{self.month_variance_letter(index)}{self.rows.total_expenses}"
            worksheet[f"F{row}"] = f"='Monthly Entry'!{self.month_actual_letter(index)}{self.rows.total_income}"
            worksheet[f"G{row}"] = f"='Monthly Entry'!{self.month_actual_letter(index)}{self.rows.savings_rate}"
            for col in ["C", "D", "E", "F", "G"]:
                worksheet[f"{col}{row}"].font = self.styles.calc_font()
                worksheet[f"{col}{row}"].border = self.styles.thin_border
                worksheet[f"{col}{row}"].alignment = self.styles.right_alignment()
                worksheet[f"{col}{row}"].number_format = self.config.percent_format if col == "G" else self.config.currency_format
            row += 1

        self.add_variance_conditional_formatting(worksheet, header_row + 1, row - 1, columns=[5])
        return row - 1

    def _add_charts(self, worksheet, expense_breakdown_end: int, monthly_table_end: int) -> None:
        expense_chart = BarChart()
        expense_chart.type = "bar"
        expense_chart.style = 10
        expense_chart.title = "Expense Budget vs Actual by Category"
        expense_chart.y_axis.title = "Category"
        expense_chart.x_axis.title = "Amount"
        expense_data = Reference(worksheet, min_col=3, max_col=4, min_row=20, max_row=expense_breakdown_end - 1)
        expense_categories = Reference(worksheet, min_col=2, min_row=21, max_row=expense_breakdown_end - 1)
        expense_chart.add_data(expense_data, titles_from_data=True)
        expense_chart.set_categories(expense_categories)
        expense_chart.height = 9
        expense_chart.width = 18
        worksheet.add_chart(expense_chart, "H5")

        monthly_chart = LineChart()
        monthly_chart.style = 12
        monthly_chart.title = "Monthly Expense Budget vs Actual"
        monthly_chart.y_axis.title = "Amount"
        monthly_chart.x_axis.title = "Month"
        monthly_data = Reference(worksheet, min_col=3, max_col=4, min_row=expense_breakdown_end + 5, max_row=monthly_table_end)
        monthly_categories = Reference(worksheet, min_col=2, min_row=expense_breakdown_end + 6, max_row=monthly_table_end)
        monthly_chart.add_data(monthly_data, titles_from_data=True)
        monthly_chart.set_categories(monthly_categories)
        monthly_chart.height = 8
        monthly_chart.width = 18
        worksheet.add_chart(monthly_chart, f"H{expense_breakdown_end + 2}")
