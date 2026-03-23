"""Summary dashboard worksheet builder."""

from openpyxl.styles import Font

from .base import BaseSheetBuilder


class SummaryDashboardSheetBuilder(BaseSheetBuilder):
    """Build the Summary Dashboard worksheet."""

    def build(self, workbook):
        worksheet = workbook.create_sheet("Summary Dashboard")
        worksheet.sheet_view.showGridLines = False

        worksheet.column_dimensions["A"].width = 3
        worksheet.column_dimensions["B"].width = 30
        worksheet.column_dimensions["C"].width = 18
        worksheet.column_dimensions["D"].width = 18
        worksheet.column_dimensions["E"].width = 18

        worksheet.merge_cells("B2:E2")
        worksheet["B2"] = "MONTHLY SUMMARY DASHBOARD"
        worksheet["B2"].font = Font(size=18, bold=True, color=self.styles.palette.header_dark)
        worksheet["B2"].alignment = self.styles.left_alignment()
        worksheet.row_dimensions[2].height = 30

        worksheet.merge_cells("B3:E3")
        worksheet["B3"] = (
            "Overview of your financial performance - updates automatically from Monthly Entry"
        )
        worksheet["B3"].font = Font(size=10, italic=True, color="666666")

        worksheet["B5"] = "KEY METRICS (Year-to-Date)"
        worksheet["B5"].font = Font(size=14, bold=True, color=self.styles.palette.header_dark)
        worksheet.row_dimensions[5].height = 25

        metrics = [
            ("Total Income (Net + Bonus)", "='Monthly Entry'!O45", self.config.currency_format),
            ("Total Expenses", "='Monthly Entry'!O44", self.config.currency_format),
            ("Net Savings (Transferred)", "='Monthly Entry'!O32", self.config.currency_format),
            ("Total Bonuses Received", "='Monthly Entry'!O36", self.config.currency_format),
            ("Unexpected Expenses", "='Monthly Entry'!O39", self.config.currency_format),
            ("Year-End Surplus/(Deficit)", "='Monthly Entry'!O46", self.config.currency_format),
            ("Average Savings Rate", "='Monthly Entry'!O47", self.config.percent_format),
            ("Year-End Savings Balance", "='Monthly Entry'!O48", self.config.currency_format),
        ]

        row = 7
        for label, formula, number_format in metrics:
            worksheet[f"B{row}"] = label
            worksheet[f"B{row}"].font = Font(size=11, color=self.styles.palette.text_dark)
            worksheet[f"B{row}"].alignment = self.styles.left_alignment()
            worksheet[f"B{row}"].border = self.styles.thin_border

            worksheet.merge_cells(f"C{row}:D{row}")
            worksheet[f"C{row}"] = formula
            worksheet[f"C{row}"].font = Font(size=12, bold=True, color=self.styles.palette.header_light)
            worksheet[f"C{row}"].alignment = self.styles.right_alignment()
            worksheet[f"C{row}"].number_format = number_format
            worksheet[f"C{row}"].border = self.styles.thin_border

            if row % 2 == 0:
                worksheet[f"B{row}"].fill = self.styles.alt_row_fill()
                worksheet[f"C{row}"].fill = self.styles.alt_row_fill()
            row += 1

        row += 2
        worksheet[f"B{row}"] = "EXPENSE BREAKDOWN BY CATEGORY"
        worksheet[f"B{row}"].font = Font(size=14, bold=True, color=self.styles.palette.header_dark)
        row += 2

        worksheet[f"B{row}"] = "Category"
        worksheet[f"C{row}"] = "Year Total"
        worksheet[f"D{row}"] = "% of Total"
        worksheet[f"B{row}"].font = self.styles.header_font()
        worksheet[f"C{row}"].font = self.styles.header_font()
        worksheet[f"D{row}"].font = self.styles.header_font()
        worksheet[f"B{row}"].fill = self.styles.header_fill()
        worksheet[f"C{row}"].fill = self.styles.header_fill()
        worksheet[f"D{row}"].fill = self.styles.header_fill()
        worksheet[f"B{row}"].border = self.styles.thin_border
        worksheet[f"C{row}"].border = self.styles.thin_border
        worksheet[f"D{row}"].border = self.styles.thin_border
        worksheet[f"B{row}"].alignment = self.styles.left_alignment()
        worksheet[f"C{row}"].alignment = self.styles.right_alignment()
        worksheet[f"D{row}"].alignment = self.styles.right_alignment()
        worksheet.row_dimensions[row].height = 22

        expense_categories = [
            ("Data", "='Monthly Entry'!O11"),
            ("Transport", "='Monthly Entry'!O12"),
            ("Subscriptions", "='Monthly Entry'!O13"),
            ("TFG Debit", "='Monthly Entry'!O14"),
            ("Family Support", "='Monthly Entry'!O15"),
            ("Tithe", "='Monthly Entry'!O16"),
            ("Groceries", "='Monthly Entry'!O19"),
            ("Eating Out", "='Monthly Entry'!O20"),
            ("Lunch", "='Monthly Entry'!O21"),
            ("Haircuts", "='Monthly Entry'!O22"),
            ("Clothing", "='Monthly Entry'!O23"),
            ("Random Spending", "='Monthly Entry'!O24"),
            ("Rent Paid", "='Monthly Entry'!O29"),
            ("Unexpected Expenses", "='Monthly Entry'!O39"),
        ]

        row += 1
        for category, formula in expense_categories:
            worksheet[f"B{row}"] = category
            worksheet[f"B{row}"].font = Font(size=10, color=self.styles.palette.text_dark)
            worksheet[f"B{row}"].border = self.styles.thin_border
            worksheet[f"B{row}"].alignment = self.styles.left_alignment()

            worksheet[f"C{row}"] = formula
            worksheet[f"C{row}"].font = self.styles.calc_font()
            worksheet[f"C{row}"].border = self.styles.thin_border
            worksheet[f"C{row}"].alignment = self.styles.right_alignment()
            worksheet[f"C{row}"].number_format = self.config.currency_format

            worksheet[f"D{row}"] = f"=IF($C${row + 14}=0,0,C{row}/$C${row + 14})"
            worksheet[f"D{row}"].font = self.styles.calc_font()
            worksheet[f"D{row}"].border = self.styles.thin_border
            worksheet[f"D{row}"].alignment = self.styles.right_alignment()
            worksheet[f"D{row}"].number_format = self.config.percent_format

            if row % 2 == 0:
                worksheet[f"B{row}"].fill = self.styles.alt_row_fill()
                worksheet[f"C{row}"].fill = self.styles.alt_row_fill()
                worksheet[f"D{row}"].fill = self.styles.alt_row_fill()
            row += 1

        worksheet[f"B{row}"] = "TOTAL EXPENSES"
        worksheet[f"B{row}"].font = Font(bold=True, color=self.styles.palette.white)
        worksheet[f"B{row}"].fill = self.styles.header_fill()
        worksheet[f"B{row}"].border = self.styles.thin_border
        worksheet[f"B{row}"].alignment = self.styles.left_alignment()

        worksheet[f"C{row}"] = "='Monthly Entry'!O44"
        worksheet[f"C{row}"].font = Font(bold=True, color=self.styles.palette.white)
        worksheet[f"C{row}"].fill = self.styles.header_fill()
        worksheet[f"C{row}"].border = self.styles.thin_border
        worksheet[f"C{row}"].alignment = self.styles.right_alignment()
        worksheet[f"C{row}"].number_format = self.config.currency_format

        worksheet[f"D{row}"] = "=IF(C33=0,0,C33/C33)"
        worksheet[f"D{row}"].font = Font(bold=True, color=self.styles.palette.white)
        worksheet[f"D{row}"].fill = self.styles.header_fill()
        worksheet[f"D{row}"].border = self.styles.thin_border
        worksheet[f"D{row}"].alignment = self.styles.right_alignment()
        worksheet[f"D{row}"].number_format = self.config.percent_format

        return worksheet
