"""Cover worksheet builder."""

from openpyxl.styles import Alignment, Font, PatternFill

from .base import BaseSheetBuilder


class CoverSheetBuilder(BaseSheetBuilder):
    """Build the workbook cover sheet."""

    def build(self, workbook):
        worksheet = workbook.active
        worksheet.title = "Cover"
        worksheet.sheet_view.showGridLines = False

        worksheet.column_dimensions["A"].width = 3
        worksheet.column_dimensions["B"].width = 35
        worksheet.column_dimensions["C"].width = 52
        worksheet.column_dimensions["D"].width = 3

        worksheet.merge_cells("B2:C2")
        worksheet["B2"] = "PERSONAL BUDGET WORKBOOK"
        worksheet["B2"].font = Font(size=24, bold=True, color=self.styles.palette.header_dark)
        worksheet["B2"].alignment = self.styles.center_alignment()
        worksheet.row_dimensions[2].height = 40

        worksheet.merge_cells("B3:C3")
        worksheet["B3"] = "Budget vs Actual | Dynamic Detail Tables | Manual Monthly Entry"
        worksheet["B3"].font = Font(size=12, italic=True, color="666666")
        worksheet["B3"].alignment = self.styles.center_alignment()

        worksheet.merge_cells("B5:C5")
        worksheet["B5"] = (
            "Track planned vs actual cash flow, monitor grocery detail, and keep dashboards formula-driven."
        )
        worksheet["B5"].font = Font(size=11, color="444444")
        worksheet["B5"].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        worksheet["B7"] = "KEY FEATURES"
        worksheet["B7"].font = Font(size=14, bold=True, color=self.styles.palette.header_dark)
        worksheet.row_dimensions[7].height = 25

        features = [
            ("✓ Budget vs Actual", "Every month tracks Budget, Actual, and Variance columns"),
            ("✓ Dynamic Grocery Detail", "Groceries roll up from a structured detail table with starter rows"),
            ("✓ Expandable Variable Expenses", "Add custom variable rows inside Excel tables without changing totals"),
            ("✓ Automatic Dashboards", "Charts and summary metrics update from formula-driven rollups"),
            ("✓ Savings & Projection", "Emergency fund and projection sheets use the updated budget model"),
        ]

        row = 9
        for feature, description in features:
            worksheet[f"B{row}"] = feature
            worksheet[f"B{row}"].font = Font(size=11, bold=True, color=self.styles.palette.header_light)
            worksheet[f"C{row}"] = description
            worksheet[f"C{row}"].font = Font(size=10, color="555555")
            row += 1

        worksheet[f"B{row + 1}"] = "WORKBOOK STRUCTURE"
        worksheet[f"B{row + 1}"].font = Font(size=14, bold=True, color=self.styles.palette.header_dark)

        sheet_index = [
            ("Monthly Entry", "Manual entry plus grocery/custom detail tables and automatic variance formulas"),
            ("Summary Dashboard", "Budget vs Actual metrics, monthly comparison table, and charts"),
            ("Trend Analysis", "Variance trends and month-over-month actual changes"),
            ("Bonus Tracker", "Monthly and quarterly bonus rollups from Monthly Entry"),
            ("Emergency Fund", "Budget vs Actual emergency fund readiness"),
            ("5-Year Projection", "Projection table comparing budget and actual savings paths"),
        ]

        row += 3
        worksheet[f"B{row}"] = "Sheet Name"
        worksheet[f"B{row}"].font = Font(bold=True, color=self.styles.palette.header_dark)
        worksheet[f"B{row}"].fill = PatternFill(start_color="D6E3F8", end_color="D6E3F8", fill_type="solid")
        worksheet[f"C{row}"] = "Purpose"
        worksheet[f"C{row}"].font = Font(bold=True, color=self.styles.palette.header_dark)
        worksheet[f"C{row}"].fill = PatternFill(start_color="D6E3F8", end_color="D6E3F8", fill_type="solid")
        row += 1

        for sheet_name, purpose in sheet_index:
            worksheet[f"B{row}"] = sheet_name
            worksheet[f"B{row}"].font = Font(size=10, bold=True)
            worksheet[f"C{row}"] = purpose
            worksheet[f"C{row}"].font = Font(size=10, color="555555")
            if row % 2 == 0:
                worksheet[f"B{row}"].fill = self.styles.alt_row_fill()
                worksheet[f"C{row}"].fill = self.styles.alt_row_fill()
            row += 1

        worksheet[f"B{row + 2}"] = "HOW TO USE THIS WORKBOOK"
        worksheet[f"B{row + 2}"].font = Font(size=14, bold=True, color=self.styles.palette.header_dark)
        instructions = [
            "1. WHITE/BLUE entry cells are manual inputs for Budget and Actual amounts only",
            "2. Variance cells calculate automatically as Actual - Budget",
            "3. Add grocery rows inside GroceryDetailTable and custom rows inside VariableExpenseTable",
            "4. Extend months safely by updating WorkbookConfig.months and regenerating the workbook",
        ]
        row += 4
        for instruction in instructions:
            worksheet[f"B{row}"] = instruction
            worksheet[f"B{row}"].font = Font(size=10, color="444444")
            row += 1

        worksheet[f"B{row + 2}"] = "COLOR LEGEND"
        worksheet[f"B{row + 2}"].font = Font(size=12, bold=True, color=self.styles.palette.header_dark)
        row += 4
        worksheet[f"B{row}"] = "Blue Text"
        worksheet[f"B{row}"].font = Font(color=self.styles.palette.input_blue, bold=True)
        worksheet[f"C{row}"] = "Manual Budget/Actual input cells"
        worksheet[f"C{row}"].font = Font(size=10)
        worksheet[f"B{row + 1}"] = "Black Text"
        worksheet[f"B{row + 1}"].font = Font(color=self.styles.palette.text_dark, bold=True)
        worksheet[f"C{row + 1}"] = "Formula-driven cells and rollups"
        worksheet[f"C{row + 1}"].font = Font(size=10)
        worksheet[f"B{row + 2}"] = "Green Highlight"
        worksheet[f"B{row + 2}"].fill = PatternFill(start_color=self.styles.palette.green_fill, end_color=self.styles.palette.green_fill, fill_type="solid")
        worksheet[f"C{row + 2}"] = "Variance below budget or favorable movement"
        worksheet[f"C{row + 2}"].font = Font(size=10)
        worksheet[f"B{row + 3}"] = "Red Highlight"
        worksheet[f"B{row + 3}"].fill = PatternFill(start_color=self.styles.palette.red_fill, end_color=self.styles.palette.red_fill, fill_type="solid")
        worksheet[f"C{row + 3}"] = "Variance above budget or unfavorable movement"
        worksheet[f"C{row + 3}"].font = Font(size=10)
        return worksheet
