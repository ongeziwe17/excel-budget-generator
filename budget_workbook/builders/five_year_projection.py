"""5-year projection worksheet builder."""

from openpyxl.styles import Font

from .base import BaseSheetBuilder


class FiveYearProjectionSheetBuilder(BaseSheetBuilder):
    """Build the 5-Year Projection worksheet."""

    def build(self, workbook):
        worksheet = workbook.create_sheet("5-Year Projection")
        worksheet.sheet_view.showGridLines = False

        worksheet.column_dimensions["A"].width = 3
        worksheet.column_dimensions["B"].width = 30
        worksheet.column_dimensions["C"].width = 18
        worksheet.column_dimensions["D"].width = 18
        worksheet.column_dimensions["E"].width = 18
        worksheet.column_dimensions["F"].width = 18

        worksheet.merge_cells("B2:F2")
        worksheet["B2"] = "5-YEAR SAVINGS PROJECTION"
        worksheet["B2"].font = Font(size=18, bold=True, color=self.styles.palette.header_dark)
        worksheet["B2"].alignment = self.styles.left_alignment()
        worksheet.row_dimensions[2].height = 30

        worksheet.merge_cells("B3:F3")
        worksheet["B3"] = "Based on actual savings entered | 8% annual growth assumption"
        worksheet["B3"].font = Font(size=10, italic=True, color="666666")

        self._create_assumptions(worksheet)
        self._create_projection_table(worksheet)
        self._create_summary(worksheet)
        self._create_notes(worksheet)
        return worksheet

    def _create_assumptions(self, worksheet) -> None:
        worksheet["B5"] = "PROJECTION ASSUMPTIONS"
        worksheet["B5"].font = Font(size=12, bold=True, color=self.styles.palette.header_dark)

        worksheet["B6"] = "Annual Growth Rate"
        worksheet["B6"].font = Font(size=10, color=self.styles.palette.text_dark)
        worksheet["B6"].border = self.styles.thin_border

        worksheet["C6"] = self.config.annual_growth_rate
        worksheet["C6"].font = self.styles.input_font()
        worksheet["C6"].border = self.styles.thin_border
        worksheet["C6"].alignment = self.styles.right_alignment()
        worksheet["C6"].number_format = self.config.percent_format

        worksheet["B7"] = "Monthly Savings (avg from Year 1)"
        worksheet["B7"].font = Font(size=10, color=self.styles.palette.text_dark)
        worksheet["B7"].border = self.styles.thin_border
        worksheet["B7"].fill = self.styles.alt_row_fill()

        worksheet["C7"] = "='Monthly Entry'!O32/12"
        worksheet["C7"].font = self.styles.calc_font()
        worksheet["C7"].border = self.styles.thin_border
        worksheet["C7"].alignment = self.styles.right_alignment()
        worksheet["C7"].number_format = self.config.currency_format
        worksheet["C7"].fill = self.styles.alt_row_fill()

        worksheet["B8"] = "Starting Balance (Year-End)"
        worksheet["B8"].font = Font(size=10, color=self.styles.palette.text_dark)
        worksheet["B8"].border = self.styles.thin_border

        worksheet["C8"] = "='Monthly Entry'!O48"
        worksheet["C8"].font = self.styles.calc_font()
        worksheet["C8"].border = self.styles.thin_border
        worksheet["C8"].alignment = self.styles.right_alignment()
        worksheet["C8"].number_format = self.config.currency_format

    def _create_projection_table(self, worksheet) -> None:
        worksheet["B11"] = "YEAR-BY-YEAR PROJECTION"
        worksheet["B11"].font = Font(size=14, bold=True, color=self.styles.palette.header_dark)
        worksheet.row_dimensions[11].height = 25

        headers = [
            ("B13", "Year", self.styles.left_alignment()),
            ("C13", "Starting Balance", self.styles.right_alignment()),
            ("D13", "Annual Contribution", self.styles.right_alignment()),
            ("E13", "Growth (8%)", self.styles.right_alignment()),
            ("F13", "Ending Balance", self.styles.right_alignment()),
        ]
        for cell_ref, label, alignment in headers:
            worksheet[cell_ref] = label
            worksheet[cell_ref].font = self.styles.header_font()
            worksheet[cell_ref].fill = self.styles.header_fill()
            worksheet[cell_ref].border = self.styles.thin_border
            worksheet[cell_ref].alignment = alignment
        worksheet.row_dimensions[13].height = 22

        row = 14
        worksheet[f"B{row}"] = "Year 1 (Current)"
        worksheet[f"B{row}"].font = Font(size=10, bold=True, color=self.styles.palette.text_dark)
        worksheet[f"B{row}"].border = self.styles.thin_border

        worksheet[f"C{row}"] = "='Monthly Entry'!C33"
        worksheet[f"C{row}"].font = self.styles.calc_font()
        worksheet[f"C{row}"].border = self.styles.thin_border
        worksheet[f"C{row}"].alignment = self.styles.right_alignment()
        worksheet[f"C{row}"].number_format = self.config.currency_format

        worksheet[f"D{row}"] = "='Monthly Entry'!O32"
        worksheet[f"D{row}"].font = self.styles.calc_font()
        worksheet[f"D{row}"].border = self.styles.thin_border
        worksheet[f"D{row}"].alignment = self.styles.right_alignment()
        worksheet[f"D{row}"].number_format = self.config.currency_format

        worksheet[f"E{row}"] = f"=(C{row}+D{row}/2)*$C$6"
        worksheet[f"E{row}"].font = self.styles.calc_font()
        worksheet[f"E{row}"].border = self.styles.thin_border
        worksheet[f"E{row}"].alignment = self.styles.right_alignment()
        worksheet[f"E{row}"].number_format = self.config.currency_format

        worksheet[f"F{row}"] = "='Monthly Entry'!O48"
        worksheet[f"F{row}"].font = Font(bold=True, size=11, color=self.styles.palette.positive_green)
        worksheet[f"F{row}"].border = self.styles.thin_border
        worksheet[f"F{row}"].alignment = self.styles.right_alignment()
        worksheet[f"F{row}"].number_format = self.config.currency_format

        for year in range(2, 6):
            row += 1
            worksheet[f"B{row}"] = f"Year {year} (Projected)"
            worksheet[f"B{row}"].font = Font(size=10, color=self.styles.palette.text_dark)
            worksheet[f"B{row}"].border = self.styles.thin_border

            worksheet[f"C{row}"] = f"=F{row - 1}"
            worksheet[f"C{row}"].font = self.styles.calc_font()
            worksheet[f"C{row}"].border = self.styles.thin_border
            worksheet[f"C{row}"].alignment = self.styles.right_alignment()
            worksheet[f"C{row}"].number_format = self.config.currency_format

            worksheet[f"D{row}"] = "=$C$7*12"
            worksheet[f"D{row}"].font = self.styles.calc_font()
            worksheet[f"D{row}"].border = self.styles.thin_border
            worksheet[f"D{row}"].alignment = self.styles.right_alignment()
            worksheet[f"D{row}"].number_format = self.config.currency_format

            worksheet[f"E{row}"] = f"=(C{row}+D{row}/2)*$C$6"
            worksheet[f"E{row}"].font = self.styles.calc_font()
            worksheet[f"E{row}"].border = self.styles.thin_border
            worksheet[f"E{row}"].alignment = self.styles.right_alignment()
            worksheet[f"E{row}"].number_format = self.config.currency_format

            worksheet[f"F{row}"] = f"=C{row}+D{row}+E{row}"
            worksheet[f"F{row}"].font = Font(bold=True, size=11, color=self.styles.palette.header_light)
            worksheet[f"F{row}"].border = self.styles.thin_border
            worksheet[f"F{row}"].alignment = self.styles.right_alignment()
            worksheet[f"F{row}"].number_format = self.config.currency_format

            if row % 2 == 0:
                for col in ["B", "C", "D", "E", "F"]:
                    worksheet[f"{col}{row}"].fill = self.styles.alt_row_fill()

        row += 1
        worksheet[f"B{row}"] = "5-Year Total"
        worksheet[f"B{row}"].font = Font(bold=True, color=self.styles.palette.white)
        worksheet[f"B{row}"].fill = self.styles.header_fill()
        worksheet[f"B{row}"].border = self.styles.thin_border

        worksheet[f"C{row}"].fill = self.styles.header_fill()
        worksheet[f"C{row}"].border = self.styles.thin_border

        worksheet[f"D{row}"] = "=SUM(D14:D18)"
        worksheet[f"D{row}"].font = Font(bold=True, color=self.styles.palette.white)
        worksheet[f"D{row}"].fill = self.styles.header_fill()
        worksheet[f"D{row}"].border = self.styles.thin_border
        worksheet[f"D{row}"].alignment = self.styles.right_alignment()
        worksheet[f"D{row}"].number_format = self.config.currency_format

        worksheet[f"E{row}"] = "=SUM(E14:E18)"
        worksheet[f"E{row}"].font = Font(bold=True, color=self.styles.palette.white)
        worksheet[f"E{row}"].fill = self.styles.header_fill()
        worksheet[f"E{row}"].border = self.styles.thin_border
        worksheet[f"E{row}"].alignment = self.styles.right_alignment()
        worksheet[f"E{row}"].number_format = self.config.currency_format

        worksheet[f"F{row}"] = "=F18"
        worksheet[f"F{row}"].font = Font(bold=True, color=self.styles.palette.white)
        worksheet[f"F{row}"].fill = self.styles.header_fill()
        worksheet[f"F{row}"].border = self.styles.thin_border
        worksheet[f"F{row}"].alignment = self.styles.right_alignment()
        worksheet[f"F{row}"].number_format = self.config.currency_format

    def _create_summary(self, worksheet) -> None:
        row = 22
        worksheet[f"B{row}"] = "PROJECTION SUMMARY"
        worksheet[f"B{row}"].font = Font(size=14, bold=True, color=self.styles.palette.header_dark)

        metrics = [
            ("Starting Balance (Today)", "=C14", self.config.currency_format),
            ("Projected Balance (Year 5)", "=F18", self.config.currency_format),
            ("Total Contributions (5 Years)", "=SUM(D14:D18)", self.config.currency_format),
            ("Total Growth/Earnings (5 Years)", "=SUM(E14:E18)", self.config.currency_format),
            ("Growth Multiple", "=IF(C14=0,0,F18/C14)", '0.00"x"'),
        ]

        row += 2
        for label, formula, number_format in metrics:
            worksheet[f"B{row}"] = label
            worksheet[f"B{row}"].font = Font(size=10, color=self.styles.palette.text_dark)
            worksheet[f"B{row}"].border = self.styles.thin_border

            worksheet[f"C{row}"] = formula
            worksheet[f"C{row}"].font = Font(size=11, bold=True, color=self.styles.palette.header_light)
            worksheet[f"C{row}"].border = self.styles.thin_border
            worksheet[f"C{row}"].alignment = self.styles.right_alignment()
            worksheet[f"C{row}"].number_format = number_format

            if row % 2 == 0:
                worksheet[f"B{row}"].fill = self.styles.alt_row_fill()
                worksheet[f"C{row}"].fill = self.styles.alt_row_fill()
            row += 1

    def _create_notes(self, worksheet) -> None:
        row = 29
        worksheet[f"B{row}"] = "IMPORTANT NOTES"
        worksheet[f"B{row}"].font = Font(size=12, bold=True, color=self.styles.palette.header_dark)

        notes = [
            "• This projection assumes consistent monthly savings at your Year 1 average",
            "• 8% annual return is an estimate - actual returns will vary",
            "• Update your Monthly Entry sheet regularly for accurate projections",
            "• Consider increasing savings rate as income grows",
            "• This is a simplified projection - consult a financial advisor for detailed planning",
        ]
        row += 2
        for note in notes:
            worksheet[f"B{row}"] = note
            worksheet[f"B{row}"].font = Font(size=10, color="555555")
            worksheet.merge_cells(f"B{row}:F{row}")
            row += 1
