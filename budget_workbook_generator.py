"""
Personal Budget Workbook Generator
===================================
A maintainable Python script to generate a dynamic Excel budgeting workbook
for South African Rand (ZAR) with manual entry fields and automatic calculations.

Author: AI Assistant
Date: 2026-03-21
"""

from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Border, Side, Alignment
from openpyxl.formatting.rule import CellIsRule
from openpyxl.utils import get_column_letter


# =============================================================================
# CONFIGURATION - Modify these values to customize your workbook
# =============================================================================

class Config:
    """Configuration constants for the workbook."""
    
    # Currency settings
    CURRENCY_FORMAT = 'R #,##0.00'
    PERCENT_FORMAT = '0.0%'
    
    # Financial profile (for reference/expected values only)
    MONTHLY_GROSS_SALARY = 25000
    MONTHLY_NET_INCOME = 20339
    
    # Expected bonus percentages (for Bonus Tracker reference)
    BONUS_Q2_PCT = 0.25  # 25% of gross
    BONUS_Q3_PCT = 0.50  # 50% of gross
    BONUS_Q4_PCT = 0.83  # 83% of gross
    
    # Default rent amount (can be changed in workbook)
    DEFAULT_RENT_AMOUNT = 8000
    
    # Projection assumptions
    ANNUAL_GROWTH_RATE = 0.08  # 8%
    
    # Output file path
    OUTPUT_PATH = "Personal_Budget_Workbook.xlsx"


# =============================================================================
# STYLE DEFINITIONS
# =============================================================================

class Styles:
    """Color palette and styling definitions."""
    
    # Colors
    HEADER_DARK = "1F4E79"
    HEADER_LIGHT = "2E75B6"
    ACCENT_BLUE = "5B9BD5"
    BG_LIGHT = "F5F5F5"
    TEXT_DARK = "000000"
    WHITE = "FFFFFF"
    POSITIVE_GREEN = "27AE60"
    NEGATIVE_RED = "E74C3C"
    INPUT_BLUE = "0066CC"
    
    # Conditional formatting colors
    GREEN_FILL = "D4EDDA"
    RED_FILL = "F8D7DA"
    
    # Border style
    THIN_BORDER = Border(
        left=Side(style='thin', color='D0D0D0'),
        right=Side(style='thin', color='D0D0D0'),
        top=Side(style='thin', color='D0D0D0'),
        bottom=Side(style='thin', color='D0D0D0')
    )
    
    # Font styles
    @classmethod
    def header_font(cls):
        return Font(color=cls.WHITE, bold=True, size=11)
    
    @classmethod
    def section_font(cls):
        return Font(color=cls.WHITE, bold=True, size=10)
    
    @classmethod
    def input_font(cls):
        return Font(color=cls.INPUT_BLUE, size=11)
    
    @classmethod
    def calc_font(cls):
        return Font(color=cls.TEXT_DARK, size=11)
    
    # Fill styles
    @classmethod
    def header_fill(cls):
        return PatternFill(start_color=cls.HEADER_DARK, end_color=cls.HEADER_DARK, fill_type="solid")
    
    @classmethod
    def section_fill(cls):
        return PatternFill(start_color=cls.HEADER_LIGHT, end_color=cls.HEADER_LIGHT, fill_type="solid")
    
    @classmethod
    def alt_row_fill(cls):
        return PatternFill(start_color=cls.BG_LIGHT, end_color=cls.BG_LIGHT, fill_type="solid")
    
    @classmethod
    def positive_fill(cls):
        return PatternFill(start_color=cls.POSITIVE_GREEN, end_color=cls.POSITIVE_GREEN, fill_type="solid")
    
    @classmethod
    def negative_fill(cls):
        return PatternFill(start_color=cls.NEGATIVE_RED, end_color=cls.NEGATIVE_RED, fill_type="solid")


# =============================================================================
# MONTHLY ENTRY ROW REFERENCE
# These are the row numbers in the Monthly Entry sheet for formula references
# =============================================================================

class MonthlyEntryRows:
    """Row numbers for key items in the Monthly Entry sheet."""
    NET_INCOME = 8
    DATA = 11
    TRANSPORT = 12
    SUBSCRIPTIONS = 13
    TFG_DEBIT = 14
    FAMILY_SUPPORT = 15
    TITHE = 16
    GROCERIES = 19
    EATING_OUT = 20
    LUNCH = 21
    HAIRCUTS = 22
    CLOTHING = 23
    RANDOM_SPENDING = 24
    RENT_TOGGLE = 27
    RENT_AMOUNT = 28
    RENT_PAID = 29
    SAVINGS_TRANSFER = 32
    STARTING_SAVINGS = 33
    BONUS = 36
    UNEXPECTED_EXPENSES = 39
    TOTAL_FIXED = 42
    TOTAL_VARIABLE = 43
    TOTAL_EXPENSES = 44
    TOTAL_INCOME = 45
    SURPLUS_DEFICIT = 46
    SAVINGS_RATE = 47
    RUNNING_SAVINGS = 48


# =============================================================================
# HELPER FUNCTIONS
# =============================================================================

def apply_border_range(ws, start_row, end_row, start_col, end_col):
    """Apply thin border to a range of cells."""
    for row in range(start_row, end_row + 1):
        for col in range(start_col, end_col + 1):
            ws.cell(row=row, column=col).border = Styles.THIN_BORDER


def create_section_header(ws, row, title, start_col=2, end_col=15):
    """Create a colored section header row."""
    ws.merge_cells(start_row=row, start_column=start_col, end_row=row, end_column=end_col)
    cell = ws.cell(row=row, column=start_col)
    cell.value = title
    cell.font = Styles.section_font()
    cell.fill = Styles.section_fill()
    cell.alignment = Alignment(horizontal='left', vertical='center')
    apply_border_range(ws, row, row, start_col, end_col)
    ws.row_dimensions[row].height = 22


def create_input_row(ws, row, label, indent=False, start_col=2, num_months=12):
    """
    Create a row for manual data entry.
    
    Args:
        ws: Worksheet object
        row: Row number
        label: Label text for the category
        indent: Whether to indent the label (for sub-categories)
        start_col: Starting column (default 2 = column B)
        num_months: Number of month columns (default 12)
    """
    # Label column
    label_cell = ws.cell(row=row, column=start_col)
    label_cell.value = f"  {label}" if indent else label
    label_cell.font = Font(size=10, bold=not indent, color=Styles.TEXT_DARK)
    label_cell.alignment = Alignment(horizontal='left', vertical='center')
    label_cell.border = Styles.THIN_BORDER
    
    # Month columns (for manual input - blue font)
    for i in range(num_months):
        col = start_col + 1 + i
        cell = ws.cell(row=row, column=col)
        cell.border = Styles.THIN_BORDER
        cell.alignment = Alignment(horizontal='right', vertical='center')
        cell.number_format = Config.CURRENCY_FORMAT
        cell.font = Styles.input_font()
    
    # Year Total column (formula - black font)
    total_col = start_col + 1 + num_months
    total_cell = ws.cell(row=row, column=total_col)
    total_cell.border = Styles.THIN_BORDER
    total_cell.alignment = Alignment(horizontal='right', vertical='center')
    total_cell.number_format = Config.CURRENCY_FORMAT
    total_cell.font = Styles.calc_font()
    start_month_col = get_column_letter(start_col + 1)
    end_month_col = get_column_letter(start_col + num_months)
    total_cell.value = f"=SUM({start_month_col}{row}:{end_month_col}{row})"
    
    # Alternating row color
    if row % 2 == 0:
        for col in range(start_col, total_col + 1):
            ws.cell(row=row, column=col).fill = Styles.alt_row_fill()


def create_formula_row(ws, row, label, formula_template, indent=False, 
                       use_percent=False, start_col=2, num_months=12):
    """
    Create a row with formulas for each month.
    
    Args:
        ws: Worksheet object
        row: Row number
        label: Label text
        formula_template: Formula template with {col} placeholder (e.g., "={col}8*0.1")
        indent: Whether to indent the label
        use_percent: Whether to use percentage format
        start_col: Starting column
        num_months: Number of month columns
    """
    # Label column
    label_cell = ws.cell(row=row, column=start_col)
    label_cell.value = f"  {label}" if indent else label
    label_cell.font = Font(size=10, bold=not indent, color=Styles.TEXT_DARK)
    label_cell.alignment = Alignment(horizontal='left', vertical='center')
    label_cell.border = Styles.THIN_BORDER
    
    # Month columns with formulas
    for i in range(num_months):
        col = start_col + 1 + i
        col_letter = get_column_letter(col)
        cell = ws.cell(row=row, column=col)
        cell.value = formula_template.replace('{col}', col_letter)
        cell.font = Styles.calc_font()
        cell.border = Styles.THIN_BORDER
        cell.alignment = Alignment(horizontal='right', vertical='center')
        cell.number_format = Config.PERCENT_FORMAT if use_percent else Config.CURRENCY_FORMAT
    
    # Year Total
    total_col = start_col + 1 + num_months
    total_cell = ws.cell(row=row, column=total_col)
    start_month_col = get_column_letter(start_col + 1)
    end_month_col = get_column_letter(start_col + num_months)
    total_cell.value = f"=SUM({start_month_col}{row}:{end_month_col}{row})"
    total_cell.font = Styles.calc_font()
    total_cell.border = Styles.THIN_BORDER
    total_cell.alignment = Alignment(horizontal='right', vertical='center')
    total_cell.number_format = Config.PERCENT_FORMAT if use_percent else Config.CURRENCY_FORMAT
    
    if row % 2 == 0:
        for col in range(start_col, total_col + 1):
            ws.cell(row=row, column=col).fill = Styles.alt_row_fill()


# =============================================================================
# SHEET CREATION FUNCTIONS
# =============================================================================

def create_cover_sheet(wb):
    """Create the Cover sheet with overview and instructions."""
    ws = wb.active
    ws.title = "Cover"
    ws.sheet_view.showGridLines = False
    
    # Column widths
    ws.column_dimensions['A'].width = 3
    ws.column_dimensions['B'].width = 35
    ws.column_dimensions['C'].width = 50
    ws.column_dimensions['D'].width = 3
    
    # Title
    ws.merge_cells('B2:C2')
    ws['B2'] = "PERSONAL BUDGET WORKBOOK"
    ws['B2'].font = Font(size=24, bold=True, color=Styles.HEADER_DARK)
    ws['B2'].alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[2].height = 40
    
    # Subtitle
    ws.merge_cells('B3:C3')
    ws['B3'] = "South African Rand (ZAR) | Manual Entry System"
    ws['B3'].font = Font(size=12, italic=True, color="666666")
    ws['B3'].alignment = Alignment(horizontal='center')
    
    # Description
    ws.merge_cells('B5:C5')
    ws['B5'] = "Track your actual income, expenses, and savings with dynamic calculations."
    ws['B5'].font = Font(size=11, color="444444")
    ws['B5'].alignment = Alignment(horizontal='center', wrap_text=True)
    
    # Key Features Section
    ws['B7'] = "KEY FEATURES"
    ws['B7'].font = Font(size=14, bold=True, color=Styles.HEADER_DARK)
    ws.row_dimensions[7].height = 25
    
    features = [
        ("✓ Manual Monthly Entry", "Enter your actual income and expenses each month"),
        ("✓ Automatic Calculations", "All totals, variances, and rates calculate automatically"),
        ("✓ Trend Analysis", "Month-to-month spending and savings comparisons"),
        ("✓ Emergency Fund Tracking", "Monitor your 6-month expense buffer progress"),
        ("✓ 5-Year Projection", "Long-term savings growth at 8% annual return"),
    ]
    
    row = 9
    for feature, desc in features:
        ws[f'B{row}'] = feature
        ws[f'B{row}'].font = Font(size=11, bold=True, color=Styles.HEADER_LIGHT)
        ws[f'C{row}'] = desc
        ws[f'C{row}'].font = Font(size=10, color="555555")
        row += 1
    
    # Sheet Index
    ws[f'B{row+1}'] = "WORKBOOK STRUCTURE"
    ws[f'B{row+1}'].font = Font(size=14, bold=True, color=Styles.HEADER_DARK)
    
    sheet_index = [
        ("Monthly Entry", "Input all income, expenses, savings, and bonuses"),
        ("Summary Dashboard", "Overview with expense breakdown and key metrics"),
        ("Trend Analysis", "Month-to-month changes with conditional formatting"),
        ("Bonus Tracker", "Track quarterly and annual bonus income"),
        ("Emergency Fund", "Monitor emergency fund progress vs 6-month target"),
        ("5-Year Projection", "Long-term savings projection at 8% growth"),
    ]
    
    row += 3
    ws[f'B{row}'] = "Sheet Name"
    ws[f'B{row}'].font = Font(bold=True, color=Styles.HEADER_DARK)
    ws[f'B{row}'].fill = PatternFill(start_color="D6E3F8", end_color="D6E3F8", fill_type="solid")
    ws[f'C{row}'] = "Purpose"
    ws[f'C{row}'].font = Font(bold=True, color=Styles.HEADER_DARK)
    ws[f'C{row}'].fill = PatternFill(start_color="D6E3F8", end_color="D6E3F8", fill_type="solid")
    row += 1
    
    for sheet_name, purpose in sheet_index:
        ws[f'B{row}'] = sheet_name
        ws[f'B{row}'].font = Font(size=10, bold=True)
        ws[f'C{row}'] = purpose
        ws[f'C{row}'].font = Font(size=10, color="555555")
        if row % 2 == 0:
            ws[f'B{row}'].fill = Styles.alt_row_fill()
            ws[f'C{row}'].fill = Styles.alt_row_fill()
        row += 1
    
    # Instructions
    ws[f'B{row+2}'] = "HOW TO USE THIS WORKBOOK"
    ws[f'B{row+2}'].font = Font(size=14, bold=True, color=Styles.HEADER_DARK)
    
    instructions = [
        "1. BLUE TEXT cells are for YOUR INPUT - enter actual amounts only",
        "2. BLACK TEXT cells contain FORMULAS - do not modify these",
        "3. Enter data month by month in the 'Monthly Entry' sheet",
        "4. All other sheets update automatically based on your entries",
    ]
    
    row += 4
    for instruction in instructions:
        ws[f'B{row}'] = instruction
        ws[f'B{row}'].font = Font(size=10, color="444444")
        row += 1
    
    # Color Legend
    ws[f'B{row+2}'] = "COLOR LEGEND"
    ws[f'B{row+2}'].font = Font(size=12, bold=True, color=Styles.HEADER_DARK)
    
    row += 4
    ws[f'B{row}'] = "Blue Text"
    ws[f'B{row}'].font = Font(color=Styles.INPUT_BLUE, bold=True)
    ws[f'C{row}'] = "Manual input cells - enter your actual amounts here"
    ws[f'C{row}'].font = Font(size=10)
    
    ws[f'B{row+1}'] = "Black Text"
    ws[f'B{row+1}'].font = Font(color=Styles.TEXT_DARK, bold=True)
    ws[f'C{row+1}'] = "Formula cells - calculations happen automatically"
    ws[f'C{row+1}'].font = Font(size=10)
    
    ws[f'B{row+2}'] = "Green Highlight"
    ws[f'B{row+2}'].fill = PatternFill(start_color=Styles.GREEN_FILL, end_color=Styles.GREEN_FILL, fill_type="solid")
    ws[f'C{row+2}'] = "Positive trend - spending decreased or savings increased"
    ws[f'C{row+2}'].font = Font(size=10)
    
    ws[f'B{row+3}'] = "Red Highlight"
    ws[f'B{row+3}'].fill = PatternFill(start_color=Styles.RED_FILL, end_color=Styles.RED_FILL, fill_type="solid")
    ws[f'C{row+3}'] = "Negative trend - spending increased or savings decreased"
    ws[f'C{row+3}'].font = Font(size=10)
    
    return ws


def create_monthly_entry_sheet(wb):
    """Create the Monthly Entry sheet for data input."""
    ws = wb.create_sheet("Monthly Entry")
    ws.sheet_view.showGridLines = False
    
    # Column widths
    ws.column_dimensions['A'].width = 3
    ws.column_dimensions['B'].width = 28
    for i in range(12):
        ws.column_dimensions[get_column_letter(3 + i)].width = 14
    ws.column_dimensions['O'].width = 16
    
    # Hidden reference columns
    ws.column_dimensions['Q'].hidden = True
    ws.column_dimensions['R'].hidden = True
    
    # Title
    ws.merge_cells('B2:O2')
    ws['B2'] = "MONTHLY DATA ENTRY"
    ws['B2'].font = Font(size=18, bold=True, color=Styles.HEADER_DARK)
    ws['B2'].alignment = Alignment(horizontal='left', vertical='center')
    ws.row_dimensions[2].height = 30
    
    # Subtitle
    ws.merge_cells('B3:O3')
    ws['B3'] = "Enter your actual income and expenses in the BLUE cells below."
    ws['B3'].font = Font(size=10, italic=True, color="666666")
    
    # Month headers
    ws['B5'] = "CATEGORY"
    ws['B5'].font = Styles.header_font()
    ws['B5'].fill = Styles.header_fill()
    ws['B5'].alignment = Alignment(horizontal='left', vertical='center')
    ws['B5'].border = Styles.THIN_BORDER
    
    months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
    for i, month in enumerate(months):
        col = get_column_letter(3 + i)
        ws[f'{col}5'] = month
        ws[f'{col}5'].font = Styles.header_font()
        ws[f'{col}5'].fill = Styles.header_fill()
        ws[f'{col}5'].alignment = Alignment(horizontal='center', vertical='center')
        ws[f'{col}5'].border = Styles.THIN_BORDER
    
    ws['O5'] = "YEAR TOTAL"
    ws['O5'].font = Styles.header_font()
    ws['O5'].fill = Styles.header_fill()
    ws['O5'].alignment = Alignment(horizontal='center', vertical='center')
    ws['O5'].border = Styles.THIN_BORDER
    ws.row_dimensions[5].height = 25
    
    row = 7
    
    # INCOME SECTION
    create_section_header(ws, row, "INCOME")
    row += 1
    create_input_row(ws, row, "Net Income Received")
    net_income_row = row
    row += 1
    
    # FIXED EXPENSES
    row += 1
    create_section_header(ws, row, "FIXED EXPENSES")
    row += 1
    
    fixed_expenses = ['Data', 'Transport', 'Subscriptions', 'TFG Debit', 'Family Support']
    for expense in fixed_expenses:
        create_input_row(ws, row, expense, indent=True)
        row += 1
    
    # Tithe (auto-calculated at 10% of Net Income)
    create_formula_row(ws, row, "Tithe (10% of Net Income)", 
                       f"={{col}}{net_income_row}*0.1", indent=True)
    row += 1
    
    # VARIABLE EXPENSES
    row += 1
    create_section_header(ws, row, "VARIABLE EXPENSES")
    row += 1
    
    variable_expenses = ['Groceries', 'Eating Out', 'Lunch', 'Haircuts', 'Clothing', 'Random Spending']
    for expense in variable_expenses:
        create_input_row(ws, row, expense, indent=True)
        row += 1
    
    # FUTURE RENT (Toggle Section)
    row += 1
    create_section_header(ws, row, "FUTURE RENT (Toggle On/Off)")
    row += 1
    
    # Rent toggle (1=Yes, 0=No)
    ws[f'B{row}'] = "  Rent Active? (1=Yes, 0=No)"
    ws[f'B{row}'].font = Font(size=10, color=Styles.TEXT_DARK)
    ws[f'B{row}'].alignment = Alignment(horizontal='left', vertical='center')
    ws[f'B{row}'].border = Styles.THIN_BORDER
    
    rent_toggle_row = row
    for i in range(12):
        col = get_column_letter(3 + i)
        cell = ws[f'{col}{row}']
        cell.value = 0
        cell.font = Styles.input_font()
        cell.border = Styles.THIN_BORDER
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.number_format = '0'
    row += 1
    
    # Rent amount
    ws[f'B{row}'] = f"  Rent Amount (R{Config.DEFAULT_RENT_AMOUNT:,})"
    ws[f'B{row}'].font = Font(size=10, color=Styles.TEXT_DARK)
    ws[f'B{row}'].alignment = Alignment(horizontal='left', vertical='center')
    ws[f'B{row}'].border = Styles.THIN_BORDER
    
    rent_amount_row = row
    for i in range(12):
        col = get_column_letter(3 + i)
        cell = ws[f'{col}{row}']
        cell.value = Config.DEFAULT_RENT_AMOUNT
        cell.font = Styles.input_font()
        cell.border = Styles.THIN_BORDER
        cell.alignment = Alignment(horizontal='right', vertical='center')
        cell.number_format = Config.CURRENCY_FORMAT
    row += 1
    
    # Rent paid (calculated based on toggle)
    ws[f'B{row}'] = "  Rent Paid (calculated)"
    ws[f'B{row}'].font = Font(size=10, color=Styles.TEXT_DARK)
    ws[f'B{row}'].alignment = Alignment(horizontal='left', vertical='center')
    ws[f'B{row}'].border = Styles.THIN_BORDER
    
    for i in range(12):
        col = get_column_letter(3 + i)
        cell = ws[f'{col}{row}']
        cell.value = f"=IF({col}{rent_toggle_row}=1,{col}{rent_amount_row},0)"
        cell.font = Styles.calc_font()
        cell.border = Styles.THIN_BORDER
        cell.alignment = Alignment(horizontal='right', vertical='center')
        cell.number_format = Config.CURRENCY_FORMAT
    
    ws[f'O{row}'].value = f"=SUM(C{row}:N{row})"
    ws[f'O{row}'].font = Styles.calc_font()
    ws[f'O{row}'].border = Styles.THIN_BORDER
    ws[f'O{row}'].alignment = Alignment(horizontal='right', vertical='center')
    ws[f'O{row}'].number_format = Config.CURRENCY_FORMAT
    row += 1
    
    # SAVINGS & TRANSFERS
    row += 1
    create_section_header(ws, row, "SAVINGS & TRANSFERS")
    row += 1
    create_input_row(ws, row, "Transfer to Savings")
    savings_row = row
    row += 1
    create_input_row(ws, row, "Starting Savings Balance")
    starting_savings_row = row
    row += 1
    
    # BONUSES
    row += 1
    create_section_header(ws, row, "BONUSES RECEIVED")
    row += 1
    create_input_row(ws, row, "Bonus Received")
    bonus_row = row
    row += 1
    
    # UNEXPECTED EXPENSES
    row += 1
    create_section_header(ws, row, "UNEXPECTED EXPENSES")
    row += 1
    create_input_row(ws, row, "Unexpected Expenses")
    unexpected_row = row
    row += 1
    
    # SUMMARY TOTALS
    row += 1
    create_section_header(ws, row, "SUMMARY TOTALS")
    row += 1
    
    # Total Fixed Expenses
    ws[f'B{row}'] = "Total Fixed Expenses"
    ws[f'B{row}'].font = Font(size=10, bold=True, color=Styles.TEXT_DARK)
    ws[f'B{row}'].alignment = Alignment(horizontal='left', vertical='center')
    ws[f'B{row}'].border = Styles.THIN_BORDER
    
    total_fixed_row = row
    for i in range(12):
        col = get_column_letter(3 + i)
        cell = ws[f'{col}{row}']
        # Sum of Data, Transport, Subscriptions, TFG, Family, Tithe
        cell.value = f"={col}{MonthlyEntryRows.DATA}+{col}{MonthlyEntryRows.TRANSPORT}+{col}{MonthlyEntryRows.SUBSCRIPTIONS}+{col}{MonthlyEntryRows.TFG_DEBIT}+{col}{MonthlyEntryRows.FAMILY_SUPPORT}+{col}{MonthlyEntryRows.TITHE}"
        cell.font = Styles.calc_font()
        cell.border = Styles.THIN_BORDER
        cell.alignment = Alignment(horizontal='right', vertical='center')
        cell.number_format = Config.CURRENCY_FORMAT
    
    ws[f'O{row}'].value = f"=SUM(C{row}:N{row})"
    ws[f'O{row}'].font = Styles.calc_font()
    ws[f'O{row}'].border = Styles.THIN_BORDER
    ws[f'O{row}'].alignment = Alignment(horizontal='right', vertical='center')
    ws[f'O{row}'].number_format = Config.CURRENCY_FORMAT
    row += 1
    
    # Total Variable Expenses
    ws[f'B{row}'] = "Total Variable Expenses"
    ws[f'B{row}'].font = Font(size=10, bold=True, color=Styles.TEXT_DARK)
    ws[f'B{row}'].alignment = Alignment(horizontal='left', vertical='center')
    ws[f'B{row}'].border = Styles.THIN_BORDER
    
    for i in range(12):
        col = get_column_letter(3 + i)
        cell = ws[f'{col}{row}']
        cell.value = f"={col}{MonthlyEntryRows.GROCERIES}+{col}{MonthlyEntryRows.EATING_OUT}+{col}{MonthlyEntryRows.LUNCH}+{col}{MonthlyEntryRows.HAIRCUTS}+{col}{MonthlyEntryRows.CLOTHING}+{col}{MonthlyEntryRows.RANDOM_SPENDING}"
        cell.font = Styles.calc_font()
        cell.border = Styles.THIN_BORDER
        cell.alignment = Alignment(horizontal='right', vertical='center')
        cell.number_format = Config.CURRENCY_FORMAT
    
    ws[f'O{row}'].value = f"=SUM(C{row}:N{row})"
    ws[f'O{row}'].font = Styles.calc_font()
    ws[f'O{row}'].border = Styles.THIN_BORDER
    ws[f'O{row}'].alignment = Alignment(horizontal='right', vertical='center')
    ws[f'O{row}'].number_format = Config.CURRENCY_FORMAT
    row += 1
    
    # TOTAL ALL EXPENSES (highlighted)
    ws[f'B{row}'] = "TOTAL ALL EXPENSES"
    ws[f'B{row}'].font = Font(size=11, bold=True, color=Styles.WHITE)
    ws[f'B{row}'].fill = Styles.header_fill()
    ws[f'B{row}'].alignment = Alignment(horizontal='left', vertical='center')
    ws[f'B{row}'].border = Styles.THIN_BORDER
    
    for i in range(12):
        col = get_column_letter(3 + i)
        cell = ws[f'{col}{row}']
        cell.value = f"={col}{total_fixed_row}+{col}{row-1}+{col}{row-13}+{col}{unexpected_row}"
        cell.font = Font(bold=True, size=11, color=Styles.WHITE)
        cell.fill = Styles.header_fill()
        cell.border = Styles.THIN_BORDER
        cell.alignment = Alignment(horizontal='right', vertical='center')
        cell.number_format = Config.CURRENCY_FORMAT
    
    ws[f'O{row}'].value = f"=SUM(C{row}:N{row})"
    ws[f'O{row}'].font = Font(bold=True, size=11, color=Styles.WHITE)
    ws[f'O{row}'].fill = Styles.header_fill()
    ws[f'O{row}'].border = Styles.THIN_BORDER
    ws[f'O{row}'].alignment = Alignment(horizontal='right', vertical='center')
    ws[f'O{row}'].number_format = Config.CURRENCY_FORMAT
    row += 1
    
    # TOTAL INCOME (highlighted green)
    ws[f'B{row}'] = "TOTAL INCOME"
    ws[f'B{row}'].font = Font(size=11, bold=True, color=Styles.WHITE)
    ws[f'B{row}'].fill = Styles.positive_fill()
    ws[f'B{row}'].alignment = Alignment(horizontal='left', vertical='center')
    ws[f'B{row}'].border = Styles.THIN_BORDER
    
    for i in range(12):
        col = get_column_letter(3 + i)
        cell = ws[f'{col}{row}']
        cell.value = f"={col}{net_income_row}+{col}{bonus_row}"
        cell.font = Font(bold=True, size=11, color=Styles.WHITE)
        cell.fill = Styles.positive_fill()
        cell.border = Styles.THIN_BORDER
        cell.alignment = Alignment(horizontal='right', vertical='center')
        cell.number_format = Config.CURRENCY_FORMAT
    
    ws[f'O{row}'].value = f"=SUM(C{row}:N{row})"
    ws[f'O{row}'].font = Font(bold=True, size=11, color=Styles.WHITE)
    ws[f'O{row}'].fill = Styles.positive_fill()
    ws[f'O{row}'].border = Styles.THIN_BORDER
    ws[f'O{row}'].alignment = Alignment(horizontal='right', vertical='center')
    ws[f'O{row}'].number_format = Config.CURRENCY_FORMAT
    row += 1
    
    # SURPLUS / DEFICIT
    ws[f'B{row}'] = "SURPLUS / (DEFICIT)"
    ws[f'B{row}'].font = Font(size=11, bold=True, color=Styles.WHITE)
    ws[f'B{row}'].fill = PatternFill(start_color=Styles.ACCENT_BLUE, end_color=Styles.ACCENT_BLUE, fill_type="solid")
    ws[f'B{row}'].alignment = Alignment(horizontal='left', vertical='center')
    ws[f'B{row}'].border = Styles.THIN_BORDER
    
    for i in range(12):
        col = get_column_letter(3 + i)
        cell = ws[f'{col}{row}']
        cell.value = f"={col}{row-2}-{col}{row-3}-{col}{savings_row}"
        cell.font = Font(bold=True, size=11, color=Styles.WHITE)
        cell.fill = PatternFill(start_color=Styles.ACCENT_BLUE, end_color=Styles.ACCENT_BLUE, fill_type="solid")
        cell.border = Styles.THIN_BORDER
        cell.alignment = Alignment(horizontal='right', vertical='center')
        cell.number_format = Config.CURRENCY_FORMAT
    
    ws[f'O{row}'].value = f"=SUM(C{row}:N{row})"
    ws[f'O{row}'].font = Font(bold=True, size=11, color=Styles.WHITE)
    ws[f'O{row}'].fill = PatternFill(start_color=Styles.ACCENT_BLUE, end_color=Styles.ACCENT_BLUE, fill_type="solid")
    ws[f'O{row}'].border = Styles.THIN_BORDER
    ws[f'O{row}'].alignment = Alignment(horizontal='right', vertical='center')
    ws[f'O{row}'].number_format = Config.CURRENCY_FORMAT
    row += 1
    
    # Savings Rate %
    ws[f'B{row}'] = "Savings Rate %"
    ws[f'B{row}'].font = Font(size=10, bold=True, color=Styles.TEXT_DARK)
    ws[f'B{row}'].alignment = Alignment(horizontal='left', vertical='center')
    ws[f'B{row}'].border = Styles.THIN_BORDER
    
    for i in range(12):
        col = get_column_letter(3 + i)
        cell = ws[f'{col}{row}']
        cell.value = f"=IF({col}{row-4}=0,0,{col}{savings_row}/{col}{row-4})"
        cell.font = Styles.calc_font()
        cell.border = Styles.THIN_BORDER
        cell.alignment = Alignment(horizontal='right', vertical='center')
        cell.number_format = Config.PERCENT_FORMAT
    
    ws[f'O{row}'].value = f"=IF(O{row-4}=0,0,O{savings_row}/O{row-4})"
    ws[f'O{row}'].font = Styles.calc_font()
    ws[f'O{row}'].border = Styles.THIN_BORDER
    ws[f'O{row}'].alignment = Alignment(horizontal='right', vertical='center')
    ws[f'O{row}'].number_format = Config.PERCENT_FORMAT
    row += 1
    
    # Running Savings Balance
    ws[f'B{row}'] = "Running Savings Balance"
    ws[f'B{row}'].font = Font(size=10, bold=True, color=Styles.TEXT_DARK)
    ws[f'B{row}'].alignment = Alignment(horizontal='left', vertical='center')
    ws[f'B{row}'].border = Styles.THIN_BORDER
    
    # January = Starting Balance + Jan Savings
    ws[f'C{row}'].value = f"=C{starting_savings_row}+C{savings_row}"
    ws[f'C{row}'].font = Styles.calc_font()
    ws[f'C{row}'].border = Styles.THIN_BORDER
    ws[f'C{row}'].alignment = Alignment(horizontal='right', vertical='center')
    ws[f'C{row}'].number_format = Config.CURRENCY_FORMAT
    
    # Feb-Dec = Previous Month + Current Month Savings
    for i in range(1, 12):
        col = get_column_letter(3 + i)
        prev_col = get_column_letter(3 + i - 1)
        cell = ws[f'{col}{row}']
        cell.value = f"={prev_col}{row}+{col}{savings_row}"
        cell.font = Styles.calc_font()
        cell.border = Styles.THIN_BORDER
        cell.alignment = Alignment(horizontal='right', vertical='center')
        cell.number_format = Config.CURRENCY_FORMAT
    
    ws[f'O{row}'].value = f"=N{row}"
    ws[f'O{row}'].font = Styles.calc_font()
    ws[f'O{row}'].border = Styles.THIN_BORDER
    ws[f'O{row}'].alignment = Alignment(horizontal='right', vertical='center')
    ws[f'O{row}'].number_format = Config.CURRENCY_FORMAT
    
    # Store key row references (hidden)
    ws['Q5'] = "KEY ROW REFERENCE"
    ws['Q5'].font = Font(bold=True, color=Styles.HEADER_DARK)
    
    refs = [
        ("Net Income Row:", MonthlyEntryRows.NET_INCOME),
        ("Total Expenses Row:", MonthlyEntryRows.TOTAL_EXPENSES),
        ("Total Income Row:", MonthlyEntryRows.TOTAL_INCOME),
        ("Savings Row:", MonthlyEntryRows.SAVINGS_TRANSFER),
        ("Bonus Row:", MonthlyEntryRows.BONUS),
        ("Surplus Row:", MonthlyEntryRows.SURPLUS_DEFICIT),
        ("Savings Rate Row:", MonthlyEntryRows.SAVINGS_RATE),
        ("Running Savings Row:", MonthlyEntryRows.RUNNING_SAVINGS),
    ]
    
    for i, (label, row_num) in enumerate(refs):
        ws[f'Q{7+i}'] = label
        ws[f'R{7+i}'] = row_num
    
    return ws


def create_summary_dashboard(wb):
    """Create the Summary Dashboard sheet."""
    ws = wb.create_sheet("Summary Dashboard")
    ws.sheet_view.showGridLines = False
    
    ws.column_dimensions['A'].width = 3
    ws.column_dimensions['B'].width = 30
    ws.column_dimensions['C'].width = 18
    ws.column_dimensions['D'].width = 18
    ws.column_dimensions['E'].width = 18
    
    ws.merge_cells('B2:E2')
    ws['B2'] = "MONTHLY SUMMARY DASHBOARD"
    ws['B2'].font = Font(size=18, bold=True, color=Styles.HEADER_DARK)
    ws['B2'].alignment = Alignment(horizontal='left', vertical='center')
    ws.row_dimensions[2].height = 30
    
    ws.merge_cells('B3:E3')
    ws['B3'] = "Overview of your financial performance - updates automatically from Monthly Entry"
    ws['B3'].font = Font(size=10, italic=True, color="666666")
    
    ws['B5'] = "KEY METRICS (Year-to-Date)"
    ws['B5'].font = Font(size=14, bold=True, color=Styles.HEADER_DARK)
    ws.row_dimensions[5].height = 25
    
    metrics = [
        ("Total Income (Net + Bonus)", "='Monthly Entry'!O45", Config.CURRENCY_FORMAT),
        ("Total Expenses", "='Monthly Entry'!O44", Config.CURRENCY_FORMAT),
        ("Net Savings (Transferred)", "='Monthly Entry'!O32", Config.CURRENCY_FORMAT),
        ("Total Bonuses Received", "='Monthly Entry'!O36", Config.CURRENCY_FORMAT),
        ("Unexpected Expenses", "='Monthly Entry'!O39", Config.CURRENCY_FORMAT),
        ("Year-End Surplus/(Deficit)", "='Monthly Entry'!O46", Config.CURRENCY_FORMAT),
        ("Average Savings Rate", "='Monthly Entry'!O47", Config.PERCENT_FORMAT),
        ("Year-End Savings Balance", "='Monthly Entry'!O48", Config.CURRENCY_FORMAT),
    ]
    
    row = 7
    for label, formula, fmt in metrics:
        ws[f'B{row}'] = label
        ws[f'B{row}'].font = Font(size=11, color=Styles.TEXT_DARK)
        ws[f'B{row}'].alignment = Alignment(horizontal='left', vertical='center')
        ws[f'B{row}'].border = Styles.THIN_BORDER
        
        ws.merge_cells(f'C{row}:D{row}')
        ws[f'C{row}'] = formula
        ws[f'C{row}'].font = Font(size=12, bold=True, color=Styles.HEADER_LIGHT)
        ws[f'C{row}'].alignment = Alignment(horizontal='right', vertical='center')
        ws[f'C{row}'].number_format = fmt
        ws[f'C{row}'].border = Styles.THIN_BORDER
        
        if row % 2 == 0:
            ws[f'B{row}'].fill = Styles.alt_row_fill()
            ws[f'C{row}'].fill = Styles.alt_row_fill()
        row += 1
    
    # Expense Breakdown
    row += 2
    ws[f'B{row}'] = "EXPENSE BREAKDOWN BY CATEGORY"
    ws[f'B{row}'].font = Font(size=14, bold=True, color=Styles.HEADER_DARK)
    
    row += 2
    ws[f'B{row}'] = "Category"
    ws[f'B{row}'].font = Styles.header_font()
    ws[f'B{row}'].fill = Styles.header_fill()
    ws[f'B{row}'].border = Styles.THIN_BORDER
    ws[f'B{row}'].alignment = Alignment(horizontal='left', vertical='center')
    
    ws[f'C{row}'] = "Year Total"
    ws[f'C{row}'].font = Styles.header_font()
    ws[f'C{row}'].fill = Styles.header_fill()
    ws[f'C{row}'].border = Styles.THIN_BORDER
    ws[f'C{row}'].alignment = Alignment(horizontal='right', vertical='center')
    
    ws[f'D{row}'] = "% of Total"
    ws[f'D{row}'].font = Styles.header_font()
    ws[f'D{row}'].fill = Styles.header_fill()
    ws[f'D{row}'].border = Styles.THIN_BORDER
    ws[f'D{row}'].alignment = Alignment(horizontal='right', vertical='center')
    ws.row_dimensions[row].height = 22
    
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
        ws[f'B{row}'] = category
        ws[f'B{row}'].font = Font(size=10, color=Styles.TEXT_DARK)
        ws[f'B{row}'].border = Styles.THIN_BORDER
        ws[f'B{row}'].alignment = Alignment(horizontal='left', vertical='center')
        
        ws[f'C{row}'] = formula
        ws[f'C{row}'].font = Styles.calc_font()
        ws[f'C{row}'].border = Styles.THIN_BORDER
        ws[f'C{row}'].alignment = Alignment(horizontal='right', vertical='center')
        ws[f'C{row}'].number_format = Config.CURRENCY_FORMAT
        
        ws[f'D{row}'] = f"=IF($C${row+14}=0,0,C{row}/$C${row+14})"
        ws[f'D{row}'].font = Styles.calc_font()
        ws[f'D{row}'].border = Styles.THIN_BORDER
        ws[f'D{row}'].alignment = Alignment(horizontal='right', vertical='center')
        ws[f'D{row}'].number_format = Config.PERCENT_FORMAT
        
        if row % 2 == 0:
            ws[f'B{row}'].fill = Styles.alt_row_fill()
            ws[f'C{row}'].fill = Styles.alt_row_fill()
            ws[f'D{row}'].fill = Styles.alt_row_fill()
        row += 1
    
    # Total row
    ws[f'B{row}'] = "TOTAL EXPENSES"
    ws[f'B{row}'].font = Font(bold=True, color=Styles.WHITE)
    ws[f'B{row}'].fill = Styles.header_fill()
    ws[f'B{row}'].border = Styles.THIN_BORDER
    ws[f'B{row}'].alignment = Alignment(horizontal='left', vertical='center')
    
    ws[f'C{row}'] = "='Monthly Entry'!O44"
    ws[f'C{row}'].font = Font(bold=True, color=Styles.WHITE)
    ws[f'C{row}'].fill = Styles.header_fill()
    ws[f'C{row}'].border = Styles.THIN_BORDER
    ws[f'C{row}'].alignment = Alignment(horizontal='right', vertical='center')
    ws[f'C{row}'].number_format = Config.CURRENCY_FORMAT
    
    ws[f'D{row}'] = "=IF(C33=0,0,C33/C33)"
    ws[f'D{row}'].font = Font(bold=True, color=Styles.WHITE)
    ws[f'D{row}'].fill = Styles.header_fill()
    ws[f'D{row}'].border = Styles.THIN_BORDER
    ws[f'D{row}'].alignment = Alignment(horizontal='right', vertical='center')
    ws[f'D{row}'].number_format = Config.PERCENT_FORMAT
    
    return ws


def create_trend_analysis(wb):
    """Create the Trend Analysis sheet with conditional formatting."""
    ws = wb.create_sheet("Trend Analysis")
    ws.sheet_view.showGridLines = False
    
    ws.column_dimensions['A'].width = 3
    ws.column_dimensions['B'].width = 28
    for i in range(11):
        ws.column_dimensions[get_column_letter(3 + i)].width = 14
    
    ws.merge_cells('B2:N2')
    ws['B2'] = "TREND ANALYSIS"
    ws['B2'].font = Font(size=18, bold=True, color=Styles.HEADER_DARK)
    ws['B2'].alignment = Alignment(horizontal='left', vertical='center')
    ws.row_dimensions[2].height = 30
    
    ws.merge_cells('B3:N3')
    ws['B3'] = "Month-to-month changes - GREEN means improvement, RED means increase in spending"
    ws['B3'].font = Font(size=10, italic=True, color="666666")
    
    ws['B4'] = "▼ = Spending Decreased (Good) | ▲ = Spending Increased (Warning)"
    ws['B4'].font = Font(size=9, color="666666")
    
    # Headers
    ws['B6'] = "METRIC"
    ws['B6'].font = Styles.header_font()
    ws['B6'].fill = Styles.header_fill()
    ws['B6'].border = Styles.THIN_BORDER
    ws['B6'].alignment = Alignment(horizontal='left', vertical='center')
    
    change_labels = ['Feb-Jan', 'Mar-Feb', 'Apr-Mar', 'May-Apr', 'Jun-May', 
                     'Jul-Jun', 'Aug-Jul', 'Sep-Aug', 'Oct-Sep', 'Nov-Oct', 'Dec-Nov']
    for i, label in enumerate(change_labels):
        col = get_column_letter(3 + i)
        ws[f'{col}6'] = label
        ws[f'{col}6'].font = Styles.header_font()
        ws[f'{col}6'].fill = Styles.header_fill()
        ws[f'{col}6'].border = Styles.THIN_BORDER
        ws[f'{col}6'].alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[6].height = 25
    
    def create_trend_row(row, label, monthly_row, use_percent=False):
        ws[f'B{row}'] = label
        ws[f'B{row}'].font = Font(size=10, color=Styles.TEXT_DARK)
        ws[f'B{row}'].border = Styles.THIN_BORDER
        ws[f'B{row}'].alignment = Alignment(horizontal='left', vertical='center')
        
        month_pairs = [('D', 'C'), ('E', 'D'), ('F', 'E'), ('G', 'F'), ('H', 'G'),
                       ('I', 'H'), ('J', 'I'), ('K', 'J'), ('L', 'K'), ('M', 'L'), ('N', 'M')]
        
        for curr_col, prev_col in month_pairs:
            cell = ws[f'{curr_col}{row}']
            cell.value = f"='Monthly Entry'!{curr_col}{monthly_row}-'Monthly Entry'!{prev_col}{monthly_row}"
            cell.font = Styles.calc_font()
            cell.border = Styles.THIN_BORDER
            cell.alignment = Alignment(horizontal='right', vertical='center')
            cell.number_format = Config.PERCENT_FORMAT if use_percent else Config.CURRENCY_FORMAT
        
        if row % 2 == 0:
            ws[f'B{row}'].fill = Styles.alt_row_fill()
    
    row = 8
    
    # SPENDING TRENDS
    ws.merge_cells(f'B{row}:N{row}')
    ws[f'B{row}'] = "SPENDING TRENDS (RED = Increased, GREEN = Decreased)"
    ws[f'B{row}'].font = Styles.section_font()
    ws[f'B{row}'].fill = Styles.negative_fill()
    for col in range(2, 15):
        ws.cell(row=row, column=col).border = Styles.THIN_BORDER
    row += 1
    
    create_trend_row(row, "Total Spending Change", MonthlyEntryRows.TOTAL_EXPENSES)
    row += 1
    create_trend_row(row, "Fixed Expenses Change", MonthlyEntryRows.TOTAL_FIXED)
    row += 1
    create_trend_row(row, "Variable Expenses Change", MonthlyEntryRows.TOTAL_VARIABLE)
    row += 1
    
    individual_expenses = [
        ("  Data", MonthlyEntryRows.DATA),
        ("  Transport", MonthlyEntryRows.TRANSPORT),
        ("  Subscriptions", MonthlyEntryRows.SUBSCRIPTIONS),
        ("  TFG Debit", MonthlyEntryRows.TFG_DEBIT),
        ("  Family Support", MonthlyEntryRows.FAMILY_SUPPORT),
        ("  Tithe", MonthlyEntryRows.TITHE),
        ("  Groceries", MonthlyEntryRows.GROCERIES),
        ("  Eating Out", MonthlyEntryRows.EATING_OUT),
        ("  Lunch", MonthlyEntryRows.LUNCH),
        ("  Haircuts", MonthlyEntryRows.HAIRCUTS),
        ("  Clothing", MonthlyEntryRows.CLOTHING),
        ("  Random Spending", MonthlyEntryRows.RANDOM_SPENDING),
    ]
    
    for label, monthly_row in individual_expenses:
        create_trend_row(row, label, monthly_row)
        row += 1
    
    # SAVINGS TRENDS
    row += 1
    ws.merge_cells(f'B{row}:N{row}')
    ws[f'B{row}'] = "SAVINGS TRENDS (GREEN = Increased, RED = Decreased)"
    ws[f'B{row}'].font = Styles.section_font()
    ws[f'B{row}'].fill = Styles.positive_fill()
    for col in range(2, 15):
        ws.cell(row=row, column=col).border = Styles.THIN_BORDER
    row += 1
    
    create_trend_row(row, "Savings Transfer Change", MonthlyEntryRows.SAVINGS_TRANSFER)
    row += 1
    create_trend_row(row, "Savings Rate % Change", MonthlyEntryRows.SAVINGS_RATE, use_percent=True)
    row += 1
    
    # INCOME TRENDS
    row += 1
    ws.merge_cells(f'B{row}:N{row}')
    ws[f'B{row}'] = "INCOME TRENDS"
    ws[f'B{row}'].font = Styles.section_font()
    ws[f'B{row}'].fill = Styles.section_fill()
    for col in range(2, 15):
        ws.cell(row=row, column=col).border = Styles.THIN_BORDER
    row += 1
    
    create_trend_row(row, "Net Income Change", MonthlyEntryRows.NET_INCOME)
    row += 1
    create_trend_row(row, "Bonus Received Change", MonthlyEntryRows.BONUS)
    row += 1
    create_trend_row(row, "Total Income Change", MonthlyEntryRows.TOTAL_INCOME)
    
    # CONDITIONAL FORMATTING
    green_fill = PatternFill(start_color=Styles.GREEN_FILL, end_color=Styles.GREEN_FILL, fill_type="solid")
    red_fill = PatternFill(start_color=Styles.RED_FILL, end_color=Styles.RED_FILL, fill_type="solid")
    
    # Expense trends: Positive = RED (bad), Negative = GREEN (good)
    for r in range(9, 24):
        for col_letter in ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N']:
            ws.conditional_formatting.add(f'{col_letter}{r}', 
                CellIsRule(operator='greaterThan', formula=['0'], fill=red_fill))
            ws.conditional_formatting.add(f'{col_letter}{r}', 
                CellIsRule(operator='lessThan', formula=['0'], fill=green_fill))
    
    # Savings trends: Positive = GREEN (good), Negative = RED (bad)
    for r in [25, 26]:
        for col_letter in ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N']:
            ws.conditional_formatting.add(f'{col_letter}{r}', 
                CellIsRule(operator='greaterThan', formula=['0'], fill=green_fill))
            ws.conditional_formatting.add(f'{col_letter}{r}', 
                CellIsRule(operator='lessThan', formula=['0'], fill=red_fill))
    
    return ws


def create_bonus_tracker(wb):
    """Create the Bonus Tracker sheet."""
    ws = wb.create_sheet("Bonus Tracker")
    ws.sheet_view.showGridLines = False
    
    ws.column_dimensions['A'].width = 3
    ws.column_dimensions['B'].width = 25
    ws.column_dimensions['C'].width = 18
    ws.column_dimensions['D'].width = 18
    ws.column_dimensions['E'].width = 25
    
    ws.merge_cells('B2:E2')
    ws['B2'] = "BONUS TRACKER"
    ws['B2'].font = Font(size=18, bold=True, color=Styles.HEADER_DARK)
    ws['B2'].alignment = Alignment(horizontal='left', vertical='center')
    ws.row_dimensions[2].height = 30
    
    ws.merge_cells('B3:E3')
    ws['B3'] = "Track your quarterly bonuses - enter actual amounts received"
    ws['B3'].font = Font(size=10, italic=True, color="666666")
    
    ws['B5'] = "EXPECTED BONUS SCHEDULE (based on your profile)"
    ws['B5'].font = Font(size=12, bold=True, color=Styles.HEADER_DARK)
    
    ws['B6'] = "Quarter"
    ws['B6'].font = Styles.header_font()
    ws['B6'].fill = Styles.header_fill()
    ws['B6'].border = Styles.THIN_BORDER
    
    ws['C6'] = "% of Gross"
    ws['C6'].font = Styles.header_font()
    ws['C6'].fill = Styles.header_fill()
    ws['C6'].border = Styles.THIN_BORDER
    ws['C6'].alignment = Alignment(horizontal='center')
    
    ws['D6'] = "Expected Amount"
    ws['D6'].font = Styles.header_font()
    ws['D6'].fill = Styles.header_fill()
    ws['D6'].border = Styles.THIN_BORDER
    ws['D6'].alignment = Alignment(horizontal='right')
    
    ws['E6'] = "Actual Received"
    ws['E6'].font = Styles.header_font()
    ws['E6'].fill = Styles.header_fill()
    ws['E6'].border = Styles.THIN_BORDER
    ws['E6'].alignment = Alignment(horizontal='right')
    ws.row_dimensions[6].height = 22
    
    bonus_schedule = [
        ("Q1 (Jan-Mar)", "0%", 0),
        ("Q2 (Apr-Jun)", "25%", Config.MONTHLY_GROSS_SALARY * Config.BONUS_Q2_PCT),
        ("Q3 (Jul-Sep)", "50%", Config.MONTHLY_GROSS_SALARY * Config.BONUS_Q3_PCT),
        ("Q4 (Oct-Dec)", "83%", Config.MONTHLY_GROSS_SALARY * Config.BONUS_Q4_PCT),
    ]
    
    row = 7
    for quarter, pct, expected in bonus_schedule:
        ws[f'B{row}'] = quarter
        ws[f'B{row}'].font = Font(size=10, color=Styles.TEXT_DARK)
        ws[f'B{row}'].border = Styles.THIN_BORDER
        
        ws[f'C{row}'] = pct
        ws[f'C{row}'].font = Font(size=10, color=Styles.TEXT_DARK)
        ws[f'C{row}'].border = Styles.THIN_BORDER
        ws[f'C{row}'].alignment = Alignment(horizontal='center')
        
        ws[f'D{row}'] = expected
        ws[f'D{row}'].font = Font(size=10, color="666666")
        ws[f'D{row}'].border = Styles.THIN_BORDER
        ws[f'D{row}'].alignment = Alignment(horizontal='right')
        ws[f'D{row}'].number_format = Config.CURRENCY_FORMAT
        
        ws[f'E{row}'].border = Styles.THIN_BORDER
        ws[f'E{row}'].alignment = Alignment(horizontal='right')
        ws[f'E{row}'].number_format = Config.CURRENCY_FORMAT
        ws[f'E{row}'].font = Styles.input_font()
        
        if row % 2 == 0:
            for c in ['B', 'C', 'D', 'E']:
                ws[f'{c}{row}'].fill = Styles.alt_row_fill()
        row += 1
    
    # Total row
    ws[f'B{row}'] = "TOTAL YEARLY BONUS"
    ws[f'B{row}'].font = Font(bold=True, color=Styles.WHITE)
    ws[f'B{row}'].fill = Styles.header_fill()
    ws[f'B{row}'].border = Styles.THIN_BORDER
    
    ws[f'C{row}'].border = Styles.THIN_BORDER
    ws[f'C{row}'].fill = Styles.header_fill()
    
    total_expected = Config.MONTHLY_GROSS_SALARY * (Config.BONUS_Q2_PCT + Config.BONUS_Q3_PCT + Config.BONUS_Q4_PCT)
    ws[f'D{row}'] = total_expected
    ws[f'D{row}'].font = Font(bold=True, color=Styles.WHITE)
    ws[f'D{row}'].fill = Styles.header_fill()
    ws[f'D{row}'].border = Styles.THIN_BORDER
    ws[f'D{row}'].alignment = Alignment(horizontal='right')
    ws[f'D{row}'].number_format = Config.CURRENCY_FORMAT
    
    ws[f'E{row}'] = "=SUM(E7:E10)"
    ws[f'E{row}'].font = Font(bold=True, color=Styles.WHITE)
    ws[f'E{row}'].fill = Styles.header_fill()
    ws[f'E{row}'].border = Styles.THIN_BORDER
    ws[f'E{row}'].alignment = Alignment(horizontal='right')
    ws[f'E{row}'].number_format = Config.CURRENCY_FORMAT
    
    row += 3
    ws[f'B{row}'] = "VARIANCE FROM EXPECTED"
    ws[f'B{row}'].font = Font(size=12, bold=True, color=Styles.HEADER_DARK)
    
    row += 2
    ws[f'B{row}'] = "Difference (Actual - Expected)"
    ws[f'B{row}'].font = Font(size=10, color=Styles.TEXT_DARK)
    ws[f'B{row}'].border = Styles.THIN_BORDER
    
    ws[f'E{row}'] = "=E11-D11"
    ws[f'E{row}'].font = Styles.calc_font()
    ws[f'E{row}'].border = Styles.THIN_BORDER
    ws[f'E{row}'].alignment = Alignment(horizontal='right')
    ws[f'E{row}'].number_format = Config.CURRENCY_FORMAT
    
    row += 3
    ws[f'B{row}'] = "MONTHLY BONUS TRACKING (from Monthly Entry)"
    ws[f'B{row}'].font = Font(size=12, bold=True, color=Styles.HEADER_DARK)
    
    row += 2
    ws[f'B{row}'] = "Month"
    ws[f'B{row}'].font = Styles.header_font()
    ws[f'B{row}'].fill = Styles.header_fill()
    ws[f'B{row}'].border = Styles.THIN_BORDER
    
    ws[f'C{row}'] = "Bonus Amount"
    ws[f'C{row}'].font = Styles.header_font()
    ws[f'C{row}'].fill = Styles.header_fill()
    ws[f'C{row}'].border = Styles.THIN_BORDER
    ws[f'C{row}'].alignment = Alignment(horizontal='right')
    ws.row_dimensions[row].height = 22
    row += 1
    
    months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
    for i, month in enumerate(months):
        col = get_column_letter(3 + i)
        ws[f'B{row}'] = month
        ws[f'B{row}'].font = Font(size=10, color=Styles.TEXT_DARK)
        ws[f'B{row}'].border = Styles.THIN_BORDER
        
        ws[f'C{row}'] = f"='Monthly Entry'!{col}{MonthlyEntryRows.BONUS}"
        ws[f'C{row}'].font = Styles.calc_font()
        ws[f'C{row}'].border = Styles.THIN_BORDER
        ws[f'C{row}'].alignment = Alignment(horizontal='right')
        ws[f'C{row}'].number_format = Config.CURRENCY_FORMAT
        
        if row % 2 == 0:
            ws[f'B{row}'].fill = Styles.alt_row_fill()
            ws[f'C{row}'].fill = Styles.alt_row_fill()
        row += 1
    
    ws[f'B{row}'] = "Total"
    ws[f'B{row}'].font = Font(bold=True, color=Styles.TEXT_DARK)
    ws[f'B{row}'].border = Styles.THIN_BORDER
    
    ws[f'C{row}'] = "='Monthly Entry'!O36"
    ws[f'C{row}'].font = Font(bold=True, color=Styles.TEXT_DARK)
    ws[f'C{row}'].border = Styles.THIN_BORDER
    ws[f'C{row}'].alignment = Alignment(horizontal='right')
    ws[f'C{row}'].number_format = Config.CURRENCY_FORMAT
    
    return ws


def create_emergency_fund(wb):
    """Create the Emergency Fund Tracker sheet."""
    ws = wb.create_sheet("Emergency Fund")
    ws.sheet_view.showGridLines = False
    
    ws.column_dimensions['A'].width = 3
    ws.column_dimensions['B'].width = 35
    ws.column_dimensions['C'].width = 20
    ws.column_dimensions['D'].width = 20
    ws.column_dimensions['E'].width = 20
    
    ws.merge_cells('B2:E2')
    ws['B2'] = "EMERGENCY FUND TRACKER"
    ws['B2'].font = Font(size=18, bold=True, color=Styles.HEADER_DARK)
    ws['B2'].alignment = Alignment(horizontal='left', vertical='center')
    ws.row_dimensions[2].height = 30
    
    ws.merge_cells('B3:E3')
    ws['B3'] = "Monitor your progress toward a 6-month expense buffer"
    ws['B3'].font = Font(size=10, italic=True, color="666666")
    
    ws['B5'] = "CURRENT STATUS"
    ws['B5'].font = Font(size=14, bold=True, color=Styles.HEADER_DARK)
    ws.row_dimensions[5].height = 25
    
    ws['B7'] = "Average Monthly Expenses"
    ws['B7'].font = Font(size=11, color=Styles.TEXT_DARK)
    ws['B7'].border = Styles.THIN_BORDER
    ws['B7'].alignment = Alignment(horizontal='left', vertical='center')
    
    ws['C7'] = "='Monthly Entry'!O44/12"
    ws['C7'].font = Font(size=12, bold=True, color=Styles.HEADER_LIGHT)
    ws['C7'].border = Styles.THIN_BORDER
    ws['C7'].alignment = Alignment(horizontal='right', vertical='center')
    ws['C7'].number_format = Config.CURRENCY_FORMAT
    
    ws['B8'] = "6-Month Emergency Fund Target"
    ws['B8'].font = Font(size=11, color=Styles.TEXT_DARK)
    ws['B8'].border = Styles.THIN_BORDER
    ws['B8'].alignment = Alignment(horizontal='left', vertical='center')
    ws['B8'].fill = Styles.alt_row_fill()
    
    ws['C8'] = "=C7*6"
    ws['C8'].font = Font(size=12, bold=True, color=Styles.HEADER_DARK)
    ws['C8'].border = Styles.THIN_BORDER
    ws['C8'].alignment = Alignment(horizontal='right', vertical='center')
    ws['C8'].number_format = Config.CURRENCY_FORMAT
    ws['C8'].fill = Styles.alt_row_fill()
    
    ws['B9'] = "Current Savings Balance"
    ws['B9'].font = Font(size=11, color=Styles.TEXT_DARK)
    ws['B9'].border = Styles.THIN_BORDER
    ws['B9'].alignment = Alignment(horizontal='left', vertical='center')
    
    ws['C9'] = "='Monthly Entry'!O48"
    ws['C9'].font = Font(size=12, bold=True, color=Styles.POSITIVE_GREEN)
    ws['C9'].border = Styles.THIN_BORDER
    ws['C9'].alignment = Alignment(horizontal='right', vertical='center')
    ws['C9'].number_format = Config.CURRENCY_FORMAT
    
    ws['B10'] = "Progress Toward Target"
    ws['B10'].font = Font(size=11, color=Styles.TEXT_DARK)
    ws['B10'].border = Styles.THIN_BORDER
    ws['B10'].alignment = Alignment(horizontal='left', vertical='center')
    ws['B10'].fill = Styles.alt_row_fill()
    
    ws['C10'] = "=IF(C8=0,0,C9/C8)"
    ws['C10'].font = Font(size=12, bold=True, color=Styles.HEADER_LIGHT)
    ws['C10'].border = Styles.THIN_BORDER
    ws['C10'].alignment = Alignment(horizontal='right', vertical='center')
    ws['C10'].number_format = Config.PERCENT_FORMAT
    ws['C10'].fill = Styles.alt_row_fill()
    
    ws['B11'] = "Amount Still Needed"
    ws['B11'].font = Font(size=11, color=Styles.TEXT_DARK)
    ws['B11'].border = Styles.THIN_BORDER
    ws['B11'].alignment = Alignment(horizontal='left', vertical='center')
    
    ws['C11'] = "=MAX(0,C8-C9)"
    ws['C11'].font = Font(size=12, bold=True, color=Styles.NEGATIVE_RED)
    ws['C11'].border = Styles.THIN_BORDER
    ws['C11'].alignment = Alignment(horizontal='right', vertical='center')
    ws['C11'].number_format = Config.CURRENCY_FORMAT
    
    ws['B13'] = "EMERGENCY FUND STATUS"
    ws['B13'].font = Font(size=12, bold=True, color=Styles.HEADER_DARK)
    
    ws['B14'] = '=IF(C10>=1,"FULLY FUNDED",IF(C10>=0.5,"HALFWAY THERE","BUILDING..."))'
    ws['B14'].font = Font(size=14, bold=True)
    ws['B14'].alignment = Alignment(horizontal='center', vertical='center')
    ws.merge_cells('B14:C14')
    ws.row_dimensions[14].height = 30
    
    ws['B17'] = "MONTHLY SAVINGS PROGRESS"
    ws['B17'].font = Font(size=14, bold=True, color=Styles.HEADER_DARK)
    
    ws['B19'] = "Month"
    ws['B19'].font = Styles.header_font()
    ws['B19'].fill = Styles.header_fill()
    ws['B19'].border = Styles.THIN_BORDER
    
    ws['C19'] = "Monthly Savings"
    ws['C19'].font = Styles.header_font()
    ws['C19'].fill = Styles.header_fill()
    ws['C19'].border = Styles.THIN_BORDER
    ws['C19'].alignment = Alignment(horizontal='right')
    
    ws['D19'] = "Cumulative Savings"
    ws['D19'].font = Styles.header_font()
    ws['D19'].fill = Styles.header_fill()
    ws['D19'].border = Styles.THIN_BORDER
    ws['D19'].alignment = Alignment(horizontal='right')
    
    ws['E19'] = "% of Target"
    ws['E19'].font = Styles.header_font()
    ws['E19'].fill = Styles.header_fill()
    ws['E19'].border = Styles.THIN_BORDER
    ws['E19'].alignment = Alignment(horizontal='right')
    ws.row_dimensions[19].height = 22
    
    row = 20
    months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
    for i, month in enumerate(months):
        col = get_column_letter(3 + i)
        
        ws[f'B{row}'] = month
        ws[f'B{row}'].font = Font(size=10, color=Styles.TEXT_DARK)
        ws[f'B{row}'].border = Styles.THIN_BORDER
        
        ws[f'C{row}'] = f"='Monthly Entry'!{col}{MonthlyEntryRows.SAVINGS_TRANSFER}"
        ws[f'C{row}'].font = Styles.calc_font()
        ws[f'C{row}'].border = Styles.THIN_BORDER
        ws[f'C{row}'].alignment = Alignment(horizontal='right')
        ws[f'C{row}'].number_format = Config.CURRENCY_FORMAT
        
        ws[f'D{row}'] = f"='Monthly Entry'!{col}{MonthlyEntryRows.RUNNING_SAVINGS}"
        ws[f'D{row}'].font = Styles.calc_font()
        ws[f'D{row}'].border = Styles.THIN_BORDER
        ws[f'D{row}'].alignment = Alignment(horizontal='right')
        ws[f'D{row}'].number_format = Config.CURRENCY_FORMAT
        
        ws[f'E{row}'] = f"=IF($C$8=0,0,D{row}/$C$8)"
        ws[f'E{row}'].font = Styles.calc_font()
        ws[f'E{row}'].border = Styles.THIN_BORDER
        ws[f'E{row}'].alignment = Alignment(horizontal='right')
        ws[f'E{row}'].number_format = Config.PERCENT_FORMAT
        
        if row % 2 == 0:
            for c in ['B', 'C', 'D', 'E']:
                ws[f'{c}{row}'].fill = Styles.alt_row_fill()
        row += 1
    
    row += 2
    ws[f'B{row}'] = "GUIDELINES"
    ws[f'B{row}'].font = Font(size=12, bold=True, color=Styles.HEADER_DARK)
    
    guidelines = [
        "• Emergency fund should cover 6 months of essential expenses",
        "• Keep emergency fund in a separate, easily accessible savings account",
        "• Only use for true emergencies (job loss, medical, major repairs)",
        "• Replenish immediately after any withdrawal",
    ]
    
    row += 2
    for guideline in guidelines:
        ws[f'B{row}'] = guideline
        ws[f'B{row}'].font = Font(size=10, color="555555")
        ws.merge_cells(f'B{row}:E{row}')
        row += 1
    
    return ws


def create_five_year_projection(wb):
    """Create the 5-Year Projection sheet."""
    ws = wb.create_sheet("5-Year Projection")
    ws.sheet_view.showGridLines = False
    
    ws.column_dimensions['A'].width = 3
    ws.column_dimensions['B'].width = 30
    ws.column_dimensions['C'].width = 18
    ws.column_dimensions['D'].width = 18
    ws.column_dimensions['E'].width = 18
    ws.column_dimensions['F'].width = 18
    
    ws.merge_cells('B2:F2')
    ws['B2'] = "5-YEAR SAVINGS PROJECTION"
    ws['B2'].font = Font(size=18, bold=True, color=Styles.HEADER_DARK)
    ws['B2'].alignment = Alignment(horizontal='left', vertical='center')
    ws.row_dimensions[2].height = 30
    
    ws.merge_cells('B3:F3')
    ws['B3'] = "Based on actual savings entered | 8% annual growth assumption"
    ws['B3'].font = Font(size=10, italic=True, color="666666")
    
    ws['B5'] = "PROJECTION ASSUMPTIONS"
    ws['B5'].font = Font(size=12, bold=True, color=Styles.HEADER_DARK)
    
    ws['B6'] = "Annual Growth Rate"
    ws['B6'].font = Font(size=10, color=Styles.TEXT_DARK)
    ws['B6'].border = Styles.THIN_BORDER
    
    ws['C6'] = Config.ANNUAL_GROWTH_RATE
    ws['C6'].font = Styles.input_font()
    ws['C6'].border = Styles.THIN_BORDER
    ws['C6'].alignment = Alignment(horizontal='right')
    ws['C6'].number_format = Config.PERCENT_FORMAT
    
    ws['B7'] = "Monthly Savings (avg from Year 1)"
    ws['B7'].font = Font(size=10, color=Styles.TEXT_DARK)
    ws['B7'].border = Styles.THIN_BORDER
    ws['B7'].fill = Styles.alt_row_fill()
    
    ws['C7'] = "='Monthly Entry'!O32/12"
    ws['C7'].font = Styles.calc_font()
    ws['C7'].border = Styles.THIN_BORDER
    ws['C7'].alignment = Alignment(horizontal='right')
    ws['C7'].number_format = Config.CURRENCY_FORMAT
    ws['C7'].fill = Styles.alt_row_fill()
    
    ws['B8'] = "Starting Balance (Year-End)"
    ws['B8'].font = Font(size=10, color=Styles.TEXT_DARK)
    ws['B8'].border = Styles.THIN_BORDER
    
    ws['C8'] = "='Monthly Entry'!O48"
    ws['C8'].font = Styles.calc_font()
    ws['C8'].border = Styles.THIN_BORDER
    ws['C8'].alignment = Alignment(horizontal='right')
    ws['C8'].number_format = Config.CURRENCY_FORMAT
    
    ws['B11'] = "YEAR-BY-YEAR PROJECTION"
    ws['B11'].font = Font(size=14, bold=True, color=Styles.HEADER_DARK)
    ws.row_dimensions[11].height = 25
    
    ws['B13'] = "Year"
    ws['B13'].font = Styles.header_font()
    ws['B13'].fill = Styles.header_fill()
    ws['B13'].border = Styles.THIN_BORDER
    ws['B13'].alignment = Alignment(horizontal='left', vertical='center')
    
    ws['C13'] = "Starting Balance"
    ws['C13'].font = Styles.header_font()
    ws['C13'].fill = Styles.header_fill()
    ws['C13'].border = Styles.THIN_BORDER
    ws['C13'].alignment = Alignment(horizontal='right', vertical='center')
    
    ws['D13'] = "Annual Contribution"
    ws['D13'].font = Styles.header_font()
    ws['D13'].fill = Styles.header_fill()
    ws['D13'].border = Styles.THIN_BORDER
    ws['D13'].alignment = Alignment(horizontal='right', vertical='center')
    
    ws['E13'] = "Growth (8%)"
    ws['E13'].font = Styles.header_font()
    ws['E13'].fill = Styles.header_fill()
    ws['E13'].border = Styles.THIN_BORDER
    ws['E13'].alignment = Alignment(horizontal='right', vertical='center')
    
    ws['F13'] = "Ending Balance"
    ws['F13'].font = Styles.header_font()
    ws['F13'].fill = Styles.header_fill()
    ws['F13'].border = Styles.THIN_BORDER
    ws['F13'].alignment = Alignment(horizontal='right', vertical='center')
    ws.row_dimensions[13].height = 22
    
    # Year 1 (Current - from actual data)
    row = 14
    ws[f'B{row}'] = "Year 1 (Current)"
    ws[f'B{row}'].font = Font(size=10, bold=True, color=Styles.TEXT_DARK)
    ws[f'B{row}'].border = Styles.THIN_BORDER
    
    ws[f'C{row}'] = "='Monthly Entry'!C33"
    ws[f'C{row}'].font = Styles.calc_font()
    ws[f'C{row}'].border = Styles.THIN_BORDER
    ws[f'C{row}'].alignment = Alignment(horizontal='right')
    ws[f'C{row}'].number_format = Config.CURRENCY_FORMAT
    
    ws[f'D{row}'] = "='Monthly Entry'!O32"
    ws[f'D{row}'].font = Styles.calc_font()
    ws[f'D{row}'].border = Styles.THIN_BORDER
    ws[f'D{row}'].alignment = Alignment(horizontal='right')
    ws[f'D{row}'].number_format = Config.CURRENCY_FORMAT
    
    ws[f'E{row}'] = f"=(C{row}+D{row}/2)*$C$6"
    ws[f'E{row}'].font = Styles.calc_font()
    ws[f'E{row}'].border = Styles.THIN_BORDER
    ws[f'E{row}'].alignment = Alignment(horizontal='right')
    ws[f'E{row}'].number_format = Config.CURRENCY_FORMAT
    
    ws[f'F{row}'] = "='Monthly Entry'!O48"
    ws[f'F{row}'].font = Font(bold=True, size=11, color=Styles.POSITIVE_GREEN)
    ws[f'F{row}'].border = Styles.THIN_BORDER
    ws[f'F{row}'].alignment = Alignment(horizontal='right')
    ws[f'F{row}'].number_format = Config.CURRENCY_FORMAT
    
    # Years 2-5 (Projected)
    for year in range(2, 6):
        row += 1
        ws[f'B{row}'] = f"Year {year} (Projected)"
        ws[f'B{row}'].font = Font(size=10, color=Styles.TEXT_DARK)
        ws[f'B{row}'].border = Styles.THIN_BORDER
        
        ws[f'C{row}'] = f"=F{row-1}"
        ws[f'C{row}'].font = Styles.calc_font()
        ws[f'C{row}'].border = Styles.THIN_BORDER
        ws[f'C{row}'].alignment = Alignment(horizontal='right')
        ws[f'C{row}'].number_format = Config.CURRENCY_FORMAT
        
        ws[f'D{row}'] = "=$C$7*12"
        ws[f'D{row}'].font = Styles.calc_font()
        ws[f'D{row}'].border = Styles.THIN_BORDER
        ws[f'D{row}'].alignment = Alignment(horizontal='right')
        ws[f'D{row}'].number_format = Config.CURRENCY_FORMAT
        
        ws[f'E{row}'] = f"=(C{row}+D{row}/2)*$C$6"
        ws[f'E{row}'].font = Styles.calc_font()
        ws[f'E{row}'].border = Styles.THIN_BORDER
        ws[f'E{row}'].alignment = Alignment(horizontal='right')
        ws[f'E{row}'].number_format = Config.CURRENCY_FORMAT
        
        ws[f'F{row}'] = f"=C{row}+D{row}+E{row}"
        ws[f'F{row}'].font = Font(bold=True, size=11, color=Styles.HEADER_LIGHT)
        ws[f'F{row}'].border = Styles.THIN_BORDER
        ws[f'F{row}'].alignment = Alignment(horizontal='right')
        ws[f'F{row}'].number_format = Config.CURRENCY_FORMAT
        
        if row % 2 == 0:
            for c in ['B', 'C', 'D', 'E', 'F']:
                ws[f'{c}{row}'].fill = Styles.alt_row_fill()
    
    # Total row
    row += 1
    ws[f'B{row}'] = "5-Year Total"
    ws[f'B{row}'].font = Font(bold=True, color=Styles.WHITE)
    ws[f'B{row}'].fill = Styles.header_fill()
    ws[f'B{row}'].border = Styles.THIN_BORDER
    
    ws[f'C{row}'].fill = Styles.header_fill()
    ws[f'C{row}'].border = Styles.THIN_BORDER
    
    ws[f'D{row}'] = f"=SUM(D14:D18)"
    ws[f'D{row}'].font = Font(bold=True, color=Styles.WHITE)
    ws[f'D{row}'].fill = Styles.header_fill()
    ws[f'D{row}'].border = Styles.THIN_BORDER
    ws[f'D{row}'].alignment = Alignment(horizontal='right')
    ws[f'D{row}'].number_format = Config.CURRENCY_FORMAT
    
    ws[f'E{row}'] = f"=SUM(E14:E18)"
    ws[f'E{row}'].font = Font(bold=True, color=Styles.WHITE)
    ws[f'E{row}'].fill = Styles.header_fill()
    ws[f'E{row}'].border = Styles.THIN_BORDER
    ws[f'E{row}'].alignment = Alignment(horizontal='right')
    ws[f'E{row}'].number_format = Config.CURRENCY_FORMAT
    
    ws[f'F{row}'] = f"=F18"
    ws[f'F{row}'].font = Font(bold=True, color=Styles.WHITE)
    ws[f'F{row}'].fill = Styles.header_fill()
    ws[f'F{row}'].border = Styles.THIN_BORDER
    ws[f'F{row}'].alignment = Alignment(horizontal='right')
    ws[f'F{row}'].number_format = Config.CURRENCY_FORMAT
    
    # Summary Metrics
    row = 22
    ws[f'B{row}'] = "PROJECTION SUMMARY"
    ws[f'B{row}'].font = Font(size=14, bold=True, color=Styles.HEADER_DARK)
    
    summary_metrics = [
        ("Starting Balance (Today)", "=C14", Config.CURRENCY_FORMAT),
        ("Projected Balance (Year 5)", "=F18", Config.CURRENCY_FORMAT),
        ("Total Contributions (5 Years)", "=SUM(D14:D18)", Config.CURRENCY_FORMAT),
        ("Total Growth/Earnings (5 Years)", "=SUM(E14:E18)", Config.CURRENCY_FORMAT),
        ("Growth Multiple", "=IF(C14=0,0,F18/C14)", '0.00"x"'),
    ]
    
    row += 2
    for label, formula, fmt in summary_metrics:
        ws[f'B{row}'] = label
        ws[f'B{row}'].font = Font(size=10, color=Styles.TEXT_DARK)
        ws[f'B{row}'].border = Styles.THIN_BORDER
        
        ws[f'C{row}'] = formula
        ws[f'C{row}'].font = Font(size=11, bold=True, color=Styles.HEADER_LIGHT)
        ws[f'C{row}'].border = Styles.THIN_BORDER
        ws[f'C{row}'].alignment = Alignment(horizontal='right')
        ws[f'C{row}'].number_format = fmt
        
        if row % 2 == 0:
            ws[f'B{row}'].fill = Styles.alt_row_fill()
            ws[f'C{row}'].fill = Styles.alt_row_fill()
        row += 1
    
    # Notes
    row += 2
    ws[f'B{row}'] = "IMPORTANT NOTES"
    ws[f'B{row}'].font = Font(size=12, bold=True, color=Styles.HEADER_DARK)
    
    notes = [
        "• This projection assumes consistent monthly savings at your Year 1 average",
        "• 8% annual return is an estimate - actual returns will vary",
        "• Update your Monthly Entry sheet regularly for accurate projections",
        "• Consider increasing savings rate as income grows",
        "• This is a simplified projection - consult a financial advisor for detailed planning",
    ]
    
    row += 2
    for note in notes:
        ws[f'B{row}'] = note
        ws[f'B{row}'].font = Font(size=10, color="555555")
        ws.merge_cells(f'B{row}:F{row}')
        row += 1
    
    return ws


# =============================================================================
# MAIN FUNCTION
# =============================================================================

def create_budget_workbook(output_path=None):
    """
    Create the complete Personal Budget Workbook.
    
    Args:
        output_path: Path to save the workbook. If None, uses Config.OUTPUT_PATH.
    
    Returns:
        Path to the created workbook.
    """
    if output_path is None:
        output_path = Config.OUTPUT_PATH
    
    wb = Workbook()
    
    # Create all sheets
    create_cover_sheet(wb)
    create_monthly_entry_sheet(wb)
    create_summary_dashboard(wb)
    create_trend_analysis(wb)
    create_bonus_tracker(wb)
    create_emergency_fund(wb)
    create_five_year_projection(wb)
    
    # Save the workbook
    wb.save(output_path)
    print(f"Budget workbook created successfully: {output_path}")
    
    return output_path


# =============================================================================
# ENTRY POINT
# =============================================================================

if __name__ == "__main__":
    create_budget_workbook()
