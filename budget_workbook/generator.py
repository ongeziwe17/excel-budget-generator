"""Application service for budget workbook generation."""

from pathlib import Path

from openpyxl import Workbook

from budget_workbook.builders.bonus_tracker import BonusTrackerSheetBuilder
from budget_workbook.builders.cover import CoverSheetBuilder
from budget_workbook.builders.emergency_fund import EmergencyFundSheetBuilder
from budget_workbook.builders.five_year_projection import FiveYearProjectionSheetBuilder
from budget_workbook.builders.monthly_entry import MonthlyEntrySheetBuilder
from budget_workbook.builders.summary_dashboard import SummaryDashboardSheetBuilder
from budget_workbook.builders.trend_analysis import TrendAnalysisSheetBuilder
from budget_workbook.config import WorkbookConfig
from budget_workbook.rows import MonthlyEntryRows
from budget_workbook.styles import WorkbookStyles


class BudgetWorkbookGenerator:
    """Create the complete personal budget workbook."""

    def __init__(self, config: WorkbookConfig | None = None) -> None:
        self.config = config or WorkbookConfig()
        self.styles = WorkbookStyles()
        self.rows = MonthlyEntryRows()
        self.sheet_builders = [
            CoverSheetBuilder(self.config, self.styles, self.rows),
            MonthlyEntrySheetBuilder(self.config, self.styles, self.rows),
            SummaryDashboardSheetBuilder(self.config, self.styles, self.rows),
            TrendAnalysisSheetBuilder(self.config, self.styles, self.rows),
            BonusTrackerSheetBuilder(self.config, self.styles, self.rows),
            EmergencyFundSheetBuilder(self.config, self.styles, self.rows),
            FiveYearProjectionSheetBuilder(self.config, self.styles, self.rows),
        ]

    def create_workbook(self, output_path: str | None = None) -> Path:
        """Generate the workbook and return the saved file path."""
        target_path = Path(output_path or self.config.output_path)
        workbook = Workbook()

        for builder in self.sheet_builders:
            builder.build(workbook)

        workbook.save(target_path)
        return target_path
