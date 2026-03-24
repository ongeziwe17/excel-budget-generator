"""Configuration models for the budget workbook generator."""

from dataclasses import dataclass, field
from pathlib import Path

from .versioning import WorkbookVersion


@dataclass(frozen=True)
class WorkbookConfig:
    """Configuration values used throughout workbook generation."""

    currency_format: str = "R #,##0.00"
    percent_format: str = "0.0%"
    monthly_gross_salary: int = 25_000
    monthly_net_income: int = 20_339
    bonus_q2_pct: float = 0.25
    bonus_q3_pct: float = 0.50
    bonus_q4_pct: float = 0.83
    default_rent_amount: int = 8_000
    annual_growth_rate: float = 0.08
    workbook_name_prefix: str = "Personal_Budget_Workbook"
    workbook_version: WorkbookVersion = field(default_factory=WorkbookVersion)
    output_dir: Path = field(
        default_factory=lambda: Path(__file__).resolve().parent.parent / "workbooks"
    )
    months: tuple[str, ...] = field(
        default_factory=lambda: (
            "Jan",
            "Feb",
            "Mar",
            "Apr",
            "May",
            "Jun",
            "Jul",
            "Aug",
            "Sep",
            "Oct",
            "Nov",
            "Dec",
        )
    )

    @property
    def output_filename(self) -> str:
        return f"{self.workbook_name_prefix}_{self.workbook_version}.xlsx"

    @property
    def output_path(self) -> Path:
        return self.output_dir / self.output_filename
