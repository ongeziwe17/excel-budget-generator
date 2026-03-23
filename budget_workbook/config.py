"""Configuration models for the budget workbook generator."""

from dataclasses import dataclass, field


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
    output_path: str = "Personal_Budget_Workbook.xlsx"
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
