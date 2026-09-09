"""Configuration models for the budget workbook generator."""

from dataclasses import dataclass, field
from enum import Enum
from pathlib import Path

from .versioning import WorkbookVersion


class WorkbookMode(str, Enum):
    """Supported workbook layouts."""

    SINGLE = "single"
    COUPLE = "couple"


class ContributionMethod(str, Enum):
    """Ways a couple can split shared household costs."""

    EQUAL = "Equal"
    INCOME_PROPORTIONAL = "Income proportional"
    CUSTOM = "Custom"


@dataclass(frozen=True)
class PersonConfig:
    """Configuration for one member of a couple."""

    name: str
    monthly_gross_income: int = 25_000
    monthly_net_income: int = 20_000
    opening_savings_balance: int = 0
    custom_contribution_share: float | None = None

    def __post_init__(self) -> None:
        if not self.name.strip():
            raise ValueError("A partner name cannot be blank.")
        for field_name in (
            "monthly_gross_income",
            "monthly_net_income",
            "opening_savings_balance",
        ):
            if getattr(self, field_name) < 0:
                raise ValueError(f"{field_name} cannot be negative.")
        if self.monthly_net_income > self.monthly_gross_income:
            raise ValueError("monthly_net_income cannot exceed monthly_gross_income.")
        if self.custom_contribution_share is not None and not 0 <= self.custom_contribution_share <= 1:
            raise ValueError("custom_contribution_share must be between 0 and 1.")


@dataclass(frozen=True)
class CoupleConfig:
    """Household settings for a two-person budget."""

    partner_one: PersonConfig = field(default_factory=lambda: PersonConfig(name="Partner 1"))
    partner_two: PersonConfig = field(default_factory=lambda: PersonConfig(name="Partner 2"))
    contribution_method: ContributionMethod = ContributionMethod.INCOME_PROPORTIONAL
    opening_joint_savings_balance: int = 0
    emergency_fund_months: int = 6

    def __post_init__(self) -> None:
        if self.partner_one.name.strip().casefold() == self.partner_two.name.strip().casefold():
            raise ValueError("Partner names must be different.")
        if self.opening_joint_savings_balance < 0:
            raise ValueError("opening_joint_savings_balance cannot be negative.")
        if self.emergency_fund_months <= 0:
            raise ValueError("emergency_fund_months must be greater than zero.")
        if self.contribution_method is ContributionMethod.CUSTOM:
            shares = (
                self.partner_one.custom_contribution_share,
                self.partner_two.custom_contribution_share,
            )
            if None in shares or abs(sum(share for share in shares if share is not None) - 1) > 0.000001:
                raise ValueError("Custom contribution shares must be provided and add up to 100%.")


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
    mode: WorkbookMode = WorkbookMode.SINGLE
    couple: CoupleConfig = field(default_factory=CoupleConfig)
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
