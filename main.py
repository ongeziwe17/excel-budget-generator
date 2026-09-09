"""CLI entry point for generating personal or couple budget workbooks."""

import argparse

from budget_workbook import (
    BudgetWorkbookGenerator,
    ContributionMethod,
    CoupleConfig,
    PersonConfig,
    WorkbookConfig,
    WorkbookMode,
    WorkbookVersion,
)


def parse_args() -> argparse.Namespace:
    """Parse workbook-generation options."""

    parser = argparse.ArgumentParser(description="Generate an Excel budget workbook.")
    parser.add_argument("--mode", choices=[mode.value for mode in WorkbookMode], default="single")
    parser.add_argument("--partner-one-name", default="Partner 1")
    parser.add_argument("--partner-two-name", default="Partner 2")
    parser.add_argument("--partner-one-gross-income", type=int)
    parser.add_argument("--partner-two-gross-income", type=int)
    parser.add_argument("--partner-one-net-income", type=int, default=20_000)
    parser.add_argument("--partner-two-net-income", type=int, default=20_000)
    parser.add_argument(
        "--contribution-method",
        choices=[method.value for method in ContributionMethod],
        default=ContributionMethod.INCOME_PROPORTIONAL.value,
    )
    parser.add_argument("--partner-one-share", type=float)
    parser.add_argument("--partner-two-share", type=float)
    parser.add_argument("--output", help="Optional output .xlsx path.")
    return parser.parse_args()


def build_config(args: argparse.Namespace) -> WorkbookConfig:
    """Create validated workbook configuration from CLI arguments."""

    mode = WorkbookMode(args.mode)
    if mode is WorkbookMode.SINGLE:
        return WorkbookConfig()
    couple = CoupleConfig(
        partner_one=PersonConfig(
            name=args.partner_one_name,
            monthly_gross_income=args.partner_one_gross_income or args.partner_one_net_income,
            monthly_net_income=args.partner_one_net_income,
            custom_contribution_share=args.partner_one_share,
        ),
        partner_two=PersonConfig(
            name=args.partner_two_name,
            monthly_gross_income=args.partner_two_gross_income or args.partner_two_net_income,
            monthly_net_income=args.partner_two_net_income,
            custom_contribution_share=args.partner_two_share,
        ),
        contribution_method=ContributionMethod(args.contribution_method),
    )
    return WorkbookConfig(
        mode=mode,
        couple=couple,
        workbook_name_prefix="Household_Budget_Workbook",
        workbook_version=WorkbookVersion(0, 1, 0),
    )


def main() -> None:
    """Generate the requested budget workbook."""
    args = parse_args()
    try:
        config = build_config(args)
    except ValueError as exc:
        raise SystemExit(f"Configuration error: {exc}") from None
    output_path = BudgetWorkbookGenerator(config).create_workbook(args.output)
    print(f"Budget workbook created successfully: {output_path}")


if __name__ == "__main__":
    main()
