"""Row references used across workbook formulas."""

from dataclasses import dataclass


@dataclass(frozen=True)
class MonthlyEntryRows:
    """Key row numbers in the Monthly Entry sheet."""

    net_income: int = 8
    data: int = 11
    transport: int = 12
    subscriptions: int = 13
    tfg_debit: int = 14
    family_support: int = 15
    tithe: int = 16
    groceries: int = 19
    eating_out: int = 20
    lunch: int = 21
    haircuts: int = 22
    clothing: int = 23
    random_spending: int = 24
    rent_toggle: int = 27
    rent_amount: int = 28
    rent_paid: int = 29
    savings_transfer: int = 32
    starting_savings: int = 33
    bonus: int = 36
    unexpected_expenses: int = 39
    total_fixed: int = 42
    total_variable: int = 43
    total_expenses: int = 44
    total_income: int = 45
    surplus_deficit: int = 46
    savings_rate: int = 47
    running_savings: int = 48
