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
    rent_housing: int = 17
    groceries: int = 20
    eating_out: int = 21
    lunch: int = 22
    haircuts: int = 23
    clothing: int = 24
    random_spending: int = 25
    custom_variable_total: int = 26
    savings_transfer: int = 29
    starting_savings: int = 30
    bonus: int = 33
    unexpected_expenses: int = 36
    total_fixed: int = 39
    total_variable: int = 40
    total_expenses: int = 41
    total_income: int = 42
    surplus_deficit: int = 43
    savings_rate: int = 44
    running_savings: int = 45
