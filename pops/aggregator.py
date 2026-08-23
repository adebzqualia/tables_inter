from __future__ import annotations

from openpyxl.utils import get_column_letter


def build_sum_formula(
    column: int,
    row_by_country: dict[str, int],
    members: tuple[str, ...],
    round_values: bool = False,
    round_digits: int = 0,
) -> str | None:
    """Build an Excel ``SUM`` formula for a configured country group.

    The formula references intermediary country rows rather than copying a
    Python-computed total.

    :param column: One-based intermediary worksheet column.
    :param row_by_country: Mapping from country to intermediary row number.
    :param members: Ordered configured group members.
    :param round_values: Whether to round the generated total.
    :param round_digits: Decimal digits used when rounding.
    :return: Excel formula, or ``None`` for an empty group.
    """

    references = [
        f"{get_column_letter(column)}{row_by_country[country]}"
        for country in members
    ]
    if not references:
        return None

    expression = f"SUM({','.join(references)})"
    if round_values:
        expression = f"ROUND({expression},{round_digits})"
    return f"={expression}"


def build_ratio_formula(
    numerator_ref: str,
    denominator_ref: str,
    multiplier: float = 1.0,
    percent: bool = False,
    round_values: bool = False,
    round_digits: int = 0,
) -> str:
    """Build a guarded Excel formula for a simple group-level ratio.

    :param numerator_ref: Intermediary total-cell reference for the numerator.
    :param denominator_ref: Intermediary total-cell reference for the denominator.
    :param multiplier: Optional scale applied after division.
    :param percent: Whether the result is stored as an Excel percentage decimal.
    :param round_values: Whether to round the ratio result.
    :param round_digits: Decimal digits used when rounding.
    :return: Excel formula.
    """

    expression = f"({numerator_ref}/{denominator_ref})"
    if multiplier != 1.0:
        expression = f"({expression}*{multiplier:g})"

    if round_values:
        if percent:
            expression = f"(ROUND(({expression})*100,{round_digits})/100)"
        else:
            expression = f"ROUND({expression},{round_digits})"

    return f'=IFERROR({expression},"")'
