from __future__ import annotations

from openpyxl.utils import get_column_letter


def build_sum_formula(
    column: int,
    row_by_country: dict[str, int],
    members: tuple[str, ...],
) -> str | None:
    """Build an Excel ``SUM`` formula for a configured country group.

    The formula references the intermediary country rows rather than copying a
    Python-computed total. This keeps the aggregation transparent in Excel.

    :param column: One-based intermediary worksheet column.
    :param row_by_country: Mapping from country to intermediary row number.
    :param members: Ordered configured group members.
    :return: Excel formula, or ``None`` for an empty group.
    """

    references = [
        f"{get_column_letter(column)}{row_by_country[country]}"
        for country in members
    ]
    if not references:
        return None
    return f"=SUM({','.join(references)})"
