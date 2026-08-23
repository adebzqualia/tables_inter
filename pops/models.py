from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True)
class ValidationIssue:
    """One recoverable validation or extraction issue.

    :param country: Country/entity, if applicable.
    :param source_sheet: Physical source worksheet name, if applicable.
    :param kpi: KPI display label, if applicable.
    :param period: Period label, if applicable.
    :param issue: Stable issue code.
    :param details: Human-readable details including coordinates where useful.
    """

    country: str | None
    source_sheet: str | None
    kpi: str | None
    period: str | None
    issue: str
    details: str


@dataclass(frozen=True)
class ExtractionRecord:
    """One country/KPI/period extraction result.

    ``kpi_index`` is the KPI position in configuration. It makes duplicate KPI
    names unambiguous without forcing operational users to maintain technical IDs.

    :param country: Country/entity.
    :param source: Logical source-sheet name.
    :param source_sheet: Physical Excel worksheet name.
    :param kpi_index: Zero-based KPI position in source configuration.
    :param kpi: Raw KPI label searched in Excel.
    :param kpi_display: Display label used in generated outputs.
    :param period: Period label.
    :param value: Last cached numeric value, or ``None`` if invalid/missing.
    :param coordinate: Resolved source cell coordinate, if available.
    """

    country: str
    source: str
    source_sheet: str
    kpi_index: int
    kpi: str
    kpi_display: str
    period: str
    value: int | float | None
    coordinate: str | None
