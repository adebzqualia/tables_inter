from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True)
class ValidationIssue:
    """One recoverable validation or extraction issue.

    :param country: Country/entity, if applicable.
    :param source_sheet: Physical source worksheet name, if applicable.
    :param kpi: KPI label, if applicable.
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

    :param country: Country/entity.
    :param source: Logical source-sheet name.
    :param kpi: KPI label.
    :param period: Period label.
    :param value: Valid numeric value, or ``None`` if extraction failed/missing.
    :param coordinate: Source cell coordinate when a cell was resolved.
    """

    country: str
    source: str
    kpi: str
    period: str
    value: int | float | None
    coordinate: str | None


@dataclass(frozen=True)
class GroupTotal:
    """One configured country-group total.

    :param source: Logical source-sheet name.
    :param kpi: KPI label.
    :param group: Country-group name.
    :param period: Period label.
    :param value: Sum of valid numeric country observations, or ``None`` if none exist.
    """

    source: str
    kpi: str
    group: str
    period: str
    value: int | float | None
