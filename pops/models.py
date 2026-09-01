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
    :param number_format: Source Excel number format, if a cell was resolved.
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
    number_format: str | None = None

@dataclass(frozen=True)
class ResolvedRatioTotal:
    """Resolved group-total rule for one ratio KPI.

    The rule may come from explicit YAML configuration or from strict formula
    inference against the source workbook.

    :param numerator_index: Configured additive KPI index used as numerator.
    :param denominator_index: Configured additive KPI index used as denominator.
    :param multiplier: Numeric scale applied after division.
    :param percent: Whether the result is stored/displayed as an Excel percentage.
    :param origin: ``configured`` or ``inferred``.
    :param evidence_countries: Countries that confirmed an inferred rule.
    """

    numerator_index: int
    denominator_index: int
    multiplier: float = 1.0
    percent: bool = True
    origin: str = "configured"
    evidence_countries: tuple[str, ...] = ()



@dataclass(frozen=True)
class LineageNode:
    """One generated KPI node in the recursive dependency graph.

    :param node_id: Stable internal node identifier.
    :param source: Logical source sheet containing the KPI.
    :param name: KPI name.
    :param occurrence: One-based occurrence for duplicate names.
    :param display_name: Human-readable intermediary title.
    :param configured: Whether the node comes from ``kpis.yaml``.
    :param configured_index: Configured KPI index when applicable.
    :param default_aggregation: ``sum``, ``average``, ``formula`` or ``unresolved``.
    :param percent: Whether the KPI is formatted as a percentage in source cells.
    """

    node_id: str
    source: str
    name: str
    occurrence: int
    display_name: str
    configured: bool
    configured_index: int | None
    default_aggregation: str
    percent: bool = False


@dataclass(frozen=True)
class LineageRecord:
    """One country/period source location for an automatically discovered KPI.

    :param node_id: Dependency node identifier.
    :param country: Country/entity.
    :param period: Configured period.
    :param source_sheet: Physical source worksheet.
    :param coordinate: Source coordinate, if resolved.
    :param number_format: Source number format.
    :param formula: Source formula expression, if present.
    :param value: Last cached value, if numeric.
    """

    node_id: str
    country: str
    period: str
    source_sheet: str
    coordinate: str | None
    number_format: str | None = None
    formula: str | None = None
    value: int | float | None = None


@dataclass(frozen=True)
class LineagePlan:
    """Recursive formula-lineage plan used when writing group totals.

    :param nodes: Automatically discovered dependency nodes.
    :param records: Country/period source locations for dependency nodes.
    :param formulas: Symbolic formulas keyed by ``(node_id, period)``.
    :param configured_formulas: Symbolic formulas for configured derived KPIs,
        keyed by ``(source, kpi_index, period)``.
    :param placeholder_targets: Mapping from symbolic placeholder token to
        ``(node_id, period)``.
    :param dependency_order: Auto dependency node IDs in dependency-first order.
    """

    nodes: tuple[LineageNode, ...]
    records: tuple[LineageRecord, ...]
    formulas: dict[tuple[str, str], str]
    configured_formulas: dict[tuple[str, int, str], str]
    placeholder_targets: dict[str, tuple[str, str]]
    dependency_order: tuple[str, ...]
