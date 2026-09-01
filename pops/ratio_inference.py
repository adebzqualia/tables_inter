from __future__ import annotations

import re
from collections import Counter, defaultdict
from dataclasses import dataclass

from openpyxl.formula import Tokenizer

from .config import AppConfig, KPIConfig, resolve_ratio_total_indices
from .models import ExtractionRecord, ResolvedRatioTotal, ValidationIssue

_CELL_RE = re.compile(r"^\$?([A-Za-z]{1,3})\$?(\d+)$")
_ALLOWED_FUNCTIONS = {"IF", "IFERROR", "ROUND"}
_ARITHMETIC_OPERATORS = {"+", "-", "*", "/", "^"}


@dataclass(frozen=True)
class _FormulaRatio:
    """Simple cell-over-cell relation parsed from one Excel formula.

    :param numerator_sheet: Explicit numerator sheet name, or ``None``.
    :param numerator_coordinate: Normalized numerator cell coordinate.
    :param denominator_sheet: Explicit denominator sheet name, or ``None``.
    :param denominator_coordinate: Normalized denominator cell coordinate.
    :param multiplier: Numeric multiplier applied to the direct ratio.
    """

    numerator_sheet: str | None
    numerator_coordinate: str
    denominator_sheet: str | None
    denominator_coordinate: str
    multiplier: float = 1.0


def _next_non_space(tokens, index: int, step: int) -> int | None:
    """Return the next non-space token index.

    :param tokens: Formula tokens.
    :param index: Starting token index.
    :param step: ``1`` for right or ``-1`` for left.
    :return: Token index, or ``None`` when exhausted.
    """

    index += step
    while 0 <= index < len(tokens):
        if tokens[index].type != "WSPACE":
            return index
        index += step
    return None


def _range_operand(tokens, slash_index: int, step: int) -> tuple[int, int, str] | None:
    """Resolve a direct range operand adjacent to a division operator.

    One pair of simple parentheses around the cell reference is accepted.

    :param tokens: Formula tokens.
    :param slash_index: Index of the division operator.
    :param step: ``-1`` for numerator or ``1`` for denominator.
    :return: ``(start_index, end_index, token_value)`` or ``None``.
    """

    index = _next_non_space(tokens, slash_index, step)
    if index is None:
        return None

    token = tokens[index]
    if token.type == "OPERAND" and token.subtype == "RANGE":
        return index, index, token.value

    expected_outer = "CLOSE" if step < 0 else "OPEN"
    expected_inner = "OPEN" if step < 0 else "CLOSE"
    if token.type != "PAREN" or token.subtype != expected_outer:
        return None

    operand_index = _next_non_space(tokens, index, step)
    if operand_index is None:
        return None
    operand = tokens[operand_index]
    if operand.type != "OPERAND" or operand.subtype != "RANGE":
        return None

    other_paren_index = _next_non_space(tokens, operand_index, step)
    if other_paren_index is None:
        return None
    other_paren = tokens[other_paren_index]
    if other_paren.type != "PAREN" or other_paren.subtype != expected_inner:
        return None

    start = min(index, other_paren_index)
    end = max(index, other_paren_index)
    return start, end, operand.value


def _cell_reference(value: str) -> tuple[str | None, str] | None:
    """Parse one direct A1 cell reference.

    External workbooks, ranges, named ranges, and structured references are
    intentionally rejected.

    :param value: ``OPERAND/RANGE`` token value.
    :return: Optional sheet name and normalized coordinate, or ``None``.
    """

    if "[" in value or "]" in value or ":" in value or "," in value:
        return None

    sheet_name: str | None = None
    coordinate = value
    if "!" in value:
        sheet_part, coordinate = value.rsplit("!", 1)
        if sheet_part.startswith("'") and sheet_part.endswith("'"):
            sheet_part = sheet_part[1:-1].replace("''", "'")
        sheet_name = sheet_part

    match = _CELL_RE.fullmatch(coordinate)
    if match is None:
        return None
    return sheet_name, f"{match.group(1).upper()}{int(match.group(2))}"


def _numeric_token(token) -> float | None:
    """Return a numeric token value when possible.

    :param token: Formula token.
    :return: Parsed number or ``None``.
    """

    if token.type != "OPERAND" or token.subtype != "NUMBER":
        return None
    try:
        return float(token.value)
    except ValueError:
        return None


def _parse_simple_ratio_formula(formula: object) -> _FormulaRatio | None:
    """Parse a strict simple ``cell / cell`` Excel formula.

    Common guards/wrappers such as ``IF``, ``IFERROR`` and ``ROUND`` are allowed.
    A numeric multiplier immediately before or after the ratio is also supported.
    Formulas with multiple divisions or other arithmetic are rejected.

    :param formula: Raw Excel formula value.
    :return: Parsed relation or ``None`` when the formula is not safely simple.
    """

    if not isinstance(formula, str) or not formula.startswith("="):
        return None

    try:
        tokens = [token for token in Tokenizer(formula).items if token.type != "WSPACE"]
    except Exception:
        return None

    for token in tokens:
        if token.type == "FUNC" and token.subtype == "OPEN":
            function_name = token.value[:-1].upper()
            if function_name not in _ALLOWED_FUNCTIONS:
                return None
        if token.type == "OPERATOR-PREFIX" and token.value in {"+", "-"}:
            return None

    slash_indices = [
        index
        for index, token in enumerate(tokens)
        if token.type == "OPERATOR-INFIX" and token.value == "/"
    ]
    if len(slash_indices) != 1:
        return None

    slash_index = slash_indices[0]
    numerator = _range_operand(tokens, slash_index, -1)
    denominator = _range_operand(tokens, slash_index, 1)
    if numerator is None or denominator is None:
        return None

    numerator_start, _numerator_end, numerator_value = numerator
    _denominator_start, denominator_end, denominator_value = denominator
    numerator_ref = _cell_reference(numerator_value)
    denominator_ref = _cell_reference(denominator_value)
    if numerator_ref is None or denominator_ref is None:
        return None

    consumed_operators = {slash_index}
    multiplier = 1.0

    # Accept ``100 * numerator / denominator``.
    before_numerator = numerator_start - 1
    if before_numerator >= 1:
        op = tokens[before_numerator]
        number = _numeric_token(tokens[before_numerator - 1])
        if op.type == "OPERATOR-INFIX" and op.value == "*" and number is not None:
            multiplier *= number
            consumed_operators.add(before_numerator)

    # Accept ``numerator / denominator * 100``.
    after_denominator = denominator_end + 1
    if after_denominator + 1 < len(tokens):
        op = tokens[after_denominator]
        number = _numeric_token(tokens[after_denominator + 1])
        if op.type == "OPERATOR-INFIX" and op.value == "*" and number is not None:
            multiplier *= number
            consumed_operators.add(after_denominator)

    for index, token in enumerate(tokens):
        if (
            token.type == "OPERATOR-INFIX"
            and token.value in _ARITHMETIC_OPERATORS
            and index not in consumed_operators
        ):
            return None

    return _FormulaRatio(
        numerator_sheet=numerator_ref[0],
        numerator_coordinate=numerator_ref[1],
        denominator_sheet=denominator_ref[0],
        denominator_coordinate=denominator_ref[1],
        multiplier=multiplier,
    )


def _same_sheet(explicit_sheet: str | None, source_sheet: str) -> bool:
    """Check whether a formula reference stays on the source worksheet.

    :param explicit_sheet: Parsed sheet name, or ``None`` for a local reference.
    :param source_sheet: Current physical source-sheet name.
    :return: Whether the reference is local to the same sheet.
    """

    return explicit_sheet is None or explicit_sheet.casefold() == source_sheet.casefold()


def _configured_rule(
    source_name: str,
    kpis: tuple[KPIConfig, ...],
    kpi: KPIConfig,
) -> ResolvedRatioTotal:
    """Convert an explicit YAML ratio rule to resolved KPI indexes.

    :param source_name: Logical source sheet.
    :param kpis: Configured KPIs for the source.
    :param kpi: Ratio KPI owning the rule.
    :return: Resolved rule.
    """

    indices = resolve_ratio_total_indices(kpis, kpi, source_name)
    if indices is None or kpi.ratio_total is None:
        raise ValueError("Configured ratio rule unexpectedly resolved to None")
    numerator_index, denominator_index = indices
    return ResolvedRatioTotal(
        numerator_index=numerator_index,
        denominator_index=denominator_index,
        multiplier=kpi.ratio_total.multiplier,
        percent=kpi.ratio_total.percent,
        origin="configured",
    )


def _infer_one_ratio(
    formulas_wb,
    config: AppConfig,
    source_name: str,
    kpis: tuple[KPIConfig, ...],
    ratio_kpi: KPIConfig,
    records: list[ExtractionRecord],
) -> tuple[ResolvedRatioTotal | None, ValidationIssue]:
    """Infer one missing ratio-total rule from source formulas.

    A country confirms a relationship only when all successfully parsed periods
    for that country map to the same additive numerator/denominator pair and
    multiplier. At least two countries must agree when multiple countries are
    configured.

    :param formulas_wb: Formula-preserving workbook.
    :param config: Complete application configuration.
    :param source_name: Logical source sheet.
    :param kpis: KPIs configured for the source.
    :param ratio_kpi: Ratio KPI without an explicit rule.
    :param records: Extraction records.
    :return: Optional inferred rule and one summary diagnostic.
    """

    relevant = [
        record
        for record in records
        if record.source == source_name
        and record.kpi_index == ratio_kpi.index
        and record.coordinate is not None
    ]

    coordinate_lookup = {
        (record.country, record.period, record.coordinate.replace("$", "").upper()): record
        for record in records
        if record.source == source_name and record.coordinate is not None
    }

    country_relations: dict[str, set[tuple[int, int, float]]] = defaultdict(set)
    failure_counts: Counter[str] = Counter()
    parsed_cells = 0

    for record in relevant:
        if record.source_sheet not in formulas_wb.sheetnames:
            failure_counts["source sheet unavailable"] += 1
            continue

        formula_cell = formulas_wb[record.source_sheet][record.coordinate]
        relation = _parse_simple_ratio_formula(formula_cell.value)
        if relation is None:
            failure_counts["not a supported simple cell/cell formula"] += 1
            continue

        if not _same_sheet(relation.numerator_sheet, record.source_sheet) or not _same_sheet(
            relation.denominator_sheet, record.source_sheet
        ):
            failure_counts["formula references another worksheet"] += 1
            continue

        numerator_record = coordinate_lookup.get(
            (record.country, record.period, relation.numerator_coordinate)
        )
        denominator_record = coordinate_lookup.get(
            (record.country, record.period, relation.denominator_coordinate)
        )
        if numerator_record is None or denominator_record is None:
            failure_counts["formula cells do not map to configured KPI intersections"] += 1
            continue

        numerator_kpi = kpis[numerator_record.kpi_index]
        denominator_kpi = kpis[denominator_record.kpi_index]
        if numerator_kpi.aggregation != "sum" or denominator_kpi.aggregation != "sum":
            failure_counts["numerator or denominator is not configured as sum"] += 1
            continue

        parsed_cells += 1
        country_relations[record.country].add(
            (
                numerator_record.kpi_index,
                denominator_record.kpi_index,
                relation.multiplier,
            )
        )

    conflicting_countries = [
        country for country, relations in country_relations.items() if len(relations) > 1
    ]
    if conflicting_countries:
        return None, ValidationIssue(
            country=None,
            source_sheet=source_name,
            kpi=ratio_kpi.display_name(config.workbook.kpi_title_separator),
            period=None,
            issue="RATIO_TOTAL_INFERENCE_CONFLICT",
            details=(
                "Source formulas imply different numerator/denominator relationships "
                f"within: {', '.join(conflicting_countries)}. No group-total rule was inferred."
            ),
        )

    country_evidence = {
        country: next(iter(relations))
        for country, relations in country_relations.items()
        if len(relations) == 1
    }
    distinct_relations = set(country_evidence.values())
    minimum_countries = 1 if len(config.countries.countries) == 1 else 2

    if len(distinct_relations) != 1 or len(country_evidence) < minimum_countries:
        reason_parts: list[str] = []
        if len(distinct_relations) > 1:
            reason_parts.append("different relationships were found across countries")
        elif not distinct_relations:
            reason_parts.append("no source formula could be mapped safely")
        else:
            reason_parts.append(
                f"only {len(country_evidence)} country confirmed the relationship; "
                f"at least {minimum_countries} are required"
            )
        if failure_counts:
            top_failures = ", ".join(
                f"{reason}: {count}" for reason, count in failure_counts.most_common(3)
            )
            reason_parts.append(f"formula checks: {top_failures}")

        return None, ValidationIssue(
            country=None,
            source_sheet=source_name,
            kpi=ratio_kpi.display_name(config.workbook.kpi_title_separator),
            period=None,
            issue="RATIO_TOTAL_INFERENCE_FAILED",
            details="; ".join(reason_parts) + ". TOP8/TOP9/ALL remain blank.",
        )

    numerator_index, denominator_index, multiplier = next(iter(distinct_relations))
    numerator = kpis[numerator_index]
    denominator = kpis[denominator_index]
    evidence_countries = tuple(
        country
        for country in config.countries.countries
        if country in country_evidence
    )

    rule = ResolvedRatioTotal(
        numerator_index=numerator_index,
        denominator_index=denominator_index,
        multiplier=multiplier,
        percent=True,
        origin="inferred",
        evidence_countries=evidence_countries,
    )
    issue = ValidationIssue(
        country=None,
        source_sheet=source_name,
        kpi=ratio_kpi.display_name(config.workbook.kpi_title_separator),
        period=None,
        issue="RATIO_TOTAL_INFERRED",
        details=(
            f"Inferred {numerator.display_name(config.workbook.kpi_title_separator)} / "
            f"{denominator.display_name(config.workbook.kpi_title_separator)}"
            + (f" * {multiplier:g}" if multiplier != 1.0 else "")
            + f" from source formulas; confirmed in {len(evidence_countries)}/"
            f"{len(config.countries.countries)} countries across {parsed_cells} formula cell(s)."
        ),
    )
    return rule, issue


def resolve_ratio_total_rules(
    formulas_wb,
    config: AppConfig,
    records: list[ExtractionRecord],
) -> tuple[dict[tuple[str, int], ResolvedRatioTotal], list[ValidationIssue]]:
    """Resolve explicit and inferred ratio group-total rules.

    Explicit YAML rules always take precedence. Formula inference is attempted
    only for ratio KPIs without ``ratio_total`` and only when enabled globally.

    :param formulas_wb: Formula-preserving workbook.
    :param config: Complete application configuration.
    :param records: Extraction records containing source coordinates.
    :return: Rules keyed by ``(source_name, kpi_index)`` and diagnostics.
    """

    rules: dict[tuple[str, int], ResolvedRatioTotal] = {}
    issues: list[ValidationIssue] = []

    for source_name, source in config.workbook.sources.items():
        if not source.enabled:
            continue
        kpis = config.kpis_by_source[source_name]

        for kpi in kpis:
            if kpi.aggregation != "ratio":
                continue

            if kpi.ratio_total is not None:
                rules[(source_name, kpi.index)] = _configured_rule(
                    source_name, kpis, kpi
                )
                continue

            if config.workbook.recursive_formula_totals:
                # The recursive lineage engine handles missing rules, including
                # cross-sheet and multi-step formula chains.
                continue

            if not config.workbook.infer_ratio_totals_from_formulas:
                issues.append(
                    ValidationIssue(
                        country=None,
                        source_sheet=source_name,
                        kpi=kpi.display_name(config.workbook.kpi_title_separator),
                        period=None,
                        issue="RATIO_TOTAL_RULE_MISSING",
                        details=(
                            "No ratio_total is configured and automatic formula inference "
                            "is disabled. TOP8/TOP9/ALL remain blank."
                        ),
                    )
                )
                continue

            rule, issue = _infer_one_ratio(
                formulas_wb=formulas_wb,
                config=config,
                source_name=source_name,
                kpis=kpis,
                ratio_kpi=kpi,
                records=records,
            )
            issues.append(issue)
            if rule is not None:
                rules[(source_name, kpi.index)] = rule

    return rules, issues
