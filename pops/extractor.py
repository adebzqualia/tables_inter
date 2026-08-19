from __future__ import annotations

from collections import Counter, defaultdict
from dataclasses import dataclass
from decimal import Decimal
from typing import Iterable

from openpyxl.cell.cell import Cell, MergedCell
from openpyxl.worksheet.worksheet import Worksheet

from .config import AppConfig, KPIConfig, SheetConfig, normalize_label
from .models import ExtractionRecord, ValidationIssue

_EXCEL_ERROR_VALUES = {
    "#NULL!",
    "#DIV/0!",
    "#VALUE!",
    "#REF!",
    "#NAME?",
    "#NUM!",
    "#N/A",
    "#GETTING_DATA",
}


@dataclass(frozen=True)
class CellRef:
    """Coordinate and raw label value for an indexed cell.

    :param row: One-based Excel row.
    :param column: One-based Excel column.
    :param coordinate: Excel coordinate such as ``J37``.
    :param value: Raw indexed value.
    """

    row: int
    column: int
    coordinate: str
    value: object


@dataclass(frozen=True)
class ResolvedLabels:
    """Resolved KPI and period labels for one worksheet.

    :param periods: Period label to resolved cell, or ``None``.
    :param kpis: KPI label to resolved cell, or ``None``.
    :param issues: Matching diagnostics.
    """

    periods: dict[str, CellRef | None]
    kpis: dict[str, CellRef | None]
    issues: tuple[ValidationIssue, ...]


def _physical_sheet_name(template: str, country: str, source: str) -> str:
    """Construct a physical source worksheet name.

    :param template: Configured sheet-name template.
    :param country: Country/entity name.
    :param source: Logical source-sheet name.
    :return: Physical worksheet name.
    """

    return template.format(country=country, sheet=source)


def build_label_index(
    values_ws: Worksheet,
    formulas_ws: Worksheet,
) -> dict[str, list[CellRef]]:
    """Scan a worksheet once and index normalized cell labels.

    Formula cells use the ``data_only`` workbook's cached result. Non-formula
    cells use the formula-preserving workbook directly.

    :param values_ws: Worksheet from the ``data_only=True`` workbook.
    :param formulas_ws: Worksheet from the formula-preserving workbook.
    :return: Mapping from normalized label to all matching coordinates.
    """

    index: dict[str, list[CellRef]] = defaultdict(list)

    for row in formulas_ws.iter_rows():
        for formula_cell in row:
            raw_value = (
                values_ws[formula_cell.coordinate].value
                if formula_cell.data_type == "f"
                else formula_cell.value
            )
            normalized = normalize_label(raw_value)
            if not normalized:
                continue
            index[normalized].append(
                CellRef(
                    row=formula_cell.row,
                    column=formula_cell.column,
                    coordinate=formula_cell.coordinate,
                    value=raw_value,
                )
            )

    return dict(index)


def _candidate_map(
    labels: Iterable[str],
    label_index: dict[str, list[CellRef]],
) -> dict[str, list[CellRef]]:
    """Return exact normalized candidates for configured labels.

    :param labels: Configured labels.
    :param label_index: Worksheet label index.
    :return: Candidate coordinates by original configured label.
    """

    return {
        label: list(label_index.get(normalize_label(label), []))
        for label in labels
    }


def _support_by_row(candidates: dict[str, list[CellRef]]) -> Counter[int]:
    """Count how many distinct configured labels occur on each row.

    :param candidates: Candidate coordinates keyed by label.
    :return: Row support counts.
    """

    support: Counter[int] = Counter()
    for refs in candidates.values():
        for row in {ref.row for ref in refs}:
            support[row] += 1
    return support


def _support_by_column(candidates: dict[str, list[CellRef]]) -> Counter[int]:
    """Count how many distinct configured labels occur in each column.

    :param candidates: Candidate coordinates keyed by label.
    :return: Column support counts.
    """

    support: Counter[int] = Counter()
    for refs in candidates.values():
        for column in {ref.column for ref in refs}:
            support[column] += 1
    return support


def _coordinates(refs: Iterable[CellRef]) -> str:
    """Format candidate coordinates for diagnostics.

    :param refs: Candidate cells.
    :return: Comma-separated coordinates.
    """

    return ", ".join(ref.coordinate for ref in refs)


def _resolve_periods(
    country: str,
    sheet_name: str,
    periods: tuple[str, ...],
    candidates: dict[str, list[CellRef]],
    kpi_candidates: dict[str, list[CellRef]],
) -> tuple[dict[str, CellRef | None], list[ValidationIssue]]:
    """Resolve configured period labels without arbitrary first-match logic.

    Duplicate candidates are first constrained to cells above the apparent KPI
    body, then by the row containing the most distinct configured periods. A
    remaining tie is reported as ambiguous.

    :param country: Country/entity name.
    :param sheet_name: Physical worksheet name.
    :param periods: Ordered configured periods.
    :param candidates: Period candidates.
    :param kpi_candidates: KPI candidates used for table context.
    :return: Resolved period cells and diagnostics.
    """

    issues: list[ValidationIssue] = []
    resolved: dict[str, CellRef | None] = {}
    row_support = _support_by_row(candidates)

    all_kpi_rows = [ref.row for refs in kpi_candidates.values() for ref in refs]
    body_top = min(all_kpi_rows) if all_kpi_rows else None

    for period in periods:
        refs = candidates[period]
        if not refs:
            resolved[period] = None
            issues.append(
                ValidationIssue(
                    country=country,
                    source_sheet=sheet_name,
                    kpi=None,
                    period=period,
                    issue="PERIOD_NOT_FOUND",
                    details="Configured period label was not found.",
                )
            )
            continue

        if len(refs) == 1:
            resolved[period] = refs[0]
            continue

        contextual = refs
        if body_top is not None:
            above_body = [ref for ref in refs if ref.row < body_top]
            if above_body:
                contextual = above_body

        chosen: CellRef | None = None
        if len(contextual) == 1:
            chosen = contextual[0]
        else:
            max_support = max(row_support[ref.row] for ref in contextual)
            strongest = [
                ref for ref in contextual if row_support[ref.row] == max_support
            ]
            if max_support > 1 and len(strongest) == 1:
                chosen = strongest[0]

        if chosen is None:
            resolved[period] = None
            issues.append(
                ValidationIssue(
                    country=country,
                    source_sheet=sheet_name,
                    kpi=None,
                    period=period,
                    issue="DUPLICATE_PERIOD_MATCH",
                    details=f"Ambiguous exact matches at: {_coordinates(refs)}",
                )
            )
            continue

        resolved[period] = chosen
        issues.append(
            ValidationIssue(
                country=country,
                source_sheet=sheet_name,
                kpi=None,
                period=period,
                issue="DUPLICATE_PERIOD_RESOLVED",
                details=(
                    f"Exact matches at: {_coordinates(refs)}. "
                    f"Resolved by table context to {chosen.coordinate}."
                ),
            )
        )

    return resolved, issues


def _resolve_kpis(
    country: str,
    sheet_name: str,
    kpis: tuple[KPIConfig, ...],
    candidates: dict[str, list[CellRef]],
    resolved_periods: dict[str, CellRef | None],
) -> tuple[dict[str, CellRef | None], list[ValidationIssue]]:
    """Resolve configured KPI labels without substring or fuzzy matching.

    Duplicate candidates are constrained below the resolved period headers and
    to the left of the resolved data columns where possible. Remaining ties use
    KPI-column support; unresolved ties are reported.

    :param country: Country/entity name.
    :param sheet_name: Physical worksheet name.
    :param kpis: Configured KPIs.
    :param candidates: KPI candidates.
    :param resolved_periods: Previously resolved period cells.
    :return: Resolved KPI cells and diagnostics.
    """

    issues: list[ValidationIssue] = []
    resolved: dict[str, CellRef | None] = {}
    column_support = _support_by_column(candidates)

    period_refs = [ref for ref in resolved_periods.values() if ref is not None]
    header_bottom = max((ref.row for ref in period_refs), default=None)
    leftmost_period_column = min((ref.column for ref in period_refs), default=None)

    for kpi in kpis:
        refs = candidates[kpi.name]
        if not refs:
            resolved[kpi.name] = None
            issues.append(
                ValidationIssue(
                    country=country,
                    source_sheet=sheet_name,
                    kpi=kpi.name,
                    period=None,
                    issue="KPI_NOT_FOUND",
                    details="Configured KPI label was not found.",
                )
            )
            continue

        if len(refs) == 1:
            resolved[kpi.name] = refs[0]
            continue

        contextual = refs
        constrained = [
            ref
            for ref in refs
            if (header_bottom is None or ref.row > header_bottom)
            and (
                leftmost_period_column is None
                or ref.column < leftmost_period_column
            )
        ]
        if constrained:
            contextual = constrained

        chosen: CellRef | None = None
        if len(contextual) == 1:
            chosen = contextual[0]
        else:
            max_support = max(column_support[ref.column] for ref in contextual)
            strongest = [
                ref
                for ref in contextual
                if column_support[ref.column] == max_support
            ]
            if max_support > 1 and len(strongest) == 1:
                chosen = strongest[0]

        if chosen is None:
            resolved[kpi.name] = None
            issues.append(
                ValidationIssue(
                    country=country,
                    source_sheet=sheet_name,
                    kpi=kpi.name,
                    period=None,
                    issue="DUPLICATE_KPI_MATCH",
                    details=f"Ambiguous exact matches at: {_coordinates(refs)}",
                )
            )
            continue

        resolved[kpi.name] = chosen
        issues.append(
            ValidationIssue(
                country=country,
                source_sheet=sheet_name,
                kpi=kpi.name,
                period=None,
                issue="DUPLICATE_KPI_RESOLVED",
                details=(
                    f"Exact matches at: {_coordinates(refs)}. "
                    f"Resolved by table context to {chosen.coordinate}."
                ),
            )
        )

    return resolved, issues


def resolve_labels(
    country: str,
    sheet_name: str,
    sheet: SheetConfig,
    kpis: tuple[KPIConfig, ...],
    label_index: dict[str, list[CellRef]],
) -> ResolvedLabels:
    """Resolve all configured periods and KPIs for one worksheet.

    :param country: Country/entity name.
    :param sheet_name: Physical worksheet name.
    :param sheet: Source-sheet configuration.
    :param kpis: KPI configuration for the source.
    :param label_index: One-pass worksheet label index.
    :return: Resolved labels and diagnostics.
    """

    period_candidates = _candidate_map(sheet.periods, label_index)
    kpi_candidates = _candidate_map((kpi.name for kpi in kpis), label_index)

    resolved_periods, period_issues = _resolve_periods(
        country=country,
        sheet_name=sheet_name,
        periods=sheet.periods,
        candidates=period_candidates,
        kpi_candidates=kpi_candidates,
    )
    resolved_kpis, kpi_issues = _resolve_kpis(
        country=country,
        sheet_name=sheet_name,
        kpis=kpis,
        candidates=kpi_candidates,
        resolved_periods=resolved_periods,
    )

    return ResolvedLabels(
        periods=resolved_periods,
        kpis=resolved_kpis,
        issues=tuple(period_issues + kpi_issues),
    )


def _merged_anchor(ws: Worksheet, row: int, column: int) -> Cell | MergedCell:
    """Return the top-left anchor when a coordinate lies in a merged range.

    :param ws: Worksheet.
    :param row: One-based row.
    :param column: One-based column.
    :return: Real cell or merged-range anchor.
    """

    cell = ws.cell(row=row, column=column)
    if not isinstance(cell, MergedCell):
        return cell

    for merged_range in ws.merged_cells.ranges:
        if cell.coordinate in merged_range:
            return ws.cell(row=merged_range.min_row, column=merged_range.min_col)
    return cell


def _is_excel_error(cell: Cell | MergedCell, value: object) -> bool:
    """Return whether a cell/value represents an Excel error.

    :param cell: Resolved worksheet cell.
    :param value: Cell value.
    :return: ``True`` for Excel error values.
    """

    if getattr(cell, "data_type", None) == "e":
        return True
    return isinstance(value, str) and value.strip().upper() in _EXCEL_ERROR_VALUES


def _extract_numeric_value(
    country: str,
    sheet_name: str,
    kpi: str,
    period: str,
    row: int,
    column: int,
    values_ws: Worksheet,
    formulas_ws: Worksheet,
) -> tuple[int | float | None, str, ValidationIssue | None]:
    """Extract and validate one additive KPI value.

    :param country: Country/entity name.
    :param sheet_name: Physical worksheet name.
    :param kpi: KPI label.
    :param period: Period label.
    :param row: Resolved KPI row.
    :param column: Resolved period column.
    :param values_ws: Worksheet containing cached formula results.
    :param formulas_ws: Formula-preserving worksheet.
    :return: Numeric value, source coordinate, and optional diagnostic.
    """

    formula_cell = _merged_anchor(formulas_ws, row, column)
    value_cell = _merged_anchor(values_ws, row, column)
    coordinate = formula_cell.coordinate
    raw_value = value_cell.value

    if _is_excel_error(value_cell, raw_value) or _is_excel_error(
        formula_cell, formula_cell.value
    ):
        return None, coordinate, ValidationIssue(
            country=country,
            source_sheet=sheet_name,
            kpi=kpi,
            period=period,
            issue="EXCEL_ERROR_VALUE",
            details=f"Excel error at {coordinate}: {raw_value!r}",
        )

    if getattr(formula_cell, "data_type", None) == "f" and raw_value is None:
        return None, coordinate, ValidationIssue(
            country=country,
            source_sheet=sheet_name,
            kpi=kpi,
            period=period,
            issue="FORMULA_WITHOUT_CACHED_VALUE",
            details=(
                f"Formula at {coordinate} has no cached calculated value. "
                "Open/recalculate/save the workbook in Excel before rerunning."
            ),
        )

    if raw_value is None:
        return None, coordinate, ValidationIssue(
            country=country,
            source_sheet=sheet_name,
            kpi=kpi,
            period=period,
            issue="BLANK_VALUE",
            details=f"Resolved source cell {coordinate} is blank.",
        )

    if isinstance(raw_value, bool) or not isinstance(
        raw_value, (int, float, Decimal)
    ):
        return None, coordinate, ValidationIssue(
            country=country,
            source_sheet=sheet_name,
            kpi=kpi,
            period=period,
            issue="NON_NUMERIC_VALUE",
            details=f"Expected a numeric value at {coordinate}, found {raw_value!r}.",
        )

    if isinstance(raw_value, Decimal):
        raw_value = float(raw_value)

    return raw_value, coordinate, None


def _empty_records_for_sheet(
    country: str,
    source: str,
    periods: tuple[str, ...],
    kpis: tuple[KPIConfig, ...],
) -> list[ExtractionRecord]:
    """Create missing records for all additive KPI intersections.

    :param country: Country/entity name.
    :param source: Logical source-sheet name.
    :param periods: Configured periods.
    :param kpis: Configured KPIs.
    :return: Missing extraction records for ``sum`` KPIs.
    """

    return [
        ExtractionRecord(
            country=country,
            source=source,
            kpi=kpi.name,
            period=period,
            value=None,
            coordinate=None,
        )
        for kpi in kpis
        if kpi.aggregation == "sum"
        for period in periods
    ]


def extract_workbook(
    values_wb,
    formulas_wb,
    config: AppConfig,
) -> tuple[list[ExtractionRecord], list[ValidationIssue]]:
    """Extract all enabled source sheets from the consolidated workbook.

    Each relevant worksheet is scanned once. Only additive KPIs produce value
    records; configured ratio/skip KPIs are recognized but deliberately not
    aggregated in this version.

    :param values_wb: Workbook loaded with ``data_only=True``.
    :param formulas_wb: Workbook loaded with ``data_only=False``.
    :param config: Complete application configuration.
    :return: Extraction records and validation issues.
    """

    records: list[ExtractionRecord] = []
    issues: list[ValidationIssue] = []

    for source_name, sheet in config.workbook.sources.items():
        if not sheet.enabled:
            continue

        kpis = config.kpis_by_source[source_name]
        sum_kpis = tuple(kpi for kpi in kpis if kpi.aggregation == "sum")

        for country in config.countries.countries:
            sheet_name = _physical_sheet_name(
                config.workbook.sheet_name_template,
                country,
                source_name,
            )

            if sheet_name not in formulas_wb.sheetnames or sheet_name not in values_wb.sheetnames:
                issues.append(
                    ValidationIssue(
                        country=country,
                        source_sheet=sheet_name,
                        kpi=None,
                        period=None,
                        issue="SOURCE_SHEET_MISSING",
                        details="Expected source worksheet does not exist.",
                    )
                )
                records.extend(
                    _empty_records_for_sheet(
                        country=country,
                        source=source_name,
                        periods=sheet.periods,
                        kpis=kpis,
                    )
                )
                continue

            formulas_ws = formulas_wb[sheet_name]
            values_ws = values_wb[sheet_name]
            label_index = build_label_index(values_ws, formulas_ws)
            resolved = resolve_labels(
                country=country,
                sheet_name=sheet_name,
                sheet=sheet,
                kpis=kpis,
                label_index=label_index,
            )
            issues.extend(resolved.issues)

            for kpi in sum_kpis:
                kpi_ref = resolved.kpis[kpi.name]
                for period in sheet.periods:
                    period_ref = resolved.periods[period]

                    if kpi_ref is None or period_ref is None:
                        records.append(
                            ExtractionRecord(
                                country=country,
                                source=source_name,
                                kpi=kpi.name,
                                period=period,
                                value=None,
                                coordinate=None,
                            )
                        )
                        continue

                    value, coordinate, issue = _extract_numeric_value(
                        country=country,
                        sheet_name=sheet_name,
                        kpi=kpi.name,
                        period=period,
                        row=kpi_ref.row,
                        column=period_ref.column,
                        values_ws=values_ws,
                        formulas_ws=formulas_ws,
                    )
                    records.append(
                        ExtractionRecord(
                            country=country,
                            source=source_name,
                            kpi=kpi.name,
                            period=period,
                            value=value,
                            coordinate=coordinate,
                        )
                    )
                    if issue is not None:
                        issues.append(issue)

    return records, issues
