from __future__ import annotations

from openpyxl.styles import Alignment, Border, Font, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from .aggregator import build_ratio_formula, build_sum_formula
from .config import AppConfig, KPIConfig, SheetConfig, resolve_ratio_total_indices
from .models import ExtractionRecord, ValidationIssue

_HEADER_FONT = Font(bold=True)
_TITLE_FONT = Font(bold=True, size=12)
_TOTAL_FONT = Font(bold=True)
_THIN_TOP_BORDER = Border(top=Side(style="thin"))
_HYPERLINK_FONT = Font(underline="single")


def _prepare_generated_sheet(
    wb: Workbook,
    name: str,
    replace_existing: bool,
) -> Worksheet:
    """Create a generated worksheet, optionally replacing a previous one.

    :param wb: Formula-preserving workbook.
    :param name: Generated worksheet name.
    :param replace_existing: Whether an existing generated sheet may be removed.
    :return: Fresh worksheet.
    :raises ValueError: If the sheet already exists and replacement is disabled.
    """

    if name in wb.sheetnames:
        if not replace_existing:
            raise ValueError(
                f"Generated worksheet {name!r} already exists and replacement is disabled"
            )
        del wb[name]
    return wb.create_sheet(title=name)


def _record_lookup(
    records: list[ExtractionRecord],
) -> dict[tuple[str, int, str, str], ExtractionRecord]:
    """Index extraction records using KPI configuration position.

    :param records: Extraction records.
    :return: Lookup keyed by source, KPI index, country, and period.
    :raises ValueError: If a logical observation occurs more than once.
    """

    lookup: dict[tuple[str, int, str, str], ExtractionRecord] = {}
    for record in records:
        key = (record.source, record.kpi_index, record.country, record.period)
        if key in lookup:
            raise ValueError(f"Duplicate extraction record detected: {key}")
        lookup[key] = record
    return lookup


def _style_header_row(ws: Worksheet, row: int, last_column: int) -> None:
    """Apply simple header formatting.

    :param ws: Output worksheet.
    :param row: Header row number.
    :param last_column: Last table column.
    """

    for column in range(1, last_column + 1):
        cell = ws.cell(row=row, column=column)
        cell.font = _HEADER_FONT
        cell.alignment = Alignment(horizontal="center")


def _quote_sheet_name(name: str) -> str:
    """Quote an Excel worksheet name for formulas and internal links.

    :param name: Worksheet name.
    :return: Safely quoted worksheet name.
    """

    return "'" + name.replace("'", "''") + "'"


def _is_percent_format(number_format: str | None) -> bool:
    """Return whether an Excel number format represents a percentage.

    :param number_format: Excel number-format code.
    :return: ``True`` when the format contains a percent token.
    """

    return bool(number_format and "%" in number_format)


def _rounded_number_format(percent: bool, digits: int) -> str:
    """Build a simple rounded output number format.

    :param percent: Whether the value is a percentage decimal.
    :param digits: Decimal digits to display.
    :return: Excel number-format code.
    """

    decimals = "." + ("0" * digits) if digits else ""
    return f"0{decimals}%" if percent else f"#,##0{decimals}"


def _source_formula(
    sheet_name: str,
    coordinate: str,
    round_values: bool,
    round_digits: int,
    percent: bool,
) -> str:
    """Build a guarded Excel formula linking to a source cell.

    A plain direct reference displays an empty Excel source cell as zero. The
    guard preserves genuine blanks and source errors as blank intermediary cells.
    When rounding is enabled, percentages are rounded in percentage points while
    preserving Excel's underlying decimal representation.

    :param sheet_name: Physical source worksheet name.
    :param coordinate: Source cell coordinate.
    :param round_values: Whether to round the linked numeric value.
    :param round_digits: Decimal digits used when rounding.
    :param percent: Whether the source value is an Excel percentage decimal.
    :return: Excel formula.
    """

    reference = f"{_quote_sheet_name(sheet_name)}!{coordinate}"
    value_expression = reference
    if round_values:
        if percent:
            value_expression = f"ROUND({reference}*100,{round_digits})/100"
        else:
            value_expression = f"ROUND({reference},{round_digits})"
    return f'=IFERROR(IF({reference}="","",{value_expression}),"")'


def _source_hyperlink(sheet_name: str, coordinate: str) -> str:
    """Build an internal hyperlink to a source cell.

    :param sheet_name: Physical source worksheet name.
    :param coordinate: Source cell coordinate.
    :return: Internal Excel hyperlink target.
    """

    return f"#{_quote_sheet_name(sheet_name)}!{coordinate}"


def _table_layout(
    kpis: list[KPIConfig],
    countries: tuple[str, ...],
    groups: dict[str, tuple[str, ...]],
    spacing_rows: int,
) -> tuple[
    dict[int, dict[str, int]],
    dict[tuple[int, str], int],
    dict[int, int],
]:
    """Pre-compute table rows so ratio formulas can reference other KPI totals.

    :param kpis: Generated KPI definitions in display order.
    :param countries: Ordered configured countries.
    :param groups: Ordered configured country groups.
    :param spacing_rows: Empty rows between KPI tables.
    :return: Country rows, group-total rows, and title rows keyed by KPI index.
    """

    country_rows: dict[int, dict[str, int]] = {}
    total_rows: dict[tuple[int, str], int] = {}
    title_rows: dict[int, int] = {}
    row = 1

    for kpi in kpis:
        title_rows[kpi.index] = row
        header_row = row + 1
        first_country_row = header_row + 1
        country_rows[kpi.index] = {
            country: first_country_row + offset
            for offset, country in enumerate(countries)
        }
        total_row = first_country_row + len(countries) + 1
        for group_name in groups:
            total_rows[(kpi.index, group_name)] = total_row
            total_row += 1
        row = total_row + spacing_rows

    return country_rows, total_rows, title_rows


def _write_source_sheet(
    ws: Worksheet,
    source_name: str,
    sheet: SheetConfig,
    config: AppConfig,
    records: list[ExtractionRecord],
) -> None:
    """Write additive and ratio KPI tables using native Excel formulas.

    Country cells always reference the resolved source cells. Additive group
    totals use Excel ``SUM`` formulas. Ratio country values also reference their
    already-calculated source ratio cells; configured ratio group totals divide
    the corresponding additive numerator total by additive denominator total.

    :param ws: Generated intermediary worksheet.
    :param source_name: Logical source-sheet name.
    :param sheet: Source-sheet configuration.
    :param config: Complete application configuration.
    :param records: Extraction records.
    """

    record_by_key = _record_lookup(records)
    configured_kpis = config.kpis_by_source[source_name]
    output_kpis = [
        kpi for kpi in configured_kpis if kpi.aggregation in {"sum", "ratio"}
    ]
    country_rows, total_rows, title_rows = _table_layout(
        output_kpis,
        config.countries.countries,
        config.countries.groups,
        config.workbook.table_spacing_rows,
    )

    last_column = 1 + len(sheet.periods)

    for kpi in output_kpis:
        row = title_rows[kpi.index]
        title = kpi.display_name(config.workbook.kpi_title_separator)
        title_cell = ws.cell(row=row, column=1, value=title)
        title_cell.font = _TITLE_FONT
        if last_column > 1:
            ws.merge_cells(
                start_row=row,
                start_column=1,
                end_row=row,
                end_column=last_column,
            )

        header_row = row + 1
        ws.cell(row=header_row, column=1, value="Countries")
        for offset, period in enumerate(sheet.periods, start=2):
            ws.cell(row=header_row, column=offset, value=period)
        _style_header_row(ws, header_row, last_column)

        for country in config.countries.countries:
            data_row = country_rows[kpi.index][country]
            ws.cell(row=data_row, column=1, value=country)

            for offset, period in enumerate(sheet.periods, start=2):
                record = record_by_key[(source_name, kpi.index, country, period)]
                cell = ws.cell(row=data_row, column=offset)

                if record.coordinate is None:
                    cell.value = None
                    continue

                source_percent = _is_percent_format(record.number_format)
                cell.value = _source_formula(
                    record.source_sheet,
                    record.coordinate,
                    round_values=config.workbook.round_values,
                    round_digits=config.workbook.round_digits,
                    percent=source_percent,
                )

                if config.workbook.round_values:
                    cell.number_format = _rounded_number_format(
                        source_percent,
                        config.workbook.round_digits,
                    )
                elif record.number_format:
                    cell.number_format = record.number_format

                if config.workbook.add_source_hyperlinks:
                    cell.hyperlink = _source_hyperlink(
                        record.source_sheet,
                        record.coordinate,
                    )
                    cell.font = _HYPERLINK_FONT

        ratio_indices = (
            resolve_ratio_total_indices(configured_kpis, kpi, source_name)
            if kpi.aggregation == "ratio" and kpi.ratio_total is not None
            else None
        )

        for group_name, members in config.countries.groups.items():
            total_row = total_rows[(kpi.index, group_name)]
            label_cell = ws.cell(
                row=total_row,
                column=1,
                value=f"TOTAL {group_name}",
            )
            label_cell.font = _TOTAL_FONT
            label_cell.border = _THIN_TOP_BORDER

            for offset, _period in enumerate(sheet.periods, start=2):
                cell = ws.cell(row=total_row, column=offset)

                if kpi.aggregation == "sum":
                    cell.value = build_sum_formula(
                        offset,
                        country_rows[kpi.index],
                        members,
                        round_values=config.workbook.round_values,
                        round_digits=config.workbook.round_digits,
                    )
                    if config.workbook.round_values:
                        cell.number_format = _rounded_number_format(
                            False,
                            config.workbook.round_digits,
                        )

                elif ratio_indices is not None:
                    numerator_index, denominator_index = ratio_indices
                    numerator_row = total_rows[(numerator_index, group_name)]
                    denominator_row = total_rows[(denominator_index, group_name)]
                    column_letter = get_column_letter(offset)
                    cell.value = build_ratio_formula(
                        f"{column_letter}{numerator_row}",
                        f"{column_letter}{denominator_row}",
                        multiplier=kpi.ratio_total.multiplier,
                        percent=kpi.ratio_total.percent,
                        round_values=config.workbook.round_values,
                        round_digits=config.workbook.round_digits,
                    )
                    if kpi.ratio_total.percent:
                        cell.number_format = _rounded_number_format(
                            True,
                            config.workbook.round_digits,
                        )
                    elif config.workbook.round_values:
                        cell.number_format = _rounded_number_format(
                            False,
                            config.workbook.round_digits,
                        )

                else:
                    cell.value = None

                cell.font = _TOTAL_FONT
                cell.border = _THIN_TOP_BORDER

    ws.column_dimensions["A"].width = 32
    for column in range(2, last_column + 1):
        period = sheet.periods[column - 2]
        ws.column_dimensions[get_column_letter(column)].width = max(12, len(period) + 2)


def _write_validation_sheet(
    ws: Worksheet,
    issues: list[ValidationIssue],
) -> None:
    """Write validation diagnostics to a dedicated worksheet.

    :param ws: Validation worksheet.
    :param issues: Ordered validation issues.
    """

    headers = [
        "Country",
        "Source Sheet",
        "KPI",
        "Period",
        "Issue",
        "Cell/Details",
    ]
    ws.append(headers)
    _style_header_row(ws, 1, len(headers))

    for issue in issues:
        ws.append(
            [
                issue.country,
                issue.source_sheet,
                issue.kpi,
                issue.period,
                issue.issue,
                issue.details,
            ]
        )

    widths = {
        "A": 16,
        "B": 28,
        "C": 65,
        "D": 16,
        "E": 36,
        "F": 100,
    }
    for column, width in widths.items():
        ws.column_dimensions[column].width = width

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions


def write_generated_sheets(
    wb: Workbook,
    config: AppConfig,
    records: list[ExtractionRecord],
    issues: list[ValidationIssue],
) -> None:
    """Add formula-linked intermediary and validation sheets to the workbook.

    Existing source worksheets are not recreated or rewritten. Only configured
    generated sheets are created/replaced.

    :param wb: Formula-preserving workbook to modify and save.
    :param config: Complete application configuration.
    :param records: Extraction records containing resolved source coordinates.
    :param issues: Validation issues.
    """

    replace_existing = config.workbook.replace_existing_generated_sheets

    for source_name, sheet in config.workbook.sources.items():
        if not sheet.enabled:
            continue
        if sheet.output_sheet is None:
            raise ValueError(f"Enabled source {source_name!r} has no output sheet")

        ws = _prepare_generated_sheet(
            wb,
            name=sheet.output_sheet,
            replace_existing=replace_existing,
        )
        _write_source_sheet(
            ws=ws,
            source_name=source_name,
            sheet=sheet,
            config=config,
            records=records,
        )

    validation_ws = _prepare_generated_sheet(
        wb,
        name=config.workbook.validation_sheet,
        replace_existing=replace_existing,
    )
    _write_validation_sheet(validation_ws, issues)

    calculation = getattr(wb, "calculation", None)
    if calculation is not None:
        calculation.calcMode = "auto"
        calculation.fullCalcOnLoad = True
        calculation.forceFullCalc = True
