from __future__ import annotations

from openpyxl.styles import Alignment, Border, Font, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from .config import AppConfig, SheetConfig
from .models import ExtractionRecord, GroupTotal, ValidationIssue

_HEADER_FONT = Font(bold=True)
_TITLE_FONT = Font(bold=True, size=12)
_TOTAL_FONT = Font(bold=True)
_THIN_TOP_BORDER = Border(top=Side(style="thin"))


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
) -> dict[tuple[str, str, str, str], ExtractionRecord]:
    """Index extraction records for output writing.

    :param records: Extraction records.
    :return: Lookup keyed by source, KPI, country, and period.
    """

    return {
        (record.source, record.kpi, record.country, record.period): record
        for record in records
    }


def _total_lookup(
    totals: list[GroupTotal],
) -> dict[tuple[str, str, str, str], GroupTotal]:
    """Index group totals for output writing.

    :param totals: Group totals.
    :return: Lookup keyed by source, KPI, group, and period.
    """

    return {
        (total.source, total.kpi, total.group, total.period): total
        for total in totals
    }


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


def _write_source_sheet(
    ws: Worksheet,
    source_name: str,
    sheet: SheetConfig,
    config: AppConfig,
    records: list[ExtractionRecord],
    totals: list[GroupTotal],
) -> None:
    """Write all additive KPI tables for one logical source sheet.

    :param ws: Generated intermediary worksheet.
    :param source_name: Logical source-sheet name.
    :param sheet: Source-sheet configuration.
    :param config: Complete application configuration.
    :param records: Extraction records.
    :param totals: Configured group totals.
    """

    record_by_key = _record_lookup(records)
    total_by_key = _total_lookup(totals)
    sum_kpis = [
        kpi
        for kpi in config.kpis_by_source[source_name]
        if kpi.aggregation == "sum"
    ]

    last_column = 1 + len(sheet.periods)
    row = 1

    for kpi in sum_kpis:
        title_cell = ws.cell(row=row, column=1, value=kpi.name)
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

        data_row = header_row + 1
        for country in config.countries.countries:
            ws.cell(row=data_row, column=1, value=country)
            for offset, period in enumerate(sheet.periods, start=2):
                record = record_by_key[(source_name, kpi.name, country, period)]
                ws.cell(row=data_row, column=offset, value=record.value)
            data_row += 1

        total_row = data_row + 1
        for group_name in config.countries.groups:
            label_cell = ws.cell(
                row=total_row,
                column=1,
                value=f"TOTAL {group_name}",
            )
            label_cell.font = _TOTAL_FONT
            label_cell.border = _THIN_TOP_BORDER

            for offset, period in enumerate(sheet.periods, start=2):
                total = total_by_key[(source_name, kpi.name, group_name, period)]
                cell = ws.cell(row=total_row, column=offset, value=total.value)
                cell.font = _TOTAL_FONT
                cell.border = _THIN_TOP_BORDER
            total_row += 1

        row = total_row + config.workbook.table_spacing_rows

    ws.column_dimensions["A"].width = 28
    for column in range(2, last_column + 1):
        period = sheet.periods[column - 2]
        ws.column_dimensions[get_column_letter(column)].width = max(
            12,
            len(period) + 2,
        )



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
        "C": 55,
        "D": 16,
        "E": 32,
        "F": 90,
    }
    for column, width in widths.items():
        ws.column_dimensions[column].width = width

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions


def write_generated_sheets(
    wb: Workbook,
    config: AppConfig,
    records: list[ExtractionRecord],
    totals: list[GroupTotal],
    issues: list[ValidationIssue],
) -> None:
    """Add intermediary and validation sheets to the existing workbook.

    Existing source worksheets are not recreated or rewritten. Only configured
    generated sheets are created/replaced.

    :param wb: Formula-preserving workbook to modify and save.
    :param config: Complete application configuration.
    :param records: Extraction records.
    :param totals: Configured group totals.
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
            totals=totals,
        )

    validation_ws = _prepare_generated_sheet(
        wb,
        name=config.workbook.validation_sheet,
        replace_existing=replace_existing,
    )
    _write_validation_sheet(validation_ws, issues)
