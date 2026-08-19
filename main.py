from __future__ import annotations

from pathlib import Path

from openpyxl import load_workbook

from pops.aggregator import compute_group_totals
from pops.config import load_app_config
from pops.extractor import extract_workbook
from pops.models import ValidationIssue
from pops.writer import write_generated_sheets

CONFIG_DIR = Path(__file__).resolve().parent / "config"


def _configuration_warnings(config) -> list[ValidationIssue]:
    """Create non-fatal diagnostics for incomplete optional business groups.

    Empty groups are allowed because TOP8/TOP9 membership is a business rule
    that must not be invented by the program.

    :param config: Complete application configuration.
    :return: Configuration warnings suitable for the validation sheet.
    """

    return [
        ValidationIssue(
            country=None,
            source_sheet=None,
            kpi=None,
            period=None,
            issue="EMPTY_COUNTRY_GROUP",
            details=(
                f"Country group {group_name!r} has no members. "
                "Its generated totals will remain blank until configured."
            ),
        )
        for group_name, members in config.countries.groups.items()
        if not members
    ]


def main() -> None:
    """Generate intermediary KPI tables in a copy of Consolidated POPS.

    :raises FileNotFoundError: If the configured input workbook does not exist.
    """

    config = load_app_config(CONFIG_DIR)
    input_path = config.runtime.input_workbook
    output_path = config.runtime.output_workbook

    if not input_path.is_file():
        raise FileNotFoundError(f"Input workbook not found: {input_path}")

    output_path.parent.mkdir(parents=True, exist_ok=True)
    keep_vba = input_path.suffix.lower() == ".xlsm"

    formulas_wb = load_workbook(
        input_path,
        data_only=False,
        keep_vba=keep_vba,
        keep_links=True,
    )
    values_wb = load_workbook(
        input_path,
        data_only=True,
        keep_vba=keep_vba,
        keep_links=True,
    )

    records, issues = extract_workbook(
        values_wb=values_wb,
        formulas_wb=formulas_wb,
        config=config,
    )
    issues = _configuration_warnings(config) + issues
    totals = compute_group_totals(records, config)

    write_generated_sheets(
        wb=formulas_wb,
        config=config,
        records=records,
        totals=totals,
        issues=issues,
    )
    formulas_wb.save(output_path)

    ratio_count = sum(
        1
        for source_name, source in config.workbook.sources.items()
        if source.enabled
        for kpi in config.kpis_by_source[source_name]
        if kpi.aggregation == "ratio"
    )
    skip_count = sum(
        1
        for source_name, source in config.workbook.sources.items()
        if source.enabled
        for kpi in config.kpis_by_source[source_name]
        if kpi.aggregation == "skip"
    )

    print(f"Created: {output_path}")
    print(f"Validation issues: {len(issues)}")
    print(f"Configured ratio KPIs skipped: {ratio_count}")
    print(f"Explicitly skipped KPIs: {skip_count}")


if __name__ == "__main__":
    main()
