from __future__ import annotations

from .config import AppConfig
from .models import ExtractionRecord, GroupTotal


def _build_record_lookup(
    records: list[ExtractionRecord],
) -> dict[tuple[str, str, str, str], ExtractionRecord]:
    """Index extraction records and reject duplicate logical observations.

    :param records: Extraction records.
    :return: Records keyed by source, KPI, country, and period.
    :raises ValueError: If a logical observation occurs more than once.
    """

    lookup: dict[tuple[str, str, str, str], ExtractionRecord] = {}
    for record in records:
        key = (record.source, record.kpi, record.country, record.period)
        if key in lookup:
            raise ValueError(f"Duplicate extraction record detected: {key}")
        lookup[key] = record
    return lookup


def compute_group_totals(
    records: list[ExtractionRecord],
    config: AppConfig,
) -> list[GroupTotal]:
    """Compute configured country-group totals for additive KPIs.

    Missing country values remain missing at country level. Group totals sum only
    valid numeric observations. If a group has no valid observations for a
    KPI/period, its total is ``None`` rather than zero.

    :param records: Validated extraction records.
    :param config: Complete application configuration.
    :return: Ordered group totals.
    """

    lookup = _build_record_lookup(records)
    totals: list[GroupTotal] = []

    for source_name, sheet in config.workbook.sources.items():
        if not sheet.enabled:
            continue

        sum_kpis = [
            kpi
            for kpi in config.kpis_by_source[source_name]
            if kpi.aggregation == "sum"
        ]

        for kpi in sum_kpis:
            for group_name, members in config.countries.groups.items():
                for period in sheet.periods:
                    values = [
                        lookup[(source_name, kpi.name, country, period)].value
                        for country in members
                    ]
                    valid_values = [value for value in values if value is not None]
                    total = sum(valid_values) if valid_values else None
                    totals.append(
                        GroupTotal(
                            source=source_name,
                            kpi=kpi.name,
                            group=group_name,
                            period=period,
                            value=total,
                        )
                    )

    return totals
