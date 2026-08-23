from __future__ import annotations

import math
import re
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Any

import yaml

SUPPORTED_AGGREGATIONS = {"sum", "ratio", "skip"}
_INVALID_SHEET_CHARS = re.compile(r"[\\/*?:\[\]]")


@dataclass(frozen=True)
class RuntimeConfig:
    """Runtime workbook paths.

    :param input_workbook: Source Consolidated POPS workbook.
    :param output_workbook: Destination workbook.
    """

    input_workbook: Path
    output_workbook: Path


@dataclass(frozen=True)
class CountryConfig:
    """Configured countries and country groups.

    :param countries: Ordered source country/entity names.
    :param groups: Ordered mapping of group names to members.
    """

    countries: tuple[str, ...]
    groups: dict[str, tuple[str, ...]]


@dataclass(frozen=True)
class SheetConfig:
    """Configuration for one logical source sheet.

    :param name: Original source-sheet name, for example ``ID Card``.
    :param enabled: Whether the sheet is processed in the current run.
    :param periods: Ordered configured period labels.
    :param output_sheet: Generated intermediary sheet name.
    """

    name: str
    enabled: bool
    periods: tuple[str, ...]
    output_sheet: str | None


@dataclass(frozen=True)
class WorkbookConfig:
    """Workbook-level configuration.

    :param sheet_name_template: Template used to construct source sheet names.
    :param validation_sheet: Generated validation sheet name.
    :param replace_existing_generated_sheets: Whether generated sheets may be recreated.
    :param table_spacing_rows: Empty rows between intermediary KPI tables.
    :param add_source_hyperlinks: Whether intermediary cells link to their source cell.
    :param kpi_title_separator: Separator used between type/subtype/name in titles.
    :param round_values: Whether generated numeric values are rounded in Excel formulas.
    :param round_digits: Decimal digits used for non-ratio values when rounding is enabled.
    :param ratio_round_digits: Decimal digits used for ratio percentages.
    :param sources: Ordered logical source-sheet definitions.
    """

    sheet_name_template: str
    validation_sheet: str
    replace_existing_generated_sheets: bool
    table_spacing_rows: int
    add_source_hyperlinks: bool
    kpi_title_separator: str
    round_values: bool
    round_digits: int
    ratio_round_digits: int
    sources: dict[str, SheetConfig]


@dataclass(frozen=True)
class KPIRefConfig:
    """Reference to another KPI occurrence in the same source sheet.

    :param name: Configured KPI name.
    :param occurrence: One-based occurrence among same-name configured KPIs.
    """

    name: str
    occurrence: int = 1


@dataclass(frozen=True)
class RatioTotalConfig:
    """Rule for aggregating a simple ratio at country-group level.

    :param numerator: Additive KPI used as numerator.
    :param denominator: Additive KPI used as denominator.
    :param multiplier: Optional numeric scale applied after division.
    :param percent: Whether the ratio is an Excel percentage value.
    """

    numerator: KPIRefConfig
    denominator: KPIRefConfig
    multiplier: float = 1.0
    percent: bool = False


@dataclass(frozen=True)
class KPIConfig:
    """Configuration for one KPI occurrence.

    Duplicate ``name`` values are intentionally allowed. They are matched to
    worksheet occurrences in top-to-bottom order using configuration order.
    ``type`` and ``subtype`` are display metadata only and never affect matching.

    :param index: Zero-based position in the source KPI configuration.
    :param name: KPI label to match in Excel.
    :param aggregation: Aggregation method such as ``sum`` or ``ratio``.
    :param type: Optional display-only KPI type/context.
    :param subtype: Optional display-only KPI subtype/context.
    :param note: Optional business note.
    :param ratio_total: Optional group-total rule for simple ratios.
    """

    index: int
    name: str
    aggregation: str
    type: str | None = None
    subtype: str | None = None
    note: str | None = None
    ratio_total: RatioTotalConfig | None = None

    def display_name(self, separator: str = " | ") -> str:
        """Build the intermediary title without repeating identical metadata.

        :param separator: Text separating optional metadata and KPI name.
        :return: Display label.
        """

        parts: list[str] = []
        for part in (self.type, self.subtype, self.name):
            if not part:
                continue
            if parts and normalize_label(parts[-1]) == normalize_label(part):
                continue
            parts.append(part)
        return separator.join(parts)


@dataclass(frozen=True)
class AppConfig:
    """Complete application configuration.

    :param runtime: Input/output workbook paths.
    :param countries: Country and group definitions.
    :param workbook: Sheet and output definitions.
    :param kpis_by_source: KPI definitions keyed by logical source-sheet name.
    """

    runtime: RuntimeConfig
    countries: CountryConfig
    workbook: WorkbookConfig
    kpis_by_source: dict[str, tuple[KPIConfig, ...]]


def normalize_label(value: object) -> str:
    """Normalize an Excel label for controlled exact matching.

    The function normalizes harmless presentation differences only. It does not
    remove punctuation or use fuzzy matching.

    :param value: Raw Excel cell value.
    :return: Normalized label.
    """

    if value is None:
        return ""

    if isinstance(value, bool):
        text = str(value)
    elif isinstance(value, int):
        text = str(value)
    elif isinstance(value, float) and math.isfinite(value) and value.is_integer():
        text = str(int(value))
    elif isinstance(value, (date, datetime)):
        text = value.isoformat()
    else:
        text = str(value)

    text = text.replace("\u00a0", " ")
    text = text.replace("\r", " ").replace("\n", " ").replace("\t", " ")
    return " ".join(text.split()).casefold()


def _load_yaml(path: Path) -> dict[str, Any]:
    """Load one YAML mapping.

    :param path: YAML file path.
    :return: Parsed YAML mapping.
    :raises ValueError: If the file does not contain a mapping.
    """

    with path.open("r", encoding="utf-8") as handle:
        data = yaml.safe_load(handle) or {}

    if not isinstance(data, dict):
        raise ValueError(f"Expected a YAML mapping in {path}")
    return data


def _resolve_path(raw_path: str, project_dir: Path) -> Path:
    """Resolve a configured path relative to the project directory.

    :param raw_path: Configured absolute or relative path.
    :param project_dir: Project root directory.
    :return: Absolute normalized path.
    """

    path = Path(raw_path).expanduser()
    if not path.is_absolute():
        path = project_dir / path
    return path.resolve()


def _default_output_path(input_path: Path) -> Path:
    """Build the default non-destructive output workbook path.

    :param input_path: Source workbook path.
    :return: Derived destination path.
    """

    return input_path.with_name(
        f"{input_path.stem}_with_intermediary_tables{input_path.suffix}"
    )


def _ensure_unique_labels(values: list[str], context: str) -> None:
    """Reject duplicate configured labels after normalization.

    This is used for countries and periods. KPI names are deliberately excluded
    because repeated KPI labels are supported positionally.

    :param values: Labels to validate.
    :param context: Human-readable configuration context.
    :raises ValueError: If two labels normalize to the same value.
    """

    seen: dict[str, str] = {}
    for value in values:
        normalized = normalize_label(value)
        if not normalized:
            raise ValueError(f"Blank label is not allowed in {context}")
        if normalized in seen:
            raise ValueError(
                f"Duplicate normalized label in {context}: "
                f"{seen[normalized]!r} and {value!r}"
            )
        seen[normalized] = value


def _validate_sheet_name(name: str, context: str) -> None:
    """Validate an Excel worksheet name.

    :param name: Worksheet name.
    :param context: Configuration context.
    :raises ValueError: If the name is invalid for Excel.
    """

    if not name or len(name) > 31 or _INVALID_SHEET_CHARS.search(name):
        raise ValueError(f"Invalid Excel worksheet name for {context}: {name!r}")


def _load_runtime(path: Path, project_dir: Path) -> RuntimeConfig:
    """Load runtime path configuration.

    :param path: ``runtime.yaml`` path.
    :param project_dir: Project root directory.
    :return: Validated runtime configuration.
    """

    data = _load_yaml(path)
    raw_input = data.get("input_workbook")
    if not raw_input:
        raise ValueError("runtime.yaml must define input_workbook")

    input_path = _resolve_path(str(raw_input), project_dir)
    raw_output = data.get("output_workbook")
    output_path = (
        _resolve_path(str(raw_output), project_dir)
        if raw_output
        else _default_output_path(input_path)
    )

    if input_path == output_path:
        raise ValueError(
            "Input and output workbook paths must differ. "
            "This tool writes to a new file by design."
        )
    if input_path.suffix.lower() not in {".xlsx", ".xlsm"}:
        raise ValueError("Only .xlsx and .xlsm workbooks are supported")
    if output_path.suffix.lower() != input_path.suffix.lower():
        raise ValueError("Input and output workbook extensions must match")

    return RuntimeConfig(input_workbook=input_path, output_workbook=output_path)


def _load_countries(path: Path) -> CountryConfig:
    """Load countries and country groups.

    :param path: ``countries.yaml`` path.
    :return: Validated country configuration.
    """

    data = _load_yaml(path)
    countries_raw = data.get("countries") or []
    groups_raw = data.get("country_groups") or {}

    if not isinstance(countries_raw, list) or not all(
        isinstance(value, str) for value in countries_raw
    ):
        raise ValueError("countries must be a list of strings")
    if not countries_raw:
        raise ValueError("At least one country must be configured")
    _ensure_unique_labels(countries_raw, "countries")

    countries = tuple(countries_raw)
    country_set = set(countries)

    if not isinstance(groups_raw, dict):
        raise ValueError("country_groups must be a mapping")

    groups: dict[str, tuple[str, ...]] = {}
    for group_name, members_raw in groups_raw.items():
        if not isinstance(group_name, str) or not group_name.strip():
            raise ValueError("Country-group names must be non-empty strings")
        if not isinstance(members_raw, list) or not all(
            isinstance(member, str) for member in members_raw
        ):
            raise ValueError(f"Country group {group_name!r} must be a list of strings")
        if len(members_raw) != len(set(members_raw)):
            raise ValueError(f"Country group {group_name!r} contains duplicate countries")

        unknown = [member for member in members_raw if member not in country_set]
        if unknown:
            raise ValueError(
                f"Country group {group_name!r} contains unknown countries: {unknown}"
            )
        groups[group_name] = tuple(members_raw)

    if not groups:
        raise ValueError("At least one country group must be configured")

    return CountryConfig(countries=countries, groups=groups)


def _load_workbook_config(path: Path) -> WorkbookConfig:
    """Load source-sheet and generated-sheet configuration.

    :param path: ``sheets.yaml`` path.
    :return: Validated workbook configuration.
    """

    data = _load_yaml(path)
    template = str(data.get("sheet_name_template", "{country}_{sheet}"))
    if "{country}" not in template or "{sheet}" not in template:
        raise ValueError(
            "sheet_name_template must contain both {country} and {sheet} placeholders"
        )

    validation_sheet = str(data.get("validation_sheet", "INTER_VALIDATION"))
    _validate_sheet_name(validation_sheet, "validation_sheet")

    replace_existing = bool(data.get("replace_existing_generated_sheets", True))
    table_spacing_rows = int(data.get("table_spacing_rows", 3))
    if table_spacing_rows < 1:
        raise ValueError("table_spacing_rows must be at least 1")

    add_source_hyperlinks = bool(data.get("add_source_hyperlinks", True))
    kpi_title_separator = str(data.get("kpi_title_separator", " | "))
    round_values = bool(data.get("round_values", False))
    round_digits = int(data.get("round_digits", 0))
    ratio_round_digits = int(data.get("ratio_round_digits", 1))
    if round_digits < 0 or round_digits > 10:
        raise ValueError("round_digits must be between 0 and 10")
    if ratio_round_digits < 0 or ratio_round_digits > 10:
        raise ValueError("ratio_round_digits must be between 0 and 10")

    sources_raw = data.get("sources") or {}
    if not isinstance(sources_raw, dict) or not sources_raw:
        raise ValueError("sheets.yaml must define at least one source sheet")

    sources: dict[str, SheetConfig] = {}
    output_names: list[str] = [validation_sheet]

    for source_name, source_raw in sources_raw.items():
        if not isinstance(source_name, str) or not source_name.strip():
            raise ValueError("Source-sheet names must be non-empty strings")
        if source_raw is None:
            source_raw = {}
        if not isinstance(source_raw, dict):
            raise ValueError(f"Source {source_name!r} must be a mapping")

        enabled = bool(source_raw.get("enabled", False))
        periods_raw = source_raw.get("periods") or []
        output_sheet_raw = source_raw.get("output_sheet")

        if not isinstance(periods_raw, list) or not all(
            isinstance(period, (str, int, float)) for period in periods_raw
        ):
            raise ValueError(f"Periods for source {source_name!r} must be a list")

        periods = tuple(str(period) for period in periods_raw)
        if periods:
            _ensure_unique_labels(list(periods), f"periods for {source_name}")

        output_sheet = str(output_sheet_raw) if output_sheet_raw else None
        if output_sheet:
            _validate_sheet_name(output_sheet, f"output_sheet for {source_name}")
            output_names.append(output_sheet)

        if enabled and not periods:
            raise ValueError(f"Enabled source {source_name!r} must define periods")
        if enabled and not output_sheet:
            raise ValueError(f"Enabled source {source_name!r} must define output_sheet")

        sources[source_name] = SheetConfig(
            name=source_name,
            enabled=enabled,
            periods=periods,
            output_sheet=output_sheet,
        )

    if len(output_names) != len(set(output_names)):
        raise ValueError("Generated output/validation worksheet names must be unique")

    return WorkbookConfig(
        sheet_name_template=template,
        validation_sheet=validation_sheet,
        replace_existing_generated_sheets=replace_existing,
        table_spacing_rows=table_spacing_rows,
        add_source_hyperlinks=add_source_hyperlinks,
        kpi_title_separator=kpi_title_separator,
        round_values=round_values,
        round_digits=round_digits,
        ratio_round_digits=ratio_round_digits,
        sources=sources,
    )


def _optional_text(raw: dict[str, Any], field: str, kpi_name: str) -> str | None:
    """Read one optional text field from a KPI definition.

    :param raw: Raw KPI mapping.
    :param field: Field name.
    :param kpi_name: KPI name for error messages.
    :return: Stripped text or ``None``.
    """

    value = raw.get(field)
    if value is None or value == "":
        return None
    if not isinstance(value, str):
        raise ValueError(f"KPI {field} for {kpi_name!r} must be a string")
    return value.strip() or None


def _load_kpi_ref(value: object, field: str, kpi_name: str) -> KPIRefConfig:
    """Load a simple KPI reference used by a ratio-total rule.

    A reference can be either a KPI name string or a mapping with ``name`` and
    optional one-based ``occurrence`` for duplicate configured KPI names.

    :param value: Raw YAML reference.
    :param field: Reference field name for diagnostics.
    :param kpi_name: Ratio KPI owning the reference.
    :return: Parsed KPI reference.
    """

    if isinstance(value, str):
        name = value.strip()
        occurrence = 1
    elif isinstance(value, dict):
        name_raw = value.get("name")
        if not isinstance(name_raw, str) or not name_raw.strip():
            raise ValueError(
                f"ratio_total.{field} for KPI {kpi_name!r} must define a KPI name"
            )
        name = name_raw.strip()
        occurrence = int(value.get("occurrence", 1))
    else:
        raise ValueError(
            f"ratio_total.{field} for KPI {kpi_name!r} must be a string or mapping"
        )

    if occurrence < 1:
        raise ValueError(
            f"ratio_total.{field}.occurrence for KPI {kpi_name!r} must be >= 1"
        )
    return KPIRefConfig(name=name, occurrence=occurrence)


def _load_ratio_total(raw: dict[str, Any], kpi_name: str) -> RatioTotalConfig | None:
    """Load an optional simple-ratio group-total rule.

    :param raw: Raw KPI YAML mapping.
    :param kpi_name: KPI name for diagnostics.
    :return: Parsed rule or ``None``.
    """

    value = raw.get("ratio_total")
    if value is None:
        return None
    if not isinstance(value, dict):
        raise ValueError(f"ratio_total for KPI {kpi_name!r} must be a mapping")

    if "numerator" not in value or "denominator" not in value:
        raise ValueError(
            f"ratio_total for KPI {kpi_name!r} must define numerator and denominator"
        )

    multiplier = float(value.get("multiplier", 1.0))
    return RatioTotalConfig(
        numerator=_load_kpi_ref(value["numerator"], "numerator", kpi_name),
        denominator=_load_kpi_ref(value["denominator"], "denominator", kpi_name),
        multiplier=multiplier,
        percent=bool(value.get("percent", False)),
    )


def _resolve_kpi_ref_index(
    kpis: tuple[KPIConfig, ...],
    ref: KPIRefConfig,
    owner: KPIConfig,
    source_name: str,
) -> int:
    """Resolve and validate one ratio dependency against configured KPIs.

    :param kpis: KPIs configured for one source sheet.
    :param ref: KPI reference to resolve.
    :param owner: Ratio KPI containing the rule.
    :param source_name: Logical source-sheet name.
    :return: Zero-based configured KPI index.
    :raises ValueError: If the reference is missing, ambiguous, or non-additive.
    """

    matches = [
        kpi for kpi in kpis if normalize_label(kpi.name) == normalize_label(ref.name)
    ]
    if len(matches) < ref.occurrence:
        raise ValueError(
            f"Ratio KPI {owner.name!r} in {source_name!r} references "
            f"{ref.name!r} occurrence {ref.occurrence}, but only {len(matches)} exist"
        )
    target = matches[ref.occurrence - 1]
    if target.aggregation != "sum":
        raise ValueError(
            f"Ratio KPI {owner.name!r} in {source_name!r} must reference additive "
            f"KPIs; {target.name!r} is configured as {target.aggregation!r}"
        )
    return target.index


def resolve_ratio_total_indices(
    kpis: tuple[KPIConfig, ...],
    ratio_kpi: KPIConfig,
    source_name: str,
) -> tuple[int, int] | None:
    """Resolve numerator/denominator KPI indexes for a configured ratio rule.

    :param kpis: KPIs configured for one source sheet.
    :param ratio_kpi: Ratio KPI.
    :param source_name: Logical source-sheet name.
    :return: ``(numerator_index, denominator_index)`` or ``None``.
    """

    rule = ratio_kpi.ratio_total
    if rule is None:
        return None
    return (
        _resolve_kpi_ref_index(kpis, rule.numerator, ratio_kpi, source_name),
        _resolve_kpi_ref_index(kpis, rule.denominator, ratio_kpi, source_name),
    )

def _load_kpis(path: Path, workbook: WorkbookConfig) -> dict[str, tuple[KPIConfig, ...]]:
    """Load KPI configuration by source sheet.

    Duplicate KPI names are valid. Their configuration order is significant and
    is matched against same-name worksheet occurrences from top to bottom.

    :param path: ``kpis.yaml`` path.
    :param workbook: Workbook configuration used for cross-validation.
    :return: KPI definitions keyed by source-sheet name.
    """

    data = _load_yaml(path)
    sources_raw = data.get("sources") or {}
    if not isinstance(sources_raw, dict):
        raise ValueError("kpis.yaml sources must be a mapping")

    unknown_sources = [name for name in sources_raw if name not in workbook.sources]
    if unknown_sources:
        raise ValueError(f"KPI configuration references unknown sources: {unknown_sources}")

    result: dict[str, tuple[KPIConfig, ...]] = {}
    for source_name, kpis_raw in sources_raw.items():
        if not isinstance(kpis_raw, list):
            raise ValueError(f"KPIs for source {source_name!r} must be a list")

        kpis: list[KPIConfig] = []
        for index, raw in enumerate(kpis_raw):
            if not isinstance(raw, dict):
                raise ValueError(
                    f"KPI #{index + 1} for source {source_name!r} must be a mapping"
                )

            name = raw.get("name")
            aggregation = raw.get("aggregation")
            if not isinstance(name, str) or not name.strip():
                raise ValueError(
                    f"KPI #{index + 1} for source {source_name!r} must define a name"
                )
            name = name.strip()
            if aggregation not in SUPPORTED_AGGREGATIONS:
                raise ValueError(
                    f"Unsupported aggregation {aggregation!r} for KPI {name!r}. "
                    f"Allowed values: {sorted(SUPPORTED_AGGREGATIONS)}"
                )

            kpis.append(
                KPIConfig(
                    index=index,
                    name=name,
                    aggregation=aggregation,
                    type=_optional_text(raw, "type", name),
                    subtype=_optional_text(raw, "subtype", name),
                    note=_optional_text(raw, "note", name),
                    ratio_total=_load_ratio_total(raw, name),
                )
            )

        configured = tuple(kpis)
        for kpi in configured:
            if kpi.ratio_total is not None:
                if kpi.aggregation != "ratio":
                    raise ValueError(
                        f"KPI {kpi.name!r} in {source_name!r} defines ratio_total "
                        "but is not configured as aggregation: ratio"
                    )
                resolve_ratio_total_indices(configured, kpi, source_name)
        result[source_name] = configured

    for source_name, source in workbook.sources.items():
        if source.enabled and source_name not in result:
            raise ValueError(
                f"Enabled source {source_name!r} has no KPI definitions in kpis.yaml"
            )

    return result


def load_app_config(config_dir: Path) -> AppConfig:
    """Load and validate all application configuration files.

    :param config_dir: Directory containing the YAML configuration files.
    :return: Complete validated application configuration.
    """

    config_dir = config_dir.resolve()
    project_dir = config_dir.parent

    runtime = _load_runtime(config_dir / "runtime.yaml", project_dir)
    countries = _load_countries(config_dir / "countries.yaml")
    workbook = _load_workbook_config(config_dir / "sheets.yaml")
    kpis = _load_kpis(config_dir / "kpis.yaml", workbook)

    return AppConfig(
        runtime=runtime,
        countries=countries,
        workbook=workbook,
        kpis_by_source=kpis,
    )
