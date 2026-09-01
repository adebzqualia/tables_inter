# POPS intermediary-table generator — v8

This project creates auditable intermediary KPI tables in a copy of the
Consolidated POPS workbook.

## What changed in v8

The program now includes a **recursive formula-lineage engine**.

A business KPI no longer has to be calculated directly in its source sheet. The
program can follow chains such as:

```text
GRANTING KPI
  -> country_KPI sheet
      -> another KPI formula
          -> NBI
          -> Outstanding
```

or:

```text
source KPI
  -> country_KPI
      -> formula referring back to country_GRANTING / country_CS / country_CORE
```

The engine recursively follows cell references until it reaches semantic KPI
values that can be aggregated. KPI-sheet cells are identified by the column
whose header is configured as `KPI`.

Only supporting KPIs that are actually needed by a derived formula are
materialized in `INTER_DEPENDENCIES`.

## Generated sheets

In addition to the configured `INTER_<SOURCE>` worksheets, v8 generates:

- `INTER_CONFIG`: live Excel controls for ordinary KPI aggregation;
- `INTER_DEPENDENCIES`: automatically discovered prerequisite KPI tables;
- `INTER_VALIDATION`: extraction and formula-lineage diagnostics.

All names are configurable in `config/sheets.yaml`.

## SUM / AVERAGE directly in Excel

For ordinary value KPIs, TOP8/TOP9/ALL aggregation can be changed **after the
Python run**.

`INTER_CONFIG` contains an `Aggregation` column with an Excel dropdown:

```text
SUM
AVERAGE
```

Changing the dropdown immediately changes the corresponding TOTAL formulas. No
Python rerun is required.

The YAML `aggregation: sum` / `aggregation: average` value therefore acts as the
initial default. Automatically discovered leaf dependencies default to
`default_dependency_aggregation` from `sheets.yaml`.

Derived ratios/formulas are marked `FORMULA` and are not aggregated by SUM or
AVERAGE.

## Recursive ratio / derived-formula totals

For a configured `aggregation: ratio` KPI:

1. an explicit `ratio_total` YAML rule still has highest priority;
2. otherwise the recursive lineage engine inspects the real source formula;
3. formula references may cross worksheets and may themselves refer to formulas;
4. KPI-sheet references are converted into semantic dependency KPIs;
5. formulas are compared across countries to make sure the business relation is
   consistent;
6. TOP8/TOP9/ALL are calculated by applying the traced formula to the relevant
   intermediary group totals.

Example source chain:

```excel
'PFF_GRANTING'!J20 = 'PFF_KPI'!J50
'PFF_KPI'!J50      = IF(J40=0,0,J41/J40)
```

where row 41 is `NBI (M€)` and row 40 is `Outstanding (M€)` in the KPI column.

The generated group formula is based on:

```text
TOTAL NBI / TOTAL Outstanding
```

rather than averaging the country percentages.

The same logic works if the formula points back to another operational sheet.
For example, a KPI-sheet formula can depend on a configured `GRANTING`, `CS`,
`CORE`, etc. KPI; the generated formula references that KPI's intermediary TOTAL
row.

## Dependency organization

`INTER_DEPENDENCIES` contains only KPIs discovered while tracing formulas.
Dependencies are ordered with prerequisite/leaf tables before derived tables
where possible.

Each dependency table keeps the same audit behavior as normal intermediary
sheets:

- country cells are Excel formulas pointing to the actual source cells;
- optional click-to-source hyperlinks are preserved;
- ordinary dependency totals use the live SUM/AVERAGE setting from
  `INTER_CONFIG`;
- formula-derived dependency totals use the recursively reconstructed formula.

This means the full calculation chain remains inspectable inside the output
workbook.

## KPI-sheet discovery

The main KPI catalogue is configured in `sheets.yaml`:

```yaml
kpi_source_name: KPI
kpi_name_header: KPI
```

For a referenced cell on `{country}_KPI`, the program:

1. locates the column headed `KPI`;
2. reads the KPI name from the referenced cell's row;
3. identifies the corresponding configured period column;
4. creates/reuses that semantic dependency KPI;
5. resolves the same KPI independently in every country's KPI sheet;
6. recursively analyzes its formula if it has one.

Duplicate KPI names on the KPI sheet are distinguished by top-to-bottom
occurrence number.

## Safe formula handling

The recursive engine is deliberately conservative. It supports:

- single-cell references;
- references to another country-prefixed or logical source sheet;
- nested formulas;
- common Excel functions such as `IF`, `IFERROR`, `ROUND`, etc.;
- small ranges inside `SUM`, `AVERAGE`, `MIN`, `MAX`, `COUNT`, and `COUNTA`;
- constants and same-workbook helper formulas.

It does **not** silently reinterpret unsafe constructs. Unsupported external
workbooks, named/structured references, very large ranges, unresolved helper
values, cycles, or inconsistent formulas are reported in `INTER_VALIDATION`.
The affected group total remains blank rather than being guessed.

Percentage-formatted discovered KPIs without a traceable formula are also left
unaggregated: the program does not average percentages merely because a formula
could not be found.

## Source navigation

```yaml
add_source_hyperlinks: true
```

Set it to `false` to disable click-to-source navigation. The actual source
coordinate remains present in the intermediary Excel formula either way.

## Rounding

```yaml
round_values: true
round_digits: 0
ratio_round_digits: 1
```

Ordinary values use `round_digits`. Ratios/percentages retain one decimal
percentage point by default, e.g. `12.34% -> 12.3%`.

## Main lineage settings

```yaml
recursive_formula_totals: true
dependency_sheet: INTER_DEPENDENCIES
aggregation_control_sheet: INTER_CONFIG
kpi_source_name: KPI
kpi_name_header: KPI
formula_max_depth: 20
default_dependency_aggregation: sum
```

`formula_max_depth` prevents accidental infinite traversal through pathological
formula chains.

## Existing behavior retained

- Every country/source sheet is scanned independently; coordinates are never
  assumed to be identical across countries.
- Repeated configured KPI names are mapped by configuration order to worksheet
  top-to-bottom order.
- Country intermediary values link to the resolved source cells.
- Missing source values remain distinct from actual numeric zero.
- Explicit ratio rules remain supported.
- `type` and `subtype` remain display-only metadata.
- Excel is requested to fully recalculate the generated workbook when opened.

## Run

1. Set the workbook paths in `config/runtime.yaml`.
2. Confirm TOP8/TOP9 memberships in `config/countries.yaml`.
3. Review the initial KPI metadata in `config/kpis.yaml`.
4. Install requirements:

```bash
pip install -r requirements.txt
```

5. Run:

```bash
python main.py
```

6. In the generated workbook, use `INTER_CONFIG` to switch ordinary KPI totals
   between SUM and AVERAGE if required.

## Important Excel limitation

`openpyxl` does not calculate Excel formulas. It preserves/writes formulas and
can read the last cached result saved by Excel. The output workbook therefore
requests automatic full recalculation when it is opened in Excel.
