# POPS intermediary-table generator — v3

This project creates auditable intermediary tables in a copy of the Consolidated
POPS workbook.

## Main behavior

- Every country/source worksheet is scanned independently; KPI rows and period
  columns are located independently per country.
- Intermediary country values are Excel formulas linked to the detected source
  cells.
- `add_source_hyperlinks` in `config/sheets.yaml` controls only the on-click
  navigation. Setting it to `false` removes the clickable jump while **keeping
  the source location in the Excel formula**.
- `TOTAL <group>` rows are native Excel `SUM(...)` formulas over the configured
  intermediary country rows.
- KPI `type` and `subtype` are optional display-only YAML fields.
- If `type`/`subtype` repeats the KPI name, the title shows the identical text
  only once.
- Repeated KPI names are supported and matched positionally:
  configuration order -> worksheet order from top to bottom, separately for
  every country.
- If a country has a different number of occurrences for a repeated KPI name,
  no positional assignment is made for that name and
  `KPI_OCCURRENCE_COUNT_MISMATCH` is written to `INTER_VALIDATION`.

## Source navigation switch

In `config/sheets.yaml`:

```yaml
add_source_hyperlinks: true
```

Use `true` to make intermediary values clickable and jump to their source cells.
Use `false` to disable that click behavior. In both cases, a resolved value still
contains a formula such as:

```excel
=IFERROR(IF('Belgium_ID Card'!J37="","",'Belgium_ID Card'!J37),"")
```

The guard preserves source blanks/errors as blank intermediary cells while a
real source zero remains zero.

## KPI catalog

`config/kpis.yaml` now contains the complete ordered KPI catalog supplied for:

- ID Card
- OBS KPI
- FTE
- CVC
- GRANTING
- CS
- CORE
- NPL SALES

`type` and `subtype` are display metadata only and never determine which source
cell is selected.

The current engine still implements additive aggregation only:

- `aggregation: sum` -> linked intermediary table + Excel group SUM formulas
- `aggregation: ratio` -> recognized but deliberately not aggregated yet
- `aggregation: skip` -> catalogued non-additive/ambiguous KPI awaiting its
  approved business aggregation rule

The initial `sum` / `ratio` / `skip` classifications in the expanded catalog are
explicit configuration and should be reviewed with the business before
production use.

## Enabling the additional source sheets

Only the ID Card period/header list was previously provided. Therefore ID Card
remains enabled, while the new source KPI catalogs are present but their source
sheets remain disabled in `config/sheets.yaml` rather than guessing period
headers.

For example, once the exact GRANTING period labels are known:

```yaml
GRANTING:
  enabled: true
  output_sheet: INTER_GRANTING
  periods:
    - "2024"
    - "2025"
    - "PLAN 2026"
```

No Python change is required.

## Formula calculation

The generated workbook requests automatic/full recalculation when opened in
Excel. `openpyxl` itself does not calculate Excel formulas.

## Run

1. Configure the input/output workbook in `config/runtime.yaml`.
2. Fill TOP8/TOP9 memberships in `config/countries.yaml`.
3. Review/enable the desired source sheets and their period labels in
   `config/sheets.yaml`.
4. Install dependencies:

```bash
pip install -r requirements.txt
```

5. Run:

```bash
python main.py
```


## Rounding

`config/sheets.yaml` now contains:

```yaml
round_values: true
round_digits: 0
ratio_round_digits: 1
```

When enabled, generated country values and totals use Excel `ROUND`. Ordinary
numeric KPIs use `round_digits` (default `0`), while ratio KPIs use
`ratio_round_digits` (default `1`) and are displayed as percentages. For example,
`12.34%` becomes `12.3%`, while `4331.9` becomes `4332`. Excel's underlying
percentage representation remains decimal. Set `round_values: false` to preserve
the source decimals and formats.

## Simple ratio totals

All configured `aggregation: ratio` KPIs now get intermediary country rows linked
to the already-computed ratio cell in each country sheet. Group totals are **not**
summed or averaged. For a simple ratio, configure its additive components:

```yaml
- name: "NBI / Outstanding (%)"
  aggregation: ratio
  ratio_total:
    numerator: "NBI (M€)"
    denominator: "OUTSTANDING - Av. (M€)"
    percent: true
```

Then `TOTAL TOP8`, `TOTAL TOP9`, and `TOTAL ALL` are Excel formulas equivalent to
`TOTAL <group> NBI / TOTAL <group> Outstanding`, referencing the intermediary
additive total rows. Ratio KPIs without `ratio_total` still show country values,
but group totals remain blank and `INTER_VALIDATION` reports
`RATIO_TOTAL_RULE_MISSING`.
