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
