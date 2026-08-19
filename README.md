# POPS intermediary-table generator

This version creates auditable intermediary tables in the existing Consolidated
POPS workbook.

## What changed

- Country values are Excel formulas linked to the detected source cells.
- The linked cells can also be clicked to jump to the source location.
- `TOTAL <group>` rows are native Excel `SUM(...)` formulas.
- KPI `type` and `subtype` are optional display-only YAML fields.
- Repeated KPI names are supported and matched positionally:
  configuration order -> worksheet order from top to bottom.
- If the number of same-name KPI occurrences differs between configuration and a
  country worksheet, no positional assignment is made for that name and
  `KPI_OCCURRENCE_COUNT_MISMATCH` is written to `INTER_VALIDATION`.

## Formula behavior

For a resolved source such as `Belgium_ID Card!J37`, the intermediary cell uses
an Excel formula equivalent to:

```excel
=IFERROR(IF('Belgium_ID Card'!J37="","",'Belgium_ID Card'!J37),"")
```

The guard is intentional: a plain direct reference to an empty Excel cell would
show `0`. Here, an actual source zero stays zero, while a blank/error stays blank.
Validation still records source problems.

Group totals reference the intermediary rows, for example:

```excel
=SUM(B3,B4,B6,B8)
```

The workbook is marked for automatic/full recalculation when opened in Excel.
`openpyxl` itself does not calculate formulas.

## Optional KPI display metadata

```yaml
- name: "Number of accounts"
  type: "OUTSTANDING"
  subtype: "AUTO"
  aggregation: sum

- name: "Number of accounts"
  type: "PRODUCTION"
  subtype: "AUTO"
  aggregation: sum
```

The titles become:

```text
OUTSTANDING | AUTO | Number of accounts
PRODUCTION | AUTO | Number of accounts
```

`type` and `subtype` do not participate in matching. They are only display text.

## Run

1. Put the source workbook at the path configured in `config/runtime.yaml`.
2. Fill `TOP8`/`TOP9` memberships in `config/countries.yaml`.
3. Install dependencies:

```bash
pip install -r requirements.txt
```

4. Run:

```bash
python main.py
```
