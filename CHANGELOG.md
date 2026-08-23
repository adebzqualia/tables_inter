# v5

- Added configurable rounding (`round_values`, `round_digits`).
- Ratio KPI country values are now generated as source-linked formulas.
- Added `ratio_total` numerator/denominator rules for group-level ratio aggregation.
- Ratio group totals reference intermediary additive TOTAL rows rather than summing/averaging percentages.
- Added validation warnings for ratio KPIs without a configured group-total rule.
- Added initial explicit simple-ratio rules where numerator/denominator are unambiguous from the supplied KPI catalog.

# Changelog

## v3

- Added/confirmed `add_source_hyperlinks` as an explicit navigation-only switch.
  Source-cell formulas are retained whether clickable navigation is on or off.
- Added the complete ordered KPI catalogs supplied for ID Card, OBS KPI, FTE,
  CVC, GRANTING, CS, CORE, and NPL SALES.
- Added type/subtype metadata throughout those catalogs.
- KPI titles now avoid repeating identical adjacent type/subtype/name text.
- Preserved positional duplicate-name matching independently per country.
- Added output-sheet placeholders for each newly catalogued source.
- Kept non-ID Card sources disabled until their exact period lists are supplied,
  rather than assuming they use ID Card periods.

## v2

- Country intermediary values use native Excel formulas referencing resolved
  source cells.
- Optional internal hyperlinks jump from intermediary cells to source cells.
- Group totals use native Excel `SUM` formulas.
- Added display-only KPI type/subtype fields.
- Added positional duplicate KPI-name handling with count-mismatch validation.
- Generated workbooks request automatic/full Excel recalculation on open.
