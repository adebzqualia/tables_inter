# Changelog

## v8

- Added recursive, cross-sheet formula lineage for derived/ratio KPIs.
- Formula tracing can follow operational sheet -> KPI sheet -> other KPI formulas
  -> configured operational KPIs recursively.
- Automatically discovers KPI-sheet dependencies using the `KPI` name column.
- Added `INTER_DEPENDENCIES` containing only prerequisite KPI tables required by
  traced formulas.
- Added `INTER_CONFIG` with live Excel SUM/AVERAGE dropdowns for ordinary value
  KPIs; group totals update without rerunning Python.
- Added `average` as a supported initial KPI aggregation type.
- Formula-derived dependency totals use the reconstructed source formula rather
  than summing/averaging country ratios.
- Added cross-country semantic-formula consensus checks.
- Added cycle, recursion-depth, external-reference, range, and unresolved-lineage
  diagnostics.
- Percentage dependencies without a traceable formula remain blank at group
  level instead of being averaged automatically.

## v7

- Added strict formula-based inference for simple ratio totals.
