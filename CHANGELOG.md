# Changelog

## v2

- Intermediary country values now use Excel formulas linked to resolved source cells.
- Optional internal hyperlinks jump from intermediary cells to source locations.
- Country-group totals are native Excel `SUM` formulas.
- Added optional display-only KPI `type` and `subtype` fields.
- Duplicate KPI names are allowed and matched in configuration order to worksheet order.
- Added `KPI_OCCURRENCE_COUNT_MISMATCH` validation when duplicate counts differ.
- Generated workbooks request automatic/full Excel recalculation on open.
