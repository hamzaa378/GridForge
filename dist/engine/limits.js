/** Sparse worksheet bounds and storage limits shared across import paths. */
// Match the Calc/Excel address grid; populated data remains sparse and is
// separately bounded so the larger theoretical grid does not allocate densely.
export const MAX_WORKSHEET_ROWS = 1_048_576;
export const MAX_WORKSHEET_COLUMNS = 16_384;
export const MAX_WORKSHEET_POPULATED_CELLS = 2_000_000;
export const MAX_XLSX_CELLS_PER_SHEET = 2_000_000;
/** Excel/Calc do not impose a small fixed sheet count; bound hostile files while allowing practical workbooks. */
export const MAX_WORKBOOK_SHEETS = 256;
/** Aggregate import budget prevents many-sheet workbooks from multiplying per-sheet limits into memory exhaustion. */
export const MAX_WORKBOOK_POPULATED_CELLS = 8_000_000;
/** Bound expanded row-dimension runs from compressed ODS files. */
export const MAX_ODS_DIMENSION_ROWS_PER_SHEET = 100_000;
