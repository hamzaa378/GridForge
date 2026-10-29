# GridForge 11.6.0 Release Notes

## Named ranges and Name Manager

- Define a worksheet name for the active cell or selected range, with an optional explanatory comment.
- Review, edit and delete active-sheet names in Name Manager.
- Create names from selected top-row and/or left-column headings. Duplicate headings receive a unique suffix.
- Insert a selected name from Use in Formula.
- Use names in formulas, including range calculations such as `=SUM(NetSales)`. Changing a definition recalculates dependents and can be undone or redone.
- Preserve supported worksheet-scoped cell/range names in GridForge JSON, native XLSX, and macro-preserving XLSM edits.

The name model is worksheet-scoped and cell/range-based. Workbook-scoped names and names referring to constants or arbitrary formulas are not supported in this release. GridForge remains standalone and does not provide full Excel feature parity.
