# Limitations

- **Formulas only, no recalculation.** openpyxl reads formula text. A formula that would evaluate to an error but has no stored error text is not flagged.
- **Inconsistency needs two matching neighbors.** A cell is flagged only when the cells on both sides (above and below, or left and right) share a pattern it breaks. Edge cells, first and last rows, and two broken cells in a row are missed.
- **Hard-coded numbers are judged by text.** Every constant except 0 and 1 is flagged, including fair ones like `/12` or `*100`. Expect noise on real models.
- **Empty references skip ranges.** `SUM(A1:A10)` with blank cells is not flagged; only single-cell references are checked. Named ranges and defined names are not resolved.
- **External links are matched by the `[n]Sheet!` pattern** in formula text. Links hidden in defined names, charts, or data connections are not found.
- **Number-as-text** only checks plain digits, commas and a decimal point. Currency symbols, percents and dates stored as text are missed.
- **No .xls, .xlsb or .ods.** Only files openpyxl can open.
- **The benchmark is synthetic** and the plants were written alongside the checks, so the catch rates are a best case.
