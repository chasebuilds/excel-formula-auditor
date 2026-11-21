# excel-formula-auditor

[![ci](https://github.com/chasebuilds/excel-formula-auditor/actions/workflows/ci.yml/badge.svg)](https://github.com/chasebuilds/excel-formula-auditor/actions/workflows/ci.yml) ![Python 3.11+](https://img.shields.io/badge/python-3.11%2B-blue) [![MIT license](https://img.shields.io/badge/license-MIT-green)](LICENSE)

Scans an .xlsx workbook and reports formula risks: error values (#REF!, #DIV/0!, #N/A, #VALUE! and the rest), hard-coded numbers inside formulas, formulas that differ from their row or column neighbors, references to empty cells, links to other workbooks, and numbers stored as text. The result is an Excel report (sheet, cell, issue, formula) and a one-line summary. Everything runs offline.

## Measured result

Seeded synthetic workbooks with known planted issues were scanned and every finding was checked against the planted list. The table is written by the benchmark script from a real run. Method and limits: [BENCHMARK.md](BENCHMARK.md).

<!-- benchmark:start -->
<!-- benchmark:end -->

The plants are clean and spaced apart, so this is a best case. [LIMITATIONS.md](LIMITATIONS.md) lists what a real workbook can do to it.

## Quick start

Windows (on macOS or Linux use `.venv/bin/` in place of `.venv\Scripts\`):

```
python -m venv .venv
.venv\Scripts\python -m pip install -e .
.venv\Scripts\xlaudit book.xlsx -o findings.xlsx
```

It prints a summary line such as:

```
book.xlsx: 6 findings in 2 sheets: 2 error_value, 1 hardcoded_number, 1 inconsistent_formula, 1 external_link, 1 number_as_text
```

## Checks

| Issue | What it means |
|---|---|
| error_value | A cell holds an Excel error, or a formula contains one such as `#REF!` |
| hardcoded_number | A formula contains a constant other than 0 or 1, like `*1.08` |
| inconsistent_formula | Neighbors above and below (or left and right) share one formula pattern and this cell breaks it |
| empty_reference | A formula points at a single cell that is empty |
| external_link | A formula reads from another workbook (`[1]Sheet!A1`) |
| number_as_text | A text cell that looks like a number, like `"1234"` |

## Development

```
.venv\Scripts\python -m pip install -e ".[dev]"
.venv\Scripts\python -m ruff check .
.venv\Scripts\python -m mypy
.venv\Scripts\python -m pytest --cov
.venv\Scripts\python -m benchmark.run_benchmark
```

## License

MIT, Chase Builds.
