# Benchmark

Written by `python -m benchmark.run_benchmark`. Seeds 0 to 49; each seed makes one clean workbook and one planted workbook with a 200-row Data sheet and a Summary sheet. Every workbook is saved and reloaded before scanning. A planted issue counts as caught only when the auditor reports that exact sheet, cell and issue type.

| Issue type | Caught |
|---|---|
| error_value | **100.0%** (144 of 144) |
| hardcoded_number | **100.0%** (100 of 100) |
| inconsistent_formula | **100.0%** (100 of 100) |
| empty_reference | **100.0%** (50 of 50) |
| external_link | **100.0%** (50 of 50) |
| number_as_text | **100.0%** (99 of 99) |
| False flags on 50 clean workbooks | 0 (of 61,000 cells) |
| Unplanted flags on 50 planted workbooks | 2 |
| Time for 100 workbooks | 4.4 s |

The plants are written by the same person who wrote the checks, so this is a best case. See LIMITATIONS.md.
