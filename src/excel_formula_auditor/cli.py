"""Command line entry point: xlaudit book.xlsx -o findings.xlsx."""

from __future__ import annotations

import argparse

from excel_formula_auditor.audit import audit, summary_line
from excel_formula_auditor.report import write_report


def main(argv: list[str] | None = None) -> int:
    p = argparse.ArgumentParser(prog="xlaudit", description="Scan a workbook for formula risks.")
    p.add_argument("workbook")
    p.add_argument("-o", "--output", help="write findings to this .xlsx")
    a = p.parse_args(argv)
    findings = audit(a.workbook)
    line = summary_line(findings, a.workbook)
    if a.output:
        write_report(findings, a.output, line)
    print(line)
    return 0
