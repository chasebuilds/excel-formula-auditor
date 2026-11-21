"""Write findings to an Excel report."""

from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font

from excel_formula_auditor.audit import Finding


def write_report(findings: list[Finding], out: str | Path, summary: str) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Findings"
    ws.append(["sheet", "cell", "issue", "formula"])
    for c in ws[1]:
        c.font = Font(bold=True)
    for f in findings:
        ws.append([f.sheet, f.cell, f.issue, f.detail])
        ws.cell(ws.max_row, 4).data_type = "s"  # keep formula text as text
    ws.freeze_panes = "A2"
    for col, width in zip("ABCD", (16, 8, 22, 70), strict=True):
        ws.column_dimensions[col].width = width
    s = wb.create_sheet("Summary")
    s.append([summary])
    wb.save(out)
