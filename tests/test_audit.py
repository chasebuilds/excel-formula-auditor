from pathlib import Path

from openpyxl import Workbook, load_workbook

from excel_formula_auditor.audit import audit, hardcoded_numbers, relative_form, summary_line
from excel_formula_auditor.cli import main


def _book(tmp_path: Path) -> Path:
    wb = Workbook()
    ws = wb.active
    ws.title = "S"
    for r in range(1, 6):
        ws[f"A{r}"] = r
        ws[f"B{r}"] = f"=A{r}*2"
    ws["B3"] = "=A3+A1"
    ws["C1"] = "#DIV/0!"
    ws["C2"] = "1,234"
    ws["D1"] = "=[1]Rates!B2"
    ws["D2"] = "=A1+Z9"
    ws["D3"] = '="12"&A1'
    p = tmp_path / "b.xlsx"
    wb.save(p)
    return p


def test_each_issue_found(tmp_path: Path) -> None:
    got = {(f.cell, f.issue) for f in audit(_book(tmp_path))}
    assert ("B3", "inconsistent_formula") in got
    assert ("C1", "error_value") in got
    assert ("C2", "number_as_text") in got
    assert ("D1", "external_link") in got
    assert ("D2", "empty_reference") in got
    assert ("B2", "hardcoded_number") in got
    assert not any(c == "D3" for c, _ in got)


def test_relative_form_matches_copies() -> None:
    assert relative_form("=A2*$B$1", 2, 3) == relative_form("=A5*$B$1", 5, 3)
    assert hardcoded_numbers("=SUM(A1:A9)*1.08+1") == ["1.08"]
    assert hardcoded_numbers('=LOG10(A1)&"7"') == []


def test_cli_writes_report(tmp_path: Path, capsys: object) -> None:
    out = tmp_path / "f.xlsx"
    assert main([str(_book(tmp_path)), "-o", str(out)]) == 0
    ws = load_workbook(out)["Findings"]
    assert [c.value for c in ws[1]] == ["sheet", "cell", "issue", "formula"]
    assert ws.max_row > 5
    assert summary_line([]) == "0 findings in 0 sheets"
