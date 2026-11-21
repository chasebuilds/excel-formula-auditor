"""Formula risk checks over an openpyxl workbook."""

from __future__ import annotations

import re
from collections import Counter
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string

ERROR_CODES = {"#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A"}
ISSUES = (
    "error_value",
    "hardcoded_number",
    "inconsistent_formula",
    "empty_reference",
    "external_link",
    "number_as_text",
)

_STRING = re.compile(r'"[^"]*"')
_ERROR_IN_FORMULA = re.compile(r"#(?:NULL!|DIV/0!|VALUE!|REF!|NAME\?|NUM!|N/A)")
_EXTERNAL = re.compile(r"\[[^\]]*\][^!]*!")
_REF = re.compile(r"(?<![A-Za-z_\d.])(\$?)([A-Z]{1,3})(\$?)(\d+)(?![\d(])")
_SHEET_REF = re.compile(
    r"(?<![\]\w.])(?:(?:'([^']+)'|([A-Za-z_][\w.]*))!)?"
    r"(\$?[A-Z]{1,3}\$?\d+)(?![\d(])(:\$?[A-Z]{1,3}\$?\d+)?"
)
_NUMBER = re.compile(r"(?<![A-Za-z_\d.$])\d+(?:\.\d+)?(?![\d.A-Za-z_(])")
_TEXT_NUMBER = re.compile(r"^\s*[-+]?(?:\d{1,3}(?:,\d{3})+|\d+)?(?:\.\d+)?\s*$")


@dataclass(frozen=True)
class Finding:
    sheet: str
    cell: str
    issue: str
    detail: str


def _is_formula(value: Any) -> bool:
    return isinstance(value, str) and value.startswith("=")


def relative_form(formula: str, row: int, col: int) -> str:
    """Rewrite references as offsets from the cell, like R1C1, so copies compare equal."""

    def sub(m: re.Match[str]) -> str:
        cabs, letters, rabs, digits = m.groups()
        c = column_index_from_string(letters)
        r = int(digits)
        cpart = f"C{c}" if cabs else f"C[{c - col}]"
        rpart = f"R{r}" if rabs else f"R[{r - row}]"
        return rpart + cpart

    return _REF.sub(sub, _STRING.sub('""', formula))


def hardcoded_numbers(formula: str) -> list[str]:
    body = _STRING.sub('""', formula)
    body = _ERROR_IN_FORMULA.sub("E", body)
    body = re.sub(r"'[^']*'!", "S!", body)
    body = re.sub(r"\[[^\]]*\]", "", body)
    body = _REF.sub("R", body)
    return [n for n in _NUMBER.findall(body) if float(n) not in (0.0, 1.0)]


def audit_workbook(wb: Any) -> list[Finding]:
    findings: list[Finding] = []
    for ws in wb.worksheets:
        formulas: dict[tuple[int, int], str] = {}
        for row in ws.iter_rows():
            for cell in row:
                v = cell.value
                if _is_formula(v):
                    formulas[(cell.row, cell.column)] = v
        rel = {k: relative_form(f, k[0], k[1]) for k, f in formulas.items()}

        for row in ws.iter_rows():
            for cell in row:
                v = cell.value
                coord = cell.coordinate

                def add(issue: str, detail: str, sheet: str = ws.title, c: str = coord) -> None:
                    findings.append(Finding(sheet, c, issue, detail))

                if isinstance(v, str) and v.strip() in ERROR_CODES:
                    add("error_value", v)
                    continue
                if not _is_formula(v):
                    if (
                        isinstance(v, str)
                        and _TEXT_NUMBER.match(v)
                        and any(ch.isdigit() for ch in v)
                    ):
                        add("number_as_text", repr(v))
                    continue
                key = (cell.row, cell.column)
                if _ERROR_IN_FORMULA.search(v):
                    add("error_value", v)
                external = bool(_EXTERNAL.search(_STRING.sub('""', v)))
                if external:
                    add("external_link", v)
                nums = hardcoded_numbers(v)
                if nums:
                    add("hardcoded_number", f"{v}  (constants: {', '.join(nums)})")
                if _inconsistent(rel, key):
                    add("inconsistent_formula", v)
                if not external:
                    for empty in _empty_refs(wb, ws, v):
                        add("empty_reference", f"{v}  (empty: {empty})")
    return findings


def _inconsistent(rel: dict[tuple[int, int], str], key: tuple[int, int]) -> bool:
    r, c = key
    mine = rel[key]
    for a, b in (((r - 1, c), (r + 1, c)), ((r, c - 1), (r, c + 1))):
        x, y = rel.get(a), rel.get(b)
        if x is not None and x == y and x != mine:
            return True
    return False


def _empty_refs(wb: Any, ws: Any, formula: str) -> list[str]:
    out: list[str] = []
    for m in _SHEET_REF.finditer(_STRING.sub('""', formula)):
        quoted, bare, ref, rng = m.groups()
        if rng:
            continue
        name = quoted or bare
        if name is not None and name not in wb.sheetnames:
            continue
        target = wb[name] if name else ws
        if target[ref.replace("$", "")].value is None:
            out.append(f"{name}!{ref}" if name else ref)
    return out


def audit(path: str | Path) -> list[Finding]:
    return audit_workbook(load_workbook(path))


def summary_line(findings: list[Finding], path: str | Path = "") -> str:
    counts = Counter(f.issue for f in findings)
    parts = ", ".join(f"{counts[i]} {i}" for i in ISSUES if counts[i])
    sheets = len({f.sheet for f in findings})
    head = f"{Path(path).name}: " if path else ""
    return f"{head}{len(findings)} findings in {sheets} sheets" + (f": {parts}" if parts else "")

