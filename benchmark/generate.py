"""Seeded synthetic workbooks with known planted formula risks."""

from __future__ import annotations

import random
from dataclasses import dataclass, field
from typing import Any

from openpyxl import Workbook

ERRORS = ["#REF!", "#DIV/0!", "#N/A", "#VALUE!"]


@dataclass
class Book:
    wb: Any
    planted: set[tuple[str, str, str]] = field(default_factory=set)


def clean_book(rng: random.Random, rows: int) -> Book:
    wb = Workbook()
    d = wb.active
    d.title = "Data"
    d.append(["id", "qty", "price", "fee", "subtotal", "total"])
    n = rows + 1
    for r in range(2, n + 1):
        d.append([r - 1, rng.randint(1, 50), round(rng.uniform(2, 90), 2),
                  round(rng.uniform(0, 5), 2), f"=B{r}*C{r}", f"=E{r}+D{r}"])
    s = wb.create_sheet("Summary")
    s.append(["measure", "value"])
    for label, f in [
        ("subtotal", f"=SUM(Data!E2:E{n})"),
        ("total", f"=SUM(Data!F2:F{n})"),
        ("avg qty", f"=AVERAGE(Data!B2:B{n})"),
        ("first line", "=Data!B2*Data!C2"),
        ("max fee", f"=MAX(Data!D2:D{n})"),
        ("rows", f"=COUNT(Data!A2:A{n})"),
    ]:
        s.append([label, f])
    return Book(wb)


def planted_book(rng: random.Random, rows: int) -> Book:
    b = clean_book(rng, rows)
    d, s = b.wb["Data"], b.wb["Summary"]
    n = rows + 1

    def plant(sheet: str, cell: str, issue: str) -> None:
        b.planted.add((sheet, cell, issue))

    # Data rows: spread plants apart so neighbors stay clean.
    picks = rng.sample(range(3, n), 6) if n > 20 else []
    picks.sort()
    picks = [p for i, p in enumerate(picks) if i == 0 or p - picks[i - 1] > 1][:6]
    kinds = ["inconsistent_formula", "inconsistent_formula", "error_value",
             "number_as_text", "number_as_text", "error_value"]
    for r, kind in zip(picks, kinds, strict=False):
        if kind == "inconsistent_formula":
            d[f"E{r}"] = f"=B{r}*D{r}"
            plant("Data", f"E{r}", kind)
        elif kind == "error_value":
            d[f"D{r}"] = rng.choice(ERRORS)
            plant("Data", f"D{r}", kind)
        else:
            d[f"B{r}"] = str(d[f"B{r}"].value)
            plant("Data", f"B{r}", kind)

    # Summary plants on every other row below the clean block.
    row = 10
    for f, issue in [
        (f"=SUM(Data!F2:F{n})*{rng.choice(['1.08', '0.93', '12', '365'])}", "hardcoded_number"),
        (f"=Data!C{rng.randint(2, n)}+{rng.randint(2, 999)}", "hardcoded_number"),
        (f"=[{rng.randint(1, 3)}]Rates!B{rng.randint(2, 9)}*Data!C2", "external_link"),
        (f"=Data!B2+Data!Z{rng.randint(2, n)}", "empty_reference"),
        ("=#REF!+Data!B2", "error_value"),
    ]:
        s[f"A{row}"] = "check"
        s[f"B{row}"] = f
        plant("Summary", f"B{row}", issue)
        row += 2
    return b


def generate(seed: int, rows: int) -> tuple[Book, Book]:
    rng = random.Random(seed)
    return clean_book(rng, rows), planted_book(rng, rows)
