"""Score the auditor on seeded workbooks and write BENCHMARK.md plus the README table."""

from __future__ import annotations

import io
import re
import time
from collections import Counter
from dataclasses import dataclass, field
from pathlib import Path

from openpyxl import load_workbook

from benchmark.generate import generate
from excel_formula_auditor.audit import ISSUES, audit_workbook

ROOT = Path(__file__).resolve().parent.parent


@dataclass
class Score:
    books: int
    planted: Counter[str] = field(default_factory=Counter)
    caught: Counter[str] = field(default_factory=Counter)
    extra_on_planted: int = 0
    clean_flags: int = 0
    clean_cells: int = 0
    seconds: float = 0.0


def _roundtrip(wb: object) -> object:
    buf = io.BytesIO()
    wb.save(buf)  # type: ignore[attr-defined]
    buf.seek(0)
    return load_workbook(buf)


def score(books: int, rows: int) -> Score:
    s = Score(books)
    t = time.perf_counter()
    for seed in range(books):
        clean, planted = generate(seed, rows)
        cwb = _roundtrip(clean.wb)
        s.clean_flags += len(audit_workbook(cwb))
        s.clean_cells += sum(ws.max_row * ws.max_column for ws in cwb.worksheets)  # type: ignore[attr-defined]
        found = {(f.sheet, f.cell, f.issue) for f in audit_workbook(_roundtrip(planted.wb))}
        for p in planted.planted:
            s.planted[p[2]] += 1
            s.caught[p[2]] += p in found
        s.extra_on_planted += len(found - planted.planted)
    s.seconds = time.perf_counter() - t
    return s


def table(s: Score) -> str:
    lines = ["| Issue type | Caught |", "|---|---|"]
    for i in ISSUES:
        p, c = s.planted[i], s.caught[i]
        lines.append(f"| {i} | **{100 * c / p:.1f}%** ({c} of {p}) |")
    lines.append(f"| False flags on {s.books} clean workbooks | {s.clean_flags} "
                 f"(of {s.clean_cells:,} cells) |")
    lines.append(f"| Unplanted flags on {s.books} planted workbooks | {s.extra_on_planted} |")
    lines.append(f"| Time for {2 * s.books} workbooks | {s.seconds:.1f} s |")
    return "\n".join(lines)


def main() -> None:
    books, rows = 50, 200
    s = score(books, rows)
    t = table(s)
    (ROOT / "BENCHMARK.md").write_text(
        "# Benchmark\n\n"
        f"Written by `python -m benchmark.run_benchmark`. Seeds 0 to {books - 1}; each seed makes "
        f"one clean workbook and one planted workbook with a {rows}-row Data sheet and a "
        "Summary sheet. Every workbook is saved and reloaded before scanning. A planted issue "
        "counts as caught only when the auditor reports that exact sheet, cell and issue type.\n\n"
        f"{t}\n\n"
        "The plants are written by the same person who wrote the checks, so this is a best case. "
        "See LIMITATIONS.md.\n",
        encoding="utf-8",
    )
    readme = ROOT / "README.md"
    text = readme.read_text(encoding="utf-8")
    text = re.sub(r"(<!-- benchmark:start -->\n).*?(\n<!-- benchmark:end -->)",
                  lambda m: m.group(1) + t + m.group(2), text, flags=re.S)
    readme.write_text(text, encoding="utf-8")
    print(t)


if __name__ == "__main__":
    main()
