window.DEMO = {
  repo: "excel-formula-auditor",
  title: "excel-formula-auditor",
  tagline: "Drop in an Excel workbook. Get a list of formula risks: error values, hardcoded numbers in formulas, a formula that breaks the pattern of its column, external links, empty references and numbers stored as text.",
  pkg: "excel_formula_auditor",
  packages: ["openpyxl"],
  inputs: [
    { id: "book", label: "Excel workbook (.xlsx or .xlsm)", accept: ".xlsx,.xlsm", sample: "samples/book.xlsx" },
  ],
  options: [],
  sampleNote: 'The sample is <a href="samples/book.xlsx">book.xlsx</a>: a synthetic two-sheet workbook made by the benchmark generator (seed 11) with 10 planted risks. All 10 are found.',
  python: `
from pathlib import Path
from excel_formula_auditor.cli import main as cli_main

def run(paths, opts):
    src = paths["book"][0]
    out = "/work/" + Path(src).stem + "-findings.xlsx"
    code, log = run_cli(cli_main, [src, "-o", out])
    if code != 0 or not Path(out).exists():
        return {"ok": False, "log": log, "sheets": [], "download": None}
    return {"ok": True, "log": log, "sheets": sheets_from_xlsx(out), "download": out}
`,
};
