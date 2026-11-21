"""Scan Excel workbooks for formula risks."""

from excel_formula_auditor.audit import Finding, audit, audit_workbook

__all__ = ["Finding", "audit", "audit_workbook"]
__version__ = "0.1.0"
