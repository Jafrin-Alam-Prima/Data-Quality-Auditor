"""Excel Data Quality Auditor — audit engine package.

Layers:
    workbook.py    Excel processing: open, detect sheets, detect headers
    references.py  build the allowed-value lists from the reference sheets
    checks.py      the individual checks
    engine.py      runs the rules and returns the findings
    feedback.py    turns findings into plain-English feedback
    exporters.py   .txt / .docx / .pdf downloads
    config.py      all the rules, sheet names, column names and wording
"""

from .engine import PROGRESS_STEPS, audit
from .feedback import build_feedback, render_markdown, render_text
from .models import AuditResult, Issue

__all__ = [
    "audit", "PROGRESS_STEPS",
    "build_feedback", "render_text", "render_markdown",
    "AuditResult", "Issue",
]

__version__ = "1.0.0"
