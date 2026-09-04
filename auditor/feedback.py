"""
Feedback generator.

Turns the audit findings into a short, plain-English document written for a
data-entry team: what is wrong, and what they need to do.

Sections with no findings are left out entirely.  Counts, percentages and row
numbers are deliberately not included.
"""

from __future__ import annotations

import datetime as _dt

from .config import (
    DEFAULT_SECTION_INTRO,
    EXAMPLES_IN_TEXT,
    FALLBACK_TEXT,
    FEEDBACK_TEXT,
    MAX_LISTED_VALUES,
    SECTION_INTRO,
)
from .utils import quote_list

# Order of the problem clauses inside a sentence, taken from the config file.
_CODE_ORDER = {key: index for index, key in enumerate(FEEDBACK_TEXT)}


def _join(parts) -> str:
    parts = [p for p in parts if p]
    if not parts:
        return "problems that need correcting"
    if len(parts) == 1:
        return parts[0]
    return ", ".join(parts[:-1]) + " and " + parts[-1]


def _missing_sheets_section(issues):
    """Structural problems: the audit cannot map anything without these."""
    sheets = [v for i in issues if i.code == "missing_sheet" for v in i.values]
    columns = [v for i in issues if i.code == "missing_column" for v in i.values]
    empty = [i for i in issues if i.code == "empty_asset_sheet"]

    problems, actions = [], []
    if sheets:
        problems.append("the following reference sheets could not be found: "
                        + ", ".join(sheets))
        actions.append("Add the missing reference sheets to the workbook using "
                       "the standard template, and fill them in before "
                       "completing the Asset | GPE Information sheet.")
    if columns:
        problems.append("the following columns could not be found in the "
                        "Asset | GPE Information sheet: " + ", ".join(columns))
        actions.append("Restore the missing columns using the standard "
                       "template, keeping the original column headings.")
    if empty:
        problems.append("the Asset | GPE Information sheet contains no asset rows")
        actions.append("Fill in the Asset | GPE Information sheet with the "
                       "asset and GPE records.")

    if not problems:
        return None
    return {
        "title": "Workbook Structure",
        "problem": "This workbook does not follow the standard template: "
                   + _join(problems) + ".",
        "actions": actions,
        "details": [],
    }


def _section_body(section, issues, include_full_lists=False):
    """Build the problem sentence, the bullets and any necessary value list."""
    ordered = sorted(issues, key=lambda i: _CODE_ORDER.get((section, i.code), 999))

    problems, actions, details = [], [], []
    for issue in ordered:
        text = FEEDBACK_TEXT.get((section, issue.code), FALLBACK_TEXT)
        examples = quote_list(issue.values, EXAMPLES_IN_TEXT)
        if "{examples}" in text["problem"] and not examples:
            continue
        problems.append(text["problem"].format(examples=examples))
        for action in text["actions"]:
            filled = action.format(examples=examples)
            if filled not in actions:
                actions.append(filled)

        # The only list worth printing: values that must be added to a
        # reference sheet, or an option list the team must choose from.
        if issue.code == "ref_missing" and issue.values:
            limit = len(issue.values) if include_full_lists else MAX_LISTED_VALUES
            shown = issue.values[:limit]
            heading = ("Missing from this sheet:"
                       if len(shown) == len(issue.values)
                       else f"Missing from this sheet (first {len(shown)}):")
            details.append({"heading": heading, "items": shown,
                            "more": len(issue.values) - len(shown)})
        elif issue.code in ("ref_type_invalid", "not_in_reference") and issue.note:
            details.append({"heading": issue.note, "items": list(issue.note_items),
                            "more": 0})

    if not problems:
        return None

    intro = SECTION_INTRO.get(section, DEFAULT_SECTION_INTRO)
    return {
        "title": section,
        "problem": intro.format(section=section, problems=_join(problems)),
        "actions": actions,
        "details": details,
    }


def build_feedback(result, include_full_lists=False, date=None) -> dict:
    """Return the feedback document as a structure the renderers can use."""
    grouped = result.by_section()
    sections = []

    for section, issues in grouped.items():
        body = (_missing_sheets_section(issues) if section == "Missing Sheets"
                else _section_body(section, issues, include_full_lists))
        if body:
            sections.append(body)

    for number, section in enumerate(sections, start=1):
        section["number"] = number

    country = (result.country or "").strip()
    return {
        "title": f"Feedback on {country} Data" if country else "Feedback on the Uploaded Data",
        "date": (date or _dt.date.today()).strftime("%d %B %Y"),
        "sections": sections,
        "clean": not sections,
    }


def render_text(document: dict) -> str:
    """Render the feedback as plain text (also used for copy and .txt)."""
    lines = [document["title"], f"Date: {document['date']}", ""]

    if document["clean"]:
        lines.append("No data-quality issues were found in this workbook. "
                     "No corrections are needed.")
        return "\n".join(lines)

    for section in document["sections"]:
        lines.append(f"{section['number']}. {section['title']}")
        lines.append(section["problem"])
        lines.append("")
        lines.append("Please:")
        for action in section["actions"]:
            lines.append(f"- {action}")
        for detail in section["details"]:
            lines.append("")
            lines.append(detail["heading"])
            for item in detail["items"]:
                lines.append(f"  - {item}")
            if detail["more"]:
                lines.append(f"  - ... and {detail['more']} more")
        lines.append("")

    return "\n".join(lines).rstrip() + "\n"


def render_markdown(document: dict) -> str:
    """Render the feedback as Markdown, for on-screen display."""
    lines = [f"### {document['title']}", f"**Date:** {document['date']}", ""]

    if document["clean"]:
        lines.append("No data-quality issues were found in this workbook. "
                     "No corrections are needed.")
        return "\n".join(lines)

    for section in document["sections"]:
        lines.append(f"**{section['number']}. {section['title']}**")
        lines.append("")
        lines.append(section["problem"])
        lines.append("")
        lines.append("Please:")
        lines.append("")
        for action in section["actions"]:
            lines.append(f"- {action}")
        for detail in section["details"]:
            lines.append("")
            lines.append(f"*{detail['heading']}*")
            lines.append("")
            for item in detail["items"]:
                lines.append(f"  - {item}")
            if detail["more"]:
                lines.append(f"  - *... and {detail['more']} more*")
        lines.append("")

    return "\n".join(lines)
