"""
The audit engine.

Runs the reference-sheet checks, then the column rules from config.FIELD_RULES,
then the cross-sheet conclusions.  It reports progress through an optional
callback so the interface can show what it is doing.
"""

from __future__ import annotations

from .checks import CHECKS, FieldContext
from .config import FIELD_RULES, SUPPORT_COLUMNS
from .models import AuditResult, Issue
from .references import (
    build_category_reference,
    build_condition_reference,
    build_office_reference,
    build_supplier_reference,
    office_countries,
)
from .utils import guess_country
from .workbook import load_workbook, sheet_overview

PROGRESS_STEPS = [
    "Reading workbook",
    "Checking reference sheets",
    "Checking Asset | GPE Information",
    "Cross-checking Supplier",
    "Cross-checking Office",
    "Checking Asset Category",
    "Checking Asset | GPE Condition",
    "Generating feedback",
]


def audit(source, filename: str = "", country: str = "", progress=None) -> AuditResult:
    """Audit a workbook and return the findings.  The file is never modified."""

    def step(name):
        if progress:
            progress(name)

    step("Reading workbook")
    overview = sheet_overview(source)
    try:
        source.seek(0)
    except (AttributeError, OSError):
        pass
    book = load_workbook(source, filename=filename)

    result = AuditResult(
        filename=book.filename,
        country=country,
        sheet_overview=overview,
        resolved_sheets=dict(book.resolved),
        missing_sheets=list(book.missing),
    )
    issues: list[Issue] = []

    for label in book.missing:
        issues.append(Issue(
            section="Missing Sheets", code="missing_sheet",
            title=f"Reference sheet not found: {label}",
            count=1, values=[label],
        ))

    # ---- reference sheets first ------------------------------------------
    step("Checking reference sheets")
    conditions, _disposal, office_types = build_condition_reference(book, issues)
    suppliers = build_supplier_reference(book, issues)
    offices = build_office_reference(book, issues, office_types=office_types)
    category_names, category_codes, category_pairs = build_category_reference(book, issues)

    references = {
        "supplier": suppliers,
        "office": offices,
        "category_name": category_names,
        "category_code": category_codes,
        "condition": conditions,
    }

    if not result.country:
        result.country = guess_country(book.filename, office_countries(book))

    # ---- the Asset | GPE Information sheet -------------------------------
    asset_sheet = book.get("asset_info")
    if asset_sheet is None or not asset_sheet.rows:
        if asset_sheet is not None:
            issues.append(Issue(
                section="Missing Sheets", code="empty_asset_sheet",
                title="The Asset | GPE Information sheet has no data",
                sheet=asset_sheet.name, count=1,
            ))
        result.issues = issues
        step("Generating feedback")
        return result

    result.asset_row_count = len(asset_sheet.rows)

    columns: dict[str, str] = {}
    for rule in FIELD_RULES:
        found = asset_sheet.find_column(rule["column"])
        if found:
            columns[rule["key"]] = found
        else:
            issues.append(Issue(
                section="Missing Sheets", code="missing_column",
                title=f"Column not found: {rule['section']}",
                sheet=asset_sheet.name, count=1, values=[rule["section"]],
            ))

    support = {}
    for key, candidates in SUPPORT_COLUMNS.items():
        found = asset_sheet.find_column(candidates)
        if found:
            support[key] = found

    step_for_rule = {
        "supplier": "Cross-checking Supplier",
        "office": "Cross-checking Office",
        "asset_category": "Checking Asset Category",
        "asset_category_code": "Checking Asset Category",
        "condition": "Checking Asset | GPE Condition",
    }
    step("Checking Asset | GPE Information")
    announced = set()

    for rule in FIELD_RULES:
        column = columns.get(rule["key"])
        if not column:
            continue
        name = step_for_rule.get(rule["key"])
        if name and name not in announced:
            announced.add(name)
            step(name)

        context = FieldContext(
            sheet=asset_sheet, rule=rule, column=column,
            references=references, columns=columns, support=support,
            category_pairs=category_pairs,
        )
        for check_name in rule["checks"]:
            check = CHECKS.get(check_name)
            if check is None:
                continue
            issues.extend(check(context))

    # ---- cross-sheet conclusion: values that must be added to a reference --
    _add_missing_reference_values(issues, "supplier", "Supplier Reference Sheet",
                                  suppliers)
    _add_missing_reference_values(issues, "office", "Office Reference Sheet",
                                  offices)

    step("Generating feedback")
    result.issues = issues
    return result


def _add_missing_reference_values(issues, section_key, reference_section, reference):
    """Turn "not in the reference sheet" findings into an action on that sheet."""
    if not reference.available:
        return
    field_section = "Supplier" if section_key == "supplier" else "Office"
    values, count = [], 0
    for issue in issues:
        if issue.section == field_section and issue.code == "not_in_reference":
            values.extend(issue.values)
            count += issue.count
    if not values:
        return
    issues.append(Issue(
        section=reference_section, code="ref_missing",
        title=f"{field_section} values missing from {reference.label}",
        sheet=reference.sheet_name,
        count=count, values=sorted(set(values)),
    ))
