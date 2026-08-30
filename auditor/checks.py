"""
The individual checks.

Each check receives a FieldContext and returns a list of Issues.
Checks run in the order listed in the rule, and each one records the rows it
has already reported so that a single bad cell is never reported twice
(a blank cell is not also reported as "not in the reference sheet").

Adding a new kind of check = write a function here and register it in CHECKS.
"""

from __future__ import annotations

from dataclasses import dataclass, field

from .models import Issue, ReferenceList
from .utils import cell_text, is_blank, is_error_value, is_placeholder, is_zero, norm_key, to_number


@dataclass
class FieldContext:
    sheet: object                       # the Asset | GPE Information Sheet
    rule: dict                          # the rule from config.FIELD_RULES
    column: str                         # real header label found in the sheet
    references: dict                    # key -> ReferenceList
    columns: dict = field(default_factory=dict)   # rule key -> header label
    support: dict = field(default_factory=dict)   # support key -> header label
    category_pairs: dict = field(default_factory=dict)
    handled: set = field(default_factory=set)     # Excel rows already reported

    @property
    def section(self) -> str:
        return self.rule["section"]

    def reference(self) -> ReferenceList | None:
        key = self.rule.get("reference")
        return self.references.get(key) if key else None

    def cells(self):
        """Yield (excel_row, raw_value) for rows not yet reported."""
        for row_number, value in self.sheet.column_values(self.column):
            if row_number not in self.handled:
                yield row_number, value


def _issue(ctx, code, title, rows, values=(), note=""):
    ctx.handled.update(rows)
    return Issue(
        section=ctx.section, code=code, title=title,
        sheet=ctx.sheet.name, column=ctx.column,
        count=len(rows), rows=list(rows),
        values=sorted({cell_text(v) for v in values if cell_text(v)}),
        note=note,
    )


# ---------------------------------------------------------------------------
# Basic value checks
# ---------------------------------------------------------------------------

def check_not_blank(ctx):
    rows = [r for r, v in ctx.cells() if is_blank(v)]
    if not rows:
        return []
    return [_issue(ctx, "blank", f"Blank {ctx.column}", rows)]


def check_no_placeholder(ctx):
    hits = [(r, v) for r, v in ctx.cells()
            if not is_blank(v) and is_placeholder(v) and not is_error_value(v)]
    if not hits:
        return []
    return [_issue(ctx, "placeholder", f"Placeholder values in {ctx.column}",
                   [r for r, _ in hits], [v for _, v in hits])]


def check_no_error_value(ctx):
    hits = [(r, v) for r, v in ctx.cells() if is_error_value(v)]
    if not hits:
        return []
    return [_issue(ctx, "error_value", f"Excel error values in {ctx.column}",
                   [r for r, _ in hits], [v for _, v in hits])]


def check_not_zero(ctx):
    rows = [r for r, v in ctx.cells() if not is_blank(v) and is_zero(v)]
    if not rows:
        return []
    return [_issue(ctx, "zero", f"{ctx.column} recorded as 0", rows, ["0"])]


def check_numeric(ctx):
    """Flag descriptions and other text typed into a value column."""
    hits = [(r, v) for r, v in ctx.cells()
            if not is_blank(v) and to_number(v) is None]
    if not hits:
        return []
    return [_issue(ctx, "not_numeric", f"Text entries in {ctx.column}",
                   [r for r, _ in hits], [v for _, v in hits])]


def check_unique(ctx):
    seen: dict[str, int] = {}
    repeated_rows, repeated_values = [], []
    for row_number, value in ctx.cells():
        if is_blank(value):
            continue
        key = norm_key(value)
        if key in seen:
            repeated_rows.append(row_number)
            repeated_values.append(cell_text(value))
        else:
            seen[key] = row_number
    if not repeated_rows:
        return []
    # Report the first occurrence too, so the team can find every copy.
    first_rows = [seen[norm_key(v)] for v in repeated_values if norm_key(v) in seen]
    return [Issue(
        section=ctx.section, code="duplicate",
        title=f"Duplicate {ctx.column}",
        sheet=ctx.sheet.name, column=ctx.column,
        count=len(repeated_rows),
        values=sorted(set(repeated_values)),
        rows=sorted(set(repeated_rows) | set(first_rows)),
    )]


# ---------------------------------------------------------------------------
# Cross-sheet checks
# ---------------------------------------------------------------------------

def check_in_reference(ctx):
    reference = ctx.reference()
    if reference is None or not reference.available or not reference.values:
        return []     # the missing / empty reference sheet is reported elsewhere

    missing_rows, missing_values = [], []
    spelling_rows, spelling_values = [], []

    for row_number, value in ctx.cells():
        if is_blank(value):
            continue
        text = cell_text(value)
        if text in reference.exact:
            continue
        key = norm_key(text)
        if key in reference.normalized:
            spelling_rows.append(row_number)
            spelling_values.append(text)
        elif key in reference.aliases:
            continue          # an accepted short form, e.g. "Good" for
                               # "Good (No visible damage, no repairs completed)"
        else:
            missing_rows.append(row_number)
            missing_values.append(text)

    issues = []
    if missing_rows:
        issues.append(_issue(
            ctx, "not_in_reference",
            f"{ctx.column} values not found in {reference.label}",
            missing_rows, missing_values,
        ))
    if spelling_rows:
        issues.append(_issue(
            ctx, "case_mismatch",
            f"{ctx.column} values spelled differently from {reference.label}",
            spelling_rows, spelling_values,
        ))
    return issues


def check_code_matches_category(ctx):
    """The category code must be the code that belongs to the chosen category."""
    pairs = ctx.category_pairs
    category_column = ctx.columns.get("asset_category")
    names = ctx.references.get("category_name")
    codes = ctx.references.get("category_code")
    if not pairs or not category_column or names is None or codes is None:
        return []

    rows, values = [], []
    for row_number, row in zip(ctx.sheet.row_numbers, ctx.sheet.rows):
        if row_number in ctx.handled:
            continue
        category = row.get(category_column)
        code = row.get(ctx.column)
        if is_blank(category) or is_blank(code):
            continue
        category_key, code_key = norm_key(category), norm_key(code)
        if not names.contains(category_key) or not codes.contains(code_key):
            continue          # already reported by check_in_reference
        expected = pairs.get(category_key)
        if expected and expected != code_key:
            rows.append(row_number)
            values.append(f"{cell_text(code)} used for “{cell_text(category)[:45]}”")
    if not rows:
        return []
    return [_issue(ctx, "code_category_mismatch",
                   "Category Code does not match the Asset Category",
                   rows, values)]


def check_usd_matches_invoice(ctx):
    """When the invoice is already in USD, both value columns must agree."""
    currency_column = ctx.support.get("invoice_currency")
    invoice_column = ctx.columns.get("item_value_invoice")
    if not currency_column or not invoice_column:
        return []

    rows = []
    for row_number, row in zip(ctx.sheet.row_numbers, ctx.sheet.rows):
        if row_number in ctx.handled:
            continue
        if norm_key(row.get(currency_column)) != "usd":
            continue
        invoice = to_number(row.get(invoice_column))
        usd = to_number(row.get(ctx.column))
        if invoice is None or usd is None or invoice == 0:
            continue
        if abs(usd - invoice) / abs(invoice) > 0.01:
            rows.append(row_number)
    if not rows:
        return []
    return [_issue(ctx, "usd_mismatch",
                   "Item Value USD does not match the USD invoice value", rows)]


CHECKS = {
    "not_blank": check_not_blank,
    "no_placeholder": check_no_placeholder,
    "no_error_value": check_no_error_value,
    "not_zero": check_not_zero,
    "numeric": check_numeric,
    "unique": check_unique,
    "in_reference": check_in_reference,
    "code_matches_category": check_code_matches_category,
    "usd_matches_invoice": check_usd_matches_invoice,
}
