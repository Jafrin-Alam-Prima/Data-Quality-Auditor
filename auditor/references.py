"""
Builds the allowed-value lists from the reference sheets, and audits those
sheets while doing so (blanks, duplicates, placeholder entries).

The reference sheets are the source of truth.  If a reference sheet is
incomplete the audit says so first, because the columns that depend on it
cannot be mapped until it is fixed.
"""

from __future__ import annotations

from .config import CONDITION_SHEET_HINTS, REFERENCE_COLUMNS
from .models import Issue, ReferenceList
from .utils import cell_text, is_blank, is_placeholder, norm_key, reduce_name


def _add_short_form_aliases(reference: ReferenceList) -> None:
    """Let a short form entered in the Asset register match a reference value
    that carries a parenthetical explanation.

    Several reference sheets (Condition | Disposal Reason, and sometimes
    Asset | GPE Category) write the valid option as a long descriptive
    string, e.g. "Good (No visible damage, no repairs completed)". A value
    such as "Good" is the same option, just written without the explanation,
    and must not be reported as invalid.

    The alias is derived from the reference sheet's own text — the part
    before the first "(" — so it is never a hard-coded list of expected
    values, and it automatically follows whatever wording the reference
    sheet actually uses. A canonical (full-text) entry always takes priority:
    this only fills in a key that no real entry already occupies, so it can
    never hide or overwrite a genuine reference value.
    """
    for text in list(reference.values):
        if "(" not in text:
            continue
        short = text.split("(", 1)[0].strip()
        if not short:
            continue
        short_key = norm_key(short)
        reference.normalized.setdefault(short_key, text)


def _collect(sheet, label, ref_key, ref_label, section, issues):
    """Read one column of a reference sheet into a ReferenceList.

    Blank, duplicated and placeholder entries are recorded as issues against
    the reference sheet itself.
    """
    reference = ReferenceList(key=ref_key, label=ref_label,
                              sheet_name=sheet.name if sheet else "")
    if sheet is None or label is None:
        reference.available = False
        return reference

    blanks, placeholders, duplicates = [], [], []
    seen: dict[str, int] = {}

    for row_number, raw in sheet.column_values(label):
        text = cell_text(raw)
        if not text:
            blanks.append(row_number)
            continue
        if is_placeholder(raw):
            placeholders.append((row_number, text))
            continue
        key = norm_key(text)
        if key in seen:
            duplicates.append((row_number, text))
            continue
        seen[key] = row_number
        reference.values.append(text)
        reference.exact.add(text)
        reference.normalized[key] = text

    if blanks:
        issues.append(Issue(
            section=section, code="ref_blank",
            title=f"Blank entries in {ref_label}",
            sheet=sheet.name, column=label,
            count=len(blanks), rows=blanks,
        ))
    if placeholders:
        issues.append(Issue(
            section=section, code="ref_placeholder",
            title=f"Invalid entries in {ref_label}",
            sheet=sheet.name, column=label,
            count=len(placeholders),
            values=sorted({v for _, v in placeholders}),
            rows=[r for r, _ in placeholders],
        ))
    if duplicates:
        issues.append(Issue(
            section=section, code="ref_duplicate",
            title=f"Duplicate entries in {ref_label}",
            sheet=sheet.name, column=label,
            count=len(duplicates),
            values=sorted({v for _, v in duplicates}),
            rows=[r for r, _ in duplicates],
        ))
    _add_short_form_aliases(reference)
    return reference


def _column(sheet, ref_key, field_name):
    if sheet is None:
        return None
    return sheet.find_column(REFERENCE_COLUMNS[ref_key][field_name])


# ---------------------------------------------------------------------------
# Supplier
# ---------------------------------------------------------------------------

def build_supplier_reference(workbook, issues):
    sheet = workbook.get("supplier_ref")
    label = _column(sheet, "supplier_ref", "name")
    return _collect(sheet, label, "supplier", "the Supplier sheet",
                    "Supplier Reference Sheet", issues)


# ---------------------------------------------------------------------------
# Office
# ---------------------------------------------------------------------------

def build_office_reference(workbook, issues, office_types=None):
    sheet = workbook.get("office_ref")
    label = _column(sheet, "office_ref", "name")
    reference = _collect(sheet, label, "office", "the Office sheet",
                         "Office Reference Sheet", issues)

    # Office Type must be filled in, and must be one of the valid options when
    # the workbook provides a list of them.
    type_label = _column(sheet, "office_ref", "type")
    if sheet is not None and type_label and label:
        blanks, invalid = [], []
        for row_number, row in zip(sheet.row_numbers, sheet.rows):
            if is_blank(row.get(label)):
                continue                       # already reported as a blank office
            raw = row.get(type_label)
            if is_blank(raw) or is_placeholder(raw):
                blanks.append(row_number)
                continue
            if office_types and office_types.values:
                if norm_key(raw) not in office_types.normalized:
                    invalid.append((row_number, cell_text(raw)))
        if blanks:
            issues.append(Issue(
                section="Office Reference Sheet", code="ref_type_blank",
                title="Offices without an Office Type",
                sheet=sheet.name, column=type_label,
                count=len(blanks), rows=blanks,
            ))
        if invalid:
            issues.append(Issue(
                section="Office Reference Sheet", code="ref_type_invalid",
                title="Invalid Office Type",
                sheet=sheet.name, column=type_label,
                count=len(invalid),
                values=sorted({v for _, v in invalid}),
                rows=[r for r, _ in invalid],
                note="Valid options: " + ", ".join(office_types.values),
            ))
    return reference


def office_countries(workbook):
    sheet = workbook.get("office_ref")
    label = _column(sheet, "office_ref", "country")
    if sheet is None or not label:
        return []
    return [cell_text(v) for _, v in sheet.column_values(label) if cell_text(v)]


# ---------------------------------------------------------------------------
# Asset | GPE Category
# ---------------------------------------------------------------------------

def build_category_reference(workbook, issues):
    """Return (names, codes, name_key -> code_key map)."""
    sheet = workbook.get("category_ref")
    name_label = _column(sheet, "category_ref", "name")
    code_label = _column(sheet, "category_ref", "code")
    section = "Asset | GPE Category Reference Sheet"

    names = ReferenceList(key="category_name", label="the Asset | GPE Category sheet",
                          sheet_name=sheet.name if sheet else "")
    codes = ReferenceList(key="category_code", label="the Asset | GPE Category sheet",
                          sheet_name=sheet.name if sheet else "")
    pairs: dict[str, str] = {}

    if sheet is None or not name_label or not code_label:
        names.available = codes.available = False
        return names, codes, pairs

    blanks, duplicate_codes = [], []
    seen_codes: dict[str, str] = {}

    for row_number, row in zip(sheet.row_numbers, sheet.rows):
        name = cell_text(row.get(name_label))
        code = cell_text(row.get(code_label))
        if not name and not code:
            continue
        if not name or not code:
            blanks.append(row_number)
            continue

        name_k, code_k = norm_key(name), norm_key(code)
        if name_k not in names.normalized:
            names.values.append(name)
            names.exact.add(name)
            names.normalized[name_k] = name
        if code_k in seen_codes and seen_codes[code_k] != name_k:
            duplicate_codes.append(code)
        else:
            seen_codes[code_k] = name_k
        if code_k not in codes.normalized:
            codes.values.append(code)
            codes.exact.add(code)
            codes.normalized[code_k] = code
        pairs.setdefault(name_k, code_k)
        if "(" in name:                        # short-form alias, e.g. "Vehicles"
            short_key = norm_key(name.split("(", 1)[0])
            if short_key:
                pairs.setdefault(short_key, code_k)

    if blanks:
        issues.append(Issue(
            section=section, code="ref_blank",
            title="Categories missing a name or a code",
            sheet=sheet.name, count=len(blanks), rows=blanks,
        ))
    if duplicate_codes:
        issues.append(Issue(
            section=section, code="ref_duplicate",
            title="Category codes used for more than one category",
            sheet=sheet.name, column=code_label,
            count=len(duplicate_codes), values=sorted(set(duplicate_codes)),
        ))
    _add_short_form_aliases(names)
    _add_short_form_aliases(codes)
    return names, codes, pairs


# ---------------------------------------------------------------------------
# Condition | Disposal Reason
# ---------------------------------------------------------------------------

def _classify_condition_columns(grid):
    """Work out which column holds what in the Condition sheet.

    The sheet is usually stored without a header row: one column of asset
    conditions, one of disposal reasons and one of office types.  When a header
    row is present its labels are used instead.
    """
    if not grid:
        return {}, 0

    hint_map = {}
    for role, labels in CONDITION_SHEET_HINTS.items():
        for label in labels:
            hint_map[reduce_name(label)] = role

    # Does the first row look like a header?
    header_roles = {}
    for index, cell in enumerate(grid[0]):
        role = hint_map.get(reduce_name(cell))
        if role and role not in header_roles:
            header_roles[role] = index
    if header_roles and len(grid) > 1:
        return header_roles, 1

    # No header: classify by content.
    filled = {}
    for index in range(max(len(r) for r in grid)):
        values = [cell_text(r[index]) for r in grid
                  if index < len(r) and cell_text(r[index])]
        if values:
            filled[index] = values

    roles = {}
    remaining = []
    for index, values in sorted(filled.items()):
        if values and all("office" in v.casefold() for v in values):
            roles.setdefault("office_type", index)
        else:
            remaining.append(index)
    if remaining:
        roles["condition"] = remaining[0]
    if len(remaining) > 1:
        roles["disposal"] = remaining[1]
    return roles, 0


def build_condition_reference(workbook, issues):
    """Return (conditions, disposal_reasons, office_types)."""
    sheet = workbook.get("condition_ref")
    section = "Condition | Disposal Reason Reference Sheet"
    label = "the Condition | Disposal Reason sheet"

    empty = lambda key: ReferenceList(key=key, label=label,
                                      sheet_name=sheet.name if sheet else "",
                                      available=False)
    if sheet is None or not sheet.grid:
        return empty("condition"), empty("disposal"), empty("office_type")

    roles, first_row = _classify_condition_columns(sheet.grid)
    results = {}

    for role, key in (("condition", "condition"),
                      ("disposal", "disposal"),
                      ("office_type", "office_type")):
        reference = ReferenceList(key=key, label=label, sheet_name=sheet.name)
        index = roles.get(role)
        if index is None:
            reference.available = False
            results[role] = reference
            continue

        duplicates = []
        for offset, row in enumerate(sheet.grid[first_row:], start=first_row + 1):
            raw = row[index] if index < len(row) else None
            text = cell_text(raw)
            if not text:
                continue
            normalized = norm_key(text)
            if normalized in reference.normalized:
                duplicates.append((offset, text))
                continue
            reference.values.append(text)
            reference.exact.add(text)
            reference.normalized[normalized] = text

        if duplicates and role == "condition":
            issues.append(Issue(
                section=section, code="ref_duplicate",
                title="Repeated condition options",
                sheet=sheet.name, count=len(duplicates),
                values=sorted({v for _, v in duplicates}),
                rows=[r for r, _ in duplicates],
            ))
        _add_short_form_aliases(reference)
        results[role] = reference

    if not results["condition"].values:
        issues.append(Issue(
            section=section, code="ref_empty_sheet",
            title="No valid condition options found",
            sheet=sheet.name, count=1,
        ))
    return results["condition"], results["disposal"], results["office_type"]
