"""
Audit configuration.

EVERYTHING that describes *what* is checked lives in this file.
The engine (engine.py / checks.py) only knows *how* to run a check.

To extend the system later you normally only touch this file:

  * new placeholder word            -> add it to PLACEHOLDERS
  * new column name variation       -> add it to the "column" list of a rule
  * new sheet name variation        -> add it to the matching SHEET_SPECS entry
  * new field to validate           -> add a new dict to FIELD_RULES
  * new wording in the feedback      -> edit FEEDBACK_TEXT

A brand new *kind* of check (a check name that does not exist yet) is the only
change that also needs a small function added to checks.py.
"""

from __future__ import annotations

# ---------------------------------------------------------------------------
# 1. Placeholder / junk values
# ---------------------------------------------------------------------------
# Compared against the value after trimming and lower-casing.
# These are values that were typed *instead of* real data.

PLACEHOLDERS = {
    "n/a", "n\\a", "na", "n.a", "n.a.", "#n/a", "not applicable",
    "unknown", "unkown", "unknwon", "not known", "not available",
    "not yet", "notyet", "not yet available", "not yet known",
    "tba", "tbd", "tbc", "to be advised", "to be confirmed", "to be determined",
    "none", "null", "nil", "nan", "missing", "blank", "empty", "no data",
    "-", "--", "---", "_", "?", "??", "???", ".", "..", "...",
    "x", "xx", "xxx", "0", "00", "000",
}

# Excel error values are always invalid, wherever they appear.
ERROR_VALUES = {
    "#ref!", "#value!", "#name?", "#div/0!", "#null!", "#num!", "#n/a", "#n/a!",
    "#spill!", "#calc!", "#getting_data",
}

# Words that are removed from a file name when guessing the country / CO name.
FILENAME_NOISE_WORDS = {
    "data", "upload", "template", "asset", "assets", "register", "registry",
    "gpe", "file", "files", "final", "copy", "updated", "update", "new",
    "latest", "draft", "version", "form", "sheet", "workbook", "list",
    "submission", "submitted", "reviewed", "clean", "cleaned", "master",
}


# ---------------------------------------------------------------------------
# 2. Sheet detection
# ---------------------------------------------------------------------------
# Sheet names differ between country templates ("Asset | GPE Category",
# "Asset|GPECategory", "Asset GPE Category" ...).  Matching is done on the
# sheet name reduced to lower-case letters and digits only, so punctuation and
# spacing never matter.
#
#   exact       -> reduced name is exactly one of these  (highest confidence)
#   require_all -> every fragment must appear in the reduced name
#   any_of      -> at least one fragment must appear (empty list = no constraint)
#   forbid      -> none of these fragments may appear
#   required    -> the audit cannot run without this sheet

SHEET_SPECS = [
    {
        "key": "asset_info",
        "label": "Asset | GPE Information",
        "exact": ["assetgpeinformation", "assetgpeinfo", "assetinformation"],
        "require_all": ["asset"],
        "any_of": ["information", "info", "register", "registry"],
        "forbid": ["transfer", "categor", "dispose", "disposal", "donation",
                   "repair", "physical", "issue", "receipt", "theft", "lost"],
        "required": True,
    },
    {
        "key": "category_ref",
        "label": "Asset | GPE Category",
        "exact": ["assetgpecategory", "assetcategory", "gpecategory"],
        "require_all": ["categor"],
        "any_of": [],
        "forbid": ["information", "physical"],
        "required": True,
    },
    {
        "key": "supplier_ref",
        "label": "Supplier",
        "exact": ["supplier", "suppliers", "supplierreference", "supplierref"],
        "require_all": ["supplier"],
        "any_of": [],
        "forbid": [],
        "required": True,
    },
    {
        "key": "office_ref",
        "label": "Office",
        "exact": ["office", "offices", "officereference", "officeref",
                  "officetable"],
        "require_all": ["office"],
        "any_of": [],
        "forbid": ["asset", "transfer", "inventory"],
        "required": True,
    },
    {
        "key": "condition_ref",
        "label": "Condition | Disposal Reason",
        "exact": ["conditiondisposalreason", "conditiondisposal",
                  "conditionanddisposalreason"],
        "require_all": ["condition"],
        "any_of": [],
        "forbid": ["physical", "asset"],
        "required": True,
    },
]


# ---------------------------------------------------------------------------
# 3. Columns inside the reference sheets
# ---------------------------------------------------------------------------
# First matching name wins.  Matching ignores case, spaces and punctuation.

REFERENCE_COLUMNS = {
    "supplier_ref": {
        "name": ["SupplierName", "Supplier Name", "Supplier", "Name"],
        "status": ["Status"],
        "country": ["Country"],
    },
    "office_ref": {
        "name": ["OfficeName", "Office Name", "Office"],
        "type": ["OfficeType", "Office Type", "Type"],
        "country": ["Country"],
        "region": ["Region"],
    },
    "category_ref": {
        "name": ["Category Name", "CategoryName", "Asset | GPE Category",
                 "Category"],
        "code": ["Category code", "Category Code", "CategoryCode", "Code"],
        "type": ["Assets Type", "Asset Type", "AssetsType"],
    },
}

# The Condition | Disposal Reason sheet is often stored without a header row:
# one column of conditions, one of disposal reasons, one of office types.
# These labels are used when a header row *is* present, and as value hints
# when it is not.
CONDITION_SHEET_HINTS = {
    "condition": ["Condition", "Asset Condition", "Asset | GPE Condition",
                  "Conditions"],
    "disposal": ["Disposal Reason", "DisposalReason", "Disposal", "Reason"],
    "office_type": ["Office Type", "OfficeType", "Type"],
}


# ---------------------------------------------------------------------------
# 4. Field rules for the Asset | GPE Information sheet
# ---------------------------------------------------------------------------
# "checks" names are dispatched to functions in checks.py:
#
#   not_blank             value must not be empty
#   no_placeholder        value must not be N/A / TBA / UNKNOWN / ...
#   no_error_value        value must not be #REF! / #VALUE! / ...
#   not_zero              value must not be 0
#   unique                value must not repeat within the column
#   numeric               value must be a number, not a description
#   in_reference          value must exist in the named reference list
#   code_matches_category category code must belong to the chosen category
#   usd_matches_invoice   when the invoice currency is USD both values must agree
#
# "section" controls where the finding appears in the feedback document, and
# SECTION_ORDER below controls the order of the sections.

FIELD_RULES = [
    {
        "key": "asset_name",
        "section": "Asset Name",
        "sheet": "asset_info",
        "column": ["Asset Name", "AssetName", "Asset/GPE Name"],
        "checks": ["not_blank", "no_placeholder", "no_error_value"],
    },
    {
        "key": "asset_id",
        "section": "Asset ID",
        "sheet": "asset_info",
        "column": ["Asset ID", "AssetID", "Asset Id", "Asset No"],
        "checks": ["not_blank", "no_placeholder", "no_error_value",
                   "not_zero", "unique"],
    },
    {
        "key": "asset_category",
        "section": "Asset Category",
        "sheet": "asset_info",
        "column": ["Asset | GPE Category", "Asset GPE Category",
                   "Asset Category", "Category"],
        "checks": ["not_blank", "no_placeholder", "no_error_value",
                   "in_reference"],
        "reference": "category_name",
    },
    {
        "key": "asset_category_code",
        "section": "Asset Category Code",
        "sheet": "asset_info",
        "column": ["Asset | GPE Category Code", "Asset GPE Category Code",
                   "Asset Category Code", "Category Code"],
        "checks": ["not_blank", "no_placeholder", "no_error_value",
                   "in_reference", "code_matches_category"],
        "reference": "category_code",
    },
    {
        "key": "item_value_invoice",
        "section": "Item Value Invoice Currency",
        "sheet": "asset_info",
        "column": ["Item Value Invoice Currency", "Item Value Invoice",
                   "Value Invoice Currency"],
        "checks": ["not_blank", "numeric", "not_zero"],
    },
    {
        "key": "item_value_usd",
        "section": "Item Value USD",
        "sheet": "asset_info",
        "column": ["Item Value USD", "ItemValueUSD", "Value USD"],
        "checks": ["not_blank", "numeric", "not_zero", "usd_matches_invoice"],
    },
    {
        "key": "supplier",
        "section": "Supplier",
        "sheet": "asset_info",
        "column": ["Supplier", "Supplier Name", "SupplierName"],
        "checks": ["not_blank", "no_placeholder", "no_error_value",
                   "not_zero", "in_reference"],
        "reference": "supplier",
    },
    {
        "key": "office",
        "section": "Office",
        "sheet": "asset_info",
        "column": ["Office", "Office Name", "OfficeName"],
        "checks": ["not_blank", "no_placeholder", "no_error_value",
                   "not_zero", "in_reference"],
        "reference": "office",
    },
    {
        "key": "condition",
        "section": "Asset | GPE Condition",
        "sheet": "asset_info",
        "column": ["Asset | GPE Condition", "Asset GPE Condition",
                   "Asset Condition", "Condition"],
        "checks": ["not_blank", "no_placeholder", "no_error_value",
                   "in_reference"],
        "reference": "condition",
    },
]

# Helper columns the engine reads but does not audit on their own.
SUPPORT_COLUMNS = {
    "invoice_currency": ["Invoice Currency", "InvoiceCurrency", "Currency"],
}


# ---------------------------------------------------------------------------
# 5. Section order in the feedback document
# ---------------------------------------------------------------------------
# Reference sheets come first: they must be corrected before the columns that
# depend on them can be mapped.

SECTION_ORDER = [
    "Missing Sheets",
    "Supplier Reference Sheet",
    "Office Reference Sheet",
    "Asset | GPE Category Reference Sheet",
    "Condition | Disposal Reason Reference Sheet",
    "Asset Name",
    "Asset ID",
    "Asset Category",
    "Asset Category Code",
    "Item Value Invoice Currency",
    "Item Value USD",
    "Supplier",
    "Office",
    "Asset | GPE Condition",
]


# ---------------------------------------------------------------------------
# 6. Feedback wording
# ---------------------------------------------------------------------------
# For every (section, issue code) pair:
#   problem -> a short clause added to the section's opening sentence
#   actions -> the "Please:" bullet points
#
# {examples} is replaced by up to EXAMPLES_IN_TEXT quoted sample values.
# Duplicate bullets across codes are removed automatically.

EXAMPLES_IN_TEXT = 3
MAX_LISTED_VALUES = 12          # cap for "missing from the reference sheet" lists

FEEDBACK_TEXT = {
    # ---- reference sheets -------------------------------------------------
    ("Supplier Reference Sheet", "ref_blank"): {
        "problem": "blank supplier names",
        "actions": ["Fill in the missing supplier names, or delete the empty rows."],
    },
    ("Supplier Reference Sheet", "ref_duplicate"): {
        "problem": "the same supplier listed more than once",
        "actions": ["Remove the duplicate supplier entries so that each supplier appears only once."],
    },
    ("Supplier Reference Sheet", "ref_placeholder"): {
        "problem": "invalid entries such as {examples}",
        "actions": ["Replace invalid entries such as {examples} with the real supplier name, or delete those rows."],
    },
    ("Supplier Reference Sheet", "ref_missing"): {
        "problem": "suppliers that are used in the Asset | GPE Information sheet but are not listed here",
        "actions": ["Add the missing suppliers to the Supplier sheet so that every supplier used in the Asset | GPE Information sheet is listed."],
    },

    ("Office Reference Sheet", "ref_blank"): {
        "problem": "blank office names",
        "actions": ["Fill in the missing office names, or delete the empty rows."],
    },
    ("Office Reference Sheet", "ref_duplicate"): {
        "problem": "the same office listed more than once",
        "actions": ["Remove the duplicate office entries so that each office appears only once."],
    },
    ("Office Reference Sheet", "ref_placeholder"): {
        "problem": "invalid entries such as {examples}",
        "actions": ["Replace invalid entries such as {examples} with the real office name, or delete those rows."],
    },
    ("Office Reference Sheet", "ref_missing"): {
        "problem": "offices that are used in the Asset | GPE Information sheet but are not listed here",
        "actions": ["Add the missing offices to the Office sheet so that every office used in the Asset | GPE Information sheet is listed."],
    },
    ("Office Reference Sheet", "ref_type_blank"): {
        "problem": "a blank Office Type",
        "actions": ["Fill in the Office Type for every office."],
    },
    ("Office Reference Sheet", "ref_type_invalid"): {
        "problem": "an Office Type that is not a valid option ({examples})",
        "actions": ["Replace {examples} with the correct Office Type from the Condition | Disposal Reason sheet."],
    },

    ("Asset | GPE Category Reference Sheet", "ref_blank"): {
        "problem": "blank category names or codes",
        "actions": ["Fill in the missing category names and category codes."],
    },
    ("Asset | GPE Category Reference Sheet", "ref_duplicate"): {
        "problem": "the same category code used more than once",
        "actions": ["Make sure each category code is used for one category only."],
    },

    ("Condition | Disposal Reason Reference Sheet", "ref_blank"): {
        "problem": "blank entries",
        "actions": ["Fill in the missing condition and disposal reason options, or delete the empty rows."],
    },
    ("Condition | Disposal Reason Reference Sheet", "ref_duplicate"): {
        "problem": "repeated entries",
        "actions": ["Remove the repeated entries so that each option appears only once."],
    },
    ("Condition | Disposal Reason Reference Sheet", "ref_empty_sheet"): {
        "problem": "no valid condition options at all",
        "actions": ["Fill in the list of valid asset conditions before completing the Asset | GPE Condition column."],
    },

    # ---- Asset | GPE Information columns -----------------------------------
    ("Asset Name", "blank"): {
        "problem": "blank entries",
        "actions": ["Enter the correct asset name for every blank entry."],
    },
    ("Asset Name", "placeholder"): {
        "problem": "placeholder entries such as {examples}",
        "actions": ["Replace placeholder entries such as {examples} with the real name or description of the asset."],
    },
    ("Asset Name", "error_value"): {
        "problem": "Excel error values such as {examples}",
        "actions": ["Replace the error values with the real asset name."],
    },

    ("Asset ID", "blank"): {
        "problem": "blank entries",
        "actions": ["Enter the correct Asset ID for every blank entry."],
    },
    ("Asset ID", "placeholder"): {
        "problem": "placeholder entries such as {examples}",
        "actions": ["Replace placeholder entries such as {examples} with the correct Asset ID."],
    },
    ("Asset ID", "zero"): {
        "problem": "entries recorded as 0",
        "actions": ["Replace the 0 entries with the correct Asset ID."],
    },
    ("Asset ID", "error_value"): {
        "problem": "Excel error values such as {examples}",
        "actions": ["Replace the error values with the correct Asset ID."],
    },
    ("Asset ID", "duplicate"): {
        "problem": "the same ID used for more than one asset",
        "actions": ["Give every asset its own Asset ID. The same Asset ID must not be used twice."],
    },

    ("Asset Category", "blank"): {
        "problem": "blank entries",
        "actions": ["Select the correct category from the Asset | GPE Category sheet for every blank entry."],
    },
    ("Asset Category", "placeholder"): {
        "problem": "placeholder entries such as {examples}",
        "actions": ["Replace placeholder entries with the correct category from the Asset | GPE Category sheet."],
    },
    ("Asset Category", "error_value"): {
        "problem": "Excel error values such as {examples}",
        "actions": ["Replace the error values with a category from the Asset | GPE Category sheet."],
    },
    ("Asset Category", "not_in_reference"): {
        "problem": "categories that are not in the Asset | GPE Category sheet ({examples})",
        "actions": ["Replace {examples} with a category taken from the Asset | GPE Category sheet."],
    },
    ("Asset Category", "case_mismatch"): {
        "problem": "categories that are spelled slightly differently from the Asset | GPE Category sheet",
        "actions": ["Copy the category exactly as it is written in the Asset | GPE Category sheet."],
    },

    ("Asset Category Code", "blank"): {
        "problem": "blank entries",
        "actions": ["Enter the category code from the Asset | GPE Category sheet for every blank entry."],
    },
    ("Asset Category Code", "placeholder"): {
        "problem": "placeholder entries such as {examples}",
        "actions": ["Replace placeholder entries with the correct code from the Asset | GPE Category sheet."],
    },
    ("Asset Category Code", "error_value"): {
        "problem": "Excel error values such as {examples}",
        "actions": ["Replace the error values with the correct code from the Asset | GPE Category sheet."],
    },
    ("Asset Category Code", "not_in_reference"): {
        "problem": "codes that do not exist in the Asset | GPE Category sheet ({examples})",
        "actions": ["Replace {examples} with a code taken from the Asset | GPE Category sheet."],
    },
    ("Asset Category Code", "case_mismatch"): {
        "problem": "codes written in a different form from the Asset | GPE Category sheet",
        "actions": ["Write the code exactly as it appears in the Asset | GPE Category sheet."],
    },
    ("Asset Category Code", "code_category_mismatch"): {
        "problem": "codes that do not belong to the category chosen on the same row",
        "actions": ["Make sure the Asset Category Code is the code that belongs to the Asset Category on the same row, as shown in the Asset | GPE Category sheet."],
    },

    ("Item Value Invoice Currency", "blank"): {
        "problem": "blank entries",
        "actions": ["Enter the original purchase value. Where the purchase value is not known, use the agreed standard entry for unknown values."],
    },
    ("Item Value Invoice Currency", "not_numeric"): {
        "problem": "text entries such as {examples} instead of an amount",
        "actions": ["Replace text entries such as {examples} with the purchase amount in figures. This column must contain a value, not a description."],
    },
    ("Item Value Invoice Currency", "zero"): {
        "problem": "entries recorded as 0",
        "actions": ["Replace the 0 entries with the real purchase value, or with the agreed standard entry for unknown values."],
    },

    ("Item Value USD", "blank"): {
        "problem": "blank entries",
        "actions": ["Enter the USD value of the asset. Where the original value is not known, use the agreed standard entry for unknown values."],
    },
    ("Item Value USD", "not_numeric"): {
        "problem": "text entries such as {examples} instead of an amount",
        "actions": ["Replace text entries such as {examples} with the USD amount in figures."],
    },
    ("Item Value USD", "zero"): {
        "problem": "entries recorded as 0",
        "actions": ["Replace the 0 entries with the correct USD value."],
    },
    ("Item Value USD", "usd_mismatch"): {
        "problem": "rows where the invoice currency is already USD but the two value columns do not match",
        "actions": ["Where the Invoice Currency is USD, enter the same amount in Item Value USD as in Item Value Invoice Currency."],
    },

    ("Supplier", "blank"): {
        "problem": "blank entries",
        "actions": ["Select the correct supplier from the Supplier sheet for every blank entry."],
    },
    ("Supplier", "placeholder"): {
        "problem": "placeholder entries such as {examples}",
        "actions": ["Replace placeholder entries with the correct supplier name from the Supplier sheet."],
    },
    ("Supplier", "zero"): {
        "problem": "entries recorded as 0",
        "actions": ["Replace the 0 entries with the correct supplier name from the Supplier sheet."],
    },
    ("Supplier", "error_value"): {
        "problem": "Excel error values such as {examples}",
        "actions": ["Replace the error values with the correct supplier name from the Supplier sheet."],
    },
    ("Supplier", "not_in_reference"): {
        "problem": "supplier names that are not in the Supplier sheet ({examples})",
        "actions": ["Use only supplier names that are listed in the Supplier sheet. Where the supplier is genuinely new, add it to the Supplier sheet first and then select it here."],
    },
    ("Supplier", "case_mismatch"): {
        "problem": "supplier names spelled differently from the Supplier sheet",
        "actions": ["Write the supplier name exactly as it appears in the Supplier sheet, including spelling and spacing."],
    },

    ("Office", "blank"): {
        "problem": "blank entries",
        "actions": ["Select the correct office from the Office sheet for every blank entry."],
    },
    ("Office", "placeholder"): {
        "problem": "placeholder entries such as {examples}",
        "actions": ["Replace placeholder entries with the correct office name from the Office sheet."],
    },
    ("Office", "zero"): {
        "problem": "entries recorded as 0",
        "actions": ["Replace the 0 entries with the correct office name from the Office sheet."],
    },
    ("Office", "error_value"): {
        "problem": "Excel error values such as {examples}",
        "actions": ["Replace the error values with the correct office name from the Office sheet."],
    },
    ("Office", "not_in_reference"): {
        "problem": "office names that are not in the Office sheet ({examples})",
        "actions": ["Use only office names that are listed in the Office sheet. Where the office is genuinely new, add it to the Office sheet first and then select it here."],
    },
    ("Office", "case_mismatch"): {
        "problem": "office names spelled differently from the Office sheet",
        "actions": ["Write the office name exactly as it appears in the Office sheet."],
    },

    ("Asset | GPE Condition", "blank"): {
        "problem": "blank entries",
        "actions": ["Select the asset condition from the Condition | Disposal Reason sheet for every blank entry."],
    },
    ("Asset | GPE Condition", "placeholder"): {
        "problem": "entries such as {examples}",
        "actions": ["Replace entries such as {examples} with a valid condition from the Condition | Disposal Reason sheet."],
    },
    ("Asset | GPE Condition", "error_value"): {
        "problem": "Excel error values such as {examples}",
        "actions": ["Replace the error values with a valid condition from the Condition | Disposal Reason sheet."],
    },
    ("Asset | GPE Condition", "not_in_reference"): {
        "problem": "conditions that are not valid options ({examples})",
        "actions": ["Replace {examples} with one of the conditions listed in the Condition | Disposal Reason sheet."],
    },
    ("Asset | GPE Condition", "case_mismatch"): {
        "problem": "conditions written differently from the Condition | Disposal Reason sheet",
        "actions": ["Copy the condition exactly as it is written in the Condition | Disposal Reason sheet."],
    },
}

# Fallback wording when a (section, code) pair has no entry above.
FALLBACK_TEXT = {
    "problem": "invalid entries",
    "actions": ["Correct the invalid entries in this column."],
}

# Opening sentence of every section: "The <column> column in the <sheet> ..."
SECTION_INTRO = {
    "Supplier Reference Sheet": "The Supplier sheet has {problems}.",
    "Office Reference Sheet": "The Office sheet has {problems}.",
    "Asset | GPE Category Reference Sheet": "The Asset | GPE Category sheet has {problems}.",
    "Condition | Disposal Reason Reference Sheet": "The Condition | Disposal Reason sheet has {problems}.",
}
DEFAULT_SECTION_INTRO = "The {section} column in the Asset | GPE Information sheet has {problems}."
