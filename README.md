# Excel Data Quality Auditor

Upload a Data Upload Template workbook → the system audits it against the
asset-register rules → you get a short, plain-English feedback document you can
copy or download and send to the data-entry team.

The uploaded file is read in memory only. **It is never modified or saved.**

---

## Running it

**Windows** — double-click `run.bat`.

**macOS / Linux** — `./run.sh`

Either script creates a virtual environment the first time, installs the
packages, and opens the app in your browser at <http://localhost:8501>.

To run it by hand:

```bash
pip install -r requirements.txt
```

```bash
streamlit run app.py
```

---

## How it works

1. **Upload** an `.xlsx` file. The app lists every sheet it found, how many rows
   of data each has, and which sheet it is using as which reference.
2. **Audit** runs automatically when you press *Upload & Audit*.
3. **Review** the issues. Only sheets and columns with an actual problem are
   shown. Each card can be expanded to see the affected values and Excel rows.
4. **Feedback** is generated below the results. Copy it, or download it as
   `.txt`, `.docx` or `.pdf`. A `.csv` of detailed findings is also available
   for your own review — that one is not meant for the data-entry team.

The country / CO name is taken from the file name
(`Data Upload Template V3_Zambia RM.xlsx` → *Feedback on Zambia RM Data*).
You can override it in the sidebar.

---

## What it checks

Reference sheets are checked first and reported first, because the columns that
depend on them cannot be mapped until they are correct.

**Supplier sheet** — blank names, duplicates, placeholder entries, and suppliers
used in the asset register but not listed here.

**Office sheet** — blank names, duplicates, placeholder entries, offices used in
the asset register but not listed here, blank Office Type, and Office Type
values that are not one of the valid options.

**Asset | GPE Category sheet** — categories missing a name or a code, and codes
used for more than one category.

**Condition | Disposal Reason sheet** — repeated options, and an empty option
list. This sheet is the source of truth for valid conditions.

**Asset | GPE Information sheet**

| Column | Checked for |
|---|---|
| Asset Name | blank, placeholder words, Excel errors |
| Asset ID | blank, placeholder words, `0`, `#REF!`, duplicates |
| Asset \| GPE Category | blank, placeholder, not in the Category sheet, spelling differences |
| Asset \| GPE Category Code | blank, placeholder, not in the Category sheet, **code that does not belong to the category on the same row** |
| Item Value Invoice Currency | blank, `0`, descriptions typed instead of an amount |
| Item Value USD | blank, `0`, descriptions, and **a USD invoice whose two value columns disagree** |
| Supplier | blank, `0`, placeholder, not in the Supplier sheet, spelling differences |
| Office | blank, `0`, placeholder, not in the Office sheet, spelling differences |
| Asset \| GPE Condition | blank, placeholder, not a valid option in the Condition sheet |

Duplicate Asset **Names** are allowed and are not reported.

A cell is only reported once: a blank cell is not also reported as "not in the
reference sheet".

**Spelling differences** are separated from **missing values**. `HARARE` where
the Office sheet says `Harare` is a spelling difference; `Balaka`, which is not
in the Office sheet at all, is a missing value that must be added to the Office
sheet first.

The system never invents a correction. Where a purchase value is unknown the
feedback says to use *the agreed standard entry for unknown values* rather than
suggesting an amount.

---

## How values are matched against a reference sheet

Every cross-sheet comparison (Supplier, Office, Asset Category, Asset Category
Code, Asset | GPE Condition) goes through one reusable function,
`normalize_for_matching()` in `auditor/utils.py`. Before two values are
compared, both are: converted to a string, stripped of invisible Unicode
characters (non-breaking spaces, zero-width spaces, a BOM, soft hyphens, bidi
marks — the kind of character that looks identical in Excel but breaks an
exact match), trimmed, collapsed to single spaces, and case-folded. The original
value from the reference sheet is always the one shown as correct — this
normalized form is only ever used as a comparison key, never displayed.

A value is reported as **missing** only when it does not match a reference
value even after normalization. When it *does* match, but the raw text
differs (capitalization, spacing), it is reported separately as a
**formatting / spelling difference** — never both.

**Short forms of a reference value are also recognised.** Some reference
sheets write the valid option with an explanation in parentheses, e.g.
`Good (No visible damage, no repairs completed)` or
`Appliance Expert (Pvt) Ltd`. A register that uses the short form (`Good`,
`Appliance Expert`) is using the same, valid option and is not flagged. This
alias is derived from the reference sheet's own text — the part before the
first `(` — every time the audit runs, so it is never a hard-coded list and it
always follows whatever wording the reference sheet actually uses.

---

## Sheet detection

Sheet names differ between country templates, so matching ignores case,
spacing and punctuation. `Asset | GPE Category`, `Asset|GPECategory` and
`Asset GPE Category` all resolve to the same reference. If a required reference
sheet cannot be found, the feedback says so before anything else.

The `Condition | Disposal Reason` sheet is read whether or not it has a header
row — in the standard template it does not.

---

## Project layout

```
app.py                  Streamlit interface (upload, results, feedback, downloads)
auditor/
    config.py           ALL the rules, sheet names, column names and wording
    workbook.py         Excel processing: open, detect sheets, detect header rows
    references.py       builds the allowed-value lists from the reference sheets
    checks.py           the individual checks
    engine.py           runs the rules and returns the findings
    feedback.py         turns findings into plain-English feedback
    exporters.py        .txt / .docx / .pdf / .csv
    models.py           Issue, ReferenceList, AuditResult
tests/
    test_engine.py      rules and feedback, on workbooks built in memory
    test_app.py         the whole interface, on real workbooks
```

The interface contains no rules. The engine contains no wording. Everything you
are likely to change lives in `auditor/config.py`.

---

## Changing the rules

Open `auditor/config.py`.

| To do this | Edit this |
|---|---|
| Add a placeholder word such as `PENDING` | `PLACEHOLDERS` |
| Accept another spelling of a sheet name | the matching entry in `SHEET_SPECS` |
| Accept another spelling of a column name | the `column` list of that rule in `FIELD_RULES` |
| Validate a new column | add a dict to `FIELD_RULES` |
| Change the order of the feedback sections | `SECTION_ORDER` |
| Reword any sentence or bullet | `FEEDBACK_TEXT` |
| Change how many missing values are listed | `MAX_LISTED_VALUES` |

A rule looks like this:

```python
{
    "key": "supplier",
    "section": "Supplier",
    "sheet": "asset_info",
    "column": ["Supplier", "Supplier Name", "SupplierName"],
    "checks": ["not_blank", "no_placeholder", "no_error_value",
               "not_zero", "in_reference"],
    "reference": "supplier",
}
```

Checks run in the order listed. A brand-new *kind* of check is the only change
that also needs a small function added to `auditor/checks.py` and registered in
its `CHECKS` dictionary.

---

## Tests

```bash
python -m tests.test_engine
```

```bash
python -m tests.test_app
```

`test_engine.py` builds workbooks in memory containing each kind of problem and
checks that exactly those problems are reported — including that a clean
workbook produces no findings at all. `test_app.py` runs the real interface
against the workbooks in your Downloads folder, or against paths you pass on the
command line.
