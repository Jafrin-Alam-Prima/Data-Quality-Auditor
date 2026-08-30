"""
Excel Data Quality Auditor — web interface.

Upload -> Audit -> Review -> Generate Feedback

Run with:   streamlit run app.py
The uploaded workbook is read in memory only and is never modified or saved.
"""

from __future__ import annotations

import io
import time

import streamlit as st

from auditor import PROGRESS_STEPS, audit, build_feedback, render_markdown, render_text
from auditor.exporters import to_docx, to_findings_csv, to_pdf, to_txt

st.set_page_config(
    page_title="Excel Data Quality Auditor",
    page_icon="📋",
    layout="wide",
)

STYLE = """
<style>
  .block-container { padding-top: 2.2rem; max-width: 1180px; }
  .app-title { font-size: 2rem; font-weight: 700; margin-bottom: .1rem; }
  .app-subtitle { color: #5c6672; font-size: 1rem; margin-bottom: 1.4rem; }
  .pill {
      display:inline-block; padding:.12rem .55rem; border-radius:999px;
      font-size:.72rem; font-weight:600; letter-spacing:.02em; vertical-align:middle;
  }
  .pill-fix   { background:#fdecea; color:#a4271b; }
  .pill-ref   { background:#e8f1fd; color:#1b4f8a; }
  .card-title { font-size:1.02rem; font-weight:650; margin:0 0 .35rem 0; }
  .issue-line { font-size:.9rem; color:#37414c; margin:.16rem 0; }
  .issue-count { color:#7b8794; font-size:.82rem; }
  .doc {
      background:#ffffff; border:1px solid #e3e7ec; border-radius:10px;
      padding:1.5rem 1.8rem; line-height:1.6;
  }
  .doc h3 { margin-top:0; }
  .muted { color:#7b8794; font-size:.85rem; }
  div[data-testid="stMetricValue"] { font-size:1.5rem; }
</style>
"""
st.markdown(STYLE, unsafe_allow_html=True)

# Sections that block the rest of the correction work.
BLOCKING_SECTIONS = {
    "Missing Sheets",
    "Supplier Reference Sheet",
    "Office Reference Sheet",
    "Asset | GPE Category Reference Sheet",
    "Condition | Disposal Reason Reference Sheet",
}


# ---------------------------------------------------------------------------
# Header
# ---------------------------------------------------------------------------
st.markdown('<div class="app-title">Excel Data Quality Auditor</div>',
            unsafe_allow_html=True)
st.markdown('<div class="app-subtitle">Upload → Audit → Review → Generate Feedback</div>',
            unsafe_allow_html=True)


# ---------------------------------------------------------------------------
# Step 1 — Upload
# ---------------------------------------------------------------------------
uploaded = st.file_uploader(
    "Drag and drop the Excel file here",
    type=["xlsx", "xlsm"],
    help="The file is read in memory only. It is never changed or saved.",
)

with st.sidebar:
    st.subheader("Options")
    country_override = st.text_input(
        "Country / CO name",
        value="",
        placeholder="Detected from the file name",
        help="Used in the feedback title. Leave empty to detect it automatically.",
    )
    include_full_lists = st.checkbox(
        "List every missing supplier / office in the feedback",
        value=False,
        help="Off by default so the feedback stays short. The full list is "
             "always available in the detailed findings download.",
    )
    st.divider()
    st.caption("The audit rules live in `auditor/config.py`. "
               "New columns, sheet names, placeholder words and wording can be "
               "changed there without touching the rest of the application.")

if uploaded is None:
    st.info("Upload a Data Upload Template workbook to begin.")
    st.stop()

file_bytes = uploaded.getvalue()

if st.session_state.get("_file_signature") != (uploaded.name, len(file_bytes)):
    st.session_state["_file_signature"] = (uploaded.name, len(file_bytes))
    st.session_state.pop("result", None)

st.markdown(f"**File:** `{uploaded.name}`  ·  {len(file_bytes) / 1024:,.0f} KB")

run_audit = st.button("Upload & Audit", type="primary")


# ---------------------------------------------------------------------------
# Step 2 — Audit
# ---------------------------------------------------------------------------
if run_audit:
    seen: list[str] = []
    with st.status("Auditing the workbook…", expanded=True) as status:
        placeholder = st.empty()

        def progress(step_name: str):
            seen.append(step_name)
            lines = []
            for step in PROGRESS_STEPS:
                if step in seen and step != seen[-1]:
                    lines.append(f"✅ {step}")
                elif step == seen[-1]:
                    lines.append(f"⏳ **{step}**")
                else:
                    lines.append(f"◻️ {step}")
            placeholder.markdown("  \n".join(lines))
            time.sleep(0.05)          # let the interface repaint

        try:
            result = audit(
                io.BytesIO(file_bytes),
                filename=uploaded.name,
                country=country_override.strip(),
                progress=progress,
            )
        except Exception as error:                      # noqa: BLE001
            status.update(label="The workbook could not be read", state="error")
            st.error(f"The workbook could not be read: {error}")
            st.stop()

        placeholder.markdown("  \n".join(f"✅ {step}" for step in PROGRESS_STEPS))
        status.update(label="Audit complete", state="complete", expanded=False)

    st.session_state["result"] = result

result = st.session_state.get("result")
if result is None:
    st.stop()

if country_override.strip() and country_override.strip() != result.country:
    result.country = country_override.strip()


# ---------------------------------------------------------------------------
# Step 3 — Results
# ---------------------------------------------------------------------------
grouped = result.by_section()
finding_count = len(result.issues)

st.divider()
st.subheader("Audit complete")

left, middle, right = st.columns(3)
left.metric("Asset rows checked", f"{result.asset_row_count:,}")
middle.metric("Sections needing correction", len(grouped))
right.metric("Findings", finding_count)

with st.expander("Sheets found in this workbook", expanded=False):
    roles = {name: key for key, name in result.resolved_sheets.items()}
    role_labels = {
        "asset_info": "Asset | GPE Information",
        "category_ref": "Category reference",
        "supplier_ref": "Supplier reference",
        "office_ref": "Office reference",
        "condition_ref": "Condition | Disposal Reason reference",
    }
    import html as _html

    table_rows = "".join(
        '<tr style="border-bottom:1px solid #f0f2f5">'
        f'<td style="padding:.35rem .5rem">{_html.escape(sheet["name"])}</td>'
        f'<td style="padding:.35rem .5rem">'
        f'{_html.escape(role_labels.get(roles.get(sheet["name"], ""), "—"))}</td>'
        f'<td style="padding:.35rem .5rem;text-align:right">{sheet["rows"]:,}</td></tr>'
        for sheet in result.sheet_overview
    )
    st.markdown(
        '<div style="overflow-x:auto"><table style="width:100%;border-collapse:collapse;'
        'font-size:.9rem">'
        '<thead><tr style="text-align:left;border-bottom:1px solid #e3e7ec">'
        '<th style="padding:.35rem .5rem">Sheet</th>'
        '<th style="padding:.35rem .5rem">Used as</th>'
        '<th style="padding:.35rem .5rem;text-align:right">Rows of data</th>'
        f'</tr></thead><tbody>{table_rows}</tbody></table></div>',
        unsafe_allow_html=True,
    )
    if result.missing_sheets:
        st.warning("Reference sheets not found: " + ", ".join(result.missing_sheets))

if not result.has_issues:
    st.success("No data-quality issues were found in this workbook.")
else:
    st.markdown("#### Issues found")
    st.markdown(
        '<span class="muted">Only the sheets and columns with an actual problem '
        "are listed. The reference sheets are shown first because they must be "
        "corrected before the columns that depend on them can be mapped.</span>",
        unsafe_allow_html=True,
    )
    st.write("")

    section_names = list(grouped)
    for start in range(0, len(section_names), 2):
        for column, section in zip(st.columns(2), section_names[start:start + 2]):
            issues = grouped[section]
            with column, st.container(border=True):
                pill = ('<span class="pill pill-fix">Fix first</span>'
                        if section in BLOCKING_SECTIONS
                        else '<span class="pill pill-ref">Asset register</span>')
                st.markdown(
                    f'<div class="card-title">{section} &nbsp;{pill}</div>',
                    unsafe_allow_html=True,
                )
                for issue in issues:
                    st.markdown(
                        f'<div class="issue-line">• {issue.title} '
                        f'<span class="issue-count">({issue.count})</span></div>',
                        unsafe_allow_html=True,
                    )
                details = [i for i in issues if i.values or i.rows]
                if details:
                    with st.expander("Show the affected values and rows"):
                        for issue in details:
                            st.markdown(f"**{issue.title}**")
                            if issue.note:
                                st.caption(issue.note)
                            if issue.values:
                                shown = issue.values[:40]
                                st.write(", ".join(shown)
                                         + (f" … (+{len(issue.values) - 40} more)"
                                            if len(issue.values) > 40 else ""))
                            if issue.rows:
                                rows_sorted = sorted(set(issue.rows))
                                preview = ", ".join(str(r) for r in rows_sorted[:30])
                                if len(rows_sorted) > 30:
                                    preview += f" … (+{len(rows_sorted) - 30} more)"
                                st.caption(f"Excel rows: {preview}")


# ---------------------------------------------------------------------------
# Step 4 — Feedback
# ---------------------------------------------------------------------------
st.divider()
st.subheader("Feedback for the data-entry team")

document = build_feedback(result, include_full_lists=include_full_lists)
plain_text = render_text(document)
safe_country = (result.country or "workbook").replace(" ", "_").replace("&", "and")

view, copy_view = st.tabs(["Document", "Plain text (copy)"])
with view:
    st.markdown(f'<div class="doc">{render_markdown(document)}</div>',
                unsafe_allow_html=True)
with copy_view:
    st.caption("Use the copy icon in the top-right corner of the box below.")
    st.code(plain_text, language=None)

st.write("")
one, two, three, four = st.columns(4)
one.download_button(
    "Download .txt", data=to_txt(document),
    file_name=f"Feedback_{safe_country}.txt", mime="text/plain",
    use_container_width=True,
)
two.download_button(
    "Download .docx", data=to_docx(document),
    file_name=f"Feedback_{safe_country}.docx",
    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    use_container_width=True,
)
try:
    pdf_bytes = to_pdf(document)
    three.download_button(
        "Download .pdf", data=pdf_bytes,
        file_name=f"Feedback_{safe_country}.pdf", mime="application/pdf",
        use_container_width=True,
    )
except Exception:                                       # noqa: BLE001
    three.button("PDF unavailable", disabled=True, use_container_width=True,
                 help="Install reportlab to enable PDF downloads.")
four.download_button(
    "Detailed findings (.csv)", data=to_findings_csv(result),
    file_name=f"Findings_{safe_country}.csv", mime="text/csv",
    use_container_width=True,
    help="Every finding with its Excel row numbers — for the reviewer, "
         "not for the data-entry team.",
)
