"""
Excel Data Quality Auditor — web interface.

Two features, in tabs:
  Audit               Upload -> Audit -> Review -> Generate Feedback
  Preprocess for Migration   Upload -> Clean Supplier / Office / Asset | GPE
                              Information -> Download the cleaned workbook

Run with:   streamlit run app.py
An uploaded workbook is read in memory only. The audit tab never writes
anything back. The preprocessing tab builds a brand new workbook in memory
and offers it as a download — the file you uploaded is never modified.
"""

from __future__ import annotations

import html as _html
import io
import time

import streamlit as st

from auditor import (
    PREPROCESS_STEPS,
    PROGRESS_STEPS,
    audit,
    build_feedback,
    clean_workbook,
    render_markdown,
    render_text,
)
from auditor.exporters import (
    to_docx,
    to_findings_csv,
    to_pdf,
    to_preprocess_summary_csv,
    to_txt,
    to_unparseable_dates_csv,
)

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
  .pill-ok    { background:#e6f6ed; color:#1c6b3d; }
  .pill-warn  { background:#fff4e0; color:#8a5a00; }
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
  table.simple { width:100%; border-collapse:collapse; font-size:.9rem; }
  table.simple th { padding:.35rem .5rem; text-align:left; border-bottom:1px solid #e3e7ec; }
  table.simple td { padding:.35rem .5rem; border-bottom:1px solid #f0f2f5; }
</style>
"""
st.markdown(STYLE, unsafe_allow_html=True)

st.markdown('<div class="app-title">Excel Data Quality Auditor</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="app-subtitle">Audit a Data Upload Template workbook, or clean it up for migration</div>',
    unsafe_allow_html=True,
)

audit_tab, preprocess_tab = st.tabs(["🔍 Audit", "🧹 Preprocess for Migration"])


def _table(headers: list[str], rows: list[list[str]], right_align: set[int] = frozenset()) -> str:
    head = "".join(f'<th{" style=\"text-align:right\"" if i in right_align else ""}>'
                   f'{_html.escape(h)}</th>' for i, h in enumerate(headers))
    body = "".join(
        "<tr>" + "".join(
            f'<td{" style=\"text-align:right\"" if i in right_align else ""}>'
            f'{_html.escape(str(cell))}</td>' for i, cell in enumerate(row)
        ) + "</tr>"
        for row in rows
    )
    return f'<div style="overflow-x:auto"><table class="simple"><thead><tr>{head}</tr></thead><tbody>{body}</tbody></table></div>'


# =============================================================================
# Audit tab
# =============================================================================
with audit_tab:
    BLOCKING_SECTIONS = {
        "Missing Sheets",
        "Supplier Reference Sheet",
        "Office Reference Sheet",
        "Asset | GPE Category Reference Sheet",
        "Condition | Disposal Reason Reference Sheet",
    }

    st.markdown("Upload → Audit → Review → Generate Feedback")

    uploaded = st.file_uploader(
        "Drag and drop the Excel file here",
        type=["xlsx", "xlsm"],
        help="The file is read in memory only. It is never changed or saved.",
        key="audit_upload",
    )

    with st.sidebar:
        st.subheader("Audit options")
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
    else:
        file_bytes = uploaded.getvalue()

        if st.session_state.get("_file_signature") != (uploaded.name, len(file_bytes)):
            st.session_state["_file_signature"] = (uploaded.name, len(file_bytes))
            st.session_state.pop("result", None)

        st.markdown(f"**File:** `{uploaded.name}`  ·  {len(file_bytes) / 1024:,.0f} KB")

        run_audit = st.button("Upload & Audit", type="primary")

        # ---- Step 2 — Audit --------------------------------------------------
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
                    time.sleep(0.05)

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

        if result is not None:
            if country_override.strip() and country_override.strip() != result.country:
                result.country = country_override.strip()

            # ---- Step 3 — Results --------------------------------------------
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
                rows = [
                    [sheet["name"], role_labels.get(roles.get(sheet["name"], ""), "—"),
                    f'{sheet["rows"]:,}']
                    for sheet in result.sheet_overview
                ]
                st.markdown(_table(["Sheet", "Used as", "Rows of data"], rows, right_align={2}),
                           unsafe_allow_html=True)
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
                                        if issue.note_items:
                                            for item in issue.note_items:
                                                st.caption(f"• {item}")
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

            # ---- Step 4 — Feedback --------------------------------------------
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


# =============================================================================
# Preprocess tab
# =============================================================================
with preprocess_tab:
    st.markdown("Upload → Clean → Download the cleaned workbook")
    st.markdown(
        '<span class="muted">Cleans exactly three sheets: <b>Supplier</b> '
        "(removes blank/placeholder supplier names, and duplicate supplier names — "
        "keeping the most complete row), <b>Office</b> (removes duplicate office "
        "names — keeping the most complete row — and reformats every date column), "
        "and <b>Asset | GPE Information</b> (reformats every date column). "
        "Every other sheet in the workbook is left exactly as it was.</span>",
        unsafe_allow_html=True,
    )
    st.write("")

    pre_uploaded = st.file_uploader(
        "Drag and drop the Excel file here",
        type=["xlsx", "xlsm"],
        help="The uploaded file is never modified. A new, cleaned workbook is "
             "built in memory for you to download.",
        key="preprocess_upload",
    )

    with st.sidebar:
        st.divider()
        st.subheader("Preprocessing options")
        day_first_label = st.radio(
            "How are the dates written in this file?",
            options=["Day first — 26/12/2018 means 26 December 2018",
                    "Month first — 12/26/2018 means 26 December 2018"],
            index=0,
            help="Only matters for an ambiguous date like 01/02/2020. An "
                 "unambiguous date such as 2020-02-01 is read the same way "
                 "either way. Check a few real dates in your file if unsure.",
        )
        day_first = day_first_label.startswith("Day first")

    if pre_uploaded is None:
        st.info("Upload a Data Upload Template workbook to begin.")
    else:
        pre_bytes = pre_uploaded.getvalue()

        if st.session_state.get("_pre_file_signature") != (pre_uploaded.name, len(pre_bytes), day_first):
            st.session_state["_pre_file_signature"] = (pre_uploaded.name, len(pre_bytes), day_first)
            st.session_state.pop("preprocess_result", None)
            st.session_state.pop("preprocess_book_bytes", None)

        st.markdown(f"**File:** `{pre_uploaded.name}`  ·  {len(pre_bytes) / 1024:,.0f} KB")

        run_clean = st.button("Clean & Prepare Download", type="primary")

        if run_clean:
            seen: list[str] = []
            with st.status("Cleaning the workbook…", expanded=True) as status:
                placeholder = st.empty()

                def pre_progress(step_name: str):
                    seen.append(step_name)
                    lines = []
                    for step in PREPROCESS_STEPS:
                        if step in seen and step != seen[-1]:
                            lines.append(f"✅ {step}")
                        elif step == seen[-1]:
                            lines.append(f"⏳ **{step}**")
                        else:
                            lines.append(f"◻️ {step}")
                    placeholder.markdown("  \n".join(lines))
                    time.sleep(0.05)

                try:
                    book, pre_result = clean_workbook(
                        io.BytesIO(pre_bytes),
                        day_first=day_first,
                        filename=pre_uploaded.name,
                        progress=pre_progress,
                    )
                except Exception as error:                       # noqa: BLE001
                    status.update(label="The workbook could not be read", state="error")
                    st.error(f"The workbook could not be read: {error}")
                    st.stop()

                out_buffer = io.BytesIO()
                book.save(out_buffer)

                placeholder.markdown("  \n".join(f"✅ {step}" for step in PREPROCESS_STEPS))
                status.update(label="Cleaning complete", state="complete", expanded=False)

            st.session_state["preprocess_result"] = pre_result
            st.session_state["preprocess_book_bytes"] = out_buffer.getvalue()

        pre_result = st.session_state.get("preprocess_result")
        cleaned_bytes = st.session_state.get("preprocess_book_bytes")

        if pre_result is not None and cleaned_bytes is not None:
            st.divider()
            st.subheader("Cleaning complete")

            total_rows = sum(s.original_rows for s in pre_result.summaries)
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Sheets cleaned", len(pre_result.summaries))
            m2.metric("Rows processed", f"{total_rows:,}")
            m3.metric("Changes made", f"{pre_result.total_changes:,}")
            m4.metric("Dates left untouched", f"{pre_result.total_unparseable:,}")

            if pre_result.missing_sheets:
                st.warning("These sheets could not be found and were skipped: "
                          + ", ".join(pre_result.missing_sheets))

            sheet_labels = {
                "supplier_ref": "Supplier",
                "office_ref": "Office",
                "asset_info": "Asset | GPE Information",
            }

            for summary in pre_result.summaries:
                with st.container(border=True):
                    label = sheet_labels.get(summary.key, summary.sheet)
                    if not summary.available:
                        st.markdown(f'<div class="card-title">{label} '
                                   '<span class="pill pill-warn">Column not found</span></div>',
                                   unsafe_allow_html=True)
                        st.caption(f"Sheet “{summary.sheet}” was found, but the name column "
                                  "could not be identified, so this sheet was left as-is.")
                        continue

                    changed = (summary.removed_blank_or_placeholder or summary.removed_duplicates
                              or summary.dates_reformatted)
                    pill = (f'<span class="pill pill-ok">Cleaned</span>' if changed
                           else '<span class="pill pill-ok">Already clean</span>')
                    st.markdown(f'<div class="card-title">{label} &nbsp;{pill}</div>',
                               unsafe_allow_html=True)
                    st.caption(f"{summary.original_rows:,} rows → {summary.remaining_rows:,} rows")

                    lines = []
                    if summary.removed_blank_or_placeholder:
                        lines.append(f"• Removed {summary.removed_blank_or_placeholder:,} "
                                    f"blank / placeholder name entries")
                    if summary.removed_duplicates:
                        examples = ", ".join(summary.duplicate_examples[:6])
                        more = len(summary.duplicate_examples) - 6
                        lines.append(f"• Merged {summary.removed_duplicates:,} duplicate "
                                    f"name(s), keeping the most complete row"
                                    + (f" (e.g. {examples}"
                                       + (f", + {more} more" if more > 0 else "") + ")"
                                       if examples else ""))
                    if summary.date_columns:
                        lines.append(f"• Date columns: {', '.join(summary.date_columns)}")
                    if summary.dates_reformatted:
                        lines.append(f"• Reformatted {summary.dates_reformatted:,} dates")
                    if summary.dates_unparseable:
                        lines.append(f"• {len(summary.dates_unparseable):,} date cells could "
                                    f"not be confidently read and were left unchanged")

                    for line in lines:
                        st.markdown(f'<div class="issue-line">{line}</div>', unsafe_allow_html=True)

                    if summary.dates_unparseable:
                        with st.expander(f"Show cells that could not be read as dates "
                                        f"({len(summary.dates_unparseable):,})"):
                            by_column: dict[str, list] = {}
                            for u in summary.dates_unparseable:
                                by_column.setdefault(u.column, []).append(u)
                            for column, items in by_column.items():
                                shown = items[:15]
                                st.markdown(f"**{column}** — {len(items):,} cell(s)")
                                preview = ", ".join(
                                    f"row {u.row} (“{u.value}”)" for u in shown)
                                if len(items) > 15:
                                    preview += f" … (+{len(items) - 15} more)"
                                st.caption(preview)
                            st.caption("The full list, with every row number, is in the "
                                      "downloadable CSV below.")

            if any(s.key == "asset_info" and (s.dates_reformatted or s.dates_unparseable)
                  for s in pre_result.summaries):
                st.caption("Note: Asset | GPE Information also had formula cells "
                          "(e.g. a USD value calculated from the invoice amount). "
                          "Those are saved as plain values in the cleaned file, "
                          "the same numbers Excel was already showing.")

            st.write("")
            safe_name = (pre_uploaded.name.rsplit(".", 1)[0]
                        .replace(" ", "_").replace("&", "and"))
            d1, d2, d3 = st.columns(3)
            d1.download_button(
                "Download cleaned .xlsx", data=cleaned_bytes,
                file_name=f"Cleaned_{safe_name}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True, type="primary",
            )
            d2.download_button(
                "Summary (.csv)", data=to_preprocess_summary_csv(pre_result),
                file_name=f"Cleaning_Summary_{safe_name}.csv", mime="text/csv",
                use_container_width=True,
            )
            if pre_result.total_unparseable:
                d3.download_button(
                    "Unreadable dates (.csv)", data=to_unparseable_dates_csv(pre_result),
                    file_name=f"Unreadable_Dates_{safe_name}.csv", mime="text/csv",
                    use_container_width=True,
                )
