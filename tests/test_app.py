"""
End-to-end test of the Streamlit interface.

Runs app.py exactly as the browser would, with the file uploader replaced by a
real workbook, then clicks "Upload & Audit" and checks the results and the
feedback actually render without errors.

Run from the project folder:   python -m tests.test_app  [path-to-xlsx ...]
"""

from __future__ import annotations

import io
import pathlib
import sys

import streamlit as st
from streamlit.testing.v1 import AppTest

sys.path.insert(0, ".")

PASSED, FAILED = [], []


def expect(name, condition, extra=""):
    (PASSED if condition else FAILED).append(name)
    print(("  PASS  " if condition else "  FAIL  ") + name
          + (f"   {extra}" if extra and not condition else ""))


class FakeUpload(io.BytesIO):
    """Stands in for a Streamlit UploadedFile."""

    def __init__(self, path: pathlib.Path):
        data = path.read_bytes()
        super().__init__(data)
        self.name = path.name
        self._data = data

    def getvalue(self):
        return self._data


def run_app(path: pathlib.Path):
    upload = FakeUpload(path)
    original = st.file_uploader
    st.file_uploader = lambda *a, **k: upload      # noqa: ARG005
    try:
        app = AppTest.from_file("app.py", default_timeout=180)
        app.run()
        expect(f"{path.name}: page renders before the audit", not app.exception,
               str(app.exception))
        expect(f"{path.name}: audit button is present", len(app.button) >= 1)
        app.button[0].click().run()
        return app
    finally:
        st.file_uploader = original


def check(path: pathlib.Path):
    print(f"\n{path.name}")
    app = run_app(path)

    expect(f"{path.name}: audit runs without an exception",
           not app.exception, str(app.exception))
    expect(f"{path.name}: no error message shown",
           len(app.error) == 0, str([e.value for e in app.error]))

    text = " ".join(str(m.value) for m in app.markdown)
    expect(f"{path.name}: results heading rendered", "Issues found" in text
           or any("No data-quality issues" in str(s.value) for s in app.success))
    expect(f"{path.name}: feedback document rendered", "Feedback on" in text,
           text[:200])

    metrics = {m.label: m.value for m in app.metric}
    expect(f"{path.name}: rows-checked metric present",
           "Asset rows checked" in metrics, str(metrics))
    expect(f"{path.name}: findings metric present", "Findings" in metrics, str(metrics))

    codes = " ".join(str(c.value) for c in app.code)
    expect(f"{path.name}: plain-text copy box has the feedback",
           "Feedback on" in codes and "Please:" in codes)

    downloads = [b.label for b in app.get("download_button")]
    for wanted in ["Download .txt", "Download .docx", "Download .pdf",
                   "Detailed findings (.csv)"]:
        expect(f"{path.name}: {wanted} button present", wanted in downloads,
               str(downloads))

    print(f"    rows checked: {metrics.get('Asset rows checked')} | "
          f"sections: {metrics.get('Sections needing correction')} | "
          f"findings: {metrics.get('Findings')}")
    return app


if __name__ == "__main__":
    given = [pathlib.Path(a) for a in sys.argv[1:]]
    downloads = pathlib.Path.home() / "Downloads"
    default = [
        downloads / "Data Upload Template v3_Malawi_Zimbabwe.xlsx",
        downloads / "Data Upload Template V3_Zambia RM.xlsx",
    ]
    paths = given or [p for p in default if p.exists()]

    if not paths:
        print("No workbook found to test with.")
        sys.exit(1)

    for path in paths:
        check(path)

    print(f"\n{len(PASSED)} passed, {len(FAILED)} failed")
    if FAILED:
        for name in FAILED:
            print("  FAILED:", name)
        sys.exit(1)
