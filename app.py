import io
from collections import Counter

import pandas as pd
import streamlit as st

from schema_compare import compare, load_reference, read_structure

st.set_page_config(page_title="FBA INV Report Checker", page_icon="📊", layout="wide")


@st.cache_data(show_spinner=False, max_entries=5)
def structure_from_bytes(data: bytes) -> list[dict]:
    return read_structure(io.BytesIO(data))


reference = load_reference()
ref_tabs = reference["tabs"]

st.title("📊 FBA INV Report — Tab & Column Checker")
st.caption(
    f"Reference (source of truth): **{reference['source_file']}** — {len(ref_tabs)} tabs. "
    "Dates in the file name and tab names (e.g. `11.9.26`, `11.1-11.7.26`) are ignored. "
    "Columns are read from row 1 of each tab and compared by name and position."
)

with st.sidebar:
    st.header("Reference structure")
    for t in ref_tabs:
        with st.expander(f"{t['key']}  ·  {len(t['columns'])} columns"):
            st.caption(f"Reference tab name: {t['name']}")
            if t["columns"]:
                st.dataframe(
                    pd.DataFrame({"Column name": [c if c.strip() else "(blank)" for c in t["columns"]]},
                                 index=range(1, len(t["columns"]) + 1)),
                    width="stretch",
                )
            else:
                st.write("No columns (tab has no header row).")

uploaded = st.file_uploader("Upload an FBA INV REPORTS workbook to check", type=["xlsx", "xlsm"])

if uploaded is None:
    st.info("Upload a file to compare it against the reference.")
    st.stop()

with st.spinner(f"Reading tabs and columns from {uploaded.name} ..."):
    try:
        up_tabs = structure_from_bytes(uploaded.getvalue())
    except Exception as exc:
        st.error(f"Could not read **{uploaded.name}** as an Excel workbook: {exc}")
        st.stop()

result = compare(ref_tabs, up_tabs)
diffs = result["differences"]
tabs = result["tabs"]

st.subheader(f"Results for: {uploaded.name}")

if not diffs:
    st.success(
        f"✅ CLEARED — No differences. All {len(ref_tabs)} tabs and their columns match the reference."
    )
else:
    tabs_with_diffs = [t for t in tabs if t.differences]
    st.error(
        f"❌ {len(diffs)} difference{'s' if len(diffs) != 1 else ''} found "
        f"in {len(tabs_with_diffs)} tab{'s' if len(tabs_with_diffs) != 1 else ''}."
    )

c1, c2, c3, c4 = st.columns(4)
c1.metric("Total differences", len(diffs))
c2.metric("Tabs with differences", sum(1 for t in tabs if t.differences))
c3.metric("Tabs matching", sum(1 for t in tabs if t.status == "Match"))
c4.metric("Tabs in uploaded file", len(up_tabs))

summary = pd.DataFrame([
    {
        "Tab (date ignored)": (t.differences[0].tab if t.differences else
                               next(r["key"] for r in ref_tabs if r["name"] == t.reference_name)),
        "Reference tab name": t.reference_name or "— (not in reference)",
        "Uploaded tab name": t.uploaded_name or "— (not in upload)",
        "Status": ("✅ Match" if t.status == "Match" else f"❌ {t.status}"),
        "Differences": len(t.differences),
    }
    for t in tabs
])


def _row_style(row):
    color = "" if row["Differences"] == 0 else "background-color: rgba(255, 75, 75, 0.18)"
    return [color] * len(row)


st.markdown("### Summary by tab")
st.dataframe(summary.style.apply(_row_style, axis=1), width="stretch", hide_index=True)

if diffs:
    st.markdown("### Differences by type")
    by_kind = Counter(d.kind for d in diffs)
    st.dataframe(
        pd.DataFrame(by_kind.most_common(), columns=["Difference type", "Count"]),
        hide_index=True,
    )

    st.markdown("### All differences")
    diff_df = pd.DataFrame([
        {
            "#": i,
            "Tab": d.tab,
            "Type": d.kind,
            "Location": d.location,
            "Expected (reference)": d.expected,
            "Found (uploaded)": d.found,
            "Details": d.detail,
        }
        for i, d in enumerate(diffs, start=1)
    ])
    st.dataframe(diff_df, width="stretch", hide_index=True)
    st.download_button(
        "Download differences (CSV)",
        diff_df.to_csv(index=False).encode("utf-8-sig"),
        file_name=f"differences_{uploaded.name.rsplit('.', 1)[0]}.csv",
        mime="text/csv",
    )

st.markdown("### Column detail per tab")


def _mismatch_style(row):
    color = "" if row["Same?"] == "✅" else "background-color: rgba(255, 75, 75, 0.18)"
    return [color] * len(row)


for t in tabs:
    title = t.uploaded_name or t.reference_name
    icon = "✅" if not t.differences else "❌"
    label = f"{icon} {title} — {len(t.differences)} difference{'s' if len(t.differences) != 1 else ''}"
    with st.expander(label, expanded=bool(t.differences)):
        if t.differences:
            for d in t.differences:
                st.markdown(f"- **{d.kind}** — {d.detail}")
        else:
            st.markdown("No differences in this tab.")
        if t.column_rows:
            st.dataframe(
                pd.DataFrame(t.column_rows).style.apply(_mismatch_style, axis=1),
                width="stretch",
                hide_index=True,
            )
        elif t.status != "Missing tab":
            st.caption("Neither the reference nor the uploaded tab has any column headers.")
