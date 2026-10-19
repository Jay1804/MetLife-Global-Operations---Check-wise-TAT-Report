from datetime import date

import pandas as pd
import streamlit as st

import metlife_report as mr

st.set_page_config(page_title="MetLife Global - Check wise TAT", layout="wide")
st.title("MetLife Global Operations - Check wise TAT Report")

with st.sidebar:
    st.header("Filters")
    client_code = st.text_input("Client code", mr.DEFAULT_CLIENT_CODE)
    start = st.date_input("Received from", date(date.today().year, date.today().month, 1))
    end = st.date_input("Received to", date.today())
    run = st.button("Generate report", type="primary", width="stretch")

if run:
    if start > end:
        st.error("'Received from' must be on or before 'Received to'.")
        st.stop()
    try:
        with st.spinner("Querying database..."):
            raw = mr.fetch_checks(client_code.strip(), start, end)
            config = mr.fetch_tat_config(client_code.strip())
            holidays = mr.fetch_holidays()
        report = mr.build_report(raw, config, holidays)
        with st.spinner("Building Excel pivots..."):
            xlsx, has_pivot = mr.to_excel(report)
        st.session_state["result"] = (report, xlsx, has_pivot, config, start, end)
    except Exception as exc:  # DB / network errors should surface in the UI, not crash it
        st.error(f"Could not generate the report: {exc}")
        st.stop()

if "result" not in st.session_state:
    st.info("Choose the client code and received-date range, then click **Generate report**.")
    st.stop()

df, xlsx, has_pivot, config, d_from, d_to = st.session_state["result"]
if df.empty:
    st.warning("No checks found for the selected client and date range.")
    st.stop()

by_month = mr.summary_by_month(df)
total = by_month[by_month["Month"] == "Grand Total"].iloc[0]
c1, c2, c3, c4 = st.columns(4)
c1.metric("Total checks", f"{int(total['Grand Total']):,}")
c2.metric("In TAT", f"{int(total['IT']):,}")
c3.metric("Out of TAT", f"{int(total['OT']):,}")
c4.metric("OT %", f"{total['OT %']:.1%}")

unmapped = df[df["Check_TAT_source"] != "TAT configuration"]["check_name2"].unique()
if len(unmapped):
    st.warning("No TAT configuration match (fell back to the query's TAT): " + ", ".join(unmapped))

pct = {c: "{:.1%}" for c in ["OT %", "IT %"]}
tab_sum, tab_data = st.tabs(["Summary", "Data"])
with tab_sum:
    st.subheader("TAT by month")
    st.dataframe(by_month.style.format(pct), hide_index=True, width="stretch")
    st.subheader("TAT by check name")
    st.dataframe(mr.summary_by_check(df).style.format(pct), hide_index=True, width="stretch")
with tab_data:
    st.dataframe(df, hide_index=True, width="stretch")

if not has_pivot:
    st.warning("Microsoft Excel was not available, so the Excel Summary tab is a static table instead of a PivotTable.")
st.download_button(
    "Download Excel report",
    data=xlsx,
    file_name=f"Metlife_Global_Operations_Check wise Data_{d_from:%b-%y}_to_{d_to:%b-%y}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    type="primary",
)
