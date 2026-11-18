import io
from datetime import datetime

import streamlit as st

import build_tracker as bt

st.set_page_config(page_title="DXC Cost Tracker", layout="wide")
st.title("DXC Cost Tracker")
st.caption("Cost Approval + Client Name Disclosure (live DB)")

if st.button("Refresh from live DB"):
    st.cache_data.clear()


@st.cache_data(ttl=600, show_spinner="Querying live database...")
def load():
    return bt.build_df()


df = load()
st.metric("Total cases", len(df))
st.dataframe(df, use_container_width=True, hide_index=True)

tmp = bt.HERE / "_tmp_download.xlsx"
bt.write_excel(df, tmp)
st.download_button("Download Excel tracker", tmp.read_bytes(),
                   file_name=f"DXC_Cost_Tracker_{datetime.now():%Y%m%d_%H%M}.xlsx",
                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
