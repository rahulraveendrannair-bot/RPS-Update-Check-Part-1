import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from datetime import datetime

st.set_page_config(
    page_title="RPS-BOT Dashboard",
    page_icon="🤖",
    layout="wide"
)

st.title("🤖 RPS-BOT Dashboard")
st.caption(f"Last refreshed: {datetime.now().strftime('%d %B %Y, %H:%M:%S')}")

st.markdown("---")

# ── Load Data ──────────────────────────────────────────────────────────────
@st.cache_data
def load_data():
    wb = load_workbook("RPS-BOT.xlsx", read_only=True, data_only=True)
    ws = wb.active
    headers = [cell for cell in next(ws.iter_rows(max_row=1, values_only=True))]
    rows = [row for row in ws.iter_rows(min_row=2, values_only=True)]
    df = pd.DataFrame(rows, columns=headers)
    return df

df = load_data()

# ── Summary Metrics ────────────────────────────────────────────────────────
total      = len(df)
updated    = len(df[df["Status"] == "Updated"])
not_updated= len(df[df["Status"] == "Not Updated"])
failed     = len(df[df["Status"] == "Failed"])
no_data    = len(df[df["Status"].isna()])

col1, col2, col3, col4, col5 = st.columns(5)
col1.metric("📋 Total Sources",  total)
col2.metric("🟢 Updated",        updated)
col3.metric("🔵 Not Updated",    not_updated)
col4.metric("🔴 Failed",         failed)
col5.metric("⚪ No Data",        no_data)

st.markdown("---")

# ── Filter ─────────────────────────────────────────────────────────────────
st.subheader("📊 Source Status Table")

status_options = ["All"] + [s for s in ["Updated", "Not Updated", "Failed"] if s in df["Status"].values]
selected_status = st.selectbox("Filter by Status", status_options)

if selected_status != "All":
    filtered_df = df[df["Status"] == selected_status].reset_index(drop=True)
else:
    filtered_df = df.reset_index(drop=True)

# ── Status Badge Styling ───────────────────────────────────────────────────
def style_status(val):
    if val == "Updated":
        return "background-color: #d4edda; color: #155724; font-weight: bold; border-radius: 4px; padding: 2px 8px;"
    elif val == "Not Updated":
        return "background-color: #d6d8db; color: #383d41; font-weight: bold; border-radius: 4px; padding: 2px 8px;"
    elif val == "Failed":
        return "background-color: #f8d7da; color: #721c24; font-weight: bold; border-radius: 4px; padding: 2px 8px;"
    return ""

styled_df = filtered_df.style.applymap(style_status, subset=["Status"])
st.dataframe(styled_df, use_container_width=True, height=280)

st.markdown("---")

# ── Updated Sources Detail ─────────────────────────────────────────────────
st.subheader("✅ Updated Sources")
updated_df = df[df["Status"] == "Updated"][["RPL-TYPES", "Previous Data", "Current Data", "Tracking Number", "Creation Date"]]

if updated_df.empty:
    st.info("No sources have been updated in the latest run.")
else:
    st.dataframe(updated_df.reset_index(drop=True), use_container_width=True)

st.markdown("---")

# ── Failed Sources Detail ──────────────────────────────────────────────────
st.subheader("❌ Failed Sources")
failed_df = df[df["Status"] == "Failed"][["RPL-TYPES", "Previous Data", "Current Data"]]

if failed_df.empty:
    st.success("No sources failed in the latest run.")
else:
    st.dataframe(failed_df.reset_index(drop=True), use_container_width=True)

st.markdown("---")

# ── Tracking Numbers ───────────────────────────────────────────────────────
st.subheader("📦 Tracking Numbers")
tn_df = df[df["Tracking Number"].notna()][["RPL-TYPES", "Tracking Number", "Creation Date"]]

if tn_df.empty:
    st.info("No tracking numbers have been created yet.")
else:
    st.dataframe(tn_df.reset_index(drop=True), use_container_width=True)

st.markdown("---")
st.caption("RPS-BOT © 2026 | Data sourced from RPS-BOT.xlsx")
