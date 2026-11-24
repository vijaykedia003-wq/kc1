import io
from datetime import datetime

import pandas as pd
import plotly.express as px
import streamlit as st

st.set_page_config(
    page_title="Kedia Capital | 360 Management Dashboard",
    page_icon="📊",
    layout="wide",
)

st.markdown(
    """
    <style>
    .block-container {padding-top: 1.5rem; padding-bottom: 2rem;}
    [data-testid="stMetric"] {
        background: #f5f8fc; border: 1px solid #e5eaf2;
        padding: 14px 16px; border-radius: 12px;
    }
    h1, h2, h3 {color: #0B1F3A;}
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("Kedia Capital – 360° Management Dashboard")
st.caption("Upload the four source Excel files to generate client activity and holdings reports.")

HOLDING_COLUMNS = [
    "Sr. No.", "Back Office Code_1", "Client ID", "Name", "A/c Open Date",
    "Account Status", "Correspondence Address", "Permanent Address",
    "Back Office Code", "Branch Code", "Branch Name", "PAN No.",
    "Type", "Sub Type", "Occupation", "Holding Value",
]


def clean_key(series):
    """Normalize identifiers without converting missing values to the text 'nan'."""
    return series.fillna("").astype(str).str.strip()


def read_excel_file(uploaded_file, header=0):
    try:
        return pd.read_excel(uploaded_file, header=header)
    except Exception as exc:
        raise ValueError(f"Could not read {uploaded_file.name}: {exc}") from exc


def prepare_holding_data(uploaded_file):
    raw = read_excel_file(uploaded_file, header=None)
    if len(raw) < 4:
        raise ValueError(
            "The Holding Value workbook must contain at least four rows; "
            "the expected column headings are on the third row."
        )

    # The supplied notebook takes row index 2 as the header and data from row index 3.
    raw.columns = range(raw.shape[1])
    headers = raw.iloc[2].astype(str).str.strip().tolist()
    data = raw.iloc[3:].copy()
    if len(headers) != data.shape[1]:
        raise ValueError("Could not identify the Holding Value sheet columns.")

    # The notebook standardizes this report to 16 known columns.
    if len(headers) != len(HOLDING_COLUMNS):
        # Use the source's known 16-column layout, but report unexpected layouts clearly.
        raise ValueError(
            f"Expected 16 columns in the Holding Value report after its third-row header; "
            f"found {len(headers)}. Please check the workbook format."
        )

    data.columns = HOLDING_COLUMNS
    data = data.dropna(subset=["Back Office Code", "Client ID"], how="all").copy()
    data["Back Office Code"] = clean_key(data["Back Office Code"])
    data["Holding Value"] = pd.to_numeric(data["Holding Value"], errors="coerce").fillna(0)
    data = data[data["Back Office Code"] != ""]
    return data


def process_files(ucc_file, activation_file, hierarchy_file, holding_file):
    ucc = read_excel_file(ucc_file)
    activation = read_excel_file(activation_file)
    hierarchy = read_excel_file(hierarchy_file)
    holdings = prepare_holding_data(holding_file)

    required = {
        "UCC Combine List": (ucc, ["Client Code", "Status", "Branch", "Region"]),
        "Segment Activation": (activation, ["Client Code"]),
        "TL BM Sheet": (hierarchy, ["Code", "Name", "TL", "BM"]),
    }
    for label, (df, cols) in required.items():
        missing = [col for col in cols if col not in df.columns]
        if missing:
            raise ValueError(f"{label} is missing required column(s): {', '.join(missing)}")

    ucc = ucc.copy()
    activation = activation.copy()
    hierarchy = hierarchy.copy()

    ucc["Client Code"] = clean_key(ucc["Client Code"])
    activation["Client Code"] = clean_key(activation["Client Code"])
    ucc["Branch"] = clean_key(ucc["Branch"])
    hierarchy["Code"] = clean_key(hierarchy["Code"])

    holding_sum = (
        holdings.groupby("Back Office Code", as_index=False)["Holding Value"]
        .sum()
    )

    merged = ucc.merge(
        holding_sum,
        left_on="Client Code",
        right_on="Back Office Code",
        how="left",
    )
    merged["Holding Value"] = merged["Holding Value"].fillna(0)

    active_clients = set(activation.loc[activation["Client Code"] != "", "Client Code"])
    disabled = merged["Status"].fillna("").astype(str).str.strip().str.lower().eq("disabled")
    merged["Activity Status"] = "Inactive"
    merged.loc[merged["Client Code"].isin(active_clients), "Activity Status"] = "Active"
    merged.loc[disabled, "Activity Status"] = "Disabled"

    hierarchy_df = hierarchy[["Code", "Name", "TL", "BM"]].rename(
        columns={"Code": "Branch", "Name": "BD"}
    )
    hierarchy_df = hierarchy_df.drop_duplicates(subset=["Branch"], keep="first")
    final_df = merged.merge(hierarchy_df, on="Branch", how="left")

    for col, default in [
        ("TL", "Unknown TL"),
        ("BM", "Unknown BM"),
        ("BD", "Unknown BD"),
        ("Region", "Unknown Region"),
    ]:
        if col not in final_df.columns:
            final_df[col] = default
        final_df[col] = final_df[col].fillna(default).astype(str).str.strip()
        final_df.loc[final_df[col] == "", col] = default

    final_df["Holding Value (Lakhs)"] = final_df["Holding Value"] / 100000
    return final_df, holdings


def make_summary(data, group_col):
    counts = (
        data.groupby([group_col, "Activity Status"])
        .size()
        .unstack(fill_value=0)
    )
    for status in ["Active", "Inactive", "Disabled"]:
        if status not in counts.columns:
            counts[status] = 0

    holding_totals = data.groupby(group_col)["Holding Value (Lakhs)"].sum()
    counts["Total Clients"] = counts[["Active", "Inactive", "Disabled"]].sum(axis=1)
    counts["Total Holdings Value (Lakhs)"] = holding_totals
    return counts[
        ["Active", "Inactive", "Disabled", "Total Clients", "Total Holdings Value (Lakhs)"]
    ].sort_values("Total Holdings Value (Lakhs)", ascending=False)


def excel_export(data, summaries):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        data.to_excel(writer, sheet_name="Client Master", index=False)
        for sheet_name, summary in summaries.items():
            summary.to_excel(writer, sheet_name=sheet_name[:31])
    output.seek(0)
    return output.getvalue()


with st.sidebar:
    st.header("1. Upload source files")
    ucc_file = st.file_uploader("A. Master UCC list", type=["xlsx", "xls"], key="ucc")
    activation_file = st.file_uploader("B. Exchange-wise segment activation", type=["xlsx", "xls"], key="activation")
    hierarchy_file = st.file_uploader("C. TL / BM hierarchy sheet", type=["xlsx", "xls"], key="hierarchy")
    holding_file = st.file_uploader("D. UCC with holding value", type=["xlsx", "xls"], key="holding")
    process = st.button(
        "Process & compile dashboard",
        type="primary",
        use_container_width=True,
        disabled=not all([ucc_file, activation_file, hierarchy_file, holding_file]),
    )
    st.caption("Files are processed in this session. Do not upload files containing data you are not authorized to share.")

if process:
    try:
        with st.spinner("Processing the four Excel files..."):
            result, holding_rows = process_files(
                ucc_file, activation_file, hierarchy_file, holding_file
            )
        st.session_state["client_data"] = result
        st.session_state["processed_at"] = datetime.now().strftime("%d-%b-%Y %I:%M %p")
        st.success("Files processed successfully.")
    except Exception as exc:
        st.error(str(exc))

if "client_data" not in st.session_state:
    st.info("Upload all four Excel files in the left panel, then select **Process & compile dashboard**.")
    st.markdown(
        """
        **Expected source columns**
        - Master UCC: `Client Code`, `Status`, `Branch`, `Region`
        - Segment activation: `Client Code`
        - TL / BM hierarchy: `Code`, `Name`, `TL`, `BM`
        - Holding Value: 16-column report with its headings on the third row, including
          `Back Office Code`, `Client ID`, and `Holding Value`.
        """
    )
    st.stop()

all_data = st.session_state["client_data"].copy()

st.caption(f"Last processed: {st.session_state.get('processed_at', '—')}")

with st.sidebar:
    st.divider()
    st.header("2. Dashboard filters")
    regions = ["All"] + sorted(all_data["Region"].dropna().unique().tolist())
    selected_region = st.selectbox("Region", regions)
    region_data = all_data if selected_region == "All" else all_data[all_data["Region"] == selected_region]

    bms = ["All"] + sorted(region_data["BM"].dropna().unique().tolist())
    selected_bm = st.selectbox("Branch Manager (BM)", bms)
    bm_data = region_data if selected_bm == "All" else region_data[region_data["BM"] == selected_bm]

    tls = ["All"] + sorted(bm_data["TL"].dropna().unique().tolist())
    selected_tl = st.selectbox("Team Leader (TL)", tls)
    tl_data = bm_data if selected_tl == "All" else bm_data[bm_data["TL"] == selected_tl]

    status_options = ["All", "Active", "Inactive", "Disabled"]
    selected_status = st.selectbox("Activity Status", status_options)

filtered = tl_data if selected_status == "All" else tl_data[tl_data["Activity Status"] == selected_status]

total_clients = len(filtered)
active_count = int((filtered["Activity Status"] == "Active").sum())
inactive_count = int((filtered["Activity Status"] == "Inactive").sum())
disabled_count = int((filtered["Activity Status"] == "Disabled").sum())
total_lakhs = filtered["Holding Value (Lakhs)"].sum()

c1, c2, c3, c4, c5 = st.columns(5)
c1.metric("Total Clients", f"{total_clients:,}")
c2.metric("Active Clients", f"{active_count:,}")
c3.metric("Inactive Clients", f"{inactive_count:,}")
c4.metric("Disabled Clients", f"{disabled_count:,}")
c5.metric("Total Holdings", f"₹ {total_lakhs:,.2f} L")

tab_overview, tab_hierarchy, tab_worklist, tab_master = st.tabs(
    ["Overview & Analytics", "Management Summaries", "Reactivation Worklist", "Client Master"]
)

with tab_overview:
    left, right = st.columns(2)
    with left:
        status_counts = (
            filtered["Activity Status"].value_counts()
            .reindex(["Active", "Inactive", "Disabled"], fill_value=0)
            .rename_axis("Status").reset_index(name="Clients")
        )
        fig = px.pie(
            status_counts, names="Status", values="Clients", hole=0.48,
            title="Client Activity Distribution",
            color="Status",
            color_discrete_map={"Active": "#22A06B", "Inactive": "#E5A000", "Disabled": "#D64545"},
        )
        st.plotly_chart(fig, use_container_width=True)
    with right:
        region_summary = make_summary(filtered, "Region").reset_index()
        region_long = region_summary.melt(
            id_vars=["Region"],
            value_vars=["Active", "Inactive", "Disabled"],
            var_name="Status",
            value_name="Clients",
        )
        fig = px.bar(
            region_long, x="Region", y="Clients", color="Status",
            barmode="stack", title="Region-wise Client Distribution",
            color_discrete_map={"Active": "#22A06B", "Inactive": "#E5A000", "Disabled": "#D64545"},
        )
        st.plotly_chart(fig, use_container_width=True)

    st.subheader("Region-wise Summary")
    st.dataframe(make_summary(filtered, "Region"), use_container_width=True)

with tab_hierarchy:
    st.subheader("Branch Manager (BM) Summary")
    st.dataframe(make_summary(filtered, "BM"), use_container_width=True)
    st.subheader("Team Leader (TL) Summary")
    st.dataframe(make_summary(filtered, "TL"), use_container_width=True)
    st.subheader("Business Developer (BD) Summary")
    st.dataframe(make_summary(filtered, "BD"), use_container_width=True)

with tab_worklist:
    st.subheader("Inactive and Disabled Clients – Reactivation Worklist")
    worklist = filtered[filtered["Activity Status"].isin(["Inactive", "Disabled"])].sort_values(
        "Holding Value", ascending=False
    )
    display_cols = [
        col for col in [
            "Client Code", "Branch", "Region", "Activity Status",
            "Holding Value (Lakhs)", "TL", "BM", "BD"
        ] if col in worklist.columns
    ]
    st.dataframe(worklist[display_cols], use_container_width=True, height=500)

with tab_master:
    st.subheader("Client Master Data")
    st.dataframe(filtered, use_container_width=True, height=550)

summaries = {
    "Region Summary": make_summary(filtered, "Region"),
    "BM Summary": make_summary(filtered, "BM"),
    "TL Summary": make_summary(filtered, "TL"),
    "BD Summary": make_summary(filtered, "BD"),
}
st.divider()
st.download_button(
    "Download Excel Report",
    data=excel_export(filtered, summaries),
    file_name="Kedia_Capital_360_Management_Report.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    type="primary",
)
