import streamlit as st
import pandas as pd
import plotly.express as px


# ---------------------------------------------------------
# PAGE CONFIGURATION
# ---------------------------------------------------------

st.set_page_config(
    page_title="IT Compliance & Risk Analytics Dashboard",
    page_icon="📊",
    layout="wide"
)


# ---------------------------------------------------------
# LOAD DATA
# ---------------------------------------------------------

@st.cache_data
def load_data():
    df = pd.read_excel("ml_predictions.xlsx")

    # Clean column names
    df.columns = (
        df.columns
        .astype(str)
        .str.strip()
        .str.lower()
        .str.replace(" ", "_")
    )

    return df


df = load_data()


# ---------------------------------------------------------
# REQUIRED COLUMNS
# ---------------------------------------------------------

required_columns = [
    "server_id",
    "predicted_high_risk",
    "priority",
    "operating_system",
    "environment",
    "incident_category"
]

missing_columns = [
    col for col in required_columns
    if col not in df.columns
]


# ---------------------------------------------------------
# DATA VALIDATION
# ---------------------------------------------------------

if missing_columns:

    st.error(
        "The following required columns are missing from "
        "ml_predictions.xlsx:"
    )

    for column in missing_columns:
        st.write(f"- `{column}`")

    st.write("Available columns:")

    for column in df.columns:
        st.write(f"- `{column}`")

    st.stop()


# ---------------------------------------------------------
# DATA CLEANING
# ---------------------------------------------------------

df["predicted_high_risk"] = pd.to_numeric(
    df["predicted_high_risk"],
    errors="coerce"
).fillna(0)

df["predicted_high_risk"] = (
    df["predicted_high_risk"]
    .astype(int)
)

for column in [
    "priority",
    "operating_system",
    "environment",
    "incident_category"
]:

    df[column] = (
        df[column]
        .astype(str)
        .str.strip()
    )


# ---------------------------------------------------------
# TITLE
# ---------------------------------------------------------

st.markdown(
    """
    <h1 style="text-align:center;">
        CDC Infrastructure Compliance & High-Risk Server Dashboard
    </h1>
    """,
    unsafe_allow_html=True
)

st.markdown("---")


# ---------------------------------------------------------
# SIDEBAR FILTERS
# ---------------------------------------------------------

st.sidebar.header("Dashboard Filters")


environment_filter = st.sidebar.multiselect(
    "Environment",
    options=sorted(df["environment"].dropna().unique()),
    default=sorted(df["environment"].dropna().unique())
)


os_filter = st.sidebar.multiselect(
    "Operating System",
    options=sorted(df["operating_system"].dropna().unique()),
    default=sorted(df["operating_system"].dropna().unique())
)


priority_filter = st.sidebar.multiselect(
    "Priority",
    options=sorted(df["priority"].dropna().unique()),
    default=sorted(df["priority"].dropna().unique())
)


incident_filter = st.sidebar.multiselect(
    "Incident Category",
    options=sorted(
        df["incident_category"].dropna().unique()
    ),
    default=sorted(
        df["incident_category"].dropna().unique()
    )
)


# ---------------------------------------------------------
# APPLY FILTERS
# ---------------------------------------------------------

filtered_df = df[
    df["environment"].isin(environment_filter)
    &
    df["operating_system"].isin(os_filter)
    &
    df["priority"].isin(priority_filter)
    &
    df["incident_category"].isin(incident_filter)
].copy()


# ---------------------------------------------------------
# KPI CALCULATIONS
# ---------------------------------------------------------

total_servers = filtered_df["server_id"].count()

high_risk_servers = filtered_df[
    filtered_df["predicted_high_risk"] == 1
]["server_id"].count()


if total_servers > 0:

    high_risk_rate = (
        high_risk_servers / total_servers
    ) * 100

else:

    high_risk_rate = 0


# ---------------------------------------------------------
# KPI CARDS
# ---------------------------------------------------------

col1, col2, col3 = st.columns(3)


with col1:

    st.metric(
        label="Test Set Servers",
        value=f"{total_servers:,}"
    )


with col2:

    st.metric(
        label="Predicted High-Risk",
        value=f"{high_risk_servers:,}"
    )


with col3:

    st.metric(
        label="High Risk Rate",
        value=f"{high_risk_rate:.2f}%"
    )


st.markdown("")


# ---------------------------------------------------------
# HIGH-RISK DATA
# ---------------------------------------------------------

risk_df = filtered_df[
    filtered_df["predicted_high_risk"] == 1
].copy()


# ---------------------------------------------------------
# ROW 1 — PRIORITY + OPERATING SYSTEM
# ---------------------------------------------------------

col1, col2 = st.columns(2)


# ---------------------------------------------------------
# PRIORITY CHART
# ---------------------------------------------------------

with col1:

    priority_data = (
        risk_df
        .groupby("priority")
        .size()
        .reset_index(name="predicted_high_risk")
        .sort_values(
            "predicted_high_risk",
            ascending=False
        )
    )

    fig_priority = px.bar(
        priority_data,
        x="priority",
        y="predicted_high_risk",
        title="Predicted High-Risk Servers by Priority",
        labels={
            "priority": "Priority",
            "predicted_high_risk": "High-Risk Servers"
        }
    )

    fig_priority.update_layout(
        height=400,
        showlegend=False
    )

    st.plotly_chart(
        fig_priority,
        use_container_width=True
    )


# ---------------------------------------------------------
# OPERATING SYSTEM CHART
# ---------------------------------------------------------

with col2:

    os_data = (
        risk_df
        .groupby("operating_system")
        .size()
        .reset_index(name="predicted_high_risk")
        .sort_values(
            "predicted_high_risk",
            ascending=False
        )
    )

    fig_os = px.pie(
        os_data,
        names="operating_system",
        values="predicted_high_risk",
        hole=0.45,
        title="Predicted High-Risk Servers by Operating System"
    )

    fig_os.update_layout(
        height=400
    )

    st.plotly_chart(
        fig_os,
        use_container_width=True
    )


# ---------------------------------------------------------
# ROW 2 — ENVIRONMENT + INCIDENT CATEGORY
# ---------------------------------------------------------

col1, col2 = st.columns(2)


# ---------------------------------------------------------
# ENVIRONMENT CHART
# ---------------------------------------------------------

with col1:

    environment_data = (
        risk_df
        .groupby("environment")
        .size()
        .reset_index(name="predicted_high_risk")
        .sort_values(
            "predicted_high_risk",
            ascending=False
        )
    )

    fig_environment = px.bar(
        environment_data,
        x="environment",
        y="predicted_high_risk",
        title="Predicted High-Risk Servers by Environment",
        labels={
            "environment": "Environment",
            "predicted_high_risk": "High-Risk Servers"
        }
    )

    fig_environment.update_layout(
        height=450,
        showlegend=False
    )

    st.plotly_chart(
        fig_environment,
        use_container_width=True
    )


# ---------------------------------------------------------
# INCIDENT CATEGORY CHART
# ---------------------------------------------------------

with col2:

    incident_data = (
        risk_df
        .groupby("incident_category")
        .size()
        .reset_index(name="predicted_high_risk")
        .sort_values(
            "predicted_high_risk",
            ascending=False
        )
    )

    fig_incident = px.bar(
        incident_data,
        x="predicted_high_risk",
        y="incident_category",
        orientation="h",
        title="Predicted High-Risk by Incident Category",
        labels={
            "incident_category": "Incident Category",
            "predicted_high_risk": "High-Risk Servers"
        }
    )

    fig_incident.update_layout(
        height=450,
        showlegend=False
    )

    st.plotly_chart(
        fig_incident,
        use_container_width=True
    )


# ---------------------------------------------------------
# DATA SUMMARY
# ---------------------------------------------------------

st.markdown("---")

st.subheader("Filtered Data Summary")

st.write(
    f"Showing **{len(filtered_df):,}** records "
    f"after applying the selected filters."
)

st.dataframe(
    filtered_df,
    use_container_width=True,
    hide_index=True
)
