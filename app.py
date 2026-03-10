import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(page_title="Auto Excel Dashboard", layout="wide")

st.title("Auto Excel Analytics Dashboard")

# -------------------------
# Function: fix duplicate columns
# -------------------------

def make_unique(columns):
    seen = {}
    new_cols = []

    for col in columns:
        if col in seen:
            seen[col] += 1
            new_cols.append(f"{col}_{seen[col]}")
        else:
            seen[col] = 0
            new_cols.append(col)

    return new_cols


# -------------------------
# Upload Excel
# -------------------------

uploaded_file = st.file_uploader(
    "Upload Excel File",
    type=["xlsx","xls","xlsm","xlsb"]
)

# -------------------------
# Month Year Selection
# -------------------------

col1, col2 = st.columns(2)

month = col1.selectbox(
    "Month",
    [
        "January","February","March","April",
        "May","June","July","August",
        "September","October","November","December"
    ]
)

year = col2.selectbox(
    "Year",
    list(range(2020,2035))
)

# -------------------------
# Process Excel
# -------------------------

if uploaded_file:

    df = pd.read_excel(uploaded_file)

    # clean column names
    df.columns = df.columns.astype(str).str.replace("\n"," ").str.strip()

    # remove empty columns
    df = df.loc[:, ~df.columns.str.contains('^Unnamed')]

    # fix duplicates
    df.columns = make_unique(df.columns)

    # remove empty rows
    df = df.dropna(how="all")

    st.success(f"Loaded {len(df)} rows for {month} {year}")

    # -------------------------
    # Detect column types
    # -------------------------

    numeric_cols = df.select_dtypes(include=["number"]).columns.tolist()

    categorical_cols = df.select_dtypes(
        include=["object"]
    ).columns.tolist()

    # -------------------------
    # Sidebar Filters
    # -------------------------

    st.sidebar.header("Filters")

    filtered_df = df.copy()

    for col in categorical_cols:

        unique_vals = df[col].dropna().unique()

        if len(unique_vals) < 50:  # avoid huge filters

            selected = st.sidebar.multiselect(
                col,
                unique_vals,
                default=unique_vals
            )

            filtered_df = filtered_df[
                filtered_df[col].isin(selected)
            ]

    # -------------------------
    # KPIs
    # -------------------------

    st.subheader("Key Metrics")

    kpi_cols = st.columns(4)

    kpi_cols[0].metric(
        "Total Rows",
        len(filtered_df)
    )

    if len(numeric_cols) > 0:

        total_value = filtered_df[numeric_cols[0]].sum()

        kpi_cols[1].metric(
            f"Total {numeric_cols[0]}",
            f"{total_value:,.2f}"
        )

    if len(categorical_cols) > 0:

        kpi_cols[2].metric(
            f"Unique {categorical_cols[0]}",
            filtered_df[categorical_cols[0]].nunique()
        )

    kpi_cols[3].metric(
        "Columns",
        len(filtered_df.columns)
    )

    st.divider()

    # -------------------------
    # Auto Charts
    # -------------------------

    st.subheader("Auto Visualizations")

    chart_cols = st.columns(2)

    chart_count = 0

    for cat in categorical_cols[:4]:

        if len(numeric_cols) > 0:

            num = numeric_cols[0]

            chart_data = (
                filtered_df.groupby(cat)[num]
                .sum()
                .sort_values(ascending=False)
                .head(10)
                .reset_index()
            )

            fig = px.bar(
                chart_data,
                x=cat,
                y=num,
                title=f"{num} by {cat}"
            )

            chart_cols[chart_count % 2].plotly_chart(
                fig,
                use_container_width=True
            )

            chart_count += 1

    # -------------------------
    # Correlation chart
    # -------------------------

    if len(numeric_cols) >= 2:

        st.subheader("Correlation")

        corr = filtered_df[numeric_cols].corr()

        fig_corr = px.imshow(
            corr,
            text_auto=True,
            title="Numeric Correlation Matrix"
        )

        st.plotly_chart(fig_corr, use_container_width=True)

    st.divider()

    # -------------------------
    # Data Table
    # -------------------------

    st.subheader("Filtered Data")

    st.dataframe(
        filtered_df,
        use_container_width=True
    )

else:

    st.info("Upload an Excel file to start the dashboard.")