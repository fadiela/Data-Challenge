import streamlit as st
import pandas as pd
from sqlalchemy import create_engine


# =========================
# KONFIGURASI DASHBOARD
# =========================
st.set_page_config(
    page_title="Norm Database Dashboard",
    page_icon="📊",
    layout="wide"
)

MYSQL_USER = "root"
MYSQL_PASSWORD = "setiabudi_28"
MYSQL_HOST = "localhost"
MYSQL_PORT = "3306"
MYSQL_DATABASE = "norm_deka_insight"
MYSQL_TABLE = "norm_results"


# =========================
# KONEKSI MYSQL
# =========================
@st.cache_resource
def get_engine():
    engine = create_engine(
        f"mysql+pymysql://{MYSQL_USER}:{MYSQL_PASSWORD}@{MYSQL_HOST}:{MYSQL_PORT}/{MYSQL_DATABASE}"
    )
    return engine


@st.cache_data
def load_data():
    engine = get_engine()
    query = f"""
        SELECT
            parameter_name,
            scale_value,
            norm_grade,
            metric,
            filter_type,
            filter_value,
            base_n,
            norm_value
        FROM {MYSQL_TABLE}
    """
    df = pd.read_sql(query, engine)
    return df


# =========================
# LOAD DATA
# =========================
st.title("📊 Norm Database Dashboard")
st.caption("Dashboard benchmark norm database berdasarkan hasil Python dan MySQL.")

try:
    df = load_data()
except Exception as e:
    st.error("Gagal koneksi ke MySQL atau gagal membaca tabel.")
    st.exception(e)
    st.stop()


# =========================
# SIDEBAR FILTER
# =========================
st.sidebar.header("Filter Benchmark")

parameter_options = sorted(df["parameter_name"].dropna().unique())
selected_parameter = st.sidebar.selectbox(
    "Parameter",
    parameter_options,
    index=parameter_options.index("Overall Liking") if "Overall Liking" in parameter_options else 0
)

scale_options = sorted(
    df[df["parameter_name"] == selected_parameter]["scale_value"].dropna().unique()
)
selected_scale = st.sidebar.selectbox("Scale", scale_options)

norm_grade_options = ["Top 25%", "Average 50%", "Bottom 25%"]
selected_norm_grade = st.sidebar.selectbox("Norm Grade", norm_grade_options)

metric_options = ["TB%", "T2B%", "T3B%", "MS"]
selected_metric = st.sidebar.selectbox("Metric", metric_options)

filter_type_options = sorted(
    df[
        (df["parameter_name"] == selected_parameter) &
        (df["scale_value"] == selected_scale)
    ]["filter_type"].dropna().unique()
)
selected_filter_type = st.sidebar.selectbox("Filter Type", filter_type_options)

filter_value_options = sorted(
    df[
        (df["parameter_name"] == selected_parameter) &
        (df["scale_value"] == selected_scale) &
        (df["filter_type"] == selected_filter_type)
    ]["filter_value"].dropna().astype(str).unique()
)

selected_filter_value = st.sidebar.selectbox("Filter Value", filter_value_options)


# =========================
# FILTER DATA
# =========================
result_df = df[
    (df["parameter_name"] == selected_parameter) &
    (df["scale_value"] == selected_scale) &
    (df["norm_grade"] == selected_norm_grade) &
    (df["metric"] == selected_metric) &
    (df["filter_type"] == selected_filter_type) &
    (df["filter_value"].astype(str) == selected_filter_value)
].copy()


# =========================
# MAIN RESULT
# =========================
st.subheader("Hasil Benchmark")

col1, col2, col3 = st.columns(3)

if result_df.empty:
    col1.metric("Norm Value", "-")
    col2.metric("Base N", "-")
    col3.metric("Jumlah Baris", 0)
    st.warning("Tidak ada data untuk kombinasi filter ini.")
else:
    norm_value = result_df.iloc[0]["norm_value"]
    base_n = result_df.iloc[0]["base_n"]

    if pd.isna(norm_value):
        norm_display = "-"
    else:
        norm_display = f"{norm_value:.2f}" if selected_metric == "MS" else f"{norm_value:.2f}%"

    col1.metric("Norm Value", norm_display)
    col2.metric("Base N", int(base_n))
    col3.metric("Jumlah Baris", len(result_df))

st.divider()


# =========================
# DETAIL TABLE
# =========================
st.subheader("Detail Data Terfilter")
st.dataframe(result_df, use_container_width=True)


# =========================
# COMPARISON TABLE
# =========================
st.subheader("Perbandingan Top 25%, Average 50%, Bottom 25%")

comparison_df = df[
    (df["parameter_name"] == selected_parameter) &
    (df["scale_value"] == selected_scale) &
    (df["metric"] == selected_metric) &
    (df["filter_type"] == selected_filter_type) &
    (df["filter_value"].astype(str) == selected_filter_value)
].copy()

comparison_df = comparison_df.sort_values(["norm_grade"])

st.dataframe(comparison_df, use_container_width=True)


# =========================
# DOWNLOAD
# =========================
csv_data = comparison_df.to_csv(index=False).encode("utf-8")

st.download_button(
    label="Download hasil filter sebagai CSV",
    data=csv_data,
    file_name="hasil_filter_norm_database.csv",
    mime="text/csv"
)