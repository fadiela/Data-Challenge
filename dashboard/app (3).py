import streamlit as st
import pandas as pd
import io

# =====================================
# PAGE CONFIG
# =====================================
st.set_page_config(
    page_title="Norm Dashboard",
    page_icon="📊",
    layout="wide"
)

CSV_PATH = "output/norm_database_final.csv"

GRADE_ORDER = {"Top 25%": 1, "Average 50%": 2, "Bottom 25%": 3}

# =====================================
# HELPERS
# =====================================
@st.cache_data
def load_data():
    df = pd.read_csv(CSV_PATH)

    # Samakan nama kolom jika masih versi lama
    df = df.rename(columns={
        "parameter": "parameter_name",
        "scale": "scale_value"
    })

    df["filter_value"] = df["filter_value"].astype(str)
    df["scale_value"] = df["scale_value"].astype(str)

    return df


def format_table(df):
    display_df = df.copy()

    display_df = display_df.rename(columns={
        "parameter_name": "Parameter",
        "scale_value": "Scale",
        "norm_grade": "Norm Grade",
        "metric": "Metric",
        "filter_type": "Filter Type",
        "filter_value": "Filter Value",
        "base_n": "Base N",
        "norm_value": "Norm Value"
    })

    return display_df


def to_excel_bytes(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Hasil Filter")
    output.seek(0)
    return output


def summary_card(label, value):
    return f"""
    <div class="summary-card">
        <div class="summary-label">{label}</div>
        <div class="summary-value">{value}</div>
    </div>
    """


def metric_card(label, value):
    return f"""
    <div class="metric-card">
        <div class="metric-label">{label}</div>
        <div class="metric-value">{value}</div>
    </div>
    """


def join_or_dash(values, limit=4):
    if not values:
        return "-"
    values = list(values)
    if len(values) > limit:
        return f"{', '.join(values[:limit])}, +{len(values) - limit} lainnya"
    return ", ".join(values)


def cascading_multiselect(label, options, key_prefix, dep_key, default_all=True, help_text=None):
    """
    Multiselect whose widget key changes whenever its upstream dependency
    (dep_key, anything hashable e.g. tuple of upstream selections) changes.
    This forces Streamlit to re-evaluate the default (select-all) instead of
    crashing because a previous selection no longer exists in the new options.
    """
    options = sorted(options)
    widget_key = f"{key_prefix}__{hash(dep_key)}"

    # Only set an initial value the first time this widget_key appears
    # (i.e. right after upstream filters changed and produced a new key).
    if widget_key not in st.session_state:
        st.session_state[widget_key] = options if default_all else []

    top_row = st.columns([1, 1])
    with top_row[0]:
        select_all = st.button("Pilih Semua", key=f"{widget_key}_all", use_container_width=True)
    with top_row[1]:
        clear_all = st.button("Kosongkan", key=f"{widget_key}_clear", use_container_width=True)

    # IMPORTANT: never pass `default=` together with a `key` whose value we
    # also mutate via session_state in the same run - Streamlit raises an
    # exception if both are set in one script execution. So selection state
    # is managed entirely through session_state here.
    if select_all:
        st.session_state[widget_key] = options
    if clear_all:
        st.session_state[widget_key] = []

    selected = st.multiselect(
        label,
        options,
        key=widget_key,
        help=help_text
    )
    return selected


# =====================================
# CUSTOM CSS
# =====================================
st.markdown("""
<style>
/* Global */
.stApp {
    background: #07111f;
    color: #ffffff;
}

.block-container {
    padding-top: 2rem;
    padding-bottom: 2rem;
    max-width: 1250px;
}

/* Sidebar */
[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #19304f 0%, #162943 100%);
    border-right: 1px solid rgba(255,255,255,0.08);
}

[data-testid="stSidebar"] * {
    color: white;
}

[data-testid="stSidebar"] .stSelectbox label,
[data-testid="stSidebar"] .stMultiSelect label {
    font-weight: 700;
}

[data-testid="stSidebar"] div[data-baseweb="select"] > div {
    background-color: #233d5c !important;
    border: 1px solid #3f5f86 !important;
    border-radius: 10px !important;
}

[data-testid="stSidebar"] span[data-baseweb="tag"] {
    background-color: #2f6db5 !important;
}

[data-testid="stSidebar"] .stButton button {
    background: #233d5c !important;
    color: #ffffff !important;
    border: 1px solid #3f5f86 !important;
    border-radius: 8px !important;
    font-size: 12px !important;
    font-weight: 700 !important;
    padding: 0.25rem 0.4rem !important;
}

[data-testid="stSidebar"] .stButton button:hover {
    background: #2f547c !important;
}

/* Sidebar brand */
.sidebar-brand {
    padding: 8px 0 18px 0;
    margin-bottom: 18px;
    border-bottom: 1px solid rgba(255,255,255,0.12);
}

.brand-row {
    display: flex;
    align-items: center;
    gap: 12px;
    margin-bottom: 8px;
}

.brand-icon {
    font-size: 26px;
}

.brand-main {
    font-size: 24px;
    font-weight: 900;
    letter-spacing: 0.02em;
}

.brand-sub {
    font-size: 12px;
    font-weight: 700;
    letter-spacing: 0.18em;
    color: #dbe8f5;
    margin-left: 42px;
}

.filter-group-label {
    font-size: 13px;
    font-weight: 900;
    letter-spacing: 0.06em;
    text-transform: uppercase;
    color: #9fc1e8;
    margin-top: 18px;
    margin-bottom: 4px;
}

/* Title */
.page-title {
    font-size: 44px;
    font-weight: 900;
    color: #ffffff;
    margin-bottom: 4px;
}

.page-subtitle {
    font-size: 15px;
    color: #9fb1c4;
    margin-bottom: 28px;
}

/* Cards */
.summary-card {
    background: #ffffff;
    border-radius: 16px;
    padding: 20px 24px;
    border: 1px solid #d6dde8;
    box-shadow: 0 8px 22px rgba(0,0,0,0.15);
    min-height: 110px;
    margin-bottom: 14px;
}

.summary-label {
    font-size: 12px;
    font-weight: 900;
    letter-spacing: 0.16em;
    text-transform: uppercase;
    color: #35506f;
    margin-bottom: 12px;
}

.summary-value {
    font-size: 16px;
    font-weight: 900;
    color: #081522;
    line-height: 1.35;
    word-break: break-word;
}

.metric-card {
    background: linear-gradient(135deg, #16263d 0%, #20385b 100%);
    border: 1px solid #35537d;
    border-radius: 18px;
    padding: 24px 28px;
    box-shadow: 0 10px 24px rgba(0,0,0,0.22);
    min-height: 125px;
    margin-bottom: 6px;
}

.metric-label {
    font-size: 14px;
    font-weight: 800;
    color: #d4deea;
    margin-bottom: 14px;
}

.metric-value {
    font-size: 44px;
    font-weight: 900;
    color: #ffffff;
    line-height: 1.1;
}

/* Section title */
.section-title {
    font-size: 30px;
    font-weight: 900;
    color: #ffffff;
    margin-top: 26px;
    margin-bottom: 14px;
}

/* DataFrame */
[data-testid="stDataFrame"] {
    border: 1px solid rgba(255,255,255,0.08);
    border-radius: 14px;
    overflow: hidden;
}

/* Download buttons */
.stDownloadButton button {
    width: 100%;
    background: #d9e6f3 !important;
    color: #0b1a2b !important;
    border: 1px solid #5a7598 !important;
    border-radius: 10px !important;
    font-weight: 800 !important;
    padding: 0.55rem 1rem !important;
}

.stDownloadButton button:hover {
    background: #edf4fb !important;
}

/* Alert */
[data-testid="stInfo"], [data-testid="stWarning"] {
    border-radius: 12px;
}
</style>
""", unsafe_allow_html=True)

# =====================================
# LOAD DATA
# =====================================
try:
    df = load_data()
except Exception as e:
    st.error("Gagal membaca file CSV.")
    st.exception(e)
    st.stop()

# =====================================
# SIDEBAR
# Parameter & Scale tetap dropdown (single-select).
# Norm Grade, Metric, Filter Type, Filter Value jadi checkbox (multi-select).
# =====================================
with st.sidebar:
    st.markdown("""
    <div class="sidebar-brand">
        <div class="brand-row">
            <div class="brand-icon">🔷</div>
            <div class="brand-main">DEKA INSIGHT</div>
        </div>
        <div class="brand-sub">NORM DASHBOARD</div>
    </div>
    """, unsafe_allow_html=True)

    st.header("Filter Benchmark")
    st.caption("Norm Grade, Metric, Filter Type, dan Filter Value bisa pilih lebih dari satu.")

    # ---- Parameter (single, dropdown) ----
    parameter_options = sorted(df["parameter_name"].dropna().unique())
    selected_parameter = st.selectbox(
        "Parameter",
        parameter_options,
        index=parameter_options.index("Overall Liking") if "Overall Liking" in parameter_options else 0
    )

    # ---- Scale (single, dropdown) ----
    scale_options = sorted(
        df[df["parameter_name"] == selected_parameter]["scale_value"].dropna().unique()
    )
    selected_scale = st.selectbox("Scale", scale_options)

    # ---- Norm Grade (multi, checkbox-style) ----
    st.markdown('<div class="filter-group-label">Norm Grade</div>', unsafe_allow_html=True)
    norm_grade_options = ["Top 25%", "Average 50%", "Bottom 25%"]
    selected_norm_grades = cascading_multiselect(
        "Norm Grade", norm_grade_options, "sel_norm_grade", dep_key="static"
    )

    # ---- Metric (multi, checkbox-style) ----
    st.markdown('<div class="filter-group-label">Metric</div>', unsafe_allow_html=True)
    metric_options = ["TB%", "T2B%", "T3B%", "MS"]
    selected_metrics = cascading_multiselect(
        "Metric", metric_options, "sel_metric", dep_key="static"
    )

    # ---- Filter Type (multi, checkbox-style, cascades on Parameter+Scale) ----
    st.markdown('<div class="filter-group-label">Filter Type</div>', unsafe_allow_html=True)
    filter_type_options = df[
        (df["parameter_name"] == selected_parameter) &
        (df["scale_value"] == selected_scale)
    ]["filter_type"].dropna().unique().tolist()
    selected_filter_types = cascading_multiselect(
        "Filter Type", filter_type_options, "sel_filter_type",
        dep_key=(selected_parameter, selected_scale)
    )

    # ---- Filter Value (multi, checkbox-style, cascades on Parameter+Scale+Filter Type) ----
    st.markdown('<div class="filter-group-label">Filter Value</div>', unsafe_allow_html=True)
    fval_pool = df[
        (df["parameter_name"] == selected_parameter) &
        (df["scale_value"] == selected_scale) &
        (df["filter_type"].isin(selected_filter_types))
    ] if selected_filter_types else df.iloc[0:0]
    filter_value_options = fval_pool["filter_value"].dropna().unique().tolist()
    selected_filter_values = cascading_multiselect(
        "Filter Value", filter_value_options, "sel_filter_value",
        dep_key=(selected_parameter, selected_scale, tuple(sorted(selected_filter_types)))
    )

# =====================================
# FILTER DATA -> SATU TABEL GABUNGAN
# =====================================
has_any_selection = all([
    selected_norm_grades, selected_metrics, selected_filter_types, selected_filter_values
])

if has_any_selection:
    result_df = df[
        (df["parameter_name"] == selected_parameter) &
        (df["scale_value"] == selected_scale) &
        df["norm_grade"].isin(selected_norm_grades) &
        df["metric"].isin(selected_metrics) &
        df["filter_type"].isin(selected_filter_types) &
        df["filter_value"].isin(selected_filter_values)
    ].copy()

    result_df["__grade_order"] = result_df["norm_grade"].map(GRADE_ORDER)
    result_df = result_df.sort_values(
        ["filter_type", "filter_value", "__grade_order", "metric"]
    ).drop(columns=["__grade_order"])
else:
    result_df = df.iloc[0:0].copy()

# =====================================
# MAIN HEADER
# =====================================
st.markdown('<div class="page-title">📊 Norm Dashboard</div>', unsafe_allow_html=True)
st.markdown('<div class="page-subtitle">Dashboard benchmark norm database berdasarkan hasil Python.</div>', unsafe_allow_html=True)

# =====================================
# SUMMARY CARDS
# =====================================
c1, c2, c3, c4 = st.columns(4)

with c1:
    st.markdown(summary_card("Parameter", selected_parameter), unsafe_allow_html=True)

with c2:
    st.markdown(summary_card("Scale", selected_scale), unsafe_allow_html=True)

with c3:
    st.markdown(summary_card("Filter Type Dipilih", join_or_dash(selected_filter_types)), unsafe_allow_html=True)

with c4:
    st.markdown(summary_card("Filter Value Dipilih", join_or_dash(selected_filter_values)), unsafe_allow_html=True)

# =====================================
# BENCHMARK RESULT
# Card "Norm Value" & "Base N" hanya tampil kalau hasil persis 1 baris,
# karena dengan multi-select Norm Grade/Metric/Filter Type/Filter Value
# kombinasinya bisa banyak sekaligus.
# =====================================
st.markdown('<div class="section-title">Hasil Benchmark</div>', unsafe_allow_html=True)

if not has_any_selection:
    st.info("Pilih minimal satu Norm Grade, Metric, Filter Type, dan Filter Value pada sidebar untuk menampilkan hasil.")
elif result_df.empty:
    st.warning("Tidak ada data untuk kombinasi filter ini.")
elif len(result_df) == 1:
    row = result_df.iloc[0]
    norm_value = row["norm_value"]
    base_n = row["base_n"]
    selected_metric_single = row["metric"]

    if pd.isna(norm_value):
        norm_display = "-"
    else:
        norm_display = f"{norm_value:.2f}" if selected_metric_single == "MS" else f"{norm_value:.2f}%"

    base_n_display = "-" if pd.isna(base_n) else f"{int(base_n)}"

    m1, m2 = st.columns(2)
    with m1:
        st.markdown(metric_card("Norm Value", norm_display), unsafe_allow_html=True)
    with m2:
        st.markdown(metric_card("Base N", base_n_display), unsafe_allow_html=True)
else:
    avg_norm_value = result_df["norm_value"].mean(skipna=True)
    total_base_n = result_df["base_n"].sum(skipna=True)

    avg_norm_display = "-" if pd.isna(avg_norm_value) else f"{avg_norm_value:.2f}"
    total_base_n_display = "-" if pd.isna(total_base_n) else f"{int(total_base_n)}"

    m1, m2 = st.columns(2)
    with m1:
        st.markdown(metric_card(f"Rata-rata Norm Value ({len(result_df)} baris)", avg_norm_display), unsafe_allow_html=True)
    with m2:
        st.markdown(metric_card(f"Total Base N ({len(result_df)} baris)", total_base_n_display), unsafe_allow_html=True)

    metrics_in_result = sorted(result_df["metric"].dropna().unique())
    if len(metrics_in_result) > 1:
        st.caption(
            "⚠️ Rata-rata Norm Value di atas menggabungkan beberapa Metric sekaligus "
            f"({', '.join(metrics_in_result)}), yang skalanya berbeda (MS = 1-9, %-metric = 0-100). "
            "Untuk angka per Metric/Norm Grade yang lebih akurat, lihat tabel detail di bawah."
        )

# =====================================
# DETAIL TABLE
# =====================================
st.markdown('<div class="section-title">Detail Data Terfilter</div>', unsafe_allow_html=True)

if result_df.empty:
    st.info("Belum ada data yang dapat ditampilkan.")
else:
    display_result_df = format_table(result_df)
    st.dataframe(display_result_df, use_container_width=True, hide_index=True)

    d1, d2 = st.columns(2)

    csv_detail = display_result_df.to_csv(index=False).encode("utf-8-sig")
    excel_detail = to_excel_bytes(display_result_df)

    with d1:
        st.download_button(
            label="⬇ Download Detail (CSV)",
            data=csv_detail,
            file_name="detail_data_terfilter.csv",
            mime="text/csv"
        )

    with d2:
        st.download_button(
            label="⬇ Download Detail (Excel)",
            data=excel_detail,
            file_name="detail_data_terfilter.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
