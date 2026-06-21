import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import io

# =====================================
# PAGE CONFIG
# =====================================
st.set_page_config(
    page_title="Deka Insight Norm Dashboard",
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


def metric_card(label, value, tone=None):
    """tone: None (default navy) | 'good' | 'avg' | 'bad' -> changes border/accent color."""
    tone_class = f" metric-card--{tone}" if tone else ""
    return f"""
    <div class="metric-card{tone_class}">
        <div class="metric-label">{label}</div>
        <div class="metric-value">{value}</div>
    </div>
    """


def positioning_badge(label, tone):
    return f"""
    <div class="pos-badge pos-badge--{tone}">{label}</div>
    """


GRADE_TONE = {"Top 25%": "good", "Average 50%": "avg", "Bottom 25%": "bad"}


def weighted_grade_values(source_df, metric):
    """Hitung weighted-average norm_value per Norm Grade untuk satu metric tertentu,
    dari source_df yang sudah difilter parameter/scale/filter_type/filter_value
    (TANPA filter Norm Grade, supaya selalu dapat Top25/Avg50/Bottom25 lengkap)."""
    out = {}
    for grade in ["Top 25%", "Average 50%", "Bottom 25%"]:
        cell_df = source_df[
            (source_df["norm_grade"] == grade) & (source_df["metric"] == metric)
        ].dropna(subset=["norm_value"])
        weight_sum = cell_df["base_n"].sum(skipna=True)
        if cell_df.empty or pd.isna(weight_sum) or weight_sum == 0:
            out[grade] = None
        else:
            out[grade] = (cell_df["norm_value"] * cell_df["base_n"]).sum() / weight_sum
    return out


def classify_score(score, norms):
    """norms: dict dengan 'Top 25%', 'Average 50%', 'Bottom 25%'. Return (label, tone)."""
    top, bottom = norms.get("Top 25%"), norms.get("Bottom 25%")
    if top is None or bottom is None:
        return "Data norm tidak lengkap untuk posisi ini", "avg"
    if score >= top:
        return "Di Atas Norm — sejajar grup Top 25% (disukai)", "good"
    if score <= bottom:
        return "Di Bawah Norm — sejajar grup Bottom 25% (kurang disukai)", "bad"
    return "Sesuai Norm — sejajar grup Average 50%", "avg"


def join_or_dash(values, limit=4):
    if not values:
        return "-"
    values = list(values)
    if len(values) > limit:
        return f"{', '.join(values[:limit])}, +{len(values) - limit} lainnya"
    return ", ".join(values)


def cascading_multiselect(label, options, key_prefix, dep_key, help_text=None):
    """
    Multiselect whose widget key changes whenever its upstream dependency
    (dep_key, anything hashable e.g. tuple of upstream selections) changes.
    This forces Streamlit to re-evaluate the default instead of crashing
    because a previous selection no longer exists in the new options.

    Defaults to only the FIRST option selected (not select-all), so the
    user starts narrow and explicitly adds more values if they want to.
    """
    options = sorted(options)
    widget_key = f"{key_prefix}__{hash(dep_key)}"

    # Only set an initial value the first time this widget_key appears
    # (i.e. right after upstream filters changed and produced a new key).
    if widget_key not in st.session_state:
        st.session_state[widget_key] = options[:1]

    selected = st.multiselect(
        label,
        options,
        key=widget_key,
        help=help_text
    )
    return selected


# ---- Kombinasi 2-dimensi (mis. "Category + Gender" / "Food | Perempuan") ----
COMBO_SEP_TYPE = " + "
COMBO_SEP_VALUE = " | "


def is_combo_filter_type(filter_type: str) -> bool:
    return COMBO_SEP_TYPE in str(filter_type)


def split_combo_type(filter_type: str):
    """'Category + Gender' -> ('Category', 'Gender')"""
    a, b = str(filter_type).split(COMBO_SEP_TYPE, 1)
    return a, b


def split_combo_value(filter_value: str):
    """'Food | Perempuan' -> ('Food', 'Perempuan')"""
    a, b = str(filter_value).split(COMBO_SEP_VALUE, 1)
    return a, b


def make_combo_type(dim_a: str, dim_b: str) -> str:
    return f"{dim_a}{COMBO_SEP_TYPE}{dim_b}"


def make_combo_value(val_a: str, val_b: str) -> str:
    return f"{val_a}{COMBO_SEP_VALUE}{val_b}"


# =====================================
# CUSTOM CSS
# =====================================
st.markdown("""
<style>
/* Global */
.stApp {
    background: #f5f7fa;
    color: #1a2433;
}

.block-container {
    padding-top: 2rem;
    padding-bottom: 2rem;
    max-width: 1250px;
}

/* Sidebar */
[data-testid="stSidebar"] {
    background: #ffffff;
    border-right: 1px solid #e3e8ee;
}

[data-testid="stSidebar"] * {
    color: #1a2433;
}

[data-testid="stSidebar"] .stSelectbox label,
[data-testid="stSidebar"] .stMultiSelect label {
    font-weight: 700;
}

[data-testid="stSidebar"] div[data-baseweb="select"] > div {
    background-color: #f5f7fa !important;
    border: 1px solid #d6dde8 !important;
    border-radius: 10px !important;
}

[data-testid="stSidebar"] span[data-baseweb="tag"] {
    background-color: #2f6db5 !important;
    color: #ffffff !important;
}

/* Sidebar brand */
.sidebar-brand {
    padding: 8px 0 18px 0;
    margin-bottom: 18px;
    border-bottom: 1px solid #e3e8ee;
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
    color: #1a2433;
}

.brand-sub {
    font-size: 12px;
    font-weight: 700;
    letter-spacing: 0.18em;
    color: #5a7088;
    margin-left: 42px;
}

.filter-group-label {
    font-size: 13px;
    font-weight: 900;
    letter-spacing: 0.06em;
    text-transform: uppercase;
    color: #2f6db5;
    margin-top: 18px;
    margin-bottom: 4px;
}

/* Filter mode radio (Single Dimension / Kombinasi 2 Dimensi) */
[data-testid="stSidebar"] div[role="radiogroup"] {
    background: #f5f7fa;
    border: 1px solid #d6dde8;
    border-radius: 10px;
    padding: 6px 10px;
    margin-bottom: 6px;
}

.combo-hint {
    font-size: 12px;
    color: #5a7088;
    margin: -2px 0 10px 0;
    line-height: 1.4;
}

/* Title */
.page-title {
    font-size: 44px;
    font-weight: 900;
    color: #1a2433;
    margin-bottom: 4px;
}

.page-subtitle {
    font-size: 15px;
    color: #5a7088;
    margin-bottom: 28px;
}

/* Cards */
.summary-card {
    background: #ffffff;
    border-radius: 16px;
    padding: 20px 24px;
    border: 1px solid #d6dde8;
    box-shadow: 0 4px 12px rgba(0,0,0,0.04);
    min-height: 110px;
    margin-bottom: 14px;
}

.summary-label {
    font-size: 12px;
    font-weight: 900;
    letter-spacing: 0.16em;
    text-transform: uppercase;
    color: #5a7088;
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
    background: #ffffff;
    border: 1px solid #d6dde8;
    border-radius: 18px;
    padding: 24px 28px;
    box-shadow: 0 4px 12px rgba(0,0,0,0.04);
    min-height: 125px;
    margin-bottom: 6px;
}

.metric-label {
    font-size: 14px;
    font-weight: 800;
    color: #5a7088;
    margin-bottom: 14px;
}

.metric-value {
    font-size: 44px;
    font-weight: 900;
    color: #1a2433;
    line-height: 1.1;
}

/* Tone variants — dipakai untuk menandai posisi norm (good/avg/bad) */
.metric-card--good {
    background: #f0faf4;
    border: 1px solid #2f9e66;
}
.metric-card--good .metric-value { color: #1d7a44; }
.metric-card--avg {
    background: #ffffff;
    border: 1px solid #d6dde8;
}
.metric-card--bad {
    background: #fdf2f2;
    border: 1px solid #c0504d;
}
.metric-card--bad .metric-value { color: #b23b38; }

/* Positioning badge */
.pos-badge {
    display: inline-block;
    padding: 10px 18px;
    border-radius: 999px;
    font-weight: 900;
    font-size: 16px;
    margin-bottom: 14px;
}
.pos-badge--good { background: #f0faf4; color: #1d7a44; border: 1px solid #2f9e66; }
.pos-badge--avg  { background: #eef3fa; color: #2f6db5; border: 1px solid #a9c3e3; }
.pos-badge--bad  { background: #fdf2f2; color: #b23b38; border: 1px solid #c0504d; }

/* Section title */
.section-title {
    font-size: 30px;
    font-weight: 900;
    color: #1a2433;
    margin-top: 26px;
    margin-bottom: 14px;
}

/* DataFrame */
[data-testid="stDataFrame"] {
    border: 1px solid #e3e8ee;
    border-radius: 14px;
    overflow: hidden;
}

/* Download buttons */
.stDownloadButton button {
    width: 100%;
    background: #2f6db5 !important;
    color: #ffffff !important;
    border: 1px solid #2f6db5 !important;
    border-radius: 10px !important;
    font-weight: 800 !important;
    padding: 0.55rem 1rem !important;
}

.stDownloadButton button:hover {
    background: #25588f !important;
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
    # T3B% hanya valid untuk skala 7 ke atas (lihat dokumen "Penjelasan TB/T2B/T3B/MS")
    metric_options = ["TB%", "T2B%", "MS"]
    if int(selected_scale) >= 7:
        metric_options.append("T3B%")
    selected_metrics = cascading_multiselect(
        "Metric", metric_options, "sel_metric",
        dep_key=("metric_ge7" if int(selected_scale) >= 7 else "metric_lt7")
    )

    # ---- Mode Filter: Single Dimension vs Kombinasi 2 Dimensi ----
    st.markdown('<div class="filter-group-label">Mode Filter</div>', unsafe_allow_html=True)
    filter_mode = st.radio(
        "Mode Filter",
        ["Single Dimension", "Kombinasi 2 Dimensi"],
        key="filter_mode",
        label_visibility="collapsed",
        horizontal=False,
    )

    # Pool data untuk parameter+scale yang dipilih (dipakai kedua mode)
    base_pool = df[
        (df["parameter_name"] == selected_parameter) &
        (df["scale_value"] == selected_scale)
    ]

    if filter_mode == "Single Dimension":
        # ---- Filter Type (multi, checkbox-style, cascades on Parameter+Scale) ----
        st.markdown('<div class="filter-group-label">Filter Type</div>', unsafe_allow_html=True)
        filter_type_options = base_pool[
            ~base_pool["filter_type"].apply(is_combo_filter_type)
        ]["filter_type"].dropna().unique().tolist()
        selected_filter_types = cascading_multiselect(
            "Filter Type", filter_type_options, "sel_filter_type",
            dep_key=(selected_parameter, selected_scale)
        )

        # ---- Filter Value (multi, checkbox-style, cascades on Filter Type) ----
        st.markdown('<div class="filter-group-label">Filter Value</div>', unsafe_allow_html=True)
        fval_pool = base_pool[
            base_pool["filter_type"].isin(selected_filter_types)
        ] if selected_filter_types else base_pool.iloc[0:0]
        filter_value_options = fval_pool["filter_value"].dropna().unique().tolist()
        selected_filter_values = cascading_multiselect(
            "Filter Value", filter_value_options, "sel_filter_value",
            dep_key=(selected_parameter, selected_scale, tuple(sorted(selected_filter_types)))
        )

    else:
        # ---- Mode Kombinasi: pilih Dimensi A + Value A, Dimensi B + Value B ----
        # secara terpisah lewat dropdown biasa (user tidak perlu tahu format
        # gabungan "Category + Gender" / "Food | Perempuan" secara manual).
        st.markdown(
            '<div class="combo-hint">Pilih 2 dimensi sekaligus, mis. <b>Category</b> = Food '
            '<b>dan</b> <b>Gender</b> = Perempuan, untuk melihat norm pada irisan keduanya.</div>',
            unsafe_allow_html=True
        )

        combo_types_in_pool = base_pool[
            base_pool["filter_type"].apply(is_combo_filter_type)
        ]["filter_type"].dropna().unique().tolist()

        # Kumpulkan semua dimensi tunggal yang punya minimal 1 kombinasi tersedia
        all_dims = sorted({
            dim for ft in combo_types_in_pool for dim in split_combo_type(ft)
        })

        if not all_dims:
            st.warning("Tidak ada data kombinasi untuk Parameter & Scale ini.")
            selected_filter_types, selected_filter_values = [], []
        else:
            st.markdown('<div class="filter-group-label">Dimensi A</div>', unsafe_allow_html=True)
            dim_a_key = f"combo_dim_a__{hash((selected_parameter, selected_scale))}"
            if dim_a_key not in st.session_state:
                st.session_state[dim_a_key] = all_dims[0]
            dim_a = st.selectbox("Dimensi A", all_dims, key=dim_a_key, label_visibility="collapsed")

            # Dimensi B hanya menampilkan pasangan yang benar-benar ada datanya dengan Dimensi A
            dim_b_options = sorted({
                (set(split_combo_type(ft)) - {dim_a}).pop()
                for ft in combo_types_in_pool if dim_a in split_combo_type(ft)
            })

            st.markdown('<div class="filter-group-label">Dimensi B</div>', unsafe_allow_html=True)
            dim_b_key = f"combo_dim_b__{hash((selected_parameter, selected_scale, dim_a))}"
            if dim_b_key not in st.session_state:
                st.session_state[dim_b_key] = dim_b_options[0] if dim_b_options else None
            dim_b = st.selectbox(
                "Dimensi B", dim_b_options, key=dim_b_key, label_visibility="collapsed"
            ) if dim_b_options else None

            if dim_b is None:
                st.warning("Tidak ada pasangan kombinasi untuk Dimensi A ini.")
                selected_filter_types, selected_filter_values = [], []
            else:
                # filter_type di database tersimpan 1 arah saja (urutan sesuai
                # COMBO_DIMENSIONS saat database dibuat) -> coba kedua arah.
                resolved_combo_type = (
                    make_combo_type(dim_a, dim_b)
                    if make_combo_type(dim_a, dim_b) in combo_types_in_pool
                    else make_combo_type(dim_b, dim_a)
                )
                order_a_first = resolved_combo_type == make_combo_type(dim_a, dim_b)

                combo_pool = base_pool[base_pool["filter_type"] == resolved_combo_type]
                pair_values = combo_pool["filter_value"].dropna().unique().tolist()
                parsed_pairs = [split_combo_value(v) for v in pair_values]

                val_a_options = sorted({(p[0] if order_a_first else p[1]) for p in parsed_pairs})
                st.markdown('<div class="filter-group-label">Value Dimensi A</div>', unsafe_allow_html=True)
                val_a_selected = cascading_multiselect(
                    "Value Dimensi A", val_a_options, "sel_combo_val_a",
                    dep_key=(selected_parameter, selected_scale, dim_a, dim_b)
                )

                val_b_options = sorted({
                    (p[1] if order_a_first else p[0])
                    for p in parsed_pairs
                    if (p[0] if order_a_first else p[1]) in val_a_selected
                }) if val_a_selected else []
                st.markdown('<div class="filter-group-label">Value Dimensi B</div>', unsafe_allow_html=True)
                val_b_selected = cascading_multiselect(
                    "Value Dimensi B", val_b_options, "sel_combo_val_b",
                    dep_key=(selected_parameter, selected_scale, dim_a, dim_b, tuple(sorted(val_a_selected)))
                )

                selected_filter_types = [resolved_combo_type] if (val_a_selected and val_b_selected) else []
                if val_a_selected and val_b_selected:
                    selected_filter_values = [
                        make_combo_value(va, vb) if order_a_first else make_combo_value(vb, va)
                        for va in val_a_selected for vb in val_b_selected
                    ]
                    # hanya pertahankan pasangan yang benar-benar ada di data
                    selected_filter_values = [v for v in selected_filter_values if v in pair_values]
                else:
                    selected_filter_values = []

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
st.markdown('<div class="page-title">📊 Deka Insight Norm Dashboard</div>', unsafe_allow_html=True)

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
        st.markdown(metric_card("Norm Value", norm_display, tone=GRADE_TONE.get(row["norm_grade"])), unsafe_allow_html=True)
    with m2:
        st.markdown(metric_card("Base N", base_n_display), unsafe_allow_html=True)
else:
    # Base N berulang persis sama untuk tiap Metric pada kombinasi
    # (Norm Grade x Filter Type x Filter Value) yang sama, jadi sebelum
    # dijumlah, dedupe dulu biar tidak dihitung 4x (sekali per metric).
    unique_base_df = result_df.drop_duplicates(
        subset=["parameter_name", "scale_value", "norm_grade", "filter_type", "filter_value"]
    )
    total_base_n = unique_base_df["base_n"].sum(skipna=True)
    total_base_n_display = "-" if pd.isna(total_base_n) else f"{int(total_base_n)}"

    st.markdown(
        metric_card(f"Total Base N ({len(unique_base_df)} kombinasi unik)", total_base_n_display),
        unsafe_allow_html=True
    )

    grade_display_order = ["Top 25%", "Average 50%", "Bottom 25%"]
    metric_display_order = ["TB%", "T2B%", "T3B%", "MS"]
    grades_in_result = [g for g in grade_display_order if g in result_df["norm_grade"].unique()]
    metrics_in_result = [m for m in metric_display_order if m in result_df["metric"].unique()]

    pivot_rows = []
    chart_records = []
    for grade in grades_in_result:
        row = {"Norm Grade": grade}
        for met in metrics_in_result:
            cell_df = result_df[
                (result_df["norm_grade"] == grade) & (result_df["metric"] == met)
            ].dropna(subset=["norm_value"])

            weight_sum = cell_df["base_n"].sum(skipna=True)
            if cell_df.empty or pd.isna(weight_sum) or weight_sum == 0:
                row[met] = "-"
            else:
                weighted_avg = (cell_df["norm_value"] * cell_df["base_n"]).sum() / weight_sum
                row[met] = f"{weighted_avg:.2f}" if met == "MS" else f"{weighted_avg:.2f}%"
                chart_records.append({"Norm Grade": grade, "Metric": met, "Value": weighted_avg})
        pivot_rows.append(row)

    pivot_df = pd.DataFrame(pivot_rows)
    st.dataframe(pivot_df, use_container_width=True, hide_index=True)

    filter_types_in_result = sorted(result_df["filter_type"].dropna().unique())
    if len(filter_types_in_result) > 1:
        st.caption("⚠️ Total Base N bisa double-count karena lebih dari satu Filter Type dipilih sekaligus.")

    # ---- Charts: persen metric (TB/T2B/T3B, skala 0-100) dan MS (skala 1-9)
    # dipisah karena sumbu-y nya nggak sebanding kalau digabung.
    chart_df = pd.DataFrame(chart_records)
    if not chart_df.empty:
        pct_chart_df = chart_df[chart_df["Metric"] != "MS"]
        ms_chart_df = chart_df[chart_df["Metric"] == "MS"]

        chart_col1, chart_col2 = st.columns(2) if (not pct_chart_df.empty and not ms_chart_df.empty) else (st.container(), None)

        if not pct_chart_df.empty:
            with chart_col1:
                fig_pct = px.bar(
                    pct_chart_df, x="Norm Grade", y="Value", color="Metric",
                    barmode="group", text_auto=".1f",
                    category_orders={"Norm Grade": grades_in_result, "Metric": [m for m in metric_display_order if m != "MS"]},
                    labels={"Value": "Norm Value (%)"},
                    title="Norm Value per Norm Grade (TB% / T2B% / T3B%)"
                )
                fig_pct.update_layout(
                    plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
                    font_color="#1a2433", legend_title_text="Metric", yaxis_range=[0, 105]
                )
                st.plotly_chart(fig_pct, use_container_width=True)

        if not ms_chart_df.empty:
            target_col = chart_col2 if chart_col2 is not None else chart_col1
            with target_col:
                fig_ms = px.bar(
                    ms_chart_df, x="Norm Grade", y="Value", text_auto=".2f",
                    category_orders={"Norm Grade": grades_in_result},
                    labels={"Value": "Mean Score"},
                    title="Mean Score (MS) per Norm Grade",
                    color_discrete_sequence=["#3f8fd9"]
                )
                fig_ms.update_layout(
                    plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
                    font_color="#1a2433"
                )
                st.plotly_chart(fig_ms, use_container_width=True)


# =====================================
# POSISI SKOR PRODUK (apakah disukai / tidak vs norm)
# Independen dari pilihan Norm Grade di sidebar — selalu menarik
# Top 25% / Average 50% / Bottom 25% lengkap untuk jadi acuan posisi.
# =====================================
st.markdown('<div class="section-title">Bandingkan Skor Produk Anda</div>', unsafe_allow_html=True)

if not selected_filter_types or not selected_filter_values:
    st.info("Pilih minimal satu Filter Type dan Filter Value pada sidebar untuk membandingkan skor.")
else:
    base_filtered_df = df[
        (df["parameter_name"] == selected_parameter) &
        (df["scale_value"] == selected_scale) &
        (df["filter_type"].isin(selected_filter_types)) &
        (df["filter_value"].isin(selected_filter_values))
    ]

    pos_metric_options = ["TB%", "T2B%", "MS"]
    if int(selected_scale) >= 7:
        pos_metric_options.append("T3B%")

    pc1, pc2 = st.columns([1, 1])
    with pc1:
        pos_metric = st.selectbox("Metric acuan posisi", pos_metric_options, key="pos_metric")
    with pc2:
        score_max = float(selected_scale) if pos_metric == "MS" else 100.0
        score_step = 0.01
        score_default = round(score_max / 2, 2)
        input_score = st.number_input(
            f"Skor produk Anda ({'skala 1-' + selected_scale if pos_metric == 'MS' else '%'})",
            min_value=0.0, max_value=score_max, value=score_default, step=score_step
        )

    norms = weighted_grade_values(base_filtered_df, pos_metric)

    if all(v is None for v in norms.values()):
        st.warning("Tidak ada data norm untuk kombinasi filter & metric ini.")
    else:
        label, tone = classify_score(input_score, norms)
        st.markdown(positioning_badge(label, tone), unsafe_allow_html=True)

        top_v, avg_v, bottom_v = norms.get("Top 25%"), norms.get("Average 50%"), norms.get("Bottom 25%")

        axis_min = 0.0 if pos_metric != "MS" else 1.0
        axis_max = score_max
        zone_low = bottom_v if bottom_v is not None else axis_min
        zone_high = top_v if top_v is not None else axis_max

        unit = "" if pos_metric == "MS" else "%"
        fmt = lambda v: f"{v:.2f}{unit}" if v is not None else "-"

        # Ruler horizontal: 3 segmen warna (Bawah Norm / Sesuai Norm / Atas Norm)
        # + garis vertikal menandai posisi skor produk.
        fig_ruler = go.Figure()
        fig_ruler.add_trace(go.Bar(
            x=[max(zone_low - axis_min, 0.001)], y=["Posisi"], orientation="h",
            marker_color="#e3787a", text=["Bawah Norm"], textposition="inside",
            insidetextanchor="middle", textfont=dict(color="white", size=13),
            hoverinfo="skip", showlegend=False
        ))
        fig_ruler.add_trace(go.Bar(
            x=[max(zone_high - zone_low, 0.001)], y=["Posisi"], orientation="h",
            marker_color="#9fb8d6", text=["Sesuai Norm"], textposition="inside",
            insidetextanchor="middle", textfont=dict(color="#1a2433", size=13),
            hoverinfo="skip", showlegend=False
        ))
        fig_ruler.add_trace(go.Bar(
            x=[max(axis_max - zone_high, 0.001)], y=["Posisi"], orientation="h",
            marker_color="#5bbf8a", text=["Atas Norm"], textposition="inside",
            insidetextanchor="middle", textfont=dict(color="white", size=13),
            hoverinfo="skip", showlegend=False
        ))
        fig_ruler.update_layout(barmode="stack", height=130,
            margin=dict(t=10, b=30, l=10, r=10),
            plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
            font_color="#1a2433", showlegend=False,
            xaxis=dict(range=[axis_min, axis_max], showgrid=False, fixedrange=True),
            yaxis=dict(visible=False)
        )
        fig_ruler.add_shape(type="line", x0=input_score, x1=input_score, y0=-0.5, y1=0.5,
            line=dict(color="#1a2433", width=3, dash="solid"))
        fig_ruler.add_annotation(x=input_score, y=0.62, text=f"▼ Skor Anda: {fmt(input_score)}",
            showarrow=False, font=dict(size=13, color="#1a2433", family="Arial Black"))
        st.plotly_chart(fig_ruler, use_container_width=True)
        st.caption(f"Batas Bawah Norm: {fmt(bottom_v)} • Batas Atas Norm: {fmt(top_v)} — di antara keduanya = Sesuai Norm.")








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
