import streamlit as st
import pandas as pd
import numpy as np
import openpyxl
import re
from pathlib import Path

# =========================
# KONFIGURASI DASHBOARD
# =========================
st.set_page_config(
    page_title="Norm Database Dashboard",
    page_icon="📊",
    layout="wide"
)

RAW_DATA_PATH = Path(__file__).parent / "Deka_Insight_Data_For_Norm_Database.xlsx"

FILTER_COLS = [
    "Category", "Sub-Category", "Detail Product", "Gender",
    "Actual Age", "SES", "Occupation", "Type of Study",
    "Test Type", "Methodology", "Sub-Method", "# of Product", "Sequence"
]

# =========================
# LOAD & PARSE RAW DATA
# =========================
@st.cache_data(show_spinner="Memuat raw data dari Excel...")
def load_raw_data(path: str) -> pd.DataFrame:
    """
    Load semua sheet dari Excel, normalisasi nama parameter
    (pisahkan nama parameter dari skala pts), lalu gabungkan jadi
    satu DataFrame panjang (long-format) agar mudah di-query.
    """
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    all_dfs = []

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        rows = list(ws.iter_rows(values_only=True))

        # Baris pertama kadang kosong, baris ke-2 = header
        if len(rows) < 2:
            continue

        # Cari baris header (baris yang mulai dengan 'SbjNum')
        header_idx = None
        for i, row in enumerate(rows):
            if row and str(row[0]).strip() == "SbjNum":
                header_idx = i
                break
        if header_idx is None:
            continue

        header = rows[header_idx]
        data_rows = rows[header_idx + 1:]

        if not data_rows:
            continue

        df = pd.DataFrame(data_rows, columns=header)

        # Hanya ambil kolom yang tidak None di header
        valid_cols = [c for c in df.columns if c is not None]
        df = df[valid_cols]

        # Identifikasi kolom parameter (kolom P dst = setelah 'Sequence')
        demo_cols = ["SbjNum", "No Project", "Category", "Sub-Category",
                     "Detail Product", "Gender", "Actual Age", "SES",
                     "Occupation", "Type of Study", "Test Type",
                     "Methodology", "Sub-Method", "# of Product", "Sequence"]

        param_cols = [c for c in valid_cols if c not in demo_cols]

        # Melt ke long format
        id_cols = [c for c in demo_cols if c in valid_cols]
        df_long = df.melt(id_vars=id_cols, value_vars=param_cols,
                          var_name="raw_param", value_name="score")

        # Parse nama parameter dan skala dari string kolom, misal "Overall Liking - 7pts"
        def parse_param(s):
            s = str(s).strip()
            # Cari pola "- Npts" atau "- N pts" di akhir
            m = re.search(r'-\s*(\d+)\s*pts?\s*$', s, re.IGNORECASE)
            if m:
                scale = int(m.group(1))
                name = s[:m.start()].strip().strip('-').strip()
            else:
                # Coba tanpa unit (misal "Purchase Intent w/o Price" tanpa pts)
                scale = None
                name = s
            return name, scale

        df_long[["parameter_name", "scale_value"]] = df_long["raw_param"].apply(
            lambda x: pd.Series(parse_param(x))
        )

        # Drop baris di mana score bukan numerik
        df_long["score"] = pd.to_numeric(df_long["score"], errors="coerce")
        df_long = df_long.dropna(subset=["score"])
        df_long = df_long[df_long["score"] > 0]

        # Drop baris di mana scale_value None
        df_long = df_long.dropna(subset=["scale_value"])
        df_long["scale_value"] = df_long["scale_value"].astype(int)

        df_long["sheet"] = sheet_name
        all_dfs.append(df_long)

    if not all_dfs:
        return pd.DataFrame()

    combined = pd.concat(all_dfs, ignore_index=True)

    # Normalisasi nama parameter (strip whitespace, title case)
    combined["parameter_name"] = combined["parameter_name"].str.strip()

    # Pastikan filter cols ada
    for col in FILTER_COLS:
        if col not in combined.columns:
            combined[col] = None

    return combined


# =========================
# FUNGSI HITUNG NORM
# =========================
def compute_norm_metrics(scores: pd.Series, scale: int) -> dict:
    """Hitung TB%, T2B%, T3B%, dan Mean Score dari series skor."""
    n = len(scores)
    if n == 0:
        return {"TB%": None, "T2B%": None, "T3B%": None, "MS": None, "N": 0}

    tb  = (scores == scale).sum() / n * 100
    t2b = (scores >= scale - 1).sum() / n * 100
    t3b = (scores >= scale - 2).sum() / n * 100 if scale >= 7 else None
    ms  = scores.mean()

    return {"TB%": round(tb, 2), "T2B%": round(t2b, 2),
            "T3B%": round(t3b, 2) if t3b is not None else None,
            "MS": round(ms, 2), "N": n}


def compute_norm_by_grade(scores: pd.Series, scale: int) -> dict:
    """
    Hitung norm untuk Top 25%, Average 50%, Bottom 25%
    berdasarkan rank percentile individual (bukan nilai cutoff),
    sehingga selalu ada data di setiap segmen.
    """
    if len(scores) == 0:
        empty = {"TB%": None, "T2B%": None, "T3B%": None, "MS": None, "N": 0}
        return {"Top 25%": empty, "Average 50%": empty, "Bottom 25%": empty}

    n = len(scores)
    # Rank setiap baris (1 = terendah), method='first' hindari tie issue
    ranks = scores.rank(method="first")

    # Bagi berdasarkan rank percentile
    bottom25 = scores[ranks <= n * 0.25]
    top25    = scores[ranks > n * 0.75]
    avg50    = scores[(ranks > n * 0.25) & (ranks <= n * 0.75)]

    return {
        "Top 25%":     compute_norm_metrics(top25, scale),
        "Average 50%": compute_norm_metrics(avg50, scale),
        "Bottom 25%":  compute_norm_metrics(bottom25, scale),
    }


# =========================
# LOAD DATA
# =========================
st.title("📊 Norm Database Dashboard")
st.caption("Real-time norm benchmark dari raw data survey Deka Insight.")

if not Path(RAW_DATA_PATH).exists():
    st.error(f"File tidak ditemukan: **{RAW_DATA_PATH}**\n\n"
             "Pastikan file `Deka_Insight_Data_For_Norm_Database.xlsx` "
             "ada di direktori yang sama dengan `app.py`.")
    st.stop()

df = load_raw_data(RAW_DATA_PATH)

if df.empty:
    st.error("Gagal memuat data atau data kosong.")
    st.stop()

# =========================
# SIDEBAR FILTER
# =========================
st.sidebar.header("🔧 Filter Benchmark")

# Parameter
param_options = sorted(df["parameter_name"].dropna().unique())
selected_param = st.sidebar.selectbox(
    "Parameter",
    param_options,
    index=next((i for i, p in enumerate(param_options)
                if "overall liking" in p.lower()), 0)
)

# Scale
scale_options = sorted(
    df[df["parameter_name"] == selected_param]["scale_value"].dropna().unique()
)
selected_scale = st.sidebar.selectbox(
    "Skala (pts)",
    scale_options,
    format_func=lambda x: f"{x} pts"
)

# Metric
metric_options = ["TB%", "T2B%", "T3B%", "MS"]
if selected_scale < 7:
    metric_options = ["TB%", "T2B%", "MS"]
selected_metric = st.sidebar.selectbox("Metric", metric_options)

st.sidebar.divider()
st.sidebar.subheader("🗂️ Filter Dimensi (opsional)")

# Filter dimensi bebas
active_filters = {}
filtered_df = df[
    (df["parameter_name"] == selected_param) &
    (df["scale_value"] == selected_scale)
].copy()

for col in FILTER_COLS:
    if col not in filtered_df.columns:
        continue
    unique_vals = sorted(
        filtered_df[col].dropna().astype(str).unique()
    )
    if not unique_vals:
        continue
    selected_vals = st.sidebar.multiselect(
        col,
        options=["(Semua)"] + unique_vals,
        default=["(Semua)"]
    )
    if selected_vals and "(Semua)" not in selected_vals:
        active_filters[col] = selected_vals

# Terapkan semua filter dimensi
for col, vals in active_filters.items():
    filtered_df = filtered_df[filtered_df[col].astype(str).isin(vals)]

# =========================
# HITUNG NORM ON-THE-FLY
# =========================
norm_results = compute_norm_by_grade(filtered_df["score"], selected_scale)

# =========================
# TAMPILAN UTAMA
# =========================
st.subheader(f"📌 {selected_param} — {selected_scale} pts — {selected_metric}")

# Jumlah filter aktif
if active_filters:
    filter_summary = ", ".join(
        f"**{k}**: {', '.join(v)}" for k, v in active_filters.items()
    )
    st.markdown(f"🔍 Filter aktif: {filter_summary}")

# Total N
total_n = len(filtered_df)
st.metric("Total Responden (N)", f"{total_n:,}")

if total_n == 0:
    st.warning("Tidak ada data untuk kombinasi filter ini.")
    st.stop()

st.divider()

# Cards untuk 3 Norm Grade
col1, col2, col3 = st.columns(3)

def fmt_val(val, metric):
    if val is None:
        return "N/A"
    if metric == "MS":
        return f"{val:.2f}"
    return f"{val:.1f}%"

def render_grade_card(col, grade_label, grade_data, metric, color):
    val = grade_data.get(metric)
    n   = grade_data.get("N", 0)
    with col:
        st.markdown(
            f"""
            <div style="
                background: {color};
                border-radius: 12px;
                padding: 20px 16px;
                text-align: center;
                color: white;
                margin-bottom: 8px;
            ">
                <div style="font-size:13px; opacity:0.85; margin-bottom:4px;">{grade_label}</div>
                <div style="font-size:36px; font-weight:700;">{fmt_val(val, metric)}</div>
                <div style="font-size:12px; opacity:0.75; margin-top:4px;">N = {n:,}</div>
            </div>
            """,
            unsafe_allow_html=True
        )

render_grade_card(col1, "🏆 Top 25% (Above Average)",
                  norm_results["Top 25%"], selected_metric, "#1a7a4a")
render_grade_card(col2, "📊 Average 50%",
                  norm_results["Average 50%"], selected_metric, "#2563eb")
render_grade_card(col3, "📉 Bottom 25% (Below Average)",
                  norm_results["Bottom 25%"], selected_metric, "#b91c1c")

st.divider()

# =========================
# TABEL PERBANDINGAN SEMUA METRIC
# =========================
st.subheader("📋 Perbandingan Semua Metric")

rows_table = []
for grade in ["Top 25%", "Average 50%", "Bottom 25%"]:
    d = norm_results[grade]
    rows_table.append({
        "Norm Grade": grade,
        "N": d["N"],
        "TB%": fmt_val(d["TB%"], "TB%"),
        "T2B%": fmt_val(d["T2B%"], "T2B%"),
        "T3B%": fmt_val(d["T3B%"], "T3B%") if selected_scale >= 7 else "—",
        "Mean Score": fmt_val(d["MS"], "MS"),
    })

st.dataframe(pd.DataFrame(rows_table), use_container_width=True, hide_index=True)

st.divider()

# =========================
# DISTRIBUSI SKOR
# =========================
st.subheader("📊 Distribusi Skor Keseluruhan")

score_counts = (
    filtered_df["score"]
    .value_counts()
    .sort_index()
    .reset_index()
    .rename(columns={"score": "Skor", "count": "Frekuensi"})
)
score_counts["Skor"] = score_counts["Skor"].astype(int).astype(str)

st.bar_chart(score_counts.set_index("Skor")["Frekuensi"])

# =========================
# EKSPLORASI DATA MENTAH
# =========================
with st.expander("🔍 Lihat Data Mentah (sample 200 baris)"):
    show_cols = ["sheet", "parameter_name", "scale_value", "score"] + [
        c for c in FILTER_COLS if c in filtered_df.columns
    ]
    st.dataframe(
        filtered_df[show_cols].head(200),
        use_container_width=True
    )

# =========================
# DOWNLOAD
# =========================
st.divider()
csv_export = filtered_df.to_csv(index=False).encode("utf-8")
st.download_button(
    label="⬇️ Download data terfilter sebagai CSV",
    data=csv_export,
    file_name=f"norm_{selected_param}_{selected_scale}pts.csv",
    mime="text/csv"
)
