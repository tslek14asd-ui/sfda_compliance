
import io
import os
from datetime import date, datetime

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

st.set_page_config(
    page_title="لوحة مؤشرات أداء مكتبة الالتزام",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed",
)

BLUE = "#007DC5"
GREEN = "#00B258"
ORANGE = "#F15A31"
NAVY = "#013A5E"
BG = "#F4F7FA"
BORDER = "#DCE4EC"
TEXT = "#1B2B3A"
MUTED = "#6B7A8A"

st.markdown(f"""
<style>
html, body, [class*="css"] {{
    font-family: "Segoe UI", Tahoma, Arial, sans-serif;
}}
.stApp {{ background: {BG}; direction: rtl; }}
.block-container {{ padding: 1rem 2rem 2rem; max-width: 1500px; }}
.topbar {{
    background: linear-gradient(90deg, {NAVY}, {BLUE});
    color: white; padding: 14px 22px; border-radius: 10px;
    margin-bottom: 12px;
}}
.topbar h1 {{ font-size: 19px; margin: 0; }}
.topbar p {{ font-size: 11px; margin: 3px 0 0; opacity: .86; }}
.card {{
    background: white; border: 1px solid {BORDER}; border-radius: 10px;
    padding: 12px; box-shadow: 0 2px 6px rgba(1,58,94,.04);
}}
.card-title {{ color:{NAVY}; font-weight:700; font-size:13px; text-align:center; margin-bottom:4px; }}
.kpi-number {{ color:{NAVY}; font-size:27px; font-weight:800; text-align:center; }}
.kpi-label {{ color:{MUTED}; font-size:11px; text-align:center; }}
.small-note {{ color:{MUTED}; font-size:10px; text-align:center; }}
div[data-testid="stSelectbox"] label {{ font-size: 11px; color: {MUTED}; }}
button[kind="secondary"] {{ direction: rtl; }}
</style>
""", unsafe_allow_html=True)

# -------------------------
# Data loading
# -------------------------
@st.cache_data(ttl=300, show_spinner=False)
def load_excel(source):
    if source.startswith("http://") or source.startswith("https://"):
        import requests
        r = requests.get(source, timeout=30)
        r.raise_for_status()
        return pd.read_excel(io.BytesIO(r.content), sheet_name=None, engine="openpyxl")
    return pd.read_excel(source, sheet_name=None, engine="openpyxl")

def clean_sheet(df):
    df = df.copy()
    # Drop completely empty rows/columns.
    df = df.dropna(how="all").dropna(axis=1, how="all")
    # Strip column names.
    df.columns = [str(c).strip() for c in df.columns]
    return df

def first_data_row_sheet(df, expected_cols=None):
    """Find the first row that looks like the real header."""
    df0 = df.copy()
    if expected_cols:
        for i in range(min(len(df0), 30)):
            vals = [str(x).strip() for x in df0.iloc[i].tolist()]
            hits = sum(c in vals for c in expected_cols)
            if hits >= max(2, len(expected_cols)//3):
                out = df0.iloc[i+1:].copy()
                out.columns = vals
                return clean_sheet(out)
    return clean_sheet(df0)

def fmt_pct(x):
    return f"{x*100:.1f}%"

# Source can be supplied as a Streamlit secret:
# [data]
# excel_url = "https://..."
# Otherwise the bundled workbook is used for local testing.
try:
    excel_url = st.secrets["data"]["excel_url"]
except Exception:
    excel_url = ""

source = excel_url or os.path.join(os.path.dirname(__file__), "data", "compliance.xlsx")

try:
    sheets = load_excel(source)
except Exception as e:
    st.error("تعذر قراءة ملف Excel. تأكدي من رابط الملف وصلاحية الوصول إليه.")
    st.exception(e)
    st.stop()

library_raw = sheets.get("مكتبة الالتزام")
if library_raw is None:
    st.error("لم يتم العثور على Sheet باسم «مكتبة الالتزام».")
    st.stop()

library = clean_sheet(library_raw)

# Remove title/blank rows and locate the real header if necessary.
required = ["التسلسل", "فئة الوثيقة", "اسم الوثيقة", "الجهة المصدرة", "نفاذ الوثيقة", "قابلية الفحص"]
if "التسلسل" not in library.columns:
    library = first_data_row_sheet(library_raw, required)

# Keep rows with a document sequence/name; protect against formula error rows.
if "التسلسل" in library.columns:
    seq = pd.to_numeric(library["التسلسل"], errors="coerce")
    library = library[seq.notna()].copy()
    library["التسلسل"] = seq.loc[library.index].astype(int)

for c in library.columns:
    if library[c].dtype == "object":
        library[c] = library[c].astype(str).str.strip().replace({"nan": ""})

# -------------------------
# Header
# -------------------------
st.markdown("""
<div class="topbar">
  <h1>لوحة مؤشرات أداء مكتبة الالتزام</h1>
  <p>الهيئة العامة للغذاء والدواء | Saudi Food & Drug Authority</p>
</div>
""", unsafe_allow_html=True)

c1, c2 = st.columns([1, 5])
with c1:
    if st.button("↻ تحديث البيانات", use_container_width=True):
        load_excel.clear()
        st.rerun()
with c2:
    st.markdown(
        f"<div style='text-align:left;color:{MUTED};font-size:11px;padding-top:9px'>"
        f"آخر قراءة للبيانات: {datetime.now():%Y-%m-%d %H:%M}</div>",
        unsafe_allow_html=True,
    )

# -------------------------
# Filters
# -------------------------
filter_cols = [
    ("الجهة المصدرة", "الجهة المصدرة"),
    ("فئة الوثيقة", "فئة الوثيقة"),
    ("الأداة التنظيمية", "أداة تنظيمية"),
    ("نفاذ الوثيقة", "نفاذ الوثيقة"),
    ("قابلية الفحص", "قابلية الفحص"),
    ("مصدر الوثيقة", "مصدر الوثيقة"),
    ("نطاق التطبيق", "نطاق التطبيق"),
]

filter_values = {}
cols = st.columns(7)
for col_ui, (label, col_name) in zip(cols, filter_cols):
    with col_ui:
        if col_name in library.columns:
            vals = sorted([v for v in library[col_name].dropna().unique().tolist() if str(v).strip()])
        else:
            vals = []
        filter_values[col_name] = st.selectbox(label, ["(الكل)"] + vals, key=f"f_{col_name}")

filtered = library.copy()
for col_name, selected in filter_values.items():
    if selected != "(الكل)" and col_name in filtered.columns:
        filtered = filtered[filtered[col_name] == selected]

# -------------------------
# KPIs
# -------------------------
total = len(filtered)
valid = int((filtered.get("نفاذ الوثيقة", pd.Series(dtype=str)) == "ساري").sum()) if total else 0
inspect = int((filtered.get("قابلية الفحص", pd.Series(dtype=str)) == "قابل للفحص").sum()) if total else 0
apply = int((filtered.get("نطاق التطبيق", pd.Series(dtype=str)) == "ينطبق").sum()) if total else 0

def gauge(title, value, suffix, color, denominator=100):
    pct = 0 if denominator == 0 else max(0, min(100, value / denominator * 100))
    fig = go.Figure(go.Pie(
        values=[pct, 100-pct], hole=.74, rotation=90,
        marker=dict(colors=[color, "#E4E9F0"], line=dict(width=0)),
        textinfo="none", hoverinfo="skip",
    ))
    fig.update_layout(
        height=190, margin=dict(l=5,r=5,t=5,b=0),
        showlegend=False, paper_bgcolor="rgba(0,0,0,0)",
        annotations=[dict(text=f"<b>{value}{suffix}</b><br><span style='font-size:11px;color:{MUTED}'>{title}</span>",
                          x=.5,y=.5,showarrow=False,font=dict(size=21,color=color))]
    )
    return fig

kpis = [
    ("إجمالي عدد الوثائق", total, "", NAVY, max(total, 1)),
    ("نسبة السريان", round(valid/total*100,1) if total else 0, "%", GREEN, 100),
    ("نسبة القابلية للفحص", round(inspect/total*100,1) if total else 0, "%", BLUE, 100),
    ("نسبة الانطباق", round(apply/total*100,1) if total else 0, "%", ORANGE, 100),
]
kc = st.columns(4)
for ui, (title, value, suffix, color, denom) in zip(kc, kpis):
    with ui:
        st.markdown(f"<div class='card'><div class='card-title'>{title}</div></div>", unsafe_allow_html=True)
        st.plotly_chart(gauge(title, value, suffix, color, denom), use_container_width=True, config={"displayModeBar": False})

# -------------------------
# Bar charts
# -------------------------
def bar(title, column, height=280):
    if column not in filtered.columns:
        st.info(f"لا يوجد عمود «{column}» في البيانات.")
        return
    s = filtered[column].replace("", pd.NA).dropna().value_counts().head(10).sort_values()
    fig = go.Figure(go.Bar(
        x=s.values, y=s.index, orientation="h",
        marker=dict(color=BLUE), hovertemplate="%{y}: %{x}<extra></extra>"
    ))
    fig.update_layout(
        title=dict(text=title, x=.5, font=dict(size=13, color=NAVY)),
        height=height, margin=dict(l=10,r=10,t=40,b=10),
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        xaxis=dict(showgrid=False, showticklabels=False),
        yaxis=dict(showgrid=False, tickfont=dict(size=10)),
    )
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})

r1, r2 = st.columns(2)
with r1:
    bar("فئة الوثيقة", "فئة الوثيقة")
with r2:
    bar("الجهة المصدرة", "الجهة المصدرة")

r3, r4 = st.columns(2)
with r3:
    bar("مصدر الوثيقة", "مصدر الوثيقة")
with r4:
    bar("الأداة التنظيمية", "أداة تنظيمية")

# -------------------------
# Countdown: days decrease, completion increases
# -------------------------
st.markdown(f"<h3 style='color:{NAVY};font-size:14px;text-align:center;margin:15px 0 5px'>العدادات التنازلية</h3>", unsafe_allow_html=True)

countdown_raw = sheets.get("عداد_الأيام")
if countdown_raw is not None:
    cd = clean_sheet(countdown_raw)
    # Prefer the actual dates and total duration. Excel serials are handled by pandas.
    if "تاريخ النفاذ" in cd.columns:
        dates = pd.to_datetime(cd["تاريخ النفاذ"], errors="coerce", origin="1899-12-30", unit="D")
    else:
        dates = pd.Series(pd.NaT, index=cd.index)

    if "المدة الكلية" in cd.columns:
        durations = pd.to_numeric(cd["المدة الكلية"], errors="coerce")
    else:
        durations = pd.Series([180]*len(cd), index=cd.index)

    names = cd.get("اسم الوثيقة", pd.Series("", index=cd.index)).fillna("").astype(str)
    countdown_rows = []
    today = pd.Timestamp(date.today())

    for idx in cd.index:
        if pd.isna(dates.loc[idx]) or not names.loc[idx].strip():
            continue
        total_days = int(durations.loc[idx]) if pd.notna(durations.loc[idx]) else 180
        remaining = max(0, (dates.loc[idx].normalize() - today).days)
        elapsed = max(0, total_days - remaining)
        completion = min(100, max(0, elapsed / total_days * 100)) if total_days else 100
        countdown_rows.append((names.loc[idx], remaining, completion, dates.loc[idx]))

    cc = st.columns(max(1, min(3, len(countdown_rows))))
    for ui, (name, remaining, completion, end_date) in zip(cc, countdown_rows):
        with ui:
            fig = go.Figure(go.Pie(
                values=[completion, 100-completion], hole=.76, rotation=90,
                marker=dict(colors=[BLUE, "#E4E9F0"], line=dict(width=0)),
                textinfo="none", hoverinfo="skip"
            ))
            fig.update_layout(
                height=220, margin=dict(l=5,r=5,t=5,b=5),
                showlegend=False, paper_bgcolor="rgba(0,0,0,0)",
                annotations=[dict(
                    text=f"<b>{remaining}</b><br><span style='font-size:11px'>يوم متبقي</span><br>"
                         f"<span style='font-size:10px;color:{MUTED}'>{completion:.1f}% مكتمل</span>",
                    x=.5,y=.5,showarrow=False,font=dict(size=20,color=NAVY)
                )]
            )
            st.markdown(f"<div class='card'><div class='card-title'>{name}</div></div>", unsafe_allow_html=True)
            st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})
            st.markdown(f"<div class='small-note'>تاريخ النفاذ: {end_date:%Y-%m-%d}</div>", unsafe_allow_html=True)
else:
    st.info("لم يتم العثور على Sheet «عداد_الأيام».")

st.markdown(
    f"<div class='small-note' style='padding:15px'>"
    f"البيانات المعروضة مأخوذة مباشرة من ملف Excel، والفلاتر والرسوم والمؤشرات تتغير وفق البيانات المفلترة."
    f"</div>",
    unsafe_allow_html=True
)
