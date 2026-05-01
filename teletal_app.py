import io
from datetime import datetime

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
import streamlit as st

from teletal_core import discover_menu_sections, discover_week_options, scrape_menu_rows


def apply_theme():
    st.markdown(
        """
        <style>
        .block-container {
            max-width: 1520px;
            padding-top: 1.35rem;
            padding-bottom: 2.5rem;
        }
        [data-testid="stSidebar"] {
            background: #f6f8fb;
            border-right: 1px solid #e5e7eb;
        }
        [data-testid="stSidebar"] [data-testid="stVerticalBlock"] {
            gap: 0.75rem;
        }
        h1 {
            font-size: 1.65rem !important;
            line-height: 1.2 !important;
            margin-bottom: 0.1rem !important;
        }
        .app-subtitle {
            color: #64748b;
            font-size: 0.92rem;
            margin: 0 0 1rem 0;
        }
        div[data-testid="stMetric"] {
            background: #ffffff;
            border: 1px solid #e5e7eb;
            border-radius: 8px;
            padding: 0.8rem 0.9rem;
            box-shadow: 0 1px 2px rgba(15, 23, 42, 0.04);
        }
        div[data-testid="stMetric"] label {
            color: #64748b !important;
            font-size: 0.78rem !important;
        }
        div[data-testid="stMetricValue"] {
            font-size: 1.25rem !important;
            color: #0f172a;
        }
        div[data-testid="stExpander"] {
            border: 1px solid #dce3eb;
            border-radius: 8px;
            background: #ffffff;
            box-shadow: 0 1px 2px rgba(15, 23, 42, 0.035);
            margin-bottom: 0.55rem;
        }
        div[data-testid="stExpander"] summary {
            min-height: 42px;
            font-weight: 650;
            color: #182230;
        }
        div[data-testid="stDataFrame"] {
            border: 1px solid #e5e7eb;
            border-radius: 8px;
            overflow: hidden;
        }
        .section-meta {
            display: flex;
            flex-wrap: wrap;
            gap: 0.45rem;
            margin: 0.1rem 0 0.75rem 0;
        }
        .section-pill {
            border: 1px solid #dbe3ec;
            background: #f8fafc;
            border-radius: 999px;
            color: #334155;
            font-size: 0.78rem;
            padding: 0.22rem 0.55rem;
        }
        .stDownloadButton button {
            border-radius: 7px;
            font-weight: 650;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def build_xlsx(rows):
    fieldnames = list(rows[0].keys())
    numeric_cols = {
        "kcal_adag", "kj_adag", "zsír_g_adag", "zsír_telített_g_adag",
        "szénhidrát_g_adag", "cukor_g_adag", "rost_g_adag", "fehérje_g_adag",
        "só_g_adag", "kcal_100g", "kj_100g", "zsír_g_100g",
        "szénhidrát_g_100g", "fehérje_g_100g", "súly_g", "nap_szám", "hét", "év",
        "ár_ft",
    }
    wb = Workbook()
    ws = wb.active
    ev = rows[0].get("év", "")
    het = rows[0].get("hét", "")
    ws.title = f"{ev} {het}. hét"

    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", fgColor="2E75B6")

    for col_idx, name in enumerate(fieldnames, 1):
        cell = ws.cell(row=1, column=col_idx, value=name)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center")

    for row_idx, row in enumerate(rows, 2):
        for col_idx, key in enumerate(fieldnames, 1):
            val = row[key]
            if key in numeric_cols and val != "":
                try:
                    val = float(val) if "." in str(val) else int(val)
                except ValueError:
                    pass
            ws.cell(row=row_idx, column=col_idx, value=val)

    for col_idx, name in enumerate(fieldnames, 1):
        max_len = max(len(str(name)), *(len(str(r[name])) for r in rows))
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 2, 40)

    ws.freeze_panes = "A2"

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


@st.cache_data(ttl=24 * 60 * 60, show_spinner=False)
def cached_menu_rows(week):
    rows, pages = scrape_menu_rows(
        week=week,
        delay_seconds=0,
    )
    return rows, pages, datetime.now().strftime("%Y-%m-%d %H:%M")


@st.cache_data(ttl=24 * 60 * 60)
def week_options():
    return discover_week_options()


@st.cache_data(ttl=24 * 60 * 60)
def menu_options(week):
    sections, _ = discover_menu_sections(week)
    return sections


def dataframe_from_rows(rows):
    df = pd.DataFrame(rows)
    numeric_cols = [
        "hét", "év", "nap_szám", "ár_ft", "súly_g", "kcal_adag", "kj_adag",
        "zsír_g_adag", "zsír_telített_g_adag", "szénhidrát_g_adag",
        "cukor_g_adag", "rost_g_adag", "fehérje_g_adag", "só_g_adag",
        "kcal_100g", "kj_100g", "zsír_g_100g", "szénhidrát_g_100g",
        "fehérje_g_100g",
    ]
    for col in numeric_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")
    return df


def table_height(row_count):
    return min(540, max(180, 78 + 34 * row_count))


def table_columns(df):
    visible_columns = [
        "nap", "kod", "sor_név", "név", "ár_ft", "súly_g", "kcal_adag",
        "fehérje_g_adag", "szénhidrát_g_adag", "zsír_g_adag", "allergének",
    ]
    return [col for col in visible_columns if col in df.columns]


def table_config():
    return {
        "nap": st.column_config.TextColumn("Nap", width="small"),
        "kod": st.column_config.TextColumn("Kód", width="small"),
        "sor_név": st.column_config.TextColumn("Sor", width="medium"),
        "név": st.column_config.TextColumn("Étel", width="large"),
        "ár_ft": st.column_config.NumberColumn("Ár", format="%d Ft", width="small"),
        "súly_g": st.column_config.NumberColumn("Súly", format="%d g", width="small"),
        "kcal_adag": st.column_config.NumberColumn("kcal", format="%.0f", width="small"),
        "fehérje_g_adag": st.column_config.NumberColumn("Fehérje", format="%.1f g", width="small"),
        "szénhidrát_g_adag": st.column_config.NumberColumn("CH", format="%.1f g", width="small"),
        "zsír_g_adag": st.column_config.NumberColumn("Zsír", format="%.1f g", width="small"),
        "allergének": st.column_config.TextColumn("Allergének", width="large"),
    }


def render_menu_table(section_df):
    st.dataframe(
        section_df[table_columns(section_df)],
        use_container_width=True,
        hide_index=True,
        height=table_height(len(section_df)),
        column_config=table_config(),
    )


def format_int(value):
    return f"{int(round(value)):,}".replace(",", " ")


def section_stats(section_df):
    price_min = section_df["ár_ft"].min()
    price_max = section_df["ár_ft"].max()
    kcal_min = section_df["kcal_adag"].min()
    kcal_max = section_df["kcal_adag"].max()

    if pd.isna(price_min) or pd.isna(price_max):
        price = "-"
    else:
        price = f"{format_int(price_min)}-{format_int(price_max)} Ft" if price_min != price_max else f"{format_int(price_min)} Ft"

    if pd.isna(kcal_min) or pd.isna(kcal_max):
        kcal = "-"
    else:
        kcal = f"{format_int(kcal_min)}-{format_int(kcal_max)} kcal" if kcal_min != kcal_max else f"{format_int(kcal_min)} kcal"
    return price, kcal


def render_section_meta(section_df):
    price, kcal = section_stats(section_df)
    st.markdown(
        f"""
        <div class="section-meta">
            <span class="section-pill">{len(section_df)} tétel</span>
            <span class="section-pill">{section_df["kod"].nunique()} kód</span>
            <span class="section-pill">{section_df["nap"].nunique()} nap</span>
            <span class="section-pill">{price}</span>
            <span class="section-pill">{kcal}</span>
        </div>
        """,
        unsafe_allow_html=True,
    )


# ---------------------------------------------------------------------------
# UI
# ---------------------------------------------------------------------------

st.set_page_config(page_title="Teletál étlap", page_icon="🍽️", layout="wide")
apply_theme()

try:
    weeks, active_week = week_options()
except Exception:
    weeks, active_week = [], None

with st.sidebar:
    if st.button("Frissítés most", use_container_width=True):
        cached_menu_rows.clear()
        menu_options.clear()
        week_options.clear()
        st.rerun()

if weeks:
    week_values = [option["week"] for option in weeks]
    week_labels = {option["week"]: option["label"] for option in weeks}
    week = st.sidebar.selectbox(
        "Hét",
        week_values,
        index=week_values.index(active_week) if active_week in week_values else 0,
        format_func=lambda value: week_labels.get(value, f"{value}. hét"),
    )
else:
    week = None

if not week:
    st.error("Nem sikerült betölteni a választható heteket.")
    st.stop()

try:
    with st.spinner("Adatok betöltése..."):
        rows, pages, generated_at = cached_menu_rows(week)
        sections = menu_options(week)
except Exception as e:
    st.error(f"Hiba történt: {e}")
    st.stop()

if not rows:
    st.warning("Nem található letölthető étel.")
    st.stop()

df = dataframe_from_rows(rows)
ev = pages[0]["ev"] if pages else ""
het = pages[0]["het"] if pages else week

with st.sidebar:
    st.caption(f"Utolsó frissítés: {generated_at}")
    st.caption(f"{len(sections)} menü, {len(df)} sor/napi tétel")

    day_values = list(df["nap"].dropna().unique())
    selected_days = st.multiselect("Nap szűrő", day_values, default=day_values)

    search = st.text_input("Keresés")

filtered = df[df["nap"].isin(selected_days)]
if search:
    search_mask = filtered.astype(str).apply(
        lambda col: col.str.contains(search, case=False, na=False)
    ).any(axis=1)
    filtered = filtered[search_mask]

week_label = week_labels.get(het, f"{ev}/{het}") if weeks else f"{ev}/{het}"

export_rows = filtered.to_dict("records")
full_export_rows = df.to_dict("records")
filtered_xlsx = build_xlsx(export_rows) if export_rows else None
full_xlsx = build_xlsx(full_export_rows)

with st.sidebar:
    st.download_button(
        label="Excel letöltés",
        data=filtered_xlsx if filtered_xlsx else full_xlsx,
        file_name=f"teletal_{ev}_{het}het.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        type="primary",
    )

title_col, action_col = st.columns([0.72, 0.28], vertical_alignment="center")
with title_col:
    st.title("Teletál étlap")
    st.markdown(
        f'<p class="app-subtitle">{week_label} · frissítve: {generated_at}</p>',
        unsafe_allow_html=True,
    )

with action_col:
    st.download_button(
        label="Excel letöltés",
        data=filtered_xlsx if filtered_xlsx else full_xlsx,
        file_name=f"teletal_{ev}_{het}het.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        type="primary",
    )

col1, col2, col3, col4 = st.columns(4)
col1.metric("Menük", filtered["menü"].nunique())
col2.metric("Tételek", len(filtered))
col3.metric("Kódok", filtered["kod"].nunique())
col4.metric("Hét", week_label)

st.divider()

section_keys = set()
for section in sections:
    title = section["title"]
    section_name = section["section_name"]
    section_df = filtered[
        (filtered["menü"] == title)
        & (filtered["menü_azonosító"] == section_name)
    ].sort_values(["nap_szám", "kod", "sor_név", "név"])

    if search and section_df.empty:
        continue

    section_keys.add((title, section_name))
    section_label = title if title == section_name else f"{title} [{section_name}]"
    expander_label = f"{section_label}    {len(section_df)} tétel    {section_df['kod'].nunique()} kód"

    with st.expander(expander_label, expanded=False):
        if section_df.empty:
            st.info("Nincs találat a jelenlegi szűrőkkel.")
            continue
        render_section_meta(section_df)
        render_menu_table(section_df)

remaining = filtered[
    ~filtered.apply(lambda row: (row["menü"], row["menü_azonosító"]) in section_keys, axis=1)
]
if not remaining.empty:
    with st.expander(f"Egyéb · {len(remaining)} tétel · {remaining['kod'].nunique()} kód", expanded=False):
        render_section_meta(remaining)
        render_menu_table(remaining)

if filtered.empty:
    st.info("Nincs találat a jelenlegi szűrőkkel.")
