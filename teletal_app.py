import io
import json
import re
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
        html, body, [data-testid="stAppViewContainer"], .stApp {
            background: #eef3f7 !important;
            color: #111827 !important;
        }
        [data-testid="stHeader"] {
            background: rgba(238, 243, 247, 0.92) !important;
            border-bottom: 1px solid #dbe3ec;
        }
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
            color: #102a43 !important;
        }
        h2, h3 {
            color: #102a43 !important;
        }
        .app-subtitle {
            color: #64748b;
            font-size: 0.92rem;
            margin: 0 0 1rem 0;
        }
        .planner-note {
            color: #64748b;
            font-size: 0.88rem;
            margin-top: -0.4rem;
            margin-bottom: 0.8rem;
        }
        .planner-title {
            font-size: 1.05rem;
            font-weight: 700;
            color: #102a43;
            margin: 0.4rem 0 0.25rem 0;
        }
        .planner-grid-title {
            color: #334155;
            font-size: 0.92rem;
            font-weight: 700;
            margin: 0 0 0.45rem 0;
        }
        .planner-source-note {
            color: #64748b;
            font-size: 0.82rem;
            margin: 0.1rem 0 0.75rem 0;
        }
        .day-summary {
            align-items: center;
            background: #ffffff;
            border: 1px solid #dbe3ec;
            border-radius: 8px;
            display: grid;
            gap: 0.75rem;
            grid-template-columns: minmax(90px, 0.75fr) minmax(140px, 1fr) minmax(120px, 0.8fr) minmax(260px, 1.5fr);
            margin: 1rem 0 0.55rem 0;
            padding: 0.85rem 1rem;
        }
        .day-summary-name {
            color: #102a43;
            font-size: 1rem;
            font-weight: 800;
        }
        .day-summary-kcal {
            color: #0f766e;
            font-size: 1.55rem;
            font-weight: 850;
            line-height: 1.05;
        }
        .day-summary-price {
            color: #102a43;
            font-size: 1.15rem;
            font-weight: 800;
        }
        .day-summary-meta {
            color: #475569;
            font-size: 0.88rem;
            line-height: 1.5;
        }
        [data-testid="stSelectbox"] label,
        [data-testid="stMultiSelect"] label,
        [data-testid="stTextInput"] label {
            color: #334155 !important;
            font-weight: 650;
        }
        [data-baseweb="select"] > div,
        [data-testid="stTextInput"] input {
            background: #ffffff !important;
            border-color: #cbd5e1 !important;
            color: #0f172a !important;
            border-radius: 7px !important;
        }
        [data-testid="stAlert"] {
            background: #f8fafc;
            border: 1px solid #dbe3ec;
            color: #334155;
            border-radius: 8px;
        }
        hr {
            border-color: #dbe3ec;
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
            background: #ffffff;
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
        .stButton button {
            border-radius: 7px;
            border: 1px solid #0f766e;
            background: #0f766e;
            color: #ffffff;
            font-weight: 650;
        }
        .stButton button:hover {
            border-color: #115e59;
            background: #115e59;
            color: #ffffff;
        }
        .stDownloadButton button[kind="primary"],
        .stDownloadButton button {
            background: #0f766e;
            border-color: #0f766e;
            color: #ffffff;
        }
        .stDownloadButton button:hover {
            background: #115e59;
            border-color: #115e59;
            color: #ffffff;
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
    return buf.getvalue()


@st.cache_data(show_spinner=False)
def cached_xlsx(rows):
    return build_xlsx(rows)


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


def ordered_day_values(df):
    if "nap_szám" in df.columns:
        ordered = (
            df[["nap", "nap_szám"]]
            .dropna(subset=["nap"])
            .sort_values(["nap_szám", "nap"])
            ["nap"]
            .drop_duplicates()
            .tolist()
        )
        if ordered:
            return ordered
    return list(df["nap"].dropna().unique())


MENU_DISPLAY_NAMES = {
    "promocio": "Aktuális ajánlat",
    "Reggeli": "Pékáru és Zabkása",
    "Leves": "Levesek",
    "Kedvencek": "Kedvencek",
    "Hidegkonyha": "Hidegkonyha",
    "Főzelék": "Főzelékek",
    "Tésztás Ételek": "Tészták",
    "Olasz Tészták": "Olasz tészták",
    "Tortillák": "Tortillák",
    "Főételek": "Főételek",
    "Menü": "Hagyományos menük",
    "Vakáció": "Vakáció menü",
    "Desszert": "Desszertek és sütemények",
    "Savanyúság": "Savanyúságok",
    "Cipó": "Cipó",
    "Zero": "Zero menü",
    "ZZabkasa": "Zabkása",
    "ZTortilla": "Teljes kiőrlésű tortilla",
    "CzanikBalazs": "Czanik Balázs ajánlata",
    "KlimaAnita": "Somogyi Kovács Anita ajánlata",
    "DiaCipó": "Diabetikus cipó",
    "Z10": "Z10",
    "IR": "IR",
    "IRN": "IRN",
    "V1": "V1",
    "V2": "V2",
    "V3": "V3",
    "szejtan": "Szejtán",
    "Xixo": "Xixo italok",
}

MENU_SECTION_ORDER = {
    "Zero": 0,
    "promocio": 10,
    "Reggeli": 20,
    "Leves": 30,
    "Kedvencek": 40,
    "Hidegkonyha": 50,
    "Főzelék": 60,
    "Tésztás Ételek": 70,
    "Olasz Tészták": 80,
    "Tortillák": 90,
    "Főételek": 100,
    "Menü": 110,
    "Vakáció": 120,
    "Desszert": 130,
    "Savanyúság": 140,
    "Cipó": 150,
    "ZZabkasa": 160,
    "ZTortilla": 170,
    "CzanikBalazs": 180,
    "KlimaAnita": 190,
    "DiaCipó": 200,
    "Z10": 210,
    "IR": 220,
    "IRN": 230,
    "V1": 240,
    "V2": 250,
    "V3": 260,
    "szejtan": 270,
    "Xixo": 280,
}


def cleaned_menu_title(title):
    text = str(title or "").replace("Õ", "Ő").replace("õ", "ő")
    text = re.sub(r"\s+", " ", text).strip()
    text = re.sub(r"\s*\|\s*Hagyományos konyha\s*\|?\s*", "", text, flags=re.IGNORECASE).strip()
    text = re.sub(r"\s*\|\s*$", "", text).strip()
    text = text.replace("Z ab K ása", "Zabkása")
    if text.isupper() and len(text) > 3:
        text = text.capitalize()
    return text


def menu_display_name(title, section_name, section_df=None):
    if section_df is not None and str(section_name or "") == "promocio" and "sor_név" in section_df.columns:
        line_names = [
            str(name).strip()
            for name in section_df["sor_név"].dropna().unique()
            if str(name).strip()
        ]
        if len(line_names) == 1:
            return line_names[0]

    if section_name in MENU_DISPLAY_NAMES:
        return MENU_DISPLAY_NAMES[section_name]

    cleaned_title = cleaned_menu_title(title)
    cleaned_section = cleaned_menu_title(section_name)
    if cleaned_title and cleaned_title.lower() not in {"promocio", "menü"}:
        return cleaned_title
    return cleaned_section or cleaned_title or "Egyéb"


def menu_display_name_for_row(row):
    if str(row.get("menü_azonosító", "")) == "promocio":
        line_name = str(row.get("sor_név", "")).strip()
        if line_name:
            return line_name
    return menu_display_name(row.get("menü"), row.get("menü_azonosító"))


def menu_sort_value(section_name, fallback_index=0):
    fallback_index = int(fallback_index or 0)
    return MENU_SECTION_ORDER.get(section_name, 900 + fallback_index)


def has_zero_codes(section_df):
    if section_df is None or "kod" not in section_df.columns:
        return False
    return section_df["kod"].dropna().astype(str).str.upper().str.startswith("ZR").any()


def menu_view_sort_key(view):
    is_zero = view.get("section_name") == "Zero" or has_zero_codes(view.get("df"))
    return (
        0 if is_zero else 1,
        menu_sort_value(view.get("section_name"), view.get("source_index", 0)),
        str(view.get("label", "")).casefold(),
    )


def table_height(row_count):
    return min(540, max(180, 78 + 34 * row_count))


def planner_ids(plan_store_key):
    selected_ids = [str(row_id) for row_id in st.session_state.get(plan_store_key, [])]
    unique_ids = []
    seen = set()
    for row_id in selected_ids:
        if row_id not in seen:
            seen.add(row_id)
            unique_ids.append(row_id)
    st.session_state[plan_store_key] = unique_ids
    return unique_ids


def remember_open_section(open_section_key=None, section_id=None):
    if open_section_key and section_id:
        st.session_state[open_section_key] = section_id


def set_plan_item_selected(plan_store_key, row_id, selected, open_section_key=None, section_id=None):
    selected_ids = planner_ids(plan_store_key)
    row_id = str(row_id)
    if selected and row_id not in selected_ids:
        selected_ids.append(row_id)
        st.session_state[plan_store_key] = selected_ids
    elif not selected:
        st.session_state[plan_store_key] = [
            selected_id
            for selected_id in selected_ids
            if selected_id != row_id
        ]
    remember_open_section(open_section_key, section_id)


def add_plan_item(plan_store_key, row_id, open_section_key=None, section_id=None):
    set_plan_item_selected(plan_store_key, row_id, True, open_section_key, section_id)


def remove_plan_item(plan_store_key, row_id, open_section_key=None, section_id=None):
    set_plan_item_selected(plan_store_key, row_id, False, open_section_key, section_id)


def clear_plan(plan_store_key):
    st.session_state[plan_store_key] = []


def build_choice_lookup(df):
    lookup = {}
    for row_id, row in df.iterrows():
        lookup[str(row_id)] = row
    return lookup


def plan_dataframe(selected_rows):
    if not selected_rows:
        return pd.DataFrame()
    records = []
    for row in selected_rows:
        record = row.to_dict()
        record["_row_id"] = str(row.name)
        records.append(record)
    selected_df = pd.DataFrame(records)
    sort_cols = [col for col in ["nap_szám", "menü", "kod", "sor_név", "név"] if col in selected_df.columns]
    if sort_cols:
        selected_df = selected_df.sort_values(sort_cols)
    return selected_df.reset_index(drop=True)


def rows_from_ids(df, row_ids):
    choice_lookup = build_choice_lookup(df)
    return [choice_lookup[row_id] for row_id in row_ids if row_id in choice_lookup]


def plan_summary(selected_df):
    numeric_cols = {
        "ár_ft": "Ár (Ft)",
        "súly_g": "Súly (g)",
        "kcal_adag": "kcal",
        "kj_adag": "kJ",
        "fehérje_g_adag": "Fehérje (g)",
        "szénhidrát_g_adag": "CH (g)",
        "zsír_g_adag": "Zsír (g)",
        "zsír_telített_g_adag": "Telített zsír (g)",
        "cukor_g_adag": "Cukor (g)",
        "rost_g_adag": "Rost (g)",
        "só_g_adag": "Só (g)",
    }
    if selected_df.empty:
        return pd.DataFrame(columns=["Nap", *numeric_cols.values()])

    active_numeric_cols = {
        source_col: display_col
        for source_col, display_col in numeric_cols.items()
        if source_col in selected_df.columns
    }
    daily = selected_df.groupby("nap", sort=False)[list(active_numeric_cols)].sum(numeric_only=True).reset_index()
    daily = daily.rename(columns={"nap": "Nap", **active_numeric_cols})
    total_row = {"Nap": "Összesen"}
    for col in active_numeric_cols.values():
        total_row[col] = daily[col].sum()
    return pd.concat([daily, pd.DataFrame([total_row])], ignore_index=True)


def weekly_codes_text(selected_df, day_values):
    lines = []
    for day in day_values:
        day_df = selected_df[selected_df["nap"] == day]
        if day_df.empty:
            continue
        codes = [str(code) for code in day_df["kod"].dropna().tolist()]
        if codes:
            lines.append(f"{day}: {', '.join(codes)}")
    return "\n".join(lines)


def render_copy_codes_button(text):
    escaped_text = json.dumps(text)
    st.iframe(
        f"""
        <button id="copy-codes" style="
            width: 100%;
            min-height: 38px;
            border-radius: 7px;
            border: 1px solid #0f766e;
            background: #0f766e;
            color: white;
            font-weight: 700;
            cursor: pointer;
            font-family: sans-serif;
        ">Kódlista másolása</button>
        <span id="copy-status" style="
            display: block;
            min-height: 16px;
            margin-top: 4px;
            color: #475569;
            font: 12px sans-serif;
        "></span>
        <script>
        const text = {escaped_text};
        const button = document.getElementById("copy-codes");
        const status = document.getElementById("copy-status");
        button.addEventListener("click", async () => {{
            try {{
                await navigator.clipboard.writeText(text);
                status.textContent = "Másolva";
            }} catch (error) {{
                const area = document.createElement("textarea");
                area.value = text;
                document.body.appendChild(area);
                area.select();
                document.execCommand("copy");
                area.remove();
                status.textContent = "Másolva";
            }}
            window.setTimeout(() => status.textContent = "", 1600);
        }});
        </script>
        """,
        height=62,
    )


def summary_config():
    return {
        "Nap": st.column_config.TextColumn("Nap", width="medium"),
        "Ár (Ft)": st.column_config.NumberColumn("Ár", format="%d Ft", width="small"),
        "Súly (g)": st.column_config.NumberColumn("Súly", format="%.0f g", width="small"),
        "kcal": st.column_config.NumberColumn("kcal", format="%.0f", width="small"),
        "kJ": st.column_config.NumberColumn("kJ", format="%.0f", width="small"),
        "Fehérje (g)": st.column_config.NumberColumn("Fehérje", format="%.1f g", width="small"),
        "CH (g)": st.column_config.NumberColumn("CH", format="%.1f g", width="small"),
        "Zsír (g)": st.column_config.NumberColumn("Zsír", format="%.1f g", width="small"),
        "Telített zsír (g)": st.column_config.NumberColumn("Telített zsír", format="%.1f g", width="small"),
        "Cukor (g)": st.column_config.NumberColumn("Cukor", format="%.1f g", width="small"),
        "Rost (g)": st.column_config.NumberColumn("Rost", format="%.1f g", width="small"),
        "Só (g)": st.column_config.NumberColumn("Só", format="%.1f g", width="small"),
    }


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


def format_number_cell(value, suffix="", decimals=0):
    if pd.isna(value):
        return "-"
    if decimals:
        return f"{value:.{decimals}f}{suffix}"
    return f"{format_int(value)}{suffix}"


def render_menu_picker(section_df, plan_store_key, open_section_key, section_id):
    selected_ids = planner_ids(plan_store_key)
    selected_id_set = set(selected_ids)
    header = st.columns([0.10, 0.08, 0.37, 0.14, 0.10, 0.08, 0.08, 0.08, 0.08], gap="small")
    for col, label in zip(header, ["", "Kód", "Étel", "Sor", "Ár", "kcal", "Feh.", "CH", "Zsír"]):
        col.markdown(f"**{label}**")

    for row_id, row in section_df.iterrows():
        row_id = str(row_id)
        is_selected = row_id in selected_id_set
        cols = st.columns([0.10, 0.08, 0.37, 0.14, 0.10, 0.08, 0.08, 0.08, 0.08], gap="small", vertical_alignment="center")
        with cols[0]:
            st.button(
                "Kivesz" if is_selected else "Hozzáad",
                key=f"menu_plan_{plan_store_key}_{row_id}",
                width="stretch",
                type="secondary" if is_selected else "primary",
                on_click=remove_plan_item if is_selected else add_plan_item,
                args=(plan_store_key, row_id, open_section_key, section_id),
            )
        cols[1].write(row["kod"])
        cols[2].write(row["név"])
        cols[3].write(row["sor_név"])
        cols[4].write(format_number_cell(row["ár_ft"], " Ft"))
        cols[5].write(format_number_cell(row["kcal_adag"]))
        cols[6].write(format_number_cell(row["fehérje_g_adag"], " g", 1))
        cols[7].write(format_number_cell(row["szénhidrát_g_adag"], " g", 1))
        cols[8].write(format_number_cell(row["zsír_g_adag"], " g", 1))


def render_weekly_tray(selected_plan_df, day_values, plan_store_key):
    selected_day_count = selected_plan_df["nap"].nunique()
    total_kcal = selected_plan_df["kcal_adag"].sum()
    total_price = selected_plan_df["ár_ft"].sum()
    total_protein = selected_plan_df["fehérje_g_adag"].sum()
    total_carbs = selected_plan_df["szénhidrát_g_adag"].sum()
    total_fat = selected_plan_df["zsír_g_adag"].sum()

    t1, t2, t3, t4 = st.columns(4)
    t1.metric("Heti kcal", format_int(total_kcal))
    t2.metric("Heti ár", f"{format_int(total_price)} Ft")
    t3.metric("Napok", f"{selected_day_count}/{len(day_values)}")
    t4.metric("Tételek", len(selected_plan_df))

    t5, t6, t7 = st.columns(3)
    t5.metric("Fehérje", f"{total_protein:.1f} g")
    t6.metric("CH", f"{total_carbs:.1f} g")
    t7.metric("Zsír", f"{total_fat:.1f} g")

    for day in day_values:
        day_df = selected_plan_df[selected_plan_df["nap"] == day]
        if day_df.empty:
            continue

        kcal = day_df["kcal_adag"].sum()
        price = day_df["ár_ft"].sum()
        protein = day_df["fehérje_g_adag"].sum()
        carbs = day_df["szénhidrát_g_adag"].sum()
        fat = day_df["zsír_g_adag"].sum()
        st.markdown(
            f"""
            <div class="day-summary">
                <div class="day-summary-name">{day}<br><span class="day-summary-meta">{len(day_df)} tétel</span></div>
                <div class="day-summary-kcal">{format_int(kcal)} kcal</div>
                <div class="day-summary-price">{format_int(price)} Ft</div>
                <div class="day-summary-meta">
                    Fehérje {protein:.1f} g · CH {carbs:.1f} g · Zsír {fat:.1f} g
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        header = st.columns([0.10, 0.09, 0.36, 0.12, 0.08, 0.08, 0.08, 0.08, 0.08], gap="small")
        for col, label in zip(header, ["", "Kód", "Étel", "Menü", "kcal", "Feh.", "CH", "Zsír", "Ár"]):
            col.markdown(f"**{label}**")

        for _, row in day_df.iterrows():
            cols = st.columns([0.10, 0.09, 0.36, 0.12, 0.08, 0.08, 0.08, 0.08, 0.08], gap="small", vertical_alignment="center")
            with cols[0]:
                st.button(
                    "Kivesz",
                    key=f"tray_remove_{plan_store_key}_{row['_row_id']}",
                    width="stretch",
                    type="secondary",
                    on_click=remove_plan_item,
                    args=(plan_store_key, row["_row_id"]),
                )
            cols[1].write(row["kod"])
            cols[2].write(row["név"])
            cols[3].write(menu_display_name_for_row(row))
            cols[4].write(format_number_cell(row["kcal_adag"]))
            cols[5].write(format_number_cell(row["fehérje_g_adag"], " g", 1))
            cols[6].write(format_number_cell(row["szénhidrát_g_adag"], " g", 1))
            cols[7].write(format_number_cell(row["zsír_g_adag"], " g", 1))
            cols[8].write(format_number_cell(row["ár_ft"], " Ft"))


def day_picker_label(day, available_count, selected_count):
    if selected_count:
        return f"{day} ({selected_count}/{available_count})"
    return f"{day} ({available_count})"


def render_day_picker(day_df, sections, plan_store_key, open_section_key):
    if day_df.empty:
        st.info("Nincs találat erre a napra a jelenlegi szűrőkkel.")
        return

    views = build_day_menu_views(sections, day_df)
    if not views:
        st.info("Nincs megjeleníthető menü erre a napra.")
        return

    selected_id_set = set(planner_ids(plan_store_key))
    selected_counts = {
        view["id"]: sum(
            1
            for row_id in view["df"].index
            if str(row_id) in selected_id_set
        )
        for view in views
    }
    section_key = f"{open_section_key}_{day_df['nap'].iloc[0]}"
    view_ids = [view["id"] for view in views]
    if st.session_state.get(section_key) not in view_ids:
        st.session_state[section_key] = views[0]["id"]

    view_lookup = {view["id"]: view for view in views}

    def section_label(section_id):
        view = view_lookup[section_id]
        count = selected_counts[section_id]
        suffix = f" · kiválasztva: {count}" if count else ""
        return f"{view['label']} · {len(view['df'])} tétel · {view['df']['kod'].nunique()} kód{suffix}"

    active_section_id = st.selectbox(
        "Menü",
        view_ids,
        key=section_key,
        format_func=section_label,
    )
    active_view = view_lookup[active_section_id]
    section_df = active_view["df"]
    selected_in_section = [
        str(row_id)
        for row_id in section_df.index
        if str(row_id) in selected_id_set
    ]
    if selected_in_section:
        st.caption(f"Kiválasztva ebben a menüben: {len(selected_in_section)}")
    render_section_meta(section_df)
    render_menu_picker(section_df, plan_store_key, section_key, active_section_id)


def build_day_menu_views(sections, day_df):
    views = []
    section_keys = set()

    for index, section in enumerate(sections, start=1):
        title = section["title"]
        section_name = section["section_name"]
        section_df = day_df[
            (day_df["menü"] == title)
            & (day_df["menü_azonosító"] == section_name)
        ].sort_values(["kod", "sor_név", "név"])

        if section_df.empty:
            continue

        section_keys.add((title, section_name))
        section_label = menu_display_name(title, section_name, section_df)
        views.append(
            {
                "id": f"section_{index}_{section_name}",
                "label": section_label,
                "section_name": section_name,
                "source_index": index,
                "df": section_df,
            }
        )

    if not day_df.empty:
        remaining = day_df[
            ~day_df.apply(lambda row: (row["menü"], row["menü_azonosító"]) in section_keys, axis=1)
        ]
        if not remaining.empty:
            views.append(
                {
                    "id": "section_remaining",
                    "label": "Egyéb",
                    "section_name": "Egyéb",
                    "source_index": 999,
                    "df": remaining,
                }
            )

    return sorted(views, key=menu_view_sort_key)


@st.fragment
def render_planner_area(filtered, df, sections, day_values, selected_days, ev, het, plan_store_key, open_section_key):
    selected_plan_ids = planner_ids(plan_store_key)

    st.divider()

    planner_title_col, planner_action_col = st.columns([0.78, 0.22], vertical_alignment="center")
    with planner_title_col:
        st.markdown('<div class="planner-title">Heti tálca</div>', unsafe_allow_html=True)
        st.markdown(
            '<p class="planner-note">A napi menülapokon a Hozzáad gombbal kerülnek ide az ételek. Egy napra több tétel is választható.</p>',
            unsafe_allow_html=True,
        )
    with planner_action_col:
        st.button(
            "Terv törlése",
            key=f"clear_{plan_store_key}",
            width="stretch",
            type="secondary",
            on_click=clear_plan,
            args=(plan_store_key,),
            disabled=not selected_plan_ids,
        )

    selected_plan_df = plan_dataframe(rows_from_ids(df, selected_plan_ids))
    if selected_plan_df.empty:
        st.info("Még nincs kiválasztott heti menü.")
    else:
        render_weekly_tray(selected_plan_df, day_values, plan_store_key)

    if not selected_plan_df.empty:
        daily_summary = plan_summary(selected_plan_df)
        st.markdown('<div class="planner-grid-title">Napi és heti tápanyagok</div>', unsafe_allow_html=True)
        st.dataframe(
            daily_summary,
            width="stretch",
            hide_index=True,
            height=table_height(len(daily_summary)),
            column_config=summary_config(),
        )

        plan_export_df = selected_plan_df.drop(columns=["_row_id"], errors="ignore")
        plan_xlsx = cached_xlsx(plan_export_df.to_dict("records"))
        copy_col, excel_col = st.columns([0.42, 0.58], vertical_alignment="top")
        with copy_col:
            render_copy_codes_button(weekly_codes_text(selected_plan_df, day_values))
        with excel_col:
            st.download_button(
                label="Heti terv Excel",
                data=plan_xlsx,
                file_name=f"teletal_terv_{ev}_{het}het.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                width="stretch",
                type="primary",
            )

    st.divider()

    st.markdown('<div class="planner-grid-title">Napi menülapok</div>', unsafe_allow_html=True)
    visible_day_values = [day for day in day_values if day in selected_days]
    if visible_day_values:
        selected_counts = selected_plan_df["nap"].value_counts().to_dict() if not selected_plan_df.empty else {}
        available_counts = filtered["nap"].value_counts().to_dict()
        active_day_key = f"active_day_{ev}_{het}"
        if st.session_state.get(active_day_key) not in visible_day_values:
            st.session_state[active_day_key] = visible_day_values[0]

        active_day = st.radio(
            "Nap",
            visible_day_values,
            key=active_day_key,
            horizontal=True,
            format_func=lambda day: day_picker_label(
                day,
                int(available_counts.get(day, 0)),
                int(selected_counts.get(day, 0)),
            ),
        )
        active_day_df = filtered[filtered["nap"] == active_day].sort_values(["menü", "kod", "sor_név", "név"])
        render_day_picker(active_day_df, sections, plan_store_key, open_section_key)
    else:
        st.info("Nincs kiválasztott nap a szűrőben.")

    if filtered.empty:
        st.info("Nincs találat a jelenlegi szűrőkkel.")


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
    if st.button("Frissítés most", width="stretch"):
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

    day_values = ordered_day_values(df)
    selected_days = st.multiselect("Nap szűrő", day_values, default=day_values)

    search = st.text_input("Keresés")

filtered = df[df["nap"].isin(selected_days)]
if search:
    search_mask = filtered.astype(str).apply(
        lambda col: col.str.contains(search, case=False, na=False)
    ).any(axis=1)
    filtered = filtered[search_mask]

week_label = week_labels.get(het, f"{ev}/{het}") if weeks else f"{ev}/{het}"
plan_store_key = f"weekly_plan_{ev}_{het}"
open_section_key = f"open_section_v2_{ev}_{het}"

export_rows = filtered.to_dict("records")
full_export_rows = df.to_dict("records")
filtered_xlsx = cached_xlsx(export_rows) if export_rows else None
full_xlsx = cached_xlsx(full_export_rows)

with st.sidebar:
    st.download_button(
        label="Excel letöltés",
        data=filtered_xlsx if filtered_xlsx else full_xlsx,
        file_name=f"teletal_{ev}_{het}het.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        width="stretch",
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
        width="stretch",
        type="primary",
    )

col1, col2, col3, col4 = st.columns(4)
col1.metric("Menük", filtered["menü"].nunique())
col2.metric("Tételek", len(filtered))
col3.metric("Kódok", filtered["kod"].nunique())
col4.metric("Hét", week_label)

render_planner_area(filtered, df, sections, day_values, selected_days, ev, het, plan_store_key, open_section_key)
