import io
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
import streamlit as st

from teletal_core import MENU_SOURCES, SOURCE_BY_LABEL, discover_weeks, scrape_menu_rows


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


def scrape(progress_bar, status_text, selected_labels, week):
    sources = [SOURCE_BY_LABEL[label] for label in selected_labels]

    def update_progress(index, total, item):
        status_text.text(
            "Tápérték letöltése: "
            f"{index}/{total}  ({item['source']}, {item['kod']}, {item['nap_nev']})"
        )
        if total:
            progress_bar.progress(index / total)

    status_text.text("Étlapok letöltése...")
    rows, pages = scrape_menu_rows(
        sources=sources,
        week=week,
        progress_callback=update_progress,
        delay_seconds=0.1,
    )

    ev = pages[0]["ev"] if pages else ""
    het = pages[0]["het"] if pages else ""
    return rows, ev, het, pages


@st.cache_data(ttl=1800)
def week_options():
    return discover_weeks()


# ---------------------------------------------------------------------------
# UI
# ---------------------------------------------------------------------------

st.set_page_config(page_title="Teletál étlap letöltő", page_icon="🍽️", layout="wide")
st.title("🍽️ Teletál étlap letöltő")
st.write("Letölti a kiválasztott étlapokat a tápértékekkel együtt, és Excel fájlba menti.")

source_labels = [source.label for source in MENU_SOURCES]
selected_labels = st.multiselect("Étlapok", source_labels, default=source_labels)

try:
    weeks, active_week = week_options()
except Exception:
    weeks, active_week = [], None

if weeks:
    week = st.selectbox(
        "Hét",
        weeks,
        index=weeks.index(active_week) if active_week in weeks else 0,
        format_func=lambda value: f"{value}. hét",
    )
else:
    week = None

if st.button("📥 Letöltés indítása", type="primary", disabled=not selected_labels):
    progress_bar = st.progress(0)
    status_text = st.empty()

    try:
        rows, ev, het, pages = scrape(progress_bar, status_text, selected_labels, week)
        if not rows:
            st.warning("Nem található letölthető étel a kiválasztott étlapokon.")
        else:
            status_text.text("Excel fájl összeállítása...")
            xlsx_buf = build_xlsx(rows)

            status_text.text(f"✅ Kész! {len(rows)} étel, {ev} {het}. hét")
            progress_bar.progress(1.0)

            st.success(f"Sikeresen letöltve {len(rows)} étel a {ev}. évi {het}. hétre.")
            st.caption(
                "Források: "
                + ", ".join(f"{page['source']} ({page['item_count']})" for page in pages)
            )
            st.download_button(
                label="💾 Excel fájl letöltése",
                data=xlsx_buf,
                file_name=f"teletal_{ev}_{het}het.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

            # preview table
            import pandas as pd
            df = pd.DataFrame(rows)
            st.dataframe(df, use_container_width=True)

    except Exception as e:
        st.error(f"Hiba történt: {e}")
