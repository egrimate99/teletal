import re
import time
import io
import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
import streamlit as st

BASE_URL = "https://www.teletal.hu"
ETLAP_URL = f"{BASE_URL}/etlap"
KODINFO_URL = f"{BASE_URL}/ajax/kodinfo"

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Referer": ETLAP_URL,
}

DAY_NAMES = {
    "1": "Hétfő", "2": "Kedd", "3": "Szerda",
    "4": "Csütörtök", "5": "Péntek", "6": "Szombat", "7": "Vasárnap",
}


def extract_js_var(html, var_name):
    m = re.search(
        rf"var\s+{re.escape(var_name)}\s*=\s*['\"]?([^'\";\s]+)['\"]?\s*;", html
    )
    return m.group(1) if m else None


def fetch_kodinfo(session, ev, het, tipus, nap, kod):
    empty = dict(
        nev="", suly_g="", kcal_adag="", kj_adag="",
        zsir_adag="", zsir_telitett_adag="", ch_adag="",
        cukor_adag="", rost_adag="", feherje_adag="", so_adag="",
        kcal_100g="", kj_100g="", zsir_100g="", ch_100g="",
        feherje_100g="", allergének="",
    )
    try:
        resp = session.get(
            KODINFO_URL,
            params={"ev": ev, "het": het, "tipus": tipus, "nap": nap, "kod": kod},
            timeout=15,
        )
        if resp.status_code != 200:
            return empty
        soup = BeautifulSoup(resp.text, "html.parser")

        name_tag = soup.find("h1", class_="uk-article-title")
        nev = name_tag.get_text(strip=True) if name_tag else ""

        badge = soup.find("span", class_="uk-badge-neutral")
        suly_text = badge.get_text(strip=True) if badge else ""
        suly_m = re.search(r"(\d+)", suly_text)
        suly = suly_m.group(1) if suly_m else ""

        def span_val(spans, idx):
            if idx >= len(spans):
                return ""
            raw = spans[idx].get_text(strip=True).replace(",", "")
            m = re.search(r"[\d]+\.[\d]+", raw)
            return m.group(0) if m else ""

        adag = soup.find_all("span", class_="en_adag")
        per100 = soup.find_all("span", class_="en_100")

        allerg_tag = soup.find("strong", string=re.compile(r"Allerg", re.I))
        allergek = ""
        if allerg_tag:
            parent = allerg_tag.find_parent()
            sib = parent.find_next_sibling("span") if parent else None
            allergek = sib.get_text(strip=True) if sib else ""

        return {
            "nev": nev, "suly_g": suly,
            "kcal_adag": span_val(adag, 1), "kj_adag": span_val(adag, 2),
            "zsir_adag": span_val(adag, 3), "zsir_telitett_adag": span_val(adag, 4),
            "ch_adag": span_val(adag, 5), "cukor_adag": span_val(adag, 6),
            "rost_adag": span_val(adag, 7), "feherje_adag": span_val(adag, 8),
            "so_adag": span_val(adag, 9),
            "kcal_100g": span_val(per100, 1), "kj_100g": span_val(per100, 2),
            "zsir_100g": span_val(per100, 3), "ch_100g": span_val(per100, 5),
            "feherje_100g": span_val(per100, 8),
            "allergének": allergek,
        }
    except Exception as e:
        st.warning(f"Hiba: kod={kod} nap={nap} – {e}")
        return empty


def build_xlsx(rows):
    fieldnames = list(rows[0].keys())
    numeric_cols = {
        "kcal_adag", "kj_adag", "zsír_g_adag", "zsír_telített_g_adag",
        "szénhidrát_g_adag", "cukor_g_adag", "rost_g_adag", "fehérje_g_adag",
        "só_g_adag", "kcal_100g", "kj_100g", "zsír_g_100g",
        "szénhidrát_g_100g", "fehérje_g_100g", "súly_g", "nap_szám", "hét", "év",
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


def scrape(progress_bar, status_text):
    session = requests.Session()
    session.headers.update(HEADERS)

    status_text.text("Étlap letöltése...")
    resp = session.get(ETLAP_URL, timeout=20)
    resp.raise_for_status()
    html = resp.text
    soup = BeautifulSoup(html, "html.parser")

    ev    = extract_js_var(html, "_ev")    or "2026"
    het   = extract_js_var(html, "_het")   or "14"
    tipus = extract_js_var(html, "_tipus") or "1"

    info_items = []
    seen = set()
    for tag in soup.find_all(True, attrs={"kod": True, "nap": True}):
        kod = tag.get("kod", "").strip()
        nap = tag.get("nap", "").strip()
        if kod and nap and (kod, nap) not in seen:
            seen.add((kod, nap))
            info_items.append({"kod": kod, "nap": nap})

    name_lookup = {}
    for cell in soup.find_all(class_=re.compile(r"kod_")):
        cls_list = cell.get("class", [])
        kod_class = next((c for c in cls_list if c.startswith("kod_")), None)
        if not kod_class:
            continue
        kod = kod_class[4:]
        name_tag = cell.find(["p", "h5", "h4", "span", "div"], recursive=True)
        price_tag = cell.find(string=re.compile(r"\d+\s*Ft"))
        name_lookup[kod] = (
            name_tag.get_text(strip=True) if name_tag else "",
            price_tag.strip() if price_tag else "",
        )

    rows = []
    total = len(info_items)
    for idx, item in enumerate(info_items):
        kod = item["kod"]
        nap = item["nap"]
        status_text.text(f"Tápérték letöltése: {idx + 1}/{total}  ({kod}, {DAY_NAMES.get(nap, nap)})")
        progress_bar.progress((idx + 1) / total)

        nutr = fetch_kodinfo(session, ev, het, tipus, nap, kod)
        page_name = name_lookup.get(kod, ("", ""))
        display_name = nutr["nev"] or page_name[0]
        price = page_name[1]

        rows.append({
            "hét": het, "év": ev, "nap_szám": nap,
            "nap": DAY_NAMES.get(nap, nap),
            "kod": kod, "név": display_name, "ár_ft": price,
            "súly_g": nutr["suly_g"],
            "kcal_adag": nutr["kcal_adag"], "kj_adag": nutr["kj_adag"],
            "zsír_g_adag": nutr["zsir_adag"],
            "zsír_telített_g_adag": nutr["zsir_telitett_adag"],
            "szénhidrát_g_adag": nutr["ch_adag"],
            "cukor_g_adag": nutr["cukor_adag"],
            "rost_g_adag": nutr["rost_adag"],
            "fehérje_g_adag": nutr["feherje_adag"],
            "só_g_adag": nutr["so_adag"],
            "kcal_100g": nutr["kcal_100g"], "kj_100g": nutr["kj_100g"],
            "zsír_g_100g": nutr["zsir_100g"],
            "szénhidrát_g_100g": nutr["ch_100g"],
            "fehérje_g_100g": nutr["feherje_100g"],
            "allergének": nutr["allergének"],
        })
        time.sleep(0.3)

    return rows, ev, het


# ---------------------------------------------------------------------------
# UI
# ---------------------------------------------------------------------------

st.set_page_config(page_title="Teletál étlap letöltő", page_icon="🍽️", layout="wide")
st.title("🍽️ Teletál étlap letöltő")
st.write("Letölti az aktuális heti étlapot a tápértékekkel együtt, és Excel fájlba menti.")

if st.button("📥 Letöltés indítása", type="primary"):
    progress_bar = st.progress(0)
    status_text = st.empty()

    try:
        rows, ev, het = scrape(progress_bar, status_text)
        status_text.text("Excel fájl összeállítása...")
        xlsx_buf = build_xlsx(rows)

        status_text.text(f"✅ Kész! {len(rows)} étel, {ev} {het}. hét")
        progress_bar.progress(1.0)

        st.success(f"Sikeresen letöltve {len(rows)} étel a {ev}. évi {het}. hétre.")
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
