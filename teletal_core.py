from __future__ import annotations

import re
import time
from dataclasses import dataclass
from typing import Callable, Iterable

import requests
from bs4 import BeautifulSoup


BASE_URL = "https://www.teletal.hu"

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Referer": f"{BASE_URL}/etlap",
}

DAY_NAMES = {
    "1": "Hétfő",
    "2": "Kedd",
    "3": "Szerda",
    "4": "Csütörtök",
    "5": "Péntek",
    "6": "Szombat",
    "7": "Vasárnap",
}


@dataclass(frozen=True)
class MenuSource:
    label: str
    base_url: str
    path: str


MENU_SOURCES = (
    MenuSource("Teletál", "https://www.teletal.hu", "/etlap"),
    MenuSource("Teletál Vega", "https://www.teletal.hu", "/etlap/vega"),
    MenuSource("Alakreform / Réka Menü", "https://www.rekamenu.hu", "/etlap"),
)

SOURCE_BY_LABEL = {source.label: source for source in MENU_SOURCES}

NutritionProgress = Callable[[int, int, dict], None]


def clean_text(text: str) -> str:
    return re.sub(r"\s+", " ", text).strip()


def source_url(source: MenuSource, week: str | None = None) -> str:
    path = source.path.rstrip("/")
    if week:
        path = f"{path}/{week}"
    return f"{source.base_url}{path}"


def extract_js_var(html: str, var_name: str) -> str | None:
    pattern = rf"var\s+{re.escape(var_name)}\s*=\s*['\"]?([^'\";\s]+)['\"]?\s*;"
    match = re.search(pattern, html)
    return match.group(1) if match else None


def extract_menu_context(html: str, soup: BeautifulSoup) -> dict[str, str]:
    form = soup.find("form", id="etlap_form")
    context: dict[str, str] = {}

    for field_name in ("ev", "het", "tipus"):
        value = form.get(field_name) if form else None
        if not value:
            value = extract_js_var(html, f"_{field_name}")
        if value:
            context[field_name] = value.strip()

    missing = [field for field in ("ev", "het", "tipus") if not context.get(field)]
    if missing:
        raise ValueError(f"Nem található étlap paraméter: {', '.join(missing)}")

    return context


def normalize_number(raw: str) -> str:
    match = re.search(r"\d[\d.,]*", raw.replace("\xa0", " "))
    if not match:
        return ""

    value = match.group(0)
    if "," in value and "." in value:
        value = value.replace(",", "")
    elif "," in value:
        value = value.replace(",", ".")
    return value


def parse_price_ft(text: str) -> str:
    matches = list(re.finditer(r"(\d[\d.\s]*)(?:,-)?\s*Ft\b", text))
    if not matches:
        return ""
    return re.sub(r"\D", "", matches[-1].group(1))


def extract_item_category(tag) -> str:
    row = tag.find_parent("tr")
    if row:
        category_cell = row.find(class_=re.compile(r"\bmenu-cell-code-sub2\b"))
        if category_cell:
            category = clean_text(category_cell.get_text(" ", strip=True))
            if category:
                return category

    for parent in tag.parents:
        heading = parent.find_previous(["h2", "h3", "h4", "h5"])
        if heading:
            category = clean_text(heading.get_text(" ", strip=True))
            if category:
                return category

    return ""


def extract_page_item_name(tag) -> str:
    cell = tag.find_parent(class_=re.compile(r"\bmenu-cell-text\b"))
    if not cell:
        return ""

    name_tag = cell.find(class_=re.compile(r"\buk-text-break\b"))
    if name_tag:
        return clean_text(name_tag.get_text(" ", strip=True))

    text = clean_text(cell.get_text(" ", strip=True))
    return re.sub(r"\d[\d.\s]*(?:,-)?\s*Ft\b.*$", "", text).strip()


def extract_page_item_price(tag) -> str:
    cell = tag.find_parent(class_=re.compile(r"\bmenu-cell-text\b"))
    if not cell:
        return ""

    price_tag = cell.find(class_=re.compile(r"\bmenu-price-field\b"))
    price_text = price_tag.get_text(" ", strip=True) if price_tag else cell.get_text(" ", strip=True)
    return parse_price_ft(clean_text(price_text))


def parse_menu_items(
    soup: BeautifulSoup,
    context: dict[str, str],
    source: MenuSource,
    page_url: str,
) -> list[dict]:
    items: list[dict] = []
    seen: set[tuple[str, str, str]] = set()

    for tag in soup.find_all(True, attrs={"kod": True, "nap": True}):
        kod = tag.get("kod", "").strip()
        nap = tag.get("nap", "").strip()
        if not kod or not nap:
            continue

        key = (source.label, kod, nap)
        if key in seen:
            continue
        seen.add(key)

        items.append(
            {
                "source": source.label,
                "source_base_url": source.base_url,
                "source_url": page_url,
                "ev": tag.get("ev", context["ev"]).strip(),
                "het": tag.get("het", context["het"]).strip(),
                "tipus": tag.get("tipus", context["tipus"]).strip(),
                "nap": nap,
                "nap_nev": DAY_NAMES.get(nap, nap),
                "kod": kod,
                "kategoria": extract_item_category(tag),
                "page_name": extract_page_item_name(tag),
                "price_ft": extract_page_item_price(tag),
            }
        )

    return items


def fetch_menu_page(
    session: requests.Session,
    source: MenuSource,
    week: str | None = None,
) -> tuple[dict, list[dict]]:
    url = source_url(source, week)
    response = session.get(url, headers={"Referer": source.base_url}, timeout=20)
    response.raise_for_status()

    soup = BeautifulSoup(response.text, "html.parser")
    context = extract_menu_context(response.text, soup)
    items = parse_menu_items(soup, context, source, response.url)
    page_info = {
        "source": source.label,
        "url": response.url,
        "ev": context["ev"],
        "het": context["het"],
        "tipus": context["tipus"],
        "item_count": len(items),
    }
    return page_info, items


def nutrition_empty() -> dict[str, str]:
    return {
        "nev": "",
        "suly_g": "",
        "kcal_adag": "",
        "kj_adag": "",
        "zsir_adag": "",
        "zsir_telitett_adag": "",
        "ch_adag": "",
        "cukor_adag": "",
        "rost_adag": "",
        "feherje_adag": "",
        "so_adag": "",
        "kcal_100g": "",
        "kj_100g": "",
        "zsir_100g": "",
        "ch_100g": "",
        "feherje_100g": "",
        "allergének": "",
    }


def span_value(spans: list, index: int) -> str:
    if index >= len(spans):
        return ""
    return normalize_number(spans[index].get_text(" ", strip=True))


def extract_weight_g(soup: BeautifulSoup) -> str:
    for badge in soup.select("span.uk-badge-neutral"):
        text = badge.get_text(" ", strip=True)
        if re.search(r"\d[\d.,]*\s*g\b", text, re.IGNORECASE):
            return normalize_number(text)
    return ""


def extract_allergens(soup: BeautifulSoup) -> str:
    allergen_label = soup.find("strong", string=re.compile(r"Allerg", re.IGNORECASE))
    if not allergen_label:
        return ""

    label_parent = allergen_label.find_parent()
    allergen_span = label_parent.find_next_sibling("span") if label_parent else None
    return allergen_span.get_text(" ", strip=True) if allergen_span else ""


def fetch_kodinfo(session: requests.Session, item: dict) -> dict[str, str]:
    try:
        response = session.get(
            f"{item['source_base_url']}/ajax/kodinfo",
            params={
                "ev": item["ev"],
                "het": item["het"],
                "tipus": item["tipus"],
                "nap": item["nap"],
                "kod": item["kod"],
            },
            headers={"Referer": item["source_url"]},
            timeout=15,
        )
        if response.status_code != 200:
            return nutrition_empty()

        soup = BeautifulSoup(response.text, "html.parser")
        name_tag = soup.find("h1", class_="uk-article-title")
        adag_spans = soup.find_all("span", class_="en_adag")
        per100_spans = soup.find_all("span", class_="en_100")

        return {
            "nev": name_tag.get_text(" ", strip=True) if name_tag else "",
            "suly_g": extract_weight_g(soup),
            "kcal_adag": span_value(adag_spans, 1),
            "kj_adag": span_value(adag_spans, 2),
            "zsir_adag": span_value(adag_spans, 3),
            "zsir_telitett_adag": span_value(adag_spans, 4),
            "ch_adag": span_value(adag_spans, 5),
            "cukor_adag": span_value(adag_spans, 6),
            "rost_adag": span_value(adag_spans, 7),
            "feherje_adag": span_value(adag_spans, 8),
            "so_adag": span_value(adag_spans, 9),
            "kcal_100g": span_value(per100_spans, 1),
            "kj_100g": span_value(per100_spans, 2),
            "zsir_100g": span_value(per100_spans, 3),
            "ch_100g": span_value(per100_spans, 5),
            "feherje_100g": span_value(per100_spans, 8),
            "allergének": extract_allergens(soup),
        }
    except Exception:
        return nutrition_empty()


def row_from_item(item: dict, nutrition: dict[str, str]) -> dict:
    return {
        "forrás": item["source"],
        "kategória": item["kategoria"],
        "hét": item["het"],
        "év": item["ev"],
        "nap_szám": item["nap"],
        "nap": item["nap_nev"],
        "kod": item["kod"],
        "név": item["page_name"] or nutrition["nev"],
        "ár_ft": item["price_ft"],
        "súly_g": nutrition["suly_g"],
        "kcal_adag": nutrition["kcal_adag"],
        "kj_adag": nutrition["kj_adag"],
        "zsír_g_adag": nutrition["zsir_adag"],
        "zsír_telített_g_adag": nutrition["zsir_telitett_adag"],
        "szénhidrát_g_adag": nutrition["ch_adag"],
        "cukor_g_adag": nutrition["cukor_adag"],
        "rost_g_adag": nutrition["rost_adag"],
        "fehérje_g_adag": nutrition["feherje_adag"],
        "só_g_adag": nutrition["so_adag"],
        "kcal_100g": nutrition["kcal_100g"],
        "kj_100g": nutrition["kj_100g"],
        "zsír_g_100g": nutrition["zsir_100g"],
        "szénhidrát_g_100g": nutrition["ch_100g"],
        "fehérje_g_100g": nutrition["feherje_100g"],
        "allergének": nutrition["allergének"],
        "forrás_url": item["source_url"],
    }


def scrape_menu_rows(
    sources: Iterable[MenuSource] | None = None,
    week: str | None = None,
    progress_callback: NutritionProgress | None = None,
    delay_seconds: float = 0.1,
) -> tuple[list[dict], list[dict]]:
    selected_sources = list(sources or MENU_SOURCES)

    session = requests.Session()
    session.headers.update(HEADERS)

    page_infos: list[dict] = []
    items: list[dict] = []
    for source in selected_sources:
        page_info, page_items = fetch_menu_page(session, source, week)
        page_infos.append(page_info)
        items.extend(page_items)

    rows: list[dict] = []
    total = len(items)
    for index, item in enumerate(items, start=1):
        if progress_callback:
            progress_callback(index, total, item)
        rows.append(row_from_item(item, fetch_kodinfo(session, item)))
        if delay_seconds:
            time.sleep(delay_seconds)

    return rows, page_infos


def discover_weeks(source: MenuSource = MENU_SOURCES[0]) -> tuple[list[str], str | None]:
    session = requests.Session()
    session.headers.update(HEADERS)
    response = session.get(source_url(source), timeout=20)
    response.raise_for_status()

    soup = BeautifulSoup(response.text, "html.parser")
    context = extract_menu_context(response.text, soup)
    weeks: list[str] = []
    seen: set[str] = set()

    for link in soup.find_all("a", href=True):
        match = re.search(r"/etlap/(\d+)$", link["href"])
        if match and match.group(1) not in seen:
            week = match.group(1)
            seen.add(week)
            weeks.append(week)

    if context["het"] not in seen:
        weeks.insert(0, context["het"])

    return weeks, context["het"]
