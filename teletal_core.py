from __future__ import annotations

import re
import time
from dataclasses import dataclass
from typing import Callable, Iterable
from urllib.parse import quote

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


def format_number(value: float) -> str:
    rounded = round(value, 1)
    if rounded == int(rounded):
        return str(int(rounded))
    return f"{rounded:.1f}"


def menu_section_tags(soup: BeautifulSoup) -> list:
    return [
        section
        for section in soup.find_all("section")
        if section.get("section") and section.get("ev") and section.get("het") and section.get("ewid")
    ]


def extract_section_title(section) -> str:
    header = section.find(class_=re.compile(r"\bmenu-section-header\b"))
    if header:
        title = clean_text(header.get_text(" ", strip=True))
        if title:
            if len(title) <= 120:
                return title
            return clean_text(section.get("section", "")) or title[:120]

    heading = section.find(["h1", "h2", "h3"])
    if heading:
        title = clean_text(heading.get_text(" ", strip=True))
        if title and len(title) <= 120:
            return title

    return clean_text(section.get("section", ""))


def section_key(source: MenuSource, section, index: int) -> str:
    section_name = clean_text(section.get("section", ""))
    return f"{source.label}|{source.path}|{index}|{section_name}"


def section_summary(source: MenuSource, section, index: int) -> dict:
    dummy = section.find(class_=re.compile(r"\bdummy\b"))
    return {
        "key": section_key(source, section, index),
        "source": source.label,
        "section_name": clean_text(section.get("section", "")),
        "title": extract_section_title(section),
        "ev": section.get("ev", ""),
        "het": section.get("het", ""),
        "ewid": section.get("ewid", ""),
        "is_lazy": bool(dummy),
        "dummy_rows": int(dummy.get("db", 0)) if dummy and str(dummy.get("db", "")).isdigit() else 0,
    }


def discover_menu_sections(
    week: str | None = None,
    source: MenuSource = MENU_SOURCES[0],
) -> tuple[list[dict], dict]:
    session = requests.Session()
    session.headers.update(HEADERS)
    response = session.get(source_url(source, week), timeout=20)
    response.raise_for_status()

    soup = BeautifulSoup(response.text, "html.parser")
    context = extract_menu_context(response.text, soup)
    sections = [
        section_summary(source, section, index)
        for index, section in enumerate(menu_section_tags(soup), start=1)
    ]
    return sections, {"ev": context["ev"], "het": context["het"], "tipus": context["tipus"], "url": response.url}


def fetch_lazy_section(session: requests.Session, source: MenuSource, section) -> str:
    ev = quote(section.get("ev", ""), safe="")
    het = quote(section.get("het", ""), safe="")
    ewid = quote(section.get("ewid", ""), safe="")
    varname = quote(section.get("section", ""), safe="")
    url = f"{source.base_url}/ajax/szekcio?ev={ev}&het={het}&ewid={ewid}&varname={varname}"
    response = session.get(
        url,
        headers={"Referer": source_url(source, section.get("het")), "X-Requested-With": "XMLHttpRequest"},
        timeout=20,
    )
    response.raise_for_status()
    return response.text


def load_lazy_section(session: requests.Session, source: MenuSource, section) -> None:
    dummy = section.find(class_=re.compile(r"\bdummy\b"))
    if not dummy:
        return

    fragment_html = fetch_lazy_section(session, source, section)
    if not fragment_html.strip() or fragment_html.startswith("Wrong varname:"):
        return

    fragment = BeautifulSoup(fragment_html, "html.parser")
    for child in reversed(list(fragment.contents)):
        dummy.insert_after(child)
    dummy.decompose()


def extract_item_line_name(tag, section) -> str:
    row = tag.find_parent("tr")
    if row:
        cells = row.find_all("td", recursive=False)
        if len(cells) >= 2:
            line_name = clean_text(cells[1].get_text(" ", strip=True))
            if line_name:
                return line_name

    kod = tag.get("kod", "").strip()
    if kod:
        code_row = section.find("tr", attrs={"kod": kod})
        if code_row:
            cells = code_row.find_all("td", recursive=False)
            if len(cells) >= 2:
                line_name = clean_text(cells[1].get_text(" ", strip=True))
                if line_name:
                    return line_name

    return ""


def extract_page_item_name(tag, section) -> str:
    cell = tag.find_parent(class_=re.compile(r"\bmenu-cell-text\b"))
    if cell:
        name_tag = cell.find(class_=re.compile(r"\buk-text-break\b"))
        if name_tag:
            return clean_text(name_tag.get_text(" ", strip=True))

        text = clean_text(cell.get_text(" ", strip=True))
        name = re.sub(r"\d[\d.\s]*(?:,-)?\s*Ft\b.*$", "", text).strip()
        if name:
            return name

    kod = tag.get("kod", "").strip()
    nap = tag.get("nap", "").strip()
    if kod and nap:
        names = []
        for block in section.find_all(class_=re.compile(rf"\bhl_{re.escape(kod)}_{re.escape(nap)}\b")):
            name_tag = block.find(class_=re.compile(r"\buk-text-break\b"))
            name = clean_text(name_tag.get_text(" ", strip=True)) if name_tag else ""
            if name:
                names.append(name)
        return " | ".join(dict.fromkeys(names))

    return ""


def extract_page_item_price(tag) -> str:
    cell = tag.find_parent(class_=re.compile(r"\bmenu-cell-text\b"))
    if not cell:
        return ""

    price_tag = cell.find(class_=re.compile(r"\bmenu-price-field\b"))
    price_text = price_tag.get_text(" ", strip=True) if price_tag else cell.get_text(" ", strip=True)
    return parse_price_ft(clean_text(price_text))


def parse_section_items(
    section,
    summary: dict,
    context: dict[str, str],
    source: MenuSource,
    page_url: str,
) -> list[dict]:
    items: list[dict] = []
    seen: set[tuple[str, str, str]] = set()

    for tag in section.find_all(True, attrs={"kod": True, "nap": True}):
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
                "menu_key": summary["key"],
                "menu_title": summary["title"],
                "section_name": summary["section_name"],
                "line_name": extract_item_line_name(tag, section),
                "page_name": extract_page_item_name(tag, section),
                "price_ft": extract_page_item_price(tag),
            }
        )

    return items


def fetch_menu_page(
    session: requests.Session,
    source: MenuSource,
    week: str | None = None,
    selected_section_keys: set[str] | None = None,
) -> tuple[dict, list[dict]]:
    url = source_url(source, week)
    response = session.get(url, headers={"Referer": source.base_url}, timeout=20)
    response.raise_for_status()

    soup = BeautifulSoup(response.text, "html.parser")
    context = extract_menu_context(response.text, soup)
    items: list[dict] = []
    summaries: list[dict] = []
    for index, section in enumerate(menu_section_tags(soup), start=1):
        summary = section_summary(source, section, index)
        summaries.append(summary)
        if selected_section_keys is not None and summary["key"] not in selected_section_keys:
            continue
        load_lazy_section(session, source, section)
        items.extend(parse_section_items(section, summary, context, source, response.url))

    page_info = {
        "source": source.label,
        "url": response.url,
        "ev": context["ev"],
        "het": context["het"],
        "tipus": context["tipus"],
        "section_count": len(summaries),
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


def span_groups(spans: list) -> list[list[str]]:
    values = [normalize_number(span.get_text(" ", strip=True)) for span in spans]
    groups = []
    for index in range(0, len(values), 10):
        group = values[index : index + 10]
        if len(group) == 10:
            groups.append(group)
    return groups


def extract_weight_g(soup: BeautifulSoup) -> str:
    weights = []
    for badge in soup.select("span.uk-badge-neutral"):
        text = badge.get_text(" ", strip=True)
        if re.search(r"\d[\d.,]*\s*g\b", text, re.IGNORECASE):
            value = normalize_number(text)
            if value:
                weights.append(float(value))
    return format_number(sum(weights)) if weights else ""


def extract_allergens(soup: BeautifulSoup) -> str:
    allergens = []
    for allergen_label in soup.find_all("strong", string=re.compile(r"Allerg", re.IGNORECASE)):
        label_parent = allergen_label.find_parent()
        allergen_span = label_parent.find_next_sibling("span") if label_parent else None
        if not allergen_span:
            continue
        for allergen in clean_text(allergen_span.get_text(" ", strip=True)).split(","):
            allergen = allergen.strip()
            if allergen and allergen not in allergens:
                allergens.append(allergen)
    return ", ".join(allergens)


def aggregate_adag_values(groups: list[list[str]]) -> dict[str, str]:
    if not groups:
        return {
            "kcal_adag": "",
            "kj_adag": "",
            "zsir_adag": "",
            "zsir_telitett_adag": "",
            "ch_adag": "",
            "cukor_adag": "",
            "rost_adag": "",
            "feherje_adag": "",
            "so_adag": "",
        }

    fields = {
        "kcal_adag": 1,
        "kj_adag": 2,
        "zsir_adag": 3,
        "zsir_telitett_adag": 4,
        "ch_adag": 5,
        "cukor_adag": 6,
        "rost_adag": 7,
        "feherje_adag": 8,
        "so_adag": 9,
    }
    totals = {}
    for field, index in fields.items():
        values = [float(group[index]) for group in groups if group[index]]
        totals[field] = format_number(sum(values)) if values else ""
    return totals


def aggregate_per100_values(adag_values: dict[str, str], weight_g: str, per100_groups: list[list[str]]) -> dict[str, str]:
    fields = {
        "kcal_100g": ("kcal_adag", 1),
        "kj_100g": ("kj_adag", 2),
        "zsir_100g": ("zsir_adag", 3),
        "ch_100g": ("ch_adag", 5),
        "feherje_100g": ("feherje_adag", 8),
    }

    if weight_g:
        weight = float(weight_g)
        if weight > 0:
            return {
                output_field: format_number(float(adag_values[adag_field]) * 100 / weight)
                if adag_values.get(adag_field)
                else ""
                for output_field, (adag_field, _) in fields.items()
            }

    if not per100_groups:
        return {field: "" for field in fields}

    first_group = per100_groups[0]
    return {
        output_field: first_group[index]
        for output_field, (_, index) in fields.items()
    }


def extract_detail_name(soup: BeautifulSoup) -> str:
    headings = [clean_text(h.get_text(" ", strip=True)) for h in soup.find_all("h1", class_="uk-article-title")]
    if not headings:
        return ""
    if len(headings) > 1 and "összesítés" in headings[0].lower():
        return " | ".join(headings[1:])
    return headings[0]


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
        adag_groups = span_groups(soup.find_all("span", class_="en_adag"))
        per100_groups = span_groups(soup.find_all("span", class_="en_100"))
        weight_g = extract_weight_g(soup)
        adag_values = aggregate_adag_values(adag_groups)
        per100_values = aggregate_per100_values(adag_values, weight_g, per100_groups)

        return {
            "nev": extract_detail_name(soup),
            "suly_g": weight_g,
            **adag_values,
            **per100_values,
            "allergének": extract_allergens(soup),
        }
    except Exception:
        return nutrition_empty()


def row_from_item(item: dict, nutrition: dict[str, str]) -> dict:
    return {
        "menü": item["menu_title"],
        "menü_azonosító": item["section_name"],
        "sor_név": item["line_name"],
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
    }


def scrape_menu_rows(
    sources: Iterable[MenuSource] | None = None,
    week: str | None = None,
    selected_section_keys: set[str] | None = None,
    progress_callback: NutritionProgress | None = None,
    delay_seconds: float = 0.1,
) -> tuple[list[dict], list[dict]]:
    selected_sources = list(sources or MENU_SOURCES)

    session = requests.Session()
    session.headers.update(HEADERS)

    page_infos: list[dict] = []
    items: list[dict] = []
    for source in selected_sources:
        page_info, page_items = fetch_menu_page(session, source, week, selected_section_keys)
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
