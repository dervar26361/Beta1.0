#!/usr/bin/env python3
"""
HorgászÁr.hu – webshop termékadat-gyűjtő (v2)
==============================================

Három bemeneti mód:
  --shops shops.json   Kategóriaoldalakat jár be boltonként (lapozással), összegyűjti a
                       termék-linkeket, majd minden termékoldalról kinyeri az adatokat.
  --sitemap URL ...    XML sitemap(ek)ből veszi a linkeket.
  --urls linkek.txt    Soronként egy termék-URL.

Plusz egy hibakereső mód:
  --diagnose URL       Egy oldalt letölt, elmenti a HTML-t, és kiírja, mit talált benne
                       (JSON-LD, meta tagek, változattábla, termék-linkek, lapozás).

Kinyert adatok termékenként (változatonként külön sorban):
  bolt, URL, név, változat, cikkszám, márka, ár, eredeti (áthúzott) ár, pénznem,
  elérhetőség, hossz (m), dobósúly (g), kép, státusz, hiba, időbélyeg.

Példák:
  python horgasz_scraper.py --shops shops.json --excel
  python horgasz_scraper.py --shops shops.json --only pecaplaza.hu --limit 10
  python horgasz_scraper.py --diagnose https://tacklebait.hu/Botok/Feeder-botok-picker-botok
"""

from __future__ import annotations

import argparse
import csv
import gzip
import json
import logging
import re
import sys
import time
import xml.etree.ElementTree as ET
from dataclasses import dataclass, asdict, fields, replace
from datetime import datetime
from urllib.parse import urljoin, urlparse, urlencode, parse_qs, urlunparse, urldefrag
from urllib.robotparser import RobotFileParser

import requests
from bs4 import BeautifulSoup, Tag
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

# ---------------------------------------------------------------------------
# Alapbeállítások
# ---------------------------------------------------------------------------

USER_AGENT = "Mozilla/5.0 (compatible; HorgaszarBot/1.0; +https://horgaszar.hu)"
REQUEST_TIMEOUT = 20
DELAY_BETWEEN_REQUESTS = 1.5
MAX_SITEMAP_DEPTH = 3

log = logging.getLogger("horgaszar")

AVAILABILITY_MAP = {
    "instock": "raktáron", "in_stock": "raktáron", "limitedavailability": "korlátozottan",
    "outofstock": "nincs raktáron", "out_of_stock": "nincs raktáron", "soldout": "nincs raktáron",
    "preorder": "előrendelhető", "backorder": "rendelhető", "discontinued": "megszűnt",
}
AVAILABILITY_KEYWORDS = [
    ("nincs raktáron", "nincs raktáron"), ("elfogyott", "nincs raktáron"),
    ("nem rendelhető", "nincs raktáron"), ("nem elérhető", "nincs raktáron"),
    ("előrendel", "előrendelhető"), ("raktáron", "raktáron"), ("készleten", "raktáron"),
    ("ma feladjuk", "raktáron"), ("rendelhető", "rendelhető"), ("beszerezhető", "rendelhető"),
]

# Ezekre a linkekre sosem mondjuk, hogy termékoldal
NON_PRODUCT_HINTS = re.compile(
    r"(shop_cart|shop_reg|shop_order|shop_contact|kosar|cart|checkout|login|belepes|regisztr|"
    r"account|wishlist|kedvenc|compare|osszehasonlit|route=information|route=account|"
    r"aszf|adatvedelem|kapcsolat|szallitas|blog|hirek|\.(jpe?g|png|gif|webp|pdf)$)",
    re.I,
)
PRODUCT_CONTAINER_HINTS = re.compile(r"(product|termek|artlist|item|card|snapshot)", re.I)


# ---------------------------------------------------------------------------
# Adatmodell
# ---------------------------------------------------------------------------

@dataclass
class Product:
    shop: str
    shop_name: str
    category: str | None
    url: str
    name: str | None = None
    variant: str | None = None
    sku: str | None = None
    brand: str | None = None
    price: float | None = None
    original_price: float | None = None
    currency: str | None = None
    availability: str | None = None
    length_m: float | None = None
    casting_weight_g: str | None = None
    image: str | None = None
    status: str = "ok"              # ok / nincs_adat / hiba
    error: str | None = None
    scraped_at: str = ""


# ---------------------------------------------------------------------------
# HTTP
# ---------------------------------------------------------------------------

def make_session() -> requests.Session:
    session = requests.Session()
    retry = Retry(total=3, backoff_factor=1.0, status_forcelist=(429, 500, 502, 503, 504),
                  allowed_methods=("GET",), respect_retry_after_header=True)
    adapter = HTTPAdapter(max_retries=retry)
    session.mount("http://", adapter)
    session.mount("https://", adapter)
    session.headers.update({
        "User-Agent": USER_AGENT,
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        "Accept-Language": "hu-HU,hu;q=0.9,en;q=0.8",
    })
    return session


class Fetcher:
    """Letöltés udvarias várakozással és robots.txt ellenőrzéssel."""

    def __init__(self, session: requests.Session, delay: float, respect_robots: bool = True):
        self.session = session
        self.delay = delay
        self.respect_robots = respect_robots
        self._robots: dict[str, RobotFileParser | None] = {}
        self._last = 0.0

    def _wait(self):
        elapsed = time.monotonic() - self._last
        if elapsed < self.delay:
            time.sleep(self.delay - elapsed)
        self._last = time.monotonic()

    def allowed(self, url: str) -> bool:
        if not self.respect_robots:
            return True
        p = urlparse(url)
        base = f"{p.scheme}://{p.netloc}"
        if base not in self._robots:
            rp = RobotFileParser()
            try:
                r = self.session.get(base + "/robots.txt", timeout=REQUEST_TIMEOUT)
                if r.status_code == 200:
                    rp.parse(r.text.splitlines())
                    self._robots[base] = rp
                else:
                    self._robots[base] = None
            except requests.RequestException:
                self._robots[base] = None
        rp = self._robots[base]
        return True if rp is None else rp.can_fetch(USER_AGENT, url)

    def get(self, url: str) -> requests.Response:
        if not self.allowed(url):
            raise PermissionError("robots.txt tiltja")
        self._wait()
        resp = self.session.get(url, timeout=REQUEST_TIMEOUT)
        resp.raise_for_status()
        return resp

    def soup(self, url: str) -> tuple[BeautifulSoup, str]:
        resp = self.get(url)
        if not resp.encoding or resp.encoding.lower() == "iso-8859-1":
            resp.encoding = resp.apparent_encoding
        return BeautifulSoup(resp.text, "html.parser"), resp.url


# ---------------------------------------------------------------------------
# Segédfüggvények: ár, elérhetőség, hossz, dobósúly
# ---------------------------------------------------------------------------

PRICE_RE = re.compile(r"(\d[\d\s\xa0.,]*)\s*(?:Ft|HUF|,-)", re.I)


def parse_price(value) -> float | None:
    """'12 990 Ft', '12.990,-', '12990.00', 12990 -> 12990.0"""
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).replace("\xa0", " ").strip()
    m = PRICE_RE.search(text)
    if m:
        text = m.group(1)
    text = re.sub(r"[^\d,.\s]", "", text).strip().replace(" ", "")
    if not text:
        return None
    if "," in text and "." in text:
        if text.rfind(",") > text.rfind("."):
            text = text.replace(".", "").replace(",", ".")
        else:
            text = text.replace(",", "")
    elif "," in text:
        head, _, tail = text.rpartition(",")
        text = f"{head.replace(',', '')}.{tail}" if len(tail) in (1, 2) else text.replace(",", "")
    elif text.count(".") > 1 or (text.count(".") == 1 and len(text.rpartition(".")[2]) == 3):
        text = text.replace(".", "")
    try:
        return float(text.rstrip("."))
    except ValueError:
        return None


def find_prices(text: str) -> list[float]:
    return [p for p in (parse_price(m.group(0)) for m in PRICE_RE.finditer(text or "")) if p]


def normalize_availability(value) -> str | None:
    if not value:
        return None
    raw = str(value).strip()
    key = raw.rsplit("/", 1)[-1].lower().replace(" ", "")
    if key in AVAILABILITY_MAP:
        return AVAILABILITY_MAP[key]
    low = raw.lower()
    for kw, label in AVAILABILITY_KEYWORDS:
        if kw in low:
            return label
    return None


def parse_length_m(text: str | None) -> float | None:
    """'4,20m', '3.60M', '420cm', '390H', '420XH', '12ft' -> méter."""
    if not text:
        return None
    t = text.replace("\xa0", " ")
    candidates = [
        (r"(\d[.,]\d{1,2})\s*[-–/]\s*\d[.,]\d{1,2}\s*m(?![a-zá])", 1.0),   # 3.60-3.90 m (első)
        (r"(\d{3})\s*[-–/]\s*\d{3}\s*cm", 0.01),                          # 330/390cm (első)
        (r"(\d[.,]\d{1,2})\s*m(?![a-zá])", 1.0),                       # 3,60m / 3.9 m
        (r"(\d{3})\s*cm", 0.01),                                        # 390cm
        (r"\b(\d{3})\s*(?:XXH|XH|MH|ML|H|M|L)\b", 0.01),                # Döme 390H / 420XH
        (r"\b(\d{1,2})\s*ft\b", 0.3048),                                # 12ft
    ]
    for pattern, factor in candidates:
        m = re.search(pattern, t, re.I)
        if m:
            val = round(float(m.group(1).replace(",", ".")) * factor, 2)
            if 1.0 <= val <= 7.0:
                return val
    return None


def parse_casting_weight(text: str | None) -> str | None:
    """'20-75gr' -> '20-75', '180G' -> '180', '50-170gr' -> '50-170'."""
    if not text:
        return None
    t = text.replace("\xa0", " ")
    m = re.search(r"(\d{1,3})\s*[-–]\s*(\d{2,3})\s*g(?:r|ramm)?\b", t, re.I)
    if m:
        return f"{m.group(1)}-{m.group(2)}"
    m = re.search(r"(\d{2,3})\s*g(?:r|ramm)?\b", t, re.I)
    return m.group(1) if m else None


def clean(text: str | None) -> str | None:
    if text is None:
        return None
    text = re.sub(r"\s+", " ", str(text)).strip()
    return text or None


# ---------------------------------------------------------------------------
# Termékoldal: adatkinyerés
# ---------------------------------------------------------------------------

def _iter_jsonld(data):
    if isinstance(data, list):
        for item in data:
            yield from _iter_jsonld(item)
    elif isinstance(data, dict):
        yield data
        for key in ("@graph", "hasVariant", "mainEntity", "itemListElement"):
            if key in data:
                yield from _iter_jsonld(data[key])


def _types(obj) -> set[str]:
    t = obj.get("@type")
    return {str(x).lower() for x in (t if isinstance(t, list) else [t]) if x}


def _offer_info(offers) -> dict:
    if isinstance(offers, list):
        offers = offers[0] if offers else {}
    if not isinstance(offers, dict):
        return {}
    price = offers.get("price")
    if offers.get("@type") == "AggregateOffer":
        price = offers.get("lowPrice") or price
    if price is None:
        spec = offers.get("priceSpecification") or {}
        if isinstance(spec, list):
            spec = spec[0] if spec else {}
        price = spec.get("price") if isinstance(spec, dict) else None
    return {
        "price": parse_price(price),
        "currency": offers.get("priceCurrency"),
        "availability": normalize_availability(offers.get("availability")),
    }


def _brand(obj):
    b = obj.get("brand")
    if isinstance(b, list):
        b = b[0] if b else None
    return b.get("name") if isinstance(b, dict) else b


def _image(obj):
    img = obj.get("image")
    if isinstance(img, list):
        img = img[0] if img else None
    return img.get("url") if isinstance(img, dict) else img


def extract_jsonld(soup: BeautifulSoup) -> tuple[dict, list[dict]]:
    """Visszaad: (alapadatok, változatok listája)."""
    products = []
    for script in soup.find_all("script", type="application/ld+json"):
        try:
            data = json.loads(script.string or script.get_text() or "")
        except (json.JSONDecodeError, TypeError):
            continue
        products += [o for o in _iter_jsonld(data) if _types(o) & {"product", "productgroup"}]
    if not products:
        return {}, []

    main = products[0]
    base = {"name": main.get("name"), "brand": _brand(main), "sku": main.get("sku") or main.get("mpn"),
            "image": _image(main), **_offer_info(main.get("offers"))}

    variants = []
    # 1) ProductGroup / több Product ugyanazon az oldalon
    for p in products[1:]:
        info = _offer_info(p.get("offers"))
        if info.get("price") is not None:
            variants.append({"variant": p.get("name"), "sku": p.get("sku"), **info})
    # 2) Egy Product, több Offer (pl. hosszanként)
    offers = main.get("offers")
    if not variants and isinstance(offers, list) and len(offers) > 1:
        for o in offers:
            info = _offer_info(o)
            if info.get("price") is not None:
                variants.append({"variant": o.get("name") or o.get("sku"), "sku": o.get("sku"), **info})
    return base, variants


def extract_meta(soup: BeautifulSoup) -> dict:
    def meta(*names):
        for n in names:
            tag = soup.find("meta", attrs={"property": n}) or soup.find("meta", attrs={"name": n})
            if tag and tag.get("content"):
                return tag["content"].strip()
        return None

    def itemprop(prop):
        tag = soup.find(attrs={"itemprop": prop})
        if not tag:
            return None
        return tag.get("content") or tag.get("href") or tag.get_text(" ", strip=True) or None

    h1 = soup.find("h1")
    return {
        "name": itemprop("name") or (h1.get_text(" ", strip=True) if h1 else None) or meta("og:title"),
        "price": parse_price(itemprop("price") or meta("product:price:amount", "og:price:amount")),
        "currency": itemprop("priceCurrency") or meta("product:price:currency", "og:price:currency"),
        "availability": normalize_availability(itemprop("availability") or meta("product:availability")),
        "brand": meta("product:brand"),
        "sku": itemprop("sku"),
        "image": meta("og:image"),
    }


def extract_heuristic(soup: BeautifulSoup) -> dict:
    """Utolsó mentsvár: ár és elérhetőség keresése 'price'/'ar' osztályú elemekben."""
    out = {}
    for el in soup.select("[class*=price], [class*=Price], [id*=price], [class*=_ar], [class*=-ar]"):
        if el.find_parent(["del", "s", "strike"]) or el.name in ("del", "s", "strike"):
            continue
        cls = " ".join(el.get("class", [])).lower()
        if any(x in cls for x in ("old", "regi", "original", "crossed", "brutto-old", "list")):
            continue
        prices = find_prices(el.get_text(" ", strip=True))
        if prices:
            out["price"] = prices[-1]
            break
    for el in soup.select("[class*=stock], [class*=keszlet], [class*=avail], [class*=raktar]"):
        av = normalize_availability(el.get_text(" ", strip=True))
        if av:
            out["availability"] = av
            break
    old = soup.select_one("del, s, .price-old, [class*=old-price], [class*=oldprice]")
    if old:
        p = find_prices(old.get_text(" ", strip=True))
        if p:
            out["original_price"] = p[0]
    return out


def extract_fallback_name(soup: BeautifulSoup) -> dict:
    """Ha nincs <h1>: első h2–h6 címsor, ami a <title>-ben is szerepel, különben a <title> utolsó része."""
    title = clean(soup.title.get_text()) if soup.title else None
    for h in soup.find_all(["h2", "h3", "h4", "h5", "h6"]):
        txt = clean(h.get_text(" ", strip=True))
        if txt and title and txt.lower() in title.lower():
            return {"name": txt}
    if title:
        return {"name": re.split(r"\s+[-|»]\s+", title)[-1]}
    return {}


def extract_custom(soup: BeautifulSoup, selectors: dict) -> dict:
    if not selectors:
        return {}

    def text(key):
        css = selectors.get(key)
        el = soup.select_one(css) if css else None
        return el.get_text(" ", strip=True) if el else None

    return {
        "name": text("name"),
        "price": parse_price(text("price")),
        "original_price": parse_price(text("original_price")),
        "availability": normalize_availability(text("availability")),
        "brand": text("brand"),
        "sku": text("sku"),
    }


SKIP_COLUMNS = re.compile(r"^(kép|kosár|kosárba|mennyiség|db|akció|vásárlás|)$", re.I)
SKU_COLUMNS = re.compile(r"(cikkszám|cikkszam|kód|kod|sku|azonosító)", re.I)


def extract_variant_table(soup: BeautifulSoup) -> list[dict]:
    """Táblázatban felsorolt változatok (pl. PecaPlaza: hossz / dobósúly / ár soronként)."""
    best: list[dict] = []
    for table in soup.find_all("table"):
        rows = table.find_all("tr")
        if len(rows) < 2:
            continue
        header_cells = rows[0].find_all(["th", "td"])
        headers = [clean(c.get_text(" ", strip=True)) or "" for c in header_cells]
        if find_prices(" ".join(headers)):          # az első sor már adat, nincs fejléc
            headers, data_rows = [], rows
        else:
            data_rows = rows[1:]

        variants = []
        for tr in data_rows:
            cells = tr.find_all(["td", "th"])
            if not cells:
                continue
            price_idx, price, original = None, None, None
            for i, cell in enumerate(cells):
                txt = cell.get_text(" ", strip=True)
                prices = find_prices(txt)
                if not prices:
                    continue
                struck = cell.find(["del", "s", "strike"])
                struck_prices = find_prices(struck.get_text(" ", strip=True)) if struck else []
                current = [p for p in prices if p not in struck_prices] or prices
                price_idx, price = i, current[-1]
                original = struck_prices[0] if struck_prices else (prices[0] if len(prices) > 1 else None)
                break
            if price is None:
                continue

            attrs, sku = [], None
            for i, cell in enumerate(cells):
                if i == price_idx:
                    continue
                val = clean(cell.get_text(" ", strip=True))
                head = headers[i] if i < len(headers) else ""
                if not val or SKIP_COLUMNS.match(head or "") or val.lower() in ("kosárba", "- +", "-+"):
                    continue
                if SKU_COLUMNS.search(head or "") and not sku:
                    sku = val
                    continue
                attrs.append(f"{head}: {val}" if head else val)
            variants.append({
                "variant": "; ".join(attrs) or None, "sku": sku, "price": price,
                "original_price": original if original and original > price else None,
                "availability": normalize_availability(tr.get_text(" ", strip=True)),
            })
        if len(variants) > len(best):
            best = variants
    return best


def scrape_product(fetcher: Fetcher, url: str, shop: str, cfg: dict,
                   name_hint: str | None = None) -> list[Product]:
    now = datetime.now().isoformat(timespec="seconds")
    base = Product(shop=shop, shop_name=cfg.get("name", shop), category=cfg.get("category"),
                   url=url, scraped_at=now)
    try:
        soup, final_url = fetcher.soup(url)
    except Exception as exc:
        return [replace(base, status="hiba", error=f"{type(exc).__name__}: {exc}"[:300])]

    jsonld, jsonld_variants = extract_jsonld(soup)
    data: dict = {}
    for source in (extract_custom(soup, cfg.get("selectors", {})), jsonld, extract_meta(soup),
                   extract_heuristic(soup), {"name": name_hint}, extract_fallback_name(soup)):
        for k, v in source.items():
            if data.get(k) in (None, "") and v not in (None, ""):
                data[k] = v

    product = replace(base, url=final_url, **{k: v for k, v in data.items() if k in
                                               {f.name for f in fields(Product)}})
    product.name = clean(product.name)
    if product.price is not None and not product.currency:
        product.currency = "HUF"

    variants = extract_variant_table(soup) if cfg.get("variant_table") else []
    variants = variants or jsonld_variants

    results = []
    if variants:
        for v in variants:
            # az eredeti árat mindig a változat sorából vesszük (None is felülír)
            p = replace(product, **{k: val for k, val in v.items()
                                    if val is not None or k == "original_price"})
            p.currency = p.currency or "HUF"
            results.append(p)
    else:
        results.append(product)

    for p in results:
        text = " ".join(filter(None, [p.name, p.variant]))
        p.length_m = p.length_m or parse_length_m(p.variant) or parse_length_m(p.name)
        p.casting_weight_g = p.casting_weight_g or parse_casting_weight(p.variant) or parse_casting_weight(text)
        if p.original_price is not None and p.price is not None and p.original_price <= p.price:
            p.original_price = None
        if not p.name or p.price is None:
            p.status, p.error = "nincs_adat", "név vagy ár nem található"
    return results


# ---------------------------------------------------------------------------
# Kategóriaoldalak bejárása (lapozással)
# ---------------------------------------------------------------------------

def _same_site(url: str, domain: str) -> bool:
    return urlparse(url).netloc.removeprefix("www.") == domain


def find_product_links(soup: BeautifulSoup, page_url: str, domain: str, cfg: dict) -> dict[str, str]:
    """Visszaad: {termék_url: link_szöveg}"""
    pattern = cfg.get("product_url_pattern")
    rx = re.compile(pattern, re.I) if pattern else None
    sel = (cfg.get("selectors") or {}).get("product_link")
    anchors = soup.select(sel) if sel else soup.find_all("a", href=True)
    start_urls = {urldefrag(u)[0].rstrip("/") for u in cfg.get("start_urls", [])}

    links: dict[str, str] = {}
    for a in anchors:
        href = a.get("href")
        if not href or href.startswith(("#", "javascript:", "mailto:", "tel:")):
            continue
        url = urldefrag(urljoin(page_url, href))[0]
        if not _same_site(url, domain) or NON_PRODUCT_HINTS.search(url):
            continue
        if url.rstrip("/") in start_urls or url.rstrip("/") == page_url.rstrip("/"):
            continue
        if rx:
            if not rx.search(url):
                continue
        elif not sel:
            # Heurisztika: a link termékkártyán belül van, és nem fejlécben/menüben/láblécben
            if a.find_parent(["nav", "header", "footer"]):
                continue
            parent_ok = False
            for parent in a.parents:
                if not isinstance(parent, Tag) or parent.name in ("body", "html"):
                    break
                ident = " ".join(parent.get("class", [])) + " " + (parent.get("id") or "")
                if re.search(r"(menu|nav|breadcrumb|filter|szuro|footer|header|banner|brand)", ident, re.I):
                    parent_ok = False
                    break
                if PRODUCT_CONTAINER_HINTS.search(ident):
                    parent_ok = True
            if not parent_ok:
                continue
            if "?" in url and "route=product" not in url:
                continue
        text = clean(a.get_text(" ", strip=True)) or clean(a.get("title")) or ""
        if url not in links or len(text) > len(links[url]):
            links[url] = text
    return links


def next_page_url(soup: BeautifulSoup, current: str, start: str, page_no: int, mode: str, param: str = "page") -> str | None:
    if mode == "unas_spl":                   # UNAS márkaoldal: /spl/ID/Marka -> /spl/ID,2/Marka?infinite_scroll=1&ajax_nodesign=1
        m = re.match(r"(https?://[^/]+/spl/)(\d+)(?:,\d+)?/([^?]+)", start)
        return f"{m.group(1)}{m.group(2)},{page_no}/{m.group(3)}?infinite_scroll=1&ajax_nodesign=1" if m else None
    if mode == "comma":                      # UNAS: /Kategoria,2
        return f"{start.rstrip('/')},{page_no}"
    if mode == "query_page":                 # Shoprenter: ?page=2
        p = urlparse(start)
        q = parse_qs(p.query)
        q[param] = [str(page_no)]
        return urlunparse(p._replace(query=urlencode(q, doseq=True)))
    # auto
    link = soup.find(["link", "a"], rel=lambda r: r and "next" in (r if isinstance(r, list) else [r]))
    if link and link.get("href"):
        return urljoin(current, link["href"])
    for a in soup.find_all("a", href=True):
        txt = a.get_text(" ", strip=True).lower()
        if txt in ("»", "›", ">", "következő", "következő oldal", "tovább", "next") or \
           "next" in " ".join(a.get("class", [])).lower():
            return urljoin(current, a["href"])
    for a in soup.find_all("a", href=True):           # számozott lapozó: a következő szám
        if a.get_text(strip=True) == str(page_no):
            return urljoin(current, a["href"])
    return None


def crawl_category(fetcher: Fetcher, start: str, domain: str, cfg: dict) -> dict[str, str]:
    mode = cfg.get("pagination", "auto")
    max_pages = int(cfg.get("max_pages", 20))
    all_links: dict[str, str] = {}
    url, page_no, seen_pages = start, 1, set()
    while url and page_no <= max_pages and url not in seen_pages:
        seen_pages.add(url)
        try:
            soup, final = fetcher.soup(url)
        except Exception as exc:
            log.warning("  Kategóriaoldal hiba (%s): %s", url, exc)
            break
        links = find_product_links(soup, final, domain, cfg)
        new = {u: t for u, t in links.items() if u not in all_links}
        log.info("  %d. oldal: %d termék-link (%d új) – %s", page_no, len(links), len(new), url)
        if not new:
            break
        all_links.update(new)
        page_no += 1
        url = next_page_url(soup, final, start, page_no, mode, cfg.get("page_param", "page"))
    return all_links


# ---------------------------------------------------------------------------
# Sitemap / linklista
# ---------------------------------------------------------------------------

def _strip_ns(tag: str) -> str:
    return tag.split("}", 1)[-1]


def parse_sitemap(fetcher: Fetcher, url: str, depth: int = 0) -> list[str]:
    if depth > MAX_SITEMAP_DEPTH:
        return []
    try:
        content = fetcher.get(url).content
        if url.endswith(".gz") or content[:2] == b"\x1f\x8b":
            content = gzip.decompress(content)
        root = ET.fromstring(content)
    except Exception as exc:
        log.error("Sitemap nem olvasható (%s): %s", url, exc)
        return []
    urls = []
    for child in root:
        loc = next((c.text.strip() for c in child if _strip_ns(c.tag) == "loc" and c.text), None)
        if not loc:
            continue
        if _strip_ns(root.tag) == "sitemapindex":
            urls += parse_sitemap(fetcher, loc, depth + 1)
        else:
            urls.append(loc)
    return urls


def read_url_list(path: str) -> list[str]:
    with open(path, encoding="utf-8") as fh:
        return [ln.strip() for ln in fh if ln.strip() and not ln.lstrip().startswith("#")]


# ---------------------------------------------------------------------------
# Mentés
# ---------------------------------------------------------------------------

COLUMNS = [f.name for f in fields(Product)]


def save_csv(products: list[Product], path: str) -> None:
    with open(path, "w", newline="", encoding="utf-8-sig") as fh:
        w = csv.DictWriter(fh, fieldnames=COLUMNS, delimiter=";")
        w.writeheader()
        for p in products:
            w.writerow(asdict(p))
    log.info("CSV mentve: %s", path)


def save_excel(products: list[Product], path: str) -> None:
    try:
        from openpyxl import Workbook
        from openpyxl.styles import Font, PatternFill
        from openpyxl.utils import get_column_letter
    except ImportError:
        log.error("Excel mentéshez: pip install openpyxl")
        return
    wb = Workbook()
    ws = wb.active
    ws.title = "termekek"
    ws.append(COLUMNS)
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="1F4E78")
    for p in products:
        ws.append([getattr(p, c) for c in COLUMNS])
    for col in ("price", "original_price"):
        letter = get_column_letter(COLUMNS.index(col) + 1)
        for cell in ws[letter][1:]:
            cell.number_format = '# ##0" Ft"'
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    for i, col in enumerate(COLUMNS, 1):
        width = max([len(col)] + [len(str(getattr(p, col) or "")) for p in products])
        ws.column_dimensions[get_column_letter(i)].width = min(width + 2, 55)

    # Összesítő lap boltonként
    ws2 = wb.create_sheet("osszesito")
    ws2.append(["bolt", "sorok", "ok", "nincs_adat", "hiba", "min ár", "max ár"])
    for shop in sorted({p.shop_name for p in products}):
        rows = [p for p in products if p.shop_name == shop]
        prices = [p.price for p in rows if p.price is not None]
        ws2.append([shop, len(rows), sum(p.status == "ok" for p in rows),
                    sum(p.status == "nincs_adat" for p in rows), sum(p.status == "hiba" for p in rows),
                    min(prices) if prices else None, max(prices) if prices else None])
    wb.save(path)
    log.info("Excel mentve: %s", path)


# ---------------------------------------------------------------------------
# Diagnosztika
# ---------------------------------------------------------------------------

def diagnose(fetcher: Fetcher, url: str, shops: dict) -> None:
    domain = urlparse(url).netloc.removeprefix("www.")
    cfg = shops.get(domain, {"name": domain})
    soup, final = fetcher.soup(url)
    fname = f"diagnose_{domain.replace('.', '_')}.html"
    with open(fname, "w", encoding="utf-8") as fh:
        fh.write(str(soup))
    print(f"\n=== {final}\nHTML elmentve: {fname}")
    ld_types = []
    for s in soup.find_all("script", type="application/ld+json"):
        try:
            ld_types += sorted(_types(o) for o in _iter_jsonld(json.loads(s.string or "")))
        except Exception:
            ld_types.append({"(hibás JSON)"})
    print("JSON-LD típusok:", ld_types or "nincs")
    print("itemprop=price:", bool(soup.find(attrs={"itemprop": "price"})),
          "| og/product meta ár:", bool(soup.find("meta", attrs={"property": re.compile("price")})))
    vt = extract_variant_table(soup)
    print(f"Változattábla: {len(vt)} sor", *(f"\n   {v}" for v in vt[:5]))
    links = find_product_links(soup, final, domain, cfg)
    print(f"Termék-linkek (ha ez kategóriaoldal): {len(links)}")
    for u, t in list(links.items())[:8]:
        print(f"   {t[:50]:50} {u}")
    print("Következő oldal (auto):", next_page_url(soup, final, url, 2, "auto"))
    print("\nTermékoldalként értelmezve:")
    for p in scrape_product(fetcher, url, domain, cfg):
        print("  ", {k: v for k, v in asdict(p).items() if v not in (None, "")})


# ---------------------------------------------------------------------------
# Fő program
# ---------------------------------------------------------------------------

def load_shops(path: str | None) -> dict:
    if not path:
        return {}
    with open(path, encoding="utf-8") as fh:
        data = json.load(fh)
    return {k: v for k, v in data.items() if not k.startswith("_")}


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="HorgászÁr.hu termékadat-gyűjtő")
    src = ap.add_mutually_exclusive_group(required=True)
    src.add_argument("--shops", metavar="JSON", help="boltkonfiguráció (pl. shops.json)")
    src.add_argument("--sitemap", nargs="+", metavar="URL")
    src.add_argument("--urls", metavar="FÁJL")
    src.add_argument("--diagnose", metavar="URL", help="egy oldal elemzése hibakereséshez")
    ap.add_argument("--config", default="shops.json", help="boltkonfig a --diagnose/--sitemap/--urls módhoz")
    ap.add_argument("--only", nargs="+", metavar="DOMAIN", help="csak ezek a boltok (pl. pecaplaza.hu)")
    ap.add_argument("--filter", metavar="REGEX", help="csak az erre illeszkedő URL-ek")
    ap.add_argument("--limit", type=int, help="boltonként legfeljebb ennyi termékoldal (teszthez)")
    ap.add_argument("--output", help="kimeneti CSV (alap: termekek_DÁTUM.csv)")
    ap.add_argument("--excel", action="store_true", help="Excel (.xlsx) fájlt is készít")
    ap.add_argument("--delay", type=float, default=DELAY_BETWEEN_REQUESTS)
    ap.add_argument("--no-robots", action="store_true")
    ap.add_argument("-v", "--verbose", action="store_true")
    args = ap.parse_args(argv)

    logging.basicConfig(level=logging.DEBUG if args.verbose else logging.INFO,
                        format="%(asctime)s %(levelname)s %(message)s", datefmt="%H:%M:%S")
    fetcher = Fetcher(make_session(), args.delay, respect_robots=not args.no_robots)

    try:
        shops = load_shops(args.shops or args.config)
    except FileNotFoundError:
        if args.shops:
            log.error("Nem található: %s", args.shops)
            return 1
        shops = {}

    if args.diagnose:
        diagnose(fetcher, args.diagnose, shops)
        return 0

    # (bolt, url) párok összegyűjtése
    jobs: list[tuple[str, str, str | None]] = []
    if args.shops:
        for domain, cfg in shops.items():
            if args.only and domain not in args.only:
                continue
            log.info("== %s (%s)", cfg.get("name", domain), cfg.get("engine", "?"))
            links: dict[str, str] = {}
            for start in cfg.get("start_urls", []):
                links.update(crawl_category(fetcher, start, domain, cfg))
            must = cfg.get("name_must_contain")
            if must:
                rx = re.compile(must, re.I)
                links = {u: t for u, t in links.items() if rx.search(u) or rx.search(t)}
            urls = list(links)
            if args.filter:
                urls = [u for u in urls if re.search(args.filter, u)]
            if args.limit:
                urls = urls[: args.limit]
            log.info("   %d termékoldal lesz feldolgozva", len(urls))
            jobs += [(domain, u, links.get(u)) for u in urls]
    else:
        urls = []
        if args.sitemap:
            for sm in args.sitemap:
                log.info("Sitemap: %s", sm)
                urls += parse_sitemap(fetcher, sm)
        else:
            urls = read_url_list(args.urls)
        urls = list(dict.fromkeys(urls))
        if args.filter:
            urls = [u for u in urls if re.search(args.filter, u)]
        if args.limit:
            urls = urls[: args.limit]
        jobs = [(urlparse(u).netloc.removeprefix("www."), u, None) for u in urls]

    if not jobs:
        log.error("Nincs feldolgozandó termékoldal. Tipp: futtasd a --diagnose módot egy kategóriaoldalra.")
        return 1

    products: list[Product] = []
    for i, (domain, url, hint) in enumerate(jobs, 1):
        cfg = shops.get(domain, {"name": domain})
        rows = scrape_product(fetcher, url, domain, cfg, name_hint=hint)
        products += rows
        for p in rows:
            log.info("[%d/%d] %-10s %-45s %-30s %s", i, len(jobs), p.status, (p.name or "-")[:45],
                     (p.variant or "")[:30], f"{p.price:,.0f} Ft".replace(",", " ") if p.price else "-")

    out = args.output or f"termekek_{datetime.now():%Y%m%d_%H%M}.csv"
    save_csv(products, out)
    if args.excel:
        save_excel(products, out.rsplit(".", 1)[0] + ".xlsx")
    log.info("Kész: %d sor – %d ok, %d adat nélküli, %d hibás.", len(products),
             sum(p.status == "ok" for p in products), sum(p.status == "nincs_adat" for p in products),
             sum(p.status == "hiba" for p in products))
    return 0


if __name__ == "__main__":
    sys.exit(main())
