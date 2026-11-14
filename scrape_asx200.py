#!/usr/bin/env python3
"""
ASX 200 company directory scraper
---------------------------------
Reads the constituents table on the Wikipedia page "S&P/ASX 200", opens the
Wikipedia article of every company and pulls the infobox fields (industry,
founded, head office, employees, revenue, website) into one clean table, with
the source link on every row.

When the index table has no link for a company (or the link points to the
wrong kind of page), the scraper tries the company name as an article title
and only accepts the page if its infobox lists the same code.

Every row gets a status: OK, Check (the two pages disagree, e.g. on the
head office or the code) or Not found (no company page). A second list
shows employee and revenue figures that are missing, undated or 3+ years old.

Polite by design:
  * checks robots.txt before downloading any page
  * identifies itself with a User-Agent that includes a contact URL
  * waits --delay seconds between network requests (default 1.0)
  * retries 429 and 5xx answers with exponential backoff (honours Retry-After)
  * caches every page on disk (404s too), so a re-run does not hit the site again
  * public pages only: no logins, no personal data (people fields are skipped)

Usage:
    python scrape_asx200.py             # uses the cache when it exists
    python scrape_asx200.py --refresh   # downloads every page again

Outputs (folder out/):
    asx200_directory.csv   one row per index constituent, with a status (OK / Check / Not found)
    review.csv             rows marked Check or Not found, with the reason and what to check
    figures_to_update.csv  employee and revenue figures that are missing, undated or 3+ years old
    scrape_log.json        requests, cache hits, retries, timings and coverage
"""
from __future__ import annotations

import argparse
import hashlib
import json
import re
import time
from datetime import datetime, timezone
from pathlib import Path
from urllib import robotparser
from urllib.parse import quote, urlparse

import pandas as pd
import requests
from bs4 import BeautifulSoup

LIST_URL = "https://en.wikipedia.org/wiki/S%26P/ASX_200"
WIKI = "https://en.wikipedia.org/wiki/"
USER_AGENT = "StuardOre-portfolio-scraper/1.0 (https://github.com/StuardOre)"
BASE = Path(__file__).resolve().parent
CACHE = BASE / "cache"
OUT = BASE / "out"
RETRY_STATUS = {429, 500, 502, 503, 504}
OLD_BEFORE = 2023  # employee or revenue figures from 2022 or earlier are "old" (3+ years)


# --------------------------------------------------------------------------- #
# Fetching
# --------------------------------------------------------------------------- #
class NotFound(Exception):
    """The page does not exist (HTTP 404)."""


class Fetcher:
    """requests.Session with robots.txt check, rate limit, retries and a disk cache."""

    def __init__(self, delay: float = 1.0, refresh: bool = False, max_tries: int = 4):
        self.session = requests.Session()
        self.session.headers.update({"User-Agent": USER_AGENT, "Accept-Language": "en"})
        self.delay, self.refresh, self.max_tries = delay, refresh, max_tries
        self.robots: dict[str, robotparser.RobotFileParser] = {}
        self.last_request = 0.0
        self.stats = {"network_requests": 0, "cache_hits": 0, "retries": 0,
                      "not_found_404": 0, "blocked_by_robots": 0, "status_codes": {}}
        CACHE.mkdir(exist_ok=True)

    def _allowed(self, url: str) -> bool:
        parts = urlparse(url)
        host = f"{parts.scheme}://{parts.netloc}"
        if host not in self.robots:
            rp = robotparser.RobotFileParser()
            resp = self.session.get(host + "/robots.txt", timeout=30)
            rp.parse(resp.text.splitlines())
            self.robots[host] = rp
        return self.robots[host].can_fetch(USER_AGENT, url)

    def get(self, url: str) -> tuple[str, dict]:
        key = hashlib.sha1(url.encode("utf-8")).hexdigest()[:16]
        html_path, meta_path = CACHE / f"{key}.html", CACHE / f"{key}.json"
        if meta_path.exists() and not self.refresh:
            self.stats["cache_hits"] += 1
            meta = json.loads(meta_path.read_text())
            if meta["status"] == 404:
                self.stats["not_found_404"] += 1
                raise NotFound(url)
            return html_path.read_text(encoding="utf-8"), meta

        if not self._allowed(url):
            self.stats["blocked_by_robots"] += 1
            raise PermissionError(f"robots.txt disallows {url}")

        for attempt in range(1, self.max_tries + 1):
            wait = self.delay - (time.monotonic() - self.last_request)
            if wait > 0:
                time.sleep(wait)
            started = time.monotonic()
            resp = self.session.get(url, timeout=30)
            self.last_request = time.monotonic()
            self.stats["network_requests"] += 1
            code = str(resp.status_code)
            self.stats["status_codes"][code] = self.stats["status_codes"].get(code, 0) + 1

            if resp.status_code in RETRY_STATUS and attempt < self.max_tries:
                self.stats["retries"] += 1
                retry_after = resp.headers.get("Retry-After", "")
                time.sleep(int(retry_after) if retry_after.isdigit() else 2 ** attempt)
                continue

            meta = {"url": url, "final_url": resp.url, "status": resp.status_code,
                    "attempts": attempt, "seconds": round(self.last_request - started, 2),
                    "fetched_at": datetime.now(timezone.utc).isoformat(timespec="seconds")}
            if resp.status_code == 404:
                meta_path.write_text(json.dumps(meta))
                self.stats["not_found_404"] += 1
                raise NotFound(url)
            resp.raise_for_status()
            html_path.write_text(resp.text, encoding="utf-8")
            meta_path.write_text(json.dumps(meta))
            return resp.text, meta
        raise RuntimeError(f"giving up on {url}")


# --------------------------------------------------------------------------- #
# Cleaning helpers
# --------------------------------------------------------------------------- #
NOISE = ("style, link, script, sup.reference, sup.mw-ref, span.noprint, "
         "img, .mw-file-element, [style*='display:none'], [style*='display: none']")


def clean_text(node) -> str:
    """Visible text of an HTML cell, without footnotes, hidden spans or icons."""
    if node is None:
        return ""
    for bad in node.select(NOISE):
        bad.decompose()
    for br in node.find_all("br"):
        br.replace_with("; ")
    for li in node.find_all("li"):
        li.insert_after("; ")
    text = node.get_text(" ", strip=True).replace(" ", " ")
    text = re.sub(r"\[\s*(?:\d+|[a-z]|note \d+|citation needed)\s*\]", "", text)
    text = re.sub(r"\s+", " ", text)
    text = re.sub(r"\s+([,.;:)])", r"\1", text).replace("( ", "(")
    text = re.sub(r"(;\s*)+", "; ", text)
    return text.strip(" ;,")


def to_int(text: str) -> int | None:
    digits = re.sub(r"[^\d]", "", text or "")
    return int(digits) if digits else None


def first_year(text: str) -> int | None:
    match = re.search(r"\b(1[6-9]\d{2}|20[0-2]\d)\b", text or "")
    return int(match.group(1)) if match else None


def figure_year(text: str) -> int | None:
    """Year a figure refers to: '(2025)', '(FY2025)', '(FY 25)', '(June 2024)'."""
    if not text:
        return None
    fy = re.search(r"\bFY\s?'?(\d{2,4})\b", text, re.I)
    if fy:
        value = int(fy.group(1))
        return value + 2000 if value < 100 else value
    inside = " ".join(re.findall(r"\(([^)]*)\)", text))
    return first_year(inside) or first_year(text)


def parse_employees(text: str) -> int | None:
    """'91,304 (2025)' -> 91304. Ignores numbers inside brackets (years)."""
    if not text:
        return None
    text = re.sub(r"(?<=\d)[   ](?=\d{3}\b)", ",", text)  # '40 000' -> '40,000'
    outside = re.sub(r"\([^)]*\)", " ", text)
    match = re.search(r"(\d{1,3}(?:,\d{3})+|\d+(?:\.\d+)?)\s*(k|thousand)?\b", outside, re.I)
    if not match:
        return None
    value = float(match.group(1).replace(",", ""))
    if match.group(2):
        value *= 1000
    if 1900 <= value <= 2030 and "," not in match.group(1) and not match.group(2):
        return None  # a bare year, not a head count
    return int(round(value))


# Longer symbols first, so 'US$' is not read as 'S$' and 'CA$' is not read as 'A$'.
CURRENCIES = [("AU$", "AUD"), ("US$", "USD"), ("NZ$", "NZD"), ("CA$", "CAD"), ("HK$", "HKD"),
              ("A$", "AUD"), ("C$", "CAD"), ("S$", "SGD"), ("AUD", "AUD"), ("USD", "USD"),
              ("NZD", "NZD"), ("CAD", "CAD"), ("GBP", "GBP"), ("EUR", "EUR"), ("PGK", "PGK"),
              ("£", "GBP"), ("€", "EUR"), ("$", "$")]
UNITS = {"trillion": 1_000_000, "billion": 1_000, "bn": 1_000, "b": 1_000,
         "million": 1, "mn": 1, "m": 1, "thousand": 0.001, "k": 0.001}


def parse_money(text: str) -> tuple[float | None, str | None]:
    """'US$17.9 billion (2025)' -> (17900.0, 'USD'): amount in millions + currency."""
    if not text:
        return None, None
    first = text.split(";")[0]
    currency = next((code for symbol, code in CURRENCIES if symbol in first), None)
    match = re.search(r"(\d[\d,]*(?:\.\d+)?)\s*(trillion|billion|bn|b|million|mn|m|thousand|k)?\b",
                      re.sub(r"\([^)]*\)", " ", first), re.I)
    if not match:
        return None, currency
    amount = float(match.group(1).replace(",", ""))
    unit = (match.group(2) or "").lower()
    if unit:
        millions = amount * UNITS[unit]
    elif amount >= 100_000:
        millions = amount / 1_000_000  # full number written out
    else:
        return None, currency  # no unit and a small number: too ambiguous to guess
    return round(millions, 1), currency


def website_domain(url: str) -> str | None:
    if not url:
        return None
    if "//" not in url:
        url = "https://" + url
    host = urlparse(url).netloc.lower().split(":")[0]
    return (host[4:] if host.startswith("www.") else host) or None


AU_STATES = {"nsw": ["new south wales", "nsw"], "vic": ["victoria", "vic"], "qld": ["queensland", "qld"],
             "wa": ["western australia", "wa"], "sa": ["south australia", "sa"], "tas": ["tasmania", "tas"],
             "act": ["australian capital territory", "act"], "nt": ["northern territory", "nt"]}
AU_CITIES = {"sydney": "nsw", "newcastle": "nsw", "wollongong": "nsw", "bega": "nsw",
             "melbourne": "vic", "geelong": "vic", "bendigo": "vic", "ballarat": "vic",
             "brisbane": "qld", "gold coast": "qld", "townsville": "qld", "toowoomba": "qld",
             "perth": "wa", "adelaide": "sa", "hobart": "tas", "launceston": "tas",
             "canberra": "act", "darwin": "nt"}
US_STATES = ("alabama alaska arizona arkansas california colorado connecticut delaware florida georgia "
             "hawaii idaho illinois indiana iowa kansas kentucky louisiana maine maryland massachusetts "
             "michigan minnesota mississippi missouri montana nebraska nevada ohio oklahoma oregon "
             "pennsylvania tennessee texas utah vermont virginia washington wisconsin wyoming").split()
US_STATES += ["new hampshire", "new jersey", "new mexico", "new york", "north carolina", "north dakota",
              "rhode island", "south carolina", "south dakota", "west virginia", "u s", "usa"]
COUNTRY_ALIASES = {"united states": US_STATES, "united kingdom": ["england", "scotland", "wales", "u k", "uk"]}


def words(text: str) -> str:
    """Lower case, letters only, single spaces, padded so ' term ' matches whole words."""
    return " " + re.sub(r"\s+", " ", re.sub(r"[^a-z]+", " ", (text or "").lower())).strip() + " "


def mentions(text: str, term: str) -> bool:
    return f" {term} " in text


def head_office_conflict(hq_list: str, hq_infobox: str) -> bool:
    """True when the article's head office is in another state or country than the index table says.

    A suburb in the same state is not a conflict ('Sydney' vs 'Alexandria, New South Wales').
    """
    box = words(hq_infobox)
    if not hq_list or not box.strip():
        return False
    place, _, country = hq_list.partition(",")
    cities = [words(c).strip() for c in place.split("/") if words(c).strip()]
    if any(mentions(box, c) for c in cities):
        return False
    if country.strip():  # head office abroad, e.g. 'Las Vegas, United States'
        name = words(country).strip()
        return not (mentions(box, name) or any(mentions(box, a) for a in COUNTRY_ALIASES.get(name, [])))
    state = next((AU_CITIES[c] for c in cities if c in AU_CITIES), None)
    if state and any(mentions(box, alias) for alias in AU_STATES[state]):
        return False  # same state, different suburb
    other_state = any(mentions(box, alias) for s, aliases in AU_STATES.items() if s != state for alias in aliases)
    other_city = any(mentions(box, c) for c in AU_CITIES)
    if other_state or other_city:
        return True
    return not mentions(box, "australia")  # no Australian place at all means the article puts it abroad


# --------------------------------------------------------------------------- #
# Parsing
# --------------------------------------------------------------------------- #
def parse_list(html: str) -> list[dict]:
    soup = BeautifulSoup(html, "lxml")
    table = next(t for t in soup.select("table.wikitable")
                 if [th.get_text(strip=True) for th in t.select("tr th")][:2] == ["Code", "Company"])
    rows = []
    for tr in table.select("tr")[1:]:
        tds = tr.find_all("td")
        if len(tds) < 5:
            continue
        link = tds[1].find("a", href=True)
        classes = (link.get("class") or []) if link else []
        url = None
        if link and "new" not in classes:  # 'new' = red link, the article does not exist
            url = link["href"]
            if url.startswith("/wiki/"):
                url = "https://en.wikipedia.org" + url
        rows.append({
            "asx_code": clean_text(tds[0]),
            "company": clean_text(tds[1]),
            "sector": clean_text(tds[2]),
            "market_cap_aud": to_int(clean_text(tds[3])),
            "hq_list": clean_text(tds[4]),
            "linked_url": url,
        })
    return rows


COMPANY_LABELS = {"traded as", "industry", "headquarters", "revenue", "website",
                  "number of employees", "employees", "type", "company type"}


def company_infobox(soup):
    """First infobox that describes an organisation (not a person)."""
    for box in soup.select("table.infobox"):
        classes = box.get("class") or []
        if "biography" in classes:
            continue
        labels = {th.get_text(" ", strip=True).lower() for th in box.select("tr th")}
        if "ib-company" in classes or len(labels & COMPANY_LABELS) >= 3:
            return box
    return None


def parse_company(html: str) -> dict:
    soup = BeautifulSoup(html, "lxml")
    heading = soup.select_one("h1")
    record = {"article_title": clean_text(heading) if heading else "", "has_infobox": False}
    box = company_infobox(soup)
    if box is None:
        record["person_page"] = soup.select_one("table.infobox.biography") is not None
        return record
    record["has_infobox"] = True
    for tr in box.select("tr"):
        th, td = tr.find("th"), tr.find("td")
        if th is None or td is None:
            continue
        label = clean_text(th).lower()
        if label == "website" and "website" not in record:
            link = td.select_one("a.external[href]") or td.select_one("a[href^='http']")
            record["website"] = link["href"] if link else clean_text(td)
        elif label in ("type", "company type") and "type" not in record:
            record["type"] = clean_text(td)
        elif label == "traded as" and "traded_as" not in record:
            record["traded_as"] = clean_text(td)
        elif label == "industry" and "industry" not in record:
            # collapsible lists start with a caption such as 'List of industries'
            record["industry"] = re.sub(r"^List of [a-z]+\s+", "", clean_text(td))
        elif label in ("founded", "formation") and "founded_raw" not in record:
            record["founded_raw"] = clean_text(td)
        elif label == "headquarters" and "hq_infobox" not in record:
            record["hq_infobox"] = clean_text(td).replace("; ", ", ")
        elif label in ("number of employees", "employees") and "employees_raw" not in record:
            record["employees_raw"] = clean_text(td)
        elif label == "revenue" and "revenue_raw" not in record:
            record["revenue_raw"] = clean_text(td)
    return record


def lists_code(record: dict, code: str) -> bool | None:
    """Does the infobox 'Traded as' line show this code?

    True  = 'ASX: CODE', or the same code on another exchange (e.g. 'NZX: CEN')
    False = the line exists and shows other codes only
    None  = no 'Traded as' line, or it names exchanges without codes
    """
    traded = record.get("traded_as")
    if not traded:
        return None
    if re.search(rf"\b[A-Z]{{2,8}}\W{{0,3}}{re.escape(code)}\b", traded):
        return True
    if not re.search(r"\b[A-Z]{2,8}\s*:\s*[A-Z0-9]{1,6}\b", traded):
        return None
    return False


def candidate_urls(name: str) -> list[str]:
    """Article titles to try for a company the index table does not link."""
    base = re.sub(r"\s+", " ", name).strip()
    titles = [base]
    if not re.search(r"\b(Limited|Ltd)$", base):
        titles.append(base + " Limited")
    return [WIKI + quote(t.replace(" ", "_"), safe="_()',.-") for t in dict.fromkeys(titles)]


def read_article(fetcher: Fetcher, url: str) -> dict:
    html, meta = fetcher.get(url)
    record = parse_company(html)
    record["article_url"] = meta.get("final_url") or url
    record["fetched_at"] = meta.get("fetched_at")
    return record


# --------------------------------------------------------------------------- #
# Checks
# --------------------------------------------------------------------------- #
def review_reasons(row: dict, article_counts: dict) -> list[tuple[str, str, str]]:
    """(type, issue, what to check) for one row. Types: Not found, Conflict, Missing."""
    if not row.get("article_url"):
        return [("Not found", "No Wikipedia article for this company",
                 "Fill the fields from the company website or annual report")]
    if not row.get("has_infobox"):
        title = row.get("article_title") or "the linked page"
        issue = (f"Index link opens a page about a person ({title})" if row.get("person_page")
                 else f"Index link opens {title}, a page with no company infobox")
        return [("Not found", issue, "Fill the fields from the company website or annual report")]
    issues = []
    if article_counts.get(row["article_url"], 0) > 1:
        issues.append(("Conflict", "Two index codes use the same article",
                       "Check if it is a second share class or a wrong link"))
    if row.get("code_in_article") is False:
        shown = ", ".join(f"{ex}: {cd}" for ex, cd in
                          re.findall(r"\b([A-Z]{2,8})\s*:\s*([A-Z0-9]{1,6})\b", row.get("traded_as") or ""))
        issues.append(("Conflict", f"Index table says {row['asx_code']}, article lists {shown}",
                       "Confirm the ASX code on the ASX website"))
    if head_office_conflict(row.get("hq_list"), row.get("hq_infobox")):
        issues.append(("Conflict", f"Head office: index table says {row['hq_list']}, "
                       f"article says {row['hq_infobox']}", "Confirm the current head office"))
    if row.get("revenue_m") is not None and row.get("revenue_currency") == "$":
        issues.append(("Conflict", "Revenue currency is a bare $ and the head office is abroad",
                       "Confirm the currency in the annual report"))
    if not row.get("website"):
        issues.append(("Missing", "No website in the article", "Search the official site"))
    return issues


def figure_gaps(row: dict) -> list[str]:
    """Employee and revenue figures that are missing, undated or 3+ years old."""
    if not row.get("has_infobox"):
        return []
    gaps = []
    for label, value, year in (("Employee count", row.get("employees"), row.get("employees_year")),
                               ("Revenue", row.get("revenue_m"), row.get("revenue_year"))):
        if value is None:
            gaps.append(f"No {label.lower()}")
        elif year is None:
            gaps.append(f"{label} has no year")
        elif year < OLD_BEFORE:
            gaps.append(f"{label} is from {year}")
    return gaps


# --------------------------------------------------------------------------- #
# Main
# --------------------------------------------------------------------------- #
def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--delay", type=float, default=1.0, help="seconds between network requests")
    parser.add_argument("--refresh", action="store_true", help="ignore the cache and download again")
    args = parser.parse_args()

    started = datetime.now(timezone.utc)
    clock = time.monotonic()
    fetcher = Fetcher(delay=args.delay, refresh=args.refresh)
    OUT.mkdir(exist_ok=True)

    list_html, list_meta = fetcher.get(LIST_URL)
    rows = parse_list(list_html)
    print(f"index table: {len(rows)} rows, {sum(1 for r in rows if r['linked_url'])} with a link")

    failures, tried_titles = [], 0
    for number, row in enumerate(rows, 1):
        record, found_by = None, None
        if row["linked_url"]:
            try:
                record = read_article(fetcher, row["linked_url"])
                found_by = "Link in index table"
            except NotFound:
                record = None
            except Exception as exc:  # keep going; the row goes to review
                failures.append({"asx_code": row["asx_code"], "url": row["linked_url"], "error": str(exc)})
        if record is None or not record["has_infobox"] or lists_code(record, row["asx_code"]) is False:
            for url in candidate_urls(row["company"]):
                if record and url == record.get("article_url"):
                    continue
                tried_titles += 1
                try:
                    candidate = read_article(fetcher, url)
                except NotFound:
                    continue
                except Exception as exc:
                    failures.append({"asx_code": row["asx_code"], "url": url, "error": str(exc)})
                    continue
                if candidate["has_infobox"] and lists_code(candidate, row["asx_code"]):
                    record, found_by = candidate, "Name search, confirmed by ASX code"
                    break
        if record:
            row.update(record)
            row["found_by"] = found_by
            row["code_in_article"] = lists_code(record, row["asx_code"])
        if number % 25 == 0:
            print(f"  {number}/{len(rows)} companies done")

    article_counts: dict[str, int] = {}
    for row in rows:
        if row.get("article_url") and row.get("has_infobox"):
            article_counts[row["article_url"]] = article_counts.get(row["article_url"], 0) + 1

    for row in rows:
        row["founded"] = first_year(row.get("founded_raw"))
        row["employees"] = parse_employees(row.get("employees_raw"))
        row["employees_year"] = figure_year(row.get("employees_raw")) if row["employees"] else None
        row["revenue_m"], row["revenue_currency"] = parse_money(row.get("revenue_raw"))
        row["revenue_year"] = figure_year(row.get("revenue_raw")) if row["revenue_m"] else None
        row["currency_note"] = None
        australian_office = "," not in (row.get("hq_list") or "")  # the index table only adds a country abroad
        if row["revenue_currency"] == "$" and australian_office:
            row["revenue_currency"] = "AUD"
            row["currency_note"] = "Bare $ read as A$ (Australian head office)"
        row["website_domain"] = website_domain(row.get("website"))
        row["head_office"] = row.get("hq_infobox") or row.get("hq_list")
        reasons = review_reasons(row, article_counts)
        row["issue_types"] = ", ".join(dict.fromkeys(kind for kind, _, _ in reasons))
        row["issues"] = " | ".join(issue for _, issue, _ in reasons)
        row["what_to_check"] = " | ".join(dict.fromkeys(action for _, _, action in reasons))
        row["status"] = ("Not found" if row["issue_types"] == "Not found"
                         else "Check" if reasons else "OK")
        row["figures_to_update"] = " | ".join(figure_gaps(row))

    columns = ["asx_code", "company", "sector", "industry", "head_office", "hq_list", "founded",
               "employees", "employees_year", "revenue_m", "revenue_currency", "revenue_year",
               "currency_note", "website_domain", "website", "market_cap_aud", "type", "traded_as",
               "article_title", "article_url", "found_by", "code_in_article", "status", "issue_types",
               "issues", "what_to_check", "figures_to_update", "founded_raw", "employees_raw",
               "revenue_raw", "fetched_at"]
    df = pd.DataFrame(rows).reindex(columns=columns)
    df["figures_to_update"] = df["figures_to_update"].replace("", None)
    df.to_csv(OUT / "asx200_directory.csv", index=False)
    review = df[df["status"] != "OK"][["asx_code", "company", "status", "issues", "what_to_check", "article_url"]]
    review.to_csv(OUT / "review.csv", index=False)
    updates = df[df["figures_to_update"].notna()][["asx_code", "company", "employees", "employees_year",
                                                   "revenue_m", "revenue_currency", "revenue_year",
                                                   "figures_to_update", "article_url"]]
    updates.to_csv(OUT / "figures_to_update.csv", index=False)

    coverage = {}
    for col in ["article_url", "industry", "head_office", "founded", "employees", "revenue_m", "website_domain"]:
        filled = int(df[col].notna().sum())
        coverage[col] = {"filled": filled, "pct": round(100 * filled / len(df), 1)}
    issue_counts: dict[str, int] = {}
    for text in df["issues"].dropna():
        for issue in filter(None, text.split(" | ")):
            key = re.sub(r": index table says .*|, article lists .*", "", issue)
            key = "Index code differs from article" if key.startswith("Index table says") else key
            issue_counts[key] = issue_counts.get(key, 0) + 1
    type_counts = df["status"].value_counts().to_dict()
    gap_counts: dict[str, int] = {}
    for text in df["figures_to_update"].dropna():
        for gap in text.split(" | "):
            key = re.sub(r" from \d{4}$", " is 3+ years old", gap).replace(" is is", " is")
            gap_counts[key] = gap_counts.get(key, 0) + 1

    log = {
        "run_started_utc": started.isoformat(timespec="seconds"),
        "run_seconds": round(time.monotonic() - clock, 1),
        "list_url": LIST_URL,
        "list_fetched_at": list_meta.get("fetched_at"),
        "user_agent": USER_AGENT,
        "delay_seconds": args.delay,
        "rows_in_index_table": len(df),
        "unique_asx_codes": int(df["asx_code"].nunique()),
        "rows_linked_in_index_table": sum(1 for r in rows if r["linked_url"]),
        "found_by": df["found_by"].value_counts().to_dict(),
        "title_guesses_tried": tried_titles,
        "articles_used": int(df["article_url"].notna().sum()),
        "pages_with_company_infobox": sum(1 for r in rows if r.get("has_infobox")),
        "rows_by_status": type_counts,
        "issue_counts": dict(sorted(issue_counts.items(), key=lambda kv: -kv[1])),
        "rows_with_figures_to_update": int(df["figures_to_update"].notna().sum()),
        "figure_gap_counts": dict(sorted(gap_counts.items(), key=lambda kv: -kv[1])),
        "coverage": coverage,
        "fetch_stats": fetcher.stats,
        "failures": failures,
        "data_licence": "Text from Wikipedia, CC BY-SA 4.0 (https://creativecommons.org/licenses/by-sa/4.0/)",
    }
    (OUT / "scrape_log.json").write_text(json.dumps(log, indent=2))
    print(json.dumps({k: log[k] for k in ("rows_in_index_table", "found_by", "rows_by_status",
                                           "issue_counts", "rows_with_figures_to_update",
                                           "run_seconds")}, indent=1))
    print("fetch:", fetcher.stats)


if __name__ == "__main__":
    main()
