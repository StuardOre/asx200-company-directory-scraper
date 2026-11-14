#!/usr/bin/env python3
"""
Builds ASX200_Company_Directory.xlsx from the scraper output in out/.

Sheets: Summary (formulas), Directory, Review, Figures to update, By sector (formulas), Log.
Run after scrape_asx200.py:  python build_xlsx.py
"""
from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

BASE = Path(__file__).resolve().parent
OUT = BASE / "out"
XLSX = BASE / "ASX200_Company_Directory.xlsx"
LAST = 1000  # formula ranges reach row 1000, so the sheet still works if rows are added

NAVY, WHITE = "1F3B57", "FFFFFF"
FONT = "Arial"
F_BASE = Font(name=FONT, size=10)
F_BOLD = Font(name=FONT, size=10, bold=True)
F_HEAD = Font(name=FONT, size=10, bold=True, color=WHITE)
F_TITLE = Font(name=FONT, size=16, bold=True, color=NAVY)
F_SECTION = Font(name=FONT, size=11, bold=True, color=NAVY)
F_NOTE = Font(name=FONT, size=9, italic=True, color="555555")
F_LINK = Font(name=FONT, size=10, color="0563C1", underline="single")
FILL_HEAD = PatternFill("solid", fgColor=NAVY)
FILL_OK = PatternFill("solid", fgColor="D9EAD3")
FILL_CHECK = PatternFill("solid", fgColor="FCE8B2")
FILL_NOTFOUND = PatternFill("solid", fgColor="E7E6E6")
FILL_OLD = PatternFill("solid", fgColor="FFF2CC")
THIN = Side(style="thin", color="D0D7DE")
BORDER = Border(bottom=THIN)
WRAP = Alignment(wrap_text=True, vertical="top")
TOP = Alignment(vertical="top")


def header(ws, row: int, titles: list[str], widths: list[float]) -> None:
    for col, (title, width) in enumerate(zip(titles, widths), 1):
        cell = ws.cell(row=row, column=col, value=title)
        cell.font, cell.fill = F_HEAD, FILL_HEAD
        cell.alignment = Alignment(wrap_text=True, vertical="center")
        ws.column_dimensions[get_column_letter(col)].width = width
    ws.row_dimensions[row].height = 30


def put(ws, row: int, col: int, value, fmt: str | None = None, font: Font = F_BASE,
        link: str | None = None, align: Alignment = TOP):
    if isinstance(value, float) and pd.isna(value):
        value = None
    cell = ws.cell(row=row, column=col, value=value)
    cell.font, cell.alignment, cell.border = font, align, BORDER
    if fmt:
        cell.number_format = fmt
    if link:
        cell.hyperlink = link
        cell.font = F_LINK
    return cell


def clean(value):
    return None if value is None or (isinstance(value, float) and pd.isna(value)) else value


def main() -> None:
    df = pd.read_csv(OUT / "asx200_directory.csv")
    log = json.loads((OUT / "scrape_log.json").read_text())
    run = datetime.fromisoformat(log["run_started_utc"])

    wb = Workbook()
    summary = wb.active
    summary.title = "Summary"
    directory = wb.create_sheet("Directory")
    review = wb.create_sheet("Review")
    updates = wb.create_sheet("Figures to update")
    sectors = wb.create_sheet("By sector")
    log_ws = wb.create_sheet("Log")

    # ------------------------------------------------------------------ Directory
    titles = ["ASX code", "Company", "Sector", "Industry", "Head office", "Founded", "Employees",
              "Employees year", "Revenue (millions)", "Currency", "Revenue year", "Website",
              "Market cap (A$ bn)", "Status", "Issue", "Figures to update", "Source page", "Found by"]
    widths = [8, 30, 20, 30, 30, 9, 11, 10, 12, 9, 9, 26, 11, 11, 44, 36, 34, 22]
    header(directory, 1, titles, widths)
    for r, row in enumerate(df.itertuples(index=False), 2):
        put(directory, r, 1, row.asx_code, font=F_BOLD)
        put(directory, r, 2, row.company, align=WRAP)
        put(directory, r, 3, row.sector)
        put(directory, r, 4, clean(row.industry), align=WRAP)
        put(directory, r, 5, clean(row.head_office), align=WRAP)
        put(directory, r, 6, clean(row.founded), "0")
        put(directory, r, 7, clean(row.employees), "#,##0")
        put(directory, r, 8, clean(row.employees_year), "0")
        put(directory, r, 9, clean(row.revenue_m), "#,##0.0")
        put(directory, r, 10, clean(row.revenue_currency))
        put(directory, r, 11, clean(row.revenue_year), "0")
        put(directory, r, 12, clean(row.website_domain), link=clean(row.website))
        cap = clean(row.market_cap_aud)
        put(directory, r, 13, round(cap / 1e9, 2) if cap else None, "#,##0.00")
        put(directory, r, 14, row.status, font=F_BOLD)
        put(directory, r, 15, clean(row.issues), align=WRAP)
        put(directory, r, 16, clean(row.figures_to_update), align=WRAP)
        article = clean(row.article_url)
        put(directory, r, 17, clean(row.article_title) or ("No page" if not article else article), link=article)
        put(directory, r, 18, clean(row.found_by))
    n = len(df) + 1
    directory.freeze_panes = "C2"
    directory.auto_filter.ref = f"A1:R{n}"
    directory.conditional_formatting.add(f"N2:N{n}", CellIsRule(operator="equal", formula=['"OK"'], fill=FILL_OK))
    directory.conditional_formatting.add(f"N2:N{n}", CellIsRule(operator="equal", formula=['"Check"'], fill=FILL_CHECK))
    directory.conditional_formatting.add(f"N2:N{n}", CellIsRule(operator="equal", formula=['"Not found"'], fill=FILL_NOTFOUND))
    for col in ("H", "K"):
        directory.conditional_formatting.add(
            f"{col}2:{col}{n}", FormulaRule(formula=[f"AND(ISNUMBER({col}2),{col}2<2023)"], fill=FILL_OLD))

    # ------------------------------------------------------------------ Review
    titles = ["ASX code", "Company", "Status", "Issue", "What to check", "Source page"]
    header(review, 1, titles, [8, 30, 11, 62, 44, 34])
    rev = df[df["status"] != "OK"].copy()
    rev["order"] = rev["status"].map({"Check": 0, "Not found": 1})
    rev = rev.sort_values(["order", "asx_code"])
    for r, row in enumerate(rev.itertuples(index=False), 2):
        put(review, r, 1, row.asx_code, font=F_BOLD)
        put(review, r, 2, row.company)
        put(review, r, 3, row.status, font=F_BOLD)
        put(review, r, 4, clean(row.issues), align=WRAP)
        put(review, r, 5, clean(row.what_to_check), align=WRAP)
        article = clean(row.article_url)
        put(review, r, 6, clean(row.article_title) or "No page", link=article)
    m = len(rev) + 1
    review.freeze_panes = "C2"
    review.auto_filter.ref = f"A1:F{m}"
    review.conditional_formatting.add(f"C2:C{m}", CellIsRule(operator="equal", formula=['"Check"'], fill=FILL_CHECK))
    review.conditional_formatting.add(f"C2:C{m}", CellIsRule(operator="equal", formula=['"Not found"'], fill=FILL_NOTFOUND))

    # ------------------------------------------------------------------ Figures to update
    titles = ["ASX code", "Company", "Employees", "Employees year", "Revenue (millions)", "Currency",
              "Revenue year", "What to update", "Source page"]
    header(updates, 1, titles, [8, 30, 11, 10, 12, 9, 9, 48, 34])
    upd = df[df["figures_to_update"].notna()].sort_values("asx_code")
    for r, row in enumerate(upd.itertuples(index=False), 2):
        put(updates, r, 1, row.asx_code, font=F_BOLD)
        put(updates, r, 2, row.company)
        put(updates, r, 3, clean(row.employees), "#,##0")
        put(updates, r, 4, clean(row.employees_year), "0")
        put(updates, r, 5, clean(row.revenue_m), "#,##0.0")
        put(updates, r, 6, clean(row.revenue_currency))
        put(updates, r, 7, clean(row.revenue_year), "0")
        put(updates, r, 8, row.figures_to_update, align=WRAP)
        put(updates, r, 9, clean(row.article_title), link=clean(row.article_url))
    k = len(upd) + 1
    updates.freeze_panes = "C2"
    updates.auto_filter.ref = f"A1:I{k}"
    for col in ("D", "G"):
        updates.conditional_formatting.add(
            f"{col}2:{col}{k}", FormulaRule(formula=[f"AND(ISNUMBER({col}2),{col}2<2023)"], fill=FILL_OLD))

    # ------------------------------------------------------------------ By sector (formulas)
    titles = ["Sector", "Companies", "OK", "Check", "Not found", "With website", "Market cap (A$ bn)"]
    header(sectors, 1, titles, [26, 11, 8, 8, 10, 12, 14])
    D = "Directory!"
    rng = lambda col: f"{D}${col}$2:${col}${LAST}"
    order = df["sector"].value_counts().index.tolist()
    for r, sector in enumerate(order, 2):
        put(sectors, r, 1, sector)
        put(sectors, r, 2, f"=COUNTIF({rng('C')},A{r})", "0")
        put(sectors, r, 3, f'=COUNTIFS({rng("C")},A{r},{rng("N")},"OK")', "0")
        put(sectors, r, 4, f'=COUNTIFS({rng("C")},A{r},{rng("N")},"Check")', "0")
        put(sectors, r, 5, f'=COUNTIFS({rng("C")},A{r},{rng("N")},"Not found")', "0")
        put(sectors, r, 6, f'=COUNTIFS({rng("C")},A{r},{rng("L")},"<>")', "0")
        put(sectors, r, 7, f"=SUMIF({rng('C')},A{r},{rng('M')})", "#,##0.0")
    t = len(order) + 2
    put(sectors, t, 1, "Total", font=F_BOLD)
    for col in "BCDEFG":
        put(sectors, t, "BCDEFG".index(col) + 2, f"=SUM({col}2:{col}{t - 1})",
            "#,##0.0" if col == "G" else "0", font=F_BOLD)
    put(sectors, t + 2, 1, "Market cap as shown in the index table on Wikipedia (A$ billions).", font=F_NOTE,
        align=Alignment(wrap_text=False))
    sectors.freeze_panes = "B2"

    # ------------------------------------------------------------------ Summary (formulas)
    ws = summary
    ws.column_dimensions["A"].width = 44
    ws.column_dimensions["B"].width = 16
    ws.column_dimensions["C"].width = 16
    ws.column_dimensions["D"].width = 60
    ws["A1"] = "ASX 200 company directory"
    ws["A1"].font = F_TITLE
    ws["A2"] = (f"Scraped from public Wikipedia pages on {run:%d %b %Y} with Python (requests + BeautifulSoup). "
                "One row per index company, with a status and a source link.")
    ws["A2"].font = F_NOTE

    def section(row, title, cols=("",)):
        ws.cell(row=row, column=1, value=title).font = F_SECTION
        for i, name in enumerate(cols, 2):
            if name:
                c = ws.cell(row=row, column=i, value=name)
                c.font = F_BOLD
                c.alignment = Alignment(horizontal="right")

    def line(row, label, formula, fmt="0", note=None):
        put(ws, row, 1, label)
        put(ws, row, 2, formula, fmt, font=F_BOLD)
        if note:
            put(ws, row, 4, note, font=F_NOTE)

    section(4, "Result", ("Rows",))
    line(5, "Companies in the index table", f"=COUNTA({rng('A')})")
    line(6, "Company pages read", "=B7+B8")
    line(7, "Status OK", f'=COUNTIF({rng("N")},"OK")', note="Page found, no conflicts")
    line(8, "Status Check", f'=COUNTIF({rng("N")},"Check")',
         note="The index table and the company page disagree (see Review)")
    line(9, "Status Not found", f'=COUNTIF({rng("N")},"Not found")',
         note="No company page; fill from the company website")
    line(10, "Rows with figures to update", f'=COUNTIF({rng("P")},"?*")',
         note="Employees or revenue missing, undated or 3+ years old")

    section(12, "Field coverage", ("Rows filled", "% of companies"))
    fields = [("Website", "COUNTA", "L"), ("Industry", "COUNTA", "D"), ("Head office", "COUNTA", "E"),
              ("Founded (year)", "COUNT", "F"), ("Employees", "COUNT", "G"), ("Revenue", "COUNT", "I")]
    for i, (label, func, col) in enumerate(fields, 13):
        put(ws, i, 1, label)
        put(ws, i, 2, f"={func}({rng(col)})", "0", font=F_BOLD)
        put(ws, i, 3, f"=IF($B$5=0,0,B{i}/$B$5)", "0.0%")
    ws["D18"] = "Head office comes from the index table when there is no company page."
    ws["D18"].font = F_NOTE

    stats = log["fetch_stats"]
    codes = stats.get("status_codes", {})
    minutes, seconds = divmod(int(round(log["run_seconds"])), 60)
    section(20, "Scrape run (from out/scrape_log.json)")
    run_rows = [
        ("Run started (UTC)", f"{run:%Y-%m-%d %H:%M}"),
        ("Pages requested", stats["network_requests"]),
        ("Answered 200 (page) / 404 (no such page)", f"{codes.get('200', 0)} / {codes.get('404', 0)}"),
        ("Retries needed (429 or 5xx)", stats["retries"]),
        ("Pages blocked by robots.txt", stats["blocked_by_robots"]),
        ("Delay between requests", f"{log['delay_seconds']:g} second"),
        ("Run time", f"{minutes} min {seconds} s"),
        ("User-Agent", log["user_agent"]),
    ]
    for i, (label, value) in enumerate(run_rows, 21):
        put(ws, i, 1, label)
        put(ws, i, 2, value, font=F_BOLD,
            align=Alignment(horizontal="left" if label == "User-Agent" else "right", vertical="top"))
    ws["D22"] = "Hardcoded values copied from the run log; the counts above are live formulas."
    ws["D22"].font = F_NOTE

    section(30, "How it works")
    found = log.get("found_by", {})
    steps = [
        f"1. Reads the index table on the S&P/ASX 200 page ({log['rows_in_index_table']} rows, "
        f"{log['rows_linked_in_index_table']} with a link to a company page).",
        "2. Opens every linked page. For the others it tries the company name as a page title and keeps the page "
        f"only if it lists the same code ({found.get('Name search, confirmed by ASX code', 0)} more found).",
        "3. Cleans the infobox: founding year, head count and its year, revenue in millions with currency, "
        "website domain.",
        "4. Checks every row: head office in the same state or country, same code on both pages, links that open "
        "a person's page, and figures that are 3+ years old.",
    ]
    for i, text in enumerate(steps, 31):
        ws.cell(row=i, column=1, value=text).font = F_BASE

    section(36, "Legend")
    legend = [("OK", FILL_OK, "Company page found and it agrees with the index table"),
              ("Check", FILL_CHECK, "The two pages disagree: confirm before using the row"),
              ("Not found", FILL_NOTFOUND, "No company page: only index table data (code, sector, head office)"),
              ("2021", FILL_OLD, "Amber year: employee or revenue figure is 3+ years old")]
    for i, (label, fill, text) in enumerate(legend, 37):
        c = ws.cell(row=i, column=1, value=label)
        c.font, c.fill = F_BOLD, fill
        ws.cell(row=i, column=2, value=text).font = F_BASE

    section(42, "Source")
    ws["A43"] = ("Wikipedia: the S&P/ASX 200 page and each company article. Text under CC BY-SA 4.0 "
                 "(https://creativecommons.org/licenses/by-sa/4.0/).")
    ws["A44"] = (f"Index table as published on {run:%d %b %Y}; index membership and company figures may have "
                 "changed since. Figures should be confirmed in annual reports before use.")
    for ref in ("A43", "A44"):
        ws[ref].font = F_BASE

    # ------------------------------------------------------------------ Log
    log_ws.column_dimensions["A"].width = 44
    log_ws.column_dimensions["B"].width = 70
    header(log_ws, 1, ["Item", "Value"], [44, 70])
    items = [
        ("List page", log["list_url"]),
        ("Run started (UTC)", log["run_started_utc"]),
        ("Run seconds", log["run_seconds"]),
        ("User-Agent", log["user_agent"]),
        ("Delay between requests (s)", log["delay_seconds"]),
        ("Rows in index table", log["rows_in_index_table"]),
        ("Unique ASX codes", log["unique_asx_codes"]),
        ("Rows with a link in the index table", log["rows_linked_in_index_table"]),
        ("Title guesses tried", log["title_guesses_tried"]),
        ("Company pages used", log["pages_with_company_infobox"]),
        ("Network requests", stats["network_requests"]),
        ("Cache hits", stats["cache_hits"]),
        ("Retries", stats["retries"]),
        ("404 answers (page does not exist)", stats["not_found_404"]),
        ("Blocked by robots.txt", stats["blocked_by_robots"]),
        ("Failures", len(log["failures"])),
    ]
    items += [(f"Rows by status: {k}", v) for k, v in log["rows_by_status"].items()]
    items += [(f"Issue: {k}", v) for k, v in log["issue_counts"].items()]
    items += [(f"Figure gap: {k}", v) for k, v in log["figure_gap_counts"].items()]
    items += [("Data licence", log["data_licence"])]
    for r, (label, value) in enumerate(items, 2):
        put(log_ws, r, 1, label)
        put(log_ws, r, 2, value, align=Alignment(horizontal="left", vertical="top"))

    for sheet in wb.worksheets:
        sheet.sheet_view.showGridLines = sheet.title not in ("Summary",)
        sheet.page_setup.orientation = "landscape"
        sheet.page_setup.fitToWidth = 1
        sheet.page_setup.fitToHeight = 0
        sheet.sheet_properties.pageSetUpPr.fitToPage = True
    directory.sheet_properties.tabColor = NAVY
    review.sheet_properties.tabColor = "F1C232"
    wb.save(XLSX)
    print(f"saved {XLSX.name}: {len(df)} companies, {len(rev)} to review, {len(upd)} figures to update")


if __name__ == "__main__":
    main()
