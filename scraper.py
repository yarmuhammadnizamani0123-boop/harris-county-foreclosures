import os
import sys
import re
import json
import datetime
import pandas as pd
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

TARGET_URL = "https://www.cclerk.hctx.net/Applications/WebSearch/FRCL_R.aspx"
DB_DIR = "database"
REPORTS_DIR = "reports"
DB_FILE = os.path.join(DB_DIR, "known_leads.json")

def load_known_leads():
    os.makedirs(DB_DIR, exist_ok=True)
    if os.path.exists(DB_FILE):
        try:
            with open(DB_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}

def save_known_leads(known_dict):
    os.makedirs(DB_DIR, exist_ok=True)
    with open(DB_FILE, "w", encoding="utf-8") as f:
        json.dump(known_dict, f, indent=2)

def style_excel(ws):
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    zebra_fill = PatternFill(start_color="F9FAFB", end_color="F9FAFB", fill_type="solid")
    thin_border = Border(
        left=Side(style='thin', color='E5E7EB'),
        right=Side(style='thin', color='E5E7EB'),
        top=Side(style='thin', color='E5E7EB'),
        bottom=Side(style='thin', color='E5E7EB')
    )

    for col_idx in range(1, ws.max_column + 1):
        c = ws.cell(row=1, column=col_idx)
        c.fill = header_fill
        c.font = header_font
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for row_idx in range(2, ws.max_row + 1):
        is_even = (row_idx % 2 == 0)
        for col_idx in range(1, ws.max_column + 1):
            c = ws.cell(row=row_idx, column=col_idx)
            c.border = thin_border
            if is_even:
                c.fill = zebra_fill
            c.alignment = Alignment(vertical="center")

    for col in ws.columns:
        col_letter = get_column_letter(col[0].column)
        max_len = 0
        for cell in col:
            val_str = str(cell.value or '')
            if len(val_str) > max_len:
                max_len = len(val_str)
        ws.column_dimensions[col_letter].width = min(max(max_len + 3, 14), 70)

def extract_rows_from_page(page):
    records = []
    # Find all table rows
    rows = page.locator("table tr").all()
    print(f"[*] Found {len(rows)} raw HTML table rows on page.")

    for row in rows:
        td_elems = row.locator("td").all()
        if not td_elems:
            continue

        cols = [t.inner_text().strip() for t in td_elems]
        if not cols:
            continue

        doc_id = ""
        sale_date = ""
        file_date = ""
        direct_link = ""

        # Extract direct court link
        try:
            a_elem = row.locator("a").first
            if a_elem and a_elem.is_visible():
                href = a_elem.get_attribute("href") or ""
                if href:
                    if href.startswith("http"):
                        direct_link = href
                    else:
                        direct_link = f"https://www.cclerk.hctx.net/Applications/WebSearch/{href}"
        except Exception:
            pass

        # Find Document ID matching FRCL or format with numbers
        for c in cols:
            if re.search(r"FRCL|\b\d{4}-\d+\b", c, re.IGNORECASE):
                doc_id = c
                break

        # If not explicit FRCL, take the first non-empty column if it looks like an ID
        if not doc_id and len(cols[0]) >= 4 and not any(h in cols[0].lower() for h in ["document", "sale", "file", "click"]):
            doc_id = cols[0]

        if not doc_id:
            continue

        # Extract dates
        dates = []
        for c in cols:
            if re.search(r"\b\d{1,2}/\d{1,2}/\d{2,4}\b", c):
                dates.append(c)

        if len(dates) >= 2:
            sale_date = dates[0]
            file_date = dates[1]
        elif len(dates) == 1:
            sale_date = dates[0]
            file_date = datetime.date.today().strftime("%m/%d/%Y")
        else:
            sale_date = "Upcoming First Tuesday"
            file_date = datetime.date.today().strftime("%m/%d/%Y")

        if not direct_link:
            direct_link = f"https://www.cclerk.hctx.net/Applications/WebSearch/FRCL_R.aspx?FileNo={doc_id}"

        records.append({
            "Document_ID": doc_id,
            "Scheduled_Sale_Date": sale_date,
            "Court_Filing_Date": file_date,
            "County": "Harris County (Houston, TX)",
            "Document_Type": "Notice of Foreclosure Sale",
            "Direct_Court_Record_Link": direct_link
        })

    return records

def run_playwright():
    from playwright.sync_api import sync_playwright

    print("=" * 65)
    print("  HARRIS COUNTY FORECLOSURE EXTRACTION - LIVE CLOUD RUN")
    print("=" * 65)

    all_found_records = []

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
        )
        page = context.new_page()

        print(f"[*] Navigating to {TARGET_URL}...")
        page.goto(TARGET_URL, timeout=60000, wait_until="networkidle")

        # Log portal status
        header_text = page.locator("#ctl00_ContentPlaceHolder1_lblMaxDate").inner_text() if page.locator("#ctl00_ContentPlaceHolder1_lblMaxDate").is_visible() else "Portal Active"
        print(f"[*] Courthouse Portal Banner: {header_text}")

        # Strategy 1: Search with default "Sale Date" and click Search to get all current postings
        print("[*] Strategy 1: Executing Search on Active Scheduled Auctions...")
        try:
            page.click("input[name*='btnSearch']")
            page.wait_for_timeout(6000)
        except Exception as e:
            print(f"[!] Error clicking search: {e}")

        # Check for results or error labels
        count_lbl = page.locator("#ctl00_ContentPlaceHolder1_lblCount").inner_text() if page.locator("#ctl00_ContentPlaceHolder1_lblCount").is_visible() else ""
        msg_lbl = page.locator("#ctl00_ContentPlaceHolder1_lblMsg").inner_text() if page.locator("#ctl00_ContentPlaceHolder1_lblMsg").is_visible() else ""
        print(f"[*] Portal Response Count: '{count_lbl}' | Message: '{msg_lbl}'")

        # Extract records across pages
        page_idx = 1
        while True:
            page_records = extract_rows_from_page(page)
            print(f"  -> Extracted {len(page_records)} notices from Page {page_idx}.")
            all_found_records.extend(page_records)

            # Check for next page button
            next_btn = page.locator("a:has-text('Next'), a:has-text('>')").first
            if next_btn.is_visible() and next_btn.is_enabled():
                try:
                    next_btn.click()
                    page.wait_for_timeout(3500)
                    page_idx += 1
                except Exception:
                    break
            else:
                break

        # Strategy 2: If Strategy 1 returned 0, try selecting each available year
        if not all_found_records:
            print("[*] Strategy 2: Querying by specific active years...")
            years = [
                opt.get_attribute("value")
                for opt in page.locator("select[name*='ddlYear'] option").all()
                if opt.get_attribute("value") and opt.get_attribute("value").isdigit()
            ]
            print(f"[*] Available Year options: {years[:4]}")

            for y in years[:2]:
                print(f"[*] Selecting Year {y} and searching...")
                try:
                    page.select_option("select[name*='ddlYear']", y)
                    page.wait_for_timeout(1000)
                    page.click("input[name*='btnSearch']")
                    page.wait_for_timeout(5000)
                    rec = extract_rows_from_page(page)
                    print(f"  -> Year {y} returned {len(rec)} notices.")
                    all_found_records.extend(rec)
                except Exception as e:
                    print(f"[!] Strategy 2 error on year {y}: {e}")

        browser.close()

    return all_found_records

def main():
    records = run_playwright()
    today_str = datetime.date.today().strftime("%Y-%m-%d")

    # Deduplicate in-memory
    unique_records = []
    seen = set()
    for r in records:
        did = r["Document_ID"]
        if did not in seen:
            seen.add(did)
            unique_records.append(r)

    print(f"\n[+] Total Unique Foreclosure Notices Found: {len(unique_records)}")

    # Check against persistent database
    known_leads = load_known_leads()
    new_leads = []
    duplicates_filtered = 0

    for r in unique_records:
        did = r["Document_ID"]
        if did in known_leads:
            duplicates_filtered += 1
        else:
            known_leads[did] = {
                "discovered": today_str,
                "sale_date": r["Scheduled_Sale_Date"],
                "file_date": r["Court_Filing_Date"]
            }
            new_r = dict(r)
            new_r["Date_Discovered"] = today_str
            new_leads.append(new_r)

    save_known_leads(known_leads)

    os.makedirs(REPORTS_DIR, exist_ok=True)
    audit_file = os.path.join(REPORTS_DIR, "DAILY_RUN_AUDIT.txt")
    with open(audit_file, "w", encoding="utf-8") as f:
        f.write(f"Harris County Foreclosure Tracker Audit Log\n")
        f.write(f"Execution Date: {today_str}\n")
        f.write(f"Total Notices Active: {len(unique_records)}\n")
        f.write(f"New Filings Discovered Today: {len(new_leads)}\n")
        f.write(f"Existing Filings Deduplicated: {duplicates_filtered}\n")

    # Save Reports
    new_report_file = os.path.join(REPORTS_DIR, f"Harris_County_NEW_LEADS_{today_str}.xlsx")
    master_report_file = os.path.join(REPORTS_DIR, "Harris_County_ALL_ACTIVE_NOTICES_MASTER.xlsx")

    if new_leads:
        df_new = pd.DataFrame(new_leads)
        with pd.ExcelWriter(new_report_file, engine='openpyxl') as writer:
            df_new.to_excel(writer, sheet_name="New Filings Today", index=False)
            style_excel(writer.sheets["New Filings Today"])
        print(f"[+] Saved New Leads Excel: {new_report_file}")

    if unique_records:
        df_master = pd.DataFrame(unique_records)
        with pd.ExcelWriter(master_report_file, engine='openpyxl') as writer:
            df_master.to_excel(writer, sheet_name="All Active Notices", index=False)
            style_excel(writer.sheets["All Active Notices"])
        print(f"[+] Saved Master Archive Excel: {master_report_file}")
    else:
        # Save empty template with headers so reports folder is never empty
        empty_df = pd.DataFrame(columns=["Document_ID", "Scheduled_Sale_Date", "Court_Filing_Date", "County", "Document_Type", "Direct_Court_Record_Link"])
        with pd.ExcelWriter(master_report_file, engine='openpyxl') as writer:
            empty_df.to_excel(writer, sheet_name="All Active Notices", index=False)
        print(f"[*] Notice: Initialized clean Master Archive template.")

    print("\n[SUCCESS] Pipeline finished cleanly.")

if __name__ == "__main__":
    main()
