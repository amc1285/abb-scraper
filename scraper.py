import requests
from bs4 import BeautifulSoup
import time
import random
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime
import os

BASE_URL = "https://audiobookbay.lu"
HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
}

# Read config from environment variables (set in GitHub Actions workflow)
MAX_PAGES = int(os.environ.get("MAX_PAGES", "50"))
DELAY_MIN = float(os.environ.get("DELAY_MIN", "2.0"))
DELAY_MAX = float(os.environ.get("DELAY_MAX", "4.0"))


def get_page(url, retries=3):
    for attempt in range(retries):
        try:
            resp = requests.get(url, headers=HEADERS, timeout=15)
            resp.raise_for_status()
            return resp.text
        except Exception as e:
            print(f"  Attempt {attempt+1} failed for {url}: {e}")
            if attempt < retries - 1:
                time.sleep(5)
    print(f"  [FAILED] Could not fetch {url}")
    return None


def parse_listing_page(html):
    soup = BeautifulSoup(html, "html.parser")
    books = []

    for post in soup.select("div.post"):
        try:
            title_tag = post.select_one("div.postTitle h2 a")
            if not title_tag:
                continue
            full_title = title_tag.get_text(strip=True)
            url = title_tag.get("href", "")
            if url and not url.startswith("http"):
                url = BASE_URL + url

            if " - " in full_title:
                parts = full_title.rsplit(" - ", 1)
                title = parts[0].strip()
                author = parts[1].strip()
            else:
                title = full_title
                author = ""

            categories = ""
            keywords = ""
            language = ""

            details = post.select_one("div.postDetails")
            if details:
                for line in details.get_text("\n").split("\n"):
                    line = line.strip()
                    if line.startswith("Category:"):
                        categories = line.replace("Category:", "").strip()
                    elif line.startswith("Keywords:"):
                        keywords = line.replace("Keywords:", "").strip()
                    elif line.startswith("Language:"):
                        language = line.replace("Language:", "").strip()

            format_ = ""
            bitrate = ""
            file_size = ""
            date_posted = ""

            post_content = post.select_one("div.postContent")
            if post_content:
                for line in post_content.get_text("\n").split("\n"):
                    line = line.strip()
                    if line.startswith("Format:"):
                        parts = line.split("/")
                        format_ = parts[0].replace("Format:", "").strip()
                        if len(parts) > 1 and "Bitrate:" in parts[1]:
                            bitrate = parts[1].replace("Bitrate:", "").strip()
                    elif line.startswith("File Size:"):
                        file_size = line.replace("File Size:", "").strip()
                    elif line.startswith("Posted:"):
                        date_posted = line.replace("Posted:", "").strip()

            books.append({
                "Title": title,
                "Author": author,
                "URL": url,
                "Categories": categories,
                "Keywords": keywords,
                "Language": language,
                "Format": format_,
                "Bitrate": bitrate,
                "File Size": file_size,
                "Date Posted": date_posted,
            })
        except Exception as e:
            print(f"  [PARSE ERROR] {e}")
            continue

    return books


def get_total_pages(html):
    soup = BeautifulSoup(html, "html.parser")
    max_page = 1
    for a in soup.select("div.pagination a"):
        try:
            href = a.get("href", "")
            if "/page/" in href:
                num = int(href.rstrip("/").split("/")[-1])
                if num > max_page:
                    max_page = num
        except:
            pass
    return max_page


def save_xlsx(books, filename):
    wb = Workbook()
    ws = wb.active
    ws.title = "Audiobooks"

    header_font = Font(name="Arial", bold=True, color="FFFFFF", size=11)
    header_fill = PatternFill("solid", start_color="1F4E79")
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left = Alignment(horizontal="left", vertical="center", wrap_text=True)
    thin = Side(style="thin", color="CCCCCC")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    columns = ["Title", "Author", "URL", "Categories", "Keywords", "Language",
               "Format", "Bitrate", "File Size", "Date Posted"]
    col_widths = [45, 25, 60, 30, 40, 12, 10, 10, 12, 15]

    for col_idx, (col_name, width) in enumerate(zip(columns, col_widths), 1):
        cell = ws.cell(row=1, column=col_idx, value=col_name)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center
        cell.border = border
        ws.column_dimensions[get_column_letter(col_idx)].width = width

    ws.row_dimensions[1].height = 22
    ws.freeze_panes = "A2"

    fill_light = PatternFill("solid", start_color="EEF4FB")
    fill_white = PatternFill("solid", start_color="FFFFFF")

    for row_idx, book in enumerate(books, 2):
        fill = fill_light if row_idx % 2 == 0 else fill_white
        for col_idx, col_name in enumerate(columns, 1):
            val = book.get(col_name, "")
            cell = ws.cell(row=row_idx, column=col_idx, value=val)
            cell.font = Font(name="Arial", size=10)
            cell.fill = fill
            cell.border = border
            cell.alignment = left if col_name != "Date Posted" else center

    ws2 = wb.create_sheet("Summary")
    ws2["A1"] = "ABB Scrape Summary"
    ws2["A1"].font = Font(name="Arial", bold=True, size=14)
    ws2["A2"] = "Total Books Scraped:"
    ws2["B2"] = len(books)
    ws2["A3"] = "Scrape Date:"
    ws2["B3"] = datetime.now().strftime("%Y-%m-%d %H:%M")
    ws2["A4"] = "Source:"
    ws2["B4"] = BASE_URL
    ws2["A5"] = "Pages Scraped:"
    ws2["B5"] = MAX_PAGES if MAX_PAGES else "All"
    ws2.column_dimensions["A"].width = 22
    ws2.column_dimensions["B"].width = 25

    wb.save(filename)
    print(f"Saved {len(books)} books to {filename}")


def scrape():
    print(f"Fetching page 1...")
    first_html = get_page(f"{BASE_URL}/")
    if not first_html:
        print("Could not reach site. Exiting.")
        return []

    total_pages = get_total_pages(first_html)
    end_page = min(total_pages, MAX_PAGES) if MAX_PAGES else total_pages
    print(f"Site has {total_pages} pages. Scraping pages 1 to {end_page}.")

    all_books = []

    for page in range(1, end_page + 1):
        if page == 1:
            html = first_html
        else:
            url = f"{BASE_URL}/page/{page}/"
            html = get_page(url)
            if not html:
                continue

        books = parse_listing_page(html)
        all_books.extend(books)
        print(f"Page {page}/{end_page}: {len(books)} books | Total so far: {len(all_books)}")

        if page < end_page:
            time.sleep(random.uniform(DELAY_MIN, DELAY_MAX))

    return all_books


if __name__ == "__main__":
    print("=" * 50)
    print("AudioBook Bay Scraper - GitHub Actions Run")
    print(f"Config: MAX_PAGES={MAX_PAGES}, DELAY={DELAY_MIN}-{DELAY_MAX}s")
    print("=" * 50)

    books = scrape()

    if books:
        os.makedirs("output", exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M")
        filename = f"output/audiobookbay_{timestamp}.xlsx"
        save_xlsx(books, filename)
    else:
        print("No books scraped.")
