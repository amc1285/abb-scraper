# AudioBook Bay Scraper

Scrapes listing metadata from AudioBook Bay and exports to a formatted Excel spreadsheet.

## What it collects
Each row in the spreadsheet contains: **Title, Author, URL, Categories, Keywords, Language, Format, Bitrate, File Size, Date Posted**

## How to run

### Option A: Run manually on GitHub (recommended)
1. Go to the **Actions** tab in this repo
2. Click **AudioBook Bay Scraper** in the left sidebar
3. Click **Run workflow**
4. Enter how many pages to scrape (default: 50, max: 500)
5. Click the green **Run workflow** button
6. Wait ~2-5 minutes, then download your Excel file from the run's Artifacts section

### Option B: Automatic weekly run
The scraper runs automatically every Sunday at 6am UTC. Just check the Actions tab to download the latest file.

### Option C: Run locally
```bash
pip install requests beautifulsoup4 openpyxl
MAX_PAGES=50 python scraper.py
```

## Configuration
Edit the workflow file or set these environment variables:
- `MAX_PAGES` — number of pages to scrape (each page has ~9 books). Set to `0` or leave blank for all 500 pages.
- `DELAY_MIN` / `DELAY_MAX` — seconds to wait between page requests (default: 2–4s)

## Notes
- For personal/research use only
- The scraper collects metadata only — no files are downloaded
- GitHub Actions gives you 2,000 free minutes/month (each full scrape takes ~5 min)
