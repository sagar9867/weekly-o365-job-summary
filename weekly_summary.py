import requests
from bs4 import BeautifulSoup
from datetime import datetime, timedelta
from dateutil import parser as dateparser
import re
import openpyxl

# Search keywords and locations
KEYWORDS = ["office 365", "o365", "microsoft 365", "sharepoint", "power apps", "power platform", "exchange"]
TARGET_LOCATIONS = ["india", "remote", "work from home"]

# Days window
DAYS_BACK = 7
SINCE_DT = datetime.utcnow() - timedelta(days=DAYS_BACK)

HEADERS = {
    "User-Agent": "Mozilla/5.0"
}

# Job portals
LINKEDIN_URL = "https://www.linkedin.com/jobs/search/?keywords=Office%20365&location=India"
NAUKRI_URLS = [
    "https://www.naukri.com/office-365-jobs",
    "https://www.naukri.com/sharepoint-jobs",
    "https://www.naukri.com/exchange-jobs",
    "https://www.naukri.com/powerapps-jobs",
]

def fetch(url):
    try:
        r = requests.get(url, headers=HEADERS, timeout=20)
        if r.status_code == 200:
            return r.text
    except:
        return ""
    return ""

# Extract posting date from job detail page
def extract_post_date(url):
    html = fetch(url)
    if not html:
        return None

    soup = BeautifulSoup(html, "html.parser")
    text = soup.get_text(separator=" ")

    # Try "Posted x days ago"
    m = re.search(r"(\d+)\s+days?\s+ago", text, re.IGNORECASE)
    if m:
        return datetime.utcnow() - timedelta(days=int(m.group(1)))

    # Try explicit date
    for label in ["Posted on", "Posted"]:
        idx = text.find(label)
        if idx != -1:
            snippet = text[idx: idx+50]
            try:
                dt = dateparser.parse(snippet, fuzzy=True)
                return dt
            except:
                pass

    return None

def parse_linkedin(html):
    soup = BeautifulSoup(html, "html.parser")
    jobs = []

    for a in soup.select("a[href]"):
        title = a.get_text(strip=True)
        href = a.get("href")
        if not href or not title:
            continue

        t_low = title.lower()

        if not any(k in t_low for k in KEYWORDS):
            continue

        m = re.search(r"/jobs/view/(\d+)", href)
        if not m:
            continue

        job_id = m.group(1)
        clean_url = f"https://www.linkedin.com/jobs/view/{job_id}"

        jobs.append({
            "title": title,
            "company": "",
            "location": "",
            "url": clean_url,
            "source": "LinkedIn"
        })

    return jobs

def parse_naukri(html):
    soup = BeautifulSoup(html, "html.parser")
    jobs = []

    for job in soup.select("div.row"):
        a = job.select_one("a.title.fw500")
        if not a:
            continue

        title = a.get_text(strip=True)
        href = a.get("href")
        if not href:
            continue

        loc_tag = job.select_one("li.loc")
        location = loc_tag.get_text(strip=True) if loc_tag else ""

        comp_tag = job.select_one("li.company")
        company = comp_tag.get_text(strip=True) if comp_tag else ""

        if not any(k in title.lower() for k in KEYWORDS):
            continue

        jobs.append({
            "title": title,
            "company": company,
            "location": location,
            "url": href,
            "source": "Naukri"
        })

    return jobs

def export_to_excel(jobs):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Jobs"

    ws.append(["Title", "Company", "Location", "URL", "Source", "Posted Date"])

    for j in jobs:
        ws.append([
            j["title"], j["company"], j["location"], j["url"], j["source"], j.get("posted", "")
        ])

    wb.save("weekly_jobs.xlsx")

def main():
    all_jobs = []

    # Naukri
    for url in NAUKRI_URLS:
        html = fetch(url)
        all_jobs += parse_naukri(html)

    # LinkedIn
    html = fetch(LINKEDIN_URL)
    all_jobs += parse_linkedin(html)

    # Filter by date
    final = []
    for j in all_jobs:
        dt = extract_post_date(j["url"])
        if dt and dt >= SINCE_DT:
            j["posted"] = dt.strftime("%Y-%m-%d")
            final.append(j)

    export_to_excel(final)
    print(f"Exported {len(final)} jobs to weekly_jobs.xlsx")

if __name__ == "__main__":
    main()

