import requests
from bs4 import BeautifulSoup
import openpyxl
import re

HEADERS = {"User-Agent": "Mozilla/5.0"}

KEYWORDS = ["office", "365", "o365", "m365", "sharepoint", "power", "exchange", "teams"]

NAUKRI_URLS = [
    "https://www.naukri.com/office-365-jobs",
    "https://www.naukri.com/sharepoint-jobs",
    "https://www.naukri.com/exchange-jobs",
    "https://www.naukri.com/powerapps-jobs",
    "https://www.naukri.com/o365-jobs?k=o365&nignbevent_src=jobsearchDeskGNB",
        "https://www.naukri.com/m365-jobs?k=m365",
    "https://www.naukri.com/modern-workplace-jobs",
    "https://www.naukri.com/microsoft-admin-jobs",
    "https://in.indeed.com/jobs?q=office+365&l=India",
    "https://www.shine.com/job-search/office-365-jobs",
    "https://www.foundit.in/search/office-365-jobs-in-mumbai",
"https://www.foundit.in/search/office-365-jobs-in-pune",

]

LINKEDIN_URL = "https://www.linkedin.com/jobs/search/?keywords=Office%20365&location=India"

def fetch(url):
    try:
        r = requests.get(url, headers=HEADERS, timeout=20)
        if r.status_code == 200:
            return r.text
    except:
        pass
    return ""

def parse_linkedin(html):
    soup = BeautifulSoup(html, "html.parser")
    jobs = []
    for a in soup.select("a[href]"):
        title = a.get_text(strip=True)
        href = a.get("href")
        if not href or not title:
            continue
        t = title.lower()
        if any(k in t for k in KEYWORDS):
            m = re.search(r"/jobs/view/(\d+)", href)
            if m:
                job_id = m.group(1)
                url = f"https://www.linkedin.com/jobs/view/{job_id}"
                jobs.append({
                    "title": title,
                    "company": "",
                    "location": "",
                    "url": url,
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
        t = title.lower()
        if any(k in t for k in KEYWORDS):
            loc_tag = job.select_one("li.loc")
            company_tag = job.select_one("li.company")
            jobs.append({
                "title": title,
                "company": company_tag.get_text(strip=True) if company_tag else "",
                "location": loc_tag.get_text(strip=True) if loc_tag else "",
                "url": href,
                "source": "Naukri"
            })
    return jobs

def export_excel(jobs):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Jobs"
    ws.append(["Title", "Company", "Location", "URL", "Source"])
    for j in jobs:
        ws.append([j["title"], j["company"], j["location"], j["url"], j["source"]])
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

    export_excel(all_jobs)
    print(f"Saved {len(all_jobs)} jobs into weekly_jobs.xlsx")

if __name__ == "__main__":
    main()
