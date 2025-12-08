import requests
from bs4 import BeautifulSoup
import openpyxl
import re

# -----------------------
# CONFIG
# -----------------------
HEADERS = {"User-Agent": "Mozilla/5.0"}

KEYWORDS = [
    "office", "365", "o365", "m365",
    "sharepoint", "power", "exchange", "teams"
]

NAUKRI_URLS = [
    "https://www.naukri.com/office-365-jobs",
    "https://www.naukri.com/sharepoint-jobs",
    "https://www.naukri.com/exchange-jobs",
    "https://www.naukri.com/powerapps-jobs",
    "https://www.naukri.com/o365-jobs?k=o365",
    "https://www.naukri.com/jobapi/v3/search?keyword=o365&location=India&pageNo=1&noOfResults=40"
]

LINKEDIN_URL = "https://www.linkedin.com/jobs/search/?keywords=Office%20365&location=India"

INDEED_URLS = [
    "https://in.indeed.com/jobs?q=office+365&l=India",
]

MONSTER_URLS = [
    "https://www.foundit.in/search/office-365-jobs-in-india"
]

SHINE_URLS = [
    "https://www.shine.com/job-search/office-365-jobs"
]

# -----------------------
# FETCH FUNCTION WITH DEBUG
# -----------------------
def fetch(url):
    print(f"\n🔵 Fetching URL: {url}")
    try:
        r = requests.get(url, headers=HEADERS, timeout=20)
        print("   ↳ Status code:", r.status_code)

        snippet = r.text[:300].replace("\n", " ").replace("\t", " ")
        print("   ↳ HTML Snippet:", snippet[:200], "...\n")

        if r.status_code == 200:
            return r.text
        else:
            return ""
    except Exception as e:
        print("   ❌ Fetch error:", e)
        return ""

# -----------------------
# LINKEDIN PARSER
# -----------------------
def parse_linkedin(html):
    print("🔍 Parsing LinkedIn...")
    if not html:
        print("   ❌ No HTML received.")
        return []

    soup = BeautifulSoup(html, "html.parser")
    jobs = []

    for a in soup.select("a[href]"):
        title = a.get_text(strip=True)
        href = a.get("href")

        if not title or not href:
            continue

        if any(k in title.lower() for k in KEYWORDS):
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

    print(f"   ✔ Found {len(jobs)} LinkedIn jobs.")
    return jobs

# -----------------------
# NAUKRI PARSER
# -----------------------
def parse_naukri(html):
    print("🔍 Parsing Naukri...")
    if not html:
        print("   ❌ No HTML received.")
        return []

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

        if not any(k in title.lower() for k in KEYWORDS):
            continue

        loc_tag = job.select_one("li.loc")
        company_tag = job.select_one("li.company")

        jobs.append({
            "title": title,
            "company": company_tag.get_text(strip=True) if company_tag else "",
            "location": loc_tag.get_text(strip=True) if loc_tag else "",
            "url": href,
            "source": "Naukri"
        })

    print(f"   ✔ Found {len(jobs)} Naukri jobs.")
    return jobs

# -----------------------
# INDEED PARSER
# -----------------------
def parse_indeed(html):
    print("🔍 Parsing Indeed...")
    if not html:
        print("   ❌ No HTML received.")
        return []

    soup = BeautifulSoup(html, "html.parser")
    jobs = []

    for div in soup.select("div.cardOutline"):
        title_tag = div.select_one("h2.jobTitle")
        if not title_tag:
            continue

        title = title_tag.get_text(strip=True)

        if not any(k in title.lower() for k in KEYWORDS):
            continue

        link = div.select_one("a")
        href = "https://in.indeed.com" + link.get("href") if link else ""

        company = div.select_one("span.companyName")
        location = div.select_one("div.companyLocation")

        jobs.append({
            "title": title,
            "company": company.get_text(strip=True) if company else "",
            "location": location.get_text(strip=True) if location else "",
            "url": href,
            "source": "Indeed"
        })

    print(f"   ✔ Found {len(jobs)} Indeed jobs.")
    return jobs

# -----------------------
# MONSTER PARSER
# -----------------------
def parse_monster(html):
    print("🔍 Parsing Monster...")
    if not html:
        print("   ❌ No HTML received.")
        return []

    soup = BeautifulSoup(html, "html.parser")
    jobs = []

    for job in soup.select("div.srp-result-card"):
        title_tag = job.select_one("h3")
        if not title_tag:
            continue

        title = title_tag.get_text(strip=True)

        if not any(k in title.lower() for k in KEYWORDS):
            continue

        link = job.select_one("a")
        href = link.get("href") if link else ""

        company = job.select_one("span.company-name")
        location = job.select_one("span.job-location")

        jobs.append({
            "title": title,
            "company": company.get_text(strip=True) if company else "",
            "location": location.get_text(strip=True) if location else "",
            "url": href,
            "source": "Monster"
        })

    print(f"   ✔ Found {len(jobs)} Monster jobs.")
    return jobs

# -----------------------
# SHINE PARSER
# -----------------------
def parse_shine(html):
    print("🔍 Parsing Shine...")
    if not html:
        print("   ❌ No HTML received.")
        return []

    soup = BeautifulSoup(html, "html.parser")
    jobs = []

    for job in soup.select("li.jobCard"):
        title_tag = job.select_one("strong.title")
        if not title_tag:
            continue

        title = title_tag.get_text(strip=True)

        if not any(k in title.lower() for k in KEYWORDS):
            continue

        link = job.select_one("a")
        href = "https://www.shine.com" + link.get("href") if link else ""

        company = job.select_one("span.jobCardCompanyName")
        location = job.select_one("span.jobCardJobLocation")

        jobs.append({
            "title": title,
            "company": company.get_text(strip=True) if company else "",
            "location": location.get_text(strip=True) if location else "",
            "url": href,
            "source": "Shine"
        })

    print(f"   ✔ Found {len(jobs)} Shine jobs.")
    return jobs

# -----------------------
# EXPORT TO EXCEL
# -----------------------
def export_excel(jobs):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Jobs"

    ws.append(["Title", "Company", "Location", "URL", "Source"])

    for j in jobs:
        ws.append([
            j["title"], j["company"], j["location"], j["url"], j["source"]
        ])

    wb.save("weekly_jobs.xlsx")
    print(f"\n🔥 Saved {len(jobs)} total jobs to weekly_jobs.xlsx\n")

# -----------------------
# MAIN
# -----------------------
def main():
    all_jobs = []

    # Naukri
    for url in NAUKRI_URLS:
        html = fetch(url)
        all_jobs += parse_naukri(html)

    # LinkedIn
    html = fetch(LINKEDIN_URL)
    all_jobs += parse_linkedin(html)

    # Indeed
    for url in INDEED_URLS:
        html = fetch(url)
        all_jobs += parse_indeed(html)

    # Monster
    for url in MONSTER_URLS:
        html = fetch(url)
        all_jobs += parse_monster(html)

    # Shine
    for url in SHINE_URLS:
        html = fetch(url)
        all_jobs += parse_shine(html)

    export_excel(all_jobs)

# Execute script
if __name__ == "__main__":
    main()
