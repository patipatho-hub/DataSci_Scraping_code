# -*- coding: utf-8 -*-
"""
workventure_scraper.py  -  เก็บประกาศงานจาก Workventure (ไอที / เขียนโปรแกรม) -> job_data_workventure.xlsx

ใช้ร่วมกับ job_scraper.py (ต้องวางไว้โฟลเดอร์เดียวกัน) เพื่อใช้ฟังก์ชัน parse เดิม
และ export เป็น Excel รูปแบบเดียวกัน (sheet: jobs / traceability / skipped)

ติดตั้งครั้งแรก:
    pip install playwright
    playwright install chromium

วิธีใช้:
    python workventure_scraper.py            # เก็บข้อมูลจริง
    python workventure_scraper.py debug      # ดูว่าหน้าเว็บมีข้อมูล/ลิงก์อะไรให้ดึง (รันอันนี้ก่อนเสมอ)

เหตุผลที่ต้องใช้ Playwright: หน้ารายการงานของ Workventure วาดด้วย JavaScript
(requests ธรรมดาจะได้ HTML เปล่า) สคริปต์จึงเปิดเบราว์เซอร์จริงแบบ headless แล้วอ่านข้อมูล
จาก JSON-LD (schema.org JobPosting) ในหน้า ซึ่งเป็นข้อมูลที่เว็บเตรียมไว้ให้เครื่องมือค้นหาอ่านอยู่แล้ว

หลักการ: ตรวจ robots.txt ทุก URL, เว้นระยะระหว่างหน้า, เก็บเฉพาะข้อมูลประกาศงาน (ไม่เก็บ resume/อีเมล/เบอร์โทร)
"""
import json
import logging
import re
import sys
import time
import random
from datetime import datetime
from urllib.parse import urljoin, urlparse, urlencode, parse_qsl, urlunparse

from bs4 import BeautifulSoup

# ใช้ฟังก์ชัน parse + export เดิมจาก job_scraper.py
import job_scraper as js

# ============================ CONFIG ============================
START_URL = ("https://www.workventure.com/jobs/function/"
             "%E0%B9%84%E0%B8%AD%E0%B8%97%E0%B8%B5-%E0%B9%80%E0%B8%82%E0%B8%B5%E0%B8%A2%E0%B8%99"
             "%E0%B9%82%E0%B8%9B%E0%B8%A3%E0%B9%81%E0%B8%81%E0%B8%A3%E0%B8%A1")
PAGE_PARAM = "page"            # พารามิเตอร์เลขหน้า (?page=2) ถ้าเว็บใช้ชื่ออื่นให้แก้ตรงนี้
MAX_PAGES = 11                 # ~20 งาน/หน้า (เว็บแจ้ง 215 ตำแหน่งตอนที่ผมดู) ปรับได้
MAX_JOBS = 250
DELAY_SECONDS = 10             # เว้นระยะระหว่างหน้า (ใช้ค่าที่มากกว่า Crawl-delay ใน robots.txt)
OPEN_DETAIL_IF_MISSING = True  # เปิดหน้ารายละเอียดงาน ถ้าจาก JSON-LD ยังไม่ได้ years/education
OUTPUT_FILE = "job_data_workventure.xlsx"
SAVE_EVERY = 20
SITE_NAME = "Workventure"
HEADLESS = True
# ================================================================

log = logging.getLogger("workventure")
logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")

EMP_MAP = {
    "FULL_TIME": "Full-time", "PART_TIME": "Part-time", "CONTRACTOR": "Contract", "CONTRACT": "Contract",
    "TEMPORARY": "Contract", "INTERN": "Internship", "INTERNSHIP": "Internship", "FREELANCE": "Contract",
}


# ============================ browser ============================
class Browser:
    """ครอบ Playwright: เปิดหน้า, รอให้ JS วาดเสร็จ, คืน HTML ที่ render แล้ว + เก็บรายการ XHR/JSON ไว้ debug"""

    def __init__(self):
        try:
            from playwright.sync_api import sync_playwright
        except ImportError:
            sys.exit("ยังไม่ได้ติดตั้ง Playwright:  pip install playwright && playwright install chromium")
        self._pw = sync_playwright().start()
        self._browser = self._pw.chromium.launch(headless=HEADLESS)
        ctx = self._browser.new_context(user_agent=js.USER_AGENT, locale="th-TH")
        self.page = ctx.new_page()
        self.json_calls = []
        self.page.on("response", self._on_response)

    def _on_response(self, resp):
        try:
            if "json" in (resp.headers.get("content-type") or "") and resp.request.resource_type in ("xhr", "fetch"):
                self.json_calls.append(resp.url)
        except Exception:
            pass

    def get(self, url):
        if not js.robots.allowed(url):
            log.warning("ถูกห้ามโดย robots.txt ข้าม: %s", url)
            return None
        time.sleep(js.robots.delay(url) + random.uniform(0, 3))   # delay() = max(DELAY_SECONDS ของ job_scraper, Crawl-delay)
        try:
            self.page.goto(url, wait_until="networkidle", timeout=60000)
            # เลื่อนลงเพื่อให้ lazy-load ครบ
            for _ in range(4):
                self.page.mouse.wheel(0, 4000)
                self.page.wait_for_timeout(800)
            return self.page.content()
        except Exception as e:
            log.warning("เปิดหน้าไม่สำเร็จ %s: %s", url, e)
            return None

    def close(self):
        try:
            self._browser.close()
            self._pw.stop()
        except Exception:
            pass


# ============================ parse JSON-LD ============================
def _walk(obj):
    """เดินหา dict ที่เป็น JobPosting ในโครงสร้าง JSON-LD ที่ซ้อนกัน (list / @graph / ItemList)"""
    if isinstance(obj, list):
        for x in obj:
            yield from _walk(x)
    elif isinstance(obj, dict):
        t = obj.get("@type")
        types = t if isinstance(t, list) else [t]
        if "JobPosting" in types or ("title" in obj and ("hiringOrganization" in obj or "occupationalCategory" in obj)):
            yield obj
        for k in ("@graph", "itemListElement", "item", "mainEntity"):
            if k in obj:
                yield from _walk(obj[k])


def jobposting_items(soup):
    out = []
    for tag in soup.find_all("script", type="application/ld+json"):
        try:
            data = json.loads(tag.string or tag.get_text() or "")
        except (ValueError, TypeError):
            continue
        out.extend(_walk(data))
    return out


def _s(x):
    return x.get("name") if isinstance(x, dict) else x


def item_to_fields(item):
    """แปลง JobPosting 1 รายการเป็นค่าที่ใช้ต่อ"""
    org = item.get("hiringOrganization")
    loc = item.get("jobLocation")
    loc = loc[0] if isinstance(loc, list) and loc else loc
    location = None
    if isinstance(loc, dict):
        a = loc.get("address") or {}
        if isinstance(a, dict):
            location = ", ".join(str(_s(v)) for v in (a.get("addressLocality"), a.get("addressRegion")) if v) or None
        elif isinstance(a, str):
            location = a
    sal = item.get("baseSalary") or {}
    val = sal.get("value", sal) if isinstance(sal, dict) else {}
    smin = smax = None
    if isinstance(val, dict):
        smin, smax = val.get("minValue"), val.get("maxValue")
        if smin is None and isinstance(val.get("value"), (int, float)):
            smin = val["value"]
    elif isinstance(val, (int, float)):
        smin = val
    emp = item.get("employmentType")
    emp = emp[0] if isinstance(emp, list) and emp else emp
    desc = item.get("description") or ""
    if "<" in desc:                                   # description บางเว็บเป็น HTML
        desc = BeautifulSoup(desc, "lxml").get_text("\n")
    return {
        "title": item.get("title"),
        "url": item.get("url") or (item.get("mainEntityOfPage") if isinstance(item.get("mainEntityOfPage"), str) else None),
        "company": _s(org) if org else None,
        "location": location,
        "salary_min": smin, "salary_max": smax,
        "employment": EMP_MAP.get(str(emp).upper().replace("-", "_")) if emp else None,
        "category": item.get("occupationalCategory"),
        "skills": item.get("skills"),
        "remote": "Remote" if item.get("jobLocationType") == "TELECOMMUTE" else None,
        "description": desc,
    }


def find_detail_link(soup, title, base_url):
    """หา <a> ของการ์ดงานที่ข้อความตรงกับชื่อตำแหน่ง (ไม่ต้องเดา pattern ของ URL ประกาศ)"""
    if not title:
        return None
    t = title.strip().lower()
    for a in soup.select("a[href]"):
        if t and t in a.get_text(" ", strip=True).lower():
            href = urljoin(base_url, a["href"]).split("#")[0]
            if urlparse(href).netloc.endswith("workventure.com"):
                return href
    return None


# ============================ แปลงเป็นแถวข้อมูล ============================
def build_row(f, link):
    desc = f["description"] or ""
    years, years_src = js.parse_years(desc, None)
    lo, hi = f["salary_min"], f["salary_max"]
    mid = None
    if lo or hi:
        lo = lo or hi
        hi = hi or lo
        mid = (float(lo) + float(hi)) / 2
    neg = js.NEGOTIABLE.search(desc) if desc else None
    row = {
        "job_title": f["title"],
        "job_link": link,
        "salary/THB": mid if mid is not None else ("สามารถตกลงได้" if neg else None),
        "company_location": f["location"],
        "company_size": None,
        "years_experience": years if years is not None else "ไม่ระบุ",
        "experience_level": js.parse_level(f["title"] or "", None, years),
        "education_level": js.parse_education(desc),
        "company_name": f["company"],
        "industry": None,
        "remote_status": f["remote"] or js.parse_remote(desc),
        "employment_type": f["employment"] or js.parse_employment(f["title"], None),
        "ข้อมูลจาก": SITE_NAME,
    }
    extra = {
        "job_url": link, "salary_min": f["salary_min"], "salary_max": f["salary_max"],
        "salary_raw": None, "experience_raw": None,
        "years_source": years_src or "not found", "status": "open",
        "category": f["category"], "skills": f["skills"],
        "scraped_at": datetime.now().strftime("%Y-%m-%d %H:%M"),
    }
    return row, extra


def enrich_from_detail(browser, row, extra, link):
    """เปิดหน้ารายละเอียดงาน เพื่อเติม years/education/company_size/industry ที่ JSON-LD หน้ารายการไม่มี"""
    html = browser.get(link)
    if not html:
        return
    soup = BeautifulSoup(html, "lxml")
    items = jobposting_items(soup)
    desc = ""
    if items:
        f = item_to_fields(items[0])
        desc = f["description"]
        if f["salary_min"] and row["salary/THB"] in (None, "สามารถตกลงได้"):
            lo, hi = f["salary_min"], f["salary_max"] or f["salary_min"]
            row["salary/THB"] = (float(lo) + float(hi)) / 2
        row["company_location"] = row["company_location"] or f["location"]
        row["employment_type"] = row["employment_type"] or f["employment"]
    lines = js.get_lines(BeautifulSoup(html, "lxml"))
    text = "\n".join(js.body_lines(lines)) if lines else ""
    blob = desc or text
    if row["years_experience"] == "ไม่ระบุ":
        exp_text = js.find_value(lines, js.LABELS["experience"]) or js.regex_field(lines, js.EXP_RE)
        yrs, src = js.parse_years(blob, exp_text)
        if yrs is not None:
            row["years_experience"], extra["years_source"] = yrs, src
            row["experience_level"] = js.parse_level(row["job_title"] or "", None, yrs)
        extra["experience_raw"] = exp_text
    row["education_level"] = row["education_level"] or js.parse_education(blob)
    row["company_size"] = row["company_size"] or js.find_value(lines, js.LABELS["company_size"])
    row["industry"] = row["industry"] or js.find_value(lines, js.LABELS["industry"])


# ============================ เก็บรายการ ============================
def page_url(base, n):
    if n == 1:
        return base
    p = urlparse(base)
    q = dict(parse_qsl(p.query))
    q[PAGE_PARAM] = str(n)
    return urlunparse(p._replace(query=urlencode(q)))


def main():
    rows, extras, skipped, seen = [], [], [], set()
    b = Browser()
    try:
        for n in range(1, MAX_PAGES + 1):
            url = page_url(START_URL, n)
            log.info("หน้า %s: %s", n, url)
            html = b.get(url)
            if not html:
                skipped.append({"job_url": url, "site": SITE_NAME, "reason": "เปิดหน้ารายการไม่ได้ (robots/timeout)"})
                break
            soup = BeautifulSoup(html, "lxml")
            items = jobposting_items(soup)
            if not items:
                with open(f"debug_workventure_page{n}.html", "w", encoding="utf-8") as fh:
                    fh.write(html)
                log.warning("หน้า %s ไม่พบ JSON-LD JobPosting (บันทึก HTML ไว้ debug_workventure_page%s.html) "
                            "รัน 'python workventure_scraper.py debug' เพื่อตรวจ", n, n)
                break
            new = 0
            for it in items:
                f = item_to_fields(it)
                link = f["url"] or find_detail_link(soup, f["title"], url) or url
                key = (link, f["title"], f["company"])
                if key in seen or not f["title"]:
                    continue
                seen.add(key)
                new += 1
                row, extra = build_row(f, link)
                need = row["years_experience"] == "ไม่ระบุ" or not row["education_level"]
                if OPEN_DETAIL_IF_MISSING and need and link != url:
                    try:
                        enrich_from_detail(b, row, extra, link)
                    except Exception as e:
                        log.exception("เติมข้อมูลจากหน้ารายละเอียดพัง %s: %s", link, e)
                        skipped.append({"job_url": link, "site": SITE_NAME, "reason": f"detail error: {e}"})
                rows.append(row)
                extras.append(extra)
                if len(rows) % SAVE_EVERY == 0:
                    js.export_excel(rows, extras, skipped, path=OUTPUT_FILE)
                if len(rows) >= MAX_JOBS:
                    break
            log.info("หน้า %s ได้ %s งานใหม่ (รวม %s)", n, new, len(rows))
            if new == 0 or len(rows) >= MAX_JOBS:
                break
    finally:
        b.close()
    js.export_excel(rows, extras, skipped, path=OUTPUT_FILE)


def debug():
    """ตรวจหน้าแรก: มี JSON-LD ไหม, เว็บเรียก API (XHR) ที่ไหนบ้าง, ลิงก์ที่น่าจะเป็นหน้าประกาศงาน"""
    b = Browser()
    try:
        html = b.get(START_URL)
        if not html:
            print("เปิดหน้าไม่ได้ (robots.txt ห้าม หรือ timeout)")
            return
        open("debug_workventure.html", "w", encoding="utf-8").write(html)
        soup = BeautifulSoup(html, "lxml")
        items = jobposting_items(soup)
        print("ความยาว HTML ที่ render แล้ว:", len(html))
        print("จำนวน JobPosting ใน JSON-LD:", len(items))
        for it in items[:3]:
            print("  ตัวอย่าง:", json.dumps(item_to_fields(it), ensure_ascii=False)[:300])
        print("\nXHR/fetch ที่คืน JSON (อาจใช้เป็นแหล่งข้อมูลแทนได้):")
        for u in dict.fromkeys(b.json_calls):
            print("  ", u)
        print("\nลิงก์ที่น่าจะเป็นหน้าประกาศงาน:")
        hrefs = [urljoin(START_URL, a["href"]) for a in soup.select("a[href]")]
        cand = [h for h in dict.fromkeys(hrefs) if re.search(r"/jobs?/", h) and "/function/" not in h and "/for/" not in h]
        for h in cand[:15]:
            print("  ", h)
        print("\nบันทึก HTML ไว้ที่ debug_workventure.html")
    finally:
        b.close()


if __name__ == "__main__":
    if len(sys.argv) >= 2 and sys.argv[1] == "debug":
        debug()
    else:
        main()
