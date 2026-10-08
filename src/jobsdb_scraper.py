# -*- coding: utf-8 -*-
"""
jobsdb_scraper.py  -  เก็บประกาศงานจาก JobsDB Thailand (หน้ารายการงาน) -> job_data_jobsdb.xlsx

วางไว้โฟลเดอร์เดียวกับ job_scraper.py (ใช้ฟังก์ชัน parse + export เดิม)
รัน:  python jobsdb_scraper.py          เก็บข้อมูลจริง
      python jobsdb_scraper.py debug    ตรวจหน้าเว็บ (รันก่อนเสมอ)

!!! สำคัญเรื่อง robots.txt ของ th.jobsdb.com (User-agent: *)
    - Disallow */job/        -> ห้ามเปิดหน้ารายละเอียดงาน (/th/job/123456) สคริปต์นี้จึง "ไม่เปิด" หน้านั้นเลย
                                ข้อมูลทั้งหมดมาจากการ์ดงานในหน้ารายการเท่านั้น (ปีประสบการณ์จึงได้เท่าที่การ์ดบอก)
                                job_link ที่ใส่ในไฟล์ไว้ให้คุณคลิกเปิดดูเอง ไม่ได้ให้สคริปต์ดึง
    - Disallow *?            -> ห้าม URL ที่มี query string (?page=2) ยกเว้น Allow *?keywords และ *?advertiserid
                                กฎที่ "ยาวกว่า" ชนะ ดังนั้น ?keywords=... ผ่าน แต่ ?page=2 เฉยๆ ไม่ผ่าน
    - python urllib.robotparser ไม่รองรับ wildcard (*) จึงตรวจผิดกับไฟล์นี้ สคริปต์นี้เลยมีตัวตรวจของตัวเอง

ถ้าเว็บตอบกลับเป็นหน้า bot detection / 403 สคริปต์จะหยุดทันที ไม่พยายามหลบการตรวจจับ
"""
import json
import logging
import random
import re
import sys
import time
from datetime import datetime
from urllib.parse import urljoin, urlparse

import requests
from bs4 import BeautifulSoup

import job_scraper as js

# ============================ CONFIG ============================
# หน้าแรกของแต่ละลิสต์ (ไม่มี query string)
START_URLS = [
    "https://th.jobsdb.com/th/Data-Analyst-jobs-in-information-communication-technology",
]
# หน้า 2 ขึ้นไป: ใช้ ?keywords=...&page=N (ผ่าน Allow *?keywords) -- สคริปต์จะตรวจ robots ทุก URL
# ถ้าตรวจแล้วห้าม จะหยุดที่หน้า 1 เอง  ตั้งเป็น None เพื่อเก็บแค่หน้าแรก
PAGED_URL = "https://th.jobsdb.com/th/Data-Analyst-jobs-in-information-communication-technology?keywords=Data+Analyst&page={page}"
MAX_PAGES = 5
MAX_JOBS = 150
DELAY_SECONDS = 10
OUTPUT_FILE = "job_data_jobsdb.xlsx"
SITE_NAME = "JobsDB"
SAVE_EVERY = 20
# ================================================================

log = logging.getLogger("jobsdb")
logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")


# ============================ robots.txt แบบรองรับ wildcard ============================
class WildRobots:
    """ตรวจ robots.txt ตามหลัก Google: รองรับ * และ $, กฎที่ pattern ยาวกว่าชนะ, ความยาวเท่ากัน Allow ชนะ"""

    def __init__(self, session, user_agent):
        self.session, self.ua, self.cache = session, user_agent.lower(), {}

    def _rules(self, base):
        if base in self.cache:
            return self.cache[base]
        rules = []
        try:
            r = self.session.get(base + "/robots.txt", timeout=15)
            text = r.text if r.status_code == 200 else ""
        except requests.RequestException:
            text = ""
        rules = self.parse(text, self.ua)
        self.cache[base] = rules
        log.info("loaded robots.txt: %s (%s กฎ)", base, len(rules))
        return rules

    @staticmethod
    def parse(text, ua="*"):
        groups, agents, rules, last_was_agent = [], [], [], False
        for raw in text.splitlines():
            line = raw.split("#", 1)[0].strip()
            if ":" not in line:
                continue
            k, v = (x.strip() for x in line.split(":", 1))
            k = k.lower()
            if k == "user-agent":
                if not last_was_agent and agents:
                    groups.append((agents, rules)); agents, rules = [], []
                agents.append(v.lower()); last_was_agent = True
            elif k in ("allow", "disallow"):
                last_was_agent = False
                if v:                                   # Disallow: (ว่าง) = อนุญาตทั้งหมด ไม่ต้องเก็บกฎ
                    rules.append((k == "allow", v))
        if agents:
            groups.append((agents, rules))
        chosen = None
        for ag, rl in groups:                           # กลุ่มที่ชื่อ bot ตรงกับ UA ของเรา (ถ้ามี)
            if any(a != "*" and a in ua for a in ag):
                chosen = rl
        if chosen is None:
            chosen = [r for ag, rl in groups if "*" in ag for r in rl]
        return chosen

    @staticmethod
    def _match(pattern, target):
        rx = re.escape(pattern).replace(r"\*", ".*")
        if rx.endswith(r"\$"):
            rx = rx[:-2] + "$"
        return re.match(rx, target) is not None

    def allowed(self, url):
        p = urlparse(url)
        rules = self._rules(f"{p.scheme}://{p.netloc}")
        target = p.path + (("?" + p.query) if p.query else "")
        best_len, best_allow = -1, True
        for allow, pat in rules:
            if self._match(pat, target):
                L = len(pat)
                if L > best_len or (L == best_len and allow):
                    best_len, best_allow = L, allow
        return best_allow


session = requests.Session()
session.headers.update({"User-Agent": js.USER_AGENT, "Accept-Language": "th,en;q=0.8"})
robots = WildRobots(session, js.USER_AGENT)

BLOCK_HINT = re.compile(r"captcha|access denied|just a moment|verify you are human|unusual traffic|bot detection", re.I)


def fetch(url):
    """ดึงหน้ารายการอย่างสุภาพ; คืน (html, สถานะ) สถานะ: ok | robots | blocked | error"""
    if not robots.allowed(url):
        log.warning("ถูกห้ามโดย robots.txt ข้าม: %s", url)
        return None, "robots"
    time.sleep(DELAY_SECONDS + random.uniform(0, 3))
    try:
        r = session.get(url, timeout=30)
    except requests.RequestException as e:
        log.warning("request error: %s", e)
        return None, "error"
    if r.status_code in (401, 403, 429) or (r.status_code == 200 and BLOCK_HINT.search(r.text[:5000])):
        log.warning("เว็บบล็อกคำขอ (HTTP %s / bot detection) หยุดเก็บ ไม่พยายามหลบ: %s", r.status_code, url)
        return None, "blocked"
    if r.status_code != 200:
        log.warning("HTTP %s ข้าม: %s", r.status_code, url)
        return None, "error"
    r.encoding = r.apparent_encoding or "utf-8"
    return r.text, "ok"


# ============================ parse การ์ดงาน ============================
def _find_job_dicts(obj, out):
    """เดินหา dict ที่หน้าตาเหมือนประกาศงาน (มี title + ชื่อบริษัท/คำอธิบายสั้น) ใน JSON ที่ฝังในหน้า"""
    if isinstance(obj, list):
        for x in obj:
            _find_job_dicts(x, out)
    elif isinstance(obj, dict):
        title = obj.get("title") or obj.get("jobTitle")
        if isinstance(title, str) and (obj.get("companyName") or obj.get("advertiser") or obj.get("hiringOrganization")) \
                and (obj.get("id") or obj.get("jobId") or obj.get("url")):
            out.append(obj)
        for v in obj.values():
            if isinstance(v, (dict, list)):
                _find_job_dicts(v, out)


def jobs_from_json(soup):
    out = []
    for tag in soup.find_all("script"):
        txt = tag.string or ""
        if not txt or "{" not in txt:
            continue
        cands = []
        if tag.get("type") in ("application/ld+json", "application/json"):
            cands.append(txt)
        else:
            for m in re.finditer(r"(?:SEEK_REDUX_DATA|__NEXT_DATA__|SEEK_APOLLO_DATA)\s*=\s*(\{.*\})\s*;?\s*$", txt, re.S):
                cands.append(m.group(1))
        for c in cands:
            try:
                _find_job_dicts(json.loads(c), out)
            except ValueError:
                continue
    return out


def _lab(x):
    if isinstance(x, dict):
        return x.get("label") or x.get("name") or x.get("description")
    return x


def norm_json_job(d):
    adv = d.get("advertiser") or d.get("hiringOrganization") or {}
    locs = d.get("locations") or d.get("location") or []
    locs = locs if isinstance(locs, list) else [locs]
    cls = d.get("classifications") or []
    cls_txt = ", ".join(
        str(_lab((c.get("classification") if isinstance(c, dict) else c) or {})) for c in cls if c) or d.get("occupationalCategory")
    wt = d.get("workTypes") or d.get("employmentType") or []
    jid = d.get("id") or d.get("jobId")
    return {
        "title": d.get("title") or d.get("jobTitle"),
        "company": d.get("companyName") or _lab(adv),
        "location": ", ".join(str(_lab(l)) for l in locs if l) or None,
        "salary": d.get("salaryLabel") or d.get("salary"),
        "text": " ".join([str(d.get("teaser") or "")] + [str(b) for b in (d.get("bulletPoints") or [])]),
        "work_type": " ".join(str(w) for w in (wt if isinstance(wt, list) else [wt])),
        "classification": cls_txt,
        "link": d.get("url") or (f"https://th.jobsdb.com/th/job/{jid}" if jid else None),   # ลิงก์ไว้ให้คลิกเอง ไม่ถูกดึง
    }


def jobs_from_dom(soup, base_url):
    """สำรอง: อ่านจากการ์ดงานตาม attribute data-automation (ใช้ได้ถ้าหน้ามีการ์ดใน HTML)"""
    def txt(card, name):
        el = card.select_one(f'[data-automation="{name}"]')
        return el.get_text(" ", strip=True) if el else None
    out = []
    for card in soup.select('[data-automation="normalJob"], [data-automation="premiumJob"], article[data-card-type="JobCard"]'):
        a = card.select_one('a[data-automation="jobTitle"], a[href*="/job/"]')
        out.append({
            "title": txt(card, "jobTitle"), "company": txt(card, "jobCompany"),
            "location": txt(card, "jobLocation"), "salary": txt(card, "jobSalary"),
            "text": " ".join(filter(None, [txt(card, "jobShortDescription"), card.get_text(" ", strip=True)])),
            "work_type": card.get_text(" ", strip=True), "classification": txt(card, "jobClassification"),
            "link": urljoin(base_url, a["href"]).split("?")[0] if a and a.get("href") else None,
        })
    return out


def parse_listing(html, url):
    soup = BeautifulSoup(html, "lxml")
    jobs = [norm_json_job(d) for d in jobs_from_json(soup)]
    src = "json"
    if not jobs:
        jobs, src = jobs_from_dom(soup, url), "dom"
    return [j for j in jobs if j.get("title")], src


EMP = [("Internship", r"intern|ฝึกงาน|สหกิจ"), ("Part-time", r"part[\s-]?time|พาร์ท?ไทม์|พาร์ทไทม์|casual"),
       ("Contract", r"contract|สัญญาจ้าง|ชั่วคราว|freelance"), ("Full-time", r"full[\s-]?time|เต็มเวลา|งานประจำ")]


def to_row(j):
    text = f"{j['title']} {j['text']}"
    sal_text = j.get("salary") or ""
    lo, hi, mid = js.parse_salary(sal_text)
    if mid is not None and re.search(r"year|annum|ต่อปี|/ปี", sal_text, re.I):
        mid, lo, hi = mid / 12, lo / 12, hi / 12                     # แปลงเป็นรายเดือน
    elif mid is not None and re.search(r"hour|day|ชั่วโมง|วัน", sal_text, re.I):
        mid = lo = hi = None                                         # รายชั่วโมง/รายวัน ไม่เทียบกับเงินเดือน
    if mid is None and js.NEGOTIABLE.search(sal_text):
        sal_val = "สามารถตกลงได้"
    else:
        sal_val = mid
    years, years_src = js.parse_years(j["text"], None)
    emp = next((n for n, p in EMP if re.search(p, j["work_type"], re.I)), None) or js.parse_employment(j["title"], None)
    row = {
        "job_title": j["title"], "job_link": j["link"], "salary/THB": sal_val,
        "company_location": j["location"], "company_size": None,
        "years_experience": years if years is not None else "ไม่ระบุ",
        "experience_level": js.parse_level(j["title"], None, years),
        "education_level": js.parse_education(j["text"]),
        "company_name": j["company"], "industry": None,
        "remote_status": js.parse_remote(f"{text} {j['location'] or ''}"),
        "employment_type": emp, "ข้อมูลจาก": SITE_NAME,
    }
    extra = {"job_url": j["link"], "salary_min": lo, "salary_max": hi, "salary_raw": sal_text or None,
             "experience_raw": None, "years_source": years_src or "not found", "status": "open",
             "classification": j["classification"], "scraped_at": datetime.now().strftime("%Y-%m-%d %H:%M")}
    return row, extra


# ============================ main / debug ============================
def page_urls():
    for base in START_URLS:
        yield 1, base
        if PAGED_URL and "Data-Analyst" in base:
            for n in range(2, MAX_PAGES + 1):
                yield n, PAGED_URL.format(page=n)


def main():
    rows, extras, skipped, seen = [], [], [], set()
    for n, url in page_urls():
        if len(rows) >= MAX_JOBS:
            break
        log.info("หน้า %s: %s", n, url)
        html, status = fetch(url)
        if status == "blocked":
            skipped.append({"job_url": url, "site": SITE_NAME, "reason": "bot detection / HTTP 403-429"})
            break
        if not html:
            skipped.append({"job_url": url, "site": SITE_NAME, "reason": f"ดึงไม่ได้ ({status})"})
            if n == 1:
                continue
            break
        jobs, src = parse_listing(html, url)
        if not jobs:
            open(f"debug_jobsdb_page{n}.html", "w", encoding="utf-8").write(html)
            log.warning("หน้า %s ไม่พบการ์ดงาน บันทึก HTML ไว้ debug_jobsdb_page%s.html (รัน debug เพื่อตรวจ)", n, n)
            break
        new = 0
        for j in jobs:
            key = j["link"] or (j["title"], j["company"])
            if key in seen:
                continue
            seen.add(key); new += 1
            row, extra = to_row(j)
            rows.append(row); extras.append(extra)
            if len(rows) % SAVE_EVERY == 0:
                js.export_excel(rows, extras, skipped, path=OUTPUT_FILE)
        log.info("หน้า %s (%s) ได้ %s งานใหม่ (รวม %s)", n, src, new, len(rows))
        if new == 0:
            break
    js.export_excel(rows, extras, skipped, path=OUTPUT_FILE)


def debug():
    for u in START_URLS[:1]:
        print("robots.txt อนุญาตหน้ารายการ:", robots.allowed(u))
        print("robots.txt อนุญาตหน้ารายละเอียดงาน (/th/job/123):", robots.allowed("https://th.jobsdb.com/th/job/123"))
        print("robots.txt อนุญาต ?page=2:", robots.allowed(u + "?page=2"))
        if PAGED_URL:
            print("robots.txt อนุญาต PAGED_URL:", robots.allowed(PAGED_URL.format(page=2)))
        html, status = fetch(u)
        print("สถานะการดึง:", status)
        if not html:
            return
        open("debug_jobsdb.html", "w", encoding="utf-8").write(html)
        jobs, src = parse_listing(html, u)
        print("ความยาว HTML:", len(html), "| อ่านการ์ดงานได้:", len(jobs), "จาก", src)
        for j in jobs[:3]:
            print("  ", json.dumps(j, ensure_ascii=False)[:300])
        print("บันทึก HTML ไว้ที่ debug_jobsdb.html")


if __name__ == "__main__":
    debug() if len(sys.argv) >= 2 and sys.argv[1] == "debug" else main()
