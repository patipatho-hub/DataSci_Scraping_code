# -*- coding: utf-8 -*-
"""
jobsdb_manual_fill.py  -  ช่วยกรอก years_experience / education_level ของ JobsDB จากข้อความ "คุณสมบัติ (Qualification)"
                          ที่คุณ "เปิดอ่านและคัดลอกเอง" จากเบราว์เซอร์

สคริปต์นี้ "ไม่ได้เชื่อมต่ออินเทอร์เน็ตและไม่ได้เปิดหน้า JobsDB" เลย
(หน้ารายละเอียดงาน /th/job/... ถูกห้ามโดย robots.txt สำหรับโปรแกรมอัตโนมัติ)
มันแค่รับข้อความที่คุณวาง แล้วอ่านปีประสบการณ์ + วุฒิการศึกษา และเขียนกลับลง Excel ให้

วิธีใช้:
    python jobsdb_manual_fill.py                      # ใช้ไฟล์ job_data_jobsdb.xlsx
    python jobsdb_manual_fill.py ชื่อไฟล์.xlsx

ทีละแถวที่ years_experience = "ไม่ระบุ" หรือ education_level ว่าง สคริปต์จะแสดงชื่อตำแหน่ง + job_link
  1) กดเปิดลิงก์ในเบราว์เซอร์ของคุณเอง   2) คัดลอกหัวข้อ "คุณสมบัติ/Qualification"
  3) วางในหน้าต่างนี้ แล้วพิมพ์ END ในบรรทัดใหม่ + Enter
  (Enter เปล่าๆ = ข้ามแถวนี้ | q = บันทึกแล้วออก)  บันทึกไฟล์ทุกแถว ปิดโปรแกรมกลางทางได้
"""
import re
import sys

import pandas as pd

import job_scraper as js

DEFAULT_FILE = "job_data_jobsdb.xlsx"
NUM_YEARS = re.compile(
    r"(?:อย่างน้อย|ไม่น้อยกว่า|มากกว่า|at least|minimum|min\.?)?\s*(\d{1,2})(?:\s*[-–~]\s*\d{1,2})?\s*\+?\s*(?:ปี|years?|yrs?)")
NUM_MONTHS = re.compile(r"(\d{1,2})\s*(?:เดือน|months?)")


def years_from_qualification(text):
    """คืน (ปีขั้นต่ำ, บรรทัดต้นฉบับที่ใช้) จากข้อความคุณสมบัติ
    - ใช้บรรทัดแรกที่พูดถึง 'ประสบการณ์/experience' และมีตัวเลข+ปี (ช่วง 1-3 ปี -> ใช้ค่าต่ำสุด)
    - ถ้าไม่มีตัวเลข แต่ระบุว่าไม่ต้องมีประสบการณ์/จบใหม่ -> 0
    - ไม่พบอะไรเลย -> (None, None) ปล่อยให้ว่าง ไม่เดา"""
    t = re.sub(r"(\d)\s*\n\s*(ปี|years?|yrs?)", r"\1 \2", text, flags=re.I)   # ตัวเลขกับ 'ปี' ที่ถูกตัดขึ้นบรรทัดใหม่
    for line in t.splitlines():
        low = line.lower()
        if not re.search(r"ประสบการณ์|experience", low):
            continue
        m = NUM_YEARS.search(low)
        if m:
            return int(m.group(1)), line.strip()
        m = NUM_MONTHS.search(low)
        if m:
            return round(int(m.group(1)) / 12, 2), line.strip()
    if re.search(r"ไม่มีประสบการณ์|ไม่จำเป็นต้องมีประสบการณ์|fresh\s*grad|จบใหม่|no experience", t.lower()):
        return 0, None
    return None, None


def read_block():
    lines = []
    while True:
        try:
            s = input()
        except EOFError:
            break
        if s.strip().upper() == "END":
            break
        if not lines and s.strip().lower() == "q" :
            return "q"
        if not lines and s == "":
            return ""
        lines.append(s)
    return "\n".join(lines)


def main(path):
    book = pd.read_excel(path, sheet_name=None)
    jobs = book["jobs"].astype(object).where(book["jobs"].notna(), None)
    trace = book.get("traceability")
    trace = trace.astype(object).where(trace.notna(), None) if trace is not None else None
    skipped = book.get("skipped")
    synced = trace is not None and len(trace) == len(jobs)

    todo = [i for i, r in jobs.iterrows()
            if r["years_experience"] in (None, "ไม่ระบุ") or not r["education_level"]]
    print(f"แถวที่ยังขาดข้อมูล: {len(todo)} จาก {len(jobs)}\n")

    def save():
        rows = jobs.to_dict("records")
        extras = trace.to_dict("records") if trace is not None else []
        sk = skipped.to_dict("records") if skipped is not None else []
        js.export_excel(rows, extras, sk, path=path)

    for n, i in enumerate(todo, 1):
        r = jobs.loc[i]
        print(f"[{n}/{len(todo)}] {r['job_title']} | {r['company_name']}")
        print(f"   ลิงก์: {r['job_link']}")
        print("   วางข้อความคุณสมบัติ แล้วพิมพ์ END (Enter เปล่า=ข้าม, q=ออก):")
        text = read_block()
        if text == "q":
            break
        if not text.strip():
            print("   ข้าม\n")
            continue
        yrs, raw = years_from_qualification(text)
        edu = js.parse_education(text)
        if yrs is not None:
            jobs.at[i, "years_experience"] = yrs
            jobs.at[i, "experience_level"] = js.parse_level(r["job_title"] or "", None, yrs)
        if edu:
            jobs.at[i, "education_level"] = edu
        if synced:
            trace.at[i, "years_source"] = "manual" if yrs is not None else trace.at[i, "years_source"]
            trace.at[i, "experience_raw"] = raw
        print(f"   -> years_experience={yrs if yrs is not None else 'ไม่พบ'} | education_level={edu or 'ไม่พบ'}\n")
        save()
    save()
    print("เสร็จแล้ว บันทึกที่", path)


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else DEFAULT_FILE)
