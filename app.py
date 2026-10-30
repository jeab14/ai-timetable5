# -*- coding: utf-8 -*-
"""
ระบบจัดตารางสอนอัตโนมัติ (AI Timetable Scheduler)
เวอร์ชัน 7.1 (ธีมม่วงพาสเทล + ไอคอน Material): รองรับชุดข้อมูลการแข่งขันทักษะเทคโนโลยีปัญญาประดิษฐ์ (7 ไฟล์)
    teacher / room / student_group / subject / teach / register / timeslot

  - ผู้ดูแลระบบ/ฝ่ายวิชาการ อัปโหลดไฟล์ทั้งหมดพร้อมกันครั้งเดียว (CSV / Excel / ZIP กี่ไฟล์ก็ได้)
    ระบบแยกประเภทไฟล์เอง เติมข้อมูลที่ขาด แล้วจัดตารางให้ทันที
  - ตรวจเงื่อนไขพื้นฐานข้อ 1-8 + เงื่อนไขพิเศษข้อ 9-13 (เลือกตามที่จับสลากได้)
  - ส่งออก output.csv ตามรูปแบบที่กำหนด + Excel / PDF แบบตารางเรียนที่อ่านง่าย
  - เข้าสู่ระบบอย่างเดียว: ครูใช้รหัสครู นักเรียนใช้รหัสกลุ่มเรียน (หรือรหัสนักเรียน)
    รหัสผ่านเริ่มต้นเหมือนกันทุกคน เปลี่ยนเองได้ และผู้ดูแลระบบรีเซ็ตกลับเป็นค่าเริ่มต้นได้

วิธีรัน:
    pip install -U streamlit pandas openpyxl fpdf2
    streamlit run app.py
"""

import hashlib
import hmac
import html as html_lib
import importlib.util
import io
import json
import os
import random
import re
import secrets
import time
import zipfile
from collections import defaultdict
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo

import pandas as pd
import streamlit as st

EXCEL_SUPPORTED = importlib.util.find_spec("openpyxl") is not None
PDF_SUPPORTED = importlib.util.find_spec("fpdf") is not None

st.set_page_config(
    page_title="ระบบจัดตารางสอนอัตโนมัติ (AI Timetable Scheduler)",
    page_icon=":material/calendar_month:",
    layout="wide",
    initial_sidebar_state="expanded",
)

APP_TITLE = "ระบบจัดตารางสอนอัตโนมัติ (AI Timetable Scheduler)"
APP_DIR = Path(__file__).resolve().parent
TZ = ZoneInfo("Asia/Bangkok")

# ---------------------------------------------------------------------------
# ค่าคงที่ตามเงื่อนไขการแข่งขัน (แก้ได้ที่นี่ที่เดียว)
# ---------------------------------------------------------------------------
BREAK_PERIODS = {5}            # ข้อ 2: คาบที่ 5 ของทุกวันเป็นคาบพัก
MAX_PERIODS_PER_DAY = 10       # ข้อ 5: กลุ่มเรียนเรียนได้ไม่เกิน 10 คาบ/วัน
LEADER_MEETING = ("Tue", [8])  # ข้อ 10: หัวหน้า (Leader) ประชุมวันอังคาร คาบ 8
HOMEROOM_PERIODS = 1           # ข้อ 11: โฮมรูมอย่างน้อย 1 คาบ/สัปดาห์
HOMEROOM_ID = "HOMEROOM"
ACTIVITY_SLOT = ("Wed", [8, 9])  # ข้อ 12: วิชากิจกรรม วันพุธ คาบ 8-9
GENERAL_PREFIXES = ("20000", "30000")  # ข้อ 13: รหัสวิชาสามัญ
DEFAULT_DAYS = ["Mon", "Tue", "Wed", "Thu", "Fri"]
DEFAULT_N_PERIODS = 12
DEFAULT_START_HOUR = 8

DAYS = [("Mon", "จันทร์"), ("Tue", "อังคาร"), ("Wed", "พุธ"), ("Thu", "พฤหัสบดี"),
        ("Fri", "ศุกร์"), ("Sat", "เสาร์"), ("Sun", "อาทิตย์")]
DAY_CODES = [d for d, _ in DAYS]
DAY_TH = dict(DAYS)
DAY_ORDER = {d: i for i, d in enumerate(DAY_CODES)}
DAY_ALIASES = {}
for _code, _th in DAYS:
    for _alias in (_code, _th, "วัน" + _th):
        DAY_ALIASES[_alias.lower()] = _code
DAY_ALIASES.update({
    "monday": "Mon", "tuesday": "Tue", "tues": "Tue", "wednesday": "Wed", "thursday": "Thu",
    "thur": "Thu", "thurs": "Thu", "friday": "Fri", "saturday": "Sat", "sunday": "Sun",
    "จ": "Mon", "อ": "Tue", "พ": "Wed", "พฤ": "Thu", "ศ": "Fri", "ส": "Sat", "อา": "Sun",
    "จ.": "Mon", "อ.": "Tue", "พ.": "Wed", "พฤ.": "Thu", "ศ.": "Fri", "ส.": "Sat", "อา.": "Sun",
})

SPECIAL_RULES = {
    9: "รายวิชาทฤษฎีต้องใช้ห้องทฤษฎีเท่านั้น วิชาที่มีชั่วโมงปฏิบัติใช้ห้องทฤษฎีไม่ได้",
    10: "หัวหน้าแผนก/หัวหน้างาน (Leader) ประชุมทุกวันอังคาร 15.00-16.00 น. (คาบ 8)",
    11: "ทุกกลุ่มเรียนต้องมีคาบโฮมรูม (พบครูที่ปรึกษา) อย่างน้อย 1 คาบ/สัปดาห์",
    12: "รายวิชากิจกรรมต้องจัดวันพุธ 15.00-17.00 น. (คาบ 8-9)",
    13: "รายวิชาสามัญ (รหัสขึ้นต้น 20000/30000) ต้องเรียนร่วมกัน 2 กลุ่มในวิชาเดียวกัน",
}

# ---------------------------------------------------------------------------
# ประเภทไฟล์ข้อมูล
# ---------------------------------------------------------------------------
KINDS = {
    "teacher": {"label": "ครูผู้สอน", "file": "teacher.csv", "icon": "person",
                "cols": ["teacher_id", "teacher_name", "role"]},
    "room": {"label": "ห้องเรียน", "file": "room.csv", "icon": "meeting_room",
             "cols": ["room_id", "room_name", "room_type"]},
    "student_group": {"label": "กลุ่มเรียน", "file": "student_group.csv", "icon": "groups",
                      "cols": ["group_id", "group_name", "student_count", "advisor"]},
    "subject": {"label": "รายวิชา", "file": "subject.csv", "icon": "menu_book",
                "cols": ["subject_id", "subject_name", "theory", "practice", "credit"]},
    "teach": {"label": "วิชาที่ครูรับผิดชอบ", "file": "teach.csv", "icon": "co_present",
              "cols": ["teacher_id", "subject_id"]},
    "register": {"label": "การลงทะเบียนเรียน", "file": "register.csv", "icon": "edit_note",
                 "cols": ["group_id", "subject_id"]},
    "timeslot": {"label": "คาบเรียน", "file": "timeslot.csv", "icon": "schedule",
                 "cols": ["timeslot_id", "day", "period", "start", "end"]},
    "student": {"label": "รายชื่อนักเรียน (ไม่บังคับ)", "file": "student.csv", "icon": "school",
                "cols": ["student_id", "student_name", "group_id"]},
}
CORE_KINDS = [k for k in KINDS if k != "student"]

# ชื่อไฟล์/ชื่อชีต -> ประเภท (ตรวจตามลำดับ)
NAME_PATTERNS = [
    ("output", r"output|result"),
    ("student_group", r"student_?groups?|groups?|class_?groups?"),
    ("timeslot", r"time_?slots?|periods?"),
    ("register", r"regist\w*|enroll\w*"),
    ("teacher", r"teachers?"),
    ("teach", r"teach(es|ing)?"),
    ("subject", r"subjects?|courses?"),
    ("room", r"rooms?"),
    ("student", r"students?"),
]

# หัวคอลัมน์ภาษาไทย/ชื่ออื่นที่พบบ่อย -> ชื่อมาตรฐาน
COL_ALIASES = {
    "รหัสครู": "teacher_id", "ชื่อครู": "teacher_name", "ชื่อ_นามสกุล": "teacher_name", "ตำแหน่ง": "role",
    "รหัสห้อง": "room_id", "ชื่อห้อง": "room_name", "ประเภทห้อง": "room_type", "ชนิดห้อง": "room_type",
    "รหัสกลุ่ม": "group_id", "รหัสกลุ่มเรียน": "group_id", "ชื่อกลุ่ม": "group_name",
    "ชื่อกลุ่มเรียน": "group_name", "จำนวนนักเรียน": "student_count", "ครูที่ปรึกษา": "advisor",
    "รหัสวิชา": "subject_id", "ชื่อวิชา": "subject_name", "ทฤษฎี": "theory", "ปฏิบัติ": "practice",
    "หน่วยกิต": "credit", "วัน": "day", "คาบ": "period", "คาบที่": "period", "เวลาเริ่ม": "start",
    "เวลาสิ้นสุด": "end", "รหัสคาบ": "timeslot_id", "รหัสคาบเรียน": "timeslot_id",
    "รหัสนักเรียน": "student_id", "รหัสนักศึกษา": "student_id", "ชื่อนักเรียน": "student_name",
    "start_time": "start", "end_time": "end", "time_slot_id": "timeslot_id", "slot_id": "timeslot_id",
    "groupid": "group_id", "teacherid": "teacher_id", "subjectid": "subject_id", "roomid": "room_id",
    "timeslotid": "timeslot_id", "studentid": "student_id",
}


def now_text():
    return datetime.now(TZ).strftime("%d/%m/%Y %H:%M")


def df_to_csv_bytes(df):
    """CSV แบบ utf-8-sig เพื่อให้เปิดภาษาไทยใน Excel ได้ถูกต้อง"""
    return df.to_csv(index=False).encode("utf-8-sig")


# ---------------------------------------------------------------------------
# 1) อ่านไฟล์ที่อัปโหลด (CSV / Excel หลายชีต / ZIP)
# ---------------------------------------------------------------------------
def norm_col(c):
    s = str(c).replace("﻿", "").strip().lower()
    s = re.sub(r"[\s\-./]+", "_", s).strip("_")
    return COL_ALIASES.get(s, s)


def clean_df(df):
    df = df.copy()
    df.columns = [norm_col(c) for c in df.columns]
    df = df.loc[:, [c for c in df.columns if c and not c.startswith("unnamed")]]
    df = df.loc[:, ~pd.Index(df.columns).duplicated()]
    for col in df.columns:
        df[col] = df[col].map(lambda v: v.strip() if isinstance(v, str) else v)
        df[col] = df[col].replace({"": pd.NA, "nan": pd.NA, "NaN": pd.NA, "None": pd.NA})
    return df.dropna(how="all").reset_index(drop=True)


def read_csv_bytes(raw):
    for enc in ("utf-8-sig", "utf-8", "cp874", "tis-620"):
        try:
            text = raw.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    else:
        raise ValueError("อ่านไฟล์ไม่ได้ (ไม่ใช่ UTF-8 หรือ TIS-620)")
    first = text.splitlines()[0] if text.strip() else ""
    sep = ";" if first.count(";") > first.count(",") else ("\t" if first.count("\t") > first.count(",") else ",")
    return pd.read_csv(io.StringIO(text), dtype=str, sep=sep, skipinitialspace=True)


def read_file_bytes(name, raw):
    """คืนค่า [(ชื่อที่แสดง, [ชื่อที่ใช้เดาประเภท], DataFrame)]"""
    lower = name.lower()
    stem = Path(name).stem
    if lower.endswith(".zip"):
        out = []
        with zipfile.ZipFile(io.BytesIO(raw)) as zf:
            for member in zf.namelist():
                if member.endswith("/") or "__MACOSX" in member or Path(member).name.startswith((".", "~$")):
                    continue
                if Path(member).suffix.lower() in (".csv", ".txt", ".xlsx", ".xlsm", ".xls"):
                    out += read_file_bytes(member, zf.read(member))
        return out
    if lower.endswith((".xlsx", ".xlsm", ".xls")):
        sheets = pd.read_excel(io.BytesIO(raw), sheet_name=None, dtype=str)
        many = len(sheets) > 1
        return [(f"{Path(name).name} › {sh}" if many else Path(name).name, [str(sh), stem] if many else [stem, str(sh)], df)
                for sh, df in sheets.items() if not df.dropna(how="all").empty]
    return [(Path(name).name, [stem], read_csv_bytes(raw))]


def name_kind(name):
    stem = re.sub(r"[^a-z0-9]+", "_", str(name).lower())
    for kind, pat in NAME_PATTERNS:
        if re.search(rf"(^|_)({pat})(_|$)", stem):
            return kind
    return None


def columns_fit(kind, cols, strict=True):
    c = set(cols)
    has = lambda *xs: all(x in c for x in xs)
    anyof = lambda *xs: any(x in c for x in xs)
    if kind == "output":
        return has("group_id", "timeslot_id")
    if kind == "teacher":
        return has("teacher_id") and "subject_id" not in c and (anyof("teacher_name", "role") or not strict)
    if kind == "teach":
        return has("teacher_id", "subject_id") and "group_id" not in c
    if kind == "room":
        return has("room_id") and not anyof("group_id", "timeslot_id") and (anyof("room_name", "room_type") or not strict)
    if kind == "student_group":
        return has("group_id") and "subject_id" not in c and \
            (anyof("group_name", "student_count", "advisor") or not strict)
    if kind == "subject":
        return has("subject_id") and not anyof("group_id", "teacher_id") and \
            (anyof("subject_name", "theory", "practice", "credit") or not strict)
    if kind == "register":
        return has("group_id", "subject_id") and "timeslot_id" not in c
    if kind == "timeslot":
        return (has("timeslot_id") or has("day", "period")) and not anyof("group_id", "subject_id")
    if kind == "student":
        return has("student_id")
    return False


def detect_kind(names, df):
    cols = list(df.columns)
    for nm in names:   # 1) ดูจากชื่อชีต/ชื่อไฟล์ก่อน แต่คอลัมน์ต้องสอดคล้อง
        kind = name_kind(nm)
        if kind and columns_fit(kind, cols, strict=False):
            return kind
    for kind in ("output", "timeslot", "teach", "register", "subject", "teacher", "room",
                 "student_group", "student"):   # 2) ดูจากคอลัมน์อย่างเดียว
        if columns_fit(kind, cols):
            return kind
    return None


def ingest_files(files):
    """files = [(ชื่อไฟล์, bytes)] -> (raw: {kind: [(ที่มา, df)]}, report: DataFrame)"""
    raw, report = defaultdict(list), []
    for fname, data in files:
        try:
            parts = read_file_bytes(fname, data)
        except Exception as exc:
            report.append({"ไฟล์": fname, "ประเภทที่ตรวจพบ": "อ่านไม่ได้", "จำนวนแถว": 0, "หมายเหตุ": str(exc)})
            continue
        if not parts:
            report.append({"ไฟล์": fname, "ประเภทที่ตรวจพบ": "ไม่มีข้อมูล", "จำนวนแถว": 0, "หมายเหตุ": ""})
        for label, names, df in parts:
            df = clean_df(df)
            kind = detect_kind(names, df)
            if kind == "output":
                note, shown = "เป็นไฟล์ผลลัพธ์ตารางเรียน — ไม่นำมาใช้เป็นข้อมูลตั้งต้น", "ข้าม (output)"
            elif kind is None:
                note, shown = "คอลัมน์ไม่ตรงกับไฟล์ข้อมูลประเภทใด: " + ", ".join(df.columns[:8]), "ไม่ทราบประเภท"
            else:
                raw[kind].append((label, df))
                shown, note = f"{kind}.csv · {KINDS[kind]['label']}", ""
                missing = [c for c in KINDS[kind]["cols"] if c not in df.columns]
                if missing:
                    note = "ไม่มีคอลัมน์: " + ", ".join(missing) + " (ระบบเติมให้)"
            report.append({"ไฟล์": label, "ประเภทที่ตรวจพบ": shown, "จำนวนแถว": len(df), "หมายเหตุ": note})
    return dict(raw), pd.DataFrame(report, columns=["ไฟล์", "ประเภทที่ตรวจพบ", "จำนวนแถว", "หมายเหตุ"])


# ---------------------------------------------------------------------------
# 2) รวมข้อมูล + เติมข้อมูลที่ขาด (ข้อมูลไม่ครบก็ยังจัดตารางได้มากที่สุด)
# ---------------------------------------------------------------------------
def to_int(v):
    n = pd.to_numeric(v, errors="coerce")
    return None if pd.isna(n) else int(round(float(n)))


def to_str(v):
    if v is None or (not isinstance(v, str) and pd.isna(v)):
        return None
    s = str(v).strip()
    if re.fullmatch(r"-?\d+\.0", s):   # Excel อ่าน 5306 เป็น 5306.0
        s = s[:-2]
    return s or None


def parse_day(v):
    s = to_str(v)
    if not s:
        return None
    return DAY_ALIASES.get(s.lower()) or DAY_ALIASES.get(s.lower()[:3])


def parse_time(v):
    s = to_str(v)
    if not s:
        return None
    m = re.search(r"(\d{1,2})[:.](\d{2})", s)
    return f"{int(m.group(1)):02d}:{m.group(2)}" if m else None


def default_time(period):
    h = DEFAULT_START_HOUR + period - 1
    return f"{h:02d}:00", f"{h + 1:02d}:00"


def classify_room(rtype, rname):
    """แบ่งห้องเป็น theory / practice / general (ไม่ทราบ ใช้ได้ทั้งสองแบบ)"""
    practice_kw = ("practice", "practic", "lab", "ปฏิบัติ", "computer", "คอม", "workshop", "shop", "studio", "โรงฝึก")
    theory_kw = ("theory", "lecture", "ทฤษฎี", "บรรยาย", "lec")
    for text in (rtype, rname):
        t = (to_str(text) or "").lower()
        if not t:
            continue
        if any(k in t for k in practice_kw):
            return "practice"
        if any(k in t for k in theory_kw):
            return "theory"
    return "general"


def is_general_subject(sid):
    return str(sid).startswith(GENERAL_PREFIXES)


def is_activity_subject(sid, name):
    sid, name = str(sid), str(name or "")
    if "กิจกรรม" in name or "activity" in name.lower():
        return True
    return bool(re.match(r"^[23]0000-2\d{3}$", sid))   # รหัสกลุ่มกิจกรรมเสริมหลักสูตร


def merge_frames(raw, kind):
    frames = [df for _, df in raw.get(kind, [])]
    if not frames:
        return None
    df = pd.concat(frames, ignore_index=True)
    for c in KINDS[kind]["cols"]:
        if c not in df.columns:
            df[c] = pd.NA
    return df.drop_duplicates().reset_index(drop=True)


def build_dataset(raw):
    """คืนค่า dataset (dict ของ DataFrame มาตรฐาน) + notes [(ระดับ, ข้อความ)]"""
    notes = []
    note = lambda level, text: notes.append((level, text))
    got = {k: merge_frames(raw, k) for k in KINDS}

    # ---- timeslot ----
    ts_rows = []
    if got["timeslot"] is not None:
        bad = 0
        for r in got["timeslot"].itertuples(index=False):
            day, period = parse_day(r.day), to_int(r.period)
            if day is None or period is None or period <= 0:
                bad += 1
                continue
            d_start, d_end = default_time(period)
            ts_rows.append({"timeslot_id": to_str(r.timeslot_id), "day": day, "period": period,
                            "start": parse_time(r.start) or d_start, "end": parse_time(r.end) or d_end})
        if bad:
            note("warn", f"timeslot: ข้าม {bad} แถวที่ไม่มีวัน/คาบที่อ่านได้")
    if not ts_rows:
        note("info", "ไม่พบไฟล์ timeslot ที่ใช้ได้ — ใช้คาบมาตรฐาน จันทร์-ศุกร์ วันละ 12 คาบ เริ่ม 08:00 คาบละ 1 ชั่วโมง")
        n = 0
        for d in DEFAULT_DAYS:
            for p in range(1, DEFAULT_N_PERIODS + 1):
                n += 1
                s, e = default_time(p)
                ts_rows.append({"timeslot_id": str(n), "day": d, "period": p, "start": s, "end": e})
    timeslots = pd.DataFrame(ts_rows).drop_duplicates(["day", "period"])
    timeslots = timeslots.assign(_d=timeslots["day"].map(DAY_ORDER)).sort_values(["_d", "period"]).drop(columns="_d")
    if timeslots["timeslot_id"].isna().any() or timeslots["timeslot_id"].duplicated().any():
        if got["timeslot"] is not None:
            note("info", "timeslot: บางแถวไม่มีรหัสคาบหรือรหัสซ้ำ — ระบบกำหนดรหัสคาบใหม่เรียงตามวันและคาบ")
        timeslots["timeslot_id"] = [str(i) for i in range(1, len(timeslots) + 1)]
    timeslots["is_break"] = timeslots["period"].isin(BREAK_PERIODS)
    timeslots = timeslots.reset_index(drop=True)

    # ---- subject ----
    subj = {}
    if got["subject"] is not None:
        for r in got["subject"].itertuples(index=False):
            sid = to_str(r.subject_id)
            if not sid or sid in subj:
                continue
            subj[sid] = {"subject_id": sid, "subject_name": to_str(r.subject_name) or sid,
                         "theory": to_int(r.theory), "practice": to_int(r.practice), "credit": to_int(r.credit)}
    else:
        note("warn", "ไม่พบไฟล์ subject — ไม่ทราบจำนวนคาบของแต่ละวิชา (วิชาที่ไม่มีข้อมูลคาบจะยังจัดไม่ได้)")

    # ---- teach / register ----
    def pairs(kind, a, b):
        df = got[kind]
        if df is None:
            return []
        out = []
        for r in df.itertuples(index=False):
            x, y = to_str(getattr(r, a)), to_str(getattr(r, b))
            if x and y and (x, y) not in out:
                out.append((x, y))
        return out

    teach = pairs("teach", "teacher_id", "subject_id")
    register = pairs("register", "group_id", "subject_id")
    if got["teach"] is None:
        note("warn", "ไม่พบไฟล์ teach — ไม่ทราบว่าครูคนใดสอนวิชาใด รายวิชาจะถูกจัดโดยไม่มีครูผู้สอน")
    if got["register"] is None:
        note("error", "ไม่พบไฟล์ register — ไม่ทราบว่ากลุ่มเรียนใดลงทะเบียนวิชาใด จึงยังจัดตารางไม่ได้")

    for sid in {s for _, s in register} | {s for _, s in teach}:
        if sid not in subj:
            subj[sid] = {"subject_id": sid, "subject_name": sid, "theory": None, "practice": None, "credit": None}
            if got["subject"] is not None:
                note("warn", f"วิชา {sid} ไม่มีในไฟล์ subject")
    rows = []
    for s in subj.values():
        th, pr = s["theory"], s["practice"]
        if th is None and pr is None:
            periods = s["credit"] if s["credit"] else None
            if periods:
                note("warn", f"วิชา {s['subject_id']} ไม่มีจำนวนคาบทฤษฎี/ปฏิบัติ — ประมาณจากหน่วยกิต = {periods} คาบ")
                th, pr = periods, 0
        else:
            th, pr = th or 0, pr or 0
            periods = th + pr
        rows.append({**s, "theory": th, "practice": pr, "periods": periods,
                     "is_general": is_general_subject(s["subject_id"]),
                     "is_activity": is_activity_subject(s["subject_id"], s["subject_name"])})
    subjects = pd.DataFrame(rows, columns=["subject_id", "subject_name", "theory", "practice", "credit",
                                           "periods", "is_general", "is_activity"])

    # ---- teacher ----
    teachers = {}
    if got["teacher"] is not None:
        for r in got["teacher"].itertuples(index=False):
            tid = to_str(r.teacher_id)
            if tid and tid not in teachers:
                role_raw = (to_str(r.role) or "").lower()
                role = "Leader" if ("leader" in role_raw or "หัวหน้า" in role_raw) else "Teacher"
                teachers[tid] = {"teacher_id": tid, "teacher_name": to_str(r.teacher_name) or tid, "role": role}
    else:
        note("info", "ไม่พบไฟล์ teacher — สร้างรายชื่อครูจากรหัสครูในไฟล์ teach (ใช้รหัสแทนชื่อ)")
    for tid, _ in teach:
        if tid not in teachers:
            teachers[tid] = {"teacher_id": tid, "teacher_name": tid, "role": "Teacher"}
            if got["teacher"] is not None:
                note("warn", f"รหัสครู {tid} ในไฟล์ teach ไม่มีในไฟล์ teacher")
    teachers_df = pd.DataFrame(list(teachers.values()), columns=["teacher_id", "teacher_name", "role"])
    name_to_tid = {str(n).replace(" ", ""): t for t, n in zip(teachers_df["teacher_id"], teachers_df["teacher_name"])}

    # ---- student_group ----
    groups = {}
    if got["student_group"] is not None:
        for r in got["student_group"].itertuples(index=False):
            gid = to_str(r.group_id)
            if gid and gid not in groups:
                adv = to_str(r.advisor)
                adv_id = adv if adv in teachers else name_to_tid.get((adv or "").replace(" ", ""))
                groups[gid] = {"group_id": gid, "group_name": to_str(r.group_name) or gid,
                               "student_count": to_int(r.student_count), "advisor": adv, "advisor_id": adv_id}
    else:
        note("info", "ไม่พบไฟล์ student_group — สร้างกลุ่มเรียนจากรหัสกลุ่มในไฟล์ register")
    for gid, _ in register:
        if gid not in groups:
            groups[gid] = {"group_id": gid, "group_name": gid, "student_count": None, "advisor": None,
                           "advisor_id": None}
    groups_df = pd.DataFrame(list(groups.values()),
                             columns=["group_id", "group_name", "student_count", "advisor", "advisor_id"])

    # ---- room ----
    rooms = {}
    if got["room"] is not None:
        for r in got["room"].itertuples(index=False):
            rid = to_str(r.room_id)
            if rid and rid not in rooms:
                rooms[rid] = {"room_id": rid, "room_name": to_str(r.room_name) or rid,
                              "room_type": to_str(r.room_type) or "-",
                              "category": classify_room(r.room_type, r.room_name)}
    else:
        note("warn", "ไม่พบไฟล์ room — จะจัดตารางโดยไม่ระบุห้องเรียน")
    rooms_df = pd.DataFrame(list(rooms.values()), columns=["room_id", "room_name", "room_type", "category"])
    if len(rooms_df) and (rooms_df["category"] == "general").all():
        note("info", "ระบุประเภทห้อง (ทฤษฎี/ปฏิบัติ) จาก room_type ไม่ได้ — ถือว่าทุกห้องใช้ได้ทั้งทฤษฎีและปฏิบัติ")

    # ---- student (ไม่บังคับ) ----
    students_df = pd.DataFrame(columns=["student_id", "student_name", "group_id"])
    if got["student"] is not None:
        st_rows = []
        for r in got["student"].itertuples(index=False):
            sid = to_str(r.student_id)
            gid = to_str(r.group_id)
            if sid and gid in groups:
                st_rows.append({"student_id": sid, "student_name": to_str(r.student_name) or sid, "group_id": gid})
        students_df = pd.DataFrame(st_rows, columns=students_df.columns).drop_duplicates("student_id")

    # ---- ตรวจความสมบูรณ์ของการลงทะเบียน ----
    periods_of = dict(zip(subjects["subject_id"], subjects["periods"]))
    no_hours = sorted({s for _, s in register if not periods_of.get(s)})
    if no_hours:
        note("warn", "วิชาที่ไม่ทราบจำนวนคาบ (ยังจัดไม่ได้): " + ", ".join(no_hours))
    teach_subjects = {s for _, s in teach}
    no_teacher = sorted({s for _, s in register if s not in teach_subjects})
    if teach and no_teacher:
        note("warn", "วิชาที่ลงทะเบียนแต่ไม่มีครูผู้สอนในไฟล์ teach (จะจัดโดยไม่มีครู): " + ", ".join(no_teacher))

    received = {k: (got[k] is not None) for k in KINDS}
    sources = {k: [lbl for lbl, _ in raw.get(k, [])] for k in KINDS}
    return {
        "teachers": teachers_df, "rooms": rooms_df, "groups": groups_df, "subjects": subjects,
        "teach": pd.DataFrame(teach, columns=["teacher_id", "subject_id"]),
        "register": pd.DataFrame(register, columns=["group_id", "subject_id"]),
        "timeslots": timeslots, "students": students_df,
        "received": received, "sources": sources, "notes": notes,
        "can_schedule": bool(register),
    }


# =============================================================================
# 3) อัลกอริทึมจัดตาราง: Constraint Satisfaction แบบ Greedy + Heuristic + Random Restart
#
#   1) แปลงการลงทะเบียน (register) เป็น "รายวิชา-กลุ่ม" (course) จำนวนคาบ = theory + practice
#      - ข้อ 13: วิชาสามัญจับคู่ 2 กลุ่มเรียนร่วมกัน (ครู ห้อง เวลาเดียวกัน)
#      - ข้อ 11: เพิ่มคาบโฮมรูมให้ทุกกลุ่ม โดยครูที่ปรึกษา
#   2) เลือกครูผู้สอนจาก teach.csv เท่านั้น (ข้อ 8) ถ้ามีหลายคนเลือกคนที่ภาระงานน้อยที่สุด
#   3) แตกคาบเป็น "บล็อก" คาบติดกัน แล้วจัดบล็อกที่ยากที่สุดก่อน (Most-Constrained-First)
#      บล็อกต้องผ่าน Hard Constraints ทั้งหมด แล้วเลือกช่องที่คะแนน Soft Constraints ดีที่สุด
#   4) วางไม่ได้ -> แบ่งบล็อกครึ่งแล้วลองใหม่ เหลือ 1 คาบยังวางไม่ได้ -> รายงานพร้อมสาเหตุ
#   5) ทำซ้ำหลายรอบด้วยลำดับสุ่มต่างกัน เก็บผลที่ดีที่สุด แล้วตรวจผลด้วย verify_schedule() อีกชั้น
# =============================================================================
DEFAULT_SETTINGS = {
    "rules": [],          # เงื่อนไขพิเศษที่เลือก (9-13)
    "max_block": 5,       # เรียนติดกันสูงสุด (คาบ/ครั้ง)
    "attempts": 30,       # จำนวนรอบค้นหา
    "seed": 42,
}

SCHEDULE_COLS = ["group_id", "timeslot_id", "day", "period", "subject_id", "teacher_id", "room_id",
                 "session", "kind"]
OUTPUT_COLS = ["group_id", "timeslot_id", "subject_id", "teacher_id", "room_id"]


def split_into_blocks(hours, max_block):
    n_blocks = -(-hours // max_block)
    base, extra = divmod(hours, n_blocks)
    return [base + 1] * extra + [base] * (n_blocks - extra)


def build_courses(ds, rules):
    subjects = ds["subjects"].set_index("subject_id")
    groups_order = {g: i for i, g in enumerate(ds["groups"]["group_id"])}
    teach_map = defaultdict(list)
    for t, s in ds["teach"].itertuples(index=False):
        teach_map[s].append(t)

    by_subject, skipped = defaultdict(list), []
    for g, s in ds["register"].itertuples(index=False):
        periods = subjects.at[s, "periods"] if s in subjects.index else None
        if periods is None or pd.isna(periods) or int(periods) <= 0:
            skipped.append({"กลุ่มเรียน": g, "รหัสวิชา": s, "สาเหตุ": "ไม่ทราบจำนวนคาบ (theory + practice)"})
            continue
        by_subject[s].append(g)

    courses = []
    for s, glist in by_subject.items():
        glist = sorted(set(glist), key=lambda g: groups_order.get(g, 9999))
        info = subjects.loc[s]
        kind = "activity" if info["is_activity"] else "normal"
        needs = "practice" if (info["practice"] or 0) > 0 else "theory"
        cands = list(teach_map.get(s, []))
        if 12 in rules and kind == "activity":
            # ข้อ 12: ทุกกลุ่มทำกิจกรรมพร้อมกันวันพุธคาบ 8-9 — ครู 1 คนดูแลหลายกลุ่มในคาบเดียว
            # จึงรวมเป็นกิจกรรมร่วม แบ่งกลุ่มตามจำนวนครูที่รับผิดชอบวิชานั้น
            k = max(1, len(cands))
            chunks = [(tuple(glist[i::k]), [cands[i]] if cands else []) for i in range(k) if glist[i::k]]
        elif 13 in rules and info["is_general"] and len(glist) >= 2:
            chunks = [(tuple(glist[i:i + 2]), cands) for i in range(0, len(glist), 2)]
        else:
            chunks = [((g,), cands) for g in glist]
        for chunk, chunk_cands in chunks:
            courses.append({"groups": chunk, "subject_id": s, "periods": int(info["periods"]),
                            "candidates": list(chunk_cands), "kind": kind, "needs": needs})

    if 11 in rules:
        for g in ds["groups"].itertuples(index=False):
            adv = g.advisor_id if isinstance(g.advisor_id, str) else None
            courses.append({"groups": (g.group_id,), "subject_id": HOMEROOM_ID, "periods": HOMEROOM_PERIODS,
                            "candidates": [adv] if adv else [], "kind": "homeroom", "needs": "theory"})
    return courses, skipped


def build_env(ds, settings):
    rules = set(settings["rules"])
    ts = ds["timeslots"]
    days = [d for d in DAY_CODES if d in set(ts["day"])]
    day_periods = {d: sorted(ts[(ts["day"] == d) & (~ts["is_break"])]["period"].tolist()) for d in days}
    ts_id = {(d, p): i for i, d, p in zip(ts["timeslot_id"], ts["day"], ts["period"])}

    def contiguous(block):
        return all(all(q in BREAK_PERIODS for q in range(a + 1, b)) for a, b in zip(block, block[1:]))

    leaders = set(ds["teachers"].loc[ds["teachers"]["role"] == "Leader", "teacher_id"])
    blocked = set()
    if 10 in rules:
        d, ps = LEADER_MEETING
        blocked = {(t, d, p) for t in leaders for p in ps}

    rooms = ds["rooms"]
    room_cat = dict(zip(rooms["room_id"], rooms["category"]))
    return {"rules": rules, "days": days, "day_periods": day_periods, "ts_id": ts_id,
            "contiguous": contiguous, "blocked": blocked, "room_cat": room_cat,
            "room_ids": list(rooms["room_id"]), "has_rooms": len(rooms) > 0, "leaders": leaders}


def eligible_rooms(course, env):
    if not env["has_rooms"]:
        return None, set()
    match = [r for r in env["room_ids"] if env["room_cat"][r] in (course["needs"], "general")]
    if 9 in env["rules"]:
        return match, set(match)
    return env["room_ids"], set(match)   # ไม่บังคับข้อ 9: ใช้ได้ทุกห้อง แต่ชอบห้องที่ตรงประเภท


def assign_teachers(courses, rng):
    load = defaultdict(int)
    chosen = {}
    order = sorted(range(len(courses)),
                   key=lambda i: (len(courses[i]["candidates"]) or 99, -courses[i]["periods"], rng.random()))
    for i in order:
        cands = courses[i]["candidates"]
        if not cands:
            chosen[i] = None
            continue
        t = min(cands, key=lambda c: (load[c], rng.random()))
        chosen[i] = t
        load[t] += courses[i]["periods"]
    return chosen


def make_tasks(courses, teachers, env, settings):
    tasks = []
    for i, c in enumerate(courses):
        rooms, pref = eligible_rooms(c, env)
        remaining, fixed = c["periods"], []
        if c["kind"] == "activity" and 12 in env["rules"]:
            day, ps = ACTIVITY_SLOT
            avail = [p for p in ps if p in env["day_periods"].get(day, [])]
            take = min(remaining, len(avail))
            if take:
                fixed.append((day, avail[:take]))
                remaining -= take
        base = {"course": i, "groups": c["groups"], "subject_id": c["subject_id"], "teacher": teachers[i],
                "kind": c["kind"], "rooms": rooms, "pref": pref}
        for day, block in fixed:
            tasks.append(dict(base, size=len(block), fixed=(day, block)))
        if remaining:
            for size in split_into_blocks(remaining, settings["max_block"]):
                tasks.append(dict(base, size=size, fixed=None))
    return tasks


def schedule_once(tasks, env, rng):
    group_busy, teacher_busy, room_busy = set(), set(), set()
    group_load, teacher_load, room_load = defaultdict(int), defaultdict(int), defaultdict(int)
    course_days = defaultdict(set)
    placements, unplaced = [], []
    penalty = 0
    days, day_periods, blocked = env["days"], env["day_periods"], env["blocked"]

    def difficulty(t):
        n_rooms = len(t["rooms"]) if t["rooms"] is not None else 50
        return ((1000 if t["fixed"] else 0) + t["size"] * 100 + len(t["groups"]) * 60
                + 80 / max(n_rooms, 1) + (20 if t["teacher"] in env["leaders"] else 0))

    queue = sorted(tasks, key=lambda t: (-difficulty(t), rng.random()))
    while queue:
        task = queue.pop(0)
        size, tid, groups = task["size"], task["teacher"], task["groups"]
        options = []
        if task["fixed"]:
            options = [(task["fixed"][0], task["fixed"][1])]
        else:
            for d in days:
                plist = day_periods[d]
                for i in range(len(plist) - size + 1):
                    block = plist[i:i + size]
                    if env["contiguous"](block):
                        options.append((d, block))

        best, best_score = None, None
        for d, block in options:
            if any(group_load[(g, d)] + size > MAX_PERIODS_PER_DAY for g in groups):
                continue
            if any((g, d, p) in group_busy for g in groups for p in block):
                continue
            if tid and any((tid, d, p) in teacher_busy or (tid, d, p) in blocked for p in block):
                continue
            base = ((30 if d in course_days[task["course"]] else 0)
                    + sum(group_load[(g, d)] for g in groups) * 3
                    + (teacher_load[(tid, d)] * 2 if tid else 0)
                    + block[0] * 0.4
                    + (1 if block[-1] - block[0] >= size else 0))
            if task["rooms"] is None:
                score = base + rng.random()
                if best_score is None or score < best_score:
                    best, best_score = (d, block, None), score
                continue
            for r in task["rooms"]:
                if any((r, d, p) in room_busy for p in block):
                    continue
                score = base + (0 if r in task["pref"] else 25) + room_load[r] * 0.05 + rng.random()
                if best_score is None or score < best_score:
                    best, best_score = (d, block, r), score

        if best is None:
            if size > 1 and not task["fixed"]:
                half = size // 2
                queue[0:0] = [dict(task, size=size - half), dict(task, size=half)]
                penalty += 5
            else:
                unplaced.append(task)
            continue

        d, block, room = best
        if d in course_days[task["course"]]:
            penalty += 10
        for p in block:
            for g in groups:
                group_busy.add((g, d, p))
            if tid:
                teacher_busy.add((tid, d, p))
            if room:
                room_busy.add((room, d, p))
        for g in groups:
            group_load[(g, d)] += size
        if tid:
            teacher_load[(tid, d)] += size
        if room:
            room_load[room] += size
        course_days[task["course"]].add(d)
        placements.append(dict(task, day=d, block=block, room=room))
    return placements, unplaced, penalty, (group_busy, teacher_busy)


def unplaced_reason(task, env, busy):
    group_busy, teacher_busy = busy
    if task["rooms"] is not None and not task["rooms"]:
        return "ไม่มีห้องประเภทที่ใช้ได้ (ข้อ 9)" if 9 in env["rules"] else "ไม่มีห้องเรียน"
    if task["fixed"]:
        return f"ช่วงวัน{DAY_TH[task['fixed'][0]]} คาบ {'-'.join(map(str, task['fixed'][1]))} ไม่ว่าง (ข้อ 12)"
    free_together = any(
        not any((g, d, p) in group_busy for g in task["groups"])
        and not (task["teacher"] and ((task["teacher"], d, p) in teacher_busy or (task["teacher"], d, p) in env["blocked"]))
        for d in env["days"] for p in env["day_periods"][d]
    )
    if not free_together:
        return "ครูและกลุ่มเรียนไม่มีคาบว่างตรงกัน"
    return "ห้องเรียนที่ใช้ได้ไม่ว่างในช่วงที่ครูและกลุ่มเรียนว่าง"


def generate_timetable(ds, settings, progress_cb=None):
    env = build_env(ds, settings)
    courses, skipped = build_courses(ds, env["rules"])
    rng = random.Random(settings["seed"])
    required = sum(c["periods"] * len(c["groups"]) for c in courses)

    best, best_key, best_teachers = None, None, {}
    attempts = settings["attempts"]
    for i in range(attempts):
        teachers = assign_teachers(courses, rng)
        tasks = make_tasks(courses, teachers, env, settings)
        placements, unplaced, penalty, busy = schedule_once(tasks, env, rng)
        key = (sum(t["size"] * len(t["groups"]) for t in unplaced), penalty)
        if best_key is None or key < best_key:
            best, best_key, best_teachers = (placements, unplaced, penalty, busy), key, teachers
        if progress_cb:
            progress_cb(i + 1, attempts, best_key)
        if best_key[0] == 0 and (best_key[1] == 0 or i + 1 >= min(8, attempts)):
            break
    placements, unplaced, penalty, busy = best

    # ---- ตรวจความจุ: ภาระที่ต้องจัดเกินคาบที่มีจริงหรือไม่ (ช่วยอธิบายว่าทำไมจัดไม่ครบ) ----
    slots = sum(len(v) for v in env["day_periods"].values())
    capacity = []
    t_need = defaultdict(int)
    for i, c in enumerate(courses):
        if best_teachers.get(i):
            t_need[best_teachers[i]] += c["periods"]
    for t, n in sorted(t_need.items(), key=lambda x: -x[1]):
        avail = slots - sum(1 for (tt, _, _) in env["blocked"] if tt == t)
        if n > avail:
            capacity.append(f"ครู {t} ต้องสอน {n} คาบ แต่มีคาบให้สอนเพียง {avail} คาบ/สัปดาห์ — ควรเพิ่มครูผู้สอนใน teach")
    g_need = defaultdict(int)
    for c in courses:
        for g in c["groups"]:
            g_need[g] += c["periods"]
    g_cap = sum(min(len(v), MAX_PERIODS_PER_DAY) for v in env["day_periods"].values())
    for g, n in g_need.items():
        if n > g_cap:
            capacity.append(f"กลุ่ม {g} ต้องเรียน {n} คาบ แต่เรียนได้สูงสุด {g_cap} คาบ/สัปดาห์")

    rows = []
    for n, pl in enumerate(placements, 1):
        for p in pl["block"]:
            for g in pl["groups"]:
                rows.append({"group_id": g, "timeslot_id": env["ts_id"][(pl["day"], p)], "day": pl["day"],
                             "period": p, "subject_id": pl["subject_id"], "teacher_id": pl["teacher"] or "",
                             "room_id": pl["room"] or "", "session": n, "kind": pl["kind"]})
    schedule = pd.DataFrame(rows, columns=SCHEDULE_COLS)
    if not schedule.empty:
        g_order = {g: i for i, g in enumerate(ds["groups"]["group_id"])}
        schedule = (schedule.assign(_g=schedule["group_id"].map(g_order), _d=schedule["day"].map(DAY_ORDER))
                    .sort_values(["_g", "_d", "period"]).drop(columns=["_g", "_d"]).reset_index(drop=True))

    subj_name = dict(zip(ds["subjects"]["subject_id"], ds["subjects"]["subject_name"]))
    subj_name[HOMEROOM_ID] = "โฮมรูม (พบครูที่ปรึกษา)"
    agg = defaultdict(lambda: {"periods": 0, "reason": None})
    for t in unplaced:
        key = ("+".join(t["groups"]), t["subject_id"], t["teacher"] or "-")
        agg[key]["periods"] += t["size"]
        agg[key]["reason"] = agg[key]["reason"] or unplaced_reason(t, env, busy)
    unplaced_df = pd.DataFrame(
        [{"กลุ่มเรียน": k[0], "รหัสวิชา": k[1], "ชื่อวิชา": subj_name.get(k[1], k[1]), "ครูผู้สอน": k[2],
          "คาบที่ยังจัดไม่ได้": v["periods"], "สาเหตุ": v["reason"]} for k, v in agg.items()],
        columns=["กลุ่มเรียน", "รหัสวิชา", "ชื่อวิชา", "ครูผู้สอน", "คาบที่ยังจัดไม่ได้", "สาเหตุ"])
    return {
        "schedule": schedule, "unplaced": unplaced_df,
        "skipped": pd.DataFrame(skipped, columns=["กลุ่มเรียน", "รหัสวิชา", "สาเหตุ"]),
        "capacity": capacity,
        "stats": {"required": required, "placed": len(schedule), "unplaced": required - len(schedule),
                  "courses": len(courses), "penalty": penalty},
        "settings": dict(settings),
    }


def output_df(schedule):
    """output.csv ตามรูปแบบที่กำหนด: group_id, timeslot_id, subject_id, teacher_id, room_id"""
    return schedule[OUTPUT_COLS].copy() if not schedule.empty else pd.DataFrame(columns=OUTPUT_COLS)


def output_detail_df(schedule):
    cols = ["group_id", "timeslot_id", "day", "period", "subject_id", "teacher_id", "room_id"]
    return schedule[cols].copy() if not schedule.empty else pd.DataFrame(columns=cols)


# ---------------------------------------------------------------------------
# 4) ตรวจผลลัพธ์ตามเงื่อนไขทุกข้อ (ตรวจแยกจากอัลกอริทึม เพื่อยืนยันผลอีกชั้น)
# ---------------------------------------------------------------------------
def verify_schedule(result, ds):
    sch = result["schedule"]
    rules = set(result["settings"]["rules"])
    rows = []

    def add(no, text, bad, detail_ok, detail_bad, examples=()):
        ex = list(examples)[:6]
        rows.append({"ข้อ": no, "เงื่อนไข": text, "ผล": "ผ่าน" if not bad else "ไม่ผ่าน",
                     "รายละเอียด": detail_ok if not bad else detail_bad + (": " + ", ".join(map(str, ex)) if ex else "")})

    subj = ds["subjects"].set_index("subject_id")
    reg = [(g, s) for g, s in ds["register"].itertuples(index=False)]
    need = {(g, s): int(subj.at[s, "periods"]) for g, s in reg
            if s in subj.index and pd.notna(subj.at[s, "periods"]) and subj.at[s, "periods"]}
    normal = sch[sch["kind"] != "homeroom"]
    got = normal.groupby(["group_id", "subject_id"]).size().to_dict()

    wrong = [f"{g}/{s} ({got.get((g, s), 0)}/{n})" for (g, s), n in need.items() if got.get((g, s), 0) != n]
    add(1, "จำนวนคาบของแต่ละวิชา = theory + practice", wrong,
        f"ครบทุกวิชา ({len(need)} รายการ)", f"ไม่ตรง {len(wrong)} รายการ", wrong)

    brk = sch[sch["period"].isin(BREAK_PERIODS)]
    add(2, f"ไม่จัดวิชาใดลงคาบพัก (คาบ {', '.join(map(str, sorted(BREAK_PERIODS)))})", len(brk),
        "ไม่มีวิชาในคาบพัก", f"พบ {len(brk)} คาบ")

    reg_count = defaultdict(set)
    for (g, s) in reg:   # นับทุกวิชาที่ลงทะเบียน รวมวิชาที่ข้อมูลไม่พอ (จะได้เห็นว่าขาด)
        reg_count[g].add(s)
    got_count = normal.groupby("group_id")["subject_id"].nunique().to_dict()
    bad3 = [f"{g} ({got_count.get(g, 0)}/{len(v)})" for g, v in reg_count.items() if got_count.get(g, 0) != len(v)]
    add(3, "จำนวนวิชาในตาราง = จำนวนวิชาที่ลงทะเบียน", bad3, "ครบทุกกลุ่มเรียน", "ไม่ครบ", bad3)

    tot_need = defaultdict(int)
    for (g, _), n in need.items():
        tot_need[g] += n
    tot_got = normal.groupby("group_id").size().to_dict()
    bad4 = [f"{g} ({tot_got.get(g, 0)}/{n})" for g, n in tot_need.items() if tot_got.get(g, 0) != n]
    add(4, "จำนวนคาบรวม = ผลรวมคาบของวิชาที่ลงทะเบียน", bad4, "ครบทุกกลุ่มเรียน", "ไม่เท่ากัน", bad4)

    per_day = sch.groupby(["group_id", "day"]).size()
    over = [f"{g} วัน{DAY_TH[d]} ({n})" for (g, d), n in per_day.items() if n > MAX_PERIODS_PER_DAY]
    add(5, f"แต่ละวันเรียนไม่เกิน {MAX_PERIODS_PER_DAY} คาบ", over,
        f"มากสุด {int(per_day.max()) if len(per_day) else 0} คาบ/วัน", "เกิน", over)

    t = sch[sch["teacher_id"] != ""]
    tc = t.groupby(["teacher_id", "day", "period"])["session"].nunique()
    tc = tc[tc > 1]
    add(6, "ครูไม่สอนชนกัน (1 คาบ สอนได้ 1 วิชา)", len(tc), "ไม่มีครูสอนชน",
        f"ชน {len(tc)} คาบ", [f"{a} วัน{DAY_TH[b]} คาบ{c}" for a, b, c in tc.index])

    r = sch[sch["room_id"] != ""]
    rc = r.groupby(["room_id", "day", "period"])["session"].nunique()
    rc = rc[rc > 1]
    add(7, "ห้องเรียน 1 คาบใช้ได้ 1 วิชา (ห้องไม่ทับกัน)", len(rc),
        "ไม่มีห้องทับกัน" + (" (วิชาสามัญเรียนร่วม 2 กลุ่มตามข้อ 13)" if 13 in rules else ""),
        f"ทับกัน {len(rc)} คาบ", [f"{a} วัน{DAY_TH[b]} คาบ{c}" for a, b, c in rc.index])

    allowed = set(map(tuple, ds["teach"][["teacher_id", "subject_id"]].values.tolist()))
    tt = normal[normal["teacher_id"] != ""].drop_duplicates(["teacher_id", "subject_id"])
    bad8 = [f"{a}-{b}" for a, b in zip(tt["teacher_id"], tt["subject_id"]) if (a, b) not in allowed]
    no_t = normal[normal["teacher_id"] == ""]["subject_id"].unique().tolist()
    add(8, "ครูสอนเฉพาะวิชาที่กำหนดในไฟล์ teach", bad8,
        "ถูกต้องทุกวิชา" + (f" (ไม่มีครูใน teach: {', '.join(no_t)})" if no_t else ""), "ไม่ตรงกับ teach", bad8)

    if 9 in rules:
        cat = dict(zip(ds["rooms"]["room_id"], ds["rooms"]["category"]))
        needs = {s: ("practice" if (subj.at[s, "practice"] or 0) > 0 else "theory") for s in subj.index}
        needs[HOMEROOM_ID] = "theory"
        rr = sch[sch["room_id"] != ""].drop_duplicates(["subject_id", "room_id", "group_id"])
        bad9 = [f"{s}@{rm}" for s, rm in zip(rr["subject_id"], rr["room_id"])
                if cat.get(rm, "general") not in (needs.get(s, "theory"), "general")]
        add(9, SPECIAL_RULES[9], bad9, "ทุกวิชาใช้ห้องตรงประเภท", "ใช้ห้องผิดประเภท", bad9)
    if 10 in rules:
        leaders = set(ds["teachers"].loc[ds["teachers"]["role"] == "Leader", "teacher_id"])
        d, ps = LEADER_MEETING
        hit = sch[sch["teacher_id"].isin(leaders) & (sch["day"] == d) & sch["period"].isin(ps)]
        add(10, SPECIAL_RULES[10], len(hit), f"หัวหน้า {len(leaders)} คนว่างตรงเวลาประชุม",
            "มีคาบสอนชนเวลาประชุม", hit["teacher_id"].unique())
    if 11 in rules:
        hr = sch[sch["kind"] == "homeroom"].groupby("group_id").size()
        miss = [g for g in ds["groups"]["group_id"] if hr.get(g, 0) < HOMEROOM_PERIODS]
        add(11, SPECIAL_RULES[11], miss, f"ทุกกลุ่มมีโฮมรูม ({len(hr)} กลุ่ม)", "ยังไม่มีโฮมรูม", miss)
    if 12 in rules:
        d, ps = ACTIVITY_SLOT
        act = sch[sch["kind"] == "activity"]
        out = act[(act["day"] != d) | (~act["period"].isin(ps))]
        acts = ds["subjects"].loc[ds["subjects"]["is_activity"], "subject_id"].tolist()
        add(12, SPECIAL_RULES[12], len(out),
            f"วิชากิจกรรม {', '.join(acts) or '-'} อยู่วันพุธคาบ 8-9 ทั้งหมด", "อยู่นอกช่วงที่กำหนด",
            (out["group_id"] + " " + out["subject_id"]).unique())
    if 13 in rules:
        gen = sch[sch["subject_id"].map(is_general_subject)]
        sess = gen.groupby(["subject_id", "session"])["group_id"].agg(lambda x: tuple(sorted(set(x))))
        reg_groups = ds["register"].groupby("subject_id")["group_id"].nunique()
        solo = defaultdict(set)
        for (s, _), gs in sess.items():
            if len(gs) == 1:
                solo[s] |= set(gs)
        # กลุ่มที่เรียนเดี่ยวได้ = เศษเมื่อจับคู่ (จำนวนกลุ่มที่ลงทะเบียนเป็นเลขคี่ -> เหลือ 1 กลุ่ม)
        bad13 = [f"{s} ({', '.join(sorted(gs))})" for s, gs in solo.items()
                 if len(gs) > (reg_groups.get(s, 0) % 2 if reg_groups.get(s, 0) >= 2 else 1)]
        add(13, SPECIAL_RULES[13], bad13, "วิชาสามัญเรียนร่วมกันอย่างน้อย 2 กลุ่ม (ยกเว้นกลุ่มที่เหลือเศษ)",
            "มีกลุ่มที่เรียนเดี่ยว", bad13)
    return pd.DataFrame(rows, columns=["ข้อ", "เงื่อนไข", "ผล", "รายละเอียด"])


# =============================================================================
# 5) การแสดงผลตาราง + ส่งออก Excel / PDF / CSV
# =============================================================================
VIEWS = {
    "group": {"col": "group_id", "noun": "กลุ่มเรียน", "title": "ตารางเรียน", "icon": "groups"},
    "teacher": {"col": "teacher_id", "noun": "ครูผู้สอน", "title": "ตารางสอน", "icon": "co_present"},
    "room": {"col": "room_id", "noun": "ห้องเรียน", "title": "ตารางการใช้ห้อง", "icon": "meeting_room"},
}
SUBJECT_COLORS = ["#EDE4FF", "#FFE1EE", "#DDF3EA", "#E0ECFF", "#FFEBD9",
                  "#F6E3FA", "#E6F6FB", "#FFF4D6", "#E8E4FF", "#FBE4E9"]


def lookups(ds):
    subj_name = dict(zip(ds["subjects"]["subject_id"], ds["subjects"]["subject_name"]))
    subj_name[HOMEROOM_ID] = "โฮมรูม (พบครูที่ปรึกษา)"
    return {
        "subject": subj_name,
        "teacher": dict(zip(ds["teachers"]["teacher_id"], ds["teachers"]["teacher_name"])),
        "group": dict(zip(ds["groups"]["group_id"], ds["groups"]["group_name"])),
        "room": dict(zip(ds["rooms"]["room_id"], ds["rooms"]["room_name"])),
    }


def entity_options(ds, schedule, view):
    lk = lookups(ds)
    if view == "group":
        ids = list(ds["groups"]["group_id"])
        return {g: (f"{lk['group'][g]} ({g})" if lk["group"].get(g, g) != g else g) for g in ids}
    if view == "teacher":
        ids = list(ds["teachers"]["teacher_id"])
        return {t: f"{lk['teacher'].get(t, t)} ({t})" for t in ids}
    ids = list(ds["rooms"]["room_id"])
    extra = [r for r in schedule["room_id"].unique() if r and r not in ids] if not schedule.empty else []
    return {r: (f"{r} · {lk['room'][r]}" if lk["room"].get(r, r) != r else r) for r in ids + extra}


def entity_title(ds, view, eid):
    lk = lookups(ds)
    name = {"group": lk["group"], "teacher": lk["teacher"], "room": lk["room"]}[view].get(eid, eid)
    return f"{VIEWS[view]['title']}: {name}" + (f" ({eid})" if name != eid else "")


def timeslot_layout(ds):
    ts = ds["timeslots"]
    days = [d for d in DAY_CODES if d in set(ts["day"])]
    periods = set(ts["period"])
    lo, hi = min(periods), max(periods)
    periods = sorted(periods | {p for p in BREAK_PERIODS if lo <= p <= hi})
    times = {}
    for p in periods:
        hit = ts[ts["period"] == p]
        if len(hit):
            times[p] = f"{hit.iloc[0]['start']}-{hit.iloc[0]['end']}"
        else:
            s, e = default_time(p)
            times[p] = f"{s}-{e}"
    return days, periods, times


def grid_data(ds, schedule, view, eid):
    days, periods, times = timeslot_layout(ds)
    lk = lookups(ds)
    sub = schedule[schedule[VIEWS[view]["col"]] == eid] if not schedule.empty else schedule
    cells = {}
    for (d, p), rows in sub.groupby(["day", "period"]):
        r = rows.iloc[0]
        groups = "+".join(sorted(rows["group_id"].unique()))
        room = f"[{r.room_id}]" if r.room_id else ""
        if view == "group":
            lines = [r.subject_id, r.teacher_id or "-", room]
        elif view == "teacher":
            lines = [r.subject_id, groups, room]
        else:
            lines = [r.subject_id, groups, r.teacher_id or "-"]
        cells[(d, p)] = {"lines": lines, "subject_id": r.subject_id, "kind": r.kind,
                         "title": lk["subject"].get(r.subject_id, r.subject_id)}
    return {"days": days, "periods": periods, "times": times, "cells": cells}


def legend_df(ds, schedule, view, eid):
    lk = lookups(ds)
    subj = ds["subjects"].set_index("subject_id")
    sub = schedule[schedule[VIEWS[view]["col"]] == eid] if not schedule.empty else schedule
    rows = []
    for sid, part in sub.groupby("subject_id", sort=False):
        if sid in subj.index:
            s = subj.loc[sid]
            tpc = f"{s['theory']}-{s['practice']}-{s['credit'] if pd.notna(s['credit']) else '-'}"
        else:
            tpc = "-"
        row = {"รหัสวิชา": sid, "ชื่อวิชา": lk["subject"].get(sid, sid), "ท-ป-น": tpc,
               "คาบ/สัปดาห์": part[["day", "period"]].drop_duplicates().shape[0]}
        if view != "teacher":
            row["ครูผู้สอน"] = ", ".join(f"{lk['teacher'].get(t, t)} ({t})" for t in part["teacher_id"].unique() if t) or "-"
        if view != "group":
            row["กลุ่มเรียน"] = ", ".join(sorted(part["group_id"].unique()))
        if view != "room":
            row["ห้องเรียน"] = ", ".join(r for r in part["room_id"].unique() if r) or "-"
        rows.append(row)
    return pd.DataFrame(rows)


def color_map(schedule):
    ids = sorted(schedule["subject_id"].unique()) if not schedule.empty else []
    return {sid: SUBJECT_COLORS[i % len(SUBJECT_COLORS)] for i, sid in enumerate(ids)}


GRID_CSS = """
.tt-wrap{overflow-x:auto;margin:.3rem 0 1.1rem;border-radius:18px;border:1px solid #e6dcf8;background:#fff;
  box-shadow:0 10px 28px -18px rgba(111,82,192,.45)}
table.tt{border-collapse:collapse;width:100%;min-width:900px;table-layout:fixed;font-size:12px;
  background:#fff;color:#3a2c5a;font-family:'Prompt','Sarabun','Tahoma',sans-serif}
table.tt th,table.tt td{border:1px solid #e9e1f8;padding:4px 2px;text-align:center;vertical-align:middle}
table.tt tr > *:first-child{border-left:none} table.tt tr > *:last-child{border-right:none}
table.tt thead tr:first-child > *{border-top:none} table.tt tbody tr:last-child > *{border-bottom:none}
table.tt thead tr.time th{background:#f6f1ff;font-weight:400;font-size:10.5px;color:#857a9e}
table.tt thead tr.num th{background:#ebe1fe;font-size:15px;font-weight:600;color:#6f52c0}
table.tt th.day{width:92px;background:#faf7ff;color:#6f52c0;font-weight:600;font-size:13px}
table.tt .brk{background:#f3eef9 !important;color:#a99cc4;width:52px;font-weight:500}
table.tt td.brk{background:repeating-linear-gradient(45deg,#faf7ff,#faf7ff 6px,#f3edfc 6px,#f3edfc 12px) !important}
table.tt tbody td{height:60px}
table.tt td.slot{line-height:1.25;transition:filter .2s}
table.tt td.slot:hover{filter:brightness(.97) saturate(1.2)}
table.tt td.slot b{display:block;color:#5b3fa8;font-size:12px;font-weight:600}
table.tt td.slot span{display:block;font-size:11px;color:#8b6dda}
table.tt td.slot i{display:block;font-style:normal;font-size:10.5px;color:#857a9e}
table.tt td.dim{opacity:.22}
"""


def grid_html(ds, schedule, view, eid, search="", colors=None):
    g = grid_data(ds, schedule, view, eid)
    esc = html_lib.escape
    q = search.strip().lower()
    h = ['<div class="tt-wrap"><table class="tt"><thead><tr class="time"><th class="day"></th>']
    for p in g["periods"]:
        h.append(f'<th class="{"brk" if p in BREAK_PERIODS else ""}">{esc(g["times"][p])}</th>')
    h.append('</tr><tr class="num"><th class="day"></th>')
    for p in g["periods"]:
        h.append('<th class="brk">พัก</th>' if p in BREAK_PERIODS else f"<th>{p}</th>")
    h.append("</tr></thead><tbody>")
    for d in g["days"]:
        h.append(f'<tr><th class="day">วัน{DAY_TH[d]}</th>')
        for p in g["periods"]:
            if p in BREAK_PERIODS:
                h.append('<td class="brk">พัก</td>')
                continue
            c = g["cells"].get((d, p))
            if not c:
                h.append("<td></td>")
                continue
            text = " ".join(map(str, c["lines"])) + " " + c["title"]
            dim = " dim" if q and q not in text.lower() else ""
            bg = f' style="background:{colors.get(c["subject_id"], "#fff")}"' if colors else ""
            l1, l2, l3 = (list(c["lines"]) + ["", "", ""])[:3]
            h.append(f'<td class="slot{dim}"{bg} title="{esc(c["title"])}"><b>{esc(str(l1))}</b>'
                     f'<span>{esc(str(l2))}</span><i>{esc(str(l3))}</i></td>')
        h.append("</tr>")
    h.append("</tbody></table></div>")
    return "".join(h)


def safe_sheet_name(name, used):
    base = re.sub(r"[\[\]:*?/\\]", "-", str(name))[:28] or "sheet"
    cand, i = base, 2
    while cand in used:
        cand = f"{base[:25]}({i})"
        i += 1
    used.add(cand)
    return cand


def build_excel(ds, schedule, view, targets, checks=None):
    """Excel: 1 ชีตต่อ 1 กลุ่ม/ครู/ห้อง (แบบตารางเรียนในโจทย์) + ชีต output + ผลตรวจเงื่อนไข"""
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter

    thin = Side(style="thin", color="9CA3AF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    fill = lambda c: PatternFill("solid", fgColor=c)
    F = lambda **kw: Font(name="Tahoma", **kw)
    wb = Workbook()
    wb.remove(wb.active)
    used = set()
    colors = color_map(schedule)

    def table(ws, df, r, title=None):
        if title:
            ws.cell(r, 1, title).font = F(size=11, bold=True)
            r += 1
        for j, col in enumerate(df.columns, 1):
            c = ws.cell(r, j, col)
            c.font, c.fill, c.border, c.alignment = F(size=10, bold=True), fill("DBEAFE"), border, center
        for row in df.itertuples(index=False):
            r += 1
            for j, v in enumerate(row, 1):
                c = ws.cell(r, j, None if (not isinstance(v, str) and pd.isna(v)) else (v.item() if hasattr(v, "item") else v))
                c.font, c.border = F(size=10), border
        return r + 2

    for eid in targets:
        g = grid_data(ds, schedule, view, eid)
        ws = wb.create_sheet(safe_sheet_name(eid, used))
        last = 1 + len(g["periods"])
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=last)
        c = ws.cell(1, 1, entity_title(ds, view, eid))
        c.font, c.alignment = F(size=14, bold=True), Alignment(horizontal="center")
        for j, p in enumerate(g["periods"], 2):
            brk = p in BREAK_PERIODS
            for rr, val, color in ((3, g["times"][p], "DBEAFE"), (4, "พัก" if brk else p, "BFDBFE")):
                c = ws.cell(rr, j, val)
                c.font = F(size=9 if rr == 3 else 12, bold=rr == 4, color="6B7280" if brk else "111827")
                c.fill, c.border, c.alignment = fill("D1D5DB" if brk else color), border, center
            ws.column_dimensions[get_column_letter(j)].width = 6 if brk else 13
        for rr in (3, 4):
            ws.cell(rr, 1).fill, ws.cell(rr, 1).border = fill("DBEAFE"), border
        for i, d in enumerate(g["days"]):
            rr = 5 + i
            c = ws.cell(rr, 1, f"วัน{DAY_TH[d]}")
            c.font, c.fill, c.border, c.alignment = F(size=11, bold=True, color="1D4ED8"), fill("F3F4F6"), border, center
            for j, p in enumerate(g["periods"], 2):
                c = ws.cell(rr, j)
                c.border, c.alignment = border, center
                if p in BREAK_PERIODS:
                    c.value, c.fill, c.font = "พัก", fill("E5E7EB"), F(size=9, color="6B7280")
                    continue
                cell = g["cells"].get((d, p))
                if cell:
                    c.value = "\n".join(str(x) for x in cell["lines"] if x)
                    c.font = F(size=9, color="1D4ED8")
                    c.fill = fill(colors.get(cell["subject_id"], "#FFFFFF").lstrip("#"))
            ws.row_dimensions[rr].height = 46
        ws.column_dimensions["A"].width = 14
        ws.row_dimensions[3].height = 18
        ws.row_dimensions[4].height = 22
        lg = legend_df(ds, schedule, view, eid)
        if len(lg):
            table(ws, lg, 6 + len(g["days"]) + 1, "รายวิชา")
        ws.page_setup.orientation = "landscape"
        ws.page_setup.paperSize = ws.PAPERSIZE_A4
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 1

    ws = wb.create_sheet("output")
    sub = schedule[schedule[VIEWS[view]["col"]].isin(targets)] if not schedule.empty else schedule
    table(ws, output_detail_df(sub), 1)
    ws.freeze_panes = "A2"
    for j, w in enumerate([10, 12, 8, 8, 14, 11, 11], 1):
        ws.column_dimensions[get_column_letter(j)].width = w
    if checks is not None and len(checks):
        ws = wb.create_sheet("ผลตรวจเงื่อนไข")
        table(ws, checks, 1)
        for j, w in enumerate([6, 60, 12, 70], 1):
            ws.column_dimensions[get_column_letter(j)].width = w
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ---------------------------------------------------------------------------
# PDF (fpdf2) — ต้องมีฟอนต์ไทย: Windows ใช้ Tahoma ได้เลย หรือวางไฟล์ .ttf ไว้ในโฟลเดอร์ fonts/
# ---------------------------------------------------------------------------
FONT_CANDIDATES = [
    ("C:/Windows/Fonts/tahoma.ttf", "C:/Windows/Fonts/tahomabd.ttf"),
    ("C:/Windows/Fonts/LeelawUI.ttf", "C:/Windows/Fonts/LeelaUIb.ttf"),
    ("C:/Windows/Fonts/leelawad.ttf", "C:/Windows/Fonts/leelawdb.ttf"),
    ("/System/Library/Fonts/Supplemental/Tahoma.ttf", "/System/Library/Fonts/Supplemental/Tahoma Bold.ttf"),
    ("/Library/Fonts/Tahoma.ttf", None),
    ("/usr/share/fonts/truetype/noto/NotoSansThai-Regular.ttf", "/usr/share/fonts/truetype/noto/NotoSansThai-Bold.ttf"),
    ("/usr/share/fonts/opentype/tlwg/Loma.otf", "/usr/share/fonts/opentype/tlwg/Loma-Bold.otf"),
    ("/usr/share/fonts/truetype/tlwg/Loma.ttf", "/usr/share/fonts/truetype/tlwg/Loma-Bold.ttf"),
]


def find_thai_font():
    local = sorted((APP_DIR / "fonts").glob("*.ttf")) if (APP_DIR / "fonts").exists() else []
    if local:
        bold = [f for f in local if "bold" in f.name.lower() or f.stem.lower().endswith(("bd", "b"))]
        reg = [f for f in local if f not in bold] or local
        return str(reg[0]), str(bold[0]) if bold else None
    for reg, bold in FONT_CANDIDATES:
        if Path(reg).exists():
            return reg, bold if bold and Path(bold).exists() else None
    return None, None


def build_pdf(ds, schedule, view, targets):
    from fpdf import FPDF
    from fpdf.fonts import FontFace

    reg, bold = find_thai_font()
    if not reg:
        raise RuntimeError("ไม่พบฟอนต์ภาษาไทย — วางไฟล์ฟอนต์ .ttf (เช่น THSarabunNew.ttf) ไว้ในโฟลเดอร์ fonts/ ข้าง app.py")
    pdf = FPDF(orientation="L", unit="mm", format="A4")
    pdf.set_auto_page_break(auto=True, margin=10)
    pdf.add_font("TH", "", reg)
    pdf.add_font("TH", "B", bold or reg)
    try:
        pdf.set_text_shaping(True)   # จัดวางสระ/วรรณยุกต์ไทยให้ถูกต้อง (ต้องมี uharfbuzz)
    except Exception:
        pass
    colors = color_map(schedule)
    hexrgb = lambda h: tuple(int(h.lstrip("#")[i:i + 2], 16) for i in (0, 2, 4))

    for eid in targets:
        g = grid_data(ds, schedule, view, eid)
        pdf.add_page()
        pdf.set_font("TH", "B", 15)
        pdf.cell(0, 9, entity_title(ds, view, eid), align="C", new_x="LMARGIN", new_y="NEXT")
        x0, y0 = pdf.l_margin, pdf.get_y() + 2
        day_w = 22
        widths = [8 if p in BREAK_PERIODS else 1 for p in g["periods"]]
        unit = (pdf.w - pdf.l_margin - pdf.r_margin - day_w - sum(w for w in widths if w == 8)) / max(
            sum(1 for w in widths if w == 1), 1)
        col_w = [8 if p in BREAK_PERIODS else unit for p in g["periods"]]
        row_h = 17

        def box(x, y, w, h, fill_rgb, lines, sizes, rgbs, styles):
            pdf.set_fill_color(*fill_rgb)
            pdf.set_draw_color(156, 163, 175)
            pdf.rect(x, y, w, h, style="DF")
            total = sum(s * 0.42 for s in sizes)
            yy = y + (h - total) / 2
            for text, size, rgb, sty in zip(lines, sizes, rgbs, styles):
                pdf.set_font("TH", sty, size)
                pdf.set_text_color(*rgb)
                pdf.set_xy(x, yy)
                pdf.cell(w, size * 0.42, str(text), align="C")
                yy += size * 0.42

        x = x0 + day_w
        box(x0, y0, day_w, 12, (219, 234, 254), [""], [8], [(0, 0, 0)], [""])
        for p, w in zip(g["periods"], col_w):
            brk = p in BREAK_PERIODS
            box(x, y0, w, 5, (209, 213, 219) if brk else (219, 234, 254), ["" if brk else g["times"][p]],
                [6.5], [(55, 65, 81)], [""])
            box(x, y0 + 5, w, 7, (209, 213, 219) if brk else (191, 219, 254), ["พัก" if brk else p],
                [8 if brk else 11], [(107, 114, 128) if brk else (17, 24, 39)], ["B"])
            x += w
        y = y0 + 12
        for d in g["days"]:
            box(x0, y, day_w, row_h, (243, 244, 246), [f"วัน{DAY_TH[d]}"], [10], [(29, 78, 216)], ["B"])
            x = x0 + day_w
            for p, w in zip(g["periods"], col_w):
                if p in BREAK_PERIODS:
                    box(x, y, w, row_h, (229, 231, 235), ["พัก"], [7], [(107, 114, 128)], [""])
                else:
                    c = g["cells"].get((d, p))
                    if c:
                        lines = [str(v) for v in c["lines"] if v]
                        box(x, y, w, row_h, hexrgb(colors.get(c["subject_id"], "#FFFFFF")), lines,
                            [8] + [7] * (len(lines) - 1), [(29, 78, 216)] + [(55, 65, 81)] * (len(lines) - 1),
                            ["B"] + [""] * (len(lines) - 1))
                    else:
                        box(x, y, w, row_h, (255, 255, 255), [""], [7], [(0, 0, 0)], [""])
                x += w
            y += row_h

        lg = legend_df(ds, schedule, view, eid)
        if len(lg):
            pdf.set_xy(x0, y + 4)
            pdf.set_text_color(17, 24, 39)
            pdf.set_font("TH", "", 8)
            with pdf.table(col_widths=[18, 52, 12, 14] + [52] * (len(lg.columns) - 4), text_align="LEFT",
                           line_height=4.6,
                           headings_style=FontFace(emphasis="BOLD", fill_color=(219, 234, 254))) as tbl:
                hdr = tbl.row()
                for col in lg.columns:
                    hdr.cell(str(col))
                for row in lg.itertuples(index=False):
                    r = tbl.row()
                    for v in row:
                        r.cell(str(v))
        pdf.set_text_color(0, 0, 0)
    return bytes(pdf.output())


def build_html_report(ds, schedule, view, targets):
    """HTML พร้อมพิมพ์ (A4 แนวนอน) — เปิดในเบราว์เซอร์แล้วกด Ctrl+P บันทึกเป็น PDF ได้"""
    esc = html_lib.escape
    colors = color_map(schedule)
    pages = []
    for eid in targets:
        lg = legend_df(ds, schedule, view, eid)
        pages.append(f'<section class="page"><h2>{esc(entity_title(ds, view, eid))}</h2>'
                     f'{grid_html(ds, schedule, view, eid, colors=colors)}'
                     f'{lg.to_html(index=False, classes="lg", border=0) if len(lg) else ""}</section>')
    css = (GRID_CSS + "@page{size:A4 landscape;margin:10mm}body{font-family:'Sarabun','Tahoma',sans-serif;"
           "margin:20px;color:#111827;background:#fff}h2{text-align:center;margin:0 0 8px;font-size:20px}"
           ".page{page-break-after:always;margin-bottom:30px}.page:last-child{page-break-after:auto}"
           "table.tt{min-width:0}table.lg{border-collapse:collapse;font-size:12px;margin-top:6px}"
           "table.lg th,table.lg td{border:1px solid #9ca3af;padding:3px 8px;text-align:left}"
           "table.lg th{background:#dbeafe}@media print{body{margin:0}}")
    return ('<!doctype html><html lang="th"><head><meta charset="utf-8"><title>ตารางเรียน</title>'
            '<link href="https://fonts.googleapis.com/css2?family=Sarabun:wght@400;700&display=swap" rel="stylesheet">'
            f"<style>{css}</style></head><body>{''.join(pages)}</body></html>")


# ---------------------------------------------------------------------------
# 6) สถิติช่วยตัดสินใจ
# ---------------------------------------------------------------------------
def build_stats(ds, schedule):
    days, periods, _ = timeslot_layout(ds)
    teach_slots = sum(1 for d in days for p in periods if p not in BREAK_PERIODS
                      and ((ds["timeslots"]["day"] == d) & (ds["timeslots"]["period"] == p)).any())
    sch = schedule if not schedule.empty else pd.DataFrame(columns=SCHEDULE_COLS)
    lk = lookups(ds)

    t_rows = []
    for t in ds["teachers"].itertuples(index=False):
        sub = sch[sch["teacher_id"] == t.teacher_id]
        uniq = sub.drop_duplicates(["day", "period"])
        t_rows.append({"รหัสครู": t.teacher_id, "ชื่อครู": t.teacher_name,
                       "ตำแหน่ง": "หัวหน้า (Leader)" if t.role == "Leader" else "ครูผู้สอน",
                       "คาบสอน/สัปดาห์": len(uniq), "จำนวนวิชา": sub["subject_id"].nunique(),
                       "กลุ่มที่สอน": sub["group_id"].nunique(), "วันที่มาสอน": uniq["day"].nunique()})
    r_rows = []
    for r in ds["rooms"].itertuples(index=False):
        used = sch[sch["room_id"] == r.room_id].drop_duplicates(["day", "period"]).shape[0]
        r_rows.append({"รหัสห้อง": r.room_id, "ชื่อห้อง": r.room_name, "ประเภทห้อง": r.room_type,
                       "ระบบจัดเป็น": {"theory": "ทฤษฎี", "practice": "ปฏิบัติ", "general": "ใช้ได้ทั้งสองแบบ"}[r.category],
                       "คาบที่ใช้": used, "คาบว่าง": teach_slots - used,
                       "อัตราการใช้ (%)": round(used / teach_slots * 100, 1) if teach_slots else 0})
    g_rows = []
    for g in ds["groups"].itertuples(index=False):
        sub = sch[sch["group_id"] == g.group_id]
        per_day = sub.groupby("day").size()
        g_rows.append({"รหัสกลุ่ม": g.group_id, "ชื่อกลุ่ม": g.group_name,
                       "ครูที่ปรึกษา": lk["teacher"].get(g.advisor_id, g.advisor) if (g.advisor_id or g.advisor) else "-",
                       "จำนวนวิชา": sub.loc[sub["kind"] != "homeroom", "subject_id"].nunique(),
                       "คาบเรียน/สัปดาห์": len(sub), "วันที่มาเรียน": sub["day"].nunique(),
                       "เรียนมากสุด/วัน": int(per_day.max()) if len(per_day) else 0})
    return {"teacher_df": pd.DataFrame(t_rows), "room_df": pd.DataFrame(r_rows),
            "group_df": pd.DataFrame(g_rows), "slots": teach_slots, "days": days}


# =============================================================================
# 7) บัญชีผู้ใช้ / รหัสผ่าน / การเผยแพร่
#
#   - ผู้ดูแลระบบ / ฝ่ายวิชาการ : ชื่อผู้ใช้ admin
#   - ครูผู้สอน                 : ชื่อผู้ใช้ = รหัสครู (เช่น T01) หรือชื่อ-นามสกุล
#   - นักเรียน                   : ชื่อผู้ใช้ = รหัสกลุ่มเรียน (เช่น G1) / ชื่อกลุ่ม
#                                   หรือรหัสนักเรียน (ถ้าอัปโหลดไฟล์ student.csv)
#   รหัสผ่านเริ่มต้นเหมือนกันทุกบัญชี (DEFAULT_PASSWORD) — เปลี่ยนเองได้ภายหลัง
#   ผู้ดูแลระบบรีเซ็ตรหัสผ่านกลับเป็นค่าเริ่มต้นได้
#   บัญชีครู/นักเรียนสร้างจากข้อมูลในตารางที่เผยแพร่ล่าสุด (ไม่ต้องสมัครเอง)
# =============================================================================
DATA_DIR = APP_DIR / "app_data"
PUBLISHED_FILE = DATA_DIR / "published.pkl"
DRAFT_FILE = DATA_DIR / "draft.pkl"
PASSWORD_FILE = DATA_DIR / "passwords.json"
LOG_FILE = DATA_DIR / "activity_log.json"
ADMIN_USERNAME = "admin"
MAX_LOGIN_FAIL = 5
LOCK_SECONDS = 30
MIN_PASSWORD_LEN = 4
PBKDF2_ROUNDS = 120_000

ROLE_INFO = {
    "admin": {"label": "ผู้ดูแลระบบ / ฝ่ายวิชาการ", "icon": "admin_panel_settings", "tone": "lv",
              "desc": "นำเข้าข้อมูล จัดตาราง ตรวจสอบ ส่งออก อนุมัติเผยแพร่ และจัดการรหัสผ่านผู้ใช้"},
    "teacher": {"label": "ครูผู้สอน", "icon": "co_present", "tone": "pink", "desc": "ดูตารางสอนและภาระงานสอนของตนเอง"},
    "student": {"label": "นักเรียน / นักศึกษา", "icon": "school", "tone": "mint", "desc": "ดูตารางเรียนและรายวิชาของกลุ่มเรียนตนเอง"},
}

MENU_HOME = "ภาพรวม & เผยแพร่"
MENU_DATA = "นำเข้าข้อมูล & จัดตาราง"
MENU_VIEW = "ตารางเรียน/ตารางสอน & ส่งออก"
MENU_USERS = "ผู้ใช้งาน & รหัสผ่าน"
MENU_T_TT = "ตารางสอนของฉัน"
MENU_T_LOAD = "ภาระงานสอนของฉัน"
MENU_S_TT = "ตารางเรียนของฉัน"
MENU_S_SUBJ = "รายวิชาและครูผู้สอน"
MENU_HIST = "ประวัติข้อมูล & ลบข้อมูล"
MENU_ICONS = {MENU_HOME: "space_dashboard", MENU_DATA: "upload_file", MENU_VIEW: "calendar_view_week",
              MENU_USERS: "manage_accounts", MENU_T_TT: "event_note", MENU_T_LOAD: "assignment_ind",
              MENU_S_TT: "event_note", MENU_S_SUBJ: "auto_stories", MENU_HIST: "inventory_2"}
ROLE_MENUS = {
    "admin": [MENU_HOME, MENU_DATA, MENU_VIEW, MENU_HIST, MENU_USERS],
    "teacher": [MENU_T_TT, MENU_T_LOAD],
    "student": [MENU_S_TT, MENU_S_SUBJ],
}


def default_password():
    try:
        pw = st.secrets.get("DEFAULT_PASSWORD")
    except Exception:
        pw = None
    return str(pw or os.environ.get("DEFAULT_PASSWORD") or "1234")


def read_json(path, default):
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return default


def write_json(path, data):
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(".tmp")
    tmp.write_text(json.dumps(data, ensure_ascii=False, indent=1), encoding="utf-8")
    tmp.replace(path)


def save_pickle(path, obj):
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(".tmp")
    pd.to_pickle(obj, tmp)
    tmp.replace(path)


def load_pickle(path):
    try:
        return pd.read_pickle(path) if path.exists() else None
    except Exception:
        return None


def load_published():
    return load_pickle(PUBLISHED_FILE)


def log_event(user, action, detail=""):
    try:
        logs = read_json(LOG_FILE, [])
        logs.append({"เวลา": now_text(), "ผู้ใช้": (user or {}).get("username", "-"),
                     "สิทธิ์": ROLE_INFO.get((user or {}).get("role"), {}).get("label", "-"),
                     "กิจกรรม": action, "รายละเอียด": detail})
        write_json(LOG_FILE, logs[-500:])
    except Exception:
        pass


# ---- ประวัติชุดข้อมูล: เก็บทุกชุดที่นำเข้า (ไฟล์ต้นฉบับ + ข้อมูล + ตารางล่าสุด) เพื่อนำกลับมาใช้ใหม่ ----
HISTORY_DIR = DATA_DIR / "history"
HISTORY_INDEX = HISTORY_DIR / "index.json"


def history_index():
    return read_json(HISTORY_INDEX, [])


def write_history_index(items):
    HISTORY_DIR.mkdir(parents=True, exist_ok=True)
    write_json(HISTORY_INDEX, items)


def history_path(hid):
    return HISTORY_DIR / f"{hid}.pkl"


def save_history_entry(hid, data):
    HISTORY_DIR.mkdir(parents=True, exist_ok=True)
    tmp = history_path(hid).with_suffix(".tmp")
    pd.to_pickle(data, tmp)
    tmp.replace(history_path(hid))


def ds_counts(ds):
    return {"teachers": len(ds["teachers"]), "groups": len(ds["groups"]), "subjects": len(ds["subjects"]),
            "rooms": len(ds["rooms"]), "register": len(ds["register"])}


def archive_upload(files, ds, report):
    """เก็บชุดข้อมูลที่นำเข้าลงประวัติ (ไฟล์ชุดเดิมซ้ำ -> อัปเดตรายการเดิม ไม่สร้างใหม่) คืนค่า id"""
    h = hashlib.sha1()
    for name, data in sorted(files):
        h.update(name.encode("utf-8") + b"\0" + data)
    digest = h.hexdigest()
    items = history_index()
    hit = next((it for it in items if it.get("digest") == digest), None)
    hid = hit["id"] if hit else datetime.now(TZ).strftime("%Y%m%d%H%M%S") + "_" + secrets.token_hex(3)
    old = (load_pickle(history_path(hid)) or {}) if hit else {}
    old.update({"files": list(files), "ds": ds, "file_report": report})
    save_history_entry(hid, old)
    meta = {"updated_at": now_text(), "files": [n for n, _ in files], "counts": ds_counts(ds),
            "size": sum(len(d) for _, d in files)}
    if hit:
        hit.update(meta)
    else:
        items.append({"id": hid, "name": f"ชุดข้อมูล {now_text()}", "created_at": now_text(), "digest": digest,
                      "placed": None, "required": None, **meta})
    write_history_index(items)
    return hid


def save_history_result(hid, result, checks):
    data = load_pickle(history_path(hid))
    if not data:
        return
    data.update({"result": result, "checks": checks})
    save_history_entry(hid, data)
    items = history_index()
    for it in items:
        if it["id"] == hid:
            it.update({"placed": result["stats"]["placed"], "required": result["stats"]["required"],
                       "updated_at": now_text()})
    write_history_index(items)


def history_zip(hid):
    """ไฟล์ต้นฉบับของชุดข้อมูล รวมเป็น ZIP (อัปโหลดกลับเข้าระบบได้ทันที)"""
    data = load_pickle(history_path(hid)) or {}
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        used = set()
        for name, raw in data.get("files", []):
            base = Path(name).name
            n, i = base, 2
            while n in used:
                n = f"{Path(base).stem}({i}){Path(base).suffix}"
                i += 1
            used.add(n)
            zf.writestr(n, raw)
        if data.get("result") is not None:
            zf.writestr("output.csv", df_to_csv_bytes(output_df(data["result"]["schedule"])))
    return buf.getvalue()


def delete_history(hids):
    hids = set(hids)
    for hid in hids:
        try:
            history_path(hid).unlink(missing_ok=True)
        except OSError:
            pass
    write_history_index([it for it in history_index() if it["id"] not in hids])
    if st.session_state.get("history_id") in hids:
        st.session_state.pop("history_id", None)


# ---- รหัสผ่าน: เก็บเฉพาะบัญชีที่เปลี่ยนรหัสแล้ว (hash PBKDF2 + salt) ----
def hash_pw(password, salt=None):
    salt = salt or secrets.token_hex(16)
    return salt, hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), bytes.fromhex(salt), PBKDF2_ROUNDS).hex()


def check_password(key, password):
    store = read_json(PASSWORD_FILE, {})
    if key in store:
        return hmac.compare_digest(hash_pw(password, store[key]["salt"])[1], store[key]["hash"])
    return hmac.compare_digest(str(password), default_password())


def uses_default(key):
    return key not in read_json(PASSWORD_FILE, {})


def set_password(key, password):
    store = read_json(PASSWORD_FILE, {})
    salt, h = hash_pw(password)
    store[key] = {"salt": salt, "hash": h, "changed_at": now_text()}
    write_json(PASSWORD_FILE, store)


def reset_passwords(keys):
    store = read_json(PASSWORD_FILE, {})
    for k in keys:
        store.pop(k, None)
    write_json(PASSWORD_FILE, store)


def squash(text):
    return re.sub(r"\s+", "", str(text or "")).lower()


def account_directory(pub):
    accounts = [{"key": "admin", "username": ADMIN_USERNAME, "role": "admin", "link": None,
                 "name": ROLE_INFO["admin"]["label"], "aliases": []}]
    if pub:
        ds = pub["ds"]
        for t in ds["teachers"].itertuples(index=False):
            accounts.append({"key": f"teacher:{t.teacher_id}", "username": t.teacher_id, "role": "teacher",
                             "link": t.teacher_id, "name": t.teacher_name, "aliases": [t.teacher_name]})
        for g in ds["groups"].itertuples(index=False):
            accounts.append({"key": f"group:{g.group_id}", "username": g.group_id, "role": "student",
                             "link": g.group_id, "name": f"นักเรียนกลุ่ม {g.group_name}", "aliases": [g.group_name]})
        for s in ds["students"].itertuples(index=False):
            accounts.append({"key": f"student:{s.student_id}", "username": s.student_id, "role": "student",
                             "link": s.group_id, "name": s.student_name, "aliases": [s.student_name]})
    return accounts


def find_account(username, pub):
    u = squash(username)
    if not u:
        return None
    accounts = account_directory(pub)
    for acc in accounts:
        if squash(acc["username"]) == u:
            return acc
    for acc in accounts:
        if any(squash(a) == u for a in acc["aliases"]):
            return acc
    return None


def do_login(username, password):
    ss = st.session_state
    if time.time() < ss.get("login_lock_until", 0):
        ss["login_error"] = f"ระงับการเข้าสู่ระบบชั่วคราว กรุณารออีก {int(ss['login_lock_until'] - time.time()) + 1} วินาที"
        return
    pub = load_published()
    acc = find_account(username, pub)
    if acc is None or not check_password(acc["key"], password):
        ss["login_fail"] = ss.get("login_fail", 0) + 1
        log_event({"username": str(username).strip() or "-"}, "เข้าสู่ระบบไม่สำเร็จ")
        if ss["login_fail"] >= MAX_LOGIN_FAIL:
            ss["login_fail"], ss["login_lock_until"] = 0, time.time() + LOCK_SECONDS
            ss["login_error"] = f"ใส่ผิด {MAX_LOGIN_FAIL} ครั้ง ระงับการเข้าสู่ระบบ {LOCK_SECONDS} วินาที"
        else:
            hint = "" if pub or acc else " — ครู/นักเรียนจะเข้าสู่ระบบได้หลังฝ่ายวิชาการเผยแพร่ตาราง"
            ss["login_error"] = f"ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง (ลองได้อีก {MAX_LOGIN_FAIL - ss['login_fail']} ครั้ง){hint}"
        return
    user = {k: acc[k] for k in ("key", "username", "role", "link", "name")}
    ss["auth_user"], ss["login_fail"] = user, 0
    ss.pop("login_error", None)
    ss["main_menu"] = ROLE_MENUS[user["role"]][0]
    log_event(user, "เข้าสู่ระบบ")


def do_logout():
    log_event(st.session_state.get("auth_user"), "ออกจากระบบ")
    for k in ("auth_user", "main_menu", "flash"):
        st.session_state.pop(k, None)


def go_to(menu):
    st.session_state["main_menu"] = menu


# =============================================================================
# 7.5) ธีมหน้าเว็บ — ม่วงพาสเทล (Pastel Lavender) + ไอคอน Material Symbols
#      ไอคอนใน widget ของ Streamlit ใช้รูปแบบ ":material/ชื่อไอคอน:"
#      ไอคอนใน HTML ที่เขียนเอง ใช้ mi("ชื่อไอคอน")
#      รายชื่อไอคอนทั้งหมด: https://fonts.google.com/icons (สไตล์ Rounded)
# =============================================================================
THEME_CSS = """
@import url('https://fonts.googleapis.com/css2?family=Outfit:wght@500;600;700&family=Prompt:wght@300;400;500;600&family=Inter:wght@300;400;500;600&display=swap');
@import url('https://fonts.googleapis.com/css2?family=Material+Symbols+Rounded:opsz,wght,FILL,GRAD@20..48,300..600,0..1,0&display=block');

:root{
  --lv-50:#fbf8ff; --lv-100:#f5efff; --lv-200:#ebe1fe; --lv-300:#dccbfb; --lv-400:#c3a9f5;
  --lv-500:#a68aec; --lv-600:#8b6dda; --lv-700:#6f52c0; --ink:#3a2c5a; --muted:#857a9e;
  --pink:#fbd3e6; --peach:#ffe3d3; --mint:#d4f2e7; --sky:#dce9ff; --line:#ece4fb;
  --ease:cubic-bezier(.16,1,.3,1);
  --shadow:0 10px 30px -12px rgba(111,82,192,.25);
  --shadow-lg:0 22px 45px -18px rgba(111,82,192,.38);
  --font-body:'Inter','Prompt',system-ui,sans-serif;
  --font-head:'Outfit','Prompt',system-ui,sans-serif;
}

/* ---------- พื้นหลัง + ฟอนต์ ---------- */
html,body,.stApp{font-family:var(--font-body);color:var(--ink)}
.stApp{background:
  radial-gradient(42rem 32rem at 0% -10%,#efe5ff 0%,transparent 60%),
  radial-gradient(36rem 30rem at 100% 0%,#ffe6f2 0%,transparent 58%),
  radial-gradient(40rem 34rem at 60% 115%,#e2eeff 0%,transparent 60%),
  var(--lv-50);background-attachment:fixed}
.stApp p,.stApp li,.stApp label,.stApp input,.stApp textarea,.stApp button,.stApp select{font-family:var(--font-body)}
.stApp h1,.stApp h2,.stApp h3,.stApp h4,.stApp h5{font-family:var(--font-head);color:var(--ink);letter-spacing:-.01em}
.stApp h3{font-weight:600;font-size:1.35rem;margin-top:.6rem}
.stApp h4{font-size:1.12rem;font-weight:600}
.stApp h5{color:var(--lv-700);font-weight:600}
[data-testid="stIconMaterial"],.mi{font-family:'Material Symbols Rounded' !important;font-weight:400;
  font-style:normal;line-height:1;letter-spacing:normal;text-transform:none;white-space:nowrap;
  font-feature-settings:'liga';-webkit-font-smoothing:antialiased;font-variation-settings:'FILL' 0,'wght' 400}
.stApp h1 span[role="img"],.stApp h2 span[role="img"],.stApp h3 span[role="img"],.stApp h4 span[role="img"],
.stApp h5 span[role="img"],.stApp h1 [data-testid="stIconMaterial"],.stApp h3 [data-testid="stIconMaterial"]{color:var(--lv-600)}
header[data-testid="stHeader"]{background:transparent}
[data-testid="stMainBlockContainer"],.block-container{padding-top:2.2rem;max-width:1320px}
::selection{background:var(--lv-300)}

/* ---------- แถบด้านข้าง ---------- */
section[data-testid="stSidebar"]{background:rgba(255,255,255,.72);backdrop-filter:blur(14px);
  border-right:1px solid var(--line)}
section[data-testid="stSidebar"] [data-testid="stRadio"] > label{display:none}
section[data-testid="stSidebar"] [role="radiogroup"]{gap:.3rem;width:100%;display:flex;flex-direction:column;align-items:stretch}
section[data-testid="stSidebar"] [data-testid="stElementContainer"]:has([data-testid="stRadio"]),
section[data-testid="stSidebar"] [data-testid="stRadio"],section[data-testid="stSidebar"] [data-testid="stRadio"] > div{width:100% !important}
section[data-testid="stSidebar"] [role="radiogroup"] > div{width:100%}
section[data-testid="stSidebar"] [data-testid="stRadioOption"]{width:100%;box-sizing:border-box;display:flex;margin:0;padding:.62rem .85rem;
  border-radius:14px;cursor:pointer;transition:all .35s var(--ease);border:1px solid transparent}
section[data-testid="stSidebar"] [data-testid="stRadioOption"] > div > div:first-child:not([data-testid]){display:none}
section[data-testid="stSidebar"] [data-testid="stRadioOption"] [data-testid="stMarkdownContainer"] p{font-size:.94rem;font-weight:500;
  color:var(--ink);margin:0;display:flex;align-items:center;gap:.5rem;line-height:1.35}
section[data-testid="stSidebar"] [data-testid="stRadioOption"] > div{width:100%}
section[data-testid="stSidebar"] [data-testid="stRadioOption"] span[role="img"]{color:var(--lv-600);font-size:1.3rem;flex:none}
section[data-testid="stSidebar"] [data-testid="stRadioOption"]:hover{background:var(--lv-100);transform:translateX(3px)}
section[data-testid="stSidebar"] [data-testid="stRadioOption"][data-selected="true"]{
  background:linear-gradient(135deg,#b9a0f5 0%,#9a7de6 100%);box-shadow:var(--shadow)}
section[data-testid="stSidebar"] [data-testid="stRadioOption"][data-selected="true"] p,
section[data-testid="stSidebar"] [data-testid="stRadioOption"][data-selected="true"] span[role="img"]{color:#fff}
.side-brand{display:flex;align-items:center;gap:.7rem;margin:.2rem 0 1rem}
.side-brand .logo{width:44px;height:44px;border-radius:14px;display:grid;place-items:center;color:#fff;
  background:linear-gradient(135deg,#c7b1f8,#9a7de6);box-shadow:var(--shadow)}
.side-brand .logo .mi{font-size:26px}
.side-brand b{display:block;font-family:var(--font-head);font-size:1.15rem;color:var(--ink);line-height:1.1}
.side-brand small{color:var(--muted);font-size:.78rem}
.side-label{font-size:.72rem;text-transform:uppercase;letter-spacing:.12em;color:var(--muted);
  font-weight:600;margin:.6rem 0 .35rem .2rem}
.user-card{display:flex;align-items:center;gap:.7rem}
.user-card .av{width:42px;height:42px;flex:none;border-radius:50%;display:grid;place-items:center;
  background:linear-gradient(135deg,var(--pink),var(--lv-300));color:var(--lv-700)}
.user-card .av .mi{font-size:22px}
.user-card b{display:block;font-size:.95rem;line-height:1.25;color:var(--ink)}
.user-card small{color:var(--muted);font-size:.78rem;line-height:1.3;display:block}

/* ---------- ปุ่ม (ทรงแคปซูล) ---------- */
.stButton > button,.stDownloadButton > button,.stFormSubmitButton > button,[data-testid="stBaseButton-secondary"],
[data-testid="stBaseButton-primary"],[data-testid="stBaseButton-secondaryFormSubmit"],[data-testid="stBaseButton-primaryFormSubmit"]{
  border-radius:999px !important;font-weight:500;padding:.5rem 1.15rem;min-height:2.6rem;
  transition:transform .35s var(--ease),box-shadow .35s var(--ease),background .35s var(--ease)}
[data-testid="stBaseButton-secondary"],[data-testid="stBaseButton-secondaryFormSubmit"]{
  background:#fff;border:1px solid var(--lv-300);color:var(--lv-700)}
[data-testid="stBaseButton-secondary"]:hover,[data-testid="stBaseButton-secondaryFormSubmit"]:hover{
  background:var(--lv-100);border-color:var(--lv-500);color:var(--lv-700);transform:translateY(-2px);box-shadow:var(--shadow)}
[data-testid="stBaseButton-primary"],[data-testid="stBaseButton-primaryFormSubmit"]{
  background:linear-gradient(135deg,#b79df4 0%,#8f71df 100%);border:none;color:#fff;box-shadow:var(--shadow)}
[data-testid="stBaseButton-primary"]:hover,[data-testid="stBaseButton-primaryFormSubmit"]:hover{
  background:linear-gradient(135deg,#a98cef 0%,#7f60d6 100%);color:#fff;transform:translateY(-2px);box-shadow:var(--shadow-lg)}
button:disabled{opacity:.5;transform:none !important;box-shadow:none !important}
button:focus-visible{outline:3px solid var(--lv-300) !important;outline-offset:2px}

/* ---------- การ์ด / กล่อง ---------- */
[class*="st-key-card-"],div[data-testid="stForm"]{background:rgba(255,255,255,.86);border:1px solid var(--line) !important;
  border-radius:22px !important;box-shadow:0 6px 24px -14px rgba(111,82,192,.22);
  transition:box-shadow .45s var(--ease),transform .45s var(--ease);animation:fadeUp .8s var(--ease) both}
[class*="st-key-card-"]:hover{box-shadow:var(--shadow)}
[class*="st-key-card-"]{padding:1.1rem 1.25rem !important}
section[data-testid="stSidebar"] .st-key-card-user{padding:.9rem !important;animation:none}
div[data-testid="stForm"]{padding:1.4rem 1.4rem 1.1rem}

/* ---------- ตัวเลขสรุป (metric) ---------- */
[data-testid="stMetric"]{background:#fff;border:1px solid var(--line);border-radius:20px;padding:1rem 1.15rem;
  box-shadow:0 6px 22px -14px rgba(111,82,192,.25);transition:transform .45s var(--ease),box-shadow .45s var(--ease);
  animation:fadeUp .8s var(--ease) both}
[data-testid="stMetric"]:hover{transform:translateY(-4px);box-shadow:var(--shadow)}
[data-testid="stMetricLabel"] p{color:var(--muted);font-size:.84rem;font-weight:500}
[data-testid="stMetricLabel"] span[role="img"]{color:var(--lv-600);margin-right:.15rem}
[data-testid="stMetricValue"]{font-family:var(--font-head);color:var(--ink);font-weight:600}

/* ---------- ช่องกรอก / ตัวเลือก ---------- */
[data-baseweb="input"],[data-baseweb="select"] > div,[data-baseweb="textarea"],[data-testid="stNumberInputContainer"]{
  border-radius:14px !important;border-color:var(--lv-200) !important;background:#fff !important}
[data-baseweb="input"]:focus-within,[data-baseweb="select"] > div:focus-within{border-color:var(--lv-500) !important;
  box-shadow:0 0 0 3px rgba(166,138,236,.18)}
[data-baseweb="tag"]{background:var(--lv-200) !important;color:var(--lv-700) !important;border-radius:999px !important}
[data-baseweb="tag"] span{color:var(--lv-700) !important}
[data-testid="stFileUploaderDropzone"]{background:linear-gradient(135deg,#fdfbff,#f6f0ff);border:2px dashed var(--lv-300);
  border-radius:20px;transition:all .35s var(--ease)}
[data-testid="stFileUploaderDropzone"]:hover{border-color:var(--lv-500);background:var(--lv-100)}
[data-testid="stFileUploaderDropzone"] [data-testid="stIconMaterial"]{color:var(--lv-500)}

/* radio แนวนอนในหน้าหลัก -> ปุ่มแคปซูล */
[data-testid="stMain"] [role="radiogroup"]{gap:.45rem;flex-wrap:wrap}
[data-testid="stMain"] [data-testid="stRadioOption"]{margin:0;padding:.42rem 1rem .42rem .7rem;
  background:#fff;border:1px solid var(--lv-200);border-radius:999px;transition:all .3s var(--ease);cursor:pointer}
[data-testid="stMain"] [data-testid="stRadioOption"]:hover{border-color:var(--lv-400);background:var(--lv-100)}
[data-testid="stMain"] [data-testid="stRadioOption"][data-selected="true"]{
  background:var(--lv-200);border-color:var(--lv-500);box-shadow:0 4px 14px -8px rgba(111,82,192,.5)}
[data-testid="stMain"] [data-testid="stRadioOption"][data-selected="true"] p{color:var(--lv-700);font-weight:600}
[data-testid="stMain"] [data-testid="stRadioOption"] span[role="img"]{color:var(--lv-600)}

/* pills / segmented control */
[data-testid="stButtonGroup"] button{border-radius:999px !important;border-color:var(--lv-200) !important}
[data-testid="stButtonGroup"] button[aria-checked="true"],[data-testid="stButtonGroup"] button[kind*="Active"],
[data-testid="stButtonGroup"] button[aria-pressed="true"],[data-testid="stButtonGroup"] [data-testid*="Active"]{
  background:var(--lv-200) !important;border-color:var(--lv-500) !important;color:var(--lv-700) !important;opacity:1 !important;font-weight:600}
[data-testid="stButtonGroup"] [data-testid*="Active"] *{color:var(--lv-700) !important}

/* ---------- แท็บ ---------- */
[data-baseweb="tab-list"]{gap:.4rem;background:var(--lv-100);padding:.3rem;border-radius:999px;width:fit-content;max-width:100%}
[data-baseweb="tab"]{border-radius:999px !important;padding:.35rem 1rem !important;height:auto !important;transition:all .3s var(--ease)}
[data-baseweb="tab"][aria-selected="true"]{background:#fff;box-shadow:0 4px 12px -6px rgba(111,82,192,.45)}
[data-baseweb="tab"][aria-selected="true"] p{color:var(--lv-700);font-weight:600}
[data-baseweb="tab-highlight"],[data-baseweb="tab-border"]{display:none}

/* ---------- กล่องข้อความแจ้งเตือน ---------- */
[data-testid="stAlert"] > div,[data-testid="stAlertContainer"]{border-radius:18px !important;border:none}
[data-testid="stAlertContentInfo"]{color:#4a3d78}
[data-testid="stAlert"] [data-testid="stAlertContainer"]{box-shadow:0 4px 16px -12px rgba(111,82,192,.4)}
div[data-testid="stAlertContainer"]:has([data-testid="stAlertContentInfo"]){background:#efe8ff}
div[data-testid="stAlertContainer"]:has([data-testid="stAlertContentSuccess"]){background:#e3f6ee}
div[data-testid="stAlertContainer"]:has([data-testid="stAlertContentWarning"]){background:#fff3dc}
div[data-testid="stAlertContainer"]:has([data-testid="stAlertContentError"]){background:#ffe6ee}

/* ---------- ตาราง / expander ---------- */
[data-testid="stDataFrame"],[data-testid="stTable"]{border-radius:16px;overflow:hidden;border:1px solid var(--line);
  box-shadow:0 6px 20px -16px rgba(111,82,192,.35);background:#fff}
[data-testid="stExpander"] details{border-radius:18px !important;border:1px solid var(--line) !important;background:rgba(255,255,255,.8)}
[data-testid="stExpander"] summary:hover{color:var(--lv-700)}
[data-testid="stExpander"] summary [data-testid="stIconMaterial"]{color:var(--lv-600)}
hr{border-color:var(--line) !important}
[data-testid="stCaptionContainer"]{color:var(--muted)}
[data-testid="stCaptionContainer"] span[role="img"]{color:var(--lv-500)}
[data-baseweb="tab"] span[role="img"],[data-testid="stExpander"] summary span[role="img"]{color:var(--lv-600)}
code{background:var(--lv-100) !important;color:var(--lv-700) !important;border-radius:8px}

/* ---------- Hero ---------- */
.hero{position:relative;overflow:hidden;border-radius:28px;padding:1.35rem 2rem;margin:0 0 1.2rem;
  background:linear-gradient(120deg,#efe6ff 0%,#f9e9f6 55%,#e7efff 100%);border:1px solid #fff;
  box-shadow:var(--shadow);animation:fadeUp .9s var(--ease) both}
.hero .blob{position:absolute;border-radius:50%;filter:blur(38px);opacity:.75;animation:morph 14s ease-in-out infinite alternate}
.hero .b1{width:280px;height:280px;background:#d5c2ff;top:-120px;left:-60px}
.hero .b2{width:240px;height:240px;background:#ffcce4;bottom:-140px;left:42%;animation-delay:-4s}
.hero .b3{width:200px;height:200px;background:#c9ddff;top:-70px;right:16%;animation-delay:-8s}
.hero-inner{position:relative;z-index:2;max-width:calc(100% - 150px)}
.hero-tag{display:inline-flex;align-items:center;gap:.4rem;background:rgba(255,255,255,.75);color:var(--lv-700);
  border-radius:999px;padding:.28rem .8rem;font-size:.74rem;font-weight:600;letter-spacing:.1em;text-transform:uppercase}
.hero-tag .mi{font-size:16px}
.hero h1{font-family:var(--font-head);font-size:clamp(1.35rem,2.4vw,2.05rem);line-height:1.2;margin:.5rem 0 .3rem;
  color:var(--ink);font-weight:700;padding:0}
.hero h1 .accent{background:linear-gradient(90deg,#8b6dda,#d77fb3);-webkit-background-clip:text;background-clip:text;color:transparent}
.hero h1 .streak{display:inline-block;vertical-align:middle;height:.32em;width:clamp(60px,9vw,160px);margin-left:.4rem;
  border-radius:0 999px 999px 0;background:linear-gradient(90deg,#b49af3,#f3b5d4);transform-origin:left;
  animation:streak 1.5s var(--ease) .3s both}
.hero p{margin:0;color:#6a5e86;font-size:.95rem;display:flex;align-items:center;gap:.35rem;flex-wrap:wrap}
.hero p .mi{font-size:18px;color:var(--lv-600)}
.hero.big p{display:block}
.hero .chips{display:flex;flex-wrap:wrap;gap:.45rem;margin-top:.9rem}
.hero .chip{display:inline-flex;align-items:center;gap:.35rem;background:rgba(255,255,255,.8);border-radius:999px;
  padding:.32rem .8rem;font-size:.82rem;color:var(--ink)}
.hero .chip .mi{font-size:17px;color:var(--lv-600)}
.badge{position:absolute;right:1.6rem;top:50%;width:104px;height:104px;margin-top:-52px;z-index:2}
.hero.big .badge{width:124px;height:124px;margin-top:-62px}
.badge svg{width:100%;height:100%;animation:spin 14s linear infinite}
.badge text{font-family:var(--font-head);font-size:10.5px;letter-spacing:3.2px;fill:var(--lv-700);font-weight:600}
.badge .core{position:absolute;inset:27%;border-radius:50%;display:grid;place-items:center;color:#fff;
  background:linear-gradient(135deg,#c2a8f7,#9a7de6);box-shadow:var(--shadow)}
.badge .core .mi{font-size:30px}
.hero.big{padding:2.6rem 2.4rem;min-height:300px;display:flex;align-items:center}
.hero.big h1{font-size:clamp(2rem,3.6vw,3rem)}

/* หัวเรื่องของแต่ละหน้า */
.page-head{display:flex;align-items:center;gap:.9rem;margin:.2rem 0 1.1rem;animation:fadeUp .7s var(--ease) both}
.page-head .ic{width:52px;height:52px;flex:none;border-radius:17px;display:grid;place-items:center;color:var(--lv-700);
  background:linear-gradient(135deg,#f0e7ff,#ffe4f1);border:1px solid #fff;box-shadow:var(--shadow)}
.page-head .ic .mi{font-size:28px}
.page-head h2{font-family:var(--font-head);font-size:1.6rem;margin:0;padding:0;color:var(--ink);font-weight:600}
.page-head p{margin:.1rem 0 0;color:var(--muted);font-size:.92rem}

/* การ์ดบทบาท / ขั้นตอน / คาบวันนี้ */
.mini-card{background:#fff;border:1px solid var(--line);border-radius:20px;padding:1.1rem 1.15rem;margin-bottom:.55rem;
  box-shadow:0 6px 22px -16px rgba(111,82,192,.35);transition:transform .45s var(--ease),box-shadow .45s var(--ease);
  animation:fadeUp .8s var(--ease) both}
.mini-card:hover{transform:translateY(-4px);box-shadow:var(--shadow)}
.mini-card .ic{width:42px;height:42px;border-radius:13px;display:grid;place-items:center;margin-bottom:.6rem;color:var(--lv-700)}
.mini-card .ic .mi{font-size:23px}
.mini-card b{display:block;color:var(--ink);font-size:.98rem;margin-bottom:.15rem}
.mini-card small{display:block;color:var(--muted);font-size:.82rem;line-height:1.45}
.mini-card code{font-size:.8rem}
.mini-card.done{background:linear-gradient(160deg,#fff 30%,#f1ebff)}
.tone-lv{background:var(--lv-200)} .tone-pink{background:var(--pink)} .tone-mint{background:var(--mint)}
.tone-sky{background:var(--sky)} .tone-peach{background:var(--peach)}
.status-pill{display:inline-flex;align-items:center;gap:.25rem;font-size:.74rem;font-weight:600;border-radius:999px;
  padding:.15rem .55rem;margin-top:.5rem}
.status-pill .mi{font-size:15px}
.status-pill.ok{background:#dcf3e9;color:#2f7d5d} .status-pill.wait{background:var(--lv-100);color:var(--muted)}
.class-card .time{display:flex;align-items:center;gap:.35rem;color:var(--lv-700);font-weight:600;font-size:.9rem}
.class-card .time .mi{font-size:18px}
.class-card .subj{font-family:var(--font-head);font-size:1.05rem;color:var(--ink);margin:.35rem 0 .2rem;line-height:1.3}
.class-card .meta{display:flex;flex-wrap:wrap;gap:.6rem;color:var(--muted);font-size:.82rem}
.class-card .meta span{display:inline-flex;align-items:center;gap:.2rem}
.class-card .meta .mi{font-size:16px}

.hist-row b{display:block;font-family:var(--font-head);font-size:1.05rem;color:var(--ink);margin-bottom:.2rem}
.hist-row small{display:flex;align-items:center;gap:.3rem;color:var(--muted);font-size:.82rem;line-height:1.6}
.hist-row small .mi{font-size:15px;color:var(--lv-500)}
.hist-row .status-pill{margin-right:.3rem}
[class*="st-key-card-hist-"]{animation:fadeUp .6s var(--ease) both}
.st-key-card-danger{border-color:#f6d3df !important;background:linear-gradient(160deg,#fff 40%,#fff4f7) !important}
[data-testid="stPopoverButton"],[data-testid="stPopover"] button{border-radius:999px !important}
[data-testid="stBaseButton-tertiary"]{color:var(--lv-700) !important}

/* แถบข้อความวิ่ง (หน้าเข้าสู่ระบบ) */
.marquee{overflow:hidden;white-space:nowrap;margin:1.4rem 0 .4rem;mask-image:linear-gradient(90deg,transparent,#000 8%,#000 92%,transparent)}
.marquee .track{display:inline-block;animation:marquee 28s linear infinite}
.marquee span{font-family:var(--font-head);font-weight:700;font-size:clamp(1.4rem,3vw,2.3rem);color:transparent;
  -webkit-text-stroke:1.2px var(--lv-400);margin-right:1.4rem;text-transform:uppercase;letter-spacing:.02em}
.marquee span.fill{color:var(--lv-300);-webkit-text-stroke:0}

@keyframes fadeUp{from{opacity:0;transform:translateY(24px)}to{opacity:1;transform:none}}
@keyframes streak{from{transform:scaleX(0)}to{transform:scaleX(1)}}
@keyframes spin{to{transform:rotate(360deg)}}
@keyframes marquee{from{transform:translateX(0)}to{transform:translateX(-50%)}}
@keyframes morph{0%{transform:translate(0,0) scale(1)}50%{transform:translate(40px,25px) scale(1.15)}
  100%{transform:translate(-30px,10px) scale(.95)}}
@media (prefers-reduced-motion:reduce){*{animation:none !important;transition:none !important}}
@media (max-width:768px){.hero-inner{max-width:100%}.badge{display:none}.hero{padding:1.5rem}}
"""


def mi(name, cls=""):
    """ไอคอน Material Symbols สำหรับ HTML ที่เขียนเอง"""
    return f'<span class="mi {cls}" aria-hidden="true">{name}</span>'


def inject_theme():
    st.html(f"<style>{THEME_CSS}</style>")


def html_block(markup):
    st.markdown(markup, unsafe_allow_html=True)


def badge_svg(text="AI TIMETABLE • SMART SCHEDULER • ", icon="auto_awesome"):
    return ('<div class="badge" aria-hidden="true"><svg viewBox="0 0 120 120">'
            '<defs><path id="badgeCircle" d="M60,60 m-47,0 a47,47 0 1,1 94,0 a47,47 0 1,1 -94,0"/></defs>'
            f'<text><textPath href="#badgeCircle">{html_lib.escape(text)}</textPath></text></svg>'
            f'<div class="core">{mi(icon)}</div></div>')


def page_header(icon, title, desc=""):
    html_block(f'<div class="page-head"><div class="ic">{mi(icon)}</div><div><h2>{html_lib.escape(title)}</h2>'
               + (f"<p>{html_lib.escape(desc)}</p>" if desc else "") + "</div></div>")


def mini_card(icon, title, body="", tone="lv", extra="", cls=""):
    return (f'<div class="mini-card {cls}"><div class="ic tone-{tone}">{mi(icon)}</div>'
            f'<b>{html_lib.escape(title)}</b><small>{body}</small>{extra}</div>')


def style_status(df, cols):
    """ระบายสีพาสเทลให้คอลัมน์สถานะในตาราง (ผ่าน = เขียว, ไม่ผ่าน/ไม่มี = ชมพู, เติมให้ = ม่วง)"""
    def color(v):
        s = str(v)
        if any(k in s for k in ("ไม่ผ่าน", "อ่านไม่ได้", "ไม่มีข้อมูล", "จำเป็น")) or s == "ไม่มี":
            return "background-color:#ffe3ec;color:#a23a63;font-weight:600"
        if s.startswith(("ผ่าน", "ได้รับ", "เปลี่ยนแล้ว")):
            return "background-color:#e0f5ec;color:#2f7d5d;font-weight:600"
        if s.startswith(("ไม่มี", "ข้าม", "ไม่ทราบ", "รหัสเริ่มต้น")):
            return "background-color:#f3edff;color:#6f52c0"
        return ""
    cols = [c for c in cols if c in df.columns]
    return df.style.map(color, subset=cols) if cols and len(df) else df


# =============================================================================
# 8) หน้าเข้าสู่ระบบ
# =============================================================================
def page_login():
    st.html("<style>[data-testid='stSidebar'],[data-testid='stSidebarCollapsedControl'],"
            "[data-testid='stExpandSidebarButton']{display:none}</style>")
    pub = load_published()
    left, right = st.columns([1.35, 1], gap="large", vertical_alignment="center")
    with left:
        html_block(
            '<div class="hero big"><div class="blob b1"></div><div class="blob b2"></div><div class="blob b3"></div>'
            f'<div class="hero-inner"><span class="hero-tag">{mi("auto_awesome")} AI Timetable Scheduler</span>'
            '<h1>ระบบจัดตารางสอน<br><span class="accent">อัตโนมัติ</span><span class="streak"></span></h1>'
            '<p>อัปโหลดข้อมูลครั้งเดียว ระบบจัดตาราง ตรวจเงื่อนไข และส่งออกให้ครบในไม่กี่วินาที</p>'
            f'<div class="chips"><span class="chip">{mi("bolt")} จัดตารางทันที</span>'
            f'<span class="chip">{mi("verified")} ตรวจเงื่อนไขครบทุกข้อ</span>'
            f'<span class="chip">{mi("download")} Excel · PDF · CSV</span></div></div>'
            f'{badge_svg()}</div>')
    with right:
        with st.form("login_form"):
            html_block(f'<div class="page-head" style="margin:0 0 .6rem"><div class="ic">{mi("lock_open")}</div>'
                       '<div><h2>เข้าสู่ระบบ</h2><p>เมนูจะแสดงตามสิทธิ์ของผู้ใช้</p></div></div>')
            username = st.text_input("ชื่อผู้ใช้", placeholder="รหัสครู / รหัสกลุ่มเรียน / รหัสนักเรียน")
            password = st.text_input("รหัสผ่าน", type="password")
            ok = st.form_submit_button("เข้าสู่ระบบ", type="primary", width="stretch", icon=":material/login:")
            st.caption(":material/help: ลืมรหัสผ่าน? ติดต่อฝ่ายวิชาการเพื่อรีเซ็ตเป็นรหัสผ่านเริ่มต้น")
        if ok:
            do_login(username, password)
            if st.session_state.get("auth_user"):
                st.rerun()
        if st.session_state.get("login_error"):
            err = st.session_state["login_error"]
            st.error(err, icon=":material/lock_clock:" if "ระงับ" in err else ":material/error:")

    html_block('<div class="marquee" aria-hidden="true"><div class="track">'
               + "".join(f'<span class="{"fill" if i % 2 else ""}">{w}</span>'
                         for i, w in enumerate(["Smart Timetable", "จัดตารางอัตโนมัติ", "No Conflicts",
                                                "ตรวจเงื่อนไขครบ", "Export Anywhere", "ใช้งานง่าย"] * 2))
               + "</div></div>")
    st.markdown("#### :material/badge: ชื่อผู้ใช้ของแต่ละกลุ่ม")
    unames = {"admin": "<code>admin</code>", "teacher": "รหัสครู เช่น <code>T01</code> หรือชื่อ-นามสกุล",
              "student": "รหัสกลุ่มเรียน เช่น <code>G1</code> หรือรหัสนักเรียน"}
    for col, role in zip(st.columns(3), ROLE_INFO):
        with col:
            info = ROLE_INFO[role]
            html_block(mini_card(info["icon"], info["label"], f"ชื่อผู้ใช้: {unames[role]}<br>{info['desc']}",
                                 tone=info["tone"]))
    st.caption(":material/info: ทุกบัญชีเริ่มต้นด้วยรหัสผ่านเดียวกัน (สอบถามฝ่ายวิชาการ) และควรเปลี่ยนรหัสผ่านหลังเข้าใช้ครั้งแรก")
    if pub:
        st.success(f"ตารางเรียนฉบับปัจจุบันเผยแพร่โดย {pub['approved_by']} เมื่อ {pub['approved_at']}"
                   + (f" · {pub['note']}" if pub.get("note") else ""), icon=":material/campaign:")
    else:
        st.info("ยังไม่มีตารางเรียนที่เผยแพร่ — ครูและนักเรียนจะเข้าสู่ระบบได้หลังฝ่ายวิชาการเผยแพร่ตาราง",
                icon=":material/campaign:")


# =============================================================================
# 9) หน้านำเข้าข้อมูล & จัดตาราง (ผู้ดูแลระบบ)
# =============================================================================
SETTING_KEYS = {"rules": [], "max_block": DEFAULT_SETTINGS["max_block"],
                "attempts": DEFAULT_SETTINGS["attempts"], "seed": DEFAULT_SETTINGS["seed"]}


def current_settings():
    """ค่าตั้งการจัดตาราง เก็บใน cfg_* (ไม่หายเมื่อสลับหน้า ต่างจากค่าของ widget)"""
    ss = st.session_state
    cfg = {k: ss.get(f"cfg_{k}", d) for k, d in SETTING_KEYS.items()}
    return {"rules": sorted(cfg["rules"] or []), "max_block": int(cfg["max_block"]),
            "attempts": int(cfg["attempts"]), "seed": int(cfg["seed"])}


def sync_setting(k):
    st.session_state[f"cfg_{k}"] = st.session_state[f"w_{k}"]


def new_seed():
    seed = random.randint(1, 99999)
    st.session_state["cfg_seed"] = st.session_state["w_seed"] = seed


def run_schedule():
    ss = st.session_state
    settings = current_settings()
    t0 = time.perf_counter()
    with st.spinner(":material/smart_toy: กำลังจัดตารางเรียนและตรวจสอบเงื่อนไข..."):
        result = generate_timetable(ss["ds"], settings)
        checks = verify_schedule(result, ss["ds"])
    result["elapsed"] = time.perf_counter() - t0
    result["generated_at"] = now_text()
    ss["result"], ss["checks"] = result, checks
    ss["result_sig"] = (ss.get("ds_version"), json.dumps(settings, sort_keys=True))
    try:
        save_pickle(DRAFT_FILE, {"ds": ss["ds"], "file_report": ss.get("file_report"), "result": result,
                                 "checks": checks, "ds_version": ss.get("ds_version"),
                                 "history_id": ss.get("history_id")})
    except OSError:
        pass
    if ss.get("history_id"):
        try:
            save_history_result(ss["history_id"], result, checks)
        except Exception:
            pass
    log_event(ss.get("auth_user"), "จัดตาราง",
              f"{result['stats']['placed']}/{result['stats']['required']} คาบ · เงื่อนไขพิเศษ {settings['rules'] or '-'}")


def restore_draft():
    ss = st.session_state
    if ss.get("ds") is None and not ss.get("draft_checked"):
        ss["draft_checked"] = True
        draft = load_pickle(DRAFT_FILE)
        if draft:
            for k in ("ds", "file_report", "result", "checks", "ds_version", "history_id"):
                ss[k] = draft.get(k)
            if draft.get("result"):
                s = draft["result"]["settings"]
                for k in SETTING_KEYS:
                    ss[f"cfg_{k}"] = s[k]
                ss["result_sig"] = (ss["ds_version"], json.dumps(current_settings(), sort_keys=True))


def clear_data():
    ss = st.session_state
    for k in ("ds", "file_report", "result", "checks", "result_sig", "upload_sig", "history_id"):
        ss.pop(k, None)
    ss["uploader_v"] = ss.get("uploader_v", 0) + 1
    try:
        DRAFT_FILE.unlink(missing_ok=True)
    except OSError:
        pass


def render_completeness(ds):
    rows = []
    fallback = {
        "teacher": "สร้างจากรหัสครูในไฟล์ teach", "room": "จัดตารางโดยไม่ระบุห้อง",
        "student_group": "สร้างจากรหัสกลุ่มในไฟล์ register", "subject": "ไม่ทราบจำนวนคาบ (วิชาเหล่านั้นจัดไม่ได้)",
        "teach": "จัดโดยไม่มีครูผู้สอน", "register": "จำเป็น — ไม่มีข้อมูลนี้จัดตารางไม่ได้",
        "timeslot": "ใช้คาบมาตรฐาน จ.-ศ. วันละ 12 คาบ", "student": "นักเรียนใช้รหัสกลุ่มเรียนเข้าสู่ระบบ",
    }
    count = {"teacher": len(ds["teachers"]), "room": len(ds["rooms"]), "student_group": len(ds["groups"]),
             "subject": len(ds["subjects"]), "teach": len(ds["teach"]), "register": len(ds["register"]),
             "timeslot": len(ds["timeslots"]), "student": len(ds["students"])}
    for k, info in KINDS.items():
        got = ds["received"][k]
        rows.append({"ไฟล์": info["file"], "ข้อมูล": info["label"],
                     "สถานะ": "ได้รับ" if got else ("ไม่มี" if k == "register" else
                                                 "ไม่มี (ไม่บังคับ)" if k == "student" else "ไม่มี — ระบบเติมให้"),
                     "จำนวนรายการที่ใช้": count[k],
                     "ที่มา / การจัดการ": ", ".join(ds["sources"][k]) if got else fallback[k]})
    st.dataframe(style_status(pd.DataFrame(rows), ["สถานะ"]), hide_index=True, width="stretch")
    errs = [t for lv, t in ds["notes"] if lv == "error"]
    warns = [t for lv, t in ds["notes"] if lv == "warn"]
    infos = [t for lv, t in ds["notes"] if lv == "info"]
    if errs:
        st.error("**ต้องแก้ไข**\n\n" + "\n".join(f"- {t}" for t in errs), icon=":material/error:")
    if warns:
        st.warning("**ข้อมูลไม่ครบ/ไม่สอดคล้อง (ระบบจัดเท่าที่ทำได้)**\n\n" + "\n".join(f"- {t}" for t in warns[:25])
                   + (f"\n- ... และอีก {len(warns) - 25} รายการ" if len(warns) > 25 else ""), icon=":material/warning:")
    if infos:
        st.info("**ระบบเติมข้อมูลให้อัตโนมัติ**\n\n" + "\n".join(f"- {t}" for t in infos), icon=":material/auto_fix_high:")


def render_data_preview(ds):
    with st.expander(":material/search: ดูข้อมูลที่ระบบใช้จัดตาราง (หลังรวมไฟล์และเติมข้อมูลแล้ว)"):
        tabs = st.tabs(["ครู", "ห้อง", "กลุ่มเรียน", "รายวิชา", "teach", "register", "คาบเรียน", "นักเรียน"])
        subj = ds["subjects"].assign(
            ประเภทวิชา=lambda d: d.apply(lambda r: ("กิจกรรม" if r.is_activity else "")
                                          + (" สามัญ" if r.is_general else "") or "วิชาชีพ", axis=1))
        frames = [ds["teachers"], ds["rooms"], ds["groups"], subj.drop(columns=["is_general", "is_activity"]),
                  ds["teach"], ds["register"], ds["timeslots"], ds["students"]]
        for tab, df in zip(tabs, frames):
            with tab:
                st.dataframe(df, hide_index=True, width="stretch", height=min(38 * (len(df) + 1), 360))
        st.caption("room.category: theory = ห้องทฤษฎี · practice = ห้องปฏิบัติ · general = ระบุไม่ได้ (ใช้ได้ทั้งสองแบบ) — "
                   "ระบบดูจากคำใน room_type/room_name เช่น Theory, Lab, Practice, ปฏิบัติ")


def render_result(ds, result, checks):
    stats = result["stats"]
    hard_fail = checks[checks["ข้อ"].isin([2, 5, 6, 7]) & checks["ผล"].str.contains("ไม่")]
    fails = checks[checks["ผล"].str.contains("ไม่")]
    if fails.empty:
        st.success(f"จัดตารางสำเร็จครบ {stats['placed']} คาบ และผ่านเงื่อนไขทุกข้อ ({len(checks)} ข้อ)",
                   icon=":material/celebration:")
    elif hard_fail.empty:
        st.warning(f"จัดได้ {stats['placed']} จาก {stats['required']} คาบ — ไม่มีครู/ห้อง/กลุ่มเรียนชนกัน "
                   f"แต่ยังไม่ผ่าน {len(fails)} ข้อ (ดูสาเหตุด้านล่าง)", icon=":material/warning:")
    else:
        st.error("พบการละเมิดเงื่อนไขหลัก — กรุณาตรวจสอบข้อมูลและจัดตารางใหม่", icon=":material/error:")
    c = st.columns(5)
    c[0].metric(":material/target: คาบที่ต้องจัด", f"{stats['required']}")
    c[1].metric(":material/check_circle: จัดได้", f"{stats['placed']}",
                delta=f"{stats['placed'] / stats['required'] * 100:.0f}%" if stats["required"] else None)
    c[2].metric(":material/hourglass_top: ยังจัดไม่ได้", f"{stats['unplaced']}")
    c[3].metric(":material/rule: เงื่อนไขที่ผ่าน", f"{len(checks) - len(fails)}/{len(checks)}")
    c[4].metric(":material/timer: เวลาประมวลผล", f"{result.get('elapsed', 0):.2f} วิ")

    st.markdown("##### :material/verified_user: ผลตรวจเงื่อนไขการจัดตาราง")
    st.dataframe(style_status(checks, ["ผล"]), hide_index=True, width="stretch",
                 column_config={"ข้อ": st.column_config.NumberColumn(width="small"),
                                "เงื่อนไข": st.column_config.TextColumn(width="large")})
    if result.get("capacity"):
        st.error("**ข้อมูลเกินความจุของตาราง (จัดให้ครบไม่ได้ไม่ว่าจะจัดแบบใด)**\n\n"
                 + "\n".join(f"- {t}" for t in result["capacity"]), icon=":material/trending_down:")
    if len(result["unplaced"]):
        st.markdown("##### :material/pending_actions: รายวิชาที่ยังจัดไม่ครบ พร้อมสาเหตุ")
        st.dataframe(result["unplaced"], hide_index=True, width="stretch")
    if len(result["skipped"]):
        st.markdown("##### :material/block: รายการลงทะเบียนที่ข้าม (ข้อมูลไม่พอ)")
        st.dataframe(result["skipped"], hide_index=True, width="stretch")

    st.markdown("##### :material/download: ดาวน์โหลดผลลัพธ์")
    stamp = datetime.now(TZ).strftime("%Y%m%d_%H%M")
    sch = result["schedule"]
    d1, d2, d3 = st.columns(3)
    d1.download_button("output.csv (ตามรูปแบบที่กำหนด)", df_to_csv_bytes(output_df(sch)), "output.csv",
                       "text/csv", type="primary", width="stretch", icon=":material/description:")
    d2.download_button("output_detail.csv (มีวัน/คาบ)", df_to_csv_bytes(output_detail_df(sch)),
                       f"output_detail_{stamp}.csv", "text/csv", width="stretch", icon=":material/description:")
    if EXCEL_SUPPORTED:
        d3.download_button("Excel ตารางเรียนทุกกลุ่ม", build_excel(ds, sch, "group", list(ds["groups"]["group_id"]), checks),
                           f"timetable_groups_{stamp}.xlsx",
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", width="stretch",
                           icon=":material/table_view:")
    with st.expander(f":material/visibility: ตัวอย่าง output.csv ({len(sch)} แถว)"):
        st.dataframe(output_df(sch), hide_index=True, width="stretch", height=300)
    st.button("ไปดูตารางแบบกริด / ส่งออก Excel·PDF", on_click=go_to, args=(MENU_VIEW,), type="primary",
              icon=":material/arrow_forward:")


def page_data():
    ss = st.session_state
    restore_draft()
    page_header("upload_file", "นำเข้าข้อมูล & จัดตาราง",
                "อัปโหลดไฟล์ทั้งหมดพร้อมกันครั้งเดียว (CSV / Excel / ZIP) ระบบแยกประเภท เติมข้อมูลที่ขาด และจัดตารางให้ทันที")
    show_flash()

    with st.container(border=True, key="card-rules"):
        ss["cfg_rules"] = []   # ไม่ใช้เงื่อนไขพิเศษ (ข้อ 9-13) — จัดตามเงื่อนไขพื้นฐานข้อ 1-8 เท่านั้น
        for k in SETTING_KEYS:   # ตั้งค่า widget จากค่าที่บันทึกไว้ (กรณีกลับมาจากหน้าอื่น)
            if f"w_{k}" not in ss:
                ss[f"w_{k}"] = ss.get(f"cfg_{k}", SETTING_KEYS[k])
        with st.expander(":material/tune: ตั้งค่าขั้นสูง"):
            a, b, c = st.columns(3)
            a.slider("เรียนติดกันสูงสุด (คาบ/ครั้ง)", 1, MAX_PERIODS_PER_DAY, key="w_max_block",
                     on_change=sync_setting, args=("max_block",), help="วิชาที่มีคาบมากกว่านี้จะถูกแบ่งเป็นหลายครั้ง")
            b.slider("จำนวนรอบการค้นหา", 1, 200, key="w_attempts", on_change=sync_setting, args=("attempts",))
            c.number_input("ค่าสุ่มตั้งต้น (Seed)", 0, 99999, key="w_seed", on_change=sync_setting, args=("seed",),
                           help="ค่าเดิมได้ผลเดิม เปลี่ยนค่าเพื่อลองตารางแบบอื่น")

    with st.container(border=True, key="card-upload"):
        st.markdown("##### :material/looks_one: อัปโหลดไฟล์ข้อมูล")
        st.caption("teacher, room, student_group, subject, teach, register, timeslot "
                   "(ขาดบางไฟล์ได้ ระบบจะเติมให้เท่าที่ทำได้)")
        files = st.file_uploader("ลากไฟล์ทั้งหมดมาวางที่นี่", accept_multiple_files=True,
                                 type=["csv", "txt", "xlsx", "xlsm", "xls", "zip"],
                                 key=f"uploader_{ss.get('uploader_v', 0)}")
        if files:
            sig = tuple(sorted((f.name, f.size) for f in files))
            if sig != ss.get("upload_sig"):
                pairs = [(f.name, f.getvalue()) for f in files]
                raw, report = ingest_files(pairs)
                ss["ds"], ss["file_report"], ss["upload_sig"] = build_dataset(raw), report, sig
                ss["ds_version"] = ss.get("ds_version", 0) + 1
                try:   # เก็บชุดข้อมูลลงประวัติอัตโนมัติ
                    ss["history_id"] = archive_upload(pairs, ss["ds"], report)
                except Exception:
                    ss.pop("history_id", None)
                log_event(ss.get("auth_user"), "นำเข้าข้อมูล", f"{len(files)} ไฟล์")
        if ss.get("ds") is not None:
            a, b = st.columns([1, 1.4], vertical_alignment="center")
            a.button("ล้างข้อมูลที่นำเข้า", on_click=clear_data, icon=":material/delete:", width="stretch")
            b.button("ดูประวัติข้อมูลเก่า", on_click=go_to, args=(MENU_HIST,), icon=":material/inventory_2:",
                     type="tertiary")
            st.caption(":material/inventory_2: ชุดข้อมูลนี้ถูกเก็บในประวัติแล้ว ล้างออกแล้วนำกลับมาใช้ใหม่ได้")

    ds = ss.get("ds")
    if ds is None:
        st.info("ยังไม่มีข้อมูล — อัปโหลดไฟล์ด้านบนเพื่อเริ่มจัดตาราง", icon=":material/cloud_upload:")
        with st.expander(":material/list_alt: รูปแบบไฟล์ที่รองรับ"):
            st.dataframe(pd.DataFrame([{"ไฟล์": v["file"], "ข้อมูล": v["label"], "คอลัมน์": ", ".join(v["cols"])}
                                       for v in KINDS.values()]), hide_index=True, width="stretch")
            st.caption("ระบบแยกประเภทจากชื่อไฟล์/ชื่อชีตและหัวคอลัมน์ · Excel 1 ไฟล์หลายชีตได้ · ZIP ที่รวมหลายไฟล์ได้")
        return

    st.subheader(":material/looks_two: ผลการตรวจข้อมูล")
    if ss.get("file_report") is not None:
        st.dataframe(style_status(ss["file_report"], ["ประเภทที่ตรวจพบ"]), hide_index=True, width="stretch")
    render_completeness(ds)
    render_data_preview(ds)

    st.subheader(":material/looks_3: ผลการจัดตาราง")
    if not ds["can_schedule"]:
        st.error("ยังจัดตารางไม่ได้ — ต้องมีข้อมูลการลงทะเบียนเรียน (register.csv)", icon=":material/error:")
        return
    want_sig = (ss.get("ds_version"), json.dumps(current_settings(), sort_keys=True))
    if ss.get("result") is None or ss.get("result_sig") != want_sig:
        run_schedule()   # ข้อมูลหรือเงื่อนไขเปลี่ยน -> จัดตารางใหม่ทันที
    st.button("จัดตารางใหม่อีกครั้ง", on_click=new_seed, icon=":material/refresh:",
              help="สุ่มค่า Seed ใหม่เพื่อลองตารางแบบอื่น")
    render_result(ds, ss["result"], ss["checks"])


# =============================================================================
# 10) หน้าตารางเรียน/ตารางสอน & ส่งออก (ผู้ดูแลระบบ)
# =============================================================================
def render_grid_block(ds, sch, view, eid, search="", show_legend=True):
    st.html(f"<style>{GRID_CSS}</style>")
    st.markdown(f"##### :material/{VIEWS[view]['icon']}: {entity_title(ds, view, eid)}")
    st.markdown(grid_html(ds, sch, view, eid, search), unsafe_allow_html=True)
    if show_legend:
        lg = legend_df(ds, sch, view, eid)
        if len(lg):
            st.dataframe(lg, hide_index=True, width="stretch")
        else:
            st.caption("ไม่มีคาบในตาราง")


def render_downloads(ds, sch, view, targets, slug, checks=None, key=""):
    stamp = datetime.now(TZ).strftime("%Y%m%d_%H%M")
    sub = sch[sch[VIEWS[view]["col"]].isin(targets)] if not sch.empty else sch
    c = st.columns(4)
    if EXCEL_SUPPORTED:
        c[0].download_button("Excel", build_excel(ds, sch, view, targets, checks), f"{slug}_{stamp}.xlsx",
                             "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                             type="primary", width="stretch", key=f"xl_{key}", icon=":material/table_view:")
    pdf_ok = PDF_SUPPORTED and find_thai_font()[0]
    if pdf_ok:
        try:
            c[1].download_button("PDF", build_pdf(ds, sch, view, targets), f"{slug}_{stamp}.pdf",
                                 "application/pdf", width="stretch", key=f"pdf_{key}", icon=":material/picture_as_pdf:")
        except Exception as exc:
            c[1].button("PDF", disabled=True, width="stretch", help=str(exc), key=f"pdf_{key}", icon=":material/picture_as_pdf:")
    else:
        c[1].button("PDF", disabled=True, width="stretch", key=f"pdf_{key}", icon=":material/picture_as_pdf:",
                    help="ติดตั้ง fpdf2 (pip install fpdf2) และต้องมีฟอนต์ไทย — ใช้ปุ่ม HTML แล้วพิมพ์เป็น PDF แทนได้")
    c[2].download_button("HTML (พิมพ์เป็น PDF)", build_html_report(ds, sch, view, targets).encode("utf-8"),
                         f"{slug}_{stamp}.html", "text/html", width="stretch", key=f"html_{key}", icon=":material/print:")
    c[3].download_button("CSV", df_to_csv_bytes(output_detail_df(sub)), f"{slug}_{stamp}.csv", "text/csv",
                         width="stretch", key=f"csv_{key}", icon=":material/description:")


def render_stats(ds, sch):
    s = build_stats(ds, sch)
    pct = st.column_config.ProgressColumn("อัตราการใช้", min_value=0, max_value=100, format="%.0f%%")
    t1, t2, t3 = st.tabs([":material/co_present: ภาระงานสอนครู", ":material/meeting_room: การใช้ห้องเรียน",
                          ":material/groups: กลุ่มเรียน"])
    with t1:
        tdf = s["teacher_df"]
        if len(tdf):
            st.bar_chart(tdf, x="ชื่อครู", y="คาบสอน/สัปดาห์", horizontal=True, height=max(220, 30 * len(tdf)),
                         color="#a68aec")
        st.dataframe(tdf, hide_index=True, width="stretch")
    with t2:
        rdf = s["room_df"]
        st.dataframe(rdf, hide_index=True, width="stretch", column_config={"อัตราการใช้ (%)": pct})
        st.caption(f"คิดจากคาบเรียนที่ใช้ได้ทั้งหมด {s['slots']} คาบ/สัปดาห์ (ไม่รวมคาบพัก)")
    with t3:
        st.dataframe(s["group_df"], hide_index=True, width="stretch")


def page_view():
    ss = st.session_state
    restore_draft()
    page_header("calendar_view_week", "ตารางเรียน / ตารางสอน & ส่งออก",
                "ดูตารางแบบกริดรายกลุ่ม รายครู รายห้อง ค้นหาแบบไฮไลต์ และส่งออก Excel / PDF / CSV")
    result, ds = ss.get("result"), ss.get("ds")
    if result is None or ds is None:
        st.warning("ยังไม่มีตาราง — ไปที่หน้า 'นำเข้าข้อมูล & จัดตาราง' แล้วอัปโหลดไฟล์ก่อน", icon=":material/info:")
        st.button("ไปหน้านำเข้าข้อมูล", on_click=go_to, args=(MENU_DATA,), type="primary", icon=":material/arrow_forward:")
        return
    sch = result["schedule"]
    fails = ss["checks"][ss["checks"]["ผล"].str.contains("ไม่")]
    (st.success if fails.empty else st.warning)(
        f"ตารางฉบับร่าง จัดเมื่อ {result.get('generated_at', '-')} · {result['stats']['placed']}/"
        f"{result['stats']['required']} คาบ · ผ่านเงื่อนไข {len(ss['checks']) - len(fails)}/{len(ss['checks'])} ข้อ",
        icon=":material/task_alt:" if fails.empty else ":material/warning:")

    view = st.radio("มุมมอง", list(VIEWS), horizontal=True, key="view_mode",
                    format_func=lambda v: f":material/{VIEWS[v]['icon']}: ราย{VIEWS[v]['noun']}")
    options = entity_options(ds, sch, view)
    if not options:
        st.info(f"ไม่มีข้อมูล{VIEWS[view]['noun']}", icon=":material/info:")
        return
    c1, c2, c3 = st.columns([2, 2, 1], vertical_alignment="bottom")
    eid = c1.selectbox(f"เลือก{VIEWS[view]['noun']}", list(options), format_func=options.get, key=f"pick_{view}")
    search = c2.text_input("ไฮไลต์คำค้นหา", key="view_search", placeholder="เช่น 31901, T03, 6204")
    show_all = c3.toggle(f"แสดงทุก{VIEWS[view]['noun']}", key=f"all_{view}")
    targets = list(options) if show_all else [eid]
    for t in targets:
        render_grid_block(ds, sch, view, t, search, show_legend=not show_all)

    st.subheader(":material/ios_share: ส่งออก")
    scope = st.radio("ขอบเขต", ["one", "all"], horizontal=True, key="export_scope",
                     format_func=lambda x: f"เฉพาะที่เลือก ({options[eid]})" if x == "one"
                     else f"ทุก{VIEWS[view]['noun']} ({len(options)} รายการ)")
    exp_targets = [eid] if scope == "one" else list(options)
    render_downloads(ds, sch, view, exp_targets, f"timetable_{view}", ss["checks"], key="admin")
    st.caption("**Excel** = 1 ชีตต่อ 1 รายการ แบบตารางในโจทย์ + ตารางรายวิชา + ชีต output + ผลตรวจเงื่อนไข · "
               "**PDF** = A4 แนวนอน 1 หน้าต่อรายการ · **CSV** = output แบบมีวัน/คาบ")
    st.divider()
    st.subheader(":material/insights: สถิติช่วยตัดสินใจ")
    render_stats(ds, sch)


# =============================================================================
# 11) หน้าภาพรวม & เผยแพร่ / ผู้ใช้งาน (ผู้ดูแลระบบ)
# =============================================================================
def page_home():
    ss = st.session_state
    restore_draft()
    user = ss["auth_user"]
    page_header("space_dashboard", "ภาพรวม & เผยแพร่", "ติดตามขั้นตอนการทำงาน อนุมัติ และเผยแพร่ตารางให้ครูและนักเรียน")
    show_flash()
    pub = load_published()
    result = ss.get("result")

    st.subheader(":material/route: ขั้นตอนการทำงาน")
    steps = [("1. นำเข้าข้อมูล", "upload_file", "lv", ss.get("ds") is not None, "นำเข้าแล้ว", "ยังไม่นำเข้า", MENU_DATA),
             ("2. จัดตาราง", "auto_awesome", "pink", result is not None, "มีตารางฉบับร่าง", "ยังไม่จัด", MENU_DATA),
             ("3. ตรวจสอบ/ส่งออก", "fact_check", "sky", result is not None, "ดูตาราง Excel/PDF", "รอขั้นตอนที่ 2", MENU_VIEW),
             ("4. เผยแพร่", "campaign", "mint", pub is not None, "ครูและนักเรียนเห็นแล้ว", "ยังไม่เผยแพร่", None)]
    for col, (name, icon, tone, ok, yes, no, menu) in zip(st.columns(4), steps):
        with col:
            pill = (f'<span class="status-pill ok">{mi("check_circle")} เสร็จแล้ว</span>' if ok
                    else f'<span class="status-pill wait">{mi("radio_button_unchecked")} รอดำเนินการ</span>')
            html_block(mini_card(icon, name, yes if ok else no, tone=tone, extra=pill, cls="done" if ok else ""))
            if menu:
                st.button("ไปหน้านี้", key=f"step_{name}", on_click=go_to, args=(menu,), width="stretch",
                          icon=":material/arrow_forward:")

    st.subheader(":material/approval: อนุมัติและเผยแพร่ตารางฉบับร่าง")
    if result is None:
        st.info("ยังไม่มีตารางฉบับร่าง — นำเข้าข้อมูลเพื่อให้ระบบจัดตารางก่อน", icon=":material/info:")
    else:
        checks = ss["checks"]
        fails = checks[checks["ผล"].str.contains("ไม่")]
        hard = checks[checks["ข้อ"].isin([2, 5, 6, 7]) & checks["ผล"].str.contains("ไม่")]
        st.markdown(f"ฉบับร่างจัดเมื่อ **{result.get('generated_at', '-')}** · {result['stats']['placed']}/"
                    f"{result['stats']['required']} คาบ · ผ่านเงื่อนไข {len(checks) - len(fails)}/{len(checks)} ข้อ")
        if len(fails):
            st.dataframe(style_status(fails, ["ผล"]), hide_index=True, width="stretch")
        note = st.text_input("หมายเหตุ (ถ้ามี)", placeholder="เช่น ตารางเรียนภาคเรียนที่ 2/2569", key="pub_note")
        confirm = True
        if len(fails) and hard.empty:
            confirm = st.checkbox("ยืนยันเผยแพร่แม้ยังไม่ผ่านบางเงื่อนไข", key="pub_confirm")
        if not hard.empty:
            st.error("เผยแพร่ไม่ได้: มีครู/ห้อง/กลุ่มเรียนชนกัน หรือจัดลงคาบพัก", icon=":material/block:")
        if st.button("อนุมัติและเผยแพร่ให้ครูและนักเรียน", type="primary", disabled=not hard.empty or not confirm,
                     icon=":material/campaign:"):
            save_pickle(PUBLISHED_FILE, {"ds": ss["ds"], "result": result, "checks": checks,
                                         "approved_by": user["name"], "approved_at": now_text(), "note": note.strip()})
            log_event(user, "เผยแพร่ตาราง", note.strip())
            ss["flash"] = "เผยแพร่เรียบร้อย — ครูและนักเรียนเข้าสู่ระบบดูตารางได้แล้ว"
            st.rerun()

    st.subheader(":material/visibility: ตารางที่ครูและนักเรียนเห็นอยู่ตอนนี้")
    if pub:
        st.success(f"เผยแพร่โดย {pub['approved_by']} เมื่อ {pub['approved_at']}" + (f" · {pub['note']}" if pub.get("note") else ""),
                   icon=":material/campaign:")
        ps = pub["result"]["stats"]
        c = st.columns(4)
        c[0].metric(":material/co_present: ครู", len(pub["ds"]["teachers"]))
        c[1].metric(":material/groups: กลุ่มเรียน", len(pub["ds"]["groups"]))
        c[2].metric(":material/meeting_room: ห้อง", len(pub["ds"]["rooms"]))
        c[3].metric(":material/check_circle: คาบที่จัด", f"{ps['placed']}/{ps['required']}")
        if st.button("ยกเลิกการเผยแพร่", icon=":material/unpublished:"):
            PUBLISHED_FILE.unlink(missing_ok=True)
            log_event(user, "ยกเลิกการเผยแพร่")
            ss["flash"] = "ยกเลิกการเผยแพร่แล้ว — ครูและนักเรียนจะเข้าสู่ระบบไม่ได้จนกว่าจะเผยแพร่ใหม่"
            st.rerun()
    else:
        st.info("ยังไม่มีตารางที่เผยแพร่", icon=":material/info:")

    st.subheader(":material/history: ประวัติการใช้งาน")
    logs = read_json(LOG_FILE, [])
    if logs:
        st.dataframe(pd.DataFrame(logs[::-1]), hide_index=True, width="stretch", height=260)
    else:
        st.caption("ยังไม่มีประวัติ")


def page_users():
    ss = st.session_state
    page_header("manage_accounts", "ผู้ใช้งาน & รหัสผ่าน", "ดูบัญชีผู้ใช้ทั้งหมด และรีเซ็ตรหัสผ่านกลับเป็นค่าเริ่มต้น")
    show_flash()
    pub = load_published()
    st.info(f"รหัสผ่านเริ่มต้นของทุกบัญชี: **`{default_password()}`** · ผู้ใช้เปลี่ยนรหัสผ่านเองได้ที่แถบด้านซ้าย · "
            "เปลี่ยนค่าเริ่มต้นได้ด้วย `DEFAULT_PASSWORD` ใน .streamlit/secrets.toml", icon=":material/key:")
    if not pub:
        st.warning("ยังไม่มีตารางที่เผยแพร่ — บัญชีครูและนักเรียนจะถูกสร้างจากข้อมูลในตารางที่เผยแพร่", icon=":material/warning:")
    store = read_json(PASSWORD_FILE, {})
    accounts = account_directory(pub)
    df = pd.DataFrame([{
        "ชื่อผู้ใช้": a["username"], "สิทธิ์": ROLE_INFO[a["role"]]["label"], "ชื่อ": a["name"],
        "ผูกกับ": a["link"] or "ทั้งระบบ",
        "รหัสผ่าน": "เปลี่ยนแล้ว (" + store[a["key"]]["changed_at"] + ")" if a["key"] in store else "รหัสเริ่มต้น",
        "_key": a["key"]} for a in accounts])
    c = st.columns(4)
    c[0].metric(":material/badge: บัญชีทั้งหมด", len(df))
    c[1].metric(":material/co_present: ครู", int((df["สิทธิ์"] == ROLE_INFO["teacher"]["label"]).sum()))
    c[2].metric(":material/school: นักเรียน", int((df["สิทธิ์"] == ROLE_INFO["student"]["label"]).sum()))
    c[3].metric(":material/key: เปลี่ยนรหัสแล้ว", int(df["_key"].isin(store).sum()))
    role_filter = st.segmented_control("กรอง", ["ทั้งหมด", "ครูผู้สอน", "นักเรียน / นักศึกษา"], default="ทั้งหมด",
                                       key="user_filter")
    shown = df if role_filter in (None, "ทั้งหมด") else df[df["สิทธิ์"] == role_filter]
    st.dataframe(style_status(shown.drop(columns="_key"), ["รหัสผ่าน"]), hide_index=True, width="stretch", height=min(38 * (len(shown) + 1), 420))

    st.markdown("##### :material/lock_reset: รีเซ็ตรหัสผ่านกลับเป็นค่าเริ่มต้น")
    label = dict(zip(df["_key"], df["ชื่อผู้ใช้"] + " · " + df["ชื่อ"]))
    pick = st.multiselect("เลือกบัญชี", list(label), format_func=label.get, key="reset_pick")
    a, b = st.columns(2)
    if a.button("รีเซ็ตบัญชีที่เลือก", disabled=not pick, width="stretch", icon=":material/lock_reset:"):
        reset_passwords(pick)
        log_event(ss["auth_user"], "รีเซ็ตรหัสผ่าน", ", ".join(label[k].split(" · ")[0] for k in pick))
        ss["flash"] = f"รีเซ็ตรหัสผ่าน {len(pick)} บัญชีเป็นค่าเริ่มต้นแล้ว"
        st.rerun()
    confirm_all = b.checkbox("ยืนยันรีเซ็ตทุกบัญชี (ยกเว้น admin)", key="reset_all_ok")
    if b.button("รีเซ็ตทุกบัญชี", disabled=not confirm_all, width="stretch", icon=":material/warning:"):
        reset_passwords([k for k in df["_key"] if k != "admin"])
        log_event(ss["auth_user"], "รีเซ็ตรหัสผ่านทุกบัญชี")
        ss["flash"] = "รีเซ็ตรหัสผ่านทุกบัญชี (ยกเว้น admin) แล้ว"
        st.rerun()


def restore_history(hid):
    """นำชุดข้อมูลจากประวัติกลับมาเป็นข้อมูลปัจจุบัน (พร้อมตารางล่าสุดของชุดนั้น ถ้ามี)"""
    ss = st.session_state
    data = load_pickle(history_path(hid))
    if not data or data.get("ds") is None:
        ss["flash"] = ("error", "เปิดชุดข้อมูลนี้ไม่ได้ (ไฟล์อาจเสียหายหรือถูกลบไปแล้ว)")
        return
    for k in ("result", "checks", "result_sig", "upload_sig"):
        ss.pop(k, None)
    ss["ds"], ss["file_report"] = data["ds"], data.get("file_report")
    ss["ds_version"] = ss.get("ds_version", 0) + 1
    ss["uploader_v"] = ss.get("uploader_v", 0) + 1
    ss["history_id"], ss["draft_checked"] = hid, True
    if data.get("result") is not None:
        s = data["result"]["settings"]
        for k in ("max_block", "attempts", "seed"):
            ss[f"cfg_{k}"] = s[k]
            ss.pop(f"w_{k}", None)
        ss["cfg_rules"] = []
        ss["result"], ss["checks"] = data["result"], data["checks"]
        ss["result_sig"] = (ss["ds_version"], json.dumps(current_settings(), sort_keys=True))
    try:
        save_pickle(DRAFT_FILE, {"ds": ss["ds"], "file_report": ss.get("file_report"), "result": ss.get("result"),
                                 "checks": ss.get("checks"), "ds_version": ss["ds_version"], "history_id": hid})
    except OSError:
        pass
    name = next((it["name"] for it in history_index() if it["id"] == hid), hid)
    log_event(ss.get("auth_user"), "นำข้อมูลเก่ากลับมาใช้", name)
    ss["flash"] = ("success", f"นำ “{name}” กลับมาใช้แล้ว — ดูผลได้ที่หน้านำเข้าข้อมูล & จัดตาราง")
    ss["main_menu"] = MENU_DATA


def rename_history(hid):
    new = str(st.session_state.get(f"rn_{hid}", "")).strip()
    if not new:
        return
    items = history_index()
    for it in items:
        if it["id"] == hid:
            it["name"] = new[:80]
    write_history_index(items)
    st.session_state["flash"] = ("success", "เปลี่ยนชื่อชุดข้อมูลแล้ว")


def delete_history_cb(hids, label):
    delete_history(hids)
    log_event(st.session_state.get("auth_user"), "ลบประวัติข้อมูล", label)
    st.session_state["flash"] = ("success", f"ลบ{label}แล้ว")


def delete_current_cb():
    clear_data()
    st.session_state.pop("history_id", None)
    log_event(st.session_state.get("auth_user"), "ลบข้อมูลปัจจุบัน")
    st.session_state["flash"] = ("success", "ลบข้อมูลปัจจุบันแล้ว — ชุดข้อมูลยังเก็บอยู่ในประวัติ นำกลับมาใช้ได้ทุกเมื่อ")


def clear_log_cb():
    write_json(LOG_FILE, [])
    st.session_state["flash"] = ("success", "ล้างประวัติการใช้งานแล้ว")


def show_flash():
    f = st.session_state.pop("flash", None)
    if isinstance(f, tuple):
        (st.error if f[0] == "error" else st.success)(
            f[1], icon=":material/error:" if f[0] == "error" else ":material/check_circle:")
    elif f:
        st.success(f, icon=":material/check_circle:")


def page_history():
    ss = st.session_state
    restore_draft()
    page_header("inventory_2", "ประวัติข้อมูล & ลบข้อมูล",
                "ทุกชุดข้อมูลที่นำเข้าถูกเก็บไว้อัตโนมัติ เลือกนำกลับมาใช้ ดาวน์โหลด เปลี่ยนชื่อ หรือลบได้")
    show_flash()
    items = sorted(history_index(), key=lambda it: it["id"], reverse=True)
    cur = ss.get("history_id")

    # ---- ข้อมูลที่ใช้อยู่ตอนนี้ ----
    st.subheader(":material/folder_open: ข้อมูลที่ใช้อยู่ตอนนี้")
    with st.container(border=True, key="card-current"):
        if ss.get("ds") is None:
            st.caption(":material/info: ยังไม่มีข้อมูลที่ใช้อยู่ — นำเข้าไฟล์ใหม่ หรือเลือกชุดข้อมูลเก่าจากประวัติด้านล่าง")
            st.button("ไปหน้านำเข้าข้อมูล", on_click=go_to, args=(MENU_DATA,), icon=":material/upload_file:")
        else:
            name = next((it["name"] for it in items if it["id"] == cur), "ข้อมูลที่นำเข้าล่าสุด")
            res = ss.get("result")
            c = ds_counts(ss["ds"])
            st.markdown(f"**{name}**")
            st.caption(f"ครู {c['teachers']} · กลุ่มเรียน {c['groups']} · รายวิชา {c['subjects']} · ห้อง {c['rooms']}"
                       + (f" · จัดได้ {res['stats']['placed']}/{res['stats']['required']} คาบ" if res else ""))
            a, b, _ = st.columns([1, 1, 2])
            a.button("ไปหน้าจัดตาราง", on_click=go_to, args=(MENU_DATA,), icon=":material/arrow_forward:",
                     width="stretch")
            with b.popover("ลบข้อมูลปัจจุบัน", icon=":material/delete:", width="stretch"):
                st.markdown("ลบข้อมูลที่นำเข้าและตารางฉบับร่างออกจากหน้าจัดตาราง")
                st.caption("ไม่กระทบตารางที่เผยแพร่แล้ว · ชุดข้อมูลยังอยู่ในประวัติ นำกลับมาใช้ได้")
                st.button("ยืนยันลบ", type="primary", on_click=delete_current_cb, key="del_current",
                          icon=":material/delete_forever:", width="stretch")

    # ---- ประวัติชุดข้อมูล ----
    st.subheader(":material/history: ชุดข้อมูลที่เคยนำเข้า")
    if not items:
        st.info("ยังไม่มีประวัติ — ทุกครั้งที่อัปโหลดไฟล์ในหน้านำเข้าข้อมูล ระบบจะเก็บชุดข้อมูลไว้ที่นี่อัตโนมัติ",
                icon=":material/inventory_2:")
    else:
        q = st.text_input("ค้นหาชุดข้อมูล", placeholder="พิมพ์ชื่อชุดข้อมูล หรือชื่อไฟล์", key="hist_q")
        if q.strip():
            ql = q.strip().lower()
            items = [it for it in items if ql in it["name"].lower() or any(ql in f.lower() for f in it["files"])]
        st.caption(f"ทั้งหมด {len(items)} ชุด · ไฟล์ชุดเดิมที่อัปโหลดซ้ำจะอัปเดตรายการเดิม ไม่สร้างซ้ำ")
        for it in items:
            hid, cnt = it["id"], it.get("counts", {})
            with st.container(border=True, key=f"card-hist-{hid}"):
                info, act = st.columns([2.2, 1.8], vertical_alignment="center")
                with info:
                    pills = ""
                    if hid == cur and ss.get("ds") is not None:
                        pills += f'<span class="status-pill ok">{mi("radio_button_checked")} กำลังใช้อยู่</span> '
                    if it.get("placed") is not None:
                        pills += (f'<span class="status-pill wait">{mi("event_available")} จัดได้ '
                                  f'{it["placed"]}/{it["required"]} คาบ</span>')
                    else:
                        pills += f'<span class="status-pill wait">{mi("schedule")} ยังไม่ได้จัดตาราง</span>'
                    files = ", ".join(it["files"][:8]) + (f" และอีก {len(it['files']) - 8} ไฟล์" if len(it["files"]) > 8 else "")
                    html_block(f'<div class="hist-row"><b>{html_lib.escape(it["name"])}</b>'
                               f'<small>{mi("schedule")} นำเข้าเมื่อ {it["created_at"]} · ใช้ล่าสุด {it["updated_at"]}</small>'
                               f'<small>{mi("description")} {html_lib.escape(files)}</small>'
                               f'<small>{mi("groups")} ครู {cnt.get("teachers", "-")} · กลุ่มเรียน {cnt.get("groups", "-")}'
                               f' · รายวิชา {cnt.get("subjects", "-")} · ห้อง {cnt.get("rooms", "-")}</small>'
                               f'<div>{pills}</div></div>')
                with act:
                    r1, r2 = st.columns(2)
                    r1.button("นำกลับมาใช้", key=f"use_{hid}", on_click=restore_history, args=(hid,), type="primary",
                              icon=":material/settings_backup_restore:", width="stretch")
                    r2.download_button("ดาวน์โหลด", data=lambda h=hid: history_zip(h), file_name=f"{hid}_data.zip",
                                       mime="application/zip", key=f"dl_{hid}", icon=":material/download:",
                                       width="stretch", on_click="ignore")
                    r3, r4 = st.columns(2)
                    with r3.popover("เปลี่ยนชื่อ", icon=":material/edit:", width="stretch"):
                        st.text_input("ชื่อใหม่", value=it["name"], key=f"rn_{hid}", max_chars=80)
                        st.button("บันทึกชื่อ", key=f"rnb_{hid}", on_click=rename_history, args=(hid,),
                                  type="primary", icon=":material/save:", width="stretch")
                    with r4.popover("ลบ", icon=":material/delete:", width="stretch"):
                        st.markdown(f"ลบ **{it['name']}** ถาวร?")
                        st.caption("ลบแล้วกู้คืนไม่ได้ · ไม่กระทบข้อมูลที่ใช้อยู่และตารางที่เผยแพร่")
                        st.button("ยืนยันลบ", key=f"del_{hid}", type="primary", on_click=delete_history_cb,
                                  args=([hid], f" “{it['name']}” "), icon=":material/delete_forever:", width="stretch")

    # ---- ลบข้อมูลทั้งหมด ----
    st.subheader(":material/delete_sweep: ลบข้อมูลทั้งหมด")
    with st.container(border=True, key="card-danger"):
        a, b = st.columns(2)
        with a:
            st.markdown("**ลบประวัติชุดข้อมูลทั้งหมด**")
            st.caption(f"ลบชุดข้อมูลเก่าทั้ง {len(history_index())} ชุดถาวร (ข้อมูลที่ใช้อยู่และตารางที่เผยแพร่ไม่หาย)")
            ok_all = st.checkbox("ยืนยัน ลบประวัติทั้งหมด", key="del_all_ok")
            st.button("ลบประวัติทั้งหมด", disabled=not ok_all or not history_index(), on_click=delete_history_cb,
                      args=([it["id"] for it in history_index()], "ประวัติชุดข้อมูลทั้งหมด"),
                      icon=":material/delete_forever:", key="del_all")
        with b:
            st.markdown("**ล้างประวัติการใช้งาน (log)**")
            st.caption(f"ลบบันทึกการเข้าสู่ระบบและกิจกรรมทั้งหมด {len(read_json(LOG_FILE, []))} รายการ")
            ok_log = st.checkbox("ยืนยัน ล้างประวัติการใช้งาน", key="del_log_ok")
            st.button("ล้างประวัติการใช้งาน", disabled=not ok_log, on_click=clear_log_cb,
                      icon=":material/delete_forever:", key="del_log")
        st.caption(":material/info: ตารางที่เผยแพร่ให้ครูและนักเรียน ยกเลิกได้ที่หน้า ภาพรวม & เผยแพร่")


# =============================================================================
# 12) หน้าของครูและนักเรียน (อ่านจากตารางที่เผยแพร่)
# =============================================================================
def published_or_warn():
    pub = load_published()
    if not pub:
        st.warning("ยังไม่มีตารางที่เผยแพร่ — กรุณาติดต่อฝ่ายวิชาการ", icon=":material/hourglass_empty:")
        return None
    st.caption(f":material/campaign: ตารางฉบับเผยแพร่โดย {pub['approved_by']} เมื่อ {pub['approved_at']}"
               + (f" · {pub['note']}" if pub.get("note") else ""))
    return pub


def render_today(ds, sch, view, eid):
    now = datetime.now(TZ)
    today = DAY_CODES[now.weekday()]
    word = "สอน" if view == "teacher" else "เรียน"
    sub = sch[sch[VIEWS[view]["col"]] == eid]
    days = [d for d in DAY_CODES if d in set(ds["timeslots"]["day"])]
    target = today if today in days else next((DAY_CODES[(now.weekday() + k) % 7] for k in range(1, 8)
                                               if DAY_CODES[(now.weekday() + k) % 7] in days), None)
    if target is None:
        return
    st.markdown(f"#### :material/today: {'วันนี้ ' if target == today else 'วัน' + word + 'ถัดไป: '}วัน{DAY_TH[target]}")
    todays = sub[sub["day"] == target].drop_duplicates(["period"]).sort_values("period")
    if todays.empty:
        st.info(f"วัน{DAY_TH[target]}ไม่มีคาบ{word} พักผ่อนให้เต็มที่นะ", icon=":material/beach_access:")
        return
    lk = lookups(ds)
    ts = ds["timeslots"].set_index(["day", "period"])
    blocks, cur = [], None
    for r in todays.itertuples(index=False):
        if cur and cur["sid"] == r.subject_id and cur["room"] == r.room_id and \
                all(q in BREAK_PERIODS for q in range(cur["end"] + 1, r.period)):
            cur["end"] = r.period
        else:
            cur = {"sid": r.subject_id, "room": r.room_id, "teacher": r.teacher_id, "start": r.period, "end": r.period}
            blocks.append(cur)
    cols = st.columns(min(len(blocks), 4))
    esc = html_lib.escape
    for i, b in enumerate(blocks):
        t0 = ts.loc[(target, b["start"]), "start"] if (target, b["start"]) in ts.index else ""
        t1 = ts.loc[(target, b["end"]), "end"] if (target, b["end"]) in ts.index else ""
        grp = "+".join(sorted(sub[(sub["day"] == target) & (sub["period"] == b["start"])]["group_id"].unique()))
        period = f"คาบ {b['start']}" + (f"-{b['end']}" if b["end"] > b["start"] else "")
        if view == "teacher":
            other = f'<span>{mi("groups")} กลุ่ม {esc(grp)}</span>'
        else:
            other = f'<span>{mi("co_present")} {esc(str(lk["teacher"].get(b["teacher"], b["teacher"] or "-")))}</span>'
        with cols[i % len(cols)]:
            html_block(f'<div class="mini-card class-card"><div class="time">{mi("schedule")} {esc(str(t0))}–{esc(str(t1))}'
                       f' <small style="display:inline;margin-left:auto">{period}</small></div>'
                       f'<div class="subj"><b style="display:inline">{esc(str(b["sid"]))}</b> '
                       f'{esc(str(lk["subject"].get(b["sid"], "")))}</div>'
                       f'<div class="meta">{other}<span>{mi("meeting_room")} ห้อง {esc(str(b["room"] or "-"))}</span></div></div>')


def page_teacher_tt():
    user = st.session_state["auth_user"]
    page_header("event_note", "ตารางสอนของฉัน", "ตารางสอนประจำสัปดาห์และคาบสอนวันนี้")
    pub = published_or_warn()
    if not pub:
        return
    ds, sch = pub["ds"], pub["result"]["schedule"]
    tid = user["link"]
    sub = sch[sch["teacher_id"] == tid]
    uniq = sub.drop_duplicates(["day", "period"])
    c = st.columns(4)
    c[0].metric(":material/timer: คาบสอน/สัปดาห์", f"{len(uniq)} คาบ")
    c[1].metric(":material/menu_book: รายวิชา", f"{sub['subject_id'].nunique()} วิชา")
    c[2].metric(":material/groups: กลุ่มที่สอน", f"{sub['group_id'].nunique()} กลุ่ม")
    c[3].metric(":material/calendar_month: วันที่มาสอน", f"{uniq['day'].nunique()} วัน")
    render_today(ds, sch, "teacher", tid)
    render_grid_block(ds, sch, "teacher", tid)
    st.markdown("#### :material/download: ดาวน์โหลดตารางสอนของฉัน")
    render_downloads(ds, sch, "teacher", [tid], f"teacher_{tid}", key="t")


def page_teacher_load():
    user = st.session_state["auth_user"]
    page_header("assignment_ind", "ภาระงานสอนของฉัน", "สรุปคาบสอน รายวิชา และการกระจายคาบสอนรายวัน")
    pub = published_or_warn()
    if not pub:
        return
    ds, sch = pub["ds"], pub["result"]["schedule"]
    tid = user["link"]
    s = build_stats(ds, sch)
    tdf = s["teacher_df"]
    me = tdf[tdf["รหัสครู"] == tid]
    if me.empty:
        st.warning("ไม่พบข้อมูลของท่านในตารางที่เผยแพร่", icon=":material/person_search:")
        return
    me = me.iloc[0]
    active = tdf[tdf["คาบสอน/สัปดาห์"] > 0]["คาบสอน/สัปดาห์"]
    avg = active.mean() if len(active) else 0
    c = st.columns(3)
    c[0].metric(":material/timer: คาบสอน/สัปดาห์", f"{me['คาบสอน/สัปดาห์']} คาบ",
                delta=f"{me['คาบสอน/สัปดาห์'] - avg:+.1f} เทียบค่าเฉลี่ย", delta_color="off")
    c[1].metric(":material/workspace_premium: ตำแหน่ง", me["ตำแหน่ง"])
    c[2].metric(":material/calendar_month: วันที่มาสอน", f"{me['วันที่มาสอน']} วัน")
    rules = set(pub["result"]["settings"]["rules"])
    if 10 in rules and me["ตำแหน่ง"].startswith("หัวหน้า"):
        st.info(f"ท่านมีประชุมหัวหน้าประจำสัปดาห์ วัน{DAY_TH[LEADER_MEETING[0]]} คาบ "
                f"{', '.join(map(str, LEADER_MEETING[1]))} (15.00-16.00 น.) ระบบไม่จัดคาบสอนในช่วงนี้",
                icon=":material/forum:")
    st.markdown("#### :material/menu_book: รายวิชาที่สอน")
    st.dataframe(legend_df(ds, sch, "teacher", tid), hide_index=True, width="stretch")
    sub = sch[sch["teacher_id"] == tid].drop_duplicates(["day", "period"])
    st.markdown("#### :material/bar_chart: จำนวนคาบสอนรายวัน")
    st.bar_chart(pd.DataFrame({"วัน": [DAY_TH[d] for d in s["days"]],
                               "คาบสอน": [int((sub["day"] == d).sum()) for d in s["days"]]}),
                 x="วัน", y="คาบสอน", height=260, sort=False, color="#a68aec")


def page_student_tt():
    user = st.session_state["auth_user"]
    page_header("event_note", "ตารางเรียนของฉัน", "ตารางเรียนประจำสัปดาห์และคาบเรียนวันนี้")
    pub = published_or_warn()
    if not pub:
        return
    ds, sch = pub["ds"], pub["result"]["schedule"]
    gid = user["link"]
    if gid not in set(ds["groups"]["group_id"]):
        st.warning(f"ไม่พบกลุ่มเรียน {gid} ในตารางที่เผยแพร่", icon=":material/person_search:")
        return
    sub = sch[sch["group_id"] == gid]
    c = st.columns(4)
    c[0].metric(":material/menu_book: รายวิชา", f"{sub.loc[sub['kind'] != 'homeroom', 'subject_id'].nunique()} วิชา")
    c[1].metric(":material/timer: คาบเรียน/สัปดาห์", f"{len(sub)} คาบ")
    c[2].metric(":material/calendar_month: วันที่มาเรียน", f"{sub['day'].nunique()} วัน")
    c[3].metric(":material/co_present: ครูผู้สอน", f"{sub.loc[sub['teacher_id'] != '', 'teacher_id'].nunique()} คน")
    render_today(ds, sch, "group", gid)
    render_grid_block(ds, sch, "group", gid)
    st.markdown("#### :material/download: ดาวน์โหลดตารางเรียน")
    render_downloads(ds, sch, "group", [gid], f"group_{gid}", key="s")


def page_student_subj():
    user = st.session_state["auth_user"]
    page_header("auto_stories", "รายวิชาและครูผู้สอน", "รายวิชาที่ลงทะเบียน จำนวนคาบ และครูผู้สอนของกลุ่มเรียน")
    pub = published_or_warn()
    if not pub:
        return
    ds, sch = pub["ds"], pub["result"]["schedule"]
    gid = user["link"]
    lk = lookups(ds)
    st.subheader(f":material/groups: กลุ่มเรียน {lk['group'].get(gid, gid)}")
    df = legend_df(ds, sch, "group", gid)
    st.dataframe(df, hide_index=True, width="stretch")
    if len(df):
        st.caption(f"รวม {len(df)} รายการ · {int(df['คาบ/สัปดาห์'].sum())} คาบ/สัปดาห์")
    g = ds["groups"].set_index("group_id")
    if gid in g.index and (g.at[gid, "advisor"] or g.at[gid, "advisor_id"]):
        adv = g.at[gid, "advisor_id"]
        st.info(f"ครูที่ปรึกษา: {lk['teacher'].get(adv, g.at[gid, 'advisor'])}", icon=":material/supervisor_account:")


# =============================================================================
# 13) โครงหลักของแอป
# =============================================================================
def render_hero(user, menu):
    hour = datetime.now(TZ).hour
    greet = "สวัสดีตอนเช้า" if hour < 12 else "สวัสดีตอนบ่าย" if hour < 17 else "สวัสดีตอนเย็น"
    info = ROLE_INFO[user["role"]]
    html_block('<div class="hero"><div class="blob b1"></div><div class="blob b2"></div><div class="blob b3"></div>'
               f'<div class="hero-inner"><span class="hero-tag">{mi("auto_awesome")} AI Timetable Scheduler</span>'
               f'<h1>{greet}, <span class="accent">{html_lib.escape(str(user["name"]))}</span>'
               '<span class="streak"></span></h1>'
               f'<p>{mi(info["icon"])} {html_lib.escape(info["label"])} · {datetime.now(TZ).strftime("%d/%m/%Y")}</p>'
               f'</div>{badge_svg(icon=MENU_ICONS.get(menu, "calendar_month"))}</div>')


def render_sidebar(user):
    with st.sidebar:
        html_block(f'<div class="side-brand"><div class="logo">{mi("calendar_month")}</div>'
                   '<div><b>AI Timetable</b><small>ระบบจัดตารางสอนอัตโนมัติ</small></div></div>')
        info = ROLE_INFO[user["role"]]
        with st.container(border=True, key="card-user"):
            sub = (f"{info['label']} · " if user["name"] != info["label"] else "") + f"ชื่อผู้ใช้ {user['username']}"
            html_block(f'<div class="user-card"><div class="av">{mi(info["icon"])}</div><div>'
                       f'<b>{html_lib.escape(str(user["name"]))}</b><small>{html_lib.escape(sub)}</small></div></div>')
            st.button("ออกจากระบบ", on_click=do_logout, width="stretch", icon=":material/logout:")
        html_block('<div class="side-label">เมนูหลัก</div>')
        menu = st.radio("เมนูหลัก", ROLE_MENUS[user["role"]], key="main_menu", label_visibility="collapsed",
                        format_func=lambda m: f":material/{MENU_ICONS[m]}: {m}")
        st.divider()
        if uses_default(user["key"]):
            st.warning("ท่านยังใช้รหัสผ่านเริ่มต้น แนะนำให้เปลี่ยนรหัสผ่าน", icon=":material/shield:")
        with st.expander(":material/key: เปลี่ยนรหัสผ่าน"):
            with st.form("pw_form", clear_on_submit=True):
                old = st.text_input("รหัสผ่านเดิม", type="password")
                new = st.text_input(f"รหัสผ่านใหม่ (อย่างน้อย {MIN_PASSWORD_LEN} ตัว)", type="password")
                new2 = st.text_input("ยืนยันรหัสผ่านใหม่", type="password")
                if st.form_submit_button("บันทึก", width="stretch", type="primary", icon=":material/save:"):
                    if not check_password(user["key"], old):
                        st.error("รหัสผ่านเดิมไม่ถูกต้อง", icon=":material/error:")
                    elif len(new) < MIN_PASSWORD_LEN:
                        st.error(f"รหัสผ่านใหม่ต้องมีอย่างน้อย {MIN_PASSWORD_LEN} ตัว", icon=":material/error:")
                    elif new != new2:
                        st.error("ยืนยันรหัสผ่านใหม่ไม่ตรงกัน", icon=":material/error:")
                    else:
                        set_password(user["key"], new)
                        log_event(user, "เปลี่ยนรหัสผ่าน")
                        st.success("เปลี่ยนรหัสผ่านแล้ว", icon=":material/check_circle:")
        return menu


def main():
    inject_theme()
    user = st.session_state.get("auth_user")
    if not user:
        page_login()
        return
    menus = ROLE_MENUS[user["role"]]
    if st.session_state.get("main_menu") not in menus:
        st.session_state["main_menu"] = menus[0]
    menu = render_sidebar(user)
    render_hero(user, menu)
    if menu not in menus:
        st.error("คุณไม่มีสิทธิ์เข้าถึงหน้านี้", icon=":material/block:")
        return
    pages = {MENU_HOME: page_home, MENU_DATA: page_data, MENU_VIEW: page_view, MENU_USERS: page_users, MENU_HIST: page_history,
             MENU_T_TT: page_teacher_tt, MENU_T_LOAD: page_teacher_load,
             MENU_S_TT: page_student_tt, MENU_S_SUBJ: page_student_subj}
    pages[menu]()


if __name__ == "__main__":
    main()
