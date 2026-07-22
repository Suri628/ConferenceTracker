"""同步 CCF-A 会议基础字段（地点/日期/主会投稿deadline）。

数据来自持续维护的 ccfddl/ccf-deadlines 仓库，本脚本只负责重建"基础信息"，
不会覆盖已抓取的 workshop_proposal_* 相关列（那部分由 check_workshop_deadlines.py 维护）。
"""
import glob
import json
import os
import re
import shutil
import subprocess
import tempfile
from datetime import datetime, timezone

import openpyxl
import yaml
from openpyxl.utils import get_column_letter

REPO_URL = "https://github.com/ccfddl/ccf-deadlines.git"
XLSX_PATH = os.path.join(os.path.dirname(__file__), "..", "data", "conferences.xlsx")
OVERRIDES_PATH = os.path.join(os.path.dirname(__file__), "..", "data", "conference_overrides.json")

# 表格只关心这一届；没有 TARGET_YEAR 数据的会议，日期/地点/官网都填 TBD，不展示旧年份的信息。
# 到了下一年需要手动把这个数字往前推一年。
TARGET_YEAR = 2027

COLUMNS = [
    "conference_id", "name", "acronym", "year", "category",
    "location_city", "location_country", "location_region",
    "date", "official_url",
    "main_cfp_deadline",
    "workshop_proposal_deadline", "workshop_proposal_status",
    "workshop_page_url", "last_checked_at", "last_changed_at",
    "source_url", "notes",
]

# check_workshop_deadlines.py 负责维护的列；同步基础信息时原样保留
WORKSHOP_COLUMNS = [
    "workshop_proposal_deadline", "workshop_proposal_status",
    "workshop_page_url", "last_checked_at", "last_changed_at",
    "source_url", "notes",
]

REGION_BY_COUNTRY = {
    "usa": "North America", "us": "North America", "united states": "North America",
    "canada": "North America",
    "mexico": "Latin America", "brazil": "Latin America", "argentina": "Latin America",
    "chile": "Latin America", "colombia": "Latin America",
    "uk": "Europe", "united kingdom": "Europe", "germany": "Europe", "france": "Europe",
    "italy": "Europe", "spain": "Europe", "netherlands": "Europe", "switzerland": "Europe",
    "austria": "Europe", "ireland": "Europe", "croatia": "Europe", "slovenia": "Europe",
    "portugal": "Europe", "greece": "Europe", "sweden": "Europe", "belgium": "Europe",
    "czechia": "Europe", "czech republic": "Europe", "morocco": "Other",
    "china": "Asia", "japan": "Asia", "south korea": "Asia", "korea": "Asia",
    "republic of korea": "Asia", "singapore": "Asia", "india": "Asia", "israel": "Asia",
    "hong kong": "Asia",
    "australia": "Other", "new zealand": "Other",
}

# ccfddl 的 place 字段格式不统一，美国场馆常以州名/州缩写结尾而不是国家名，这里做归一化
US_STATES = {
    "alabama", "alaska", "arizona", "arkansas", "california", "colorado", "connecticut",
    "delaware", "florida", "georgia", "hawaii", "idaho", "illinois", "indiana", "iowa",
    "kansas", "kentucky", "louisiana", "maine", "maryland", "massachusetts", "michigan",
    "minnesota", "mississippi", "missouri", "montana", "nebraska", "nevada",
    "new hampshire", "new jersey", "new mexico", "new york", "north carolina",
    "north dakota", "ohio", "oklahoma", "oregon", "pennsylvania", "rhode island",
    "south carolina", "south dakota", "tennessee", "texas", "utah", "vermont",
    "virginia", "washington", "west virginia", "wisconsin", "wyoming",
    "al", "ak", "az", "ar", "ca", "co", "ct", "de", "fl", "ga", "hi", "id", "il", "in",
    "ia", "ks", "ky", "la", "me", "md", "ma", "mi", "mn", "ms", "mo", "mt", "ne", "nv",
    "nh", "nj", "nm", "ny", "nc", "nd", "oh", "ok", "or", "pa", "ri", "sc", "sd", "tn",
    "tx", "ut", "vt", "va", "wa", "wv", "wi", "wy",
}


def clone_ccfddl(tmpdir: str) -> str:
    subprocess.run(
        ["git", "clone", "--depth", "1", REPO_URL, tmpdir],
        check=True, capture_output=True,
    )
    return os.path.join(tmpdir, "conference")


def slugify(title: str) -> str:
    return re.sub(r"[^a-z0-9]+", "", title.lower())


def pick_target_year_conf(confs: list, target_year: int) -> dict:
    """只要 target_year 那一届的数据；ccfddl 里没有就返回空 dict（调用方据此填 TBD）"""
    for c in confs or []:
        try:
            if int(c.get("year", 0)) == target_year:
                return c
        except (TypeError, ValueError):
            continue
    return {}


def split_place(place: str):
    if not place:
        return "", ""
    place = place.rstrip(".").strip()
    parts = [p.strip() for p in place.split(",")]
    if len(parts) < 2:
        return place, ""
    city, country = ", ".join(parts[:-1]), parts[-1]
    if country.lower() in US_STATES:
        # 州名/州缩写不是国家，把它并回城市部分，国家记为 USA
        return f"{city}, {country}", "USA"
    return city, country


def region_for(country: str) -> str:
    return REGION_BY_COUNTRY.get(country.strip().lower(), "Other")


def clean_date(value) -> str:
    if not value:
        return "TBD"
    return str(value).split(" ")[0]


MONTH_NUM = {
    "jan": 1, "january": 1, "feb": 2, "february": 2, "mar": 3, "march": 3,
    "apr": 4, "april": 4, "may": 5, "jun": 6, "june": 6, "jul": 7, "july": 7,
    "aug": 8, "august": 8, "sep": 9, "sept": 9, "september": 9,
    "oct": 10, "october": 10, "nov": 11, "november": 11, "dec": 12, "december": 12,
}

DATE_RANGE_RE = re.compile(
    r"([A-Za-z]+)\.?\s+(\d{1,2})(?:st|nd|rd|th)?"
    r"\s*(?:-|–|—|to)\s*"
    r"(?:([A-Za-z]+)\.?\s+)?(\d{1,2})(?:st|nd|rd|th)?,?\s+(\d{4})"
)


def format_date_range(raw: str) -> str:
    """把 ccfddl 里 'July 23-29, 2022' / 'October 12 - 16, 2026' 这类自由文本
    转成 mm.dd-mm.dd；解析不出来就原样返回，交给人工核实"""
    if not raw:
        return "TBD"
    m = DATE_RANGE_RE.search(raw)
    if not m:
        return raw
    mon1, day1, mon2, day2, _year = m.groups()
    mon1_num = MONTH_NUM.get(mon1.lower())
    mon2_num = MONTH_NUM.get(mon2.lower()) if mon2 else mon1_num
    if not mon1_num or not mon2_num:
        return raw
    return f"{mon1_num:02d}.{int(day1):02d}-{mon2_num:02d}.{int(day2):02d}"


def load_ccf_a_conferences(conference_dir: str) -> list:
    rows = []
    for path in sorted(glob.glob(os.path.join(conference_dir, "*", "*.yml"))):
        with open(path, encoding="utf-8") as f:
            docs = yaml.safe_load(f)
        entries = docs if isinstance(docs, list) else [docs]
        for entry in entries:
            if not entry:
                continue
            rank = entry.get("rank") or {}
            if rank.get("ccf") != "A":
                continue
            target = pick_target_year_conf(entry.get("confs") or [], TARGET_YEAR)
            if target:
                city, country = split_place(target.get("place", ""))
                timeline = (target.get("timeline") or [{}])[0]
                row = {
                    "conference_id": slugify(entry["title"]),
                    "name": entry.get("description", ""),
                    "acronym": entry["title"],
                    "year": TARGET_YEAR,
                    "category": entry.get("sub", ""),
                    "location_city": city,
                    "location_country": country,
                    "location_region": region_for(country),
                    "date": format_date_range(target.get("date", "")),
                    "official_url": target.get("link", ""),
                    "main_cfp_deadline": clean_date(timeline.get("deadline", "")),
                }
            else:
                # ccfddl 还没有 TARGET_YEAR 这一届的数据，一律 TBD，不展示旧年份信息
                row = {
                    "conference_id": slugify(entry["title"]),
                    "name": entry.get("description", ""),
                    "acronym": entry["title"],
                    "year": TARGET_YEAR,
                    "category": entry.get("sub", ""),
                    "location_city": "",
                    "location_country": "",
                    "location_region": "",
                    "date": "TBD",
                    "official_url": "",
                    "main_cfp_deadline": "TBD",
                }
            rows.append(row)
    rows.sort(key=lambda r: r["conference_id"])
    return rows


def load_overrides(path: str) -> dict:
    """人工维护的会议增删/排序配置，格式见 data/conference_overrides.json：
    - excluded_ids: 想从表里删掉的 conference_id（下次 sync 也不会再自动加回来）
    - custom_order: 想固定在前面的 conference_id 顺序，没写到的会议按字母序排在后面
    """
    if not os.path.exists(path):
        return {"excluded_ids": [], "custom_order": []}
    with open(path, encoding="utf-8") as f:
        data = json.load(f)
    return {
        "excluded_ids": data.get("excluded_ids", []),
        "custom_order": data.get("custom_order", []),
    }


def apply_overrides(rows: list, overrides: dict) -> list:
    excluded = set(overrides.get("excluded_ids", []))
    rows = [r for r in rows if r["conference_id"] not in excluded]

    custom_order = overrides.get("custom_order", [])
    order_index = {cid: i for i, cid in enumerate(custom_order)}
    rows.sort(key=lambda r: (order_index.get(r["conference_id"], len(order_index)), r["conference_id"]))
    return rows


def load_existing_workshop_data(path: str) -> dict:
    if not os.path.exists(path):
        return {}
    wb = openpyxl.load_workbook(path)
    if "会议数据" not in wb.sheetnames:
        return {}
    ws = wb["会议数据"]
    header = [c.value for c in ws[1]]
    idx = {name: i for i, name in enumerate(header) if name}
    existing = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or "conference_id" not in idx or not row[idx["conference_id"]]:
            continue
        cid = row[idx["conference_id"]]
        existing[cid] = {col: row[idx[col]] for col in WORKSHOP_COLUMNS if col in idx}
    return existing


def write_xlsx(rows: list, path: str):
    existing_workshop = load_existing_workshop_data(path)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "会议数据"
    ws.append(COLUMNS)

    for row in rows:
        prev = existing_workshop.get(row["conference_id"], {})
        for col in WORKSHOP_COLUMNS:
            default = "not_yet_announced" if col == "workshop_proposal_status" else ""
            row.setdefault(col, prev.get(col, default))
        ws.append([row.get(col, "") for col in COLUMNS])

    for col_idx in range(1, len(COLUMNS) + 1):
        ws.column_dimensions[get_column_letter(col_idx)].width = 20

    log_ws = wb.create_sheet("变更日志")
    log_ws.append(["conference_id", "field", "old_value", "new_value", "changed_at"])
    if os.path.exists(path):
        old_wb = openpyxl.load_workbook(path)
        if "变更日志" in old_wb.sheetnames:
            for row in old_wb["变更日志"].iter_rows(min_row=2, values_only=True):
                if row and row[0]:
                    log_ws.append(row)

    wb.save(path)
    print(f"Wrote {len(rows)} CCF-A conferences to {path}")


def main():
    tmpdir = tempfile.mkdtemp(prefix="ccfddl_")
    try:
        conference_dir = clone_ccfddl(tmpdir)
        rows = load_ccf_a_conferences(conference_dir)
        rows = apply_overrides(rows, load_overrides(OVERRIDES_PATH))
        os.makedirs(os.path.dirname(XLSX_PATH), exist_ok=True)
        write_xlsx(rows, XLSX_PATH)
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)


if __name__ == "__main__":
    main()
