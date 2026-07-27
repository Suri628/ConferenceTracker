"""同步会议基础字段（地点/日期/官网），数据来自持续维护的 ccfddl/ccf-deadlines 仓库。

目标文件 data/Conferences.xlsx，sheet "2027A类会议"。会议列表（数量/名字/顺序）完全以
xlsx 里已有的行为准——本脚本**不会**自动增删或调整会议顺序：
- 已有行的"刊物简称"如果能在 ccfddl 里匹配上，就刷新它的"会议地点"/"会议时间"/"会议URL"
- 匹配不上（比如不在 CCF 名单里、你手动加的会议，如 ISSCC/IEDM/VLSI/WINE）原样保留，不碰
- 匹配上了但 ccfddl 里还没有 TARGET_YEAR 这一届的数据，"会议地点"/"会议时间"填 TBD，"会议URL"留空
- "Workshop Deadline" / "workshopURL" / "备注" 这几列不碰，那部分由 check_workshop_deadlines.py 维护

只有 data/Conferences.xlsx 还不存在时，才会报错要求先手动建好这张表（这个脚本不负责初始化新表）。
"""
import glob
import os
import re
import shutil
import subprocess
import sys
import tempfile

import openpyxl
import yaml
from openpyxl.styles import Font, PatternFill

REPO_URL = "https://github.com/ccfddl/ccf-deadlines.git"
XLSX_PATH = os.path.join(os.path.dirname(__file__), "..", "data", "Conferences.xlsx")
SHEET_NAME = "2027A类会议"

# 表格只关心这一届；到了下一年需要手动把这个数字往前推一年。
TARGET_YEAR = 2027

ACRONYM_COL = "刊物简称"
LOCATION_COL = "会议地点"
DATE_COL = "会议时间"
URL_COL = "会议URL"

# 本次运行（sync_base_info.py + check_workshop_deadlines.py 算一个完整周期）里有变化的行，
# 整行标黄；每次 sync_base_info.py 开跑会先把上一轮的高亮清空，重新开始记
HIGHLIGHT_FILL = PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid")
NO_FILL = PatternFill(fill_type=None)


def set_row_highlight(row, highlighted: bool):
    fill = HIGHLIGHT_FILL if highlighted else NO_FILL
    for cell in row:
        cell.fill = fill

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

# "刊物简称"跟 ccfddl 里的 title 拼写/缩写不完全一致的，手动对齐一下。
# 值可以写 "slug" 或 "slug:CATEGORY"——ccfddl 里有撞名的情况（比如 FSE 既是软工的
# Foundations of Software Engineering 又是密码学的 Fast Software Encryption），
# 不写 CATEGORY 消歧的话，撞名时到底匹配到哪个取决于字典遍历顺序，在 Windows/Linux
# 上可能不一样（本地测试是对的，Actions 上跑出来却是错的，就是这个原因），所以必须显式指定
ALIASES = {
    "usenixatc": "sigopsatc", "fseesec": "fse:SE", "vr": "ieeevr",
    "siggraph": "acmsiggraph", "ubicomp": "ubicompiswc",
    # 不在 ccfddl / CCF 名单里的会议，明确标 None，不去猜
    "isscc": None, "wine": None, "iedm": None, "vlsi": None,
}

# ccfddl 的 place 字段是"场馆, 城市, 国家"这种三段式时，通用规则分不清楚该丢哪一段
# （跟"城市, 省, 国家"长得一样但处理方式相反），已知的几个手动订正一下
LOCATION_OVERRIDES = {
    "chi": "Pittsburgh, (PA), USA",
}

# ccfddl 的 date 字段有时候没跟着官网更新（众包维护，可能滞后），已知有问题的手动订正一下
# HPCA 2027：ccfddl 写的是 Jan 30-Feb 3，但官网 conf.researchr.org/home/hpca-2027
# 明确写的是 March 20-24（跟同期同地联合举办的 PPoPP 对得上），2026-07 核实过
DATE_OVERRIDES = {
    "hpca": "03.20-03.24",
}


def clone_ccfddl(tmpdir: str) -> str:
    subprocess.run(
        ["git", "clone", "--depth", "1", REPO_URL, tmpdir],
        check=True, capture_output=True,
    )
    return os.path.join(tmpdir, "conference")


def slugify(title: str) -> str:
    return re.sub(r"[^a-z0-9]+", "", title.lower())


def load_ccf_lookup(conference_dir: str) -> dict:
    """(slug, category) -> ccfddl entry，覆盖全部 rank（不只是 CCF-A，比如 IJCAI 是 CCF-B）。
    用 (slug, category) 做 key 而不是单纯 slug，是为了容纳 ccfddl 里撞名的会议——
    两个会议同名但不同 category 时不会互相覆盖，都能查到，靠 ALIASES 里的 CATEGORY 去挑对的那个"""
    lookup = {}
    for path in glob.glob(os.path.join(conference_dir, "*", "*.yml")):
        with open(path, encoding="utf-8") as f:
            docs = yaml.safe_load(f)
        for entry in (docs if isinstance(docs, list) else [docs]):
            if entry:
                lookup[(slugify(entry["title"]), entry.get("sub"))] = entry
    return lookup


def resolve_entry(acronym: str, lookup: dict):
    """按"刊物简称"找 ccfddl 里对应的 entry。撞名的（同一个 slug 对应多个 category）
    必须在 ALIASES 里用 "slug:CATEGORY" 显式指定，不然宁可返回 None 也不瞎猜——
    瞎猜的话结果会取决于字典遍历顺序，不同系统上可能不一样"""
    key = slugify(acronym)
    alias = ALIASES[key] if key in ALIASES else key
    if alias is None:
        return None
    if ":" in alias:
        slug, category = alias.split(":", 1)
        return lookup.get((slug, category))
    matches = [entry for (s, _c), entry in lookup.items() if s == alias]
    return matches[0] if len(matches) == 1 else None


def pick_target_year_conf(confs: list, target_year: int) -> dict:
    for c in confs or []:
        try:
            if int(c.get("year", 0)) == target_year:
                return c
        except (TypeError, ValueError):
            continue
    return {}


def format_date_range(raw: str) -> str:
    """把 ccfddl 里 'July 23-29, 2022' 这类自由文本转成 mm.dd-mm.dd；解析不出来原样返回"""
    if not raw:
        return "TBD"
    m = DATE_RANGE_RE.search(raw)
    if not m:
        return raw
    mon1, day1, mon2, day2, _year = m.groups()
    n1 = MONTH_NUM.get(mon1.lower())
    n2 = MONTH_NUM.get(mon2.lower()) if mon2 else n1
    if not n1 or not n2:
        return raw
    return f"{n1:02d}.{int(day1):02d}-{n2:02d}.{int(day2):02d}"


def format_location(place: str) -> str:
    """'Salt Lake City, Utah' -> 'Salt Lake City, (Utah), USA'；'Kyoto, Japan' 原样两段式保留"""
    if not place:
        return "TBD"
    place = place.rstrip(".").strip()
    parts = [p.strip() for p in place.split(",")]
    if len(parts) < 2:
        return parts[0] if parts else "TBD"
    city, last = ", ".join(parts[:-1]), parts[-1]
    if last.lower() in US_STATES:
        return f"{city}, ({last}), USA"
    if last.upper() in ("USA", "US") and len(parts) >= 3:
        maybe_state = parts[-2]
        if maybe_state.lower() in US_STATES:
            city2 = ", ".join(parts[:-2])
            return f"{city2}, ({maybe_state}), USA"
        return f"{city}, {last}"
    return f"{city}, {last}"


HYPERLINK_FONT = Font(color="0563C1", underline="single")


def set_url(cell, url: str):
    """把 url 写进单元格并做成可点击的超链接。
    清空时必须把 cell.hyperlink 也一起清掉——只清 value 的话，openpyxl 存盘再读回来会
    用残留的 hyperlink 把显示文本"复活"成旧链接，哪怕 value 当时明明已经设成空字符串了"""
    cell.value = url or ""
    if url:
        cell.hyperlink = url
        cell.font = HYPERLINK_FONT
    else:
        cell.hyperlink = None
        cell.font = Font()


def refresh_row(row_cells: dict, entry: dict, slug: str):
    """entry 是 ccfddl 的一条会议记录，刷新 row_cells 里的 会议地点/会议时间/会议URL"""
    target = pick_target_year_conf(entry.get("confs") or [], TARGET_YEAR)
    if target:
        row_cells[LOCATION_COL].value = LOCATION_OVERRIDES.get(slug) or format_location(target.get("place", ""))
        row_cells[DATE_COL].value = DATE_OVERRIDES.get(slug) or format_date_range(target.get("date", ""))
        set_url(row_cells[URL_COL], target.get("link", ""))
    else:
        row_cells[LOCATION_COL].value = "TBD"
        row_cells[DATE_COL].value = "TBD"
        set_url(row_cells[URL_COL], "")


def main():
    if not os.path.exists(XLSX_PATH):
        print(f"未找到 {XLSX_PATH}，本脚本不负责初始化新表，请先手动建好会议清单。", file=sys.stderr)
        sys.exit(1)

    tmpdir = tempfile.mkdtemp(prefix="ccfddl_")
    try:
        conference_dir = clone_ccfddl(tmpdir)
        lookup = load_ccf_lookup(conference_dir)

        wb = openpyxl.load_workbook(XLSX_PATH)
        ws = wb[SHEET_NAME]
        header = [c.value for c in ws[1]]
        idx = {name: i for i, name in enumerate(header) if name}

        refreshed, skipped, changed = 0, 0, 0
        for row in ws.iter_rows(min_row=2):
            acronym = row[idx[ACRONYM_COL]].value
            if not acronym:
                continue
            # 新一轮同步开始，先清掉上一轮留下的高亮
            set_row_highlight(row, False)

            entry = resolve_entry(acronym, lookup)
            if entry is None:
                skipped += 1
                continue
            override_key = slugify(acronym)
            row_cells = {col: row[idx[col]] for col in (LOCATION_COL, DATE_COL, URL_COL)}
            # openpyxl 存盘再读回来会把 "" 变成 None，比较时要当成一样，否则每次都会被误判成"变了"
            before = {col: (cell.value or "") for col, cell in row_cells.items()}
            refresh_row(row_cells, entry, override_key)
            after = {col: (cell.value or "") for col, cell in row_cells.items()}
            if before != after:
                set_row_highlight(row, True)
                changed += 1
            refreshed += 1

        wb.save(XLSX_PATH)
        print(f"Refreshed {refreshed} conferences ({changed} changed), left {skipped} manual/unmatched rows untouched.")
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)


if __name__ == "__main__":
    main()
