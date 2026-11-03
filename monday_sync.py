"""
PortCo Learning Hub - Monday.com sync
-------------------------------------
Reads your Monday.com board(s) and rewrites the "Projects" sheet in data.xlsx.
Group mode (default): each Monday GROUP is one project; tasks = ADDIE phases, subitems = steps.
The Portcos, Communications and Files sheets you maintain in Excel are kept.

Runs automatically on GitHub (see .github/workflows/monday-sync.yml).
Needs:  MONDAY_API_TOKEN   (GitHub secret)
        MONDAY_BOARD_IDS   (GitHub variable, comma separated) - or "board_ids" in monday-config.json
"""
import datetime as dt
import json
import os
import re
import sys
import urllib.request

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation

from project_rollup import assign_ids, file_rows, rollup

API_URL = "https://api.monday.com/v2"
API_VERSION = "2024-10"
HERE = os.path.dirname(os.path.abspath(__file__))
DATA_FILE = os.path.join(HERE, "data.xlsx")
CONFIG_FILE = os.path.join(HERE, "monday-config.json")

PROJECT_HEADERS = ["ProjectID", "PortcoName", "ProjectName", "ContentType", "IDOwner", "ScriptReceived",
                   "StartDate", "DueDate", "Status", "Progress", "FeedbackRounds", "Priority"]
PORTCO_HEADERS = ["PortcoID", "PortcoName", "Industry", "SPOC", "SPOCEmail"]
COMMS_HEADERS = ["Date", "PortcoName", "ProjectID", "Type", "Direction", "Subject", "Owner"]
FILES_HEADERS = ["FileName", "PortcoName", "ProjectID", "FileType", "Version", "UploadedOn", "Link"]
DASH_STATUSES = ["Script Received", "In Development", "In Review", "Portco Feedback", "Delivered", "On Hold"]
MONDAY_COMM_TYPE = "Monday update"


# ---------------------------------------------------------------- Monday API
def gql(query, variables, token):
    body = json.dumps({"query": query, "variables": variables}).encode()
    req = urllib.request.Request(API_URL, data=body, method="POST", headers={
        "Content-Type": "application/json", "Authorization": token, "API-Version": API_VERSION})
    with urllib.request.urlopen(req, timeout=60) as r:
        out = json.loads(r.read().decode())
    if out.get("errors") or out.get("error_message"):
        raise RuntimeError(f"Monday API error: {out.get('errors') or out.get('error_message')}")
    return out["data"]


def item_fields(with_updates):
    upd = "updates(limit: 50) { created_at text_body creator { name } }" if with_updates else ""
    cv = "column_values { id type text column { title } }"
    return f"id name group {{ id title }} {cv} subitems {{ id name {cv} }} {upd}"


def fetch_board(board_id, token, with_updates):
    f = item_fields(with_updates)
    first = gql(f"""query ($ids: [ID!]) {{ boards(ids: $ids) {{ id name columns {{ id title type }} groups {{ id title }}
        items_page(limit: 100) {{ cursor items {{ {f} }} }} }} }}""", {"ids": [str(board_id)]}, token)
    boards = first.get("boards") or []
    if not boards:
        raise RuntimeError(f"Board {board_id} not found, or the token has no access to it.")
    board = boards[0]
    items = list(board["items_page"]["items"])
    cursor = board["items_page"]["cursor"]
    while cursor:
        page = gql(f"""query ($c: String!) {{ next_items_page(limit: 100, cursor: $c) {{ cursor items {{ {f} }} }} }}""",
                   {"c": cursor}, token)["next_items_page"]
        items += page["items"]
        cursor = page["cursor"]
    return board, items


# ---------------------------------------------------------------- Mapping helpers
def find_column(columns, names):
    """Return the Monday column id whose title matches any name (case-insensitive)."""
    wanted = [n.strip().lower() for n in names]
    for n in wanted:
        for c in columns:
            if c["title"].strip().lower() == n:
                return c["id"]
    return None


def parse_date(text, end=False):
    """Handles '2026-11-05', '2026-11-05 10:00' and timelines '2026-11-01 - 2026-11-20'."""
    if not text:
        return None
    found = re.findall(r"\d{4}-\d{2}-\d{2}", text)
    if not found:
        return None
    try:
        return dt.datetime.strptime(found[-1] if end else found[0], "%Y-%m-%d")
    except ValueError:
        return None


def parse_number(text):
    m = re.search(r"-?\d+(\.\d+)?", text or "")
    return float(m.group()) if m else None


def map_status(raw, status_map):
    raw = (raw or "").strip()
    for k, v in status_map.items():
        if k.lower() == raw.lower():
            return v
    return raw or "Script Received"


def build_projects(board, items, cfg):
    cols = board["columns"]
    ids = {field: find_column(cols, names) for field, names in cfg["columns"].items()}
    pid_col = find_column(cols, cfg.get("project_id_column", []))
    timeline_ids = {c["id"] for c in cols if c["type"] in ("timeline", "timerange")}
    rows, comms = [], []
    for it in items:
        vals = {cv["id"]: (cv.get("text") or "").strip() for cv in it["column_values"]}
        get = lambda field: vals.get(ids.get(field) or "", "")
        portco = get("PortcoName")
        if not portco and cfg.get("use_group_name_as_portco_if_no_column", True):
            portco = (it.get("group") or {}).get("title", "")
        status = map_status(get("Status"), cfg["status_map"])
        progress = parse_number(get("Progress"))
        if progress is None:
            progress = cfg["progress_by_status"].get(status, 0)
        progress = max(0, min(100, round(progress)))
        due_id = ids.get("DueDate")
        due = parse_date(get("DueDate"), end=due_id in timeline_ids)
        start = parse_date(get("StartDate"))
        if start is None and due_id in timeline_ids:
            start = parse_date(get("DueDate"))           # timeline start
        pid = vals.get(pid_col, "") if pid_col else ""
        pid = pid or f"MON-{it['id']}"
        rows.append([pid, portco, it["name"], get("ContentType"), get("IDOwner"), parse_date(get("ScriptReceived")),
                     start, due, status, progress, int(parse_number(get("FeedbackRounds")) or 0), get("Priority")])
        for u in it.get("updates") or []:
            text = re.sub(r"\s+", " ", u.get("text_body") or "").strip()
            comms.append([parse_date(u.get("created_at")), portco, pid, MONDAY_COMM_TYPE, "Sent",
                          (text[:120] + "…") if len(text) > 120 else text, (u.get("creator") or {}).get("name", "")])
    return rows, comms


def _by_title(column_values):
    return {(cv.get("column") or {}).get("title", "").strip().lower(): (cv.get("text") or "").strip()
            for cv in column_values}


def _pick(vals, names):
    for n in names:
        if vals.get(n.lower()):
            return vals[n.lower()]
    return ""


def _timeline(vals, cfg):
    tl = _pick(vals, cfg["group_mode"]["timeline_columns"])
    return parse_date(tl), parse_date(tl, end=True)


def build_projects_by_group(board, items, cfg, portco):
    """Group mode: each Monday GROUP is one project; its items are phases, subitems are steps."""
    gm = cfg["group_mode"]
    order = [g["title"] for g in board.get("groups", [])]
    tasks_by_group = {t: [] for t in order}
    comms_raw = []
    for it in items:
        g = (it.get("group") or {}).get("title", "")
        vals = _by_title(it["column_values"])
        start, end = _timeline(vals, cfg)
        files = re.findall(r"https?://[^,]+?(?=, https?://|$)", _pick(vals, gm["files_columns"]))
        subs = []
        for sub in it.get("subitems") or []:
            sv = _by_title(sub["column_values"])
            ss, se = _timeline(sv, cfg)
            subs.append({"name": sub["name"], "status": _pick(sv, gm["status_columns"]),
                         "owner": _pick(sv, gm["owner_columns"]), "start": ss, "end": se})
        tasks_by_group.setdefault(g, []).append({
            "name": it["name"], "status": _pick(vals, gm["status_columns"]), "owner": _pick(vals, gm["owner_columns"]),
            "priority": _pick(vals, gm["priority_columns"]), "start": start, "end": end,
            "files": [f.strip() for f in files], "subitems": subs})
        for u in it.get("updates") or []:
            comms_raw.append((g, u))
    projects = [p for g, tasks in tasks_by_group.items()
                if (p := rollup(g, tasks, portco, gm.get("skip_groups", [])))]
    return projects, comms_raw


def board_portco(board, cfg):
    names = cfg["group_mode"].get("board_portcos", {})
    return names.get(str(board["id"])) or str(board["name"]).split()[0]


# ---------------------------------------------------------------- Excel
HDR_FILL = PatternFill("solid", fgColor="1F3A5F")
HDR_FONT = Font(bold=True, color="FFFFFF")


def read_sheet(wb, name):
    if name not in wb.sheetnames:
        return []
    rows = list(wb[name].iter_rows(values_only=True))
    if not rows:
        return []
    hdr = [str(h) if h is not None else "" for h in rows[0]]
    return [dict(zip(hdr, r)) for r in rows[1:] if any(v not in (None, "") for v in r)]


def write_sheet(wb, name, headers, rows, widths, date_cols=(), index=None):
    if name in wb.sheetnames:
        index = wb.sheetnames.index(name)
        del wb[name]
    ws = wb.create_sheet(name, index)
    ws.append(headers)
    for c in ws[1]:
        c.fill, c.font, c.alignment = HDR_FILL, HDR_FONT, Alignment(vertical="center")
    for r in rows:
        ws.append(r)
    for i, w in enumerate(widths):
        ws.column_dimensions[chr(65 + i)].width = w
    for col in date_cols:
        for c in ws[col][1:]:
            c.number_format = "dd-mmm-yyyy"
    ws.freeze_panes = "A2"
    return ws


def main():
    token = os.environ.get("MONDAY_API_TOKEN", "").strip()
    if not token:
        print("MONDAY_API_TOKEN is not set - skipping Monday sync (dashboard keeps current data.xlsx).")
        return 0
    with open(CONFIG_FILE, encoding="utf-8") as f:
        cfg = json.load(f)
    board_ids = [b.strip() for b in os.environ.get("MONDAY_BOARD_IDS", "").split(",") if b.strip()] \
        or [str(b) for b in cfg.get("board_ids", [])]
    if not board_ids:
        print("No board IDs. Set the MONDAY_BOARD_IDS variable in GitHub or 'board_ids' in monday-config.json.")
        return 1

    with_updates = bool(cfg.get("include_updates_as_communications"))
    group_mode = cfg.get("project_mode", "group") == "group"
    wb = load_workbook(DATA_FILE) if os.path.exists(DATA_FILE) else Workbook()
    if "Sheet" in wb.sheetnames and len(wb.sheetnames) == 1:
        wb.remove(wb["Sheet"])

    projects, monday_comms, names, monday_files = [], [], [], []
    for bid in board_ids:
        board, items = fetch_board(bid, token, with_updates)
        if group_mode:
            portco = board_portco(board, cfg)
            grp, raw = build_projects_by_group(board, items, cfg, portco)
            assign_ids(grp, read_sheet(wb, "Projects") + [{"ProjectName": p["ProjectName"], "ProjectID": p["ProjectID"]}
                                                          for p in projects])
            monday_files += file_rows(grp)
            ids = {p["ProjectName"]: p["ProjectID"] for p in grp}
            c = []
            for g, u in raw:
                if g not in ids:
                    continue
                text = re.sub(r"\s+", " ", u.get("text_body") or "").strip()
                c.append([parse_date(u.get("created_at")), portco, ids[g], MONDAY_COMM_TYPE, "Sent",
                          (text[:120] + "…") if len(text) > 120 else text, (u.get("creator") or {}).get("name", "")])
            p = [[r.get(h, "") for h in PROJECT_HEADERS] for r in grp]
        else:
            p, c = build_projects(board, items, cfg)
        projects += [{"ProjectName": r[2], "ProjectID": r[0], "row": r} for r in p]
        monday_comms += c
        names.append(board["name"])
        print(f"Board '{board['name']}' ({bid}): {len(p)} projects, {len(c)} updates")

    projects = [x["row"] for x in projects]
    if not projects:
        print("Monday returned 0 projects - data.xlsx left unchanged as a safety measure.")
        return 1

    # Projects - fully replaced from Monday
    ws = write_sheet(wb, "Projects", PROJECT_HEADERS, projects, [14, 26, 34, 15, 16, 15, 13, 13, 17, 10, 15, 10],
                     date_cols=("F", "G", "H"))
    dv = DataValidation(type="list", formula1='"' + ",".join(DASH_STATUSES) + '"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add("I2:I1000")

    # Portcos - keep Excel details, add any new portco names found in Monday
    portcos = read_sheet(wb, "Portcos")
    known = {str(p.get("PortcoName") or "").strip().lower() for p in portcos}
    for name in sorted({r[1] for r in projects if r[1]}):
        if name.strip().lower() not in known:
            portcos.append({"PortcoID": f"PC{len(portcos) + 1:02d}", "PortcoName": name})
            known.add(name.strip().lower())
    write_sheet(wb, "Portcos", PORTCO_HEADERS, [[p.get(h, "") for h in PORTCO_HEADERS] for p in portcos],
                [11, 26, 20, 16, 32])

    # Communications - keep Excel rows, refresh the Monday-update rows
    comms = [[r.get(h, "") for h in COMMS_HEADERS] for r in read_sheet(wb, "Communications")
             if str(r.get("Type") or "") != MONDAY_COMM_TYPE]
    comms += monday_comms
    comms.sort(key=lambda r: r[0] if isinstance(r[0], dt.datetime) else dt.datetime.min, reverse=True)
    write_sheet(wb, "Communications", COMMS_HEADERS, comms, [13, 26, 14, 15, 11, 50, 16], date_cols=("A",))

    # Files - keep rows you added by hand, refresh attachments that come from Monday
    files = [[f.get(h, "") for h in FILES_HEADERS] for f in read_sheet(wb, "Files")
             if "monday.com" not in str(f.get("Link") or "")] + monday_files
    write_sheet(wb, "Files", FILES_HEADERS, files, [52, 12, 11, 15, 9, 13, 60], date_cols=("F",))

    # Sync info for the dashboard
    now = dt.datetime.now(dt.timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    write_sheet(wb, "SyncInfo", ["Key", "Value"],
                [["Source", "Monday.com"], ["Boards", ", ".join(names)], ["LastSync", now],
                 ["Projects", len(projects)]], [14, 50])

    wb.save(DATA_FILE)
    print(f"data.xlsx updated: {len(projects)} projects, {len(monday_comms)} Monday updates, synced {now}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
