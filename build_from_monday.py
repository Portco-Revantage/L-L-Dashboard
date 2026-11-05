"""
Daily dashboard refresh - builds data.xlsx from a live Monday board read.

  python build_from_monday.py monday.json previous_data.xlsx new_data.xlsx [--portco ACC]

monday.json = the API response for the board, containing  boards[0].items_page.items  with:
  name, group{title}, column_values{id,text} for project_owner, project_status, project_timeline,
  project_priority, file_mkw4nbdc, and subitems{name, column_values{id,text}} for
  multiple_person_mkxfbhpb, color_mkxfypxq, timerange_mkxfarcc.
Rules: each group = one project ("Pre-Project / Kickoff" skipped); progress = % of subitems Done;
ProjectIDs are kept from previous_data.xlsx; Portcos/Communications/manual Files are kept.
Prints a JSON summary of what changed (new / changed / removed projects).
"""
import argparse, datetime as dt, json, re, shutil, sys
from openpyxl import load_workbook
import monday_sync as ms
from project_rollup import rollup, assign_ids, file_rows, norm

SKIP = ["Pre-Project / Kickoff"]
ITEM = {"owner": "project_owner", "status": "project_status", "tl": "project_timeline", "prio": "project_priority", "files": "file_mkw4nbdc"}
SUB = {"owner": "multiple_person_mkxfbhpb", "status": "color_mkxfypxq", "tl": "timerange_mkxfarcc"}


def d(tl, end=False):
    f = re.findall(r"\d{4}-\d{2}-\d{2}", tl or "")
    return dt.datetime.strptime(f[-1] if end else f[0], "%Y-%m-%d") if f else None


def cv(item, cid):
    for c in item.get("column_values") or []:
        if c.get("id") == cid:
            return c.get("text") or ""
    return ""


def items_from(payload):
    p = payload.get("data", payload)
    if "boards" in p:
        return p["boards"][0]["items_page"]["items"]
    if "items" in p:
        return p["items"]
    raise SystemExit("Unrecognised Monday response")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("monday_json"); ap.add_argument("previous"); ap.add_argument("out")
    ap.add_argument("--portco", default="ACC"); ap.add_argument("--board-name", default="ACC OnDemand Project 2026")
    a = ap.parse_args()
    items = items_from(json.load(open(a.monday_json, encoding="utf-8")))
    groups = {}
    for it in items:
        files = re.findall(r"https?://.+?(?=, https?://|$)", cv(it, ITEM["files"]))
        task = {"name": it["name"], "owner": cv(it, ITEM["owner"]), "status": cv(it, ITEM["status"]),
                "start": d(cv(it, ITEM["tl"])), "end": d(cv(it, ITEM["tl"]), True), "priority": cv(it, ITEM["prio"]),
                "files": [f.strip() for f in files],
                "subitems": [{"name": s["name"], "owner": cv(s, SUB["owner"]), "status": cv(s, SUB["status"]),
                              "start": d(cv(s, SUB["tl"])), "end": d(cv(s, SUB["tl"]), True)} for s in it.get("subitems") or []]}
        groups.setdefault(it["group"]["title"], []).append(task)
    projects = [p for g, t in groups.items() if (p := rollup(g, t, a.portco, SKIP))]
    if not projects:
        raise SystemExit("No projects found - previous data left unchanged.")

    shutil.copy(a.previous, a.out)
    wb = load_workbook(a.out)
    existing = ms.read_sheet(wb, "Projects")
    before = {norm(r.get("ProjectName")): r for r in existing}
    assign_ids(projects, existing)
    others = [r for r in existing if norm(r.get("PortcoName")) != norm(a.portco)]   # other portcos' boards stay as they are
    ms.write_sheet(wb, "Projects", ms.PROJECT_HEADERS, [[r.get(h, "") for h in ms.PROJECT_HEADERS] for r in others] +
                   [[p.get(h, "") for h in ms.PROJECT_HEADERS] for p in projects],
                   [11, 12, 46, 16, 16, 15, 13, 13, 17, 10, 15, 10], date_cols=("F", "G", "H"))
    kept = [f for f in ms.read_sheet(wb, "Files")
            if "monday.com" not in str(f.get("Link") or "") or norm(f.get("PortcoName")) != norm(a.portco)]
    ms.write_sheet(wb, "Files", ms.FILES_HEADERS, [[f.get(h, "") for h in ms.FILES_HEADERS] for f in kept] + file_rows(projects),
                   [52, 12, 11, 15, 9, 13, 60], date_cols=("F",))
    now = dt.datetime.now(dt.timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    ms.write_sheet(wb, "SyncInfo", ["Key", "Value"], [["Source", "Monday.com"], ["Boards", f"{a.board_name} (live read)"],
                   ["LastSync", now], ["Projects", len(projects) + len(others)]], [14, 50])
    wb.save(a.out)

    fields = ["Status", "Progress", "DueDate", "StartDate", "IDOwner", "Priority", "FeedbackRounds"]
    fmt = lambda v: v.strftime("%d %b %Y") if isinstance(v, dt.datetime) else ("" if v is None else v)
    new, changed = [], []
    by_id = {r.get("ProjectID"): r for r in existing}
    for p in projects:
        b = by_id.get(p["ProjectID"])
        if not b:
            new.append({"id": p["ProjectID"], "project": p["ProjectName"], "status": p["Status"], "progress": p["Progress"]}); continue
        diff = {f: [fmt(b.get(f)), fmt(p.get(f))] for f in fields if fmt(b.get(f)) != fmt(p.get(f)) and not (b.get(f) in (None, "") and p.get(f) in (None, ""))}
        if norm(b.get("ProjectName")) != norm(p["ProjectName"]):
            diff["Name"] = [b.get("ProjectName"), p["ProjectName"]]
        if diff:
            changed.append({"id": p["ProjectID"], "project": p["ProjectName"], "changes": diff})
    ids = {p["ProjectID"] for p in projects}
    removed = [r.get("ProjectName") for r in existing if r.get("ProjectID") not in ids and norm(r.get("PortcoName")) == norm(a.portco)]
    today = dt.datetime.now()
    act = [p for p in projects if p["Status"] not in ("Delivered", "On Hold")]
    overdue = [p["ProjectName"] for p in act if p["DueDate"] and p["DueDate"].date() < today.date()]
    nodates = [p["ProjectName"] for p in act if not p["DueDate"]]
    print(json.dumps({"portco": a.portco, "projects": len(projects), "delivered": sum(p["Status"] == "Delivered" for p in projects),
                      "new": new, "changed": changed, "removed": removed, "overdue": overdue, "no_dates": nodates,
                      "any_change": bool(new or changed or removed)}, indent=1, default=str))


if __name__ == "__main__":
    main()
