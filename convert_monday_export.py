"""
Convert a Monday.com "Export to Excel" file into the dashboard's data.xlsx.

    python convert_monday_export.py "ACC_OnDemand_Project_2026.xlsx" --portco ACC [--existing data.xlsx]

Each group becomes one project. Portcos / Communications sheets are kept from --existing,
Files are kept and Monday attachments are added.
"""
import argparse
import datetime as dt

from openpyxl import load_workbook

import monday_sync as ms
from project_rollup import assign_ids, file_rows, norm, rollup

SKIP = ["Pre-Project / Kickoff"]


def parse_export(path):
    ws = load_workbook(path, data_only=True).worksheets[0]
    rows = list(ws.iter_rows(values_only=True))
    groups, cur, hdr, sub_hdr, task = [], None, None, None, None
    for i, r in enumerate(rows):
        vals = list(r)
        first = vals[0]
        nonempty = [v for v in vals if v not in (None, "")]
        if i == 0 or not nonempty:
            continue
        if first == "Name":                         # task header
            hdr, sub_hdr = vals, None
            continue
        if first == "Subitems" and vals[1] == "Name":  # subitem header
            sub_hdr = vals
            continue
        if first and len(nonempty) == 1:            # group title row
            cur = {"name": str(first).strip(), "tasks": []}
            groups.append(cur)
            hdr, sub_hdr, task = None, None, None
            continue
        if first is None and sub_hdr and vals[1]:   # subitem row
            d = dict(zip(sub_hdr, vals))
            task["subitems"].append({"name": d.get("Name"), "status": d.get("Status"), "owner": d.get("Task Owner"),
                                     "start": d.get("Task Timeline - Start"), "end": d.get("Task Timeline - End")})
            continue
        if first is None:                            # group summary row
            continue
        if hdr and cur is not None:                  # task row
            d = dict(zip(hdr, vals))
            files = [u.strip() for u in str(d.get("Files") or "").split(", http") if u.strip()]
            files = [u if u.startswith("http") else "http" + u for u in files]
            task = {"name": d.get("Name"), "status": d.get("Status"), "owner": d.get("Task Owner"),
                    "priority": d.get("Task Priority"), "start": d.get("Task Timeline - Start"),
                    "end": d.get("Task Timeline - End"), "files": files, "subitems": []}
            sub_hdr = None
            cur["tasks"].append(task)
    return groups


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("export")
    ap.add_argument("--portco", required=True)
    ap.add_argument("--existing", default=None)
    ap.add_argument("--out", default="data.xlsx")
    a = ap.parse_args()

    groups = parse_export(a.export)
    projects = [p for g in groups if (p := rollup(g["name"], g["tasks"], a.portco, SKIP))]

    from openpyxl import Workbook
    wb = load_workbook(a.existing) if a.existing else Workbook()
    if "Sheet" in wb.sheetnames and len(wb.sheetnames) == 1:
        wb.remove(wb["Sheet"])
    existing = ms.read_sheet(wb, "Projects")
    assign_ids(projects, existing)

    ms.write_sheet(wb, "Projects", ms.PROJECT_HEADERS, [[p.get(h, "") for h in ms.PROJECT_HEADERS] for p in projects],
                   [11, 12, 46, 16, 16, 15, 13, 13, 17, 10, 15, 10], date_cols=("F", "G", "H"))

    portcos = ms.read_sheet(wb, "Portcos")
    if not any(norm(p.get("PortcoName")) == norm(a.portco) for p in portcos):
        portcos.append({"PortcoID": f"PC{len(portcos) + 1:02d}", "PortcoName": a.portco})
    ms.write_sheet(wb, "Portcos", ms.PORTCO_HEADERS, [[p.get(h, "") for h in ms.PORTCO_HEADERS] for p in portcos],
                   [11, 26, 20, 16, 32])

    comms = ms.read_sheet(wb, "Communications")
    ms.write_sheet(wb, "Communications", ms.COMMS_HEADERS, [[c.get(h, "") for h in ms.COMMS_HEADERS] for c in comms],
                   [13, 26, 14, 15, 11, 50, 16], date_cols=("A",))

    old_files = [f for f in ms.read_sheet(wb, "Files") if "monday.com" not in str(f.get("Link") or "")]
    rows = [[f.get(h, "") for h in ms.FILES_HEADERS] for f in old_files] + file_rows(projects)
    ms.write_sheet(wb, "Files", ms.FILES_HEADERS, rows, [52, 12, 11, 15, 9, 13, 60], date_cols=("F",))

    now = dt.datetime.now(dt.timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    ms.write_sheet(wb, "SyncInfo", ["Key", "Value"], [["Source", "Monday.com"], ["Boards", "Excel export"],
                   ["LastSync", now], ["Projects", len(projects)]], [14, 50])
    wb.save(a.out)
    for p in projects:
        print(f"{p['ProjectID']:8} {p['Status']:16} {p['Progress']:>3}%  due {p['DueDate'] and p['DueDate']:%d-%b}  {p['ProjectName']}"
              if p["DueDate"] else f"{p['ProjectID']:8} {p['Status']:16} {p['Progress']:>3}%  no dates  {p['ProjectName']}")
    print(f"Saved {a.out}: {len(projects)} projects, {len(rows)} files")


if __name__ == "__main__":
    main()
