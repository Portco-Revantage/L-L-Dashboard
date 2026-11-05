"""
Turns a Monday board where each GROUP is one project (tasks = ADDIE phases, subitems = steps)
into one dashboard row per project. Used by both monday_sync.py (live API) and
convert_monday_export.py (manual Excel export), so both give identical results.
"""
import datetime as dt
import re
from collections import Counter

PHASE_ORDER = ["analyze", "design", "develop", "implement", "evaluate"]
DONE = {"done", "completed", "complete", "delivered"}
STUCK = {"stuck", "on hold", "blocked"}
NOT_STARTED = {"", "not started", "to do", "todo"}


def norm(s):
    return str(s or "").strip().lower()


def _dates(rows):
    s = [r["start"] for r in rows if r.get("start")]
    e = [r["end"] for r in rows if r.get("end")]
    return (min(s) if s else None), (max(e) if e else None)


def rollup(group_name, tasks, portco, skip_groups=()):
    """
    tasks: [{name, status, owner, priority, start, end, files:[url], subitems:[{name,status,owner,start,end}]}]
    returns dict for the Projects sheet (without ProjectID) + list of file urls, or None if skipped.
    """
    if norm(group_name) in {norm(g) for g in skip_groups} or not tasks:
        return None
    # each task counts through its subitems; a task with no subitems counts as one step
    steps = [s for t in tasks for s in (t.get("subitems") or [t])]
    statuses = [norm(s.get("status")) for s in steps]
    total = len(steps)
    done = sum(1 for s in statuses if s in DONE)
    progress = round(done / total * 100) if total else 0

    task_status = {norm(t["name"]): norm(t.get("status")) for t in tasks}
    if any(s in STUCK for s in statuses + list(task_status.values())):
        status = "On Hold"
    elif done == total and total:
        status = "Delivered"
    elif all(s in NOT_STARTED for s in statuses):
        status = "Script Received"
    else:
        # first phase that is not finished decides the stage
        current = next((t for t in tasks if norm(t.get("status")) not in DONE), None)
        phase = norm(current["name"]) if current else ""
        if phase.startswith("implement") or phase.startswith("evaluate"):
            subs = {norm(s["name"]): norm(s.get("status")) for s in current.get("subitems", [])}
            in_feedback = any(k.startswith("feedback") and v not in NOT_STARTED for k, v in subs.items())
            status = "Portco Feedback" if in_feedback else "In Review"
        else:
            status = "In Development"

    start, end = _dates(steps + tasks)
    steps_for_people = [s for t in tasks for s in t.get("subitems", [])]
    owners = Counter()
    for t in tasks + steps_for_people:
        for o in re.split(r",\s*", t.get("owner") or ""):
            if o.strip():
                owners[o.strip()] += 1
    prios = Counter(t.get("priority") for t in tasks if t.get("priority"))
    reviews = sum(1 for s in steps if re.match(r"sme review", norm(s["name"])) and norm(s.get("status")) in DONE)
    names = " ".join(norm(s["name"]) for s in steps)
    ctype = "eLearning" if ("rise" in names or "synthesia" in names) else (
        "Guide / Document" if ("guide" in norm(group_name) or "word document" in names) else "eLearning")
    files = [u for t in tasks for u in t.get("files", [])]
    return {
        "PortcoName": portco, "ProjectName": group_name, "ContentType": ctype,
        "IDOwner": owners.most_common(1)[0][0] if owners else "Unassigned",
        "ScriptReceived": start, "StartDate": start, "DueDate": end,
        "Status": status, "Progress": progress, "FeedbackRounds": reviews,
        "Priority": prios.most_common(1)[0][0] if prios else "",
        "_files": files, "_doneDate": end if status == "Delivered" else None,
    }


def assign_ids(projects, existing, prefix="PRJ-", start=101):
    """Keep the ProjectID already used for the same portco + project name, so communications/files stay linked.
    Never gives two projects the same ID (e.g. the same course name on two portco boards)."""
    by_pair, by_name = {}, {}
    for r in existing:
        if r.get("ProjectID"):
            by_pair.setdefault((norm(r.get("PortcoName")), norm(r.get("ProjectName"))), r["ProjectID"])
            by_name.setdefault(norm(r.get("ProjectName")), []).append(r["ProjectID"])
    ids = [str(r.get("ProjectID")) for r in existing if r.get("ProjectID")]
    nums = [int(m.group(1)) for v in ids if (m := re.search(r"(\d+)$", v))]
    nxt = max(nums + [start - 1]) + 1
    used = set()
    for p in projects:
        pid = by_pair.get((norm(p.get("PortcoName")), norm(p["ProjectName"])))
        if not pid and not p.get("PortcoName"):
            pid = next((i for i in by_name.get(norm(p["ProjectName"]), []) if i not in used), None)
        if not pid or pid in used:
            pid = f"{prefix}{nxt}"
            nxt += 1
        used.add(pid)
        p["ProjectID"] = pid
    return projects


def file_rows(projects):
    rows = []
    for p in projects:
        for url in p.get("_files", []):
            name = url.rstrip("/").split("/")[-1]
            ext = name.lower().rsplit(".", 1)[-1] if "." in name else ""
            ftype = "Final Content" if ext in ("zip", "scorm") else "Document"
            rows.append([name, p["PortcoName"], p["ProjectID"], ftype, "v1", p.get("_doneDate") or p.get("DueDate"), url])
    return rows
