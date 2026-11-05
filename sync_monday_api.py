"""
Automatic Monday.com -> dashboard sync (READ ONLY - this script never changes anything in Monday).

Runs inside GitHub Actions. Reads each board listed in MONDAY_BOARDS with the token in the
MONDAY_API_TOKEN secret, then rebuilds data.xlsx with build_from_monday.py (same rules as the
daily refresh: each group = one project, "Pre-Project / Kickoff" skipped, progress = % of subitems
Done, ProjectIDs kept, Portcos / Communications / manual Files rows kept).

MONDAY_BOARDS (GitHub variable, optional) - one board per line:  board_id = Portco name
    18435079287 = ACC
If it is not set, the ACC OnDemand board is used.
"""
import http.client, json, os, subprocess, sys, tempfile, time, urllib.request, urllib.error

API = os.environ.get("MONDAY_API_URL", "https://api.monday.com/v2")
DEFAULT_BOARDS = "18435079287 = ACC"
ITEM_COLS = ["project_owner", "project_status", "project_timeline", "project_priority", "file_mkw4nbdc"]
SUB_COLS = ["multiple_person_mkxfbhpb", "color_mkxfypxq", "timerange_mkxfarcc"]
FIELDS = """cursor items { name group { title }
  column_values(ids: %s) { id text }
  subitems { name column_values(ids: %s) { id text } } }""" % (json.dumps(ITEM_COLS), json.dumps(SUB_COLS))
FIRST = "query ($id: [ID!]) { boards(ids: $id) { name items_page(limit: 100) { %s } } }" % FIELDS
NEXT = "query ($c: String!) { next_items_page(limit: 100, cursor: $c) { %s } }" % FIELDS


def gql(query, variables, token, tries=4):
    body = json.dumps({"query": query, "variables": variables}).encode()
    for attempt in range(tries):
        req = urllib.request.Request(API, data=body, method="POST", headers={
            "Authorization": token, "Content-Type": "application/json", "API-Version": "2025-04"})
        try:
            with urllib.request.urlopen(req, timeout=60) as r:
                out = json.loads(r.read().decode())
        except urllib.error.HTTPError as e:
            if e.code == 401:
                sys.exit("Monday refused the token (401). It may have been regenerated or revoked - "
                         "update the MONDAY_API_TOKEN secret.")
            if e.code in (429, 500, 502, 503, 504) and attempt < tries - 1:
                time.sleep(10 * (attempt + 1)); continue
            sys.exit(f"Monday API error {e.code}.")
        except (urllib.error.URLError, http.client.HTTPException, ConnectionError, TimeoutError) as e:
            if attempt < tries - 1:
                time.sleep(10 * (attempt + 1)); continue
            sys.exit(f"Could not reach Monday: {getattr(e, 'reason', e)}")
        if out.get("errors"):
            msg = "; ".join(str(x.get("message")) for x in out["errors"])
            if "complexity" in msg.lower() and attempt < tries - 1:
                time.sleep(30); continue
            sys.exit(f"Monday API error: {msg}")
        return out["data"]
    sys.exit("Monday API did not answer after several tries.")


def read_board(board_id, token):
    data = gql(FIRST, {"id": [str(board_id)]}, token)
    if not data.get("boards"):
        sys.exit(f"Board {board_id} not found, or this token's user is not a member of it.")
    board = data["boards"][0]
    page = board["items_page"]
    items = list(page["items"])
    while page.get("cursor"):
        page = gql(NEXT, {"c": page["cursor"]}, token)["next_items_page"]
        items += page["items"]
    return board["name"], items


def boards_config():
    raw = os.environ.get("MONDAY_BOARDS", "").strip() or DEFAULT_BOARDS
    out = []
    for line in raw.replace(";", "\n").splitlines():
        if "=" in line:
            bid, portco = (x.strip() for x in line.split("=", 1))
            if bid and portco:
                out.append((bid, portco))
    if not out:
        sys.exit("MONDAY_BOARDS is set but has no 'board_id = Portco' lines.")
    return out


def main():
    token = os.environ.get("MONDAY_API_TOKEN", "").strip()
    if not token:
        print("No MONDAY_API_TOKEN secret - skipping the Monday sync, publishing the current data.xlsx.")
        return
    summaries = []
    for bid, portco in boards_config():
        name, items = read_board(bid, token)
        print(f"Read board '{name}' ({len(items)} items) for {portco}.")
        with tempfile.TemporaryDirectory() as tmp:
            src = os.path.join(tmp, "board.json"); out = os.path.join(tmp, "new.xlsx")
            json.dump({"boards": [{"items_page": {"items": items}}]}, open(src, "w"))
            r = subprocess.run([sys.executable, "build_from_monday.py", src, "data.xlsx", out,
                                "--portco", portco, "--board-name", name], capture_output=True, text=True)
            if r.returncode != 0:
                sys.exit(f"Building {portco} failed - data.xlsx left unchanged.\n{r.stdout}\n{r.stderr}")
            summary = json.loads(r.stdout)
            os.replace(out, "data.xlsx")
        summaries.append(summary)

    lines = ["## Monday sync", "", "| Portco | Projects | Delivered | New | Changed | Removed | Overdue | No dates |",
             "|---|---|---|---|---|---|---|---|"]
    for s in summaries:
        lines.append(f"| {s['portco']} | {s['projects']} | {s['delivered']} | {len(s['new'])} | {len(s['changed'])} | "
                     f"{len(s['removed'])} | {', '.join(s['overdue']) or '-'} | {', '.join(s['no_dates']) or '-'} |")
    for s in summaries:
        for n in s["new"]:
            lines.append(f"- **New:** {n['project']} ({n['id']}) - {n['status']}, {n['progress']}%")
        for c in s["changed"]:
            parts = ", ".join(f"{k}: {v[0] or 'blank'} -> {v[1] or 'blank'}" for k, v in c["changes"].items())
            lines.append(f"- **Changed:** {c['project']} - {parts}")
        for r in s["removed"]:
            lines.append(f"- **Removed:** {r}")
    text = "\n".join(lines) + "\n"
    print(text)
    if os.environ.get("GITHUB_STEP_SUMMARY"):
        open(os.environ["GITHUB_STEP_SUMMARY"], "a").write(text)


if __name__ == "__main__":
    main()
