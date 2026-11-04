"""
PortCo Learning Hub - project health checks.
Each class tests one script or file. A check is SKIPPED (not failed) when that file is not in the repo.
Run locally:  pip install pytest openpyxl cryptography  &&  pytest -v tests
"""
import base64, datetime as dt, hashlib, json, os, re, shutil, subprocess, sys
import pytest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
STATUSES = {"Script Received", "In Development", "In Review", "Portco Feedback", "Delivered", "On Hold"}
SHEETS = {
    "Projects": ["ProjectID", "PortcoName", "ProjectName", "ContentType", "IDOwner", "ScriptReceived",
                 "StartDate", "DueDate", "Status", "Progress", "FeedbackRounds", "Priority"],
    "Portcos": ["PortcoID", "PortcoName"],
    "Communications": ["Date", "PortcoName", "ProjectID", "Type", "Direction", "Subject"],
    "Files": ["FileName", "PortcoName", "ProjectID", "FileType"],
}


def path(name):
    return os.path.join(ROOT, name)


def need(name):
    if not os.path.exists(path(name)):
        pytest.skip(f"{name} is not in the repo")


def read(name):
    with open(path(name), encoding="utf-8") as fh:
        return fh.read()


def sheet_rows(wb, name):
    ws = wb[name]
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return [], []
    head = [str(h or "").strip() for h in rows[0]]
    return head, [dict(zip(head, r)) for r in rows[1:] if any(v not in (None, "") for v in r)]


def as_date(v):
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    m = re.search(r"\d{4}-\d{2}-\d{2}", str(v or ""))
    return dt.date.fromisoformat(m.group(0)) if m else None


# --------------------------------------------------------------------------- data.xlsx
class TestDataFile:
    @pytest.fixture(scope="class")
    def wb(self):
        need("data.xlsx")
        from openpyxl import load_workbook
        return load_workbook(path("data.xlsx"), data_only=True)

    def test_required_sheets_exist(self, wb):
        missing = [s for s in SHEETS if s not in wb.sheetnames]
        assert not missing, f"Missing sheets: {missing}"

    def test_required_columns_exist(self, wb):
        problems = []
        for sheet, cols in SHEETS.items():
            if sheet in wb.sheetnames:
                head, _ = sheet_rows(wb, sheet)
                gone = [c for c in cols if c not in head]
                if gone:
                    problems.append(f"{sheet}: {gone}")
        assert not problems, "Missing columns -> " + "; ".join(problems)

    def test_has_projects(self, wb):
        _, rows = sheet_rows(wb, "Projects")
        assert rows, "The Projects sheet is empty"

    def test_project_ids_unique_and_well_formed(self, wb):
        _, rows = sheet_rows(wb, "Projects")
        ids = [str(r.get("ProjectID") or "").strip() for r in rows]
        bad = [i or "(blank)" for i in ids if not re.fullmatch(r"PRJ-\d+", i)]
        dupes = sorted({i for i in ids if ids.count(i) > 1})
        assert not bad, f"Badly formed ProjectIDs: {bad}"
        assert not dupes, f"Duplicate ProjectIDs: {dupes}"

    def test_status_values_are_known(self, wb):
        _, rows = sheet_rows(wb, "Projects")
        bad = [f'{r.get("ProjectID")}: {r.get("Status")!r}' for r in rows if str(r.get("Status") or "").strip() not in STATUSES]
        assert not bad, f"Unknown status values: {bad}"

    def test_progress_between_0_and_100(self, wb):
        _, rows = sheet_rows(wb, "Projects")
        bad = []
        for r in rows:
            try:
                ok = 0 <= float(r.get("Progress") or 0) <= 100
            except (TypeError, ValueError):
                ok = False
            if not ok:
                bad.append(f'{r.get("ProjectID")}: {r.get("Progress")!r}')
        assert not bad, f"Progress outside 0-100: {bad}"

    def test_delivered_projects_are_100_percent(self, wb):
        _, rows = sheet_rows(wb, "Projects")
        bad = [r.get("ProjectID") for r in rows if r.get("Status") == "Delivered" and float(r.get("Progress") or 0) < 100]
        assert not bad, f"Delivered but under 100%: {bad}"

    def test_start_date_not_after_due_date(self, wb):
        _, rows = sheet_rows(wb, "Projects")
        bad = []
        for r in rows:
            s, e = as_date(r.get("StartDate")), as_date(r.get("DueDate"))
            if s and e and s > e:
                bad.append(f'{r.get("ProjectID")}: {s} > {e}')
        assert not bad, f"Start date after due date: {bad}"

    def test_links_point_to_real_projects(self, wb):
        _, projects = sheet_rows(wb, "Projects")
        ids = {str(r.get("ProjectID") or "").strip() for r in projects}
        bad = []
        for sheet in ("Communications", "Files"):
            if sheet in wb.sheetnames:
                _, rows = sheet_rows(wb, sheet)
                bad += [f"{sheet}: {r.get('ProjectID')}" for r in rows
                        if r.get("ProjectID") and str(r["ProjectID"]).strip() not in ids]
        assert not bad, f"Rows pointing to a ProjectID that does not exist: {bad}"


# --------------------------------------------------------------------------- project_rollup.py
def step(name, status, start=None, end=None, owner="Arvin DSouza"):
    d = lambda x: dt.datetime.fromisoformat(x) if x else None
    return {"name": name, "status": status, "owner": owner, "start": d(start), "end": d(end)}


def task(name, status, subs=(), priority="Medium"):
    return {"name": name, "status": status, "owner": "Arvin DSouza", "priority": priority,
            "start": None, "end": None, "files": [], "subitems": list(subs)}


class TestProjectRollup:
    @pytest.fixture(scope="class")
    def pr(self):
        need("project_rollup.py")
        import project_rollup
        return project_rollup

    def test_all_steps_done_is_delivered_100(self, pr):
        p = pr.rollup("Course A", [task("Analyze", "Done", [step("Gather", "Done", "2026-01-01", "2026-01-05")]),
                                   task("Develop", "Done", [step("Build in Articulate Rise", "Done", "2026-01-06", "2026-01-20")])], "ACC")
        assert p["Status"] == "Delivered" and p["Progress"] == 100

    def test_nothing_started_is_script_received(self, pr):
        p = pr.rollup("Course B", [task("Analyze", "Not Started", [step("Gather", "Not Started"), step("Gap", "Not Started")])], "ACC")
        assert p["Status"] == "Script Received" and p["Progress"] == 0

    def test_progress_is_percent_of_steps_done(self, pr):
        p = pr.rollup("Course C", [task("Analyze", "Working on it", [step("a", "Done"), step("b", "Done"), step("c", "Working on it"), step("d", "")])], "ACC")
        assert p["Progress"] == 50 and p["Status"] == "In Development"

    def test_stuck_step_puts_project_on_hold(self, pr):
        p = pr.rollup("Course D", [task("Design", "Working on it", [step("a", "Done"), step("b", "Stuck")])], "ACC")
        assert p["Status"] == "On Hold"

    def test_feedback_started_in_implement_is_portco_feedback(self, pr):
        p = pr.rollup("Course E", [task("Develop", "Done", [step("Build", "Done")]),
                                   task("Implement", "In Progress", [step("SME Review(R1)", "Done"), step("Feedback", "In Progress")])], "ACC")
        assert p["Status"] == "Portco Feedback" and p["FeedbackRounds"] == 1

    def test_implement_without_feedback_is_in_review(self, pr):
        p = pr.rollup("Course F", [task("Develop", "Done", [step("Build", "Done")]),
                                   task("Implement", "Not Started", [step("Course Pilot Testing", "In Progress"), step("Feedback", "Not Started")])], "ACC")
        assert p["Status"] == "In Review"

    def test_dates_span_earliest_start_to_latest_end(self, pr):
        p = pr.rollup("Course G", [task("Analyze", "Done", [step("a", "Done", "2026-03-02", "2026-03-04")]),
                                   task("Develop", "Working on it", [step("b", "Working on it", "2026-03-05", "2026-03-30")])], "ACC")
        assert p["StartDate"].date() == dt.date(2026, 3, 2) and p["DueDate"].date() == dt.date(2026, 3, 30)

    def test_kickoff_group_is_skipped(self, pr):
        assert pr.rollup("Pre-Project / Kickoff", [task("Kickoff call", "Done")], "ACC", skip_groups=["Pre-Project / Kickoff"]) is None

    def test_existing_project_keeps_its_id(self, pr):
        out = pr.assign_ids([{"PortcoName": "ACC", "ProjectName": "Safety Plan"}, {"PortcoName": "ACC", "ProjectName": "Brand New"}],
                            [{"ProjectID": "PRJ-105", "PortcoName": "ACC", "ProjectName": "Safety Plan"},
                             {"ProjectID": "PRJ-110", "PortcoName": "ACC", "ProjectName": "Other"}])
        assert out[0]["ProjectID"] == "PRJ-105" and out[1]["ProjectID"] == "PRJ-111"

    def test_same_name_on_two_portcos_gets_two_ids(self, pr):
        out = pr.assign_ids([{"PortcoName": "ACC", "ProjectName": "Onboarding"}, {"PortcoName": "BxAx", "ProjectName": "Onboarding"}], [])
        assert out[0]["ProjectID"] != out[1]["ProjectID"]


# --------------------------------------------------------------------------- build_secure.py
class TestBuildSecure:
    @pytest.fixture(scope="class")
    def site(self, tmp_path_factory):
        need("build_secure.py")
        need("data.xlsx")
        from openpyxl import load_workbook
        work = tmp_path_factory.mktemp("secure")
        for f in ["build_secure.py", "data.xlsx", "index.html", "app.js", "style.css"]:
            if os.path.exists(path(f)):
                shutil.copy(path(f), work)
        wb = load_workbook(path("data.xlsx"), data_only=True)
        names = []
        for sheet in ("Portcos", "Projects"):
            if sheet in wb.sheetnames:
                for r in sheet_rows(wb, sheet)[1]:
                    n = str(r.get("PortcoName") or "").strip()
                    if n and n.lower() not in {x.lower() for x in names}:
                        names.append(n)
        pins = {n: str(1000 + i * 37) for i, n in enumerate(names)}
        access = "\n".join(f"{n} = {p}" for n, p in pins.items()) + "\nADMIN = test-admin-password"
        res = subprocess.run([sys.executable, "build_secure.py"], cwd=work, env={**os.environ, "PORTCO_ACCESS": access},
                             capture_output=True, text=True, timeout=600)
        return {"dir": work / "_site", "pins": pins, "res": res}

    def test_build_runs_without_errors(self, site):
        r = site["res"]
        assert r.returncode == 0, f"build_secure.py failed:\n{r.stdout}\n{r.stderr}"

    def test_raw_excel_is_never_published(self, site):
        found = [str(p) for p in site["dir"].rglob("*") if p.suffix.lower() in (".xlsx", ".xls", ".csv")]
        assert not found, f"Raw data published: {found}"

    def test_every_portco_has_a_login(self, site):
        listing = json.loads((site["dir"] / "portcos.json").read_text())
        assert {p["name"] for p in listing} == set(site["pins"]), "Some portcos have no login file"

    def _decrypt(self, blob, secret, label):
        from cryptography.hazmat.primitives.ciphers.aead import AESGCM
        from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
        from cryptography.hazmat.primitives import hashes
        salt = base64.b64decode(blob["salt"])
        key = PBKDF2HMAC(algorithm=hashes.SHA256(), length=32, salt=salt, iterations=blob["iter"]).derive(secret.encode())
        return json.loads(AESGCM(key).decrypt(base64.b64decode(blob["iv"]), base64.b64decode(blob["ct"]), None))

    def test_portco_data_opens_only_with_its_own_pin(self, site):
        listing = json.loads((site["dir"] / "portcos.json").read_text())
        for p in listing:
            blob = json.loads((site["dir"] / "data" / p["file"]).read_text())
            data = self._decrypt(blob, site["pins"][p["name"]], p["id"])
            others = {str(r.get("PortcoName")).strip().lower() for r in data.get("Projects", [])} - {p["name"].lower()}
            assert not others, f"{p['name']} file contains other portcos' projects: {others}"
            with pytest.raises(Exception):
                self._decrypt(blob, "0000" if site["pins"][p["name"]] != "0000" else "9999", p["id"])

    def test_admin_file_opens_with_admin_password(self, site):
        blob = json.loads((site["dir"] / "data" / "admin.json").read_text())
        assert self._decrypt(blob, "test-admin-password", "admin").get("Projects"), "Admin file has no projects"


# --------------------------------------------------------------------------- build_from_monday.py
class TestMondayRefresh:
    def _board(self):
        sub = lambda n, s, tl: {"name": n, "column_values": [{"id": "multiple_person_mkxfbhpb", "text": "Arvin DSouza"},
                                                              {"id": "color_mkxfypxq", "text": s}, {"id": "timerange_mkxfarcc", "text": tl}]}
        item = lambda n, g, s, subs: {"name": n, "group": {"title": g}, "subitems": subs, "column_values": [
            {"id": "project_owner", "text": "Arvin DSouza"}, {"id": "project_status", "text": s},
            {"id": "project_timeline", "text": ""}, {"id": "project_priority", "text": "High"}, {"id": "file_mkw4nbdc", "text": ""}]}
        return {"boards": [{"items_page": {"items": [
            item("Kickoff call", "Pre-Project / Kickoff", "Done", []),
            item("Analyze", "Test Course", "Done", [sub("Gather", "Done", "2026-01-01 - 2026-01-02")]),
            item("Develop", "Test Course", "Working on it", [sub("Build", "Working on it", "2026-01-03 - 2026-01-09")]),
        ]}}]}

    def test_builds_file_and_skips_kickoff(self, tmp_path):
        for f in ("build_from_monday.py", "project_rollup.py", "monday_sync.py", "data.xlsx"):
            need(f)
        (tmp_path / "board.json").write_text(json.dumps(self._board()))
        out = tmp_path / "new.xlsx"
        r = subprocess.run([sys.executable, path("build_from_monday.py"), str(tmp_path / "board.json"), path("data.xlsx"), str(out),
                            "--portco", "TESTCO"], capture_output=True, text=True, cwd=ROOT, timeout=300)
        assert r.returncode == 0, r.stderr
        from openpyxl import load_workbook
        _, rows = sheet_rows(load_workbook(out, data_only=True), "Projects")
        mine = [x for x in rows if x.get("PortcoName") == "TESTCO"]
        assert [x["ProjectName"] for x in mine] == ["Test Course"]
        assert mine[0]["Progress"] == 50


# --------------------------------------------------------------------------- website files
class TestDashboardPage:
    def test_index_links_app_and_style(self):
        need("index.html")
        h = read("index.html")
        assert 'src="app.js"' in h and 'href="style.css"' in h, "index.html does not load app.js and style.css from the root"

    def test_app_js_has_no_syntax_errors(self):
        need("app.js")
        if not shutil.which("node"):
            pytest.skip("Node.js not available")
        r = subprocess.run(["node", "--check", path("app.js")], capture_output=True, text=True)
        assert r.returncode == 0, r.stderr

    def test_every_element_used_by_app_js_exists(self):
        need("app.js")
        need("index.html")
        ids = set(re.findall(r'\$\("([A-Za-z0-9_-]+)"\)', read("app.js")))
        html = read("index.html")
        missing = sorted(i for i in ids if f'id="{i}"' not in html)
        assert not missing, f"app.js uses elements missing from index.html: {missing}"

    def test_no_secrets_in_published_files(self):
        hits = []
        for f in ("index.html", "app.js", "style.css"):
            if os.path.exists(path(f)):
                t = read(f)
                if re.search(r"(PORTCO_ACCESS\s*=|ADMIN\s*=\s*\S+|eyJhbGciOi)", t):
                    hits.append(f)
        assert not hits, f"Possible secret found in: {hits}"
