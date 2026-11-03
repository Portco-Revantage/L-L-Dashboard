"""
PortCo Learning Hub - secure build
----------------------------------
Builds the published site in _site/ so that:
  * the raw data.xlsx is NEVER published;
  * each portco gets its own data file, encrypted with its own PIN;
  * the admin file (all portcos) is encrypted with the admin password.

PINs and the admin password come from the GitHub secret PORTCO_ACCESS, one per line:
    ACC = 4821
    Northwind Logistics = 7310
    ADMIN = a-long-admin-password
Rules: portco PINs are digits only, 1-4 digits, and must be unique.
"""
import base64
import datetime as dt
import hashlib
import json
import os
import re
import secrets
import shutil
import sys

from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.ciphers.aead import AESGCM
from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
from openpyxl import load_workbook

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "_site")
ITERATIONS = 600_000
SITE_FILES = ["index.html", "app.js", "style.css"]
SHEETS = ["Portcos", "Projects", "Communications", "Files", "SyncInfo"]


def fail(msg):
    print(f"ERROR: {msg}")
    sys.exit(1)


def read_access():
    raw = os.environ.get("PORTCO_ACCESS", "")
    if not raw.strip():
        fail("The PORTCO_ACCESS secret is missing. Nothing was published, so no data is exposed.")
    access, admin = {}, None
    for n, line in enumerate(raw.splitlines(), 1):
        line = line.strip()
        if not line or line.startswith("#"):
            continue
        if "=" not in line:
            fail(f"PORTCO_ACCESS line {n} needs the format  Portco Name = 1234")
        name, secret = (p.strip() for p in line.rsplit("=", 1))
        if name.upper() == "ADMIN":
            admin = secret
            continue
        if not re.fullmatch(r"\d{1,4}", secret):
            fail(f"PIN for '{name}' must be digits only, maximum 4 digits.")
        access[name.lower()] = (name, secret)
    if not admin:
        fail("PORTCO_ACCESS needs an  ADMIN = <password>  line.")
    if len(admin) < 8:
        print("WARNING: the admin password is shorter than 8 characters. A longer one is strongly recommended.")
    pins = [p for _, p in access.values()]
    dupes = sorted({p for p in pins if pins.count(p) > 1})
    if dupes:
        fail(f"Each portco needs a unique PIN. Repeated PIN(s): {', '.join(dupes)}")
    return access, admin


def cell(v):
    if isinstance(v, (dt.datetime, dt.date)):
        return v.isoformat()
    return "" if v is None else v


def read_data():
    path = os.path.join(HERE, "data.xlsx")
    if not os.path.exists(path):
        fail("data.xlsx not found in the repository.")
    wb = load_workbook(path, data_only=True)
    out = {}
    for name in SHEETS:
        ws = next((wb[s] for s in wb.sheetnames if s.lower() == name.lower()), None)
        rows = list(ws.iter_rows(values_only=True)) if ws else []
        if not rows:
            out[name] = []
            continue
        hdr = [str(h).strip() if h is not None else "" for h in rows[0]]
        out[name] = [{h: cell(v) for h, v in zip(hdr, r) if h}
                     for r in rows[1:] if any(v not in (None, "") for v in r)]
    return out


def slug(name):
    s = re.sub(r"[^a-z0-9]+", "-", name.lower()).strip("-") or "portco"
    return f"{s}-{hashlib.sha256(name.encode()).hexdigest()[:6]}"


def encrypt(payload, secret, salt_label):
    # Stable salt per portco: the key stays the same between builds while the PIN is unchanged,
    # so signed-in users are not logged out every time Monday syncs.
    salt = hashlib.sha256(f"portco-learning-hub:{salt_label}".encode()).digest()[:16]
    key = PBKDF2HMAC(algorithm=hashes.SHA256(), length=32, salt=salt, iterations=ITERATIONS).derive(secret.encode())
    iv = secrets.token_bytes(12)
    ct = AESGCM(key).encrypt(iv, json.dumps(payload, default=str).encode(), None)
    b = lambda x: base64.b64encode(x).decode()
    return {"v": 1, "iter": ITERATIONS, "salt": b(salt), "iv": b(iv), "ct": b(ct)}


def main():
    access, admin = read_access()
    data = read_data()
    norm = lambda s: str(s or "").strip().lower()

    names = []
    for r in data["Portcos"] + data["Projects"]:
        n = str(r.get("PortcoName") or "").strip()
        if n and norm(n) not in {norm(x) for x in names}:
            names.append(n)
    details = {norm(r.get("PortcoName")): r for r in data["Portcos"]}
    sync_public = [r for r in data["SyncInfo"] if r.get("Key") in ("Source", "LastSync")]

    shutil.rmtree(OUT, ignore_errors=True)
    os.makedirs(os.path.join(OUT, "data"))
    for f in SITE_FILES:
        shutil.copy(os.path.join(HERE, f), OUT)

    listing, missing = [], []
    for name in names:
        if norm(name) not in access:
            missing.append(name)
            continue
        pin = access[norm(name)][1]
        mine = lambda rows: [r for r in rows if norm(r.get("PortcoName")) == norm(name)]
        payload = {"Portcos": mine(data["Portcos"]) or [{"PortcoName": name}], "Projects": mine(data["Projects"]),
                   "Communications": mine(data["Communications"]), "Files": mine(data["Files"]), "SyncInfo": sync_public}
        sid = slug(name)
        with open(os.path.join(OUT, "data", f"{sid}.json"), "w") as fh:
            json.dump(encrypt(payload, pin, sid), fh)
        listing.append({"id": sid, "name": name, "industry": str(details.get(norm(name), {}).get("Industry") or ""),
                        "file": f"{sid}.json"})

    with open(os.path.join(OUT, "data", "admin.json"), "w") as fh:
        json.dump(encrypt(data, admin, "admin"), fh)
    listing.sort(key=lambda p: p["name"].lower())
    with open(os.path.join(OUT, "portcos.json"), "w") as fh:
        json.dump(listing, fh, indent=1)
    with open(os.path.join(OUT, ".nojekyll"), "w") as fh:
        fh.write("")

    unused = [v[0] for k, v in access.items() if k not in {norm(n) for n in names}]
    print(f"Published {len(listing)} portco logins + admin.")
    if missing:
        print(f"WARNING: no PIN set for: {', '.join(missing)} - they are hidden from the login list (admin still sees them).")
    if unused:
        print(f"NOTE: PINs set for names not found in data.xlsx (check spelling): {', '.join(unused)}")


if __name__ == "__main__":
    main()
