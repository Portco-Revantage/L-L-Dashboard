# PortCo login with PINs + Admin login

1. A visitor opens the dashboard link and sees the list of portcos.
2. They pick their company and a pop-up asks for its PIN (digits only, up to 4).
3. With the right PIN they see **only their own** projects, communications and files.
4. **Admin login** (bottom of the list) uses a separate password and shows every portco.

## How the data is protected
- `data.xlsx` is **no longer published** on the website.
- When GitHub publishes, `build_secure.py` creates one data file per portco, **encrypted with that portco's PIN**
  (AES-256). A portco's PIN can't open another portco's file. The admin file is encrypted with the admin password.
- PINs live only in a GitHub **secret** (`PORTCO_ACCESS`), never in the code.
- After 5 wrong attempts the pop-up locks for 30 seconds.
- Signing in lasts until the browser tab is closed or **Log out** is clicked.

## Set the PINs (GitHub secret `PORTCO_ACCESS`)
One line per portco, using the exact PortcoName from data.xlsx / Monday, plus one ADMIN line:

    ACC = 4821
    Northwind Logistics = 7310
    ADMIN = Choose-A-Long-Admin-Password-2026

Rules (the publish stops with a clear message if one is broken, so nothing is exposed):
- Portco PINs: digits only, 1-4 digits, each one unique.
- ADMIN: any characters. Use 10+ characters; it protects all portcos' data.
- A portco without a PIN is hidden from the list (admin still sees it).

To change a PIN: Settings -> Secrets and variables -> Actions -> PORTCO_ACCESS -> Update,
then Actions -> "Monday sync & publish dashboard" -> Run workflow.

## Important limits
- **A 4-digit PIN has only 10,000 combinations.** It stops portcos casually opening each other's
  dashboards, but someone technical who downloads an encrypted file could guess the PIN offline.
  It is not enterprise-grade security.
- **Make the repository private.** In a public repo, anyone can still open data.xlsx on github.com.
  GitHub Pages on a private repo needs a paid GitHub plan (Pro / Team / Enterprise) - check with IT.
- For full production use with external portcos, move to Microsoft sign-in
  (e.g. Power BI with row-level security or SharePoint permissions).
