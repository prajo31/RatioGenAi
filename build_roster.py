"""
Build config.yaml (the login roster) from a plain roster.csv.

Why this script exists: config.yaml must contain *hashed* passwords, never
plain text. This script reads a simple CSV you fill in with real student
usernames/names/emails/passwords, hashes each password with bcrypt (via
streamlit-authenticator's Hasher), and writes config.yaml in the exact
structure app.py expects. Re-run it any time the roster changes.

Usage:
    1. Copy roster_template.csv to roster.csv (or edit roster.csv directly
       if this repo already has a starter one).
    2. Fill in one row per student: username,name,email,password
       - username: what the student types to log in (e.g. jdoe, or a
         student ID). Must be unique.
       - password: a plain-text starting password (students can't change
         it from within the app — reissue by editing the CSV and rerunning
         this script if someone needs a reset).
    3. Run:  python build_roster.py
    4. This produces config.yaml. Keep roster.csv and config.yaml OUT of
       version control / anywhere public — config.yaml contains hashed
       passwords, and roster.csv contains the plain-text originals.

For Streamlit Community Cloud deployment, you don't upload config.yaml at
all. Instead, open the app's Settings -> Secrets and paste the equivalent
TOML (this script also prints that block for you to copy).
"""
import csv
import secrets
import sys

import streamlit_authenticator as stauth
import yaml

ROSTER_CSV = "roster.csv"
CONFIG_YAML = "config.yaml"


def main():
    try:
        with open(ROSTER_CSV, newline="", encoding="utf-8") as f:
            rows = list(csv.DictReader(f))
    except FileNotFoundError:
        print(f"Couldn't find {ROSTER_CSV}. Create it with columns: username,name,email,password")
        sys.exit(1)

    if not rows:
        print(f"{ROSTER_CSV} has no rows. Add at least one student.")
        sys.exit(1)

    usernames = {}
    for row in rows:
        username = row["username"].strip()
        if not username:
            continue
        hashed = stauth.Hasher.hash(row["password"].strip())
        usernames[username] = {
            "name": row["name"].strip(),
            "email": row.get("email", "").strip(),
            "password": hashed,
        }

    config = {
        "credentials": {"usernames": usernames},
        "cookie": {
            "name": "ratio_tool_auth",
            # A random per-deployment secret used to sign the login cookie.
            # Regenerating this logs everyone out; that's fine.
            "key": secrets.token_hex(16),
            "expiry_days": 7,
        },
    }

    with open(CONFIG_YAML, "w", encoding="utf-8") as f:
        yaml.dump(config, f, default_flow_style=False)

    print(f"Wrote {CONFIG_YAML} with {len(usernames)} student login(s).")
    print()
    print("=" * 70)
    print("For Streamlit Community Cloud deployment instead, paste this into")
    print("the app's Settings -> Secrets box (do NOT commit config.yaml):")
    print("=" * 70)
    print()
    print("[cookie]")
    print(f'name = "{config["cookie"]["name"]}"')
    print(f'key = "{config["cookie"]["key"]}"')
    print(f'expiry_days = {config["cookie"]["expiry_days"]}')
    print()
    for username, info in usernames.items():
        print(f"[credentials.usernames.{username}]")
        print(f'name = "{info["name"]}"')
        print(f'email = "{info["email"]}"')
        print(f'password = "{info["password"]}"')
        print()


if __name__ == "__main__":
    main()
