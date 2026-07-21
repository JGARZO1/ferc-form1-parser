#!/usr/bin/env python3
"""Upload all text + binary files using gh api (requires gh auth login)."""
import base64, json, subprocess, sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
OWNER, REPO, BRANCH = "JGARZO1", "ferc-form1-parser", "main"

def gh_api(method, path, **kwargs):
    cmd = ["gh", "api", "--method", method, path]
    if "input" in kwargs:
        r = subprocess.run(cmd + ["--input", "-"], input=kwargs["input"].encode(), capture_output=True)
    else:
        r = subprocess.run(cmd, capture_output=True)
    if r.returncode != 0:
        raise SystemExit(r.stderr.decode() or r.stdout.decode())
    return json.loads(r.stdout) if r.stdout.strip() else {}

def put_file(rel: Path):
    data = rel.read_bytes()
    # get existing sha if any
    api_path = f"/repos/{OWNER}/{REPO}/contents/{rel.as_posix()}"
    try:
        existing = subprocess.run(
            ["gh", "api", api_path], capture_output=True, text=True
        )
        sha = json.loads(existing.stdout).get("sha") if existing.returncode == 0 else None
    except Exception:
        sha = None
    body = {
        "message": f"Add {rel.as_posix()}",
        "content": base64.b64encode(data).decode("ascii"),
        "branch": BRANCH,
    }
    if sha:
        body["sha"] = sha
    subprocess.run(
        ["gh", "api", "--method", "PUT", api_path, "--input", "-"],
        input=json.dumps(body).encode(),
        check=True,
    )
    print("OK", rel)

def main():
    skip = {".git", ".push_batches", "node_modules", "__pycache__"}
    files = []
    for p in ROOT.rglob("*"):
        if not p.is_file():
            continue
        if any(part in skip for part in p.parts):
            continue
        if p.name in {".DS_Store"}:
            continue
        rel = p.relative_to(ROOT)
        files.append(rel)
    files.sort()
    print(f"Uploading {len(files)} files...")
    for rel in files:
        put_file(ROOT / rel)
    print("Done.")

if __name__ == "__main__":
    main()
