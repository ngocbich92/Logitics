"""Backup hàng ngày: pg_dump (nén) + file đính kèm -> thư mục backup (là một repo git riêng, Private) -> commit và push lên GitHub.

Chạy:  python tools/backup.py <thư_mục_repo_backup> [--no-push] [--keep 14]
Cần: PostgreSQL đã cấu hình trong data/db_config.json, git đã đăng nhập GitHub (Git Credential Manager), repo backup đã `git remote add origin ...`."""
import argparse
import datetime as dt
import shutil
import subprocess
import sys
import zipfile
from pathlib import Path
from urllib.parse import unquote, urlparse

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
import pgcompat  # noqa: E402


def find_pg_dump():
    found = shutil.which("pg_dump")
    if found:
        return found
    for p in sorted(Path("C:/Program Files/PostgreSQL").glob("*/bin/pg_dump.exe"), reverse=True):
        return str(p)
    sys.exit("Không tìm thấy pg_dump")


def run(cmd, cwd=None, env=None):
    r = subprocess.run(cmd, cwd=cwd, env=env, capture_output=True, text=True, encoding="utf-8", errors="replace")
    if r.returncode:
        sys.exit(f"Lỗi: {' '.join(str(c) for c in cmd[:3])}...\n{r.stderr.strip()[:500]}")
    return r.stdout


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("repo")
    ap.add_argument("--no-push", action="store_true")
    ap.add_argument("--keep", type=int, default=14)
    a = ap.parse_args()
    dsn = pgcompat.dsn()
    if not dsn:
        sys.exit("Chưa cấu hình PostgreSQL (data/db_config.json)")
    u = urlparse(dsn)
    repo = Path(a.repo)
    repo.mkdir(parents=True, exist_ok=True)
    stamp = dt.datetime.now().strftime("%Y-%m-%d_%H%M")
    env = {**__import__("os").environ, "PGPASSWORD": unquote(u.password or "")}
    out = repo / f"nvl_{stamp}.dump"
    run([find_pg_dump(), "-Fc", "-Z", "9", "-h", u.hostname or "localhost", "-p", str(u.port or 5432), "-U", unquote(u.username or ""), "-d", u.path.lstrip("/"), "-f", str(out)], env=env)
    up = ROOT / "data" / "uploads"
    if up.is_dir() and any(up.iterdir()):
        with zipfile.ZipFile(repo / f"uploads_{stamp}.zip", "w", zipfile.ZIP_DEFLATED) as z:
            for f in up.rglob("*"):
                if f.is_file():
                    z.write(f, f.relative_to(up))
    for pat in ("nvl_*.dump", "uploads_*.zip"):  # chỉ giữ N bản gần nhất
        for old in sorted(repo.glob(pat))[:-a.keep]:
            old.unlink()
    print(f"Đã tạo {out.name} ({out.stat().st_size / 1e6:.1f} MB)")
    if a.no_push:
        return
    if not (repo / ".git").is_dir():
        run(["git", "init", "-b", "main"], cwd=repo)
    run(["git", "add", "-A"], cwd=repo)
    run(["git", "-c", "user.name=nvl-backup", "-c", "user.email=backup@localhost", "commit", "-q", "-m", f"backup {stamp}"], cwd=repo)
    run(["git", "push", "-u", "origin", "main"], cwd=repo)
    print("Đã đẩy lên GitHub")


if __name__ == "__main__":
    main()
