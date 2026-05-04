#!/usr/bin/env python3
"""Generate a zip-based update package from recent git commits."""
import sys
import os
import subprocess
import zipfile
from pathlib import Path


def main():
    if len(sys.argv) < 3:
        print("Употреба: python generate_update.py <root_dir> <output_zip> [commits]")
        sys.exit(1)

    root_dir = sys.argv[1]
    output_zip = sys.argv[2]
    commits = int(sys.argv[3]) if len(sys.argv) > 3 else 1

    def git(*args):
        return subprocess.run(
            ["git", "-C", root_dir] + list(args),
            capture_output=True, text=True
        )

    # Changed/added files
    r = git("diff", f"HEAD~{commits}", "HEAD", "--name-only", "--diff-filter=ACM")
    if r.returncode != 0:
        print(f"[ГРЕШКА] git diff: {r.stderr.strip()}")
        sys.exit(1)
    changed = [f.strip() for f in r.stdout.splitlines() if f.strip()]

    # Deleted files
    r_del = git("diff", f"HEAD~{commits}", "HEAD", "--name-only", "--diff-filter=D")
    deleted = [f.strip() for f in r_del.stdout.splitlines() if f.strip()]

    if not changed and not deleted:
        print("[ГРЕШКА] Няма промени в последните", commits, "commit(s).")
        sys.exit(1)

    with zipfile.ZipFile(output_zip, "w", zipfile.ZIP_DEFLATED) as zf:
        for rel_path in changed:
            abs_path = Path(root_dir) / rel_path
            if abs_path.exists():
                zf.write(abs_path, rel_path)
                print(f"  [+] {rel_path}")
            else:
                print(f"  [!] Не е намерен: {rel_path}")

        if deleted:
            zf.writestr("_deleted.txt", "\n".join(deleted))
            for d in deleted:
                print(f"  [-] {d}")

    total = len(changed) + len(deleted)
    print(f"\nГотово: {total} файл(а) в {Path(output_zip).name}")


if __name__ == "__main__":
    main()
