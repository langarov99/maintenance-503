#!/usr/bin/env python3
"""Apply a zip-based update package."""
import sys
import zipfile
from pathlib import Path


def main():
    if len(sys.argv) != 3:
        print("Употреба: python apply_update.py <update_zip> <root_dir>")
        sys.exit(1)

    zip_path = sys.argv[1]
    root_dir = Path(sys.argv[2])

    if not Path(zip_path).is_file():
        print(f"[ГРЕШКА] Файлът не е намерен: {zip_path}")
        sys.exit(1)

    updated = 0
    deleted = 0

    with zipfile.ZipFile(zip_path, "r") as zf:
        names = zf.namelist()

        if "_deleted.txt" in names:
            for name in zf.read("_deleted.txt").decode("utf-8").splitlines():
                name = name.strip()
                if not name:
                    continue
                target = root_dir / name
                if target.exists():
                    target.unlink()
                    print(f"  [-] {name}")
                    deleted += 1

        for name in names:
            if name == "_deleted.txt":
                continue
            target = root_dir / name
            target.parent.mkdir(parents=True, exist_ok=True)
            target.write_bytes(zf.read(name))
            print(f"  [+] {name}")
            updated += 1

    total = updated + deleted
    print(f"\nГотово: {updated} обновен(и), {deleted} изтрит(и) — общо {total} файл(а)")


if __name__ == "__main__":
    main()
