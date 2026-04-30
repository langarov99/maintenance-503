#!/usr/bin/env python3
"""Apply a git format-patch file without requiring git."""
import sys
import os
import re


def apply_patch(patch_path, root_dir):
    with open(patch_path, "r", encoding="utf-8", errors="replace") as f:
        content = f.read()

    file_diffs = re.split(r"(?=^diff --git )", content, flags=re.MULTILINE)

    results = []
    for block in file_diffs:
        if not block.startswith("diff --git "):
            continue
        result = _apply_file_diff(block, root_dir)
        if result:
            results.append(result)
    return results


def _apply_file_diff(block, root_dir):
    lines = block.splitlines(keepends=True)
    i = 0

    m = re.match(r"diff --git a/(.*?) b/(.*)", lines[0].rstrip())
    if not m:
        return None
    filepath = m.group(2)
    i = 1

    is_new_file = False
    is_deleted = False

    while i < len(lines):
        line = lines[i].rstrip()
        if line.startswith("new file mode"):
            is_new_file = True
        elif line.startswith("deleted file mode"):
            is_deleted = True
        elif line.startswith("--- ") or line.startswith("Binary files"):
            break
        i += 1

    if i >= len(lines):
        return None

    if lines[i].startswith("Binary files"):
        return None

    old_name = lines[i][4:].rstrip()
    i += 1
    if i >= len(lines) or not lines[i].startswith("+++ "):
        return None
    new_name = lines[i][4:].rstrip()
    i += 1

    abs_path = os.path.normpath(
        os.path.join(root_dir, filepath.replace("/", os.sep))
    )

    if is_deleted or new_name == "/dev/null":
        if os.path.exists(abs_path):
            os.remove(abs_path)
        return ("изтрит", filepath)

    if is_new_file or old_name == "/dev/null":
        file_lines = []
    else:
        try:
            with open(abs_path, "r", encoding="utf-8", errors="replace") as f:
                file_lines = f.readlines()
        except FileNotFoundError:
            file_lines = []

    result_lines = list(file_lines)
    delta = 0

    while i < len(lines):
        if not lines[i].startswith("@@ "):
            i += 1
            continue

        m = re.match(r"@@ -(\d+)(?:,(\d+))? \+(\d+)(?:,(\d+))? @@", lines[i])
        if not m:
            i += 1
            continue

        old_start = int(m.group(1)) - 1
        i += 1

        hunk_old = []
        hunk_new = []

        while i < len(lines):
            l = lines[i]
            if l.startswith("\\ No newline at end of file"):
                i += 1
                continue
            if l.startswith("@@ ") or l.startswith("diff --git "):
                break
            if l.startswith(" "):
                hunk_old.append(l[1:])
                hunk_new.append(l[1:])
            elif l.startswith("-"):
                hunk_old.append(l[1:])
            elif l.startswith("+"):
                hunk_new.append(l[1:])
            else:
                break
            i += 1

        pos = old_start + delta
        result_lines[pos : pos + len(hunk_old)] = hunk_new
        delta += len(hunk_new) - len(hunk_old)

    parent = os.path.dirname(abs_path)
    if parent:
        os.makedirs(parent, exist_ok=True)

    with open(abs_path, "w", encoding="utf-8", newline="") as f:
        f.writelines(result_lines)

    action = "нов" if is_new_file else "обновен"
    return (action, filepath)


def main():
    if len(sys.argv) != 3:
        print("Употреба: python apply_patch.py <patch_file> <root_dir>")
        sys.exit(1)

    patch_file = sys.argv[1]
    root_dir = sys.argv[2]

    if not os.path.isfile(patch_file):
        print(f"[ГРЕШКА] Patch файлът не е намерен: {patch_file}")
        sys.exit(1)

    print(f"Прилагане на: {os.path.basename(patch_file)}")
    print()

    results = apply_patch(patch_file, root_dir)

    if not results:
        print("[ГРЕШКА] Не са намерени промени в patch файла.")
        sys.exit(1)

    for action, fp in results:
        print(f"  [{action}] {fp}")

    print()
    print(f"Общо файлове: {len(results)}")
    sys.exit(0)


if __name__ == "__main__":
    main()
