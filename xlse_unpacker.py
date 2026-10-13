#!/usr/bin/env python3
"""
xlse_unpacker.py

Core engine and CLI tool for unpacking, extracting, and deconstructing Excel files
(.xlsx, .xlsm, .xlsb) into sibling folders with SHA-256 delta syncing.
"""

import argparse
import fnmatch
import hashlib
import os
import sys
import zipfile
from typing import Dict, List, Optional, Set, Tuple

EXCEL_EXTENSIONS = {".xlsx", ".xlsm", ".xlsb"}


def compute_sha256(data: bytes) -> str:
    """Compute SHA-256 checksum of raw byte content."""
    return hashlib.sha256(data).hexdigest()


def compute_file_sha256(file_path: str) -> str:
    """Compute SHA-256 checksum of an existing file on disk."""
    hasher = hashlib.sha256()
    with open(file_path, "rb") as f:
        while chunk := f.read(65536):
            hasher.update(chunk)
    return hasher.hexdigest()


def matches_patterns(name: str, patterns: List[str]) -> bool:
    """Check if name matches any of the given wildcard or substring patterns."""
    for pat in patterns:
        if fnmatch.fnmatch(name, pat) or pat.lower() in name.lower():
            return True
    return False


def find_excel_files(
    root_path: str,
    recurse: bool = False,
    filter_in: Optional[List[str]] = None,
    filter_out: Optional[List[str]] = None,
    dir_filter_in: Optional[List[str]] = None,
    dir_filter_out: Optional[List[str]] = None,
) -> List[str]:
    """
    Discover all Excel workbooks (.xlsx, .xlsm, .xlsb) in root_path according
    to recursion and filtering options.
    """
    matched_files = []
    root_path = os.path.abspath(root_path)

    if os.path.isfile(root_path):
        ext = os.path.splitext(root_path)[1].lower()
        if ext in EXCEL_EXTENSIONS:
            base = os.path.basename(root_path)
            if filter_in and not matches_patterns(base, filter_in):
                return []
            if filter_out and matches_patterns(base, filter_out):
                return []
            return [root_path]
        return []

    if not os.path.isdir(root_path):
        return []

    if not recurse:
        try:
            for entry in os.listdir(root_path):
                full_path = os.path.join(root_path, entry)
                if os.path.isfile(full_path):
                    ext = os.path.splitext(entry)[1].lower()
                    if ext in EXCEL_EXTENSIONS:
                        if filter_in and not matches_patterns(entry, filter_in):
                            continue
                        if filter_out and matches_patterns(entry, filter_out):
                            continue
                        matched_files.append(full_path)
        except OSError as e:
            print(f"[ERROR] Could not list directory {root_path}: {e}", file=sys.stderr)
        return sorted(matched_files)

    for dirpath, dirnames, filenames in os.walk(root_path):
        rel_dir = os.path.relpath(dirpath, root_path)

        if dir_filter_out and rel_dir != "." and matches_patterns(os.path.basename(dirpath), dir_filter_out):
            dirnames.clear()
            continue

        if dir_filter_in and rel_dir != "." and not matches_patterns(os.path.basename(dirpath), dir_filter_in):
            continue

        for filename in filenames:
            ext = os.path.splitext(filename)[1].lower()
            if ext in EXCEL_EXTENSIONS:
                if filter_in and not matches_patterns(filename, filter_in):
                    continue
                if filter_out and matches_patterns(filename, filter_out):
                    continue
                matched_files.append(os.path.join(dirpath, filename))

    return sorted(matched_files)


def unpack_excel_file(
    excel_path: str,
    output_dir: Optional[str] = None,
    dry_run: bool = False,
    quiet: bool = False,
) -> Tuple[int, int, int]:
    """
    Unpack an Excel file into its sibling folder.
    Returns:
        (created_count, updated_count, unchanged_count)
    """
    excel_path = os.path.abspath(excel_path)
    base_name = os.path.splitext(os.path.basename(excel_path))[0]
    parent_dir = os.path.dirname(excel_path)
    target_dir = os.path.abspath(output_dir) if output_dir else os.path.join(parent_dir, base_name)

    created = 0
    updated = 0
    unchanged = 0

    if not quiet:
        print(f"\n[UNPACK] Processing: {excel_path}")
        print(f"         Destination: {target_dir}")

    if not zipfile.is_zipfile(excel_path):
        print(f"[ERROR] {excel_path} is not a valid ZIP-based OpenXML / XLSB container.", file=sys.stderr)
        return (0, 0, 0)

    try:
        with zipfile.ZipFile(excel_path, "r") as zf:
            for item in zf.infolist():
                if item.is_dir():
                    continue

                rel_path = item.filename
                dest_file = os.path.join(target_dir, rel_path)
                content = zf.read(item)

                if os.path.exists(dest_file):
                    existing_hash = compute_file_sha256(dest_file)
                    new_hash = compute_sha256(content)
                    if existing_hash == new_hash:
                        unchanged += 1
                        continue
                    else:
                        updated += 1
                        action = "UPDATE"
                else:
                    created += 1
                    action = "CREATE"

                if not quiet:
                    print(f"  [{action}] {rel_path}")

                if not dry_run:
                    os.makedirs(os.path.dirname(dest_file), exist_ok=True)
                    with open(dest_file, "wb") as out_f:
                        out_f.write(content)

    except Exception as e:
        print(f"[ERROR] Failed to unpack {excel_path}: {e}", file=sys.stderr)
        return (created, updated, unchanged)

    if not quiet:
        print(f"  Summary: {created} created, {updated} updated, {unchanged} unchanged.")

    return (created, updated, unchanged)


def parse_args(argv: Optional[List[str]] = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Extract and unpack XLSX, XLSM, and XLSB files into sibling directories with delta sync."
    )
    parser.add_argument("-p", "--path", default=".", help="Root folder to search or specific file to unpack (default: .)")
    parser.add_argument("-r", "--recurse", action="store_true", help="Recursively traverse subdirectories")
    parser.add_argument("-o", "--output", help="Custom output directory (only recommended when unpacking a single file)")
    parser.add_argument("--filter-in", nargs="*", help="File patterns to include (e.g. '*report*', 'Budget')")
    parser.add_argument("--filter-out", nargs="*", help="File patterns to exclude (e.g. '*temp*', '~*')")
    parser.add_argument("--dir-filter-in", nargs="*", help="Subdirectory patterns to include during recursion")
    parser.add_argument("--dir-filter-out", nargs="*", help="Subdirectory patterns to exclude during recursion")
    parser.add_argument("--dry-run", action="store_true", help="Simulate actions without writing files to disk")
    parser.add_argument("-q", "--quiet", action="store_true", help="Suppress detailed per-file output")
    return parser.parse_args(argv)


def main(argv: Optional[List[str]] = None) -> int:
    args = parse_args(argv)
    files = find_excel_files(
        args.path,
        recurse=args.recurse,
        filter_in=args.filter_in,
        filter_out=args.filter_out,
        dir_filter_in=args.dir_filter_in,
        dir_filter_out=args.dir_filter_out,
    )

    if not files:
        if not args.quiet:
            print(f"[INFO] No Excel files (.xlsx, .xlsm, .xlsb) found in: {args.path}")
        return 0

    total_created, total_updated, total_unchanged = 0, 0, 0
    for f in files:
        c, u, uc = unpack_excel_file(
            f,
            output_dir=args.output,
            dry_run=args.dry_run,
            quiet=args.quiet,
        )
        total_created += c
        total_updated += u
        total_unchanged += uc

    if not args.quiet:
        print("\n" + "=" * 50)
        print(f"TOTAL: {len(files)} files processed")
        print(f"       {total_created} created, {total_updated} updated, {total_unchanged} unchanged.")
        print("=" * 50)

    return 0


if __name__ == "__main__":
    sys.exit(main())
