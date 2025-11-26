#!/usr/bin/env python3
# remove_docx_protection.py
# Remove "Restrict Editing" protection from a .docx by editing word/settings.xml
# Usage:
#   python remove_docx_protection.py "C:\path\to\protected.docx"
#
# Requirements: Python 3.x (no extra packages required)
# NOTE: Works only with .docx (OpenXML). Does NOT handle .doc or password-encrypted files.

import sys
import shutil
import zipfile
import tempfile
import os
import xml.etree.ElementTree as ET
from pathlib import Path

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
ET.register_namespace('w', W_NS)

def remove_protection_from_settings(settings_path):
    tree = ET.parse(settings_path)
    root = tree.getroot()
    removed = False

    # Remove any <w:documentProtection .../> elements
    for elem in root.findall(f"{{{W_NS}}}documentProtection"):
        root.remove(elem)
        removed = True

    # Remove other possible locking tags
    for tag in ("lockedParts", "lockedSections", "revisionProtection"):
        for elem in root.findall(f"{{{W_NS}}}{tag}"):
            root.remove(elem)
            removed = True

    if removed:
        tree.write(settings_path, encoding="utf-8", xml_declaration=True)
    return removed

def main():
    if len(sys.argv) < 2:
        print("Usage: python remove_docx_protection.py path/to/protected.docx")
        sys.exit(1)

    input_path = Path(sys.argv[1])
    if not input_path.exists():
        print("Error: file not found:", input_path)
        sys.exit(2)

    if input_path.suffix.lower() != ".docx":
        print("Error: this script works only on .docx files (OpenXML).")
        sys.exit(3)

    backup = input_path.with_suffix(input_path.suffix + ".bak.docx")
    shutil.copy2(input_path, backup)
    print(f"Backup created: {backup}")

    tmpdir = tempfile.mkdtemp(prefix="docx_edit_")
    try:
        with zipfile.ZipFile(input_path, 'r') as zin:
            zin.extractall(tmpdir)

        settings_rel = os.path.join(tmpdir, "word", "settings.xml")
        if not os.path.exists(settings_rel):
            print("No word/settings.xml found. This doesn't look like a standard .docx or it's missing settings.xml.")
            print("Restoring backup and exiting.")
            shutil.copy2(backup, input_path)
            sys.exit(4)

        changed = remove_protection_from_settings(settings_rel)
        if not changed:
            print("No protection tags found in settings.xml. Nothing changed.")
            print("You may still be protected by another mechanism or the file uses different protection.")
            print("Backup left at:", backup)
            sys.exit(5)

        # Repack into new .docx (overwrite original)
        new_path = input_path.with_name(input_path.stem + ".unlocked.docx")
        with zipfile.ZipFile(new_path, 'w', zipfile.ZIP_DEFLATED) as zout:
            for foldername, subfolders, filenames in os.walk(tmpdir):
                for filename in filenames:
                    filepath = os.path.join(foldername, filename)
                    arcname = os.path.relpath(filepath, tmpdir)
                    zout.write(filepath, arcname)

        print("Done. Unlocked file written to:", new_path)
        print("Original file preserved as backup:", backup)
    finally:
        try:
            shutil.rmtree(tmpdir)
        except Exception:
            pass

if __name__ == "__main__":
    main()
