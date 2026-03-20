#!/usr/bin/env python3
"""
Fix XML entity encoding in HWPX files.

Escapes unescaped & characters to &amp; in URLs and other content.
Common issue: URL query parameters like ?param1=value&param2=value.

Usage:
    python3 fix_xml.py <input.hwpx> [output.hwpx]
"""

import os
import sys
import zipfile
import tempfile
import shutil
import re


def fix_ampersands_in_xml(xml_content):
    """
    Replace bare & with &amp; in XML content.
    Preserves already-escaped entities: &amp; &lt; &gt; &quot; &apos; &#NNN;
    """
    pattern = r'&(?!(amp|lt|gt|quot|apos|#\d+);)'
    return re.sub(pattern, '&amp;', xml_content)


def fix_hwpx_file(input_hwpx, output_hwpx):
    """Open HWPX file, fix XML entities, save as new file."""
    print(f"Fixing HWPX XML: {input_hwpx}\n")

    temp_dir = tempfile.mkdtemp()

    try:
        with zipfile.ZipFile(input_hwpx, 'r') as zip_ref:
            zip_ref.extractall(temp_dir)

        xml_files = ['Contents/section0.xml', 'Contents/header.xml']

        for xml_file in xml_files:
            xml_path = os.path.join(temp_dir, xml_file)
            if not os.path.exists(xml_path):
                continue

            with open(xml_path, 'r', encoding='utf-8') as f:
                original_content = f.read()

            fixed_content = fix_ampersands_in_xml(original_content)

            with open(xml_path, 'w', encoding='utf-8') as f:
                f.write(fixed_content)

            print(f"Fixed: {xml_file}")

        # Repackage HWPX
        with zipfile.ZipFile(output_hwpx, 'w', zipfile.ZIP_DEFLATED) as zipf:
            mimetype_path = os.path.join(temp_dir, 'mimetype')
            if os.path.exists(mimetype_path):
                zipf.write(mimetype_path, 'mimetype',
                           compress_type=zipfile.ZIP_STORED)

            for root, dirs, files in os.walk(temp_dir):
                for file in files:
                    if file == 'mimetype':
                        continue
                    file_path = os.path.join(root, file)
                    arcname = os.path.relpath(file_path, temp_dir)
                    zipf.write(file_path, arcname)

        print(f"\nFixed file: {os.path.abspath(output_hwpx)}")

    finally:
        shutil.rmtree(temp_dir)


def main():
    if len(sys.argv) < 2:
        print("Usage: python3 fix_xml.py <input.hwpx> [output.hwpx]")
        sys.exit(1)

    input_file = sys.argv[1]
    if not os.path.exists(input_file):
        print(f"ERROR: File not found: {input_file}")
        sys.exit(1)

    if len(sys.argv) >= 3:
        output_file = sys.argv[2]
    else:
        base, ext = os.path.splitext(input_file)
        output_file = f"{base}_fixed{ext}"

    fix_hwpx_file(input_file, output_file)


if __name__ == '__main__':
    main()
