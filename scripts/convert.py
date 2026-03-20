#!/usr/bin/env python3
"""
Markdown → HWPX converter (cross-platform).

Converts a Markdown file to HWPX format using pypandoc-hwpx,
then applies post-processing (table styling, bold formatting).

Usage:
    python3 convert.py <input.md> [output_name]
"""

import os
import sys
import subprocess
import re
import unicodedata
import importlib.util
from datetime import datetime


def _load_module(module_name, filename):
    """Dynamically load a sibling script by filename."""
    script_dir = os.path.dirname(os.path.abspath(__file__))
    mod_path = os.path.join(script_dir, filename)
    spec = importlib.util.spec_from_file_location(module_name, mod_path)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def _get_env():
    """Get venv Python and pypandoc-hwpx paths."""
    env_mod = _load_module('env_detect', 'env_detect.py')
    venv_root = env_mod._find_venv_root()
    python_path = env_mod.find_venv_python(venv_root)
    pypandoc_path = env_mod.find_pypandoc_hwpx(venv_root)
    return python_path, pypandoc_path


def extract_title_from_markdown(md_file):
    """Extract title from first # heading in Markdown file."""
    with open(md_file, 'r', encoding='utf-8') as f:
        for line in f:
            line = line.strip()
            if line.startswith('# '):
                title = line[2:].strip()
                title = re.sub(r'[^\w\s-]', '', title)
                title = re.sub(r'[-\s]+', '_', title)
                return unicodedata.normalize('NFC', title)
    return os.path.splitext(os.path.basename(md_file))[0]


def convert_markdown_to_hwpx(input_md, output_dir='output', custom_name=None):
    """
    Convert Markdown to HWPX and save to output directory.

    Args:
        input_md: Input Markdown file path
        output_dir: Output directory (default: 'output')
        custom_name: Custom filename without extension (optional)

    Returns:
        Output file path on success, None on failure.
    """
    print(f"Markdown -> HWPX conversion started\n")
    print(f"Input: {input_md}")

    os.makedirs(output_dir, exist_ok=True)

    # Determine output filename
    if custom_name:
        base_name = custom_name
    else:
        title = extract_title_from_markdown(input_md)
        date_str = datetime.now().strftime('%Y%m%d')
        base_name = f"{date_str}_{title}"

    output_file = os.path.join(output_dir, f"{base_name}.hwpx")
    print(f"Output: {output_file}\n")

    # Find pypandoc-hwpx
    python_path, pypandoc_path = _get_env()

    if not python_path or not pypandoc_path:
        print("ERROR: pypandoc-hwpx is not installed.")
        print("Setup instructions:")
        if sys.platform == 'win32':
            print("  python -m venv hwpx_env")
            print("  hwpx_env\\Scripts\\activate")
        else:
            print("  python3 -m venv hwpx_env")
            print("  source hwpx_env/bin/activate")
        print("  pip install pypandoc-hwpx")
        return None

    try:
        result = subprocess.run(
            [python_path, pypandoc_path, input_md, '-o', output_file],
            capture_output=True,
            text=True
        )

        if result.returncode == 0 and os.path.exists(output_file):
            file_size = os.path.getsize(output_file)
            print("=" * 60)
            print("Conversion complete!")
            print("=" * 60)
            print(f"\nFile: {os.path.abspath(output_file)}")
            print(f"Size: {file_size:,} bytes")

            # Post-processing: table styling
            try:
                style_mod = _load_module('style_tables', 'style_tables.py')
                style_mod.center_table_headers(output_file)
            except Exception as e:
                print(f"WARNING: Table styling failed: {e}")

            # Post-processing: bold formatting
            try:
                bold_mod = _load_module('apply_bold', 'apply_bold.py')
                bold_mod.apply_bold(input_md, output_file)
            except Exception as e:
                print(f"WARNING: Bold formatting failed: {e}")

            # Post-processing: superscript footnotes
            try:
                sup_mod = _load_module('apply_superscript',
                                       'apply_superscript.py')
                sup_mod.apply_superscript(output_file)
            except Exception as e:
                print(f"WARNING: Superscript formatting failed: {e}")

            return output_file
        else:
            print(f"ERROR: Conversion failed: {result.stderr}")
            return None

    except Exception as e:
        print(f"ERROR: {e}")
        import traceback
        traceback.print_exc()
        return None


def main():
    if len(sys.argv) < 2:
        print("Usage: python3 convert.py <input.md> [output_name]")
        print("\nExamples:")
        print("  python3 convert.py document.md")
        print("  python3 convert.py document.md MyReport")
        sys.exit(1)

    input_file = sys.argv[1]
    custom_name = sys.argv[2] if len(sys.argv) > 2 else None

    if not os.path.exists(input_file):
        print(f"ERROR: File not found: {input_file}")
        sys.exit(1)

    convert_markdown_to_hwpx(input_file, custom_name=custom_name)


if __name__ == '__main__':
    main()
