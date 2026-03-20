#!/usr/bin/env python3
"""
Bold formatting post-processor for HWPX files (cross-platform).

Extracts **bold** text from original Markdown and applies bold formatting
to matching text runs in the HWPX output.

pypandoc-hwpx does not generate charPr for inline **bold** text,
so this script patches header.xml and section0.xml to add bold styling.

Usage:
    python3 apply_bold.py <original.md> <output.hwpx>
"""

import sys
import os
import re
import zipfile
import shutil
import tempfile


def extract_bold_phrases(md_path):
    """Extract **bold** text from Markdown. Returns (phrases, words_set)."""
    with open(md_path, 'r', encoding='utf-8') as f:
        content = f.read()

    phrases = re.findall(r'\*\*([^*]+?)\*\*', content)
    unique = list(dict.fromkeys(phrases))
    unique.sort(key=len, reverse=True)

    words_set = set()
    for phrase in unique:
        for word in phrase.split():
            if len(word) >= 1:
                words_set.add(word)

    return unique, words_set


def add_bold_charpr(header_xml):
    """Add a bold charPr to header.xml. Returns (modified_xml, new_id)."""
    ids = [int(x) for x in re.findall(r'<hh:charPr id="(\d+)"', header_xml)]
    new_id = max(ids) + 1

    match = re.search(
        r'(<hh:charPr id="0".*?</hh:charPr>)', header_xml, re.DOTALL
    )
    if not match:
        return header_xml, -1

    base_charpr = match.group(1)
    bold_charpr = base_charpr.replace('id="0"', f'id="{new_id}"', 1)
    bold_charpr = bold_charpr.replace(
        '<hh:underline',
        '<hh:bold />\n        <hh:underline'
    )

    header_xml = header_xml.replace(
        '</hh:charProperties>',
        f'      {bold_charpr}\n      </hh:charProperties>'
    )
    header_xml = re.sub(
        r'charProperties itemCnt="\d+"',
        f'charProperties itemCnt="{new_id + 1}"',
        header_xml
    )

    return header_xml, new_id


def _is_inside_table(xml, pos):
    """Check if position is inside an <hp:tbl> block."""
    before = xml[:pos]
    return before.count('<hp:tbl') > before.count('</hp:tbl>')


def apply_bold_to_runs(section_xml, bold_phrases, bold_words, bold_charpr_id):
    """Apply bold charPrIDRef to matching text runs in section0.xml.
    Skips runs inside <hp:tbl> blocks to preserve table charPr."""
    paras = re.split(r'(<hp:p\b[^>]*>)', section_xml)
    result = []
    total_applied = 0
    i = 0
    # Track table depth from non-para parts
    tbl_depth = 0

    while i < len(paras):
        part = paras[i]

        if not part.startswith('<hp:p '):
            # Track table open/close tags in non-para parts
            tbl_depth += part.count('<hp:tbl')
            tbl_depth -= part.count('</hp:tbl>')
            result.append(part)
            i += 1
            continue

        para_start = part
        para_body = paras[i + 1] if i + 1 < len(paras) else ""

        # Skip bold application inside tables
        if tbl_depth > 0:
            result.append(para_start)
            if i + 1 < len(paras):
                result.append(para_body)
                i += 2
            else:
                i += 1
            continue

        runs_with_text = list(re.finditer(
            r'<hp:run charPrIDRef="(\d+)"><hp:t>(.*?)</hp:t></hp:run>',
            para_body
        ))

        if not runs_with_text:
            result.append(para_start)
            if i + 1 < len(paras):
                result.append(para_body)
                i += 2
            else:
                i += 1
            continue

        full_text = ""
        char_to_run = []

        for ri, m in enumerate(runs_with_text):
            text = m.group(2)
            for ci, ch in enumerate(text):
                char_to_run.append((ri, ci))
            full_text += text

        bold_char_flags = [False] * len(full_text)

        for phrase in bold_phrases:
            start = 0
            while True:
                idx = full_text.find(phrase, start)
                if idx < 0:
                    break
                for j in range(idx, idx + len(phrase)):
                    bold_char_flags[j] = True
                start = idx + 1

        bold_runs = set()
        for ci, is_bold in enumerate(bold_char_flags):
            if is_bold:
                ri, _ = char_to_run[ci]
                bold_runs.add(ri)

        if bold_runs:
            modified_body = para_body
            for ri in sorted(bold_runs, reverse=True):
                m = runs_with_text[ri]
                old_run = m.group(0)
                old_ref = m.group(1)
                new_run = old_run.replace(
                    f'charPrIDRef="{old_ref}"',
                    f'charPrIDRef="{bold_charpr_id}"',
                    1
                )
                modified_body = (modified_body[:m.start()]
                                 + new_run + modified_body[m.end():]
                                 )
                total_applied += 1

            result.append(para_start)
            result.append(modified_body)
        else:
            result.append(para_start)
            result.append(para_body)

        i += 2

    return ''.join(result), total_applied


def apply_bold(md_path, hwpx_path):
    """
    Apply bold formatting from Markdown **bold** text to HWPX file.
    Returns: number of bold runs applied (0 if no bold text found).
    """
    if not os.path.exists(md_path) or not os.path.exists(hwpx_path):
        return 0

    bold_phrases, bold_words = extract_bold_phrases(md_path)
    if not bold_phrases:
        return 0

    extract_dir = tempfile.mkdtemp()
    with zipfile.ZipFile(hwpx_path, 'r') as z:
        z.extractall(extract_dir)

    header_path = os.path.join(extract_dir, 'Contents', 'header.xml')
    with open(header_path, 'r', encoding='utf-8') as f:
        header_xml = f.read()

    header_xml, bold_id = add_bold_charpr(header_xml)
    if bold_id < 0:
        shutil.rmtree(extract_dir)
        return 0

    with open(header_path, 'w', encoding='utf-8') as f:
        f.write(header_xml)

    section_path = os.path.join(extract_dir, 'Contents', 'section0.xml')
    with open(section_path, 'r', encoding='utf-8') as f:
        section_xml = f.read()

    section_xml, applied_count = apply_bold_to_runs(
        section_xml, bold_phrases, bold_words, bold_id
    )

    with open(section_path, 'w', encoding='utf-8') as f:
        f.write(section_xml)

    # Repackage
    original_files = []
    with zipfile.ZipFile(hwpx_path, 'r') as z:
        original_files = z.namelist()

    temp_hwpx = hwpx_path + '.tmp'
    with zipfile.ZipFile(temp_hwpx, 'w', zipfile.ZIP_DEFLATED) as zout:
        for fname in original_files:
            fpath = os.path.join(extract_dir, fname)
            if os.path.exists(fpath):
                zout.write(fpath, fname)

    shutil.move(temp_hwpx, hwpx_path)
    shutil.rmtree(extract_dir)

    print(f"Bold applied: {applied_count} run(s) "
          f"({len(bold_phrases)} phrase(s))")
    return applied_count


def main():
    if len(sys.argv) < 3:
        print("Usage: python3 apply_bold.py <original.md> <output.hwpx>")
        sys.exit(1)

    md_path = sys.argv[1]
    hwpx_path = sys.argv[2]

    if not os.path.exists(md_path):
        print(f"ERROR: {md_path} not found")
        sys.exit(1)
    if not os.path.exists(hwpx_path):
        print(f"ERROR: {hwpx_path} not found")
        sys.exit(1)

    apply_bold(md_path, hwpx_path)


if __name__ == '__main__':
    main()
