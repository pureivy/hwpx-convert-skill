#!/usr/bin/env python3
"""
Superscript formatting post-processor for HWPX files.

Finds footnote reference patterns like 1), 23), 55) within body text runs
and applies superscript formatting (70% relative size, 30% upward offset).

Only applies to mid-text footnote references, NOT to reference list entries
that begin with "1) Author..." at the start of a run.

Skips all runs inside <hp:tbl> blocks (table content).

Usage:
    python3 apply_superscript.py <output.hwpx>
"""

import sys
import os
import re
import zipfile
import shutil
import tempfile

# Pattern for footnote references: 1-3 digit number followed by )
FOOTNOTE_RE = re.compile(r'\d{1,3}\)')

# Pattern for reference list entry: starts with number) (with or without space)
REF_LIST_RE = re.compile(r'^\d{1,3}\)$|^\d{1,3}\) ')

# Pattern for consecutive footnote group: one or more N) patterns together
FOOTNOTE_GROUP_RE = re.compile(r'(\d{1,3}\))+')

LANGS = ['hangul', 'latin', 'hanja', 'japanese', 'other', 'symbol', 'user']


def add_superscript_charpr(header_xml):
    """Add a superscript charPr to header.xml.

    Clones charPr id=0, sets relSz=70 and offset=30 for all langs.
    Returns (modified_xml, new_id).
    """
    ids = [int(x) for x in re.findall(r'<hh:charPr id="(\d+)"', header_xml)]
    if not ids:
        return header_xml, -1
    new_id = max(ids) + 1

    match = re.search(
        r'(<hh:charPr id="0".*?</hh:charPr>)', header_xml, re.DOTALL
    )
    if not match:
        return header_xml, -1

    base_charpr = match.group(1)
    sup_charpr = base_charpr.replace('id="0"', f'id="{new_id}"', 1)

    # Set relSz to 70 for all langs
    sup_charpr = re.sub(
        r'<hh:relSz\b[^/]*/>', _build_lang_tag('relSz', 70), sup_charpr
    )

    # Set offset to -30 for all langs (negative = upward in HWPX)
    sup_charpr = re.sub(
        r'<hh:offset\b[^/]*/>', _build_lang_tag('offset', -30), sup_charpr
    )

    # Insert before </hh:charProperties>
    header_xml = header_xml.replace(
        '</hh:charProperties>',
        f'      {sup_charpr}\n      </hh:charProperties>'
    )

    # Increment itemCnt
    header_xml = re.sub(
        r'charProperties itemCnt="\d+"',
        f'charProperties itemCnt="{new_id + 1}"',
        header_xml
    )

    return header_xml, new_id


def _build_lang_tag(tag_name, value):
    """Build an hh: tag with the same value for all 7 language attributes."""
    attrs = ' '.join(f'{lang}="{value}"' for lang in LANGS)
    return f'<hh:relSz {attrs}/>' if tag_name == 'relSz' else f'<hh:offset {attrs}/>'


def _split_run_text(text, sup_charpr_id, original_ref):
    """Split a text string into segments: normal text and footnote groups.

    Returns a list of (text, charPrIDRef) tuples.
    """
    segments = []
    last_end = 0

    for m in FOOTNOTE_GROUP_RE.finditer(text):
        # Text before the footnote group
        if m.start() > last_end:
            segments.append((text[last_end:m.start()], original_ref))
        # The footnote group itself
        segments.append((m.group(0), str(sup_charpr_id)))
        last_end = m.end()

    # Text after the last footnote group
    if last_end < len(text):
        segments.append((text[last_end:], original_ref))

    return segments


def _is_inside_table(section_xml, run_start_pos):
    """Check if a position is inside an <hp:tbl> block.

    Counts open <hp:tbl> and </hp:tbl> tags before the position.
    If more opens than closes, we're inside a table.
    """
    before = section_xml[:run_start_pos]
    opens = len(re.findall(r'<hp:tbl\b', before))
    closes = len(re.findall(r'</hp:tbl>', before))
    return opens > closes


def apply_superscript_to_runs(section_xml, sup_charpr_id):
    """Apply superscript charPrIDRef to footnote reference patterns.

    Finds runs with text containing N) patterns (not at the start),
    splits them into normal + superscript runs.
    Skips runs inside <hp:tbl> blocks.

    Returns (modified_xml, count_of_superscript_runs).
    """
    # Find all text runs: <hp:run charPrIDRef="X"><hp:t>text</hp:t></hp:run>
    run_pattern = re.compile(
        r'<hp:run charPrIDRef="(\d+)"><hp:t>(.*?)</hp:t></hp:run>'
    )

    total_applied = 0
    result_parts = []
    last_end = 0

    # Find heading paragraph positions (paraPrIDRef 2-12)
    heading_para_ranges = []
    for pm in re.finditer(
        r'<hp:p paraPrIDRef="(\d+)"[^>]*>.*?</hp:p>',
        section_xml, re.DOTALL
    ):
        pid = int(pm.group(1))
        if 2 <= pid <= 12:
            heading_para_ranges.append((pm.start(), pm.end()))

    def _is_inside_heading(pos):
        return any(s <= pos < e for s, e in heading_para_ranges)

    for m in run_pattern.finditer(section_xml):
        original_ref = m.group(1)
        text = m.group(2)

        # Skip if inside a table
        if _is_inside_table(section_xml, m.start()):
            continue

        # Skip if inside a heading paragraph
        if _is_inside_heading(m.start()):
            continue

        # Skip if this is a reference list entry (starts with "N) ")
        if REF_LIST_RE.match(text):
            continue

        # Skip if no footnote pattern found
        if not FOOTNOTE_RE.search(text):
            continue

        # Check: only apply to mid-text footnotes.
        # If the entire text is just footnote references (e.g., "1)2)3)"),
        # and it's at position 0, that could be a standalone reference.
        # But per spec, we check if it starts with "N) " (with space) for
        # reference list entries. A bare "1)2)" without space is a footnote.

        segments = _split_run_text(text, sup_charpr_id, original_ref)

        # If splitting produced only one segment with the original ref,
        # nothing to change
        if len(segments) == 1 and segments[0][1] == original_ref:
            continue

        # Build replacement runs
        new_runs = []
        sup_count = 0
        for seg_text, seg_ref in segments:
            if seg_text:  # skip empty segments
                new_runs.append(
                    f'<hp:run charPrIDRef="{seg_ref}">'
                    f'<hp:t>{seg_text}</hp:t></hp:run>'
                )
                if seg_ref == str(sup_charpr_id):
                    sup_count += 1

        if sup_count > 0:
            # Copy everything before this match that hasn't been copied yet
            result_parts.append(section_xml[last_end:m.start()])
            result_parts.append(''.join(new_runs))
            last_end = m.end()
            total_applied += sup_count

    # Append the rest of the document
    result_parts.append(section_xml[last_end:])

    return ''.join(result_parts), total_applied


def apply_superscript(hwpx_path):
    """Apply superscript to footnote references N) in HWPX file.
    Returns number of superscript runs applied."""
    if not os.path.exists(hwpx_path):
        return 0

    extract_dir = tempfile.mkdtemp()
    with zipfile.ZipFile(hwpx_path, 'r') as z:
        z.extractall(extract_dir)

    # 1. Modify header.xml: add superscript charPr
    header_path = os.path.join(extract_dir, 'Contents', 'header.xml')
    with open(header_path, 'r', encoding='utf-8') as f:
        header_xml = f.read()

    header_xml, sup_id = add_superscript_charpr(header_xml)
    if sup_id < 0:
        shutil.rmtree(extract_dir)
        return 0

    with open(header_path, 'w', encoding='utf-8') as f:
        f.write(header_xml)

    # 2. Modify section0.xml: split runs with footnote refs
    section_path = os.path.join(extract_dir, 'Contents', 'section0.xml')
    with open(section_path, 'r', encoding='utf-8') as f:
        section_xml = f.read()

    section_xml, applied_count = apply_superscript_to_runs(section_xml, sup_id)

    with open(section_path, 'w', encoding='utf-8') as f:
        f.write(section_xml)

    # 3. Repackage HWPX (mimetype stored, same ZIP pattern as apply_bold.py)
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

    print(f"Superscript applied: {applied_count} run(s)")
    return applied_count


def main():
    if len(sys.argv) < 2:
        print("Usage: python3 apply_superscript.py <output.hwpx>")
        sys.exit(1)

    hwpx_path = sys.argv[1]

    if not os.path.exists(hwpx_path):
        print(f"ERROR: {hwpx_path} not found")
        sys.exit(1)

    apply_superscript(hwpx_path)


if __name__ == '__main__':
    main()
