#!/usr/bin/env python3
"""
HWPX table and body style post-processor (cross-platform).

Applies:
- Body text: justify alignment (except headings)
- Body charPr: 13pt, 함초롬바탕, 장평 98%, 자간 -5
- Body paragraph spacing: prev=500, next=0
- Heading charPr: per-level sizes, 한컴 소망 B font
- Table headers: center alignment + background color + 130% line spacing
- Table body rows: left alignment + 130% line spacing
- Table text: 11pt, 맑은 고딕, 문단위 0pt, 문단아래 0pt
- Table position: center on page (hp:pos horzAlign=CENTER)
- Cell vertical alignment: center (hp:subList vertAlign=CENTER)
- Cell margin: top/bottom 1.50mm (425 HWPUNIT)
- Column widths: proportional auto-sizing based on character count
  (Korean chars = 1.8 width, ASCII = 1.0 width)
- Numbered list indent: left=1000, intent=-1700

Usage:
    python3 style_tables.py <input.hwpx> [output.hwpx]
    (if output is omitted, input file is overwritten)
"""

import os
import sys
import re
import zipfile
import shutil
import tempfile

DEFAULT_HEADER_COLOR = '#BDD7EE'
HEADING_PARA_IDS = set(range(2, 13))

# Body text: 13pt, 함초롬바탕, 장평 98%, 자간 -5
CHAR_RATIO = '98'
CHAR_SPACING = '-5'
CHAR_HEIGHT = '1300'
CHAR_FONT_FACE = '함초롬바탕'

# Paragraph spacing: 문단위 5pt, 문단아래 0pt
PARA_SPACING_PREV = '500'
PARA_SPACING_NEXT = '0'

# Numbered list: 왼쪽 10pt, 내어쓰기 17pt
LIST_LEFT = '1000'
LIST_INTENT = '-1700'

# Heading styles: charPrIDRef → height, font 한컴 소망 B
HEADING_FONT_FACE = '한컴 소망 B'
# Heading paragraph spacing: paraPrIDRef → prev (HWPUNIT)
HEADING_PARA_PREV = {
    4: '2000',  # H3 (paraPrIDRef=4) — 문단위 20pt
}
HEADING_SIZES = {
    4: '2400',  # H1 — 24pt
    5: '2000',  # H2 — 20pt
    6: '1600',  # H3 — 16pt
    7: '1400',  # H4 — 14pt
    8: '1300',  # H5 — 13pt
}

# Table text: 11pt, 나눔고딕, 문단위 0pt, 문단아래 0pt
TABLE_CHAR_HEIGHT = '1100'
TABLE_FONT_FACE = '맑은 고딕'
TABLE_PARA_PREV = '0'
TABLE_PARA_NEXT = '0'


def _calc_text_width(text):
    """Estimate visual width of text. CJK=1.8, ASCII=1.0, min=3."""
    if not text:
        return 3
    w = sum(1.8 if ord(c) > 127 else 1.0 for c in text)
    return max(w, 3)


def _calc_proportional_widths(table_xml, total_width):
    """Calculate proportional column widths based on cell content."""
    rows = re.findall(r'<hp:tr>(.*?)</hp:tr>', table_xml, re.DOTALL)
    if not rows:
        return None

    col_max_width = {}
    num_cols = 0

    for row in rows:
        cells = re.findall(r'<hp:tc[^>]*>(.*?)</hp:tc>', row, re.DOTALL)
        for cell in cells:
            col_m = re.search(r'colAddr="(\d+)"', cell)
            span_m = re.search(r'colSpan="(\d+)"', cell)
            if not col_m:
                continue
            col = int(col_m.group(1))
            span = int(span_m.group(1)) if span_m else 1
            num_cols = max(num_cols, col + span)
            if span > 1:
                continue
            texts = re.findall(r'<hp:t>(.*?)</hp:t>', cell)
            text = ''.join(texts).strip()
            tw = _calc_text_width(text)
            col_max_width[col] = max(col_max_width.get(col, 3), tw)

    if num_cols < 2 or not col_max_width:
        return None

    for c in range(num_cols):
        if c not in col_max_width:
            col_max_width[c] = 3

    total_chars = sum(col_max_width.values())
    min_col_width = max(3000, total_width // (num_cols * 3))

    widths = {}
    for col in range(num_cols):
        w = int(total_width * col_max_width[col] / total_chars)
        widths[col] = max(w, min_col_width)

    current = sum(widths.values())
    if current != total_width:
        max_col = max(widths, key=widths.get)
        widths[max_col] += (total_width - current)

    return widths


def _update_ratio_spacing(xml_str, ratio, spacing):
    """Update ratio and spacing attributes across all lang groups."""
    langs = ['hangul', 'latin', 'hanja', 'japanese',
             'other', 'symbol', 'user']
    for lang in langs:
        xml_str = re.sub(
            rf'(<hh:ratio\b[^>]*\b{lang}=")[^"]+(")',
            rf'\g<1>{ratio}\2', xml_str)
        xml_str = re.sub(
            rf'(<hh:spacing\b[^>]*\b{lang}=")[^"]+(")',
            rf'\g<1>{spacing}\2', xml_str)
    return xml_str


def _update_fontref_all_langs(xml_str, font_id):
    """Update fontRef for all lang attributes to the given font_id."""
    def update_fontref(fr_match):
        tag = fr_match.group(0)
        for attr in ['hangul', 'latin', 'hanja', 'japanese',
                     'other', 'symbol', 'user']:
            tag = re.sub(
                rf'({attr}=")[^"]+(")',
                rf'\g<1>{font_id}\2', tag)
        return tag
    return re.sub(r'<hh:fontRef\b[^/]*/>', update_fontref, xml_str)


def _register_font(header_xml, font_face):
    """Register a font in all fontface lang groups. Returns (header_xml, font_id)."""
    # Check if font already exists
    font_m = re.search(
        r'<hh:font id="(\d+)" face="'
        + re.escape(font_face) + r'"',
        header_xml
    )
    if font_m:
        return header_xml, font_m.group(1)

    # Add font to each fontface group
    def add_font_to_fontface(match):
        block = match.group(0)
        ids = [int(x) for x in re.findall(
            r'<hh:font id="(\d+)"', block)]
        new_id = max(ids) + 1 if ids else 0
        new_font = (f'<hh:font id="{new_id}" '
                    f'face="{font_face}" '
                    f'type="TTF" isEmbedded="0" />')
        block = block.replace(
            '</hh:fontface>',
            f'        {new_font}\n      </hh:fontface>'
        )
        block = re.sub(
            r'fontCnt="(\d+)"',
            lambda m: f'fontCnt="{int(m.group(1)) + 1}"',
            block
        )
        return block

    header_xml = re.sub(
        r'<hh:fontface lang="[^"]*".*?</hh:fontface>',
        add_font_to_fontface, header_xml, flags=re.DOTALL
    )
    # Get the new font id from HANGUL group
    font_m = re.search(
        r'<hh:fontface lang="HANGUL".*?'
        r'<hh:font id="(\d+)" face="'
        + re.escape(font_face) + r'"',
        header_xml, re.DOTALL
    )
    font_id = font_m.group(1) if font_m else '0'
    return header_xml, font_id


def center_table_headers(hwpx_path, output_path=None,
                          header_color=DEFAULT_HEADER_COLOR):
    """
    Apply comprehensive post-processing to HWPX file:
    - Body justify alignment
    - Body charPr: 13pt, 함초롬바탕, 장평 98%, 자간 -5
    - Heading charPr: per-level sizes, 한컴 소망 B font
    - Table header: center + background color + 130% line spacing
    - Table body: justify + 130% line spacing
    - Table text: 11pt, 맑은 고딕, 문단위 0pt
    - Column widths: proportional to content

    If output_path is None, overwrites input file.
    """
    if output_path is None:
        output_path = hwpx_path

    tmp_dir = tempfile.mkdtemp(prefix='hwpx_center_')
    try:
        with zipfile.ZipFile(hwpx_path, 'r') as zf:
            zf.extractall(tmp_dir)

        header_path = os.path.join(tmp_dir, 'Contents', 'header.xml')
        section_path = os.path.join(tmp_dir, 'Contents', 'section0.xml')

        with open(header_path, 'r', encoding='utf-8') as f:
            header_xml = f.read()
        with open(section_path, 'r', encoding='utf-8') as f:
            section_xml = f.read()

        # ── 0. Body justify: non-heading paraPr → JUSTIFY ──────────
        heading_pids = set()
        for m in re.finditer(
            r'<hh:style\b[^>]*name="제목[^"]*"[^>]*paraPrIDRef="(\d+)"',
            header_xml
        ):
            heading_pids.add(int(m.group(1)))
        for m in re.finditer(
            r'<hh:style\b[^>]*name="부제목"[^>]*paraPrIDRef="(\d+)"',
            header_xml
        ):
            heading_pids.add(int(m.group(1)))
        heading_pids.update(HEADING_PARA_IDS)

        justify_count = 0

        def justify_para_pr(match):
            nonlocal justify_count
            full = match.group(0)
            pid_m = re.search(r'id="(\d+)"', full)
            if not pid_m:
                return full
            pid = int(pid_m.group(1))
            if pid in heading_pids:
                return full
            if 'horizontal="LEFT"' in full:
                justify_count += 1
                return full.replace('horizontal="LEFT"',
                                    'horizontal="JUSTIFY"')
            return full

        header_xml = re.sub(
            r'<hh:paraPr id="\d+".*?</hh:paraPr>',
            justify_para_pr, header_xml, flags=re.DOTALL
        )

        # ── 0b. Character shape: ratio/spacing for ALL charPr ──────
        #         height/font for non-heading charPr only
        heading_charpr_ids = set()
        for m in re.finditer(
            r'<hh:style\b[^>]*name="(?:제목|부제목)[^"]*"'
            r'[^>]*charPrIDRef="(\d+)"', header_xml
        ):
            heading_charpr_ids.add(int(m.group(1)))

        charpr_count = 0

        def update_charpr(match):
            nonlocal charpr_count
            full = match.group(0)
            pid_m = re.search(r'id="(\d+)"', full)
            cpid = int(pid_m.group(1)) if pid_m else -1
            is_heading = cpid in heading_charpr_ids

            # ratio + spacing for every charPr
            full = _update_ratio_spacing(full, CHAR_RATIO, CHAR_SPACING)

            if is_heading:
                # Heading-specific height
                if cpid in HEADING_SIZES:
                    full = re.sub(
                        r'(height=")[^"]+(")',
                        rf'\g<1>{HEADING_SIZES[cpid]}\2', full, count=1)
            else:
                # Body height
                full = re.sub(
                    r'(height=")[^"]+(")',
                    rf'\g<1>{CHAR_HEIGHT}\2', full, count=1)
            charpr_count += 1
            return full

        header_xml = re.sub(
            r'<hh:charPr id="\d+".*?</hh:charPr>',
            update_charpr, header_xml, flags=re.DOTALL
        )

        # ── 0b-2. Register body font (함초롬바탕) ──────────────────
        header_xml, body_font_id = _register_font(
            header_xml, CHAR_FONT_FACE)
        print(f"Font: {CHAR_FONT_FACE} (id={body_font_id}), "
              f"size={CHAR_HEIGHT} ({int(CHAR_HEIGHT)//100}pt)")

        # ── 0b-3. Register heading font (한컴 소망 B) ──────────────
        header_xml, heading_font_id = _register_font(
            header_xml, HEADING_FONT_FACE)
        print(f"Font: {HEADING_FONT_FACE} (id={heading_font_id}), "
              f"heading sizes={dict(HEADING_SIZES)}")

        # ── 0b-4. Update fontRef per charPr ────────────────────────
        #   - heading charPr → heading_font_id
        #   - non-heading charPr → body_font_id
        def update_charpr_font(match):
            full = match.group(0)
            pid_m = re.search(r'id="(\d+)"', full)
            if pid_m and int(pid_m.group(1)) in heading_charpr_ids:
                return _update_fontref_all_langs(full, heading_font_id)
            return _update_fontref_all_langs(full, body_font_id)

        header_xml = re.sub(
            r'<hh:charPr id="\d+".*?</hh:charPr>',
            update_charpr_font, header_xml, flags=re.DOTALL
        )

        # ── 0c. Paragraph spacing: prev/next (non-heading paraPr) ──
        spacing_count = 0

        def update_para_spacing(match):
            nonlocal spacing_count
            full = match.group(0)
            pid_m = re.search(r'id="(\d+)"', full)
            if not pid_m:
                return full
            pid = int(pid_m.group(1))
            if pid in heading_pids:
                return full
            full = re.sub(
                r'(<hc:prev value=")[^"]+(")',
                rf'\g<1>{PARA_SPACING_PREV}\2', full)
            full = re.sub(
                r'(<hc:next value=")[^"]+(")',
                rf'\g<1>{PARA_SPACING_NEXT}\2', full)
            spacing_count += 1
            return full

        header_xml = re.sub(
            r'<hh:paraPr id="\d+".*?</hh:paraPr>',
            update_para_spacing, header_xml, flags=re.DOTALL
        )

        # ── 0d. Heading paraPr: remove OUTLINE numbering ──────────
        #   heading type="OUTLINE" → "NONE" to remove auto-numbering indent
        heading_outline_count = 0

        def remove_outline(match):
            nonlocal heading_outline_count
            full = match.group(0)
            pid_m = re.search(r'id="(\d+)"', full)
            if not pid_m:
                return full
            pid = int(pid_m.group(1))
            if pid not in heading_pids:
                return full
            # Remove OUTLINE numbering and reset left/intent to 0
            for htype in ['OUTLINE', 'NUMBER']:
                if f'heading type="{htype}"' in full:
                    full = full.replace(
                        f'heading type="{htype}"', 'heading type="NONE"')
                    heading_outline_count += 1
            # Force left=0, intent=0 for headings
            full = re.sub(
                r'(<hc:left value=")[^"]+(")', r'\g<1>0\2', full)
            full = re.sub(
                r'(<hc:intent value=")[^"]+(")', r'\g<1>0\2', full)
            # Apply heading-specific paragraph spacing
            if pid in HEADING_PARA_PREV:
                full = re.sub(
                    r'(<hc:prev value=")[^"]+(")',
                    rf'\g<1>{HEADING_PARA_PREV[pid]}\2', full)
            return full

        header_xml = re.sub(
            r'<hh:paraPr id="\d+".*?</hh:paraPr>',
            remove_outline, header_xml, flags=re.DOTALL
        )
        if heading_outline_count:
            print(f"Heading outline removed: {heading_outline_count} paraPr(s)")

        # ── 0e. Numbered list indent ───────────────────────────────
        list_count = 0

        def update_list_indent(match):
            nonlocal list_count
            full = match.group(0)
            if 'heading type="NUMBER"' not in full:
                return full
            full = re.sub(
                r'(<hc:left value=")[^"]+(")',
                rf'\g<1>{LIST_LEFT}\2', full)
            full = re.sub(
                r'(<hc:intent value=")[^"]+(")',
                rf'\g<1>{LIST_INTENT}\2', full)
            list_count += 1
            return full

        header_xml = re.sub(
            r'<hh:paraPr id="\d+".*?</hh:paraPr>',
            update_list_indent, header_xml, flags=re.DOTALL
        )

        # ── 1. Compute new paraPr IDs ──────────────────────────────
        max_para_id = max(
            (int(m) for m in re.findall(
                r'<hh:paraPr id="(\d+)"', header_xml)),
            default=0
        )
        header_para_id = max_para_id + 1
        body_para_id = max_para_id + 2

        # ── 2. Clone paraPr id=0 for table header/body ─────────────
        para_pr_0 = re.search(
            r'(<hh:paraPr id="0".*?</hh:paraPr>)',
            header_xml, re.DOTALL
        )
        if not para_pr_0:
            print("WARNING: paraPr id=0 not found.")
            return False

        def make_para_pr(base, new_id, horizontal, prev=None, next_val=None):
            pr = base
            pr = pr.replace('id="0"', f'id="{new_id}"', 1)
            pr = re.sub(r'horizontal="[^"]+"',
                        f'horizontal="{horizontal}"', pr)
            pr = re.sub(
                r'(<hh:lineSpacing\b[^>]*\bvalue=")[^"]+(")',
                r'\g<1>130\2', pr
            )
            if prev is not None:
                pr = re.sub(
                    r'(<hc:prev value=")[^"]+(")',
                    rf'\g<1>{prev}\2', pr)
            if next_val is not None:
                pr = re.sub(
                    r'(<hc:next value=")[^"]+(")',
                    rf'\g<1>{next_val}\2', pr)
            return pr

        header_pr = make_para_pr(
            para_pr_0.group(1), header_para_id, 'CENTER',
            prev=TABLE_PARA_PREV, next_val=TABLE_PARA_NEXT
        )
        body_pr = make_para_pr(
            para_pr_0.group(1), body_para_id, 'LEFT',
            prev=TABLE_PARA_PREV, next_val=TABLE_PARA_NEXT
        )

        last_para_match = None
        for m in re.finditer(r'</hh:paraPr>', header_xml):
            last_para_match = m
        if not last_para_match:
            print("WARNING: Could not find paraPr insertion point.")
            return False

        ins = last_para_match.end()
        header_xml = (header_xml[:ins]
                      + '\n        ' + header_pr
                      + '\n        ' + body_pr
                      + header_xml[ins:])

        header_xml = re.sub(
            r'(<hh:paraProperties itemCnt=")(\d+)(")',
            lambda m: m.group(1) + str(int(m.group(2)) + 2) + m.group(3),
            header_xml
        )

        # ── 2b. Create table charPr (11pt, 맑은 고딕) ──────────────
        #   Clone charPr id=0, set height=TABLE_CHAR_HEIGHT, fontRef=0
        max_charpr_id = max(
            (int(m) for m in re.findall(
                r'<hh:charPr id="(\d+)"', header_xml)),
            default=0
        )
        table_charpr_id = max_charpr_id + 1

        charpr_0 = re.search(
            r'(<hh:charPr id="0".*?</hh:charPr>)',
            header_xml, re.DOTALL
        )
        if charpr_0:
            table_charpr = charpr_0.group(1)
            table_charpr = table_charpr.replace(
                'id="0"', f'id="{table_charpr_id}"', 1)
            # Set height to TABLE_CHAR_HEIGHT (11pt)
            table_charpr = re.sub(
                r'(height=")[^"]+(")',
                rf'\g<1>{TABLE_CHAR_HEIGHT}\2', table_charpr, count=1)
            # Set ratio and spacing
            table_charpr = _update_ratio_spacing(
                table_charpr, CHAR_RATIO, CHAR_SPACING)
            # Register and set table font (나눔고딕)
            header_xml, table_font_id = _register_font(
                header_xml, TABLE_FONT_FACE)
            table_charpr = _update_fontref_all_langs(
                table_charpr, table_font_id)

            # Insert after last charPr
            last_charpr_match = None
            for m in re.finditer(r'</hh:charPr>', header_xml):
                last_charpr_match = m
            if last_charpr_match:
                ins = last_charpr_match.end()
                header_xml = (header_xml[:ins]
                              + '\n        ' + table_charpr
                              + header_xml[ins:])
            # Update charProperties itemCnt
            header_xml = re.sub(
                r'(<hh:charProperties itemCnt=")(\d+)(")',
                lambda m: m.group(1) + str(int(m.group(2)) + 1) + m.group(3),
                header_xml
            )
            print(f"Table charPr: id={table_charpr_id}, "
                  f"height={TABLE_CHAR_HEIGHT} "
                  f"({int(TABLE_CHAR_HEIGHT)//100}pt), "
                  f"font={TABLE_FONT_FACE} (id={table_font_id})")
            # Create italic charPr for source/caption text (출처, 자료, ※)
            italic_charpr_id = table_charpr_id + 1
            italic_charpr = table_charpr.replace(
                f'id="{table_charpr_id}"',
                f'id="{italic_charpr_id}"', 1
            )
            italic_charpr = italic_charpr.replace(
                '<hh:underline',
                '<hh:italic />\n        <hh:underline'
            )
            header_xml = header_xml.replace(
                '</hh:charProperties>',
                f'      {italic_charpr}\n      </hh:charProperties>'
            )
            header_xml = re.sub(
                r'(<hh:charProperties itemCnt=")(\d+)(")',
                lambda m: m.group(1) + str(int(m.group(2)) + 1) + m.group(3),
                header_xml
            )
            print(f"Italic charPr: id={italic_charpr_id} (source/caption)")
        else:
            print("WARNING: charPr id=0 not found; table charPr skipped.")
            table_charpr_id = 0
            italic_charpr_id = 0

        # ── 3. Create borderFill for header background color ────────
        max_fill_id = max(
            (int(m) for m in re.findall(
                r'<hh:borderFill id="(\d+)"', header_xml)),
            default=0
        )
        header_fill_id = max_fill_id + 1

        tc_ref = re.search(
            r'<hp:tc\b[^>]*\bborderFillIDRef="(\d+)"', section_xml
        )
        table_fill_id = tc_ref.group(1) if tc_ref else '3'

        base_fill = re.search(
            r'<hh:borderFill id="' + re.escape(table_fill_id)
            + r'".*?</hh:borderFill>',
            header_xml, re.DOTALL
        )
        if not base_fill:
            print(f"WARNING: borderFill id={table_fill_id} not found.")
            return False

        new_fill = base_fill.group(0)
        new_fill = new_fill.replace(
            f'id="{table_fill_id}"', f'id="{header_fill_id}"', 1
        )
        if 'faceColor="none"' in new_fill:
            new_fill = new_fill.replace(
                'faceColor="none"', f'faceColor="{header_color}"', 1
            )
        else:
            new_fill = re.sub(
                r'faceColor="[^"]+"',
                f'faceColor="{header_color}"', new_fill, count=1
            )

        last_fill_match = None
        for m in re.finditer(r'</hh:borderFill>', header_xml):
            last_fill_match = m
        if not last_fill_match:
            print("WARNING: Could not find borderFill insertion point.")
            return False

        ins = last_fill_match.end()
        header_xml = (header_xml[:ins] + '\n      '
                      + new_fill + header_xml[ins:])

        header_xml = re.sub(
            r'(<hh:borderFills itemCnt=")(\d+)(")',
            lambda m: m.group(1) + str(int(m.group(2)) + 1) + m.group(3),
            header_xml
        )

        # ── 4. Update section0.xml: table styling + proportional cols
        tables_modified = 0
        cols_adjusted = 0

        def update_table(match):
            nonlocal tables_modified, cols_adjusted
            tbl = match.group(0)

            sz_m = re.search(r'<hp:sz width="(\d+)"', tbl)
            total_w = int(sz_m.group(1)) if sz_m else 45000

            prop_widths = _calc_proportional_widths(tbl, total_w)
            if prop_widths:
                def update_cell_width(cell_match):
                    cell = cell_match.group(0)
                    col_m = re.search(r'colAddr="(\d+)"', cell)
                    span_m = re.search(r'colSpan="(\d+)"', cell)
                    if not col_m:
                        return cell
                    col = int(col_m.group(1))
                    span = int(span_m.group(1)) if span_m else 1
                    new_w = sum(
                        prop_widths.get(col + i, 5000) for i in range(span)
                    )
                    cell = re.sub(
                        r'(<hp:cellSz width=")(\d+)(")',
                        lambda m: m.group(1) + str(new_w) + m.group(3),
                        cell
                    )
                    return cell

                tbl = re.sub(
                    r'<hp:tc[^>]*>.*?</hp:tc>',
                    update_cell_width, tbl, flags=re.DOTALL
                )
                cols_adjusted += 1

            # ── Table position: center on page ──────────────────────
            tbl = re.sub(
                r'(<hp:pos\b[^>]*\bhorzAlign=")[^"]+(")',
                r'\g<1>CENTER\2', tbl)

            # ── Cell vertical alignment: CENTER (on subList) ────────
            tbl = re.sub(
                r'(<hp:subList\b[^>]*\bvertAlign=")[^"]+(")',
                r'\g<1>CENTER\2', tbl)

            # ── Enable cell margins (hasMargin=1) ────────────────────
            tbl = re.sub(
                r'(<hp:tc\b[^>]*\bhasMargin=")0(")',
                r'\g<1>1\2', tbl)

            # ── Cell margin: top/bottom 1.50mm (425 HWPUNIT) ────────
            tbl = re.sub(
                r'(<hp:cellMargin\b[^/]*)\btop="[^"]+(")',
                r'\g<1>top="425\2', tbl)
            tbl = re.sub(
                r'(<hp:cellMargin\b[^/]*)\bbottom="[^"]+(")',
                r'\g<1>bottom="425\2', tbl)

            first_tr = re.search(
                r'(<hp:tr>)(.*?)(</hp:tr>)', tbl, re.DOTALL
            )
            if not first_tr:
                return tbl

            # ── First row (header): paraPr + borderFill + charPr ──
            first_inner = first_tr.group(2)
            first_inner = re.sub(
                r'paraPrIDRef="\d+"',
                f'paraPrIDRef="{header_para_id}"', first_inner
            )
            first_inner = re.sub(
                r'borderFillIDRef="\d+"',
                f'borderFillIDRef="{header_fill_id}"', first_inner
            )
            # Update charPrIDRef on all runs within header cells
            first_inner = re.sub(
                r'(<hp:run\b[^>]*charPrIDRef=")(\d+)(")',
                rf'\g<1>{table_charpr_id}\3', first_inner
            )
            new_first_tr = (first_tr.group(1)
                            + first_inner + first_tr.group(3))

            # ── Remaining rows (body): paraPr + charPr ─────────────
            after = tbl[first_tr.end():]
            after = re.sub(
                r'paraPrIDRef="\d+"',
                f'paraPrIDRef="{body_para_id}"', after
            )
            # Update charPrIDRef on all runs within body cells
            after = re.sub(
                r'(<hp:run\b[^>]*charPrIDRef=")(\d+)(")',
                rf'\g<1>{table_charpr_id}\3', after
            )

            tables_modified += 1
            return tbl[:first_tr.start()] + new_first_tr + after

        section_xml = re.sub(
            r'<hp:tbl[^>]*>.*?</hp:tbl>',
            update_table, section_xml, flags=re.DOTALL
        )

        # ── 4b. Remove [Figure N] captions ────────────────────────
        #   Text may be split across multiple runs, so join all <hp:t>
        fig_count = 0

        def remove_figure_para(match):
            nonlocal fig_count
            para = match.group(0)
            texts = re.findall(r'<hp:t>(.*?)</hp:t>', para)
            full = ''.join(texts).strip()
            if re.match(r'^\[?Figure \d+\]?\s*$', full):
                fig_count += 1
                return ''
            return para

        section_xml = re.sub(
            r'<hp:p\b[^>]*>.*?</hp:p>',
            remove_figure_para, section_xml, flags=re.DOTALL)
        if fig_count:
            print(f"Figure captions removed: {fig_count}")

        # ── 4c. Source/caption text: italic charPr ──────────────────
        #   "출처:", "자료:", "※", "주:" → italic 11pt 맑은고딕
        if italic_charpr_id:
            source_count = 0
            source_pats = ['출처:', '자료:', '※', '주:']

            def italicize_source(match):
                nonlocal source_count
                para = match.group(0)
                # Check if any source pattern in text
                texts = re.findall(r'<hp:t>(.*?)</hp:t>', para)
                full = ''.join(texts)
                if any(p in full for p in source_pats):
                    para = re.sub(
                        r'charPrIDRef="\d+"',
                        f'charPrIDRef="{italic_charpr_id}"', para)
                    source_count += 1
                return para

            # Only process paragraphs NOT inside tables
            parts = re.split(r'(<hp:tbl[^>]*>.*?</hp:tbl>)',
                             section_xml, flags=re.DOTALL)
            for idx in range(len(parts)):
                if not parts[idx].startswith('<hp:tbl'):
                    parts[idx] = re.sub(
                        r'<hp:p\b[^>]*>.*?</hp:p>',
                        italicize_source, parts[idx], flags=re.DOTALL)
            section_xml = ''.join(parts)
            if source_count:
                print(f"Source italic: {source_count} paragraph(s)")

        # ── 5. Save ────────────────────────────────────────────────
        with open(header_path, 'w', encoding='utf-8') as f:
            f.write(header_xml)
        with open(section_path, 'w', encoding='utf-8') as f:
            f.write(section_xml)

        # ── 6. Repackage ──────────────────────────────────────────
        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zf:
            mimetype_path = os.path.join(tmp_dir, 'mimetype')
            if os.path.exists(mimetype_path):
                zf.write(mimetype_path, 'mimetype',
                         compress_type=zipfile.ZIP_STORED)
            for root_dir, dirs, files in os.walk(tmp_dir):
                for fname in files:
                    if fname == 'mimetype':
                        continue
                    fpath = os.path.join(root_dir, fname)
                    arcname = os.path.relpath(fpath, tmp_dir)
                    zf.write(fpath, arcname)

        print(f"Char shape: ratio={CHAR_RATIO}%, spacing={CHAR_SPACING} "
              f"({charpr_count} charPr(s))")
        print(f"Para spacing: prev={PARA_SPACING_PREV}, next={PARA_SPACING_NEXT} "
              f"({spacing_count} paraPr(s))")
        print(f"List indent: left={LIST_LEFT}, intent={LIST_INTENT} "
              f"({list_count} paraPr(s))")
        print(f"Table styling: {tables_modified} table(s), "
              f"header bg: {header_color}, line spacing: 130%")
        print(f"Table paraPr prev={TABLE_PARA_PREV} for header+body")
        print(f"Body justify: {justify_count} paraPr(s)")
        print(f"Proportional columns: {cols_adjusted} table(s)")
        return True

    except Exception as e:
        print(f"ERROR: {e}")
        import traceback
        traceback.print_exc()
        return False
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)


def main():
    if len(sys.argv) < 2:
        print("Usage: python3 style_tables.py <input.hwpx> [output.hwpx]")
        sys.exit(1)

    input_path = sys.argv[1]
    output_path = sys.argv[2] if len(sys.argv) > 2 else input_path

    if not os.path.exists(input_path):
        print(f"ERROR: File not found: {input_path}")
        sys.exit(1)

    center_table_headers(input_path, output_path)


if __name__ == '__main__':
    main()
