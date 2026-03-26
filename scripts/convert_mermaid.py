#!/usr/bin/env python3
"""
Markdown with Mermaid diagrams → HWPX converter (cross-platform).

Detects mermaid code blocks, converts them to PNG images,
embeds them inline, then converts to HWPX.

Usage:
    python3 convert_mermaid.py <input.md> [output_name]
"""

import os
import sys
import re
import base64
import hashlib
import subprocess
import tempfile
import shutil
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


def _final_charpr_cleanup(hwpx_path):
    """Final pass: force table runs to charPr=12, heading runs to correct charPr.
    Runs AFTER all post-processors to fix any charPr corruption."""
    import zipfile, tempfile, shutil
    extract_dir = tempfile.mkdtemp()
    with zipfile.ZipFile(hwpx_path, 'r') as z:
        z.extractall(extract_dir)
        original_files = z.namelist()

    section_path = os.path.join(extract_dir, 'Contents', 'section0.xml')
    header_path = os.path.join(extract_dir, 'Contents', 'header.xml')
    with open(section_path, 'r', encoding='utf-8') as f:
        xml = f.read()
    with open(header_path, 'r', encoding='utf-8') as f:
        hxml = f.read()

    # Build heading style → charPrIDRef map from header.xml
    heading_map = {}
    for m in re.finditer(
        r'<hh:style\b[^>]*name="[^"]*제목[^"]*"'
        r'[^>]*paraPrIDRef="(\d+)"[^>]*charPrIDRef="(\d+)"', hxml
    ):
        heading_map[m.group(1)] = m.group(2)
    for m in re.finditer(
        r'<hh:style\b[^>]*name="부제목"'
        r'[^>]*paraPrIDRef="(\d+)"[^>]*charPrIDRef="(\d+)"', hxml
    ):
        heading_map[m.group(1)] = m.group(2)

    # 1. Find table charPr ID from header.xml (height=1100, added by style_tables)
    table_charpr_id = '12'  # fallback
    m = re.search(r'<hh:charPr id="(\d+)"[^>]*height="1100"', hxml)
    if m:
        table_charpr_id = m.group(1)
    tbl_fixed = 0
    def fix_table_runs(match):
        nonlocal tbl_fixed
        tbl = match.group(0)
        orig = tbl
        tbl = re.sub(
            r'(<hp:run\b[^>]*charPrIDRef=")(\d+)(")',
            rf'\g<1>{table_charpr_id}\3', tbl)
        if tbl != orig:
            tbl_fixed += 1
        return tbl

    xml = re.sub(r'<hp:tbl[^>]*>.*?</hp:tbl>', fix_table_runs,
                 xml, flags=re.DOTALL)

    # 2. Force heading paragraph runs to correct charPr
    hdg_fixed = 0
    def fix_heading_runs(match):
        nonlocal hdg_fixed
        para = match.group(0)
        pid = match.group(1)
        if pid not in heading_map:
            return para
        correct_charpr = heading_map[pid]
        orig = para
        para = re.sub(
            r'(<hp:run\b[^>]*charPrIDRef=")(\d+)(")',
            rf'\g<1>{correct_charpr}\3', para)
        if para != orig:
            hdg_fixed += 1
        return para

    xml = re.sub(
        r'<hp:p paraPrIDRef="(\d+)"[^>]*>.*?</hp:p>',
        fix_heading_runs, xml, flags=re.DOTALL)

    with open(section_path, 'w', encoding='utf-8') as f:
        f.write(xml)

    # Repackage
    temp = hwpx_path + '.tmp'
    with zipfile.ZipFile(temp, 'w', zipfile.ZIP_DEFLATED) as zout:
        for fname in original_files:
            fpath = os.path.join(extract_dir, fname)
            if os.path.exists(fpath):
                zout.write(fpath, fname)
    shutil.move(temp, hwpx_path)
    shutil.rmtree(extract_dir)
    if tbl_fixed or hdg_fixed:
        print(f"Final cleanup: table={tbl_fixed} tbl(s), "
              f"heading={hdg_fixed} para(s)")


def _get_env():
    """Get venv Python and pypandoc-hwpx paths."""
    env_mod = _load_module('env_detect', 'env_detect.py')
    venv_root = env_mod._find_venv_root()
    python_path = env_mod.find_venv_python(venv_root)
    pypandoc_path = env_mod.find_pypandoc_hwpx(venv_root)
    return python_path, pypandoc_path


def get_image_size(img_path):
    """Return (width, height) of an image file."""
    try:
        from PIL import Image
        with Image.open(img_path) as im:
            return im.size
    except Exception:
        return (600, 400)


def calc_display_width(natural_px, page_px=600, scale=3):
    """
    Calculate display width for A4 portrait layout.
    Divides by scale to get logical size, caps at page width.
    """
    logical_px = natural_px // scale
    return min(logical_px, page_px)


def mermaid_to_png_mmdc(mermaid_code, output_path):
    """
    Convert mermaid code to PNG via mmdc (Mermaid CLI).
    Viewport 600px = HWPX page content width (159mm) 1:1 match.
    Scale 3 = 3x resolution only, layout stays at 600px.
    FontSize 15px ~ 11pt (matches HWPX body text).
    """
    with tempfile.NamedTemporaryFile(suffix='.mmd', mode='w',
                                     encoding='utf-8', delete=False) as f:
        f.write(mermaid_code)
        mmd_path = f.name

    cfg = tempfile.NamedTemporaryFile(suffix='.json', mode='w',
                                      encoding='utf-8', delete=False)
    cfg.write('{"theme":"default","themeVariables":{"fontSize":"15px"}}')
    cfg.close()
    cfg_path = cfg.name

    try:
        result = subprocess.run(
            ['npx', '--yes', '@mermaid-js/mermaid-cli',
             '-i', mmd_path, '-o', output_path,
             '-b', 'white',
             '-w', '600',
             '-s', '3',
             '-c', cfg_path],
            capture_output=True, text=True, timeout=60
        )
        return result.returncode == 0 and os.path.exists(output_path)
    except Exception as e:
        print(f"   [mmdc error] {e}")
        return False
    finally:
        os.unlink(mmd_path)
        os.unlink(cfg_path)


def mermaid_to_png_api(mermaid_code, output_path):
    """Convert mermaid code to PNG via mermaid.ink API (requires internet)."""
    import urllib.request
    import ssl

    encoded = base64.urlsafe_b64encode(
        mermaid_code.encode('utf-8')
    ).decode('utf-8')
    url = f"https://mermaid.ink/img/{encoded}?type=png&bgColor=white"

    try:
        ctx = ssl.create_default_context()
        req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
        with urllib.request.urlopen(req, context=ctx, timeout=15) as resp:
            with open(output_path, 'wb') as f:
                f.write(resp.read())
        return os.path.getsize(output_path) > 1000
    except Exception as e:
        print(f"   [API error] {e}")
        return False


def _parse_mermaid_elements(mermaid_code):
    """Parse mermaid code into nodes, connections, styles."""
    lines = mermaid_code.strip().split('\n')
    first_line = lines[0].strip()
    body_lines = lines[1:]

    nodes_order = []
    node_defs = {}
    connections = []
    styles = []
    has_subgraph = False

    for line in body_lines:
        stripped = line.strip()
        if not stripped:
            continue
        if stripped.startswith('style '):
            styles.append(stripped)
            continue
        if 'subgraph' in stripped:
            has_subgraph = True
            break

        arrow_match = re.findall(
            r'(\w+)(\["[^"]*"\])?\s*(-->|---|-\.->|==>)\s*'
            r'(?:\|([^|]*)\|)?\s*(\w+)(\["[^"]*"\])?',
            stripped
        )
        if arrow_match:
            for src_id, src_def, arrow, label, dst_id, dst_def in arrow_match:
                if src_id not in node_defs:
                    nodes_order.append(src_id)
                    node_defs[src_id] = src_def if src_def else ''
                elif src_def and not node_defs[src_id]:
                    node_defs[src_id] = src_def
                if dst_id not in node_defs:
                    nodes_order.append(dst_id)
                    node_defs[dst_id] = dst_def if dst_def else ''
                elif dst_def and not node_defs[dst_id]:
                    node_defs[dst_id] = dst_def
                connections.append((src_id, dst_id, arrow, label or ''))
        else:
            node_match = re.match(r'^\s*(\w+)(\["[^"]*"\])\s*$', stripped)
            if node_match:
                nid, ndef = node_match.group(1), node_match.group(2)
                if nid not in node_defs:
                    nodes_order.append(nid)
                node_defs[nid] = ndef

    return first_line, nodes_order, node_defs, connections, styles, has_subgraph


def _build_grid_layout(nodes_order, node_defs, connections, styles, cols=2):
    """Build a grid layout mermaid code with subgraph rows."""
    n = len(nodes_order)
    rows_count = (n + cols - 1) // cols

    result = ['graph TB']

    for r in range(rows_count):
        row_nodes = nodes_order[r * cols:(r + 1) * cols]
        sg_id = f'_row{r}'
        result.append(f'    subgraph {sg_id}[" "]')
        result.append(f'        direction LR')
        for nid in row_nodes:
            ndef = node_defs.get(nid, '')
            result.append(f'        {nid}{ndef}')
        result.append(f'    end')

    for r in range(rows_count - 1):
        result.append(f'    _row{r} --> _row{r + 1}')

    row_of = {}
    for r in range(rows_count):
        for nid in nodes_order[r * cols:(r + 1) * cols]:
            row_of[nid] = r

    for src, dst, arrow, label in connections:
        if src in row_of and dst in row_of and row_of[src] == row_of[dst]:
            lbl = f'|{label}|' if label else ''
            result.append(f'    {src} {arrow} {lbl}{dst}')

    for s in styles:
        result.append(f'    {s}')

    return '\n'.join(result)


def _fix_quadrant_korean(mermaid_code):
    """Quote Korean text in quadrantChart to fix mmdc lexer errors."""
    if not mermaid_code.strip().startswith('quadrantChart'):
        return mermaid_code, False
    lines = mermaid_code.split('\n')
    fixed = []
    changed = False
    # Regex for CJK characters
    has_cjk = re.compile(r'[\u3000-\u9fff\uac00-\ud7af]')
    for line in lines:
        stripped = line.strip()
        if stripped == 'quadrantChart':
            fixed.append(line)
            continue
        # title, x-axis, y-axis, quadrant-N: quote the value part
        m = re.match(r'^(\s*)(title|x-axis|y-axis|quadrant-\d)\s+(.+)$',
                     line)
        if m and has_cjk.search(m.group(3)):
            indent, key, val = m.group(1), m.group(2), m.group(3)
            # Handle "A --> B" axis format
            if '-->' in val and key in ('x-axis', 'y-axis'):
                parts = val.split('-->')
                parts = [f'"{p.strip()}"' if has_cjk.search(p)
                         and not p.strip().startswith('"')
                         else p.strip() for p in parts]
                val = ' --> '.join(parts)
            elif not val.startswith('"'):
                val = f'"{val}"'
            fixed.append(f'{indent}{key} {val}')
            changed = True
            continue
        # Data points: "Label: [x, y]"
        m = re.match(r'^(\s*)([^:\[]+):\s*(\[.+\])\s*$', line)
        if m and has_cjk.search(m.group(2)):
            indent, label, coords = m.group(1), m.group(2).strip(), m.group(3)
            if not label.startswith('"'):
                label = f'"{label}"'
            fixed.append(f'{indent}{label}: {coords}')
            changed = True
            continue
        fixed.append(line)
    return '\n'.join(fixed), changed


def _optimize_mermaid_for_portrait(mermaid_code):
    """
    Optimize mermaid code for A4 portrait:
    0. quadrantChart Korean fix (quote CJK labels)
    1. \\n → <br/> in node labels
    2. graph LR → grid layout (subgraph rows + direction LR)
    """
    optimized = mermaid_code
    changed = []

    # 0. quadrantChart Korean fix
    optimized, qc_fixed = _fix_quadrant_korean(optimized)
    if qc_fixed:
        changed.append('quadrant-kr-fix')

    # 1. \n → <br/>
    def replace_newline_in_labels(m):
        return m.group(0).replace('\\n', '<br/>')
    new_code = re.sub(r'\["[^"]*\\n[^"]*"\]',
                      replace_newline_in_labels, optimized)
    if new_code != optimized:
        optimized = new_code
        changed.append('\\n→<br/>')

    # 2. LR/RL detection
    lr_pattern = re.compile(r'^(graph|flowchart)\s+(LR|RL)\b', re.MULTILINE)
    if not lr_pattern.search(optimized):
        return optimized, changed

    first_line, nodes_order, node_defs, connections, styles, has_subgraph = \
        _parse_mermaid_elements(optimized)

    if has_subgraph or len(nodes_order) <= 3:
        optimized = lr_pattern.sub(r'\1 TD', optimized)
        changed.append('LR→TD')
        return optimized, changed

    cols = 2
    if len(nodes_order) >= 8:
        cols = 3
    grid_code = _build_grid_layout(
        nodes_order, node_defs, connections, styles, cols
    )
    changed.append(f'LR→grid({cols}col)')
    return grid_code, changed


def convert_mermaid_blocks(content, img_dir):
    """
    Convert mermaid blocks to PNG and replace with inline image references.
    Returns: (converted_content, success_count, fail_count)
    """
    pattern = re.compile(r'```mermaid\n(.*?)\n```', re.DOTALL)
    blocks = pattern.findall(content)

    if not blocks:
        return content, 0, 0

    print(f"\nFound {len(blocks)} mermaid block(s)\n")

    success_count = 0
    fail_count = 0
    replacements = []

    for i, mermaid_code in enumerate(blocks, 1):
        mermaid_code, changes = _optimize_mermaid_for_portrait(mermaid_code)
        first_line = mermaid_code.strip().split('\n')[0].strip()
        change_str = f" [{', '.join(changes)}]" if changes else ""
        code_hash = hashlib.md5(mermaid_code.encode()).hexdigest()[:8]
        img_filename = f'diagram_{i:02d}_{code_hash}.png'
        img_path = os.path.join(img_dir, img_filename)

        print(f"[{i}/{len(blocks)}] {first_line}{change_str}")

        converted = mermaid_to_png_mmdc(mermaid_code, img_path)
        if converted:
            print(f"   OK: mmdc ({os.path.getsize(img_path):,} bytes)")

        if not converted:
            print(f"   Trying mermaid.ink API...")
            converted = mermaid_to_png_api(mermaid_code, img_path)
            if converted:
                print(f"   OK: API ({os.path.getsize(img_path):,} bytes)")

        if converted:
            natural_w, _ = get_image_size(img_path)
            display_px = calc_display_width(natural_w)
            replacements.append(
                f'**[Figure {i}]** ![]({img_path}){{width={display_px}}}'
            )
            success_count += 1
        else:
            print(f"   FAILED: using text placeholder")
            replacements.append(
                f'*[Figure {i}: {first_line} — diagram conversion failed]*'
            )
            fail_count += 1

    idx = 0
    def replace_block(match):
        nonlocal idx
        replacement = replacements[idx]
        idx += 1
        return replacement

    new_content = pattern.sub(replace_block, content)
    return new_content, success_count, fail_count


def convert_to_hwpx(md_path, output_path):
    """Convert preprocessed Markdown to HWPX via pypandoc-hwpx."""
    python_path, pypandoc_path = _get_env()

    if not python_path or not pypandoc_path:
        print("ERROR: pypandoc-hwpx is not installed.")
        return False

    result = subprocess.run(
        [python_path, pypandoc_path, md_path, '-o', output_path],
        capture_output=True, text=True
    )
    return result.returncode == 0


def main():
    if len(sys.argv) < 2:
        print("Usage: python3 convert_mermaid.py <input.md> [output_name]")
        sys.exit(1)

    input_md = sys.argv[1]
    if not os.path.exists(input_md):
        print(f"ERROR: File not found: {input_md}")
        sys.exit(1)

    custom_name = sys.argv[2] if len(sys.argv) > 2 else None

    print(f"Mermaid -> HWPX conversion started\n")
    print(f"Input: {input_md}")

    if custom_name:
        base_name = custom_name
    else:
        stem = os.path.splitext(os.path.basename(input_md))[0]
        stem = re.sub(r'[^\w\s-]', '', stem)
        stem = re.sub(r'[-\s]+', '_', stem)
        stem = unicodedata.normalize('NFC', stem)
        date_str = datetime.now().strftime('%Y%m%d')
        base_name = f"{date_str}_{stem}"

    os.makedirs('output', exist_ok=True)
    output_hwpx = os.path.join('output', f'{base_name}.hwpx')

    tmp_dir = tempfile.mkdtemp(prefix='mermaid_hwpx_')
    img_dir = os.path.join(tmp_dir, 'images')
    os.makedirs(img_dir)

    try:
        with open(input_md, 'r', encoding='utf-8') as f:
            content = f.read()

        new_content, success, fail = convert_mermaid_blocks(content, img_dir)

        tmp_md = os.path.join(tmp_dir, 'preprocessed.md')
        with open(tmp_md, 'w', encoding='utf-8') as f:
            f.write(new_content)

        print(f"\nDiagrams: {success} OK / {fail} failed")
        print(f"\nConverting to HWPX...")

        ok = convert_to_hwpx(tmp_md, output_hwpx)

        if ok and os.path.exists(output_hwpx):
            # Post-processing: table styling
            try:
                style_mod = _load_module('style_tables', 'style_tables.py')
                style_mod.center_table_headers(output_hwpx)
            except Exception as e:
                print(f"WARNING: Table styling failed: {e}")

            # Post-processing: bold formatting
            try:
                bold_mod = _load_module('apply_bold', 'apply_bold.py')
                bold_mod.apply_bold(input_md, output_hwpx)
            except Exception as e:
                print(f"WARNING: Bold formatting failed: {e}")

            # Post-processing: superscript footnotes
            try:
                sup_mod = _load_module('apply_superscript',
                                       'apply_superscript.py')
                sup_mod.apply_superscript(output_hwpx)
            except Exception as e:
                print(f"WARNING: Superscript formatting failed: {e}")

            # Post-processing: final charPr cleanup
            # Restore table charPr and heading charPr after all transforms
            try:
                _final_charpr_cleanup(output_hwpx)
            except Exception as e:
                print(f"WARNING: Final charPr cleanup failed: {e}")

            # Post-processing: inject image XML references
            # pypandoc-hwpx embeds images in BinData/ but fails to create
            # <hp:pic> XML tags in section0.xml for complex documents.
            # This replaces [Figure N] placeholder paragraphs with images.
            try:
                inject_mod = _load_module('inject_images',
                                          'inject_images.py')
                inject_mod.inject_images(output_hwpx)
            except Exception as e:
                print(f"WARNING: Image injection failed: {e}")

            size = os.path.getsize(output_hwpx)
            print(f"\n{'=' * 60}")
            print(f"Conversion complete!")
            print(f"{'=' * 60}")
            print(f"\nFile: {os.path.abspath(output_hwpx)}")
            print(f"Size: {size:,} bytes ({size/1024:.1f}KB)")
            print(f"Images: {success} embedded\n")
        else:
            print("ERROR: HWPX conversion failed")

    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)


if __name__ == '__main__':
    main()
