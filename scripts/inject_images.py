#!/usr/bin/env python3
"""
Post-processor to inject PNG images into HWPX files.

pypandoc-hwpx embeds image files in BinData/ and declares them in
content.hpf, but fails to create <hp:pic> XML references in section0.xml.
This script finds image declarations and injects proper XML tags at
placeholder positions marked by [Figure N] text.

Usage:
    python3 inject_images.py <input.hwpx> [output.hwpx]
"""

import os
import sys
import re
import zipfile
import shutil
import tempfile
import time
import random


def _get_image_size_from_file(img_path):
    """Get image dimensions from file."""
    try:
        from PIL import Image
        with Image.open(img_path) as im:
            return im.size
    except Exception:
        return (600, 400)


def _px_to_hwpunit(px):
    """Convert pixels to HWP units (96 DPI)."""
    return int(px * (25.4 * 283.465) / 96.0)


def _build_pic_xml(binary_item_id, width_hwp, height_hwp, char_pr_id=0):
    """Build <hp:pic> XML for an image."""
    ts = int(time.time() * 1000)
    rnd = random.randint(0, 999999)
    pic_id = str(ts % 100000000 + rnd)
    inst_id = str(random.randint(10000000, 99999999))

    # Cap width at 150mm (A4 content width ~159mm, leave margin)
    MAX_W = int(150 * 283.465)
    if width_hwp > MAX_W:
        ratio = MAX_W / width_hwp
        width_hwp = MAX_W
        height_hwp = int(height_hwp * ratio)

    xml = (
        f'<hp:run charPrIDRef="{char_pr_id}">'
        f'<hp:pic id="{pic_id}" zOrder="0" numberingType="NONE" '
        f'textWrap="TOP_AND_BOTTOM" textFlow="BOTH_SIDES" lock="0" '
        f'dropcapstyle="None" href="" groupLevel="0" instid="{inst_id}" '
        f'reverse="0">'
        f'<hp:offset x="0" y="0"/>'
        f'<hp:orgSz width="{width_hwp}" height="{height_hwp}"/>'
        f'<hp:curSz width="{width_hwp}" height="{height_hwp}"/>'
        f'<hp:flip horizontal="0" vertical="0"/>'
        f'<hp:rotationInfo angle="0" centerX="0" centerY="0" '
        f'rotateimage="1"/>'
        f'<hp:renderingInfo>'
        f'<hc:transMatrix e1="1" e2="0" e3="0" e4="0" e5="1" e6="0"/>'
        f'<hc:scaMatrix e1="1" e2="0" e3="0" e4="0" e5="1" e6="0"/>'
        f'<hc:rotMatrix e1="1" e2="0" e3="0" e4="0" e5="1" e6="0"/>'
        f'</hp:renderingInfo>'
        f'<hc:img binaryItemIDRef="{binary_item_id}" bright="0" '
        f'contrast="0" effect="REAL_PIC" alpha="0"/>'
        f'<hp:imgRect>'
        f'<hc:pt0 x="0" y="0"/>'
        f'<hc:pt1 x="{width_hwp}" y="0"/>'
        f'<hc:pt2 x="{width_hwp}" y="{height_hwp}"/>'
        f'<hc:pt3 x="0" y="{height_hwp}"/>'
        f'</hp:imgRect>'
        f'<hp:imgClip left="0" right="0" top="0" bottom="0"/>'
        f'<hp:inMargin left="0" right="0" top="0" bottom="0"/>'
        f'<hp:imgDim dimwidth="0" dimheight="0"/>'
        f'<hp:effects/>'
        f'<hp:sz width="{width_hwp}" widthRelTo="ABSOLUTE" '
        f'height="{height_hwp}" heightRelTo="ABSOLUTE" protect="0"/>'
        f'<hp:pos treatAsChar="1" affectLSpacing="0" flowWithText="1" '
        f'allowOverlap="1" holdAnchorAndSO="0" vertRelTo="PARA" '
        f'horzRelTo="COLUMN" vertAlign="TOP" horzAlign="LEFT" '
        f'vertOffset="0" horzOffset="0"/>'
        f'<hp:outMargin left="0" right="0" top="0" bottom="0"/>'
        f'<hp:shapeComment/>'
        f'</hp:pic>'
        f'</hp:run>'
    )
    return xml


def inject_images(hwpx_path, output_path=None):
    """
    Find image declarations in HWPX, inject <hp:pic> XML into section0.xml.
    Replaces [Figure N] placeholder paragraphs with image paragraphs.
    """
    if output_path is None:
        output_path = hwpx_path

    tmp_dir = tempfile.mkdtemp(prefix='hwpx_inject_')

    try:
        with zipfile.ZipFile(hwpx_path, 'r') as zf:
            zf.extractall(tmp_dir)
            original_files = zf.namelist()

        # 1. Find image files in BinData/
        bindata_images = []
        for fname in original_files:
            if fname.startswith('BinData/') and fname.lower().endswith(
                    ('.png', '.jpg', '.jpeg', '.gif', '.bmp')):
                img_id = os.path.splitext(os.path.basename(fname))[0]
                img_path = os.path.join(tmp_dir, fname)
                px_w, px_h = _get_image_size_from_file(img_path)
                bindata_images.append({
                    'id': img_id,
                    'fname': fname,
                    'path': img_path,
                    'width_hwp': _px_to_hwpunit(px_w),
                    'height_hwp': _px_to_hwpunit(px_h),
                })

        if not bindata_images:
            print("No images found in BinData/")
            return 0

        # 2. Read section0.xml
        section_path = os.path.join(tmp_dir, 'Contents', 'section0.xml')
        with open(section_path, 'r', encoding='utf-8') as f:
            xml = f.read()

        # 3. Check if images already have references
        existing_refs = re.findall(r'binaryItemIDRef="([^"]+)"', xml)
        if existing_refs:
            print(f"Images already referenced ({len(existing_refs)}), "
                  f"skipping injection")
            return 0

        # 4. Find [Figure N] placeholder paragraphs and replace with images
        #    Match paragraphs containing "Figure N" text
        injected = 0
        img_idx = 0

        def replace_figure_para(match):
            nonlocal injected, img_idx
            para = match.group(0)
            # Extract all text from paragraph
            texts = re.findall(r'<hp:t>(.*?)</hp:t>', para)
            full_text = ''.join(texts).strip()

            # Match [Figure N] or Figure N patterns
            if not re.search(r'\[?Figure \d+\]?', full_text):
                return para

            if img_idx >= len(bindata_images):
                return para

            img = bindata_images[img_idx]
            img_idx += 1
            injected += 1

            # Extract paraPrIDRef and styleIDRef from original paragraph
            para_pr = re.search(r'paraPrIDRef="(\d+)"', para)
            style_id = re.search(r'styleIDRef="(\d+)"', para)
            ppr = para_pr.group(1) if para_pr else '0'
            sid = style_id.group(1) if style_id else '0'

            # Build image paragraph
            pic_xml = _build_pic_xml(img['id'],
                                     img['width_hwp'],
                                     img['height_hwp'])

            return (f'<hp:p paraPrIDRef="{ppr}" styleIDRef="{sid}" '
                    f'pageBreak="0" columnBreak="0" merged="0">'
                    f'{pic_xml}</hp:p>')

        xml = re.sub(r'<hp:p\b[^>]*>.*?</hp:p>',
                     replace_figure_para, xml, flags=re.DOTALL)

        if injected == 0:
            # No Figure placeholders found, try inserting at end of document
            # as standalone paragraphs
            print("No [Figure N] placeholders found")
            shutil.rmtree(tmp_dir)
            return 0

        # 5. Write modified section0.xml
        with open(section_path, 'w', encoding='utf-8') as f:
            f.write(xml)

        # 6. Repackage HWPX
        with zipfile.ZipFile(output_path, 'w',
                             zipfile.ZIP_DEFLATED) as zout:
            for fname in original_files:
                fpath = os.path.join(tmp_dir, fname)
                if os.path.exists(fpath):
                    if fname == 'mimetype':
                        zout.write(fpath, fname,
                                   compress_type=zipfile.ZIP_STORED)
                    else:
                        zout.write(fpath, fname)

        print(f"Images injected: {injected}/{len(bindata_images)}")
        return injected

    finally:
        shutil.rmtree(tmp_dir)


def main():
    if len(sys.argv) < 2:
        print("Usage: python3 inject_images.py <input.hwpx> [output.hwpx]")
        sys.exit(1)

    hwpx_path = sys.argv[1]
    output_path = sys.argv[2] if len(sys.argv) > 2 else None

    if not os.path.exists(hwpx_path):
        print(f"ERROR: {hwpx_path} not found")
        sys.exit(1)

    inject_images(hwpx_path, output_path)


if __name__ == '__main__':
    main()
