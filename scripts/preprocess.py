#!/usr/bin/env python3
"""
Markdown preprocessor for HWPX conversion (cross-platform).

Cleans up Markdown content to prevent Hancom crashes and formatting issues:
1. □/○ symbols → Markdown list syntax
2. Special characters (■—「」★☆×→·) → ASCII equivalents
3. Number comma removal (4,825 → 4825)
3b. Footnote/reference bracket conversion ([N] → N))
4. Empty table cell filling
5. **key**: value line break insertion
6. List blank line insertion
7. NFC normalization (for macOS Korean text)

Usage:
    python3 preprocess.py <input.md> [output.md]
    (default output: /tmp/hwpx_cleaned.md)
"""

import os
import sys
import re
import unicodedata


def preprocess_markdown(content):
    """
    Apply all preprocessing steps to Markdown content.
    Returns cleaned content string.
    """
    # 0. □/○ → Markdown list conversion (must run first!)
    content = re.sub(r'^ {0,2}□\s*', '- ', content, flags=re.MULTILINE)
    content = re.sub(r'^ {2,4}○\s*', '  - ', content, flags=re.MULTILINE)

    # 1. Special character replacement
    content = content.replace('■', '-')
    content = content.replace('—', '--')
    content = content.replace('→', '->')
    content = content.replace('·', '-')
    content = content.replace('×', 'x')
    content = content.replace('「', '"').replace('」', '"')

    # 1b. Smart-quote prevention: convert ASCII quotes to Unicode curly quotes
    #     Pandoc's smart extension converts "text" into Quoted AST nodes,
    #     but pypandoc-hwpx drops Quoted nodes. Unicode curly quotes are
    #     treated as literal Str nodes, preserving the text.
    #     Skip YAML frontmatter (--- delimited) to avoid breaking YAML strings.
    if content.startswith('---'):
        fm_end = content.find('\n---', 3)
        if fm_end > 0:
            fm_end += 4  # include closing ---\n
            frontmatter = content[:fm_end]
            body = content[fm_end:]
        else:
            frontmatter = ''
            body = content
    else:
        frontmatter = ''
        body = content
    body = re.sub(r'"([^"\n]+)"', '\u201C\\1\u201D', body)
    body = re.sub(r"(?<![a-zA-Z])'([^'\n]+)'", '\u2018\\1\u2019', body)
    content = frontmatter + body

    # 2. ★☆ rating → numeric
    content = re.sub(r'★{5}', '5/5', content)
    content = re.sub(r'★{4}☆', '4/5', content)
    content = re.sub(r'★{3}☆{2}', '3/5', content)
    content = re.sub(r'★{2}☆{3}', '2/5', content)
    content = re.sub(r'★☆{4}', '1/5', content)
    content = content.replace('★', '*').replace('☆', '*')

    # 3. Number comma removal (4,825 → 4825)
    content = re.sub(r'(\d),(\d)', r'\1\2', content)

    # 3b. Footnote/reference bracket conversion: [N] → N)
    #     Line-start [N] → N\) (escaped to prevent pandoc ordered list)
    #     Mid-line [N] → N) (inline footnote, not a list)
    #     Skips markdown links like [text](url) by negative lookahead for '('
    content = re.sub(r'^(\[(\d{1,3})\]) ', r'\2\\) ', content, flags=re.MULTILINE)
    content = re.sub(r'\[(\d{1,3})\](?!\()', r'\1)', content)

    # 4. Empty table cell filling (most common crash cause)
    lines = content.split('\n')
    fixed_lines = []
    for line in lines:
        if '|' in line and re.search(r'\| *\|', line):
            if not re.match(r'^\|[\s\-:|]+\|$', line.strip()):
                while re.search(r'\| *\|', line):
                    line = re.sub(r'\| *\|', '| **-** |', line)
        fixed_lines.append(line)
    content = '\n'.join(fixed_lines)

    # 5. **key**: value pattern — insert blank lines for paragraph separation
    lines = content.split('\n')
    result_lines = []
    for i, line in enumerate(lines):
        prev_line = lines[i - 1] if i > 0 else ''
        if re.match(r'^\s*\*\*[^*]+\*\*:', line) and prev_line.strip():
            result_lines.append('')
        result_lines.append(line)
    content = '\n'.join(result_lines)

    # 6. List blank line insertion
    #    (paragraph followed directly by list item → merged in HWPX)
    lines = content.split('\n')
    result_lines2 = []
    for i, line in enumerate(lines):
        prev = lines[i - 1] if i > 0 else ''
        is_list = re.match(r'^(\s*[-*+]\s|\s*\d+\.\s)', line)
        prev_is_list = re.match(r'^(\s*[-*+]\s|\s*\d+\.\s)', prev)
        prev_is_empty = (prev.strip() == '')
        prev_is_heading = re.match(r'^#+\s', prev)
        if (is_list and not prev_is_list
                and not prev_is_empty and not prev_is_heading):
            result_lines2.append('')
        result_lines2.append(line)
    content = '\n'.join(result_lines2)

    # 7. NFC normalization
    content = unicodedata.normalize('NFC', content)

    return content


def preprocess_file(input_path, output_path=None):
    """
    Read a Markdown file, preprocess it, and write to output.
    Returns the output file path.
    """
    if output_path is None:
        import tempfile
        output_path = os.path.join(tempfile.gettempdir(), 'hwpx_cleaned.md')

    with open(input_path, 'r', encoding='utf-8') as f:
        content = f.read()

    cleaned = preprocess_markdown(content)

    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(cleaned)

    print(f"Preprocessed: {input_path}")
    print(f"Output: {output_path}")
    return output_path


def main():
    if len(sys.argv) < 2:
        print("Usage: python3 preprocess.py <input.md> [output.md]")
        sys.exit(1)

    input_file = sys.argv[1]
    output_file = sys.argv[2] if len(sys.argv) > 2 else None

    if not os.path.exists(input_file):
        print(f"ERROR: File not found: {input_file}")
        sys.exit(1)

    preprocess_file(input_file, output_file)


if __name__ == '__main__':
    main()
