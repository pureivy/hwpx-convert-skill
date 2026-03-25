---
name: hwpx-convert
description: Converts Markdown files to Korean Hancom Office HWPX documents with automatic crash prevention, table styling (header background color, proportional column widths, 130% line spacing), bold formatting, and mermaid diagram-to-PNG conversion. Cross-platform (macOS, Linux, Windows). Use this skill whenever the user wants to convert Markdown to HWPX/HWP, create a Hancom (한컴/한글) document from Markdown, or export to Korean word processor format. Trigger phrases include "한글로 변환", "HWPX로 변환", "hwpx 만들어", "한글 문서로", "한컴 파일로", "HWP로 만들어", or any mention of converting .md to .hwpx/.hwp format. Also triggers when a user has problematic content (special characters like ★☆□○■, empty table cells, mermaid diagrams) that needs safe conversion to Hancom format. This skill does NOT handle PDF, DOCX, or other non-Hancom formats, and does NOT convert FROM HWPX to Markdown (only Markdown → HWPX).
---

# hwpx-convert

Converts Markdown to HWPX (Hancom Office format) with automatic preprocessing, post-processing, and mermaid diagram support. Works on macOS, Linux, and Windows.

## How It Works

The conversion pipeline has 5 stages:

1. **Analyze** — Detect mermaid blocks, special characters, empty table cells
2. **Preprocess** — Clean the Markdown (to `/tmp`, never modifying the original)
3. **Convert** — Run pypandoc-hwpx to produce the HWPX file
4. **Post-process** — Apply table styling, bold formatting, XML fixes
5. **Open** — Launch the file in Hancom Office (if available)

## Setup (First Time Only)

Before first use, the user needs three things installed:

### 1. Pandoc (system-level)
```bash
# macOS
brew install pandoc

# Ubuntu/Debian
sudo apt install pandoc

# Windows
choco install pandoc
# or download from https://pandoc.org/installing.html
```

### 2. Python virtual environment with pypandoc-hwpx
```bash
# Create venv in the project directory
python3 -m venv hwpx_env

# Activate
source hwpx_env/bin/activate    # macOS/Linux
hwpx_env\Scripts\activate       # Windows

# Install
pip install pypandoc-hwpx
```

### 3. Node.js (optional, for mermaid diagrams)
```bash
# Only needed if the document contains ```mermaid blocks
# If not installed, the mermaid.ink online API is used as fallback
brew install node    # macOS
sudo apt install nodejs npm    # Ubuntu
```

### Verify setup
```bash
python3 <skill-scripts-dir>/env_detect.py
```

## Conversion Workflow

### Step 1: Analyze the input file

Determine the file's characteristics before converting. This decides which conversion script to use.

```bash
# Check for mermaid blocks
grep -c '```mermaid' INPUT_FILE.md

# Check for problematic patterns
grep -n '[★☆×→·□○■—「」]' INPUT_FILE.md
grep -n '| *|' INPUT_FILE.md
```

### Step 2: Preprocess

Run the preprocessor to create a clean copy. The original file is never modified.

```bash
python3 <skill-scripts-dir>/preprocess.py INPUT_FILE.md /tmp/hwpx_cleaned.md
```

The preprocessor handles:
- □/○ → Markdown list conversion (most important, must run first)
- Special characters (■—「」★☆×→·) → ASCII equivalents
- Quoted text ("text", 'text') → Unicode curly quotes (prevents text loss)
- Empty table cells → filled with **-** (prevents Hancom crashes)
- **key**: value patterns → blank line insertion for paragraph separation
- List items after paragraphs → blank line insertion
- NFC normalization for Korean text

### Step 3: Convert

Choose the right converter based on mermaid presence:

**No mermaid blocks:**
```bash
python3 <skill-scripts-dir>/convert.py /tmp/hwpx_cleaned.md "OutputName"
```

**Has mermaid blocks:**
```bash
python3 <skill-scripts-dir>/convert_mermaid.py /tmp/hwpx_cleaned.md "OutputName"
```

Both converters automatically apply post-processing:
- Body text: justify alignment (headings excluded)
- Table headers: center alignment + #BDD7EE background + 130% line spacing
- Table body: left alignment + 130% line spacing
- Table text: 맑은 고딕 11pt (charPr id auto-detected)
- Table paragraph spacing: prev=0, next=0
- Table position: center on page (horzAlign="CENTER")
- Cell vertical alignment: center (subList vertAlign="CENTER")
- Cell margin: top/bottom 1.50mm (425 HWPUNIT), hasMargin=1 enabled
- Column widths: proportional to content (Korean chars weighted 1.8x)
- Bold: **text** from original Markdown applied as charPr

Output goes to `output/OutputName.hwpx`.

### Step 4: Fix URLs (if needed)

If the document contains URLs with `&` query parameters:

```bash
# Check for unescaped &
unzip -p "output/OutputName.hwpx" Contents/section0.xml 2>/dev/null | grep -q 'https.*&.*='

# Fix if needed
python3 <skill-scripts-dir>/fix_xml.py "output/OutputName.hwpx" "output/OutputName_temp.hwpx"
mv "output/OutputName_temp.hwpx" "output/OutputName.hwpx"
```

### Step 5: Open in Hancom

```bash
# macOS
open "output/OutputName.hwpx"

# Linux (if Hancom is installed)
xdg-open "output/OutputName.hwpx"

# Windows
start "output\OutputName.hwpx"
```

## Crash Prevention Guide

These patterns cause Hancom to crash or produce corrupted files. The preprocessor handles most of them, but it helps to understand what they are.

| Priority | Pattern | Symptom | Fix |
|----------|---------|---------|-----|
| 1 | Empty table cells `\| text \| \| \|` | Hancom crash | Fill with **-** or move to body text |
| 2 | Raw mermaid code blocks | Hancom crash | Use convert_mermaid.py (PNG conversion) |
| 3 | Special chars ★☆×→·■—「」 | Crash or garbled text | Replace with ASCII equivalents |
| 3.5 | □/○ symbol lists | Content dropped or merged | Convert to `- ` / `  - ` |
| 4 | URLs with `&` | "Damaged file" error | Run fix_xml.py post-processor |
| 5 | Consecutive `**key**: value` lines | Lines merged into one | Insert blank lines between them |
| 6 | List after paragraph (no blank line) | List merged with paragraph | Insert blank line before list |
| 7 | Quoted text `"text"` or `'text'` | Text inside quotes disappears | Convert to Unicode curly quotes (auto) |

## Output Name Rules

If no output name is specified:
1. Extract title from first `# ` heading
2. Remove special characters, replace spaces with `_`
3. Prefix with date: `YYYYMMDD_Title`

## Environment Variable

Set `HWPX_VENV_PATH` to point to a custom virtual environment location:
```bash
export HWPX_VENV_PATH=/path/to/my/hwpx_env
```

If not set, the scripts search for `hwpx_env/` in:
1. The skill's parent directory
2. The current working directory
3. The user's home directory

## Bundled Scripts

All scripts are in the `scripts/` directory:

| Script | Purpose |
|--------|---------|
| `env_detect.py` | Cross-platform venv/dependency detection |
| `preprocess.py` | Markdown cleanup (special chars, empty cells, lists) |
| `convert.py` | Standard Markdown → HWPX conversion |
| `convert_mermaid.py` | Mermaid + Markdown → HWPX conversion |
| `style_tables.py` | Table styling post-processor |
| `apply_bold.py` | Bold formatting post-processor |
| `fix_xml.py` | XML entity escaping post-processor |

## Mermaid Diagram Optimization

When using `convert_mermaid.py`, diagrams are optimized for A4 portrait:

- Viewport 600px = page content width (159mm) at 1:1 scale
- Scale 3x = high resolution without layout change
- Font size 15px ≈ 11pt (matches body text)
- `\n` in node labels → `<br/>` for line breaks
- `graph LR` with 4+ nodes → 2-column grid layout
- `graph LR` with 8+ nodes → 3-column grid layout
