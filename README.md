# hwpx-convert-skill

Markdown를 한컴오피스 HWPX 문서로 변환하는 Claude Code 스킬입니다.

## 기능

- Markdown → HWPX 자동 변환 (pypandoc-hwpx 기반)
- 크래시 방지 전처리 (특수문자, 빈 셀, 따옴표 등)
- 테이블 스타일링 (헤더 배경색, 비례 컬럼 폭, 130% 줄간격)
- 볼드/위첨자 서식 적용
- Mermaid 다이어그램 → PNG 변환 후 삽입
- macOS / Linux / Windows 지원

## 설치

### 1. Pandoc

```bash
# macOS
brew install pandoc

# Ubuntu/Debian
sudo apt install pandoc
```

### 2. Python 가상환경

```bash
python3 -m venv hwpx_env
source hwpx_env/bin/activate
pip install pypandoc-hwpx
```

### 3. Node.js (mermaid 다이어그램용, 선택)

```bash
brew install node    # macOS
```

### 설치 확인

```bash
python3 scripts/env_detect.py
```

## 사용법

Claude Code에서 자동으로 트리거됩니다:

- "한글로 변환해줘"
- "HWPX로 변환"
- "한컴 파일로 만들어줘"

### 수동 실행

```bash
# 전처리
python3 scripts/preprocess.py input.md /tmp/cleaned.md

# 변환 (mermaid 없는 경우)
python3 scripts/convert.py /tmp/cleaned.md "출력파일명"

# 변환 (mermaid 있는 경우)
python3 scripts/convert_mermaid.py /tmp/cleaned.md "출력파일명"
```

결과물은 `output/` 폴더에 저장됩니다.

## 크래시 방지 처리

| 문제 | 증상 | 자동 처리 |
|------|------|----------|
| 빈 테이블 셀 | 한컴 크래시 | **-** 로 채움 |
| 특수문자 ★☆×→·■ | 깨짐/크래시 | ASCII로 치환 |
| 따옴표 "text" 'text' | 텍스트 소실 | 유니코드 curly quotes로 변환 |
| Mermaid 코드블록 | 크래시 | PNG 이미지로 변환 |
| URL의 & 문자 | 파일 손상 | XML 엔티티 이스케이프 |

## 스크립트 구성

| 스크립트 | 역할 |
|---------|------|
| `preprocess.py` | 마크다운 전처리 (특수문자, 빈 셀, 따옴표) |
| `convert.py` | 기본 변환 (Markdown → HWPX) |
| `convert_mermaid.py` | Mermaid 포함 변환 |
| `style_tables.py` | 테이블 스타일 후처리 |
| `apply_bold.py` | 볼드 서식 후처리 |
| `apply_superscript.py` | 위첨자 서식 후처리 |
| `fix_xml.py` | XML 엔티티 수정 |
| `env_detect.py` | 환경 감지 |

## 라이선스

MIT
