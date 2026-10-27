# Hermes Full-Stack Sandbox

> A complete sandbox environment for Hermes Agent, supporting PowerPoint/Office processing, data analytics, domestic Chinese databases, flowcharts/mindmaps, OCR, official-document (公文) font rendering, video processing, process management, headless browser web debugging, and Python web application development.

---

## Directory Structure

```
hermes-sandbox/
├── Dockerfile                          # Sandbox image build file
├── README.md                           # Chinese documentation
├── README_EN.md                        # This file (English)
├── .gitignore                          # Git ignore rules
├── dependencies.md                     # Sandbox dependency manifest (read by Hermes)
├── download-deps.ps1                   # External dependency download script (internet, pandoc only)
├── soffice.py                          # LibreOffice sandbox wrapper (prevents hangs)
├── verify.py                           # Package/system-tool/font/browser verification script (runtime self-check)
├── fonts/                              # Official-document fonts dir (GB2312 series, placed manually)
│   ├── FZXBSJW.TTF                     # FangZheng XiaoBiaoSong (doc title)
│   ├── simfang.ttf                     # FangSong_GB2312 (doc body)
│   ├── simhei.ttf                      # SimHei / 黑体 (level-1 heading)
│   ├── simkai.ttf                      # KaiTi_GB2312 (level-2 heading)
│   └── simsun.ttc                      # SimSun / 宋体
└── pandoc-3.1.11-linux-amd64.tar.gz  # External binary: Pandoc document converter
    # ↑ This file is excluded from version control; download manually or run download-deps.ps1
```

---

## External Dependencies

Place the following files in this directory before building:

| File | Source | Description |
|------|--------|-------------|
| `pandoc-3.1.11-linux-amd64.tar.gz` | [GitHub: jgm/pandoc/releases](https://github.com/jgm/pandoc/releases) | Document converter (Markdown / Word / PDF) |
| `fonts/` (FangSong_GB2312, KaiTi_GB2312, FangZheng XiaoBiaoSong, SimHei, SimSun) | Windows built-in fonts or your own font assets | Official-document fonts for checking document typesetting after PDF conversion |

> `download-deps.ps1` only downloads pandoc. The `fonts/` folder must be populated manually before building; the Dockerfile auto-imports it via `COPY ./fonts/*` into `/usr/share/fonts/truetype/custom` and rebuilds the font cache.

### Download Methods

**Method 1: PowerShell Script (recommended, requires internet)**

```powershell
.\download-deps.ps1
```

**Method 2: Manual Download**

Visit https://github.com/jgm/pandoc/releases and download `pandoc-3.1.11-linux-amd64.tar.gz` into this directory.

---

## Build Image

```bash
# Build
docker build -t hermes-sandbox-python312:20261027v1 -f Dockerfile .

# Verify image
docker images | grep hermes-sandbox
```

---

## Verify Sandbox

```bash
# Verify all Python packages / system tools / official-document fonts / Playwright browser
docker run --rm hermes-sandbox-python312:20261027v1 python3 /tmp/verify.py

# Verify ffmpeg video processing
docker run --rm hermes-sandbox-python312:20261027v1 ffmpeg -version

# Verify ps / grep / kill process management
docker run --rm hermes-sandbox-python312:20261027v1 ps -ef

# Verify official-document fonts
docker run --rm hermes-sandbox-python312:20261027v1 fc-list : family | grep -E "FangSong|KaiTi|SimHei|SimSun"

# Verify Chinese font rendering
docker run --rm hermes-sandbox-python312:20261027v1 python3 -c "
import matplotlib.pyplot as plt
plt.rcParams['font.sans-serif'] = ['WenQuanYi Zen Hei']
plt.text(0.5, 0.5, '中文测试', ha='center', fontsize=20)
plt.savefig('/tmp/chinese_test.png')
print('CJK rendering OK')
"

# Verify pandoc
docker run --rm hermes-sandbox-python312:20261027v1 pandoc --version

# Verify Playwright headless Chromium (20261027v1)
docker run --rm hermes-sandbox-python312:20261027v1 python3 -c "
from playwright.sync_api import sync_playwright
with sync_playwright() as p:
    browser = p.chromium.launch(headless=True)
    page = browser.new_page()
    page.goto('data:text/html,<h1>Hello</h1>')
    page.screenshot(path='/tmp/test.png')
    browser.close()
print('Playwright OK')
"

# Verify Flask / FastAPI (20261027v1)
docker run --rm hermes-sandbox-python312:20261027v1 python3 -c "from flask import Flask; print('Flask OK')"
docker run --rm hermes-sandbox-python312:20261027v1 python3 -c "from fastapi import FastAPI; print('FastAPI OK')"
```

---

## Hermes Configuration

Add the following to your Hermes `config.yaml`:

```yaml
terminal:
  backend: docker
  docker_image: hermes-sandbox-python312:20261027v1
  network: "none"
  container_cpu: 4
  container_memory: 5120
  container_disk: 51200
  container_persistent: true
```

---

## Tell Hermes What Tools Are Available

Start a new Hermes conversation and paste:

```
Please remember that `/opt/hermes/dependencies.md` in the sandbox records all available software and Python packages.
```

---

## Capability Quick Reference

| Capability | Tool/Package |
|------------|--------------|
| PPT text extraction | `markitdown[pptx]` |
| PPT generation (Python) | `python-pptx` |
| PPT generation (JS) | `pptxgenjs` |
| Word read/write | `python-docx` |
| Word template rendering (mail merge / contracts) | `docxtpl` |
| Excel read/write | `openpyxl`, `xlsxwriter`, `python-calamine` |
| PDF conversion | `soffice`, `pdftoppm` |
| PDF text/table extraction | `PyMuPDF`, `pdfplumber` |
| OCR | `tesseract-ocr`, `pytesseract` |
| Official-document (公文) typesetting check | FangZheng XiaoBiaoSong / FangSong_GB2312 / SimHei / KaiTi_GB2312 / SimSun fonts |
| Video processing | `ffmpeg`, `ffprobe` (trim/transcode then recognize) |
| Process management | `ps`, `grep`, `kill` |
| Data processing | `pandas`, `numpy`, `scipy` |
| Visualization | `matplotlib`, `seaborn`, `plotly`, `kaleido` |
| Flowcharts / architecture diagrams | `graphviz`, `diagrams`, `networkx` |
| Mind maps | `xmind`, `xmindparser` |
| Multimedia / audio playback | `pygame` |
| Databases (general) | `psycopg2`, `pymongo`, `redis`, `sqlalchemy`, `pymysql` |
| Databases (domestic Chinese) | `sqlcipher3` (WeChat/DingTalk), `oracledb`, `dmPython` |
| Document conversion | `pandoc`, `pypandoc` |
| **Headless browser (20261027v1)** | `playwright` + Chromium (web debugging, screenshots, layout inspection) |
| **Python Web frameworks (20261027v1)** | `flask`, `fastapi`, `uvicorn` |
| **Geographic calculation (20261027v1)** | `geographiclib` |
| **Encoding detection (20261027v1)** | `charset_normalizer` |

---

## Version History

| Version | Date | Notes |
|---------|------|-------|
| `20260731v1` | 2026-07-31 | Initial release: basic Office + data processing |
| `20260810v1` | 2026-08-10 | Added OCR, flowcharts, mindmaps, domestic databases, CJK fonts, updated Pandoc |
| `20260812v2` | 2026-08-12 | Added `docxtpl` Word template rendering and `pygame` multimedia engine |
| `20260901v1` | 2026-09-01 | Added official-document fonts (公文) for checking typesetting after PDF conversion, `ffmpeg`/`ffprobe` video processing, and `ps`/`grep`/`kill` process tools |
| `20261027v1` | 2026-10-27 | Added Playwright headless Chromium browser (web debugging/screenshots), `flask`/`fastapi`/`uvicorn` web frameworks, `geographiclib` geographic calculation, `charset_normalizer` encoding detection |

---

## License

External tools follow their respective original licenses (Pandoc / LibreOffice / Tesseract / Playwright, etc.).
The Dockerfile and scripts in this repository are for internal use with no additional restrictions.
