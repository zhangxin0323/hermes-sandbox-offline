# Hermes Full-Stack Sandbox

> A complete sandbox environment for Hermes Agent, supporting PowerPoint/Office processing, data analytics, domestic Chinese databases, flowcharts/mindmaps, OCR, and CJK font rendering.

---

## Directory Structure

```
hermes-sandbox/
├── Dockerfile                          # Sandbox image build file
├── README.md                           # Chinese documentation
├── README_EN.md                        # This file (English)
├── .gitignore                          # Git ignore rules
├── dependencies.md                     # Sandbox dependency manifest (read by Hermes)
├── download-deps.ps1                   # External dependency download script (requires internet)
├── soffice.py                          # LibreOffice sandbox wrapper (prevents hangs)
├── verify.py                           # Package verification script (runtime self-check)
└── pandoc-3.1.11-linux-amd64.tar.gz  # External binary: Pandoc document converter
    # ↑ This file is excluded from version control; download manually or run download-deps.ps1
```

---

## External Dependencies

Place the following file in this directory before building:

| File | Source | Description |
|------|--------|-------------|
| `pandoc-3.1.11-linux-amd64.tar.gz` | [GitHub: jgm/pandoc/releases](https://github.com/jgm/pandoc/releases) | Document converter (Markdown / Word / PDF) |

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
docker build -t hermes-sandbox-python312:20260812v2 -f Dockerfile .

# Verify image
docker images | grep hermes-sandbox
```

---

## Verify Sandbox

```bash
# Verify all Python packages
docker run --rm hermes-sandbox-python312:20260812v2 python3 /tmp/verify.py

# Verify Chinese font rendering
docker run --rm hermes-sandbox-python312:20260812v2 python3 -c "
import matplotlib.pyplot as plt
plt.rcParams['font.sans-serif'] = ['WenQuanYi Zen Hei']
plt.text(0.5, 0.5, '中文测试', ha='center', fontsize=20)
plt.savefig('/tmp/chinese_test.png')
print('CJK rendering OK')
"

# Verify pandoc
docker run --rm hermes-sandbox-python312:20260812v2 pandoc --version
```

---

## Hermes Configuration

Add the following to your Hermes `config.yaml`:

```yaml
terminal:
  backend: docker
  docker_image: hermes-sandbox-python312:20260812v2
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
| Data processing | `pandas`, `numpy`, `scipy` |
| Visualization | `matplotlib`, `seaborn`, `plotly`, `kaleido` |
| Flowcharts / architecture diagrams | `graphviz`, `diagrams`, `networkx` |
| Mind maps | `xmind`, `xmindparser` |
| Multimedia / audio playback | `pygame` |
| Databases (general) | `psycopg2`, `pymongo`, `redis`, `sqlalchemy`, `pymysql` |
| Databases (domestic Chinese) | `sqlcipher3` (WeChat/DingTalk), `oracledb`, `dmPython` |
| Document conversion | `pandoc`, `pypandoc` |

---

## Version History

| Version | Date | Notes |
|---------|------|-------|
| `20260731v1` | 2026-07-31 | Initial release: basic Office + data processing |
| `20260810v1` | 2026-08-10 | Added OCR, flowcharts, mindmaps, domestic databases, CJK fonts, updated Pandoc |
| `20260812v1` | 2026-08-12 | Fixed matplotlib CJK rendering (full matplotlibrc replacement) |
| `20260812v2` | 2026-08-12 | Added `docxtpl` Word template rendering and `pygame` multimedia engine |

---

## License

External tools follow their respective original licenses (Pandoc / LibreOffice / Tesseract, etc.).
The Dockerfile and scripts in this repository are for internal use with no additional restrictions.
