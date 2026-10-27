# Hermes Full-Stack Sandbox

> 为 Hermes Agent 提供 PowerPoint/Office 处理 + 数据处理 + 国产数据库 + 流程图/思维导图 + OCR + 公文字体渲染 + 视频处理 + 进程管理 + 无头浏览器网页调试 + Python Web 应用调试的完整沙箱环境。

---

## 目录结构

```
hermes-sandbox/
├── Dockerfile                          # 沙箱镜像构建文件
├── README.md                           # 本文件
├── README_EN.md                        # 英文文档
├── .gitignore                          # Git 忽略规则
├── dependencies.md                     # 沙箱依赖清单（供 Hermes 读取）
├── download-deps.ps1                   # 外部依赖下载脚本（需联网，仅 pandoc）
├── soffice.py                          # LibreOffice 沙箱 wrapper（防挂起）
├── verify.py                           # 包/系统工具/字体/浏览器验证脚本（构建时自检）
├── fonts/                              # 公文字体目录（GB2312 系列，需手动放置）
│   ├── FZXBSJW.TTF                     # 方正小标宋简体（公文标题）
│   ├── simfang.ttf                     # 仿宋_GB2312（公文正文）
│   ├── simhei.ttf                      # 黑体（一级标题）
│   ├── simkai.ttf                      # 楷体_GB2312（二级标题）
│   └── simsun.ttc                      # 宋体
└── pandoc-3.1.11-linux-amd64.tar.gz  # 外部二进制：Pandoc 文档转换工具
    # ↑ 此文件不入版本库，需手动下载或运行 download-deps.ps1
```

---

## 外部依赖

构建前需将以下文件放置到本目录：

| 文件 | 来源 | 说明 |
|------|------|------|
| `pandoc-3.1.11-linux-amd64.tar.gz` | [GitHub: jgm/pandoc/releases](https://github.com/jgm/pandoc/releases) | 文档转换工具（Markdown / Word / PDF） |
| `fonts/`（仿宋_GB2312、楷体_GB2312、方正小标宋简体、黑体、宋体） | 由 Windows 自带字体或自有字体资源获取 | 公文字体，供转换 PDF 后检查公文版式格式 |

> `download-deps.ps1` 目前只下载 pandoc；`fonts/` 目录需要在构建前手动放置字体文件（Dockerfile 会通过 `COPY ./fonts/*` 自动导入 `/usr/share/fonts/truetype/custom` 并重建字体缓存）。

### 下载方式

**方式一：PowerShell 脚本（推荐，需联网）**

```powershell
.\download-deps.ps1
```

**方式二：手动下载**

访问 https://github.com/jgm/pandoc/releases ，下载 `pandoc-3.1.11-linux-amd64.tar.gz` 放到本目录。

---

## 构建镜像

```bash
# 构建
docker build -t hermes-sandbox-python312:20261027v1 -f Dockerfile .

# 查看镜像
docker images | grep hermes-sandbox
```

---

## 验证沙箱

```bash
# 验证所有 Python 包 / 系统工具 / 公文字体 / Playwright 浏览器
docker run --rm hermes-sandbox-python312:20261027v1 python3 /tmp/verify.py

# 验证 ffmpeg 视频处理
docker run --rm hermes-sandbox-python312:20261027v1 ffmpeg -version

# 验证 ps / grep / kill 进程管理
docker run --rm hermes-sandbox-python312:20261027v1 ps -ef

# 验证公文字体（应看到 仿宋_GB2312 / 楷体_GB2312 / 方正小标宋简体）
docker run --rm hermes-sandbox-python312:20261027v1 fc-list : family | grep -E "仿宋|楷体|黑体|宋体|FangSong|KaiTi"

# 验证中文字体渲染
docker run --rm hermes-sandbox-python312:20261027v1 python3 -c "
import matplotlib.pyplot as plt
plt.rcParams['font.sans-serif'] = ['WenQuanYi Zen Hei']
plt.text(0.5, 0.5, '中文测试', ha='center', fontsize=20)
plt.savefig('/tmp/chinese_test.png')
print('中文渲染 OK')
"

# 验证 pandoc
docker run --rm hermes-sandbox-python312:20261027v1 pandoc --version

# 验证 Playwright 无头 Chromium（20261027v1 新增）
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

# 验证 Flask / FastAPI（20261027v1 新增）
docker run --rm hermes-sandbox-python312:20261027v1 python3 -c "from flask import Flask; print('Flask OK')"
docker run --rm hermes-sandbox-python312:20261027v1 python3 -c "from fastapi import FastAPI; print('FastAPI OK')"
```

---

## Hermes 配置

在 Hermes 的 `config.yaml` 中配置沙箱镜像：

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

## 告诉 Hermes 沙箱有哪些工具

新建 Hermes 对话，输入：

```
请你记住沙箱的 `/opt/hermes/dependencies.md` 文件记录了沙箱中所有可用的软件和 Python 包。
```

---

## 沙箱能力速查

| 能力 | 工具/包 |
|------|---------|
| PPT 文本提取 | `markitdown[pptx]` |
| PPT 生成（Python） | `python-pptx` |
| PPT 生成（JS） | `pptxgenjs` |
| Word 读写 | `python-docx` |
| Word 模板渲染（邮件合并/合同） | `docxtpl` |
| Excel 读写 | `openpyxl`, `xlsxwriter`, `python-calamine` |
| PDF 转换 | `soffice`, `pdftoppm` |
| PDF 文本/表格提取 | `PyMuPDF`, `pdfplumber` |
| OCR | `tesseract-ocr`, `pytesseract` |
| 公文版式检查 | 公文字体（方正小标宋简体 / 仿宋_GB2312 / 黑体 / 楷体_GB2312 / 宋体） |
| 视频处理 | `ffmpeg`, `ffprobe`（视频裁剪/转码后识别） |
| 系统进程管理 | `ps`, `grep`, `kill` |
| 数据处理 | `pandas`, `numpy`, `scipy` |
| 可视化 | `matplotlib`, `seaborn`, `plotly`, `kaleido` |
| 流程图/架构图 | `graphviz`, `diagrams`, `networkx` |
| 思维导图 | `xmind`, `xmindparser` |
| 数据库（通用） | `psycopg2`, `pymongo`, `redis`, `sqlalchemy`, `pymysql` |
| 数据库（国产） | `sqlcipher3`, `oracledb`, `dmPython` |
| 文档转换 | `pandoc`, `pypandoc` |
| 游戏/多媒体引擎 | `pygame` |
| **无头浏览器（20261027v1）** | `playwright` + Chromium（网页调试、截图、布局检查） |
| **Python Web 框架（20261027v1）** | `flask`, `fastapi`, `uvicorn` |
| **地理计算（20261027v1）** | `geographiclib` |
| **编码检测（20261027v1）** | `charset_normalizer` |

---

## 版本历史

| 版本 | 日期 | 说明 |
|------|------|------|
| `20260731v1` | 2026-07-31 | 初始版本，基础 Office + 数据处理 |
| `20260810v1` | 2026-08-10 | 新增 OCR、流程图、思维导图、国产数据库、中文字体、新版 Pandoc |
| `20260812v2` | 2026-08-12 | 新增 `docxtpl` Word 模板渲染、`pygame` 多媒体引擎 |
| `20260901v1` | 2026-09-01 | 新增公文字体（PDF 转换后检查公文格式）、`ffmpeg`/`ffprobe` 视频处理、`ps`/`grep`/`kill` 进程工具 |
| `20261027v1` | 2026-10-27 | 新增 Playwright 无头 Chromium 浏览器（网页调试/截图）、`flask`/`fastapi`/`uvicorn` Web 框架、`geographiclib` 地理计算、`charset_normalizer` 编码检测 |

---

## 许可证

外部工具遵循各自原始许可证（Pandoc / LibreOffice / Tesseract / Playwright 等）。
本仓库的 Dockerfile 及脚本为内部使用，无额外限制。
