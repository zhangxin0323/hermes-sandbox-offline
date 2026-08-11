# Hermes Full-Stack Sandbox

> 为 Hermes Agent 提供 PowerPoint/Office 处理 + 数据处理 + 国产数据库 + 流程图/思维导图 + OCR + 中文图表渲染的完整沙箱环境。

---

## 目录结构

```
hermes-sandbox/
├── Dockerfile                          # 沙箱镜像构建文件
├── README.md                           # 本文件
├── .gitignore                          # Git 忽略规则
├── dependencies.md                     # 沙箱依赖清单（供 Hermes 读取）
├── download-deps.ps1                   # 外部依赖下载脚本（需联网）
├── soffice.py                          # LibreOffice 沙箱 wrapper（防挂起）
├── verify.py                           # 包验证脚本（构建时自检）
├── matplotlibrc                        # matplotlib 中文字体配置
└── pandoc-3.1.11-linux-amd64.tar.gz  # 外部二进制：Pandoc 文档转换工具
    # ↑ 此文件不入版本库，需手动下载或运行 download-deps.ps1
```

---

## 外部依赖

构建前需将以下文件放置到本目录：

| 文件 | 来源 | 说明 |
|------|------|------|
| `pandoc-3.1.11-linux-amd64.tar.gz` | [GitHub: jgm/pandoc/releases](https://github.com/jgm/pandoc/releases) | 文档转换工具（Markdown / Word / PDF） |

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
# docker build -t hermes-sandbox-python312:20260810v1 -f Dockerfile .

# 查看镜像
docker images | findstr hermes-sandbox
```

---

## 验证沙箱

```bash
# 验证所有 Python 包
docker run --rm hermes-sandbox-python312:20260810v1 python3 /tmp/verify.py

# 验证中文字体
docker run --rm hermes-sandbox-python312:20260810v1 python3 -c "
import matplotlib.pyplot as plt
plt.rcParams['font.sans-serif'] = ['WenQuanYi Zen Hei']
plt.text(0.5, 0.5, '中文测试', ha='center', fontsize=20)
plt.savefig('/tmp/chinese_test.png')
print('中文渲染 OK')
"

# 验证 pandoc
docker run --rm hermes-sandbox-python312:20260810v1 pandoc --version
```

---

## Hermes 配置

在 Hermes 的 `config.yaml` 中配置沙箱镜像：

```yaml
terminal:
  backend: docker
  docker_image: hermes-sandbox-python312:20260810v1
  network: "none"
  container_cpu: 4
  container_memory: 5120
  container_disk: 51200
  container_persistent: true
```

---

## 沙箱能力速查

| 能力 | 工具/包 |
|------|---------|
| PPT 文本提取 | `markitdown[pptx]` |
| PPT 生成（Python） | `python-pptx` |
| PPT 生成（JS） | `pptxgenjs` |
| Word 读写 | `python-docx` |
| Excel 读写 | `openpyxl`, `xlsxwriter`, `python-calamine` |
| PDF 转换 | `soffice`, `pdftoppm` |
| PDF 文本/表格提取 | `PyMuPDF`, `pdfplumber` |
| OCR | `tesseract-ocr`, `pytesseract` |
| 数据处理 | `pandas`, `numpy`, `scipy` |
| 可视化 | `matplotlib`, `seaborn`, `plotly`, `kaleido` |
| 流程图/架构图 | `graphviz`, `diagrams`, `networkx` |
| 思维导图 | `xmind`, `xmindparser` |
| 数据库（通用） | `psycopg2`, `pymongo`, `redis`, `sqlalchemy`, `pymysql` |
| 数据库（国产） | `sqlcipher3`（微信/钉钉）, `oracledb`, `dmPython` |
| 文档转换 | `pandoc`, `pypandoc` |

---

## 版本历史

| 版本 | 日期 | 说明 |
|------|------|------|
| `20260731v1` | 2026-07-31 | 初始版本，基础 Office + 数据处理 |
| `20260810v1` | 2026-08-10 | 新增 OCR、流程图、思维导图、国产数据库、中文字体、新版 Pandoc |

---

## 许可证

外部工具遵循各自原始许可证（Pandoc / LibreOffice / Tesseract 等）。
本仓库的 Dockerfile 及脚本为内部使用，无额外限制。
