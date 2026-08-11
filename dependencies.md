# Hermes 沙箱依赖清单

&gt; 镜像: hermes-sandbox-python312:20260810v2  
&gt; 生成日期: 2026-08-10  
&gt; 用途: PowerPoint/Office 处理 + 数据处理 + 国产/Oracle 数据库 + 流程图/思维导图 + OCR + 中文图表渲染

---

## 一、系统工具 (apt)

| 包名 | 用途 |
|------|------|
| libreoffice / libreoffice-impress / libreoffice-calc / libreoffice-writer | PPT/Excel/Word 转换、渲染 |
| poppler-utils | PDF 转图片 (pdftoppm) |
| pandoc | Markdown ↔ Word/PDF/HTML 转换 |
| tesseract-ocr / libtesseract-dev | OCR 文字识别 |
| graphviz / libgraphviz-dev | 流程图渲染引擎 |
| libfontconfig1 | matplotlib 字体渲染 |
| libpq5 | PostgreSQL 客户端库 |
| dbus-x11 / xvfb | LibreOffice headless 虚拟显示兜底 |
| fonts-liberation / fonts-dejavu / fonts-noto-cjk | 西文字体 |
| fonts-wqy-zenhei / fonts-wqy-microhei | 中文字体 (matplotlib 中文渲染) |
| libsqlcipher-dev / sqlcipher | 加密 SQLite 库与命令行工具（支撑 Python sqlcipher3 连接微信/钉钉数据库） |

---

## 二、Python 包 (pip)

### Office / 文档处理
| 包名 | 用途 | 典型调用 |
|------|------|----------|
| markitdown[pptx] | PPT/Word/PDF 文本提取 | `MarkItDown().convert("file.pptx")` |
| python-pptx | 读写 .pptx | `Presentation("file.pptx")` |
| python-docx | 读写 .docx | `Document("file.docx")` |
| openpyxl | 读写 .xlsx | `openpyxl.load_workbook("file.xlsx")` |
| xlrd | 读取旧版 .xls | `pandas.read_excel(engine="xlrd")` |
| xlsxwriter | 高性能 Excel 写入 | `Workbook("out.xlsx")` |
| python-calamine | 快速 Excel 读取 (Rust 后端) | `pandas.read_excel(engine="calamine")` |
| olefile | WPS/老 Office OLE2 文档解析 | `OleFileIO("file.wps")` |
| PyMuPDF (fitz) | PDF 文本/图片提取 | `fitz.open("file.pdf")` |
| pdfplumber | PDF 表格精准提取 | `pdfplumber.open("file.pdf")` |
| pytesseract | Tesseract OCR Python 接口 | `pytesseract.image_to_string(img)` |
| pypandoc | Python 调用 pandoc | `pypandoc.convert_file("file.md", "docx")` |

### 数据处理 / 科学计算
| 包名 | 用途 |
|------|------|
| pandas | 表格数据处理 |
| numpy | 数值计算 |
| scipy | 科学计算、统计检验 |
| chardet | 自动检测文件编码 (GBK/UTF-8) |
| python-magic | 文件类型检测 |
| xlsx2csv | 大 Excel 快速转 CSV |
| unstructured | 统一文档文本提取 |

### 可视化 / 中文图表
| 包名 | 用途 | 典型调用 |
|------|------|----------|
| matplotlib | 基础绘图（已配置中文字体） | `plt.plot()` |
| seaborn | 统计可视化 | `sns.heatmap()` |
| plotly | 交互式图表 | `px.bar()` |
| kaleido | plotly 静态导出引擎 | `fig.write_image("out.png")` |

&gt; 中文字体配置: `/usr/local/lib/python3.12/site-packages/matplotlib/mpl-data/matplotlibrc`  
&gt; 字体栈: WenQuanYi Zen Hei → Noto Sans CJK SC → SimHei

### 流程图 / 思维导图 / 架构图
| 包名 | 用途 | 典型调用 |
|------|------|----------|
| graphviz | 流程图 (Python 接口) | `graphviz.Digraph()` |
| pygraphviz | graphviz 的 NetworkX 兼容层 | `nx.draw_graphviz()` |
| diagrams | 云架构/网络拓扑图生成 | `Diagram("web", filename="out")` |
| networkx | 图数据结构、复杂网络分析 | `nx.DiGraph()` |
| xmind | 生成 .xmind 思维导图文件 | `xmind.load()` / `save()` |
| xmindparser | 解析 .xmind 文件 | `xmindparser.xmind_to_dict()` |

### 数据库
| 包名 | 用途 | 典型调用 |
|------|------|----------|
| psycopg2-binary | PostgreSQL | `psycopg2.connect(...)` |
| pymongo | MongoDB | `pymongo.MongoClient(...)` |
| redis | Redis | `redis.Redis(...)` |
| sqlalchemy | ORM / 统一数据库接口 | `create_engine(...)` |
| pymysql | MySQL | `pymysql.connect(...)` |
| sqlcipher3 | 加密 SQLite (微信/钉钉) | `sqlcipher3.connect(...)` |
| oracledb | Oracle（纯 Python，无需 Instant Client） | `oracledb.connect(...)` |
| dmPython | 达梦数据库 | `dmPython.connect(user, pwd, server)` |

&gt; 注: dmPython 如清华源缺失，需从达梦安装盘 `$DM_HOME/drivers/python/dmPython/` 手动安装

### 常用工具
| 包名 | 用途 |
|------|------|
| requests | HTTP 请求 |
| pyyaml | YAML 解析 |
| jinja2 | 模板引擎 |
| tabulate | 表格格式化输出 |
| tqdm | 进度条 |
| defusedxml | 安全 XML 解析 |
| lxml | XML/HTML 处理 |

---

## 三、Node.js 包 (npm global)

| 包名 | 版本 | 用途 |
|------|------|------|
| pptxgenjs | 3.12.0 | 从 JS 代码生成 .pptx |

---

## 四、自定义脚本

| 路径 | 用途 |
|------|------|
| `/opt/hermes/scripts/office/soffice.py` | LibreOffice 沙箱 wrapper (防挂起/崩溃，自动 xvfb 兜底) |
| `/opt/hermes/dependencies.md` | 本文件，沙箱能力清单 |

---

## 五、快速验证命令

```bash
# 系统工具
soffice --version              # LibreOffice
pdftoppm -v                    # Poppler
pandoc --version               # Pandoc
tesseract --version            # OCR

# Python 包
python3 -c "import pandas, numpy, matplotlib, fitz, docx, pptx, graphviz, diagrams, oracledb; print('OK')"

# 中文图表测试
python3 -c "
import matplotlib.pyplot as plt
plt.rcParams['font.sans-serif'] = ['WenQuanYi Zen Hei']
plt.text(0.5, 0.5, '中文测试', ha='center', va='center', fontsize=20)
plt.savefig('/tmp/chinese_test.png')
print('中文渲染 OK')
"

# Node.js
node -e "console.log(require('pptxgenjs').VERSION)"