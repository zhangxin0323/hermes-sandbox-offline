# Hermes 沙箱依赖清单

> 镜像: hermes-sandbox-python312:20260901v1  
> 生成日期: 2026-09-01  
> 用途: PowerPoint/Office 处理 + 数据处理 + 国产/Oracle 数据库 + 流程图/思维导图 + OCR + 公文字体渲染 + 视频处理 + 进程管理

---

## 一、系统工具 (apt)

| 包名 | 用途 |
|------|------|
| libreoffice / libreoffice-impress / libreoffice-calc / libreoffice-writer | PPT/Excel/Word 转换、渲染 |
| poppler-utils | PDF 转图片 (pdftoppm) |
| pandoc | Markdown ↔ Word/PDF/HTML 转换 |
| **ffmpeg / ffprobe** | **视频裁剪、转码、提帧识别 (20260901v1 新增)** |
| **procps / psmisc** | **进程管理: `ps` / `kill` (20260901v1 新增)** |
| **grep** | **文本检索筛选 (20260901v1 新增)** |
| tesseract-ocr / libtesseract-dev | OCR 文字识别 |
| graphviz / libgraphviz-dev | 流程图渲染引擎 |
| libfontconfig1 / fontconfig | 字体缓存与渲染 (支撑公文字体与 matplotlib) |
| libpq5 | PostgreSQL 客户端库 |
| dbus-x11 / xvfb | LibreOffice headless 虚拟显示兜底 |
| fonts-liberation / fonts-dejavu / fonts-noto-cjk | 西文字体 |
| fonts-wqy-zenhei / fonts-wqy-microhei | 中文字体 (matplotlib 中文渲染) |
| **公文字体** (fonts/ 目录, 安装到 `/usr/share/fonts/truetype/custom`) | **方正小标宋简体 / 仿宋_GB2312 / 黑体 / 楷体_GB2312 / 宋体 —— 判断公文版式格式 (20260901v1 新增)** |
| libsqlcipher-dev / sqlcipher | 加密 SQLite 库与命令行工具（支撑 Python sqlcipher3 连接微信/钉钉数据库） |

> 公文字体对应关系: 标题→方正小标宋简体(FZXiaoBiaoSong-B05S), 正文→仿宋_GB2312, 一级标题→黑体(SimHei), 二级标题→楷体_GB2312, 表格/一般文字→宋体(SimSun)。转换 PDF/Word 后可用 `fc-match`、`fc-list` 核对字体是否生效。

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
| docxtpl | Word 模板渲染（Jinja2 邮件合并/合同生成） | `DocxTemplate("tpl.docx").render({"name": "x"})` |

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

> 中文字体配置: `/usr/local/lib/python3.12/site-packages/matplotlib/mpl-data/matplotlibrc`  
> 字体栈: WenQuanYi Zen Hei → Noto Sans CJK SC → SimHei → FangSong_GB2312 → KaiTi_GB2312

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

> 注: dmPython 如清华源缺失，需从达梦安装盘 `$DM_HOME/drivers/python/dmPython/` 手动安装

### 常用工具
| 包名 | 用途 | 典型调用 |
|------|------|----------|
| requests | HTTP 请求 | `requests.get(url)` |
| pyyaml | YAML 解析 | `yaml.safe_load(f)` |
| jinja2 | 模板引擎 | `Template("Hello {{name}}").render(name="x")` |
| tabulate | 表格格式化输出 | `tabulate(data, headers, tablefmt="grid")` |
| tqdm | 进度条 | `tqdm(iterable)` |
| pygame | 游戏/多媒体引擎（音频播放、简单 GUI、交互可视化） | `pygame.mixer.Sound("file.wav").play()` |

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
| `/tmp/verify.py` | 构建时自检脚本（Python 包 / 系统工具 / 公文字体） |

---

## 五、快速验证命令

```bash
# 系统工具
soffice --version              # LibreOffice
pdftoppm -v                    # Poppler
pandoc --version               # Pandoc
tesseract --version            # OCR
ffmpeg -version                # 视频处理 (20260901v1)
ffprobe -version               # 视频信息 (20260901v1)
ps -ef                         # 进程查看
grep -r "关键字" /path         # 文本检索
kill -l                       # 进程终止

# 公文字体核对 (20260901v1)
fc-list : family | grep -E "仿宋|楷体|FangSong|KaiTi"   # 查看 GB2312 公文家族
fc-match "仿宋_GB2312"                                  # 查看某字体实际匹配
fc-match "方正小标宋简体"

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
```

---

## 六、已知限制

| 场景 | 说明 | workaround |
|------|------|-------------|
| WPS 私有格式 (.wps/.et/.dps) | 无原生 Python 读写库 | 通过 `soffice --convert-to` 转为 docx/xlsx/pptx 后处理 |
| Visio (.vsdx) | 无 Python 编辑库 | 用 `diagrams` / `graphviz` 生成 PNG/SVG/XMind 替代 |
| dmPython (达梦) | 清华 PyPI 源可能缺失，无预编译 wheel | 从达梦安装盘 `$DM_HOME/drivers/python/dmPython/` 手动安装 |
| 微信/钉钉加密数据库 | SQLCipher 加密，需正确 PRAGMA key | 使用 `sqlcipher3` 连接，密码错误会返回空数据或报错 |
| 中文字体渲染 | 文泉驿字体缺少 `weight=400`，matplotlib 会 fallback 到 500 | 不影响显示，可忽略警告；或额外安装 `fonts-wqy-microhei` |
| 公文版式检查 | 仿宋/楷体等多字体依赖字体文件正确导入 | 构建时 `fc-cache` 已重建；若 PDF 中字体未生效，用 `fc-match` 核对家族名 |
| 视频处理 | 模型需先 `ffmpeg` 裁剪/抽帧再识别 | `ffmpeg -ss <起> -t <时长> -i in.mp4 out.mp4`；`ffprobe` 查看时长/分辨率 |
| 网络访问 | 沙箱默认 `network: "none"` | 如需联网，在 Hermes `config.yaml` 中改为 `network: "bridge"` |
| LibreOffice headless | 容器内无真实 X11，依赖 `xvfb-run` 虚拟显示兜底 | 已通过 `soffice.py` wrapper 自动处理，无需手动干预 |
| 沙箱生命周期 | `auto_remove: true`，容器退出后文件丢失 | 交付物必须写到 `/mnt/output` 或 `/workspace` 等持久化挂载目录 |
| 达梦/人大金仓专用驱动 | `dmPython` / `kingbase` 可能版本滞后 | 优先使用标准接口：`sqlcipher3` 替代 `pysqlcipher3`，`psycopg2` 兼容 Kingbase |
