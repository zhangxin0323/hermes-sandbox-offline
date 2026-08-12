---
name: wps-office-formats
category: office
description: >-
  创建、读写、转换 .wps 和 .et 文件。依赖 LibreOffice + xvfb。
---

# WPS Office 私有格式操作指南 (.wps / .et)

## 什么时候用这个 skill

- 你收到了 `.wps` 或 `.et` 文件（WPS Office 私有格式），需要读取内容
- 你需要在代码中从零生成 `.wps` 或 `.et` 文件
- 你需要把 `.wps`/`.et` 转为 PDF、docx、xlsx、CSV 等通用格式
- 你的环境是 Linux 无头服务器（无显示器），内网无法访问 WPS Office

## 原理概述

WPS Office 的 `.wps`（文档）和 `.et`（表格）是 **OLE2 复合文档** 格式。

```
.wps = MS Word 97 二进制格式  → LibreOffice 识别
.et  = MS Excel 97 二进制格式 → LibreOffice 识别
```

**核心流程**（两条路径都适用）：

```
┌─────────────┐    ┌──────────────────┐    ┌──────────────────┐
│  纯 XML 文本  │ → │  .odt / .ods     │ → │  .wps / .et      │
│  (手写内容)   │    │  (ZIP 打包)      │    │  (LibreOffice 转) │
└─────────────┘    └──────────────────┘    └──────────────────┘
                        ↑                          ↑
                  openpyxl/python-docx        LibreOffice 无头转换
                  → xlsx/docx               (soffice --convert-to)
```

**为什么不直接用 Python 写 .wps/.et？**  
这两个格式是 WPS 私有的二进制 OLE2 格式，没有原生 Python 读写库。LibreOffice 是唯一可靠的开源解决方案。

---

## 第一步：环境准备

所有操作都依赖 **虚拟显示**（因为 LibreOffice 在无头模式下需要 X11）：

```bash
# ── 启动虚拟显示（只需要执行一次）──
Xvfb :99 -screen 0 1024x768x24 -ac &

# ── 设置环境变量（每次新终端都要执行）──
export DISPLAY=:99
export SAL_USE_VCLPLUGIN=gen

# ── 验证 LibreOffice 可用 ──
soffice --version    # 应输出 LibreOffice 7.x
```

> ⚠️ 不要用 `xvfb-run`：沙箱环境缺 `xauth` 包。手动启动 Xvfb 是稳定方案。

---

## 第二步：.wps 文档操作

### 2.1 从零创建 .wps（不经任何中转格式）

**工作流说明**：
1. 用 Python 字符串直接构造 OpenDocument XML 内容
2. 把 XML 打包成 `.odt` 文件（实际上就是 ZIP + XML）
3. 调用 LibreOffice 把 `.odt` 转成 `.wps`（MS Word 97 格式）

```python
import zipfile, os, subprocess

OUTPUT_DIR = "/path/to/output"        # ← 改成你的输出目录

# ── 第1步：手写文档内容（OpenDocument XML 格式）──
content_xml = '''<?xml version="1.0" encoding="UTF-8"?>
<office:document-content
  xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
  office:version="1.2">
  <office:body>
    <office:text>
      <text:p>WPS 文档标题</text:p>                    <!-- 每段一个 <text:p> -->
      <text:p>这是正文内容...</text:p>
      <text:p>支持中文、英文、数字 123</text:p>
    </office:text>
  </office:body>
</office:document-content>'''

# ── 第2步：生成 .odt（ZIP 包）──
manifest_xml = '''<?xml version="1.0" encoding="UTF-8"?>
<manifest:manifest
  xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"
  manifest:version="1.2">
  <manifest:file-entry manifest:full-path="/" manifest:version="1.2"
    manifest:media-type="application/vnd.oasis.opendocument.text"/>
  <manifest:file-entry manifest:full-path="content.xml"
    manifest:media-type="text/xml"/>
</manifest:manifest>'''

odt_path = os.path.join(OUTPUT_DIR, "source.odt")
with zipfile.ZipFile(odt_path, 'w', zipfile.ZIP_DEFLATED) as zf:
    # mimetype 必须第一个写入，且不压缩
    zf.writestr("mimetype", b"application/vnd.oasis.opendocument.text",
                compress_type=zipfile.ZIP_STORED)
    zf.writestr("META-INF/manifest.xml", manifest_xml)
    zf.writestr("content.xml", content_xml)

print(f"✅ .odt 创建成功: {odt_path}")

# ── 第3步：LibreOffice 转 .wps ──
env = os.environ.copy()
env["SAL_USE_VCLPLUGIN"] = "gen"       # 防挂起
env["DISPLAY"] = ":99"                  # 连接到虚拟显示

subprocess.run([
    "soffice", "--headless",
    "--convert-to", "wps",              # 输出 .wps (MS Word 97)
    "--outdir", OUTPUT_DIR,             # 输出目录
    odt_path
], env=env, timeout=120, check=True)

# 此时 OUTPUT_DIR/source.wps 已生成
print(f"✅ .wps 创建成功: {os.path.join(OUTPUT_DIR, 'source.wps')}")
```

**执行后文件**：
```
/path/to/output/
  ├── source.odt     # 中间文件（可删除）
  └── source.wps     # 最终 .wps 文件（23KB+）
```

**快捷方式 — 从 .docx 转 .wps**（更简单，适合已有内容）：

```bash
# 命令行动作
DISPLAY=:99 SAL_USE_VCLPLUGIN=gen soffice --headless --convert-to wps input.docx
```

```python
# Python 方式：python-docx 生成 → 转 .wps
from docx import Document
import subprocess, os

doc = Document()
doc.add_heading('标题', level=1)
doc.add_paragraph('正文内容')
doc.save('/tmp/input.docx')

env = os.environ.copy()
env["SAL_USE_VCLPLUGIN"] = "gen"
env["DISPLAY"] = ":99"
subprocess.run(["soffice", "--headless", "--convert-to", "wps",
                "--outdir", "/tmp", "/tmp/input.docx"],
               env=env, timeout=120, check=True)
```

### 2.2 读取 .wps 内容

有三种方式读取 `.wps`，按场景选择：

| 方式 | 命令 | 适用场景 |
|------|------|----------|
| **转纯文本** | → .txt + cat | 快速预览内容 |
| **转 .docx** | → .docx + python-docx | 需要结构化的段落/表格数据 |
| **转 HTML** | → .html | 需要查看完整格式和样式 |

```bash
# 方式 1：转纯文本（最快）
DISPLAY=:99 SAL_USE_VCLPLUGIN=gen soffice --headless --convert-to txt input.wps
cat input.txt

# 方式 2：转 .docx 用 python-docx 读取（结构化）
DISPLAY=:99 SAL_USE_VCLPLUGIN=gen soffice --headless --convert-to docx input.wps
python3 -c "
from docx import Document
doc = Document('input.docx')
for p in doc.paragraphs:
    print(p.text)
"

# 方式 3：转 HTML（保留格式结构）
DISPLAY=:99 SAL_USE_VCLPLUGIN=gen soffice --headless --convert-to html input.wps
# 然后用浏览器打开或解析 HTML
```

### 2.3 .wps 格式转换

```bash
# .wps → PDF（最常用）
DISPLAY=:99 SAL_USE_VCLPLUGIN=gen soffice --headless --convert-to pdf  input.wps

# .wps → .docx（MS Word 2007+）
DISPLAY=:99 SAL_USE_VCLPLUGIN=gen soffice --headless --convert-to docx input.wps

# .wps → .odt（OpenDocument）
DISPLAY=:99 SAL_USE_VCLPLUGIN=gen soffice --headless --convert-to odt  input.wps

# .wps → HTML
DISPLAY=:99 SAL_USE_VCLPLUGIN=gen soffice --headless --convert-to html input.wps

# .wps → 纯文本
DISPLAY=:99 SAL_USE_VCLPLUGIN=gen soffice --headless --convert-to txt  input.wps
```

### 2.4 验证 .wps 文件结构

用 `olefile` 检查 OLE2 复合文档的内部结构：

```python
import olefile
ole = olefile.OleFileIO("input.wps")
print("文件大小:", ole._filesize, "bytes")
print("内部流:")
for s in ole.listdir():
    print("  ", "/".join(s))
# 正常 .wps 应包含：
#   CompObj, Ole, DocumentSummaryInformation,
#   SummaryInformation, 1Table, WordDocument
```

---

## 第三步：.et 表格操作

### 3.1 从零创建 .et（不经任何中转格式）

**工作流说明**：
1. 手写 OpenDocument Spreadsheet XML（定义单元格、数据类型、值）
2. 打包成 `.ods`（ZIP + XML）
3. LibreOffice 转 `.et`（MS Excel 97 格式）

```python
import zipfile, os, subprocess

OUTPUT_DIR = "/path/to/output"        # ← 改成你的输出目录

# ── 第1步：手写表格内容（XML）──
content_xml = '''<?xml version="1.0" encoding="UTF-8"?>
<office:document-content
  xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
  office:version="1.2">
  <office:body>
    <office:spreadsheet>

      <!-- 定义表：Sheet1，4列 -->
      <table:table table:name="Sheet1">
        <table:table-column table:number-columns-repeated="4"/>

        <!-- 表头行 -->
        <table:table-row>
          <table:table-cell office:value-type="string"><text:p>项目名称</text:p></table:table-cell>
          <table:table-cell office:value-type="string"><text:p>完成进度</text:p></table:table-cell>
          <table:table-cell office:value-type="string"><text:p>负责人</text:p></table:table-cell>
          <table:table-cell office:value-type="string"><text:p>备注</text:p></table:table-cell>
        </table:table-row>

        <!-- 数据行 1 -->
        <table:table-row>
          <table:table-cell office:value-type="string"><text:p>项目A</text:p></table:table-cell>
          <table:table-cell office:value-type="float" office:value="85.5"><text:p>85.5%</text:p></table:table-cell>
          <table:table-cell office:value-type="string"><text:p>张三</text:p></table:table-cell>
          <table:table-cell office:value-type="string"><text:p>进行中</text:p></table:table-cell>
        </table:table-row>

        <!-- 数据行 2 -->
        <table:table-row>
          <table:table-cell office:value-type="string"><text:p>项目B</text:p></table:table-cell>
          <table:table-cell office:value-type="float" office:value="100"><text:p>100%</text:p></table:table-cell>
          <table:table-cell office:value-type="string"><text:p>李四</text:p></table:table-cell>
          <table:table-cell office:value-type="string"><text:p>已完成</text:p></table:table-cell>
        </table:table-row>

      </table:table>
    </office:spreadsheet>
  </office:body>
</office:document-content>'''

# ── 第2步：生成 .ods（ZIP 包）──
manifest_xml = '''<?xml version="1.0" encoding="UTF-8"?>
<manifest:manifest
  xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"
  manifest:version="1.2">
  <manifest:file-entry manifest:full-path="/" manifest:version="1.2"
    manifest:media-type="application/vnd.oasis.opendocument.spreadsheet"/>
  <manifest:file-entry manifest:full-path="content.xml"
    manifest:media-type="text/xml"/>
</manifest:manifest>'''

ods_path = os.path.join(OUTPUT_DIR, "source.ods")
with zipfile.ZipFile(ods_path, 'w', zipfile.ZIP_DEFLATED) as zf:
    zf.writestr("mimetype", b"application/vnd.oasis.opendocument.spreadsheet",
                compress_type=zipfile.ZIP_STORED)
    zf.writestr("META-INF/manifest.xml", manifest_xml)
    zf.writestr("content.xml", content_xml)

print(f"✅ .ods 创建成功: {ods_path}")

# ── 第3步：LibreOffice 转 .et ──
env = os.environ.copy()
env["SAL_USE_VCLPLUGIN"] = "gen"
env["DISPLAY"] = ":99"
subprocess.run([
    "soffice", "--headless",
    "--convert-to", "et",               # 输出 .et (MS Excel 97)
    "--outdir", OUTPUT_DIR,
    ods_path
], env=env, timeout=120, check=True)

print(f"✅ .et 创建成功: {os.path.join(OUTPUT_DIR, 'source.et')}")
```

**单元格 XML 数据类型速查**：

| office:value-type | 用途 | 示例 |
|-------------------|------|------|
| `string` | 文本 | `<text:p>你好</text:p>` |
| `float` | 数字 | `office:value="123.45"` |
| `currency` | 货币 | `office:value="100" office:currency="CNY"` |
| `percentage` | 百分比 | `office:value="0.85"` |
| `date` | 日期 | `office:value="2026-08-12"` |

**快捷方式 — 从 .xlsx 转 .et**（适合已有表格数据）：

```bash
# 命令行
DISPLAY=:99 SAL_USE_VCLPLUGIN=gen soffice --headless --convert-to et input.xlsx
```

```python
# Python：openpyxl 写 → 转 .et
import openpyxl, subprocess, os

wb = openpyxl.Workbook()
ws = wb.active
ws['A1'] = '名称'; ws['B1'] = '数值'; ws['C1'] = '状态'
ws['A2'] = '产品A'; ws['B2'] = 100;    ws['C2'] = '通过'
ws['A3'] = '产品B'; ws['B3'] = 200;    ws['C3'] = '待审'
wb.save('/tmp/input.xlsx')

env = os.environ.copy()
env["SAL_USE_VCLPLUGIN"] = "gen"
env["DISPLAY"] = ":99"
subprocess.run(["soffice","--headless","--convert-to","et",
                "--outdir","/tmp","/tmp/input.xlsx"],
               env=env, timeout=120, check=True)
```

### 3.2 读取 .et 内容

三种方式，按需求选：

```bash
# ── 方式 1：转 CSV（最快速，适合预览）──
DISPLAY=:99 SAL_USE_VCLPLUGIN=gen soffice --headless \
  --convert-to 'csv:Text - txt - csv (StarCalc):44,34,76' input.et
cat input.csv
# 输出：
#   项目名称,完成进度,负责人,备注
#   项目A,85.5,张三,进行中
#   项目B,100,李四,已完成

# ── 方式 2：转 xlsx 用 openpyxl 读取（结构化数据）──
DISPLAY=:99 SAL_USE_VCLPLUGIN=gen soffice --headless --convert-to xlsx input.et
python3 -c "
import openpyxl
wb = openpyxl.load_workbook('input.xlsx', data_only=True)
ws = wb.active
for row in ws.iter_rows(values_only=True):
    print(row)
"

# ── 方式 3：pandas + calamine 引擎（一行搞定，推荐）──
python3 -c "
import pandas as pd
df = pd.read_excel('input.et', engine='calamine')
print(df)
"
```

### 3.3 .et 格式转换

```bash
# .et → .xlsx（最常用，兼容 MS Excel）
DISPLAY=:99 SAL_USE_VCLPLUGIN=gen soffice --headless --convert-to xlsx  input.et

# .et → .ods（OpenDocument 标准格式）
DISPLAY=:99 SAL_USE_VCLPLUGIN=gen soffice --headless --convert-to ods   input.et

# .et → HTML
DISPLAY=:99 SAL_USE_VCLPLUGIN=gen soffice --headless --convert-to html  input.et

# .et → CSV（UTF-8，逗号分隔）
DISPLAY=:99 SAL_USE_VCLPLUGIN=gen soffice --headless \
  --convert-to 'csv:Text - txt - csv (StarCalc):44,34,76' input.et
```

---

## 第四步：Python 一键读写 .et（完整函数）

这两个函数封装了完整流程，复制到你的代码中直接使用：

```python
import subprocess, os, tempfile, shutil, openpyxl

def read_et(filepath: str, data_only=True) -> list[dict]:
    """
    读取 .et 文件 → [{列名: 值}, ...]

    参数:
        filepath: .et 文件路径
        data_only: True=取计算后的值，False=保留公式
    返回:
        [{'名称':'项目A','完成进度':85.5,...}, ...]

    依赖: LibreOffice + Xvfb(DISPLAY=:99), openpyxl
    """
    tmpdir = tempfile.mkdtemp()
    try:
        # 1) LibreOffice 转 .xlsx
        env = os.environ.copy()
        env["SAL_USE_VCLPLUGIN"] = "gen"
        env["DISPLAY"] = os.environ.get("DISPLAY", ":99")
        subprocess.run(
            ["soffice","--headless","--convert-to","xlsx",
             "--outdir", tmpdir, filepath],
            env=env, capture_output=True, timeout=120, check=True)

        # 2) openpyxl 读取
        xlsx = os.path.join(tmpdir,
            os.path.splitext(os.path.basename(filepath))[0] + ".xlsx")
        wb = openpyxl.load_workbook(xlsx, data_only=data_only)
        ws = wb.active
        headers = [c.value for c in ws[1]]
        return [dict(zip(headers, row))
                for row in ws.iter_rows(min_row=2, values_only=True)]
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)


def write_et(data: list[dict], output_path: str):
    """
    写入 .et 文件

    参数:
        data: [{列名: 值}, ...]   # 第一行作表头
        output_path: 输出的 .et 文件路径

    用法:
        rows = [{'名称':'A','数值':100},{'名称':'B','数值':200}]
        write_et(rows, 'output.et')
    """
    import openpyxl
    tmpdir = tempfile.mkdtemp()
    try:
        # 1) openpyxl 生成 .xlsx
        headers = list(data[0].keys()) if data else []
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(headers)
        for row in data:
            ws.append([row.get(h, "") for h in headers])
        xlsx = os.path.join(tmpdir, "temp.xlsx")
        wb.save(xlsx)

        # 2) LibreOffice 转 .et
        env = os.environ.copy()
        env["SAL_USE_VCLPLUGIN"] = "gen"
        env["DISPLAY"] = os.environ.get("DISPLAY", ":99")
        subprocess.run(
            ["soffice","--headless","--convert-to","et",
             "--outdir", tmpdir, xlsx],
            env=env, capture_output=True, timeout=120, check=True)

        # 3) 复制结果到目标路径
        shutil.copy2(os.path.join(tmpdir, "temp.et"), output_path)
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)
```

**使用示例**：

```python
# 读取
data = read_et("报表.et")
print(data[0]["项目名称"])   # → "项目A"

# 写入
write_et([{"姓名":"王五","分数":95}], "成绩单.et")

# 读取后转为 CSV
import csv
with open("output.csv", "w", newline="", encoding="utf-8") as f:
    w = csv.DictWriter(f, fieldnames=data[0].keys())
    w.writeheader()
    w.writerows(data)
```

---

## 批量转换脚本

```bash
#!/bin/bash
# batch_convert.sh
# 用法: ./batch_convert.sh <输入目录> <目标格式>
# 示例: ./batch_convert.sh ./reports pdf

SRC_DIR="${1:-.}"
TARGET="${2:-docx}"

DISPLAY=:99 SAL_USE_VCLPLUGIN=gen soffice --headless \
  --convert-to "$TARGET" --outdir "$SRC_DIR" "$SRC_DIR"/*.{wps,et} 2>&1
```

---

## 已知问题与解决方法

| 问题 | 现象 | 原因 | 解决方法 |
|------|------|------|----------|
| **xvfb-run 失败** | `xauth command not found` | 沙箱未安装 xauth 包 | **手动启动**：`Xvfb :99 ... &` + 设置 DISPLAY |
| **javaldx 警告** | `Warning: failed to launch javaldx` | Java 未安装 | **忽略**，不影响文档转换功能 |
| **公式兼容** | WPS 公式变成静态值 | LibreOffice 对 WPS 公式支持有限 | 存两份：`.et` + `.xlsx` |
| **中文显示** | 中文显示为方块 | 缺少中文字体 | `apt install fonts-wqy-zenhei fonts-wqy-microhei` |
| **OLE2 结构** | 部分 WPS 元数据丢失 | 私有元数据非标准 OLE2 | 不影响内容读写，只影响 WPS 特有功能 |

---

## 快速验证（一键检查环境是否正常）

```bash
# 1. 确认 LibreOffice 已安装
soffice --version

# 2. 确认 Xvfb 可启动
Xvfb :99 -screen 0 1024x768x24 -ac &
export DISPLAY=:99
export SAL_USE_VCLPLUGIN=gen

# 3. 创建并查看测试文件
python3 << 'EOF'
import zipfile, subprocess, os

xml = '''<?xml version="1.0" encoding="UTF-8"?>
<office:document-content
  xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
  office:version="1.2">
  <office:body>
    <office:text>
      <text:p>环境验证通过 ✅</text:p>
    </office:text>
  </office:body>
</office:document-content>'''

manifest = '''<?xml version="1.0" encoding="UTF-8"?>
<manifest:manifest
  xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"
  manifest:version="1.2">
  <manifest:file-entry manifest:full-path="/" manifest:version="1.2"
    manifest:media-type="application/vnd.oasis.opendocument.text"/>
  <manifest:file-entry manifest:full-path="content.xml"
    manifest:media-type="text/xml"/>
</manifest:manifest>'''

with zipfile.ZipFile("/tmp/test.odt", "w", zipfile.ZIP_DEFLATED) as zf:
    zf.writestr("mimetype", b"application/vnd.oasis.opendocument.text",
                compress_type=zipfile.ZIP_STORED)
    zf.writestr("META-INF/manifest.xml", manifest)
    zf.writestr("content.xml", xml)

env = os.environ.copy()
env["SAL_USE_VCLPLUGIN"] = "gen"
env["DISPLAY"] = ":99"
subprocess.run(["soffice","--headless","--convert-to","wps","--outdir","/tmp","/tmp/test.odt"],
               env=env, timeout=120, check=True)
print("✅ .wps 创建成功" if os.path.exists("/tmp/test.wps") else "❌ 失败")
EOF
```