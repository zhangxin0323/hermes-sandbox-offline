#!/usr/bin/env python3
"""验证沙箱环境中所有 Python 包、系统工具及公文字体是否正确安装"""

import os
import subprocess
import sys

PACKAGES = [
    ("markitdown", "from markitdown import MarkItDown"),
    ("Pillow", "from PIL import Image"),
    ("defusedxml", "import defusedxml"),
    ("lxml", "import lxml"),
    ("python-pptx", "from pptx import Presentation"),
    ("pandas", "import pandas"),
    ("numpy", "import numpy"),
    ("scipy", "import scipy"),
    ("matplotlib", "import matplotlib; matplotlib.use('Agg')"),
    ("seaborn", "import seaborn"),
    ("openpyxl", "import openpyxl"),
    ("xlrd", "import xlrd"),
    ("xlsxwriter", "import xlsxwriter"),
    ("python-docx", "import docx"),
    ("docxtpl", "from docxtpl import DocxTemplate"),
    ("python-calamine", "import python_calamine"),
    ("psycopg2-binary", "import psycopg2"),
    ("pymongo", "import pymongo"),
    ("redis", "import redis"),
    ("sqlalchemy", "import sqlalchemy"),
    ("pymysql", "import pymysql"),
    ("requests", "import requests"),
    ("pyyaml", "import yaml"),
    ("jinja2", "import jinja2"),
    ("tabulate", "import tabulate"),
    ("tqdm", "import tqdm"),
    ("pygame", "import pygame"),
    ("pypandoc", "import pypandoc"),
    ("PyMuPDF", "import fitz"),
    ("pdfplumber", "import pdfplumber"),
    ("pytesseract", "import pytesseract"),
    ("plotly", "import plotly"),
    ("kaleido", "import kaleido"),
    ("diagrams", "import diagrams"),
    ("networkx", "import networkx"),
    ("xmind", "import xmind"),
    ("sqlcipher3", "import sqlcipher3"),
    ("oracledb", "import oracledb"),
    ("dmPython", "import dmPython"),  # 如离线安装失败则注释
]

# 20260901v1 新增: 系统工具（视频处理 / 进程管理 / 文档转换 / OCR 等）
SYSTEM_TOOLS = [
    ("ffmpeg (视频剪辑/转码)", "ffmpeg -version"),
    ("ffprobe (视频信息)", "ffprobe -version"),
    ("ps (进程查看, procps)", "ps --version"),
    ("grep (文本检索)", "grep --version"),
    ("kill (进程终止)", "kill -l"),
    ("pandoc (文档转换)", "pandoc --version"),
    ("soffice (LibreOffice)", "soffice --version"),
    ("pdftoppm (PDF转图片)", "pdftoppm -v"),
    ("tesseract (OCR)", "tesseract --version"),
]

# 20260901v1 新增: 公文字体（GB2312 系列，供 PDF 转换后检查公文版式格式）
GONGWEN_FONT_FILES = [
    ("FZXBSJW.TTF", "方正小标宋简体（公文标题）"),
    ("simfang.ttf", "仿宋_GB2312（公文正文）"),
    ("simhei.ttf", "黑体（一级标题）"),
    ("simkai.ttf", "楷体_GB2312（二级标题）"),
    ("simsun.ttc", "宋体（一般正文/表格）"),
]

GONGWEN_FONT_DIR = "/usr/share/fonts/truetype/custom"
GONGWEN_FONT_FAMILIES = [
    "FZXiaoBiaoSong-B05S", "方正小标宋简体",
    "FangSong_GB2312", "仿宋_GB2312",
    "SimHei", "黑体",
    "KaiTi_GB2312", "楷体_GB2312",
    "SimSun", "宋体",
]

failed = []


def check_python_packages():
    """逐包 import 验证。"""
    n_ok = 0
    for name, stmt in PACKAGES:
        try:
            exec(stmt)
            print(f"  OK   {name}")
            n_ok += 1
        except Exception as e:
            print(f"  FAIL {name}: {e}")
            failed.append(("Python 包", name))
    print(f"[Python 包] {n_ok}/{len(PACKAGES)} 通过")


def check_system_tools():
    """逐个执行系统命令验证可用性。"""
    n_ok = 0
    for name, cmd in SYSTEM_TOOLS:
        try:
            subprocess.run(
                cmd, shell=True, check=True,
                stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                timeout=60,
            )
            print(f"  OK   {name}")
            n_ok += 1
        except Exception:
            print(f"  FAIL {name}")
            failed.append(("系统工具", name))
    print(f"[系统工具] {n_ok}/{len(SYSTEM_TOOLS)} 通过")


def check_gongwen_fonts():
    """检查公文字体文件存在且已被 fontconfig 注册。"""
    n_ok = 0
    font_dir_ok = os.path.isdir(GONGWEN_FONT_DIR)
    for fname, desc in GONGWEN_FONT_FILES:
        path = os.path.join(GONGWEN_FONT_DIR, fname)
        if font_dir_ok and os.path.isfile(path) and os.path.getsize(path) > 0:
            print(f"  OK   {fname} ({desc})")
            n_ok += 1
        else:
            print(f"  FAIL {fname} ({desc}) -> 缺少 {path}")
            failed.append(("公文字体", fname))

    try:
        out = subprocess.run(
            "fc-list : file family", shell=True, capture_output=True,
            text=True, timeout=60,
        ).stdout
        found = [fam for fam in GONGWEN_FONT_FAMILIES if fam in out]
        missing = [fam for fam in GONGWEN_FONT_FAMILIES if fam not in out]
        print(f"  OK   fontconfig 已注册公文家族: {'/'.join(found)}")
        if missing:
            print(f"  WARN 未在 fc-list 识别的家族: {'/'.join(missing)}")
    except Exception as e:
        print(f"  WARN fc-list 检查失败: {e}")

    print(f"[公文字体] 文件 {n_ok}/{len(GONGWEN_FONT_FILES)} 通过")


def main():
    failed.clear()
    print("=" * 60)
    print("Hermes Full-Stack Sandbox 验证")
    print("=" * 60)

    print("\n[1/3] Python 包")
    check_python_packages()

    print("\n[2/3] 系统工具 (20260901v1)")
    check_system_tools()

    print("\n[3/3] 公文字体 (20260901v1)")
    check_gongwen_fonts()

    print("\n" + "=" * 60)
    if failed:
        print(f"验证失败: {len(failed)} 项")
        for group, name in failed:
            print(f"  - [{group}] {name}")
        sys.exit(1)
    else:
        print(f"全部验证通过: Python 包 {len(PACKAGES)} / "
              f"系统工具 {len(SYSTEM_TOOLS)} / 公文字体 {len(GONGWEN_FONT_FILES)}")


if __name__ == "__main__":
    main()
