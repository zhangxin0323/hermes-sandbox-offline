#!/usr/bin/env python3
"""验证沙箱环境中所有 Python 包是否正确安装"""

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
    ("docxtpl", "from docxtpl import DocxTemplate"),        # ← 新增
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
    ("pygame", "import pygame"),                              # ← 新增
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

failed = []
for name, stmt in PACKAGES:
    try:
        exec(stmt)
        print(f"  OK   {name}")
    except Exception as e:
        print(f"  FAIL {name}: {e}")
        failed.append(name)

if failed:
    print(f"\n验证失败: {len(failed)} 个包")
    for name in failed:
        print(f"  - {name}")
    exit(1)
else:
    print(f"\n全部 {len(PACKAGES)} 个包验证通过")