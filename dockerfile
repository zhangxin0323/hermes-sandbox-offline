# =============================================================================
# Hermes Full-Stack Sandbox for Windows / WSL2 / Docker Desktop
# =============================================================================
# 基础: Python 3.12 (slim) + Node.js 22 + LibreOffice + Poppler + Pandoc
# 功能: PPT/Word/Excel 处理 + 数据处理 + 数据库(含国产) + 流程图/思维导图 + OCR + 视频处理
#
# 构建:
#   docker build -t hermes-sandbox-python312:20260901v1 -f Dockerfile .
#
# 外部依赖文件（需放在 Dockerfile 同级目录）:
#   - soffice.py          LibreOffice 沙箱 wrapper
#   - verify.py           包验证脚本
#   - dependencies.md     沙箱依赖清单（供 Hermes 读取）
#   - matplotlibrc        中文字体配置
#   - fonts/              公文字体目录（需包含仿宋_GB2312、楷体_GB2312 等 .ttf/.ttc 文件）
# =============================================================================

# ---- 阶段 1: 构建阶段（编译原生扩展）----
FROM python:3.12-slim-bookworm AS builder

# 换清华源，安装编译依赖
RUN sed -i 's|http://deb.debian.org|https://mirrors.tuna.tsinghua.edu.cn|g' /etc/apt/sources.list.d/debian.sources \
    && sed -i 's|http://security.debian.org|https://mirrors.tuna.tsinghua.edu.cn|g' /etc/apt/sources.list.d/debian.sources \
    && apt-get update && apt-get install -y --no-install-recommends \
    gcc g++ make cmake python3-dev libffi-dev \
    libxml2-dev libxslt1-dev libjpeg-dev libpng-dev \
    libfreetype6-dev liblcms2-dev libopenjp2-7-dev \
    libtiff-dev libwebp-dev zlib1g-dev \
    libpq-dev libssl-dev cargo libgraphviz-dev \
    && rm -rf /var/lib/apt/lists/*

# 升级 pip（清华源）
RUN pip install --no-cache-dir -i https://mirrors.tuna.tsinghua.edu.cn/pypi/web/simple \
    --upgrade pip setuptools wheel

# 仅放必须编译的包（无预编译 wheel 或需要系统头文件）
RUN pip install --no-cache-dir -i https://mirrors.tuna.tsinghua.edu.cn/pypi/web/simple --user \
    psycopg2-binary \
    python-calamine

# ---- 阶段 2: 运行时阶段 ----
FROM python:3.12-slim-bookworm

ENV PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1 \
    PIP_NO_CACHE_DIR=1 \
    PIP_DISABLE_PIP_VERSION_CHECK=1 \
    NODE_ENV=production \
    SAL_USE_VCLPLUGIN=gen \
    WORKDIR=/workspace

# =============================================================================
# 安装新版 pandoc
# =============================================================================
COPY ./pandoc-3.1.11-linux-amd64.tar.gz /tmp/
RUN tar xvzf /tmp/pandoc-3.1.11-linux-amd64.tar.gz --strip-components 1 -C /usr/local/ \
    && rm /tmp/pandoc-3.1.11-linux-amd64.tar.gz \
    && pandoc --version

# =============================================================================
# 系统依赖（换清华源）—— 增加 ffmpeg、procps(ps)、grep、fontconfig(字体缓存)
# =============================================================================
RUN sed -i 's|http://deb.debian.org|https://mirrors.tuna.tsinghua.edu.cn|g' /etc/apt/sources.list.d/debian.sources \
    && sed -i 's|http://security.debian.org|https://mirrors.tuna.tsinghua.edu.cn|g' /etc/apt/sources.list.d/debian.sources \
    && apt-get update && apt-get install -y --no-install-recommends \
    ca-certificates curl gnupg git unzip zip \
    procps grep psmisc \
    ffmpeg \
    libreoffice libreoffice-impress libreoffice-calc libreoffice-writer libreoffice-common \
    poppler-utils \
    libjpeg62-turbo libpng16-16 libfreetype6 liblcms2-2 libopenjp2-7 \
    libtiff6 libwebp7 libwebpmux3 libwebpdemux2 libxcb1 \
    libfontconfig1 \
    libpq5 \
    fonts-liberation fonts-dejavu fonts-noto-cjk \
    fonts-wqy-zenhei fonts-wqy-microhei \
    tesseract-ocr libtesseract-dev \
    graphviz libgraphviz-dev \
    libsqlcipher-dev \
    dbus-x11 \
    xvfb \
    sqlcipher \
    libstdc++6 \
    && rm -rf /var/lib/apt/lists/* \
    && apt-get clean

# =============================================================================
# 安装公文字体（仿宋_GB2312、楷体_GB2312 等）
# 将字体文件放在 Dockerfile 同级目录的 fonts/ 文件夹内，构建时自动导入
# =============================================================================
RUN mkdir -p /usr/share/fonts/truetype/custom
COPY ./fonts/* /usr/share/fonts/truetype/custom/
RUN fc-cache -fv /usr/share/fonts/truetype/custom \
    && fc-list : family | sort | uniq

# =============================================================================
# Node.js 22 + pptxgenjs
# =============================================================================
RUN curl -fsSL https://deb.nodesource.com/setup_22.x | bash - \
    && apt-get install -y --no-install-recommends nodejs \
    && rm -rf /var/lib/apt/lists/* \
    && npm install -g npm@latest

RUN npm install -g pptxgenjs@3.12.0 && npm cache clean --force

# =============================================================================
# Python 依赖
# =============================================================================
COPY --from=builder /root/.local /root/.local
ENV PATH=/root/.local/bin:$PATH

RUN pip install --no-cache-dir -i https://mirrors.tuna.tsinghua.edu.cn/pypi/web/simple --upgrade pip

# 批次 1: Office / 文档处理
RUN pip install --no-cache-dir -i https://mirrors.tuna.tsinghua.edu.cn/pypi/web/simple \
    markitdown[pptx] Pillow defusedxml lxml python-pptx python-docx \
    openpyxl xlrd xlsxwriter python-calamine olefile \
    PyMuPDF pdfplumber pytesseract pypandoc docxtpl

# 批次 2: 数据处理 / 科学计算 / 可视化
RUN pip install --no-cache-dir -i https://mirrors.tuna.tsinghua.edu.cn/pypi/web/simple \
    pandas numpy scipy chardet python-magic xlsx2csv unstructured \
    matplotlib seaborn plotly kaleido pygame

# 批次 3: 流程图 / 思维导图
RUN pip install --no-cache-dir -i https://mirrors.tuna.tsinghua.edu.cn/pypi/web/simple \
    graphviz pygraphviz diagrams networkx xmind xmindparser

# 批次 4: 数据库（通用 + 国产/Oracle）
RUN pip install --no-cache-dir -i https://mirrors.tuna.tsinghua.edu.cn/pypi/web/simple \
    pymongo redis sqlalchemy pymysql sqlcipher3 oracledb

# 批次 5: 达梦数据库（清华源可能没有，失败不中断构建）
RUN pip install --no-cache-dir -i https://mirrors.tuna.tsinghua.edu.cn/pypi/web/simple dmPython \
    || echo "WARNING: dmPython not found in Tsinghua mirror. Install manually from DM installer."

# 批次 6: 常用工具库
RUN pip install --no-cache-dir -i https://mirrors.tuna.tsinghua.edu.cn/pypi/web/simple \
    requests pyyaml jinja2 tabulate tqdm

# =============================================================================
# 验证 + 配置文件
# =============================================================================
COPY verify.py /tmp/verify.py
RUN python3 /tmp/verify.py

COPY dependencies.md /opt/hermes/dependencies.md
RUN chmod 644 /opt/hermes/dependencies.md

# 追加中文字体配置（保留原始 matplotlibrc 完整性，避免参数缺失）
# 用 sed 修改现有配置项（避免重复键）
RUN sed -i 's/^#*font.family:.*/font.family: sans-serif/' /usr/local/lib/python3.12/site-packages/matplotlib/mpl-data/matplotlibrc && \
    sed -i 's/^#*font.sans-serif:.*/font.sans-serif: WenQuanYi Zen Hei, Noto Sans CJK SC, SimHei, FangSong_GB2312, KaiTi_GB2312, DejaVu Sans, sans-serif/' /usr/local/lib/python3.12/site-packages/matplotlib/mpl-data/matplotlibrc && \
    sed -i 's/^#*axes.unicode_minus:.*/axes.unicode_minus: False/' /usr/local/lib/python3.12/site-packages/matplotlib/mpl-data/matplotlibrc

# =============================================================================
# LibreOffice 沙箱 wrapper
# =============================================================================
RUN mkdir -p /opt/hermes/scripts/office
COPY soffice.py /opt/hermes/scripts/office/soffice.py
RUN chmod +x /opt/hermes/scripts/office/soffice.py && \
    ln -sf /opt/hermes/scripts/office/soffice.py /usr/local/bin/soffice-hermes

# =============================================================================
# 工作目录
# =============================================================================
WORKDIR /workspace
RUN mkdir -p /workspace /tmp/soffice-work && chmod 1777 /tmp/soffice-work

# =============================================================================
# 健康检查
# =============================================================================
HEALTHCHECK --interval=30s --timeout=10s --start-period=5s --retries=3 \
    CMD python3 -c "import pandas, numpy, matplotlib, pptx; print('OK')" || exit 1

CMD ["python3"]