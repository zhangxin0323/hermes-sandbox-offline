# =============================================================================
# Hermes Full-Stack Sandbox for Windows / WSL2 / Docker Desktop
# =============================================================================
# 基础: Python 3.12 (slim) + Node.js 22 + LibreOffice + Poppler + Pandoc + Chromium
# 功能: PPT/Word/Excel 处理 + 数据处理 + 数据库(含国产) + 流程图/思维导图 + OCR
#       + 视频处理 + 无头浏览器(网页调试/截图) + Python Web 框架(Flask/FastAPI)
#
# 构建:
#   docker build -t hermes-sandbox-python312:20261027v1 -f Dockerfile .
#
# 外部依赖文件（需放在 Dockerfile 同级目录）:
#   - soffice.py          LibreOffice 沙箱 wrapper
#   - verify.py           包验证脚本
#   - dependencies.md     沙箱依赖清单（供 Hermes 读取）
#   - matplotlibrc        中文字体配置
#   - fonts/              公文字体目录（需包含仿宋_GB2312、楷体_GB2312 等 .ttf/.ttc 文件）
# =============================================================================

# ---- 阶段 2: 运行时阶段 ----
FROM python:3.12-slim-bookworm

ENV PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1 \
    PIP_NO_CACHE_DIR=1 \
    PIP_DISABLE_PIP_VERSION_CHECK=1 \
    NODE_ENV=production \
    SAL_USE_VCLPLUGIN=gen \
    PLAYWRIGHT_BROWSERS_PATH=/opt/playwright-browsers \
    WORKDIR=/workspace

# =============================================================================
# 安装新版 pandoc
# =============================================================================
COPY ./pandoc-3.1.11-linux-amd64.tar.gz /tmp/
RUN tar xvzf /tmp/pandoc-3.1.11-linux-amd64.tar.gz --strip-components 1 -C /usr/local/ \
    && rm /tmp/pandoc-3.1.11-linux-amd64.tar.gz \
    && pandoc --version

# =============================================================================
# 系统依赖（换清华源）
#   - 基础工具: ffmpeg、procps(ps)、grep、fontconfig
#   - 无头浏览器: Chromium/Playwright 所需的 X11/图形/网络库
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
    # ---- Chromium / Playwright 无头浏览器系统依赖 ----
    libnss3 libnspr4 \
    libatk1.0-0 libatk-bridge2.0-0 libatspi2.0-0 \
    libcups2 libdrm2 libxkbcommon0 libxcomposite1 libxdamage1 \
    libxfixes3 libxrandr2 libgbm1 libpango-1.0-0 libcairo2 \
    libasound2 \
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
# Python 依赖 —— 合并为单次 pip install，避免重复安装/卸载
#   - psycopg2-binary / python-calamine 有预编译 wheel，无需 builder 编译阶段
#   - 新增: playwright、flask、fastapi、uvicorn、geographiclib、charset_normalizer
# =============================================================================
RUN pip install --no-cache-dir -i https://mirrors.tuna.tsinghua.edu.cn/pypi/web/simple --upgrade pip \
    && pip install --no-cache-dir -i https://mirrors.tuna.tsinghua.edu.cn/pypi/web/simple \
    # Office / 文档处理
    markitdown[pptx] Pillow defusedxml lxml python-pptx python-docx \
    openpyxl xlrd xlsxwriter python-calamine olefile \
    PyMuPDF pdfplumber pytesseract pypandoc docxtpl \
    # 数据处理 / 科学计算 / 可视化
    pandas numpy scipy chardet python-magic xlsx2csv unstructured \
    matplotlib seaborn plotly kaleido pygame \
    # 流程图 / 思维导图
    graphviz pygraphviz diagrams networkx xmind xmindparser \
    # 数据库（通用 + 国产/Oracle）
    psycopg2-binary pymongo redis sqlalchemy pymysql sqlcipher3 oracledb \
    # 常用工具库
    requests pyyaml jinja2 tabulate tqdm \
    charset_normalizer geographiclib \
    # Python Web 框架（调试网页 App）
    flask fastapi "uvicorn[standard]" \
    # 无头浏览器自动化（网页调试、截图、布局检查）
    playwright

# 批次 5: 达梦数据库（清华源可能没有，失败不中断构建）
RUN pip install --no-cache-dir -i https://mirrors.tuna.tsinghua.edu.cn/pypi/web/simple dmPython \
    || echo "WARNING: dmPython not found in Tsinghua mirror. Install manually from DM installer."

# =============================================================================
# 安装 Playwright 无头 Chromium 浏览器（含 ffmpeg 编解码器）
# PLAYWRIGHT_BROWSERS_PATH=/opt/playwright-browsers 已在 ENV 中设置
# =============================================================================
RUN PLAYWRIGHT_DOWNLOAD_HOST=https://cdn.npmmirror.com/binaries/playwright \
    playwright install chromium --with-deps \
    && playwright install-deps chromium \
    && rm -rf /var/lib/apt/lists/* \
    && apt-get clean

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
    CMD python3 -c "import pandas, numpy, matplotlib, pptx, playwright, flask, fastapi; print('OK')" || exit 1

CMD ["python3"]