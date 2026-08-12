#!/usr/bin/env python3
"""WPS Office .et 文件读写工具函数

依赖: openpyxl, olefile, LibreOffice + xvfb
安装: pip install openpyxl olefile
      apt install libreoffice-calc xvfb

用法:
    rows = read_et("input.et")
    write_et(rows, "output.et")
"""
import subprocess, os, tempfile, shutil


def read_et(filepath: str, data_only=True) -> list[dict]:
    """读取 .et (WPS Spreadsheet) 文件，返回 [{列名: 值}, ...]

    Args:
        filepath: .et 文件路径
        data_only: True=读取计算结果，False=读取公式

    注意: 需要 DISPLAY=:99, SAL_USE_VCLPLUGIN=gen, Xvfb 运行中
    """
    tmpdir = tempfile.mkdtemp()
    try:
        env = os.environ.copy()
        env["SAL_USE_VCLPLUGIN"] = "gen"
        env["DISPLAY"] = os.environ.get("DISPLAY", ":99")
        subprocess.run(
            ["soffice", "--headless", "--convert-to", "xlsx",
             "--outdir", tmpdir, filepath],
            env=env, capture_output=True, timeout=120, check=True)

        import openpyxl
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
    """写入 .et (WPS Spreadsheet) 文件

    Args:
        data: [{列名: 值}, ...]
        output_path: 输出 .et 文件路径
    """
    import openpyxl
    tmpdir = tempfile.mkdtemp()
    try:
        headers = list(data[0].keys()) if data else []
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(headers)
        for row in data:
            ws.append([row.get(h, "") for h in headers])
        xlsx = os.path.join(tmpdir, "temp.xlsx")
        wb.save(xlsx)

        env = os.environ.copy()
        env["SAL_USE_VCLPLUGIN"] = "gen"
        env["DISPLAY"] = os.environ.get("DISPLAY", ":99")
        subprocess.run(
            ["soffice", "--headless", "--convert-to", "et",
             "--outdir", tmpdir, xlsx],
            env=env, capture_output=True, timeout=120, check=True)
        shutil.copy2(os.path.join(tmpdir, "temp.et"), output_path)
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)


if __name__ == "__main__":
    import sys
    if len(sys.argv) < 2:
        print("用法: python3 et_workbook.py <file.et> [输出.csv]")
        sys.exit(1)
    rows = read_et(sys.argv[1])
    if len(sys.argv) > 2:
        import csv
        with open(sys.argv[2], "w", newline="", encoding="utf-8") as f:
            w = csv.DictWriter(f, fieldnames=rows[0].keys())
            w.writeheader()
            w.writerows(rows)
    else:
        import json
        print(json.dumps(rows, ensure_ascii=False, indent=2))