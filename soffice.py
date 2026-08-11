#!/usr/bin/env python3
"""
LibreOffice 沙箱环境 wrapper
解决 soffice 在容器/沙箱环境中无头模式挂起或崩溃
"""
import subprocess
import sys
import os
import shutil
import tempfile


def main():
    soffice_bin = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice_bin:
        print("ERROR: soffice not found", file=sys.stderr)
        sys.exit(1)

    env = os.environ.copy()
    env["SAL_USE_VCLPLUGIN"] = "gen"
    env["SAL_NO_MOUSE_GRABS"] = "1"

    # 构建命令：优先尝试直接运行，如果环境有 xvfb 则兜底
    xvfb = shutil.which("xvfb-run")
    if xvfb:
        # xvfb-run 包裹，防止 X11 初始化失败
        cmd = [xvfb, "-a", "--server-args=-screen 0 1024x768x24", soffice_bin]
    else:
        cmd = [soffice_bin]

    cmd.extend(sys.argv[1:])

    with tempfile.TemporaryDirectory(prefix="soffice_") as tmpdir:
        env["HOME"] = tmpdir
        user_config = os.path.join(tmpdir, ".config", "libreoffice", "4", "user")
        os.makedirs(user_config, exist_ok=True)
        registry = os.path.join(user_config, "registrymodifications.xcu")
        with open(registry, "w") as f:
            f.write('<?xml version="1.0" encoding="UTF-8"?>\n')
            f.write('<oor:items xmlns:oor="http://openoffice.org/2001/registry">\n')
            f.write('  <item oor:path="/org.openoffice.Office.Common/Misc"><prop oor:name="ShowTipOfTheDay" oor:op="fuse"><value>false</value></prop></item>\n')
            f.write('  <item oor:path="/org.openoffice.Office.Common/Misc"><prop oor:name="AutoSave" oor:op="fuse"><value>false</value></prop></item>\n')
            f.write('  <item oor:path="/org.openoffice.Office.Recovery/RecoveryInfo"><prop oor:name="CreateBackup" oor:op="fuse"><value>false</value></prop></item>\n')
            f.write('</oor:items>\n')

        try:
            result = subprocess.run(cmd, env=env, capture_output=True, text=True, timeout=300)
            if result.stdout:
                print(result.stdout, end="")
            if result.stderr:
                print(result.stderr, end="", file=sys.stderr)
            sys.exit(result.returncode)
        except subprocess.TimeoutExpired:
            print("ERROR: soffice timed out after 300s", file=sys.stderr)
            sys.exit(1)
        except KeyboardInterrupt:
            sys.exit(130)


if __name__ == "__main__":
    main()