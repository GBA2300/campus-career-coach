#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
export_pdf.py —— 把 .docx 简历转成 PDF（本地免费方案，文件不上传）。

按优先级自动尝试：
  1. LibreOffice（开源免费）：soffice --headless --convert-to pdf
  2. Microsoft Word（COM，装机自带）
  3. WPS Writer（COM，装机自带）

全部不可用时，打印免费的替代方案指引（不推荐任何需要付费的工具）。

用法:
    python export_pdf.py "<简历.docx>" [--out "<输出.pdf>"]
"""
import argparse
import glob
import os
import shutil
import subprocess
import sys
import tempfile

LIBRE_CANDIDATES = [
    r"C:\Program Files\LibreOffice\program\soffice.exe",
    r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    r"C:\Program Files\LibreOffice 25\program\soffice.exe",
    r"C:\Program Files\LibreOffice 7\program\soffice.exe",
]

PS_TEMPLATE = """$ErrorActionPreference = 'Stop'
$src = '{src}'
$dst = '{dst}'
$app = $null
foreach ($progid in @('{progid}', 'Word.Application', 'KWPS.Application')) {{
  try {{ $app = New-Object -ComObject $progid; break }} catch {{ }}
}}
if ($null -eq $app) {{ exit 3 }}
$app.Visible = $false
try {{
  $doc = $app.Documents.Open($src, $false, $true)
  $doc.SaveAs([ref]$dst, [ref]17)
  $doc.Close(0)
  Write-Output 'OK'
}} finally {{
  $app.Quit()
}}
"""

GUIDE = """
[提示] 本机没找到 LibreOffice / Word / WPS，无法自动转换。用下面任一**免费**方式即可：

  1) 用 Word 或 WPS 打开简历 → 文件 → 另存为 / 导出 → PDF     （装机自带，最稳，推荐）
  2) 装 LibreOffice（开源免费）https://www.libreoffice.org/ 装好后本脚本可一键转换
  3) 在线免费档：PDF24 https://tools.pdf24.org/  （无限制、无水印）
                 ILovePDF https://www.ilovepdf.com/

  隐私提醒：简历含身份证号、家庭住址时，优先用本地方式（1 或 2），不要上传不明站点。
"""


def find_libreoffice():
    for p in LIBRE_CANDIDATES:
        if os.path.exists(p):
            return p
    for p in glob.glob(r"C:\Program Files\LibreOffice*\program\soffice.exe"):
        return p
    return shutil.which("soffice") or shutil.which("libreoffice")


def convert_libreoffice(soffice, src, out_dir):
    cmd = [soffice, "--headless", "--norestore", "--convert-to", "pdf", "--outdir", out_dir, src]
    r = subprocess.run(cmd, capture_output=True, text=True, timeout=180)
    return r.returncode == 0


def convert_com(progid, src, dst):
    ps = PS_TEMPLATE.format(src=src.replace("'", "''"), dst=dst.replace("'", "''"), progid=progid)
    fd, path = tempfile.mkstemp(suffix=".ps1")
    try:
        with os.fdopen(fd, "w", encoding="utf-8-sig") as f:
            f.write(ps)
        r = subprocess.run(
            ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", path],
            capture_output=True, text=True, timeout=180)
        return "OK" in (r.stdout or "") and os.path.exists(dst)
    except Exception:
        return False
    finally:
        try:
            os.remove(path)
        except OSError:
            pass


def main():
    ap = argparse.ArgumentParser(description="docx → pdf（本地免费方案）")
    ap.add_argument("path", help="简历 .docx 路径")
    ap.add_argument("--out", help="输出 .pdf 路径（默认与 docx 同名同目录）")
    args = ap.parse_args()

    src = os.path.abspath(args.path)
    if not os.path.exists(src):
        print(f"[错误] 文件不存在: {src}", file=sys.stderr)
        sys.exit(1)
    if not src.lower().endswith(".docx"):
        print("[错误] 只支持 .docx", file=sys.stderr)
        sys.exit(1)

    dst = os.path.abspath(args.out) if args.out else os.path.splitext(src)[0] + ".pdf"
    os.makedirs(os.path.dirname(dst), exist_ok=True)

    soffice = find_libreoffice()
    if soffice:
        print(f"[尝试] LibreOffice：{soffice}")
        tmp_dir = os.path.dirname(dst)
        if convert_libreoffice(soffice, src, tmp_dir):
            produced = os.path.join(tmp_dir, os.path.splitext(os.path.basename(src))[0] + ".pdf")
            if os.path.exists(produced):
                if os.path.abspath(produced) != os.path.abspath(dst):
                    shutil.move(produced, dst)
                print(f"[成功] PDF 已生成：{dst}")
                return
        print("[失败] LibreOffice 转换未成功，尝试下一种方式…")

    for progid, name in (("Word.Application", "Microsoft Word"), ("KWPS.Application", "WPS Writer")):
        print(f"[尝试] {name} (COM)")
        if convert_com(progid, src, dst):
            print(f"[成功] PDF 已生成：{dst}")
            return
        print(f"[失败] 未检测到可用的 {name}")

    print(GUIDE)
    sys.exit(2)


if __name__ == "__main__":
    main()
