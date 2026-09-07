#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
read_resume.py —— 解析简历文件为纯文本 / 结构化大纲（纯标准库，零依赖）。

用法:
    python read_resume.py <简历路径> [--outline] [--out 输出.txt]

支持: .docx (zip+XML 解析) / .txt / .md
不支持: .pdf / 图片 —— 这类文件请用 Read 工具直接读取（多模态可读 PDF 与图片）。
        .doc (旧版二进制) —— 请让用户另存为 .docx。
"""
import argparse
import os
import re
import sys
import zipfile

try:
    import xml.etree.ElementTree as ET
except ImportError:  # pragma: no cover
    ET = None

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def _text_of(node):
    """拼接 w:t / w:tab / w:br 文本。"""
    parts = []
    for t in node.iter():
        tag = t.tag
        if tag == W + "t":
            parts.append(t.text or "")
        elif tag == W + "tab":
            parts.append("\t")
        elif tag == W + "br":
            parts.append(" ")
    return "".join(parts).strip()


def _para_style(p):
    ppr = p.find(W + "pPr")
    if ppr is None:
        return None
    st = ppr.find(W + "pStyle")
    if st is not None:
        return st.get(W + "val") or ""
    return None


def _is_list(p):
    ppr = p.find(W + "pPr")
    return ppr is not None and ppr.find(W + "numPr") is not None


def _is_bold(p):
    for rpr in p.iter(W + "rPr"):
        b = rpr.find(W + "b")
        if b is not None and (b.get(W + "val") or "1").lower() not in ("0", "false", "off"):
            return True
    return False


def parse_docx(path, outline=False):
    """返回 (纯文本, 结构化行列表, 统计信息)"""
    if not zipfile.is_zipfile(path):
        raise ValueError("不是有效的 .docx（可能是伪装的 .doc 或已损坏）")

    lines = []
    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        target = "word/document.xml"
        if target not in names:
            raise ValueError("docx 中缺少 word/document.xml")
        xml = z.read(target)
        root = ET.fromstring(xml)
        body = root.find(W + "body")
        if body is None:
            raise ValueError("docx 结构异常：无 body")

        def walk(container):
            for child in container:
                if child.tag == W + "p":
                    txt = _text_of(child)
                    if not txt:
                        continue
                    style = _para_style(child) or ""
                    if outline:
                        lvl = re.search(r"(\d)", style)
                        if "heading" in style.lower() or style.lower().startswith("1") or style.lower().startswith("2"):
                            prefix = "#" * (int(lvl.group(1)) if lvl and int(lvl.group(1)) <= 4 else 2) + " "
                        elif _is_list(child):
                            prefix = "- "
                        else:
                            prefix = ""
                        bold = "*" if _is_bold(child) and len(txt) < 60 else ""
                        lines.append(f"{prefix}{bold}{txt}{bold}")
                    else:
                        lines.append(txt)
                elif child.tag == W + "tbl":
                    for tr in child.findall(W + "tr"):
                        cells = []
                        for tc in tr.findall(W + "tc"):
                            ct = " ".join(_text_of(p) for p in tc.findall(W + "p"))
                            cells.append(ct.strip())
                        row = " | ".join(c for c in cells if c)
                        if row:
                            lines.append("[表格] " + row if outline else row)
                    lines.append("")

        walk(body)

    plain = "\n".join(l for l in lines if l is not None)
    plain = re.sub(r"\n{3,}", "\n\n", plain).strip()
    return plain


def parse_text(path):
    with open(path, "r", encoding="utf-8", errors="replace") as f:
        return f.read()


def stats(text):
    cn = len(re.findall(r"[\u4e00-\u9fff]", text))
    bullets = [l for l in text.splitlines() if l.strip().startswith(("-", "•", "·", "*"))]
    nums = len(re.findall(r"\d", text))
    return {
        "总字符数": len(text),
        "中文字符数": cn,
        "行数": len([l for l in text.splitlines() if l.strip()]),
        "项目符号条数": len(bullets),
        "含数字字符数": nums,
    }


def main():
    ap = argparse.ArgumentParser(description="解析简历文件为纯文本（纯标准库）")
    ap.add_argument("path", help="简历文件路径 (.docx/.txt/.md)")
    ap.add_argument("--outline", action="store_true", help="输出结构化大纲（标题/列表/加粗标记）")
    ap.add_argument("--out", help="同时写入到指定文件")
    args = ap.parse_args()

    p = args.path
    if not os.path.exists(p):
        print(f"[错误] 文件不存在: {p}", file=sys.stderr)
        sys.exit(1)

    ext = os.path.splitext(p)[1].lower()
    try:
        if ext == ".docx":
            text = parse_docx(p, outline=args.outline)
        elif ext in (".txt", ".md", ".markdown"):
            text = parse_text(p)
        elif ext == ".pdf":
            print("[提示] .pdf 请直接用 Read 工具读取（可读 PDF 与图片），本脚本不处理。", file=sys.stderr)
            sys.exit(2)
        elif ext in (".png", ".jpg", ".jpeg", ".webp"):
            print("[提示] 图片请直接用 Read 工具读取。", file=sys.stderr)
            sys.exit(2)
        else:
            print(f"[错误] 不支持的格式: {ext}（.doc 请让用户另存为 .docx）", file=sys.stderr)
            sys.exit(1)
    except Exception as e:
        print(f"[错误] 解析失败: {e}", file=sys.stderr)
        sys.exit(1)

    header = "=" * 60 + f"\n简历解析结果：{os.path.basename(p)}\n" + "=" * 60
    st = stats(text)
    summary = "\n".join(f"  - {k}: {v}" for k, v in st.items())
    body = f"{header}\n\n{text}\n\n" + "-" * 60 + f"\n统计：\n{summary}\n"

    print(body)
    if args.out:
        os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
        with open(args.out, "w", encoding="utf-8") as f:
            f.write(body)
        print(f"[已写入] {args.out}")


if __name__ == "__main__":
    main()
