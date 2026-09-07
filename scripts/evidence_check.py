#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
evidence_check.py —— 简历证据链校验（防面试露馅）。

把简历 JSON 里的每一条内容与 `素材库.md` 做匹配，生成「证据链表」：
简历条目 / 素材出处 / 可追问层级 / 物证 / 状态，并给出覆盖率。

用法:
    python evidence_check.py --data 简历.json [--evidence 求职档案/素材库.md] [--out 求职档案/证据链.md]

匹配为字符二元组相似度的粗略匹配，仅作提示；最终由人工（和用户）确认。
"""
import argparse
import json
import os
import re
import sys
from datetime import date


def grams(s):
    s = re.sub(r"[^\u4e00-\u9fffA-Za-z0-9]", "", s or "")
    if len(s) < 2:
        return {s} if s else set()
    return {s[i:i + 2] for i in range(len(s) - 1)}


def similarity(a, b):
    """条目 a 有多少内容被素材 b 覆盖（非对称，条目短、素材长时比 Jaccard 准）。"""
    ga, gb = grams(a), grams(b)
    if not ga or not gb:
        return 0.0
    return len(ga & gb) / len(ga)


def load_evidence(path):
    """返回 [(位置标签, 文本)]，按空行切段。"""
    if not path or not os.path.exists(path):
        return []
    with open(path, "r", encoding="utf-8", errors="replace") as f:
        raw = f.read()
    blocks = []
    for chunk in re.split(r"\n\s*\n", raw):
        c = chunk.strip()
        if len(c) >= 10:
            blocks.append(c)
    return [(f"素材库 #{i + 1}", b.splitlines()[0][:40], b) for i, b in enumerate(blocks)]


def collect_items(data):
    """从简历 JSON 抽出所有可校验的条目：[(板块, 条目文本, 是否含量化数字)]"""
    out = []
    for sec in data.get("sections", []) or []:
        title = sec.get("title", "")
        stype = (sec.get("type") or "timeline").lower()
        for it in sec.get("items") or []:
            if isinstance(it, str):
                out.append((title, it, bool(re.search(r"\d", it))))
                continue
            if stype in ("timeline", "exp", "edu", "project"):
                head = "  |  ".join(x for x in (it.get("org", ""), it.get("role", "")) if x)
                if head:
                    out.append((title, head, False))
                for b in it.get("bullets", []) or []:
                    out.append((title, b, bool(re.search(r"\d", b))))
            elif stype == "kv":
                txt = f"{it.get('k') or it.get('key', '')}：{it.get('v') or it.get('value', '')}"
                out.append((title, txt, bool(re.search(r"\d", txt))))
            else:
                txt = it.get("text") or json.dumps(it, ensure_ascii=False)
                out.append((title, txt, bool(re.search(r"\d", txt))))
    return [x for x in out if x[1] and x[1].strip()]


def match(item, evidences, threshold=0.28, strong=0.45):
    """返回 (出处, 摘要, 分数, 等级)；等级 ∈ {ok, maybe, none}"""
    best, best_sim, best_label = "", 0.0, ""
    for label, head, body in evidences:
        sim = max(similarity(item, body), similarity(item, head) * 0.9)
        if sim > best_sim:
            best, best_sim, best_label = head, sim, label
    if best_sim >= strong:
        return f"{best_label}", best, best_sim, "ok"
    if best_sim >= threshold:
        return f"疑似 {best_label}（需确认）", best, best_sim, "maybe"
    return "⚠️ 素材库无对应记录", "", best_sim, "none"


def main():
    ap = argparse.ArgumentParser(description="简历证据链校验")
    ap.add_argument("--data", required=True, help="简历 JSON")
    ap.add_argument("--evidence", help="素材库 .md 路径")
    ap.add_argument("--out", help="写出 证据链.md 的路径")
    ap.add_argument("--threshold", type=float, default=0.28, help="疑似匹配阈值，默认 0.28")
    args = ap.parse_args()

    if not os.path.exists(args.data):
        print(f"[错误] JSON 不存在: {args.data}", file=sys.stderr)
        sys.exit(1)
    with open(args.data, "r", encoding="utf-8") as f:
        data = json.load(f)

    evidences = load_evidence(args.evidence)
    if args.evidence and not evidences:
        print(f"[提示] 素材库为空或不存在：{args.evidence}（所有条目都会标为待补）")

    items = collect_items(data)
    rows = []
    for title, text, has_num in items:
        label, head, sim, level = match(text, evidences, args.threshold)
        if level == "none":
            risk = "高风险" if has_num else "待补"
        elif level == "maybe":
            risk = "待确认"
        else:
            risk = "OK"
        rows.append({
            "板块": title, "条目": text, "出处": label,
            "摘要": head, "相似度": f"{sim:.2f}", "可追问层级": "待填",
            "物证": "待填", "状态": risk,
        })

    total = len(rows)
    covered = sum(1 for r in rows if r["状态"] in ("OK", "待确认"))
    unsure = sum(1 for r in rows if r["状态"] == "待确认")
    high = [r for r in rows if r["状态"] == "高风险"]
    pct = (covered / total * 100) if total else 0

    lines = [
        f"# 证据链校验（{date.today().isoformat()}）",
        "",
        f"- 简历条目：**{total}** 条",
        f"- 能追溯到素材库：**{covered}** 条（**{pct:.0f}%**，其中待人工确认 {unsure} 条）",
        f"- 高风险（无素材但写了数字）：**{len(high)}** 条",
        "",
        "判据：任一条目被追问三层答不出 = 不安全。处理办法只有两个——补素材，或从简历删掉。",
        "",
        "| # | 板块 | 简历条目 | 素材出处 | 可追问层级 | 物证 | 状态 |",
        "|---|---|---|---|---|---|---|",
    ]
    for i, r in enumerate(rows, 1):
        cell = r["条目"].replace("|", "/").replace("\n", " ")
        if len(cell) > 60:
            cell = cell[:58] + "…"
        lines.append(f"| {i} | {r['板块']} | {cell} | {r['出处']} | {r['可追问层级']} | {r['物证']} | {r['状态']} |")

    lines += ["", "## 待处理", ""]
    if high:
        lines.append("**高风险（优先处理，写不出来就删）：**")
        for r in high:
            lines.append(f"- [ ] {r['条目'][:50]} → 补素材或改回保守版本")
        lines.append("")
    if covered < total:
        lines.append("**其余待补素材：**")
        for r in rows:
            if r["状态"] == "待补":
                lines.append(f"- [ ] {r['条目'][:50]}")
    if not high and covered == total:
        lines.append("- 全部条目都有素材支撑。下一步：逐条自述 30 秒，讲不顺的继续补。")

    report = "\n".join(lines)
    print("=" * 62)
    print(f"证据链校验：{os.path.basename(args.data)}")
    print(f"  条目 {total} 条 | 有素材 {covered} 条（{pct:.0f}%） | 高风险 {len(high)} 条")
    print("=" * 62)
    if pct < 70:
        print("⚠️  覆盖率 < 70%：这份简历有一批内容你讲不出细节，面试一追问就露馅。")
        print("    先补素材或删条目，不要投。")
    elif pct < 90:
        print("⚠️  覆盖率 < 90%：可投，但先处理高风险条目，别面第一志愿公司。")
    else:
        print("✅ 覆盖率达标。下一步：逐条自述 30 秒做口述演练。")
    if not os.path.exists(args.evidence or ""):
        print("\n[提示] 未指定有效素材库（--evidence），所有条目都按「无素材」计。")

    if args.out:
        os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
        with open(args.out, "w", encoding="utf-8") as f:
            f.write(report + "\n")
        print(f"\n[已写入] {args.out}")
    else:
        print("\n" + "-" * 62)
        print(report)


if __name__ == "__main__":
    main()
