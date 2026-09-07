#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
resume_lint.py —— 简历规则体检（纯标准库，零依赖）。

在人工诊断之前先跑一遍，把机器能查的硬伤清掉。输出问题清单（含行号）+ 扣分 + 摘要。

用法:
    python resume_lint.py <简历.txt> [--jd 关键词.txt] [--json]
    --jd   : 每行一个 JD 关键词，检查是否出现在简历中（ATS 关键词覆盖）
    --json : 以 JSON 输出，便于程序处理
"""
import argparse
import json
import os
import re
import sys

WEAK_VERBS = ["负责", "参与", "协助", "帮忙", "进行", "开展", "从事", "有关", "相关", "完成日常工作"]
EMPTY_WORDS = ["勤奋", "踏实", "吃苦耐劳", "乐观开朗", "性格开朗", "学习能力强",
               "有团队精神", "抗压能力强", "热爱", "善于与人沟通", "责任心强"]
HOBBY_HINT = ["兴趣爱好", "爱好", "特长", "自我评价"]
ID_CARD = re.compile(r"\b\d{17}[\dXx]\b")
PHONE = re.compile(r"(?<!\d)1[3-9]\d(?:[\s\-]?\d){8}(?!\d)")
EMAIL = re.compile(r"[\w.\-]+@[\w.\-]+\.\w+")
YEAR = re.compile(r"(20\d{2})[.\-/年]\s*(\d{1,2})?")
CN_NUM = re.compile(r"\d")


def load(path):
    with open(path, "r", encoding="utf-8", errors="replace") as f:
        return f.read()


def bullets(text):
    out = []
    for i, line in enumerate(text.splitlines(), 1):
        s = line.strip()
        if not s:
            continue
        if s.startswith(("-", "•", "·", "*", "▪", "●", "○")) or re.match(r"^\d+[.、)]", s):
            body = re.sub(r"^[-•·*▪●○]\s*|^\d+[.、)]\s*", "", s)
            out.append((i, body))
    return out


def check(text):
    issues = []
    lines = text.splitlines()
    bs = bullets(text)

    def add(level, lineno, msg, fix):
        issues.append({"级别": level, "行": lineno, "问题": msg, "建议": fix})

    # 1. 联系方式
    if not PHONE.search(text):
        add("错误", 0, "未检测到手机号", "补充手机号，写成 138-0000-0000 便于拨打")
    if not EMAIL.search(text):
        add("错误", 0, "未检测到邮箱", "补充邮箱；勿用 QQ 邮箱默认昵称，改为姓名拼音")
    for m in EMAIL.finditer(text):
        local = m.group(0).split("@")[0]
        if re.match(r"^\d{6,}$", local):
            add("警告", 0, f"邮箱疑似纯 QQ 号：{m.group(0)}", "改为 姓名拼音@xx.com，避免暴露 QQ 昵称")
    if ID_CARD.search(text):
        add("错误", 0, "简历中出现疑似身份证号", "删除；身份证号绝不应出现在简历里")
    if re.search(r"(家庭住址|详细地址|身份证)", text):
        add("警告", 0, "出现家庭住址/身份证等无关隐私信息", "删除，只需写意向城市")

    # 2. bullet 层面
    no_num = 0
    for ln, b in bs:
        if not CN_NUM.search(b):
            no_num += 1
    if bs:
        ratio = no_num / len(bs)
        if ratio > 0.6:
            add("错误", 0, f"{no_num}/{len(bs)} 条 bullet 完全没有数字（{int(ratio*100)}%）",
                "用量化五法补：绝对值 / 增幅 / 占比 / 对比 / 效率")
        elif ratio > 0.35:
            add("警告", 0, f"{no_num}/{len(bs)} 条 bullet 无数字（{int(ratio*100)}%）",
                "优先给最重要的 3 条补数字")

    for ln, b in bs:
        head = b[:4]
        if any(b.startswith(v) for v in WEAK_VERBS) and len(b) < 30:
            add("警告", ln, f"以弱动词开头且无结果：{b[:28]}",
                "改为 强动词+具体内容+方法+结果，如'统筹12人团队…到场600人次(+40%)'")
        if b.startswith(("我们", "咱们", "大家")):
            add("错误", ln, f"主语是「我们」：{b[:28]}", "改为主语省略或'我'，写清个人贡献")
        if len(b) > 90:
            add("警告", ln, f"bullet 过长（{len(b)} 字）", "拆成两条，或删掉修饰语，单条 ≤ 2 行")
        if re.search(r"[a-zA-Z]", b) and re.search(r"，|,", b) and re.search(r"[，]", b) and re.search(r"[,]", b):
            add("提示", ln, "中英文标点混用", "统一用中文全角标点，英文/数字间加空格")

    # 3. 整体层面
    cn = len(re.findall(r"[\u4e00-\u9fff]", text))
    if cn < 300:
        add("警告", 0, f"中文内容仅 {cn} 字，信息量可能不足", "补经历细节；一页简历通常 600-1200 字")
    if cn > 1600:
        add("警告", 0, f"中文内容 {cn} 字，大概率超出一页", "删废话，而不是缩小字号")

    for kw in HOBBY_HINT:
        for i, line in enumerate(lines, 1):
            if kw not in line or len(line) >= 200:
                continue
            # 标题行可能在上一行，故连同其后 3 行一起检查
            seg = "\n".join(lines[i - 1:i + 3]).strip()
            if any(w in seg for w in EMPTY_WORDS):
                add("警告", i, f"「{kw}」段落含无证据形容词：{seg[:34]}",
                    "删除，或每一句形容词后跟一个可验证事实")
            if re.search(r"(读书|旅游|听音乐|看电影|运动|唱歌)", seg) and not re.search(r"\d|完赛|证书|级|名", seg):
                add("提示", i, f"「{kw}」无成果：{seg[:34]}",
                    "要么删，要么改成有成果的（如'马拉松完赛3次''校辩论队一辩'）")
            break  # 每个关键词只报一次

    course_line = [ (i,l) for i,l in enumerate(lines,1) if ("课程" in l and l.count("、") >= 4) ]
    for i, l in course_line:
        n = l.count("、") + 1
        if n > 6:
            add("警告", i, f"罗列 {n} 门课程", "只留 3-5 门与岗位最相关的")

    # 4. 时间倒序（只比较每段经历的起始时间）
    starts = [int(m.group(1)) for m in re.finditer(r"(20\d{2})[.\-/年]\s*\d{1,2}\s*[-–—~至]", text)]
    if len(starts) >= 3:
        inc = sum(1 for a, b in zip(starts, starts[1:]) if b < a)
        if inc >= 2:
            add("提示", 0, f"经历顺序出现 {inc} 处时间倒退（更早的经历排在更新的之后）", "同一板块内按开始时间倒序：最近的排最前")

    # 5. 求职意向
    if not re.search(r"(求职意向|意向岗位|应聘岗位|目标岗位)", text):
        add("提示", 0, "未写明求职意向", "在个人信息下方写一行'求职意向：XX岗'（只写一个）")
    m = re.search(r"(求职意向|意向岗位|应聘岗位|目标岗位)[：: ]*(.+)", text)
    if m:   # 只取求职意向本身，忽略后面的联系方式
        first = re.split(r"\||\n", m.group(2))[0]
        opts = [o.strip() for o in re.split(r"[/、,，;；]| and ", first) if o.strip()]
        if len(opts) > 1:
            add("警告", 0, f"求职意向写了多个方向：{first[:30]}", "一份简历只投一个岗位，多投多版本")

    seen, uniq = set(), []
    for i in issues:
        key = (i["级别"], i["行"], i["问题"][:40])
        if key not in seen:
            seen.add(key)
            uniq.append(i)
    return uniq


def jd_coverage(text, jd_path):
    kws = [l.strip() for l in load(jd_path).splitlines() if l.strip()]
    hit, miss = [], []
    for k in kws:
        (hit if k in text else miss).append(k)
    return {"总数": len(kws), "命中": hit, "缺失": miss}


DEDUCT = {"错误": 8, "警告": 3, "提示": 1}


def main():
    ap = argparse.ArgumentParser(description="简历规则体检")
    ap.add_argument("path")
    ap.add_argument("--jd", help="JD 关键词文件（每行一个），检查 ATS 关键词覆盖")
    ap.add_argument("--json", action="store_true")
    args = ap.parse_args()

    if not os.path.exists(args.path):
        print(f"[错误] 文件不存在: {args.path}", file=sys.stderr)
        sys.exit(1)

    text = load(args.path)
    issues = check(text)
    score = max(0, 100 - sum(DEDUCT[i["级别"]] for i in issues))
    result = {"文件": os.path.basename(args.path), "规则体检分": score, "问题数": len(issues), "问题": issues}
    if args.jd:
        result["JD关键词覆盖"] = jd_coverage(text, args.jd)

    if args.json:
        print(json.dumps(result, ensure_ascii=False, indent=2))
        return

    print("=" * 62)
    print(f"简历体检：{result['文件']}    规则体检分 {score}/100    问题 {len(issues)} 条")
    print("=" * 62)
    order = {"错误": 0, "警告": 1, "提示": 2}
    for i in sorted(issues, key=lambda x: (order[x["级别"]], x["行"])):
        loc = f"第{i['行']}行" if i["行"] else "全文"
        print(f"[{i['级别']}] {loc}  {i['问题']}")
        print(f"        建议：{i['建议']}")
    if args.jd:
        cov = result["JD关键词覆盖"]
        print("-" * 62)
        print(f"JD 关键词覆盖：{len(cov['命中'])}/{cov['总数']}")
        if cov["缺失"]:
            print("  缺失（建议自然融入简历）：" + "、".join(cov["缺失"][:20]))
    print("-" * 62)
    print("说明：本脚本只查机器可判定的硬伤，不能替代人工判断。")
    print("      下一步 → 按 references/05-resume-review.md 的 12 项评分卡做人工诊断。")


if __name__ == "__main__":
    main()
