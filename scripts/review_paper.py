#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
review_paper.py — 审题人的机械校验工具（配合 exam-proposition-expert 技能）

对一份试卷 Markdown 做客观、可复现的检查，结果可直接贴入《审题报告》第五节：
  1. 满分与各题型分值合计是否相等（分值合计 = 满分）
  2. 各大题内题号是否连续（无跳号/重号）
  3. 选择题选项是否完整（A/B/C/D 齐全，或显式多于四项）
  4. 所有公式是否都能解析（解析失败 = docx 里公式会丢失或显示异常）
  5. 图片引用是否都存在文件

输出 Markdown 报告；亦可 `-o 审题报告_机械校验.md` 落盘。
注意：本脚本只做"机械"检查，科学性/规范性/保密性等需审题人主观判断，见 SKILL.md。
"""

import argparse
import os
import re
import sys

try:
    from latex2mathml.converter import convert as latex_to_mathml
except Exception as e:  # pragma: no cover
    sys.exit(f"[错误] 未安装 latex2mathml：请先 `pip install latex2mathml`。\n({e})")


def extract_math(md_text):
    """返回所有公式片段列表（块 $$...$$ 优先，再去重行内 $...$）。"""
    formulas = []
    # 块级 $$
    for m in re.finditer(r"\$\$([^$]+)\$\$", md_text, re.DOTALL):
        formulas.append(m.group(1).strip())
    # 行内 $...$，避免匹配到 $$ 之间
    stripped = re.sub(r"\$\$[^$]+\$\$", "", md_text)
    for m in re.finditer(r"(?<!\$)\$([^$\n]+?)\$(?!\$)", stripped):
        formulas.append(m.group(1).strip())
    return formulas


def check_scores(md_text):
    """满分 vs 各题型分值合计。返回 (total, section_totals, ok, detail)。"""
    total = None
    m = re.search(r"满分\s*(\d+)\s*分", md_text)
    if m:
        total = int(m.group(1))
    # 只在"试题部分"统计分值，忽略"参考答案与评分标准"里的分值标注（如"共 6 分"）
    ans_idx = re.search(r"##\s*参考答案", md_text)
    md_scores = md_text[:ans_idx.start()] if ans_idx else md_text
    section_totals = []
    # 形如 "（每小题 3 分，共 30 分）" 或 "（共 55 分）"
    for m in re.finditer(r"共\s*(\d+)\s*分", md_scores):
        section_totals.append(int(m.group(1)))
    # 也识别标题行里的 "（X 分）" 单值（注意：[^（\n] 不能跨行，否则会误吞答案里的"X 分"）
    for m in re.finditer(r"##\s+[^（\n]*?（\s*(\d+)\s*分）", md_scores):
        v = int(m.group(1))
        if v not in section_totals:
            section_totals.append(v)
    s = sum(section_totals)
    ok = (total is not None and s == total)
    detail = f"满分={total}，各题型分值合计={s}（{'+'.join(str(x) for x in section_totals)}）"
    return total, section_totals, ok, detail


def check_numbering(md_text):
    """各大题内题号连续性。返回 list of (section, issues)。"""
    issues = []
    # 按二级标题切分各大题
    parts = re.split(r"\n##\s+", md_text)
    for part in parts[1:]:
        head = part.splitlines()[0]
        nums = [int(m.group(1)) for m in re.finditer(r"^\s*(\d+)\.\s", part, re.MULTILINE)]
        # 也识别列表内的 "1."（解答题可能用列表）
        if not nums:
            continue
        expected = list(range(1, len(nums) + 1))
        if nums != expected:
            issues.append((head[:18], f"题号序列 {nums} 不连续（应为 {expected}）"))
    return issues


def check_choices(md_text):
    """选择题选项完整性。返回问题列表。"""
    issues = []
    # 以空行/题号切分每题；粗略：找含 'A.' 的段落
    # 先按题号切
    blocks = re.split(r"\n\s*\d+\.\s", md_text)
    opt_re = re.compile(r"([A-D])\.\s")
    for blk in blocks:
        if "A." in blk:
            found = set(opt_re.findall(blk))
            # 至少应有 A-D；若题目本身只有 A/B/C 三选项则忽略（少见）
            missing = set("ABCD") - found
            if missing and "D." not in blk:  # 可能本就是三选项
                # 只要缺了 A/B/C 任一项才算问题
                real_missing = missing & set("ABC")
                if real_missing:
                    issues.append(f"选项缺失 {sorted(real_missing)}：…{blk.strip()[:40]}")
    return issues


def check_math(md_text):
    """公式可解析性。返回 (total, failed_list)。"""
    formulas = extract_math(md_text)
    failed = []
    for f in formulas:
        try:
            latex_to_mathml(f)
        except Exception:
            failed.append(f)
    return len(formulas), failed


def check_images(md_text, base_dir):
    """图片引用是否都存在。返回 (refs, missing_list)。"""
    refs = re.findall(r"!\[[^\]]*\]\(([^)]+)\)", md_text)
    missing = []
    for p in refs:
        path = p.strip()
        if not os.path.isabs(path):
            path = os.path.join(base_dir, path)
        if not os.path.isfile(path):
            missing.append(p)
    return refs, missing


def main():
    ap = argparse.ArgumentParser(description="试卷机械校验（审题人辅助）")
    ap.add_argument("input", help="试卷 Markdown 路径（paper_draft.md 或 paper.md）")
    ap.add_argument("-o", "--output", help="将报告写入该文件（默认打印到标准输出）")
    args = ap.parse_args()
    if not os.path.isfile(args.input):
        sys.exit(f"[错误] 找不到输入文件：{args.input}")

    with open(args.input, "r", encoding="utf-8") as f:
        md = f.read()
    base = os.path.dirname(os.path.abspath(args.input))

    total, sec_totals, score_ok, score_detail = check_scores(md)
    num_issues = check_numbering(md)
    choice_issues = check_choices(md)
    math_total, math_failed = check_math(md)
    img_refs, img_missing = check_images(md, base)

    # 汇总
    results = []
    results.append(("各题分值合计 = 满分", "PASS" if score_ok else "FAIL", score_detail))
    results.append(("题号连续", "PASS" if not num_issues else "FAIL",
                    "全部连续" if not num_issues else "；".join(f"《{h}》{d}" for h, d in num_issues)))
    results.append(("选择题选项完整(A/B/C/D)", "PASS" if not choice_issues else "WARN",
                    "完整" if not choice_issues else "；".join(choice_issues)))
    results.append(("公式可解析为 Word 方程", "PASS" if not math_failed else "FAIL",
                    f"共 {math_total} 个公式" + ("" if not math_failed else f"，解析失败 {len(math_failed)} 个：{math_failed[:5]}")))
    results.append(("图片引用存在", "PASS" if not img_missing else "FAIL",
                    f"共 {len(img_refs)} 张" + ("" if not img_missing else f"，缺失：{img_missing}")))

    overall = "PASS" if all(r[1] in ("PASS", "WARN") for r in results) and score_ok and not num_issues and not math_failed and not img_missing else "NEED_FIX"

    lines = []
    lines.append("# 机械校验报告（审题人辅助）")
    lines.append("")
    lines.append(f"对象：`{os.path.basename(args.input)}`")
    lines.append("")
    lines.append("| 检查项 | 结论 | 说明 |")
    lines.append("| --- | --- | --- |")
    for name, verdict, detail in results:
        lines.append(f"| {name} | {verdict} | {detail} |")
    lines.append("")
    lines.append(f"**总体结论：{overall}**")
    lines.append("")
    lines.append("> 说明：本脚本仅做机械检查；科学性/规范性/保密性等需审题人主观判断（见 SKILL.md 第五节）。")
    report = "\n".join(lines)

    if args.output:
        with open(args.output, "w", encoding="utf-8") as f:
            f.write(report)
        print(f"[完成] 报告已写入 {args.output}")
    else:
        print(report)


if __name__ == "__main__":
    main()
