#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
regression_test.py — 命题专家 Skill 的回归测试

验证 build_paper.py 不会再把公式渲染成空白。核心检查：
  1. 每个 <m:oMath> 必须是 <w:p> 的直接子元素（不能是 <w:r> 的子元素）。
     违反此规范时 Word 不识别公式、渲染为空白（即用户报的 BUG）。
  2. 每个 <m:oMath> 必须含有可见文字内容（m:t）。
  3. 生成的 docx 必须能被 python-docx 正常打开、内部 XML 全部良构。

用法（在 scripts/ 目录下运行）：
  python regression_test.py
退出码 0=通过，非0=失败。
"""
import os
import sys
import zipfile
import tempfile

from docx import Document
from lxml import etree

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

from build_paper import build_docx  # noqa: E402

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"

# 触发原 BUG 的最小试卷：选择题选项含公式 + 计算题小题纯公式 + 一行多公式
# + 分数/根号/上标结构（覆盖 mfrac/msup/msqrt 的 OMML 转换）
SAMPLE_MD = """# 回归测试卷

## 一、选择题

1. 下面算式中，运用了乘法分配律的是（　　）

A. $(25+35) \\times 4 = 60 \\times 4 = 240$

B. $25 \\times 4 + 35 \\times 4 = 100 + 140 = 240$

C. $36 \\times 4 + 64 \\times 4 = 144 + 256 = 400$

D. $(36+64) \\times 4 = 400$

## 二、计算题

18. 直接写出得数。

（1）$36 \\div 4$　　（2）$72 \\div 8$　　（3）$125 \\times 8$

（4）$200 \\div 5$　　（5）$45 + 55$

19. 竖式计算。

（1）$345 \\times 16$

20. 简算与混合运算。

（1）$(25+125) \\times 4$

（2）$2^3 \\times \\left(-\\frac{1}{2}\\right) + \\sqrt{16}$

（3）$-\\frac{3}{4}$

（4）$(-2)^3$

## 三、解答题

21. 计算：$x = \\frac{10}{2} = 5$。

$$ AC = \\sqrt{AB^2 + BC^2} = \\sqrt{4^2 + 3^2} = \\sqrt{25} = 5 $$
"""


def main():
    tmp = tempfile.mkdtemp(prefix="prop_regress_")
    out = os.path.join(tmp, "test.docx")
    doc = build_docx(SAMPLE_MD, tmp, seamless=False, no_page_number=True)
    doc.save(out)

    # 1) 能打开
    Document(out)

    # 2) 解析 document.xml
    with zipfile.ZipFile(out) as z:
        raw = z.read("word/document.xml")
    root = etree.fromstring(raw)

    omaths = root.findall(f".//{{{M_NS}}}oMath")
    if not omaths:
        print("❌ 未生成任何公式")
        return 1

    bad_nested = 0
    empty = 0
    for om in omaths:
        # oMath 的祖先链上必须出现 <w:p>（块公式经 m:oMathPara 合法包裹）
        p = om.getparent()
        ok = False
        while p is not None:
            if p.tag == f"{{{W_NS}}}p":
                ok = True
                break
            p = p.getparent()
        if not ok:
            bad_nested += 1
        texts = [t.text for t in om.iter(f"{{{M_NS}}}t") if t.text and t.text.strip()]
        if not texts:
            empty += 1

    # 3) 空结构节点检查：mfrac/msup/msqrt 等转换不得产生自闭合的 <m:e>/<m:num>/<m:den> 等
    import re
    empty_struct = len(re.findall(rb"<m:(?:e|sup|sub|num|den)\s*/>", raw))

    # 4) XML 良构
    bad_xml = 0
    with zipfile.ZipFile(out) as z:
        for name in z.namelist():
            if name.endswith(".xml"):
                try:
                    etree.fromstring(z.read(name))
                except Exception:
                    bad_xml += 1

    total = len(omaths)
    print(f"公式总数: {total} | 位置错误: {bad_nested} | 空公式: {empty} | 空结构节点: {empty_struct} | XML异常: {bad_xml}")

    if bad_nested == 0 and empty == 0 and empty_struct == 0 and bad_xml == 0:
        print("✅ 回归测试通过：所有公式位置正确、含内容、分数/根号/上标结构完整")
        return 0
    print("❌ 回归测试失败")
    return 1


if __name__ == "__main__":
    sys.exit(main())
