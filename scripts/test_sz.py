#!/usr/bin/env python3
"""测试 m:sz（脚本字号）是否正确添加到上标/下标/分数/重音中。

使用 ElementTree 命名空间感知的方式检索 <m:sz>，避免字符串匹配因
序列化前缀（ns0: / m:）不同而误判。
"""
import sys
import os

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
# 强制重新加载
for mod in list(sys.modules):
    if "build_paper" in mod:
        del sys.modules[mod]

from build_paper import latex_to_omath, SCRIPT_SZ_VAL
import xml.etree.ElementTree as ET

M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
SZ_TAG = f"{{{M}}}sz"


def count_sz(elem):
    """统计 elem 子树中 <m:sz> 的数量，并确认所有 sz 的 val 均为 SCRIPT_SZ_VAL。"""
    cnt = 0
    all_val_ok = True
    for sz in elem.iter(SZ_TAG):
        cnt += 1
        if sz.get(f"{{{M}}}val") != SCRIPT_SZ_VAL:
            all_val_ok = False
    return cnt, all_val_ok


tests = [
    (r"90^\circ", "角度符号(msup→°)", 1),
    (r"(-2)^3", "上标(msup)", 1),
    (r"x^2", "简单上标(msup)", 1),
    (r"x_1", "下标(msub)", 1),
    (r"\frac{3}{4}", "分数(mfrac)", 2),
    (r"\sqrt{16}", "根号(msqrt)", 0),
]

all_ok = True
for latex, desc, min_sz in tests:
    om = latex_to_omath(latex)
    sz_count, val_ok = count_sz(om)
    # 验证 sz 的值是 SCRIPT_SZ_VAL
    ok = sz_count >= min_sz and val_ok
    if not ok:
        all_ok = False
    texts = "".join(t.text or "" for t in om.iter(f"{{{M}}}t"))
    # 额外检查：角度符号必须用 °(U+00B0) 而非 ∘(U+2218)
    has_bad_ring = "\u2218" in texts
    has_good_degree = "\u00b0" in texts
    if "角度" in desc and has_bad_ring:
        ok = False
        all_ok = False
    status = "✅" if ok else "❌"
    ring_info = ""
    if "角度" in desc:
        ring_info = f" | ∘(bad)={has_bad_ring} °(good)={has_good_degree}"
    print(f"{status} [sz={sz_count:2d} val_ok={val_ok}{ring_info}] {desc:20s} | '{texts}'")

print()
if all_ok:
    print("✅ 全部通过：上标/下标/分数/角度符号均添加了更小字号属性(m:sz=\"2\")，度数使用标准°符号")
else:
    print("❌ 部分公式缺少 m:sz 属性或使用了错误的度数字符")
    sys.exit(1)
