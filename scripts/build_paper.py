#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
build_paper.py — 命题专家技能配套脚本（自包含，无需 pandoc）

把按约定写好的 Markdown 试卷（含 LaTeX 公式与图片引用）转换为
可直接打印的 Word 试卷（.docx）。

原理：
  1. 轻量 Markdown 解析器把试卷拆成标题/段落/列表/表格/图片/公式块。
  2. LaTeX 公式经 latex2mathml 转 MathML，再转成 Word 原生可编辑方程（OMML），
     因此公式是 Word 里能编辑的方程，而不是图片，绝不会乱。
  3. 图片用 python-docx 直接嵌入。
  4. 后处理统一中文字体、A4 页面、页边距、页码，可选左侧"密封线"。

依赖：python-docx、latex2mathml（可选 matplotlib 用于生成配图）。
用法：
  python build_paper.py paper.md -o 试卷.docx [--seamless] [--no-page-number]
"""

import argparse
import os
import re
import sys
from lxml import etree as lxml_et
import xml.etree.ElementTree as ET

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.shared import Pt, Cm, Mm, Emu

try:
    from latex2mathml.converter import convert as latex_to_mathml
except Exception as e:  # pragma: no cover
    sys.exit(f"[错误] 未安装 latex2mathml：请先 `pip install latex2mathml`。\n({e})")

# 注册命名空间，确保序列化时使用标准前缀（m:/wp:/a:/wps:），
# 否则 lxml 会生成 ns0 等前缀，虽然 Word 仍能识别，但统一前缀更稳妥。
try:
    from lxml import etree as _lxml_etree
    for _p, _u in {
        "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
        "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
        "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
        "wps": "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
        "wpg": "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
    }.items():
        _lxml_etree.register_namespace(_p, _u)
except Exception:
    pass

# 中文字体方案（标准中文试卷排版）
BODY_ASCII = "Times New Roman"
BODY_EA = "宋体"
HEAD_EA = "黑体"
TITLE_SIZE = Pt(16)
SECTION_SIZE = Pt(13)
SUBSECTION_SIZE = Pt(11.5)
BODY_SIZE = Pt(10.5)


# ----------------------------------------------------------------------------
# 工具：命名空间
# ----------------------------------------------------------------------------
def m(local):
    """创建 OMML(math) 命名空间元素。"""
    return OxmlElement(f"m:{local}")


def localname(tag):
    if not isinstance(tag, str):
        return None
    return tag.split("}", 1)[1] if "}" in tag else tag


# ----------------------------------------------------------------------------
def _fix_degree_symbol(omath):
    """后处理：把上标/重音中的 ∘(U+2218 RING OPERATOR) 替换为 °(U+00B0 DEGREE SIGN)。

    LaTeX 的 \\circ 在数学排版中产生 U+2218（环运算符），
    但「角度/度数」场景下应该用 U+00B0（度符号）。
    U+2218 在非 Word 环境（IMA/WPS/LibreOffice 等）中常被渲染为
    大圆圈或附带多余字符，而 U+00B0 是通用标准字符，跨平台一致。
    """
    M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
    RING = "\u2218"   # ∘  RING OPERATOR
    DEGREE = "\u00b0" # °  DEGREE SIGN

    for t in omath.iter(f"{{{M_NS}}}t"):
        if t.text and RING in t.text:
            t.text = t.text.replace(RING, DEGREE)
    return omath


# MathML 结构规整化
# ----------------------------------------------------------------------------
_CLOSING = {")": "(", "]": "[", "}": "{", "\u27e9": "\u27e8",
            "\u2309": "\u2308", "\u230b": "\u230a"}  # 右→左 括号映射

_MATH = "http://www.w3.org/1998/Math/MathML"


def _normalize_mathml(math_elem):
    """规整 latex2mathml 输出的畸形 MathML 结构。

    已知问题：latex2mathml 对 ``(-2)^3`` 生成
      <mrow><mo>(</mo><mo>−</mo><mn>2</mn>
        <msup><mo>)</mo><mn>3</mn></msup></mrow>
    即右括号被塞进 msup 底数、其余内容散落在 mrow 中。
    正确结构应为整个 (-2) 作为 msup 底数。

    本函数扫描 msup/msub/msubsup，若底数为右括号类字符则向前回溯
    匹配的左括号，将中间所有元素包裹进 <mrow> 作为新底数。
    """
    import copy

    for tag in ("msup", "msub", "msubsup"):
        for node in math_elem.iter(f"{{{_MATH}}}{tag}"):
            children = list(node)
            if not children:
                continue
            base = children[0]  # msup/msub: base; msubsup: base
            base_text = (base.text or "").strip()
            # 检查底数是否是单个右括号类字符
            if base_text not in _CLOSING and len(list(base)) == 0:
                # 也检查只有文本子节点的 mo 元素
                t_nodes = list(base.iter(f"{{{_MATH}}}t"))
                if not t_nodes or (t_nodes[0].text or "").strip() not in _CLOSING:
                    continue
                base_text = (t_nodes[0].text or "").strip()

            opening = _CLOSING.get(base_text)
            if opening is None:
                continue

            parent = node.getparent()
            if parent is None:
                continue
            siblings = list(parent)

            # 找 node 在 siblings 中的位置和匹配的左括号
            idx = siblings.index(node)
            match_idx = None
            depth = 0
            # 从 node 前面向后找（跳过嵌套的同类型标签）
            for i in range(idx - 1, -1, -1):
                sib = siblings[i]
                sib_tag = localname(sib.tag)
                sib_txt = "".join(sib.itertext()).strip()
                if sib_tag == tag:
                    depth += 1
                elif depth > 0 and (sib_tag == "mrow" and any(localname(c.tag) == tag for c in sib)):
                    depth -= 1
                elif depth == 0 and sib_txt == opening:
                    match_idx = i
                    break

            if match_idx is None:
                continue

            # 把 siblings[match_idx .. idx-1] + 原底数包进 <mrow> 当新底数
            new_base = lxml_et.Element(f"{{{_MATH}}}mrow")
            for i in range(match_idx, idx):
                new_base.append(copy.deepcopy(siblings[i]))
            new_base.append(copy.deepcopy(base))  # deepcopy！lxml.append会移动元素

            # 替换底数
            node.remove(base)
            node.insert(0, new_base)

            # 删除原位置上被吸进去的兄弟节点（从后往前删避免索引漂移）
            for i in range(idx - 1, match_idx - 1, -1):
                parent.remove(siblings[i])

    return math_elem


# LaTeX -> Word 原生方程（OMML）
# ----------------------------------------------------------------------------
def latex_to_omath(latex):
    """把一段 LaTeX 转成 <m:oMath> 元素；失败则退化为纯文本。"""
    try:
        mathml = latex_to_mathml(latex)
        # 用 lxml 解析（支持 .getparent()，规整化需要）
        root = lxml_et.fromstring(mathml.encode("utf-8"))
        _normalize_mathml(root)   # 修复 latex2mathml 的畸形结构
        omath = mathml_to_omath(root)
        return _fix_degree_symbol(omath)  # 度数符号标准化
    except Exception as e:
        sys.stderr.write(f"[提示] 公式转换失败，已退化为文本：{latex} ({e})\n")
        o = m("oMath")
        r = m("r")
        t = m("t")
        t.text = latex
        r.append(t)
        o.append(r)
        return o


def _children_e(parent_omml, mm_parent):
    """把 MathML 子节点转换后追加进 OMML 父节点（mrow/mstyle 直接展开）。"""
    for child in list(mm_parent):
        ln = localname(child.tag)
        if ln is None:
            continue
        if ln in ("mrow", "mstyle"):
            _children_e(parent_omml, child)
        elif ln in ("mphantom",):
            # 占位符：忽略内容
            continue
        else:
            parent_omml.append(_convert(child))


def mathml_to_omath(math_root):
    omath = m("oMath")
    # math_root 可能是 <math> 或已是表达式
    _children_e(omath, math_root)
    return omath


# OOXML m:sz 值：2 = smaller（上标/下标标准大小）
SCRIPT_SZ_VAL = "2"


def _add_script_sz(omml_elem):
    """给 OMML 元素内所有 <m:r> 添加 <m:sz m:val="2"/>（更小字号），
    用于上标/下标内容，使其在 Word 中以正确的较小尺寸渲染。"""
    M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
    for r in omml_elem.iter(f"{{{M_NS}}}r"):
        # 如果已有 rPr，追加 sz；否则新建
        rpr = r.find(f"{{{M_NS}}}rPr")
        if rpr is None:
            rpr = m("rPr")
            # 把 rpr 插到第一个子元素前（保持 r 内顺序：rPr 在前，t 在后）
            if len(r):
                r.insert(0, rpr)
            else:
                r.append(rpr)
        sz = m("sz")
        sz.set(qn("m:val"), SCRIPT_SZ_VAL)
        rpr.append(sz)


def _make_run(elem, script_size=False):
    ln = localname(elem.tag)
    text = "".join(elem.itertext())
    r = m("r")
    sty_val = None
    mv = elem.get("mathvariant")
    if mv:
        sty_val = {"italic": "i", "bold": "b", "bold-italic": "bi",
                   "normal": "p", "double-struck": "p", "script": "p"}.get(mv)
    else:
        if ln == "mi":
            sty_val = "i"
        elif ln == "mtext":
            sty_val = "p"
    if sty_val or script_size:
        rpr = m("rPr")
        if sty_val:
            sty = m("sty")
            sty.set(qn("m:val"), sty_val)
            rpr.append(sty)
        if script_size:
            sz = m("sz")
            sz.set(qn("m:val"), SCRIPT_SZ_VAL)
            rpr.append(sz)
        r.append(rpr)
    t = m("t")
    t.set(qn("xml:space"), "preserve")
    t.text = text
    r.append(t)
    return r


def _e_arg(mm_elem):
    """生成 m:e 参数节点（用于分组 / 作为 f、sup 等的参数）。"""
    e = m("e")
    _children_e(e, mm_elem)
    return e


def _unwrap_e(elem):
    """若 elem 是 <m:e>，返回其子元素列表（拆包）；否则返回 [elem]。

    用于 msup/msub 等的 base 参数：_convert(mrow) 会返回 <m:e>，
    而外层 handler 又创建了 base=<m:e>，导致双层嵌套。
    此函数消除多余的一层。
    """
    M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
    if elem is not None and localname(elem.tag) == "e":
        return list(elem)
    return [elem]


def _e_from(child):
    """把【单个】MathML 元素转成 <m:e> 包裹的 OMML。

    正确处理叶子（mn/mi/mo）与容器（mrow/mfrac/msup 等）：
      _children_e 期望传入「容器」并迭代其 children，若传入叶子（无 children）
      则什么都不填 —— 这正是之前上标/分数/根号内容丢失的根因。
    """
    e = m("e")
    e.append(_convert(child))
    return e


def _convert(elem):
    ln = localname(elem.tag)
    kids = [c for c in list(elem) if localname(c.tag) is not None]

    if ln in ("mi", "mn", "mo", "mtext"):
        return _make_run(elem)
    if ln == "mspace":
        r = m("r")
        t = m("t")
        t.set(qn("xml:space"), "preserve")
        t.text = " "
        r.append(t)
        return r
    if ln in ("mrow", "mstyle"):
        return _e_arg(elem)
    if ln == "mfrac":
        f = m("f")
        f.append(m("fPr"))
        num = m("num")
        den = m("den")
        if len(kids) >= 1:
            num.append(_e_from(kids[0]))
            _add_script_sz(num)
        if len(kids) >= 2:
            den.append(_e_from(kids[1]))
            _add_script_sz(den)
        f.append(num)
        f.append(den)
        return f
    if ln == "msqrt":
        rad = m("rad")
        rad.append(m("radPr"))
        deg = m("deg")
        rad.append(deg)
        e = m("e")
        if kids:
            e.append(_convert(kids[0]))
        rad.append(e)
        return rad
    if ln == "mroot":
        rad = m("rad")
        rad.append(m("radPr"))
        deg = m("deg")
        e = m("e")
        if len(kids) >= 2:
            deg.append(_convert(kids[1]))
            _add_script_sz(deg)  # 根指数用更小字号
            e.append(_convert(kids[0]))
        rad.append(deg)
        rad.append(e)
        return rad
    if ln == "msup":
        wrap = m("sSup")       # 上标函数容器（不是 m:sup！）
        base = m("e")
        exp = m("sup")         # 上标内容
        if len(kids) >= 1:
            for child in _unwrap_e(_convert(kids[0])):
                base.append(child)
        if len(kids) >= 2:
            exp.append(_convert(kids[1]))
        _add_script_sz(exp)  # 上标用更小字号
        wrap.append(base)
        wrap.append(exp)
        return wrap
    if ln == "msub":
        wrap = m("sSub")       # 下标函数容器（不是 m:sub！）
        base = m("e")
        sub = m("sub")         # 下标内容
        if len(kids) >= 1:
            for child in _unwrap_e(_convert(kids[0])):
                base.append(child)
        if len(kids) >= 2:
            sub.append(_convert(kids[1]))
        _add_script_sz(sub)  # 下标用更小字号
        wrap.append(base)
        wrap.append(sub)
        return wrap
    if ln == "msubsup":
        wrap = m("subSup")
        base = m("e")
        sub = m("sub")
        sup = m("sup")
        if len(kids) >= 1:
            for child in _unwrap_e(_convert(kids[0])):
                base.append(child)
        if len(kids) >= 2:
            sub.append(_convert(kids[1]))
            _add_script_sz(sub)  # 下标用更小字号
        if len(kids) >= 3:
            sup.append(_convert(kids[2]))
            _add_script_sz(sup)  # 上标用更小字号
        wrap.append(base)
        wrap.append(sub)
        wrap.append(sup)
        return wrap
    if ln == "mfenced":
        d = m("d")
        dpr = m("dPr")
        beg = m("begChr")
        beg.set(qn("m:val"), elem.get("open", "(") or "(")
        end = m("endChr")
        end.set(qn("m:val"), elem.get("close", ")") or ")")
        dpr.append(beg)
        dpr.append(end)
        d.append(dpr)
        e = m("e")
        for c in kids:
            e.append(_convert(c))
        d.append(e)
        return d
    if ln == "mtable":
        tbl = m("m")
        tbl.append(m("mPr"))
        for row in kids:
            if localname(row.tag) != "mtr":
                continue
            mr = m("mr")
            cells = [c for c in list(row) if localname(c.tag) == "mtd"]
            for cell in cells:
                e = m("e")
                _children_e(e, cell)
                mr.append(e)
            tbl.append(mr)
        return tbl
    if ln in ("mover", "munder"):
        acc = m("acc")
        accpr = m("accPr")
        # 取 mo 作为重音符号
        mo_text = ""
        for c in kids:
            if localname(c.tag) == "mo":
                mo_text = "".join(c.itertext())
                break
        chr_el = m("chr")
        chr_el.set(qn("m:val"), mo_text or "^")
        accpr.append(chr_el)
        acc.append(accpr)
        base = m("e")
        if kids:
            base.append(_convert(kids[0]))
        acc.append(base)
        # mover 的符号（如度数圈）也用更小字号
        _add_script_sz(acc)
        return acc
    if ln == "mphantom":
        # 占位：返回空 e
        return m("e")
    # 兜底：当作分组
    return _e_arg(elem)


# ----------------------------------------------------------------------------
# 中文字体 / 排版
# ----------------------------------------------------------------------------
def set_run_fonts(run, ascii_font, ea_font, size=None):
    run.font.name = ascii_font
    if size is not None:
        run.font.size = size
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)
    rfonts.set(qn("w:ascii"), ascii_font)
    rfonts.set(qn("w:hAnsi"), ascii_font)
    rfonts.set(qn("w:eastAsia"), ea_font)
    rfonts.set(qn("w:cs"), ascii_font)


def set_style_font(style, ascii_font, ea_font, size=None, bold=None, align=None):
    try:
        rpr = style.element.get_or_add_rPr()
    except Exception:
        return
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)
    rfonts.set(qn("w:ascii"), ascii_font)
    rfonts.set(qn("w:hAnsi"), ascii_font)
    rfonts.set(qn("w:eastAsia"), ea_font)
    rfonts.set(qn("w:cs"), ascii_font)
    if size is not None:
        style.font.size = size
    if bold is not None:
        style.font.bold = bold
    if align is not None:
        style.paragraph_format.alignment = align


def style_document(doc):
    for style in doc.styles:
        try:
            set_style_font(style, BODY_ASCII, BODY_EA)
        except Exception:
            pass
    names = [s.name for s in doc.styles]
    if "Normal" in names:
        set_style_font(doc.styles["Normal"], BODY_ASCII, BODY_EA, size=BODY_SIZE)
    if "Heading 1" in names:
        set_style_font(doc.styles["Heading 1"], BODY_ASCII, HEAD_EA,
                       size=TITLE_SIZE, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    if "Heading 2" in names:
        set_style_font(doc.styles["Heading 2"], BODY_ASCII, HEAD_EA,
                       size=SECTION_SIZE, bold=True)
    if "Heading 3" in names:
        set_style_font(doc.styles["Heading 3"], BODY_ASCII, HEAD_EA,
                       size=SUBSECTION_SIZE, bold=True)
    # 正文 run 补中文
    for p in doc.paragraphs:
        for r in p.runs:
            if r._element.find(qn("m:oMath")) is not None:
                continue
            set_run_fonts(r, BODY_ASCII, BODY_EA)
    for tbl in doc.tables:
        for row in tbl.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for r in p.runs:
                        set_run_fonts(r, BODY_ASCII, BODY_EA)


def set_page_setup(doc):
    for section in doc.sections:
        section.page_width = Mm(210)
        section.page_height = Mm(297)
        section.top_margin = Cm(2.5)
        section.bottom_margin = Cm(2.5)
        section.left_margin = Cm(2.5)
        section.right_margin = Cm(2.5)


def add_field(paragraph, field):
    b = OxmlElement("w:fldChar")
    b.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = f" {field} "
    sep = OxmlElement("w:fldChar")
    sep.set(qn("w:fldCharType"), "separate")
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    r = paragraph.add_run()
    r._r.append(b)
    r._r.append(instr)
    r._r.append(sep)
    r._r.append(end)


def add_page_number_footer(doc):
    section = doc.sections[0]
    footer = section.footer
    p = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r0 = p.add_run("第 ")
    set_run_fonts(r0, BODY_ASCII, BODY_EA, Pt(9))
    add_field(p, "PAGE")
    r1 = p.add_run(" 页 / 共 ")
    set_run_fonts(r1, BODY_ASCII, BODY_EA, Pt(9))
    add_field(p, "NUMPAGES")
    r2 = p.add_run(" 页")
    set_run_fonts(r2, BODY_ASCII, BODY_EA, Pt(9))


def add_seam_line(doc):
    try:
        from docx.oxml.ns import nsmap
        nsmap.setdefault("wps",
            "http://schemas.microsoft.com/office/word/2010/wordprocessingShape")
        nsmap.setdefault("wpg",
            "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup")
        nsmap.setdefault("wpc",
            "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingCanvas")

        section = doc.sections[0]
        header = section.header
        hp = header.paragraphs[0] if header.paragraphs else header.add_paragraph()

        pos_left = Emu(int(Cm(0.4)))
        pos_top = Emu(int(Cm(105)))
        box_w = Emu(int(Cm(0.8)))
        box_h = Emu(int(Cm(85)))

        anchor = OxmlElement("wp:anchor")
        for k, v in {
            "distT": "0", "distB": "0", "distL": "0", "distR": "0",
            "simplePos": "0", "relativeHeight": "251658240",
            "behindDoc": "1", "locked": "0", "layoutInCell": "1",
            "allowOverlap": "1",
        }.items():
            anchor.set(qn("wp:" + k), v)
        sp = OxmlElement("wp:simplePos")
        sp.set(qn("wp:x"), "0")
        sp.set(qn("wp:y"), "0")
        anchor.append(sp)
        ph = OxmlElement("wp:positionH")
        ph.set(qn("wp:relativeFrom"), "page")
        ph_off = OxmlElement("wp:posOffset")
        ph_off.text = str(int(pos_left))
        ph.append(ph_off)
        anchor.append(ph)
        pv = OxmlElement("wp:positionV")
        pv.set(qn("wp:relativeFrom"), "page")
        pv_off = OxmlElement("wp:posOffset")
        pv_off.text = str(int(pos_top))
        pv.append(pv_off)
        anchor.append(pv)
        ext = OxmlElement("wp:extent")
        ext.set(qn("wp:cx"), str(int(box_w)))
        ext.set(qn("wp:cy"), str(int(box_h)))
        anchor.append(ext)
        eff = OxmlElement("wp:effectExtent")
        for k, v in {"l": "0", "t": "0", "r": "0", "b": "0"}.items():
            eff.set(qn("wp:" + k), v)
        anchor.append(eff)
        anchor.append(OxmlElement("wp:wrapNone"))
        docpr = OxmlElement("wp:docPr")
        docpr.set(qn("wp:id"), "99")
        docpr.set(qn("wp:name"), "SeamLine")
        anchor.append(docpr)
        anchor.append(OxmlElement("wp:cNvGraphicFramePr"))

        graphic = OxmlElement("a:graphic")
        gdata = OxmlElement("a:graphicData")
        gdata.set(qn("a:uri"),
                  "http://schemas.microsoft.com/office/word/2010/wordprocessingShape")
        wsp = OxmlElement("wps:wsp")
        txbx = OxmlElement("wps:txbx")
        content = OxmlElement("w:txbxContent")
        para = OxmlElement("w:p")
        ppr = OxmlElement("w:pPr")
        td = OxmlElement("w:textDirection")
        td.set(qn("w:val"), "tbRl")
        ppr.append(td)
        para.append(ppr)
        r = OxmlElement("w:r")
        rpr = OxmlElement("w:rPr")
        rfonts = OxmlElement("w:rFonts")
        rfonts.set(qn("w:eastAsia"), "宋体")
        rpr.append(rfonts)
        sz = OxmlElement("w:sz")
        sz.set(qn("w:val"), "24")
        rpr.append(sz)
        color = OxmlElement("w:color")
        color.set(qn("w:val"), "808080")
        rpr.append(color)
        r.append(rpr)
        t = OxmlElement("w:t")
        t.set(qn("xml:space"), "preserve")
        t.text = "密 封 线"
        r.append(t)
        para.append(r)
        content.append(para)
        txbx.append(content)
        wsp.append(txbx)
        gdata.append(wsp)
        graphic.append(gdata)
        anchor.append(graphic)

        drawing = OxmlElement("w:drawing")
        drawing.append(anchor)
        run = hp.add_run()
        run._r.append(drawing)
    except Exception as e:
        sys.stderr.write(f"[提示] 密封线插入失败，已跳过（不影响试卷主体）：{e}\n")


# ----------------------------------------------------------------------------
# 轻量 Markdown 解析 -> docx
# ----------------------------------------------------------------------------
INLINE_RE = re.compile(r"\$([^$]+)\$|(\*\*[^*]+\*\*)")


def tokenize_inline(s):
    tokens = []
    pos = 0
    for mm in INLINE_RE.finditer(s):
        if mm.start() > pos:
            tokens.append(("text", s[pos:mm.start()]))
        if mm.group(1) is not None:
            tokens.append(("math", mm.group(1)))
        else:
            tokens.append(("bold", mm.group(2)[2:-2]))
        pos = mm.end()
    if pos < len(s):
        tokens.append(("text", s[pos:]))
    return tokens


def fill_inline(paragraph, text):
    has_math = False
    for kind, val in tokenize_inline(text):
        if kind == "text":
            if val == "":
                continue
            r = paragraph.add_run(val)
            set_run_fonts(r, BODY_ASCII, BODY_EA)
        elif kind == "bold":
            r = paragraph.add_run(val)
            r.bold = True
            set_run_fonts(r, BODY_ASCII, BODY_EA)
        else:  # math —— 行内公式
            # 关键修复：<m:oMath> 必须是 <w:p> 的直接子元素，
            # 不能放进 <w:r> 内部，否则 Word 不识别、渲染为空白。
            has_math = True
            paragraph._p.append(latex_to_omath(val))
    # 包含内联公式的段落：强制设置行距，防止上标/度数符号顶部被截断。
    # 使用较大的 atLeast 值（约 1.9x 正文号），确保上标/角度符号完整显示。
    if has_math:
        pf = paragraph.paragraph_format
        pf.line_spacing_rule = WD_LINE_SPACING.AT_LEAST
        pf.line_spacing = Pt(20)


def is_table_start(lines, i):
    if "|" not in lines[i]:
        return False
    nxt = lines[i + 1].strip() if i + 1 < len(lines) else ""
    return bool(re.match(r"^\s*\|?[\s:\-|]+\|?\s*$", nxt)) and "-" in nxt


def remove_table_borders(tbl):
    tblPr = tbl._tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        e = OxmlElement(f"w:{edge}")
        e.set(qn("w:val"), "none")
        e.set(qn("w:sz"), "0")
        e.set(qn("w:space"), "0")
        e.set(qn("w:color"), "auto")
        borders.append(e)
    tblPr.append(borders)


def build_docx(md_text, md_dir, seamless, no_page_number):
    doc = Document()
    lines = md_text.split("\n")
    n = len(lines)
    i = 0
    while i < n:
        line = lines[i]
        strip = line.strip()
        if strip == "":
            i += 1
            continue
        if strip.startswith("#"):
            level = len(strip) - len(strip.lstrip("#"))
            title = strip[level:].strip()
            doc.add_heading(title, level=min(level, 3))
            i += 1
        elif strip in ("---", "***", "___"):
            p = doc.add_paragraph()
            pPr = p._p.get_or_add_pPr()
            pBdr = OxmlElement("w:pBdr")
            bottom = OxmlElement("w:bottom")
            bottom.set(qn("w:val"), "single")
            bottom.set(qn("w:sz"), "6")
            bottom.set(qn("w:space"), "1")
            bottom.set(qn("w:color"), "999999")
            pBdr.append(bottom)
            pPr.append(pBdr)
            i += 1
        elif re.match(r"^!\[[^\]]*\]\([^)]+\)$", strip):
            m = re.match(r"^!\[[^\]]*\]\(([^)]+)\)$", strip)
            path = m.group(1).strip()
            if not os.path.isabs(path):
                path = os.path.join(md_dir, path)
            if os.path.isfile(path):
                p = doc.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                try:
                    p.add_run().add_picture(path, width=Cm(9))
                except Exception:
                    p.add_run(f"[图片缺失: {path}]")
            else:
                p = doc.add_paragraph()
                p.add_run(f"[图片缺失: {path}]")
                set_run_fonts(p.runs[0], BODY_ASCII, BODY_EA)
            i += 1
        elif strip.startswith("$$") and strip.endswith("$$"):
            latex = strip[2:-2].strip()
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            op = OxmlElement("m:oMathPara")
            op.append(latex_to_omath(latex))
            p._p.append(op)
            i += 1
        elif is_table_start(lines, i):
            # 解析表格
            header = [c.strip() for c in strip.strip().strip("|").split("|")]
            i += 2  # 跳过分隔行
            body = []
            while i < n and "|" in lines[i] and lines[i].strip():
                row = [c.strip() for c in lines[i].strip().strip("|").split("|")]
                body.append(row)
                i += 1
            cols = max(len(header), *(len(r) for r in body), 1)
            tbl = doc.add_table(rows=1, cols=cols)
            try:
                tbl.style = "Table Grid"
            except Exception:
                pass
            # 表头
            for c in range(cols):
                cell = tbl.rows[0].cells[c]
                cell.text = ""
                fill_inline(cell.paragraphs[0], header[c] if c < len(header) else "")
            for r in body:
                cells = tbl.add_row().cells
                for c in range(cols):
                    cells[c].text = ""
                    fill_inline(cells[c].paragraphs[0], r[c] if c < len(r) else "")
            # 考生信息栏（含下划线填空）去边框
            joined = " ".join(header + [x for r in body for x in r])
            if "____" in joined:
                remove_table_borders(tbl)
        elif re.match(r"^\d+\.\s", strip) or strip.startswith("- ") or strip.startswith("* "):
            # 列表
            items = []
            while i < n and lines[i].strip():
                s = lines[i].strip()
                mm = re.match(r"^(\d+\.)\s+(.*)$", s)
                if mm:
                    items.append(("num", mm.group(2)))
                    i += 1
                elif s.startswith("- ") or s.startswith("* "):
                    items.append(("bul", s[2:]))
                    i += 1
                else:
                    break
            for kind, content in items:
                style = "List Number" if kind == "num" else "List Bullet"
                try:
                    p = doc.add_paragraph(style=style)
                except Exception:
                    p = doc.add_paragraph()
                fill_inline(p, content)
        else:
            p = doc.add_paragraph()
            fill_inline(p, line)
            i += 1

    style_document(doc)
    set_page_setup(doc)
    if not no_page_number:
        add_page_number_footer(doc)
    if seamless:
        add_seam_line(doc)
    return doc


def main():
    ap = argparse.ArgumentParser(description="Markdown 试卷 -> 可打印 docx（自包含）")
    ap.add_argument("input", help="输入的 paper.md 路径")
    ap.add_argument("-o", "--output", required=True, help="输出 docx 路径")
    ap.add_argument("--seamless", action="store_true", help="添加左侧竖排'密封线'")
    ap.add_argument("--no-page-number", action="store_true", help="不加页码")
    args = ap.parse_args()

    if not os.path.isfile(args.input):
        sys.exit(f"[错误] 找不到输入文件：{args.input}")

    with open(args.input, "r", encoding="utf-8") as f:
        md_text = f.read()

    print("[1/3] 解析 Markdown 并转换公式为 Word 原生方程...")
    doc = build_docx(md_text, os.path.dirname(os.path.abspath(args.input)),
                     args.seamless, args.no_page_number)

    out_dir = os.path.dirname(os.path.abspath(args.output))
    os.makedirs(out_dir, exist_ok=True)
    print("[2/3] 统一中文排版 / A4 / 页码...")
    doc.save(args.output)
    print(f"[3/3] 完成 -> {args.output}")


if __name__ == "__main__":
    main()
