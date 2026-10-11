# -*- coding: utf-8 -*-
"""从原始《项目报告模板.docx》生成生成器底版 assets/template-base.docx。

底版保留：标题封面（无边框表格）、封面间隔段、三个分节属性（页码分段）、
目录/图表目录骨架（标题段 + TOC 域占位段）、空的正文节。
删除：签字封面表格、示例正文、示例目录条目。

一般无需手动运行——assets/template-base.docx 已随 skill 分发。
仅当模板文件更新后需要重新生成底版时运行：
    python make_template_base.py
"""
import copy
import sys
from pathlib import Path

from docx import Document
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls, qn

SKILL_DIR = Path(__file__).resolve().parents[1]
ORIGINAL = SKILL_DIR / "assets" / "项目报告模板.docx"
OUT = SKILL_DIR / "assets" / "template-base.docx"

NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def W(tag):
    return "{%s}%s" % (NS, tag)


def para_text(el):
    # 注意：python-docx 的元素类重写了 .text，itertext() 会把 p/r/t 三层文本
    # 重复统计，这里只取 w:t 节点。
    return "".join(t.text or "" for t in el.findall(".//" + W("t"))).strip()


def main():
    doc = Document(ORIGINAL)
    body = doc.element.body
    children = list(body)

    # ---- 索引定位（对应原模板的顶层结构）----
    cover_tbl = children[0]        # 封面表格（标题封面，签字栏留空）
    spacer_p = children[1]         # 封面表格后的间隔空段（style 21）
    sect1_p = children[5]          # 携带第 1 节 sectPr 的空段
    empty_p = children[6]          # 目录节起始空段
    sdt = children[7]              # 目录 sdt（从中取"目录"标题段）
    pagebrk_p = children[8]        # 目录后的手动分页段
    fig_title_p = children[9]      # "图片"标题段（af2 + 黑体 + 居中）
    sect2_p = children[29]         # 携带第 2 节 sectPr 的段落
    final_sectPr = children[-1]    # body 级第 3 节 sectPr

    # ---- 取"目录"标题段（pStyle TOC + 黑体 + 居中）----
    toc_title_p = None
    for p in sdt.iter(W("p")):
        if para_text(p) == "目录":
            toc_title_p = p
            break
    if toc_title_p is None:
        sys.exit("未在 sdt 中找到“目录”标题段，原模板结构可能已变化")

    # ---- 复制需要的元素 ----
    keep = [copy.deepcopy(el) for el in
            (cover_tbl, spacer_p, sect1_p, empty_p, toc_title_p,
             pagebrk_p, fig_title_p, sect2_p, final_sectPr)]
    cover_tbl_c, spacer_c, sect1_c, empty_c, toc_title_c, pagebrk_c, \
        fig_title_c, sect2_c, final_sectPr_c = keep

    # "图片"标题段去掉渲染缓存标记
    for el in fig_title_c.findall(".//" + W("lastRenderedPageBreak")):
        el.getparent().remove(el)

    # 第 2 节分节段里的手动分页符删除：分节符本身就会另起新页，
    # 保留会多出一张空白罗马页。
    for r in sect2_c.findall(W("r")):
        sect2_c.remove(r)

    # 复制"表格"标题段（与"图片"标题段同构）
    tab_title_c = copy.deepcopy(fig_title_c)
    for t in tab_title_c.iter(W("t")):
        t.text = "表格"

    # ---- 清空 body 并重建 ----
    for ch in children:
        body.remove(ch)

    toc_main_host = parse_xml(
        '<w:p %s><w:pPr><w:pStyle w:val="TOC1"/></w:pPr>'
        '<w:r><w:t>[[TOC_MAIN]]</w:t></w:r></w:p>' % nsdecls("w"))
    fig_host = parse_xml(
        '<w:p %s><w:pPr><w:pStyle w:val="af2"/>'
        '<w:tabs><w:tab w:val="right" w:leader="dot" w:pos="8296"/></w:tabs>'
        '<w:ind w:left="960" w:hanging="480"/></w:pPr>'
        '<w:r><w:t>[[TOC_FIG]]</w:t></w:r></w:p>' % nsdecls("w"))
    tab_host = parse_xml(
        '<w:p %s><w:pPr><w:pStyle w:val="af2"/>'
        '<w:tabs><w:tab w:val="right" w:leader="dot" w:pos="8296"/></w:tabs>'
        '<w:ind w:left="960" w:hanging="480"/></w:pPr>'
        '<w:r><w:t>[[TOC_TAB]]</w:t></w:r></w:p>' % nsdecls("w"))

    for el in (cover_tbl_c, spacer_c, sect1_c, empty_c, toc_title_c,
               toc_main_host, pagebrk_c, fig_title_c, fig_host,
               tab_title_c, tab_host, sect2_c, final_sectPr_c):
        body.append(el)

    # ---- settings.xml 注入 updateFields（打开时自动刷新目录/SEQ/页码域）----
    settings = doc.settings.element
    if settings.find(W("updateFields")) is None:
        uf = parse_xml('<w:updateFields %s w:val="true"/>' % nsdecls("w"))
        anchor = None
        for tag in ("hdrShapeDefaults", "footnotePr", "endnotePr",
                    "compat", "rsids"):
            anchor = settings.find(W(tag))
            if anchor is not None:
                break
        if anchor is not None:
            anchor.addprevious(uf)
        else:
            settings.append(uf)

    doc.save(OUT)
    print("written:", OUT)


if __name__ == "__main__":
    main()
