from __future__ import annotations

import re
import shutil
import zipfile
from pathlib import Path
from tempfile import NamedTemporaryFile


REF_HEAD = (
    r'<w:p><w:pPr><w:pStyle w:val="1" /></w:pPr>'
    r'<w:r><w:rPr><w:rFonts w:hint="eastAsia" /></w:rPr>'
    r'<w:t xml:space="preserve">参考文献</w:t></w:r></w:p>'
)

NEXT_HEAD = r'<w:p><w:pPr><w:pStyle w:val="1" /></w:pPr>'
TOC_TITLE_EN = '<w:t xml:space="preserve">Table of Contents</w:t>'
TOC_TITLE_CN = '<w:t xml:space="preserve">目录</w:t>'
ABSTRACT_HEAD_CN = (
    '<w:p><w:pPr><w:pStyle w:val="1" /></w:pPr><w:r><w:rPr><w:rFonts w:hint="eastAsia" /></w:rPr>'
    '<w:t xml:space="preserve">摘要</w:t></w:r></w:p>'
)
ABSTRACT_HEAD_CN_FIXED = (
    '<w:p><w:pPr><w:pStyle w:val="AbstractTitle" /></w:pPr><w:r><w:rPr><w:rFonts w:hint="eastAsia" /></w:rPr>'
    '<w:t xml:space="preserve">摘要</w:t></w:r></w:p>'
)
ABSTRACT_HEAD_EN = (
    '<w:p><w:pPr><w:pStyle w:val="1" /></w:pPr><w:r>'
    '<w:t xml:space="preserve">Abstract</w:t></w:r></w:p>'
)
ABSTRACT_HEAD_EN_FIXED = (
    '<w:p><w:pPr><w:pStyle w:val="AbstractTitle" /></w:pPr><w:r>'
    '<w:t xml:space="preserve">Abstract</w:t></w:r></w:p>'
)
ABSTRACT_END_MARK = '<w:bookmarkEnd w:id="21" />'
FIRST_CHAPTER_MARK = 'w:name="绪论"'
KEYWORD_PARA_RE = re.compile(
    r'(<w:p><w:pPr>)<w:pStyle w:val="a0" />(</w:pPr><w:r><w:rPr><w:rStyle w:val="(?:AbstractKeywordLabel|EnglishAbstractKeywordLabel)" />)',
)
PARAGRAPH_RE = re.compile(r'<w:p>.*?</w:p>', re.S)
PSTYLE_RE = re.compile(r'<w:pStyle w:val="[^"]+" />')
TABLE_RE = re.compile(r'<w:tbl>.*?</w:tbl>', re.S)


def _ensure_tbl_center(tbl: str) -> str:
    tblpr_match = re.search(r"<w:tblPr>.*?</w:tblPr>", tbl, re.S)
    if tblpr_match:
        tblpr = tblpr_match.group(0)
        if "<w:jc " in tblpr:
            new_tblpr = re.sub(r'<w:jc w:val="[^"]+" ?/?>', '<w:jc w:val="center"/>', tblpr, count=1)
        else:
            new_tblpr = tblpr.replace("<w:tblPr>", '<w:tblPr><w:jc w:val="center"/>', 1)
        tbl = tbl[: tblpr_match.start()] + new_tblpr + tbl[tblpr_match.end() :]
    return tbl


def _remove_table_borders(tbl: str) -> str:
    no_border = (
        '<w:tblBorders>'
        '<w:top w:val="nil"/><w:left w:val="nil"/><w:bottom w:val="nil"/><w:right w:val="nil"/>'
        '<w:insideH w:val="nil"/><w:insideV w:val="nil"/>'
        '</w:tblBorders>'
    )
    if "<w:tblBorders>" in tbl:
        tbl = re.sub(r"<w:tblBorders>.*?</w:tblBorders>", no_border, tbl, count=1, flags=re.S)
    else:
        tbl = tbl.replace("<w:tblPr>", f"<w:tblPr>{no_border}", 1)
    return tbl


def _set_table_layout(tbl: str, left_cell_width: int, right_cell_width: int) -> str:
    tbl = re.sub(r"<w:tblGrid>.*?</w:tblGrid>", "", tbl, count=1, flags=re.S)
    tbl = tbl.replace(
        "</w:tblPr>",
        '<w:tblW w:type="dxa" w:w="7200"/><w:tblLayout w:type="fixed"/></w:tblPr>',
        1,
    )
    tbl = re.sub(
        r"<w:tcPr\s*/>",
        lambda m: (
            f'<w:tcPr><w:tcW w:type="dxa" w:w="{left_cell_width if _set_table_layout._toggle % 2 == 0 else right_cell_width}"/></w:tcPr>'
        ),
        tbl,
    )
    return tbl


_set_table_layout._toggle = 0


def _replace_tcpr(match: re.Match[str], left_cell_width: int, right_cell_width: int) -> str:
    width = left_cell_width if _replace_tcpr.toggle % 2 == 0 else right_cell_width
    _replace_tcpr.toggle += 1
    return f'<w:tcPr><w:tcW w:type="dxa" w:w="{width}"/></w:tcPr>'


_replace_tcpr.toggle = 0


def rewrite_bibliography_segment(segment: str) -> str:
    segment = re.sub(
        r'<w:pStyle w:val="FirstParagraph" />',
        '<w:pStyle w:val="Bibliography" />',
        segment,
    )
    segment = re.sub(
        r'<w:pStyle w:val="a0" />',
        '<w:pStyle w:val="Bibliography" />',
        segment,
    )
    return segment


def normalize_toc_styles(styles_xml: str) -> str:
    toc_indents = {
        "TOC1": '<w:ind w:left="0" w:right="0" w:firstLine="0" w:hanging="0"/>',
        "TOC2": '<w:ind w:left="240" w:right="0" w:firstLine="0" w:hanging="0"/>',
        "TOC3": '<w:ind w:left="480" w:right="0" w:firstLine="0" w:hanging="0"/>',
    }

    def repl_toc(match: re.Match[str], style_id: str) -> str:
        style_xml = match.group(0)
        ppr_match = re.search(r"<w:pPr>.*?</w:pPr>", style_xml, re.S)
        if not ppr_match:
            return style_xml
        ppr = ppr_match.group(0)
        ppr = re.sub(r"<w:ind[^>]*/>", "", ppr)
        ppr = re.sub(r"<w:jc[^>]*/>", "", ppr)
        tabs_match = re.search(r"<w:tabs>.*?</w:tabs>", ppr, re.S)
        tabs = tabs_match.group(0) if tabs_match else ""
        new_ppr = (
            "<w:pPr>"
            f"{tabs}"
            '<w:spacing w:before="0" w:after="0" w:line="360" w:lineRule="auto"/>'
            f'{toc_indents[style_id]}'
            "</w:pPr>"
        )
        return style_xml[: ppr_match.start()] + new_ppr + style_xml[ppr_match.end() :]

    for style_id in ("TOC1", "TOC2", "TOC3"):
        styles_xml = re.sub(
            rf'<w:style[^>]*w:styleId="{style_id}".*?</w:style>',
            lambda m, sid=style_id: repl_toc(m, sid),
            styles_xml,
            flags=re.S,
        )

    styles_xml = re.sub(
        r'<w:style[^>]*w:styleId="TOCHeading".*?</w:style>',
        lambda m: re.sub(
            r"<w:pPr>.*?</w:pPr>",
            '<w:pPr><w:pageBreakBefore/><w:jc w:val="center"/><w:spacing w:before="0" w:after="0" w:line="360" w:lineRule="auto"/><w:ind w:left="0" w:right="0" w:firstLine="0"/></w:pPr>',
            m.group(0),
            count=1,
            flags=re.S,
        ),
        styles_xml,
        flags=re.S,
    )
    return styles_xml


def restyle_abstract_titles(doc_xml: str) -> str:
    doc_xml = doc_xml.replace(ABSTRACT_HEAD_CN, ABSTRACT_HEAD_CN_FIXED)
    doc_xml = doc_xml.replace(ABSTRACT_HEAD_EN, ABSTRACT_HEAD_EN_FIXED)
    doc_xml = re.sub(
        r'<w:p><w:pPr><w:pStyle w:val="AbstractTitle" /></w:pPr>',
        '<w:p><w:pPr><w:pStyle w:val="AbstractTitle" /><w:jc w:val="center"/><w:spacing w:before="0" w:after="0"/><w:ind w:left="0" w:right="0" w:firstLine="0"/></w:pPr>',
        doc_xml,
    )
    return doc_xml


def move_toc_after_abstracts(doc_xml: str) -> str:
    toc_start = doc_xml.find("<w:sdt>")
    if toc_start == -1:
        return doc_xml
    toc_end = doc_xml.find("</w:sdt>", toc_start)
    if toc_end == -1:
        return doc_xml
    toc_end += len("</w:sdt>")

    toc_block = doc_xml[toc_start:toc_end]
    if 'Table of Contents' not in toc_block and '目录' not in toc_block:
        return doc_xml
    doc_xml = doc_xml[:toc_start] + doc_xml[toc_end:]

    abstract_bookmarks = list(re.finditer(r'<w:bookmarkEnd w:id="\d+" />', doc_xml))
    abstract_titles = list(re.finditer(r'<w:p><w:pPr><w:pStyle w:val="AbstractTitle" />', doc_xml))
    chapter_heading = re.search(r'<w:p><w:pPr><w:pStyle w:val="1" /></w:pPr>', doc_xml)

    if len(abstract_titles) >= 2 and chapter_heading:
        insert_at = chapter_heading.start()
    elif abstract_bookmarks:
        insert_at = abstract_bookmarks[-1].end()
    elif chapter_heading:
        insert_at = chapter_heading.start()
    else:
        return toc_block + doc_xml

    return doc_xml[:insert_at] + toc_block + doc_xml[insert_at:]


def restyle_keyword_paragraphs(doc_xml: str) -> str:
    return KEYWORD_PARA_RE.sub(r'\1<w:pStyle w:val="AbstractKeywordParagraph" />\2', doc_xml)


def restyle_table_caption_paragraphs(doc_xml: str) -> str:
    def repl(match: re.Match[str]) -> str:
        paragraph = match.group(0)
        texts = re.findall(r'<w:t[^>]*>(.*?)</w:t>', paragraph)
        paragraph_text = "".join(texts)
        if not paragraph_text.startswith("表"):
            return paragraph
        if "表格内容在" not in paragraph_text or "导出中省略" not in paragraph_text:
            return paragraph
        if "<w:pPr>" in paragraph:
            paragraph = PSTYLE_RE.sub("", paragraph, count=1)
            paragraph = paragraph.replace("<w:pPr>", '<w:pPr><w:pStyle w:val="TableCaption" />', 1)
        else:
            paragraph = paragraph.replace("<w:p>", '<w:p><w:pPr><w:pStyle w:val="TableCaption" /></w:pPr>', 1)
        paragraph = re.sub(
            r"<w:pPr>.*?</w:pPr>",
            '<w:pPr><w:pStyle w:val="TableCaption" /><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/><w:jc w:val="center"/><w:ind w:left="0" w:right="0" w:firstLine="0"/></w:pPr>',
            paragraph,
            count=1,
            flags=re.S,
        )
        return paragraph

    return PARAGRAPH_RE.sub(repl, doc_xml)


def normalize_table_caption_spacing(doc_xml: str) -> str:
    def repl(match: re.Match[str]) -> str:
        paragraph = match.group(0)
        if '<w:pStyle w:val="TableCaption"' not in paragraph:
            return paragraph
        if "<w:pPr>" in paragraph:
            paragraph = re.sub(
                r"<w:pPr>.*?</w:pPr>",
                '<w:pPr><w:pStyle w:val="TableCaption" /><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/><w:jc w:val="center"/><w:ind w:left="0" w:right="0" w:firstLine="0"/></w:pPr>',
                paragraph,
                count=1,
                flags=re.S,
            )
        return paragraph

    return PARAGRAPH_RE.sub(repl, doc_xml)


def center_image_paragraphs(doc_xml: str) -> str:
    def repl(match: re.Match[str]) -> str:
        paragraph = match.group(0)
        if "<w:drawing>" not in paragraph:
            return paragraph
        if "<w:pPr>" in paragraph:
            if "<w:jc " in paragraph:
                paragraph = re.sub(r'<w:jc w:val="[^"]+" ?/?>', '<w:jc w:val="center"/>', paragraph, count=1)
            else:
                paragraph = paragraph.replace("<w:pPr>", '<w:pPr><w:jc w:val="center"/>', 1)
            if "<w:ind " in paragraph:
                paragraph = re.sub(r'<w:ind [^>]*/>', '<w:ind w:left="0" w:right="0" w:firstLine="0"/>', paragraph, count=1)
            else:
                paragraph = paragraph.replace("<w:pPr>", '<w:pPr><w:ind w:left="0" w:right="0" w:firstLine="0"/>', 1)
        else:
            paragraph = paragraph.replace(
                "<w:p>",
                '<w:p><w:pPr><w:jc w:val="center"/><w:ind w:left="0" w:right="0" w:firstLine="0"/></w:pPr>',
                1,
            )
        return paragraph

    return PARAGRAPH_RE.sub(repl, doc_xml)


def restyle_cover(doc_xml: str) -> str:
    abstract_idx = doc_xml.find('<w:bookmarkStart w:id="26" w:name="摘要"')
    if abstract_idx == -1:
        abstract_idx = doc_xml.find('>摘要</w:t>')
    if abstract_idx == -1:
        return doc_xml

    cover = doc_xml[:abstract_idx]
    tail = doc_xml[abstract_idx:]

    cover = re.sub(
        r"<w:p>\s*<w:pPr>\s*<w:pStyle w:val=\"Title\" />.*?</w:p>\s*<w:p>\s*<w:pPr>\s*<w:pStyle w:val=\"Author\" />.*?</w:p>",
        "",
        cover,
        count=1,
        flags=re.S,
    )

    def repl_para(match: re.Match[str]) -> str:
        paragraph = match.group(0)
        text = "".join(re.findall(r'<w:t[^>]*>(.*?)</w:t>', paragraph))
        if "本科毕业设计（论文）" in text:
            return (
                '<w:p><w:pPr><w:jc w:val="center"/><w:spacing w:before="240" w:after="240"/>'
                '<w:ind w:left="0" w:right="0" w:firstLine="0"/></w:pPr>'
                '<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:eastAsia="黑体" w:hAnsi="Times New Roman" w:cs="黑体"/>'
                '<w:b/><w:bCs/><w:sz w:val="52"/><w:szCs w:val="52"/></w:rPr>'
                '<w:t xml:space="preserve">本科毕业设计（论文）</w:t></w:r></w:p>'
            )
        if "202X年" in text or re.search(r"\d{4}年.*月", text):
            paragraph = re.sub(r"<w:pPr>.*?</w:pPr>", '<w:pPr><w:jc w:val="center"/><w:spacing w:before="120" w:after="0"/><w:ind w:left="0" w:right="0" w:firstLine="0"/></w:pPr>', paragraph, count=1, flags=re.S)
            paragraph = re.sub(r"<w:rPr>.*?</w:rPr>", '<w:rPr><w:rFonts w:ascii="Times New Roman" w:eastAsia="黑体" w:hAnsi="Times New Roman" w:cs="黑体"/><w:b/><w:bCs/><w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr>', paragraph, count=1, flags=re.S)
            return paragraph
        if text.strip() and "<w:drawing>" not in paragraph and "摘要" not in text:
            if "<w:pPr>" in paragraph:
                paragraph = re.sub(r"<w:pPr>.*?</w:pPr>", '<w:pPr><w:jc w:val="center"/><w:spacing w:before="180" w:after="260"/><w:ind w:left="0" w:right="0" w:firstLine="0"/></w:pPr>', paragraph, count=1, flags=re.S)
                paragraph = re.sub(r"<w:rPr>.*?</w:rPr>", '<w:rPr><w:rFonts w:ascii="Times New Roman" w:eastAsia="黑体" w:hAnsi="Times New Roman" w:cs="黑体"/><w:b/><w:bCs/><w:sz w:val="44"/><w:szCs w:val="44"/></w:rPr>', paragraph, count=1, flags=re.S)
            return paragraph
        return paragraph

    cover = PARAGRAPH_RE.sub(repl_para, cover)

    tables = list(TABLE_RE.finditer(cover))
    if tables:
        first = tables[0]
        tbl = first.group(0)
        tbl = _ensure_tbl_center(tbl)
        tbl = _remove_table_borders(tbl)
        tbl = re.sub(r"<w:tcPr\s*/>", lambda m: _replace_tcpr(m, 1700, 5500), tbl)
        _replace_tcpr.toggle = 0
        cell_index = {"value": 0}

        def repl_cover_paragraph(p_match: re.Match[str]) -> str:
            paragraph = p_match.group(0)
            is_label = cell_index["value"] % 2 == 0
            align = "right" if is_label else "left"
            cell_index["value"] += 1
            if "<w:pPr>" in paragraph:
                paragraph = re.sub(
                    r"<w:pPr>.*?</w:pPr>",
                    f'<w:pPr><w:spacing w:before="0" w:after="160"/><w:jc w:val="{align}"/><w:ind w:left="0" w:right="0" w:firstLine="0"/></w:pPr>',
                    paragraph,
                    count=1,
                    flags=re.S,
                )
            if is_label:
                if "<w:rPr>" in paragraph:
                    paragraph = re.sub(
                        r"<w:rPr>.*?</w:rPr>",
                        '<w:rPr><w:rFonts w:ascii="Times New Roman" w:eastAsia="黑体" w:hAnsi="Times New Roman" w:cs="黑体"/><w:b/><w:bCs/><w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr>',
                        paragraph,
                        count=1,
                        flags=re.S,
                    )
            else:
                if "<w:rPr>" in paragraph:
                    paragraph = re.sub(
                        r"<w:rPr>.*?</w:rPr>",
                        '<w:rPr><w:rFonts w:ascii="Times New Roman" w:eastAsia="宋体" w:hAnsi="Times New Roman" w:cs="宋体"/><w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr>',
                        paragraph,
                        count=1,
                        flags=re.S,
                    )
                else:
                    paragraph = paragraph.replace(
                        "<w:r>",
                        '<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:eastAsia="宋体" w:hAnsi="Times New Roman" w:cs="宋体"/><w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr>',
                        1,
                    )
            return paragraph

        tbl = PARAGRAPH_RE.sub(repl_cover_paragraph, tbl)
        cover = cover[: first.start()] + tbl + cover[first.end() :]

    return cover + tail


def center_body_tables(doc_xml: str) -> str:
    tables = list(TABLE_RE.finditer(doc_xml))
    if len(tables) <= 1:
        return doc_xml

    def style_body_table(tbl: str) -> str:
        borders = (
            '<w:tblBorders>'
            '<w:top w:val="single" w:sz="8" w:space="0" w:color="000000"/>'
            '<w:left w:val="nil"/>'
            '<w:bottom w:val="single" w:sz="8" w:space="0" w:color="000000"/>'
            '<w:right w:val="nil"/>'
            '<w:insideH w:val="nil"/>'
            '<w:insideV w:val="nil"/>'
            '</w:tblBorders>'
        )
        tbl = _ensure_tbl_center(tbl)
        tblpr_match = re.search(r"<w:tblPr>.*?</w:tblPr>", tbl, re.S)
        if tblpr_match:
            tblpr = tblpr_match.group(0)
            if "<w:tblBorders>" in tblpr:
                tblpr = re.sub(r"<w:tblBorders>.*?</w:tblBorders>", borders, tblpr, count=1, flags=re.S)
            else:
                tblpr = tblpr.replace("</w:tblPr>", borders + "</w:tblPr>", 1)
            if "<w:tblCellMar>" in tblpr:
                tblpr = re.sub(
                    r"<w:tblCellMar>.*?</w:tblCellMar>",
                    '<w:tblCellMar><w:top w:w="80" w:type="dxa"/><w:left w:w="120" w:type="dxa"/><w:bottom w:w="80" w:type="dxa"/><w:right w:w="120" w:type="dxa"/></w:tblCellMar>',
                    tblpr,
                    count=1,
                    flags=re.S,
                )
            else:
                tblpr = tblpr.replace(
                    "</w:tblPr>",
                    '<w:tblCellMar><w:top w:w="80" w:type="dxa"/><w:left w:w="300" w:type="dxa"/><w:bottom w:w="80" w:type="dxa"/><w:right w:w="300" w:type="dxa"/></w:tblCellMar></w:tblPr>',
                    1,
                )
            tbl = tbl[: tblpr_match.start()] + tblpr + tbl[tblpr_match.end() :]

        tbl = re.sub(
            r"<w:tcPr\s*/>",
            '<w:tcPr><w:vAlign w:val="center"/></w:tcPr>',
            tbl,
        )
        tbl = re.sub(
            r"<w:tcPr>(.*?)</w:tcPr>",
            lambda m: m.group(0) if "<w:vAlign" in m.group(1) else f'<w:tcPr>{m.group(1)}<w:vAlign w:val="center"/></w:tcPr>',
            tbl,
            flags=re.S,
        )

        rows = list(re.finditer(r"<w:tr>.*?</w:tr>", tbl, re.S))
        if not rows:
            return tbl

        styled_rows = []
        for row_index, row_match in enumerate(rows):
            row = row_match.group(0)
            if row_index == 0:
                if "<w:trPr>" in row:
                    if "<w:tblBorders>" not in row:
                        row = row.replace(
                            "</w:trPr>",
                            '<w:tblBorders><w:bottom w:val="single" w:sz="6" w:space="0" w:color="000000"/></w:tblBorders></w:trPr>',
                            1,
                        )
                else:
                    row = row.replace(
                        "<w:tr>",
                        '<w:tr><w:trPr><w:tblBorders><w:bottom w:val="single" w:sz="6" w:space="0" w:color="000000"/></w:tblBorders></w:trPr>',
                        1,
                    )

            def repl_para(p_match: re.Match[str]) -> str:
                paragraph = p_match.group(0)
                if "<w:pPr>" in paragraph:
                    paragraph = re.sub(
                        r"<w:pPr>.*?</w:pPr>",
                        '<w:pPr><w:pStyle w:val="Compact" /><w:jc w:val="center"/><w:spacing w:before="40" w:after="40" w:line="300" w:lineRule="exact"/><w:ind w:left="0" w:right="0" w:firstLine="0"/></w:pPr>',
                        paragraph,
                        count=1,
                        flags=re.S,
                    )
                else:
                    paragraph = paragraph.replace(
                        "<w:p>",
                        '<w:p><w:pPr><w:pStyle w:val="Compact" /><w:jc w:val="center"/><w:spacing w:before="40" w:after="40" w:line="300" w:lineRule="exact"/><w:ind w:left="0" w:right="0" w:firstLine="0"/></w:pPr>',
                        1,
                    )

                font_xml = (
                    '<w:rPr><w:rFonts w:ascii="Times New Roman" w:eastAsia="宋体" w:hAnsi="Times New Roman" w:cs="宋体"/>'
                    '<w:sz w:val="21"/><w:szCs w:val="21"/></w:rPr>'
                )

                if "<w:rPr>" in paragraph:
                    paragraph = re.sub(r"<w:rPr>.*?</w:rPr>", font_xml, paragraph, flags=re.S)
                else:
                    paragraph = paragraph.replace("<w:r>", f"<w:r>{font_xml}", 1)
                return paragraph

            styled_rows.append(PARAGRAPH_RE.sub(repl_para, row))

        new_tbl = tbl
        for old_row, new_row in zip(rows, styled_rows):
            new_tbl = new_tbl.replace(old_row.group(0), new_row, 1)
        return new_tbl

    offset = 0
    for idx, match in enumerate(tables, start=1):
        if idx == 1:
            continue
        new_tbl = style_body_table(match.group(0))
        new_tbl += '<w:p><w:pPr><w:spacing w:before="0" w:after="240" w:line="240" w:lineRule="exact"/><w:ind w:left="0" w:right="0" w:firstLine="0"/></w:pPr></w:p>'
        start = match.start() + offset
        end = match.end() + offset
        doc_xml = doc_xml[:start] + new_tbl + doc_xml[end:]
        offset += len(new_tbl) - (match.end() - match.start())
    return doc_xml


def fix_docx(path: Path) -> None:
    with zipfile.ZipFile(path, "r") as zin:
        doc_xml = zin.read("word/document.xml").decode("utf-8")
        settings_xml = zin.read("word/settings.xml").decode("utf-8")
        styles_xml = zin.read("word/styles.xml").decode("utf-8")

        doc_xml = doc_xml.replace(TOC_TITLE_EN, TOC_TITLE_CN)
        doc_xml = restyle_abstract_titles(doc_xml)
        doc_xml = move_toc_after_abstracts(doc_xml)
        doc_xml = restyle_keyword_paragraphs(doc_xml)
        doc_xml = restyle_table_caption_paragraphs(doc_xml)
        doc_xml = normalize_table_caption_spacing(doc_xml)
        doc_xml = center_image_paragraphs(doc_xml)
        doc_xml = restyle_cover(doc_xml)
        doc_xml = center_body_tables(doc_xml)
        styles_xml = normalize_toc_styles(styles_xml)

        m = re.search(REF_HEAD, doc_xml)
        if not m:
            fixed_doc_xml = doc_xml
        else:
            start = m.end()
            tail = doc_xml[start:]
            next_head = re.search(NEXT_HEAD, tail)
            if not next_head:
                fixed_doc_xml = doc_xml
            else:
                bib_segment = tail[: next_head.start()]
                fixed_segment = rewrite_bibliography_segment(bib_segment)
                fixed_doc_xml = doc_xml[:start] + fixed_segment + tail[next_head.start() :]

        with NamedTemporaryFile(delete=False, suffix=".docx", dir=path.parent) as tmp:
            tmp_path = Path(tmp.name)

        with zipfile.ZipFile(tmp_path, "w") as zout:
            for info in zin.infolist():
                data = zin.read(info.filename)
                if info.filename == "word/document.xml":
                    data = fixed_doc_xml.encode("utf-8")
                if info.filename == "word/settings.xml":
                    if "<w:updateFields" not in settings_xml:
                        settings_xml_fixed = settings_xml.replace(
                            "</w:settings>",
                            '<w:updateFields w:val="true" /></w:settings>',
                        )
                    else:
                        settings_xml_fixed = re.sub(
                            r'<w:updateFields[^>]*/>',
                            '<w:updateFields w:val="true" />',
                            settings_xml,
                        )
                    data = settings_xml_fixed.encode("utf-8")
                if info.filename == "word/styles.xml":
                    data = styles_xml.encode("utf-8")
                zout.writestr(info, data)

    shutil.move(tmp_path, path)


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser()
    parser.add_argument("docx", type=Path)
    args = parser.parse_args()
    fix_docx(args.docx.resolve())
