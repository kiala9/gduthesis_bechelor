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


def restyle_abstract_titles(doc_xml: str) -> str:
    doc_xml = doc_xml.replace(ABSTRACT_HEAD_CN, ABSTRACT_HEAD_CN_FIXED)
    doc_xml = doc_xml.replace(ABSTRACT_HEAD_EN, ABSTRACT_HEAD_EN_FIXED)
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

    abstract_end = doc_xml.find(ABSTRACT_END_MARK)
    if abstract_end != -1:
        insert_at = abstract_end + len(ABSTRACT_END_MARK)
    else:
        first_chapter = doc_xml.find(FIRST_CHAPTER_MARK)
        if first_chapter == -1:
            return doc_xml
        insert_at = max(0, doc_xml.rfind("<w:bookmarkStart", 0, first_chapter))

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


def fix_docx(path: Path) -> None:
    with zipfile.ZipFile(path, "r") as zin:
        doc_xml = zin.read("word/document.xml").decode("utf-8")
        settings_xml = zin.read("word/settings.xml").decode("utf-8")

        doc_xml = doc_xml.replace(TOC_TITLE_EN, TOC_TITLE_CN)
        doc_xml = restyle_abstract_titles(doc_xml)
        doc_xml = move_toc_after_abstracts(doc_xml)
        doc_xml = restyle_keyword_paragraphs(doc_xml)
        doc_xml = restyle_table_caption_paragraphs(doc_xml)
        doc_xml = center_image_paragraphs(doc_xml)

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
                zout.writestr(info, data)

    shutil.move(tmp_path, path)


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser()
    parser.add_argument("docx", type=Path)
    args = parser.parse_args()
    fix_docx(args.docx.resolve())
