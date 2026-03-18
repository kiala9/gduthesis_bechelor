from __future__ import annotations

import argparse
import re
from pathlib import Path


MAIN_TEMPLATE = r"""\documentclass[12pt]{{book}}
\usepackage[UTF8]{{ctex}}
\usepackage{{graphicx}}
\usepackage{{booktabs}}
\usepackage{{longtable}}
\usepackage{{array}}
\usepackage{{tabularx}}
\usepackage{{multirow}}
\usepackage{{amsmath,amssymb,bm}}

\title{{{title}}}
\author{{{author}}}
\date{{}}

\begin{{document}}

{body}

\end{{document}}
"""


def extract_macro(text: str, name: str) -> str:
    m = re.search(rf"\\{name}\s*\{{([^{{}}]*)\}}", text)
    return m.group(1).strip() if m else ""


def build_cover(main_tex: str, thesis_dir: Path) -> str:
    fields = {
        "title": extract_macro(main_tex, "ThesisTitle"),
        "college": extract_macro(main_tex, "ThesisCollege"),
        "major": extract_macro(main_tex, "ThesisMajor"),
        "year": extract_macro(main_tex, "ThesisYear"),
        "class": extract_macro(main_tex, "ThesisClass"),
        "student_id": extract_macro(main_tex, "ThesisStudentID"),
        "author": extract_macro(main_tex, "ThesisAuthor"),
        "supervisor": extract_macro(main_tex, "ThesisSupervisor"),
        "date": extract_macro(main_tex, "ThesisDate"),
    }

    logo = ""
    if (thesis_dir / "figures" / "xiaohui.jpg").exists():
        logo = (
            r"\begin{flushleft}"
            "\n"
            r"\includegraphics[width=2.2cm]{figures/xiaohui.jpg}"
            "\n"
            r"\end{flushleft}"
        )

    school_name = ""
    if (thesis_dir / "figures" / "mingchen.jpg").exists():
        school_name = (
            r"\begin{center}"
            "\n"
            r"\includegraphics[width=10.56cm]{figures/mingchen.jpg}"
            "\n"
            r"\end{center}"
        )

    info_lines = [
        (r"\textbf{学\quad 院：}", fields["college"]),
        (r"\textbf{专\quad 业：}", fields["major"]),
        (r"\textbf{年级班别：}", f"{fields['year']}级（{fields['class']}）班"),
        (r"\textbf{学\quad 号：}", fields["student_id"]),
        (r"\textbf{学生姓名：}", fields["author"]),
        (r"\textbf{指导教师：}", fields["supervisor"]),
    ]

    parts = [
        r"\begin{titlepage}",
        logo,
        school_name,
        r"\vspace*{0.5cm}",
        r"\begin{center}",
        r"{\zihao{1}\heiti\bfseries 本科毕业设计（论文）\par}",
        r"\vspace{1.2cm}",
        rf"{{\zihao{{2}}\heiti\bfseries {fields['title']}\par}}",
        r"\vspace{2cm}",
        r"\begin{tabular}{@{}r l@{}}",
        r" \\[0.8em] ".join(f"{label} & {value}" for label, value in info_lines) + r"\\",
        r"\end{tabular}",
        r"\vfill",
        rf"{{\zihao{{3}}\heiti\bfseries {fields['date']}\par}}",
        r"\end{center}",
        r"\end{titlepage}",
    ]
    return "\n".join(part for part in parts if part)


def parse_label_map(thesis_dir: Path) -> dict[str, str]:
    label_map: dict[str, str] = {}
    for aux_path in sorted(thesis_dir.glob("**/*.aux")):
        text = aux_path.read_text(encoding="utf-8", errors="ignore")
        for label, number in re.findall(r"\\newlabel\{([^}]*)\}\{\{([^}]*)\}\{", text):
            label_map.setdefault(label.strip(), number.strip())
    return label_map


def parse_bib_map(thesis_dir: Path) -> dict[str, int]:
    bbl_path = thesis_dir / "main.bbl"
    if not bbl_path.exists():
        return {}

    bbl = bbl_path.read_text(encoding="utf-8", errors="ignore")
    bib_map: dict[str, int] = {}
    for idx, key in enumerate(re.findall(r"\\bibitem(?:\[[^\]]*\])?\{([^}]*)\}", bbl), start=1):
        bib_map[key.strip()] = idx
    return bib_map


def resolve_ref(label: str, label_map: dict[str, str], fallback: str = "??") -> str:
    return label_map.get(label.strip(), fallback)


def format_cite(keys: str, bib_map: dict[str, int]) -> str:
    refs = []
    for key in keys.split(","):
        number = bib_map.get(key.strip())
        refs.append(f"[{number if number is not None else '?'}]")
    return "".join(refs)


def replace_refs_and_cites(text: str, label_map: dict[str, str], bib_map: dict[str, int]) -> str:
    def repl_figref(match: re.Match[str]) -> str:
        return f"图{resolve_ref(match.group(1), label_map)}"

    def repl_tabref(match: re.Match[str]) -> str:
        return f"表{resolve_ref(match.group(1), label_map)}"

    def repl_secref(match: re.Match[str]) -> str:
        return f"第{resolve_ref(match.group(1), label_map)}节"

    def repl_eqref(match: re.Match[str]) -> str:
        return f"({resolve_ref(match.group(1), label_map)})"

    def repl_ref(match: re.Match[str]) -> str:
        return resolve_ref(match.group(1), label_map)

    def repl_upcite(match: re.Match[str]) -> str:
        return rf"\textsuperscript{{{format_cite(match.group(1), bib_map)}}}"

    def repl_cite(match: re.Match[str]) -> str:
        return format_cite(match.group(1), bib_map)

    patterns = [
        (r"\\figref\{([^}]*)\}", repl_figref),
        (r"\\tabref\{([^}]*)\}", repl_tabref),
        (r"\\secref\{([^}]*)\}", repl_secref),
        (r"\\eqref\{([^}]*)\}", repl_eqref),
        (r"\\upcite\{([^}]*)\}", repl_upcite),
        (r"\\inlinecite\{([^}]*)\}", repl_cite),
        (r"\\cite\{([^}]*)\}", repl_cite),
        (r"\\autoref\{([^}]*)\}", repl_ref),
        (r"\\ref\{([^}]*)\}", repl_ref),
    ]
    for pattern, repl in patterns:
        text = re.sub(pattern, repl, text)
    return text


def simplify_column_spec(spec: str) -> str:
    spec = re.sub(r">\{[^{}]*\}", "", spec)
    spec = re.sub(r"<\{[^{}]*\}", "", spec)
    spec = re.sub(r"p\{[^{}]*\}", "l", spec)
    spec = re.sub(r"m\{[^{}]*\}", "l", spec)
    spec = re.sub(r"b\{[^{}]*\}", "l", spec)
    spec = spec.replace("X", "l")
    spec = spec.replace("@{}", "")
    spec = spec.replace("!", "")
    spec = spec.replace(">", "")
    spec = spec.replace("<", "")
    spec = "".join(ch for ch in spec if ch in "lcr|")
    return spec or "lll"


def read_braced(text: str, start: int) -> tuple[str, int] | None:
    if start >= len(text) or text[start] != "{":
        return None
    depth = 0
    for idx in range(start, len(text)):
        ch = text[idx]
        if ch == "{":
            depth += 1
        elif ch == "}":
            depth -= 1
            if depth == 0:
                return text[start + 1 : idx], idx + 1
    return None


def extract_tabular_block(block: str) -> tuple[str, str, str] | None:
    for env_name in ("tabularx", "tabular", "longtable"):
        token = rf"\begin{{{env_name}}}"
        start = block.find(token)
        if start == -1:
            continue

        pos = start + len(token)
        if env_name == "tabularx":
            width_arg = read_braced(block, pos)
            if not width_arg:
                return None
            _, pos = width_arg

        spec_arg = read_braced(block, pos)
        if not spec_arg:
            return None
        column_spec, pos = spec_arg

        end_token = rf"\end{{{env_name}}}"
        end = block.find(end_token, pos)
        if end == -1:
            return None

        content = block[pos:end]
        full_block = block[start : end + len(end_token)]
        return full_block, column_spec, content
    return None


def fallback_table(block: str, label_map: dict[str, str], bib_map: dict[str, int]) -> str:
    caption_match = re.search(r"\\caption\{(.*?)\}", block, re.S)
    label_match = re.search(r"\\label\{([^}]*)\}", block)
    caption = caption_match.group(1).strip() if caption_match else "表格"
    caption = replace_refs_and_cites(caption, label_map, bib_map)
    caption = latex_to_plain(caption)
    label = label_match.group(1).strip() if label_match else ""
    number = label_map.get(label, "")
    prefix = f"表{number} " if number else "表 "
    return f"\\begin{{center}}{prefix}{caption}（表格内容在 Word 导出中省略）\\end{{center}}"


def normalize_tables(text: str, label_map: dict[str, str], bib_map: dict[str, int]) -> str:
    def repl_table(match: re.Match[str]) -> str:
        block = match.group(0)
        caption_match = re.search(r"\\caption\{(.*?)\}", block, re.S)
        caption = caption_match.group(1).strip() if caption_match else ""
        if any(token in block for token in (r"\multirow", r"\multicolumn", r"\diagbox")):
            return fallback_table(block, label_map, bib_map)

        tabular_block = extract_tabular_block(block)
        if not tabular_block:
            return fallback_table(block, label_map, bib_map)
        raw_tabular, column_spec, content = tabular_block
        column_spec = simplify_column_spec(column_spec)
        content = content.replace(r"\toprule", r"\hline")
        content = content.replace(r"\midrule", r"\hline")
        content = content.replace(r"\bottomrule", r"\hline")
        content = re.sub(r"\\renewcommand\{\\arraystretch\}\{[^}]*\}", "", content)
        content = re.sub(r"\\arraybackslash", "", content)
        content = re.sub(r"\\centering", "", content)
        content = re.sub(r">\{[^{}]*\}", "", content)
        content = re.sub(r"<\{[^{}]*\}", "", content)
        content = content.strip()

        if any(token in content for token in (r"\resizebox", r"\rotatebox", r"\parbox")):
            return fallback_table(block, label_map, bib_map)

        if caption:
            caption = replace_refs_and_cites(caption, label_map, bib_map)

        parts = [r"\begin{table}[htbp]"]
        if caption:
            parts.append(rf"\caption{{{caption}}}")
        parts.append(rf"\begin{{tabular}}{{{column_spec}}}")
        parts.append(content)
        parts.append(r"\end{tabular}")
        parts.append(r"\end{table}")
        return "\n".join(parts)

    table_patterns = [
        r"\\begin\{table\}(?:\[[^\]]*\])?.*?\\end\{table\}",
        r"\\begin\{longtable\}(?:\[[^\]]*\])?.*?\\end\{longtable\}",
    ]
    for pattern in table_patterns:
        text = re.sub(pattern, repl_table, text, flags=re.S)
    return text


def resolve_graphics_paths(text: str, thesis_dir: Path) -> str:
    def repl(match: re.Match[str]) -> str:
        options = match.group(1) or ""
        path = match.group(2).strip()
        if "." in Path(path).name:
            return match.group(0)

        candidates = [thesis_dir / path, thesis_dir / "figures" / path]
        for base in candidates:
            for suffix in (".png", ".pdf", ".jpg", ".jpeg", ".svg"):
                candidate = base.with_suffix(suffix)
                if candidate.exists():
                    relative = candidate.relative_to(thesis_dir)
                    return rf"\includegraphics{options}{{{relative.as_posix()}}}"
        return match.group(0)

    return re.sub(r"\\includegraphics(\[[^\]]*\])?\{([^}]*)\}", repl, text)


def clean_common(text: str, thesis_dir: Path, label_map: dict[str, str], bib_map: dict[str, int]) -> str:
    text = replace_refs_and_cites(text, label_map, bib_map)
    text = normalize_tables(text, label_map, bib_map)
    text = resolve_graphics_paths(text, thesis_dir)
    replacements = {
        r"\enspace": " ",
        r"\quad": " ",
        r"\qquad": " ",
    }
    for src, dst in replacements.items():
        text = text.replace(src, dst)
    text = re.sub(r"\\clearpage", "\n", text)
    text = re.sub(r"\\unnumberedchapter\{([^}]*)\}", r"\\chapter*{\1}", text)
    text = re.sub(r"\\pagenumbering\{[^}]*\}", "", text)
    text = re.sub(r"\\setcounter\{page\}\{[^}]*\}", "", text)
    text = re.sub(r"\\thispagestyle\{[^}]*\}", "", text)
    text = re.sub(r"\\phantomsection", "", text)
    text = re.sub(r"\\label\{[^}]*\}", "", text)
    text = re.sub(r"\\addcontentsline\{[^}]*\}\{[^}]*\}\{[^}]*\}", "", text)
    return text


def process_abstract(
    tex: str,
    lang: str,
    thesis_dir: Path,
    label_map: dict[str, str],
    bib_map: dict[str, int],
) -> str:
    if lang == "cn":
        heading = r"\chapter*{摘要}"
        keyword_cmd = "cnkeywords"
        keyword_label = r"\textbf{关键词：}"
    else:
        heading = r"\chapter*{Abstract}"
        keyword_cmd = "enkeywords"
        keyword_label = r"\textbf{Key words:}"

    tex = re.sub(rf"\\begin\{{{'cnabstract' if lang == 'cn' else 'enabstract'}\}}", "", tex)
    tex = re.sub(rf"\\end\{{{'cnabstract' if lang == 'cn' else 'enabstract'}\}}", "", tex)

    m = re.search(rf"\\{keyword_cmd}\{{(.*?)\}}", tex, re.S)
    keywords = m.group(1).strip() if m else ""
    tex = re.sub(rf"\\{keyword_cmd}\{{.*?\}}", "", tex, flags=re.S)
    tex = clean_common(tex, thesis_dir, label_map, bib_map).strip()

    parts = [heading, tex]
    if keywords:
        parts.append(f"{keyword_label} {keywords}")
    return "\n\n".join(part for part in parts if part)


def latex_to_plain(text: str) -> str:
    text = text.replace("\n", " ")
    text = text.replace("~", " ")
    text = text.replace(r"\allowbreak", "")
    text = text.replace(r"\newblock", " ")
    text = text.replace(r"\&", "&")
    text = re.sub(r"\\url\{([^}]*)\}", r"\1", text)
    text = re.sub(r"\\nolinkurl\{([^}]*)\}", r"\1", text)
    text = re.sub(r"\\href\{[^}]*\}\{([^}]*)\}", r"\1", text)
    text = re.sub(r"\\[A-Za-z]+\*?(?:\[[^]]*\])?\{([^{}]*)\}", r"\1", text)
    text = re.sub(r"[{}]", "", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def build_bibliography_section(thesis_dir: Path) -> str:
    bbl_path = thesis_dir / "main.bbl"
    if not bbl_path.exists():
        return r"\chapter*{参考文献}" + "\n\n" + "参考文献待生成。"

    bbl = bbl_path.read_text(encoding="utf-8")
    entries = re.split(r"\\bibitem(?:\[[^\]]*\])?\{[^}]*\}", bbl)[1:]
    lines = [r"\chapter*{参考文献}"]
    for idx, entry in enumerate(entries, start=1):
        plain = latex_to_plain(entry)
        if plain:
            lines.append(f"[{idx}] {plain}")
    return "\n\n".join(lines)


def process_regular(tex: str, thesis_dir: Path, label_map: dict[str, str], bib_map: dict[str, int]) -> str:
    tex = clean_common(tex, thesis_dir, label_map, bib_map)
    return tex.strip()


def process_appendix(
    tex: str,
    appendix_index: int,
    thesis_dir: Path,
    label_map: dict[str, str],
    bib_map: dict[str, int],
) -> str:
    appendix_letter = chr(ord("A") + appendix_index)
    tex = re.sub(
        r"\\chapter\*?\{([^}]*)\}",
        rf"\\chapter*{{附录{appendix_letter} \1}}",
        tex,
        count=1,
    )
    tex = clean_common(tex, thesis_dir, label_map, bib_map)
    return tex.strip()


def load_include(
    thesis_dir: Path,
    include_name: str,
    label_map: dict[str, str],
    bib_map: dict[str, int],
    in_appendix: bool = False,
    appendix_index: int = 0,
) -> str:
    path = thesis_dir / f"{include_name}.tex"
    tex = path.read_text(encoding="utf-8")
    if path.name == "abstract_cn.tex":
        return process_abstract(tex, "cn", thesis_dir, label_map, bib_map)
    if path.name == "abstract_en.tex":
        return process_abstract(tex, "en", thesis_dir, label_map, bib_map)
    if in_appendix:
        return process_appendix(tex, appendix_index, thesis_dir, label_map, bib_map)
    return process_regular(tex, thesis_dir, label_map, bib_map)


def build_body(main_tex: str, thesis_dir: Path) -> str:
    body_match = re.search(r"\\begin\{document\}(.*)\\end\{document\}", main_tex, re.S)
    if not body_match:
        raise ValueError("Cannot locate document body in main.tex")
    body = body_match.group(1)

    label_map = parse_label_map(thesis_dir)
    bib_map = parse_bib_map(thesis_dir)
    in_appendix = False
    appendix_index = 0

    parts: list[str] = []
    for line in body.splitlines():
        stripped = line.strip()
        if not stripped or stripped.startswith("%"):
            continue
        if stripped.startswith(r"\makecover"):
            parts.append(build_cover(main_tex, thesis_dir))
            continue
        if stripped.startswith(r"\tableofcontents"):
            continue
        if stripped.startswith(r"\include{"):
            include_name = re.search(r"\\include\{([^}]*)\}", stripped).group(1)
            parts.append(
                load_include(
                    thesis_dir,
                    include_name,
                    label_map,
                    bib_map,
                    in_appendix=in_appendix,
                    appendix_index=appendix_index,
                )
            )
            if in_appendix:
                appendix_index += 1
            continue
        if stripped.startswith(r"\printreferences"):
            parts.append(build_bibliography_section(thesis_dir))
            continue
        if stripped.startswith(r"\startappendix"):
            in_appendix = True
            continue
        if stripped.startswith((r"\pagenumbering", r"\setcounter", r"\clearpage")):
            continue
        parts.append(stripped)

    return "\n\n".join(part for part in parts if part).strip()


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("thesis_dir", type=Path)
    parser.add_argument("--output", type=Path)
    args = parser.parse_args()

    thesis_dir = args.thesis_dir.resolve()
    main_path = thesis_dir / "main.tex"
    output_path = args.output.resolve() if args.output else thesis_dir / "pandoc_export.tex"

    main_tex = main_path.read_text(encoding="utf-8")
    title = ""
    author = ""
    body = build_body(main_tex, thesis_dir)

    output = MAIN_TEMPLATE.format(title=title, author=author, body=body)
    output_path.write_text(output, encoding="utf-8")


if __name__ == "__main__":
    main()
