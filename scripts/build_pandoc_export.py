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


def strip_tables(text: str, label_map: dict[str, str], bib_map: dict[str, int]) -> str:
    def repl_table(match: re.Match[str]) -> str:
        block = match.group(0)
        caption_match = re.search(r"\\caption\{(.*?)\}", block, re.S)
        label_match = re.search(r"\\label\{([^}]*)\}", block)
        caption = caption_match.group(1).strip() if caption_match else "表格"
        caption = replace_refs_and_cites(caption, label_map, bib_map)
        caption = latex_to_plain(caption)
        label = label_match.group(1).strip() if label_match else ""
        number = label_map.get(label, "")
        prefix = f"表{number} " if number else "表 "
        return f"\\begin{{center}}{prefix}{caption}（表格内容在 Word 导出中省略）\\end{{center}}"

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
    text = strip_tables(text, label_map, bib_map)
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
        if stripped.startswith(r"\makecover") or stripped.startswith(r"\tableofcontents"):
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
    title = extract_macro(main_tex, "ThesisTitle") or "Thesis Export"
    author = extract_macro(main_tex, "ThesisAuthor")
    body = build_body(main_tex, thesis_dir)

    output = MAIN_TEMPLATE.format(title=title, author=author, body=body)
    output_path.write_text(output, encoding="utf-8")


if __name__ == "__main__":
    main()
