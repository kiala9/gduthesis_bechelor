# 毕业设计（论文）LaTeX 模板 for 2026

本目录是按学院手册样张整理过的一套本科毕业设计（论文）LaTeX 模板，适合直接复制后填写自己的内容。

## 适用范围

- `XeLaTeX + BibTeX` 编译流程
- 中文论文正文、小四字号、固定正文行距
- 含封面、中英文摘要、目录、正文、结论、参考文献、致谢、附录

## 目录说明

```text
Template/latex/
├── main.tex              # 主文件，优先修改这里
├── gdthesis.cls          # 模板类文件，统一控制格式
├── references.bib        # 参考文献数据库
├── refs.bib              # 备用/历史 bib 文件
├── figures/              # 图片目录
└── chapters/             # 章节内容目录
```

## 快速开始

1. 复制 `Template/latex/` 到自己的论文目录。
2. 修改 [main.tex](/home/kiala/linux_share/GDesign/Template/latex/main.tex) 顶部的论文信息。
3. 在 `chapters/` 中填写各章节内容。

## 编译方式

推荐顺序：

```bash
xelatex main.tex
bibtex main
xelatex main.tex
xelatex main.tex
```

也可使用 `latexmk` 或编辑器内置的 `XeLaTeX` 工作流。

## 说明

- 这个模板已经按手册样张做过一轮校对，但仍建议在最终提交前自行逐页检查。
- 如果只改论文内容，不建议随意修改 `gdthesis.cls`。
- 如果发现格式问题，优先检查是否在章节文件里手动写了字号、行距或额外空行。
- 确实存在问题的，欢迎提出Issue。
