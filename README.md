# 广东工业大学本科毕业设计（论文）模板

这是一套面向广东工业大学本科毕业设计（论文）的 LaTeX 模板，并附带一条可提交 Word 稿件的导出链路。

仓库定位以 `Word` 交付为主，`PDF` 编译保留用于排版核对与模板维护。

## 功能

- 按当前整理版本生成 PDF 论文
- 生成可继续在 Word 中编辑、检查和提交的 `.docx`
- 覆盖封面、中英文摘要、目录、正文、参考文献、致谢、附录
- 对目录、摘要、关键词、图表标题、参考文献做了额外的 Word 导出适配

## 目录结构

```text
gduthesis_bechelor/
├── main.tex
├── gdthesis.cls
├── references.bib
├── chapters/
├── figures/
├── scripts/
├── tools/
├── build.bat / build.sh
└── export_word.bat / export_word.sh
```

## Windows 依赖

首版采用“预装依赖后一键导出”方案。请先安装：

- TeX Live 或 MiKTeX，且命令行可用 `latexmk`、`xelatex`、`bibtex`
- Pandoc
- Python 3
- Microsoft Word

如果命令未加入 `PATH`，`build.bat` 和 `export_word.bat` 会直接报错。

## 快速开始

1. 修改 [main.tex](/home/kiala/linux_share/GDesign/Template/gduthesis_bechelor/main.tex) 顶部的论文信息。
2. 在 `chapters/` 中填写论文内容。
3. 在 `figures/` 中放置图片。
4. 在 [references.bib](/home/kiala/linux_share/GDesign/Template/gduthesis_bechelor/references.bib) 中维护参考文献。

## 生成 PDF

Windows 下双击 [build.bat](/home/kiala/linux_share/GDesign/Template/gduthesis_bechelor/build.bat)，或在终端中执行：

```bat
build.bat
```

Linux / macOS / WSL 下执行：

```bash
sh build.sh
```

成功后会在仓库根目录生成 `main.pdf`。

## 导出 Word

Windows 下双击 [export_word.bat](/home/kiala/linux_share/GDesign/Template/gduthesis_bechelor/export_word.bat)，或在终端中执行：

```bat
export_word.bat
```

Linux / macOS / WSL 下执行：

```bash
sh export_word.sh
```

脚本会自动完成以下步骤：

1. 先运行 `latexmk`，刷新目录、交叉引用与参考文献
2. 生成用于 Pandoc 的中间文件 `pandoc_export.tex`
3. 导出 `thesis.docx`
4. 对 Word 文档做后处理修补

最终输出文件：

- `main.pdf`
- `thesis.docx`

## 可修改与不建议修改的文件

通常只需要修改：

- [main.tex](/home/kiala/linux_share/GDesign/Template/gduthesis_bechelor/main.tex)
- `chapters/` 下各章节文件
- [references.bib](/home/kiala/linux_share/GDesign/Template/gduthesis_bechelor/references.bib)
- `figures/` 下图片资源

不建议直接修改，除非你明确知道后果：

- [gdthesis.cls](/home/kiala/linux_share/GDesign/Template/gduthesis_bechelor/gdthesis.cls)
- `scripts/` 下的 Python 与 Lua 脚本
- [tools/Gthesis_reference.docx](/home/kiala/linux_share/GDesign/Template/gduthesis_bechelor/tools/Gthesis_reference.docx)

## 已知限制

- Word 导出时，PDF 格式插图可能无法正确嵌入
- 若论文最终需要提交 Word，建议插图优先使用 PNG/JPG
- 复杂表格在 Word 导出中可能退化为仅保留表题，需要人工复核
- Word 稿件虽然已做适配，但提交前仍建议逐页检查分页、图表位置与目录结果
- 模板会持续调整，若学校格式要求变动，请自行复核

## 发布建议

- 仓库发布时，建议将当前完整目录直接推送到 GitHub
- 若需要给已 clone 的同学发增量更新，建议在 GitHub Release 中同时附带 `.patch` 文件
- 可使用 [scripts/make_release_patch.sh](/home/kiala/linux_share/GDesign/Template/gduthesis_bechelor/scripts/make_release_patch.sh) 从两个 tag 或 commit 之间生成补丁
- 建议在 Release 说明里明确提示：Word 导出优先使用 PNG/JPG 图片，并对复杂表格做人工复核

## 许可

本仓库使用 MIT License。
