# .latexmkrc —— latexmk 项目配置
# 用法：latexmk main  （无需额外参数，此文件已配置好）

# ---- 编译引擎 ----
# pdf_mode = 5：xelatex → xdv → pdf（推荐 XeLaTeX 项目使用）
$pdf_mode = 5;
$xelatex  = 'xelatex -synctex=1 -interaction=nonstopmode -file-line-error %O %S';

# ---- 参考文献 ----
# bibtex_use = 1：有 .bib 文件时自动运行 bibtex
# bibtex_use = 2：强制每次都运行 bibtex（遇到文献未更新时用此值）
$bibtex_use = 1;

# ---- 主文件 ----
@default_files = ('main.tex');

# ---- 清理时额外删除的文件 ----
$clean_ext = 'synctex.gz bbl run.xml bcf fdb_latexmk tdo';
