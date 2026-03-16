# Makefile —— 论文编译
#
# 日常编译（只改正文，不动 .bib）：
#   make
#
# 更新参考文献后（改了 references.bib）：
#   make bib

.PHONY: all bib clean

# 默认：latexmk 全自动（xelatex + bibtex + xelatex × 2）
all:
	latexmk main

# 清理中间文件（保留 pdf）
clean:
	latexmk -c
