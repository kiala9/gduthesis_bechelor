%% gdthesis.cls —— 毕业设计（论文）文档类
%% 依据：2026届毕业设计（论文）手册
%% 基于 ctexbook，使用 XeLaTeX + BibTeX 编译

\NeedsTeXFormat{LaTeX2e}
\ProvidesClass{gdthesis}[2026/03/14 Graduation Design Thesis Class]

% 将未知选项透传给 ctexbook（如 draft, twoside 等）
\DeclareOption*{\PassOptionsToClass{\CurrentOption}{ctexbook}}
\ProcessOptions\relax

\LoadClass[
  a4paper,
  zihao   = -4,   % 正文默认小四号
  UTF8,
  openany,
  oneside,
]{ctexbook}

% ================================================================
%  论文信息存储（内部变量，用 \gdt@xxx 命名）
% ================================================================
\newcommand{\gdt@title}     {论文题目（主标题不超过20字）}
\newcommand{\gdt@titleEn}   {Thesis Title in English}
\newcommand{\gdt@college}   {XX学院}
\newcommand{\gdt@major}     {XX专业（XX方向）}
\newcommand{\gdt@year}      {20XX}   % 入学年份，如 2022
\newcommand{\gdt@class}     {X}      % 班别数字，如 2
\newcommand{\gdt@studentid} {2020XXXXXXXXX}
\newcommand{\gdt@author}    {学生姓名}
\newcommand{\gdt@supervisor}{指导教师姓名}
\newcommand{\gdt@date}      {2026年\enspace X\enspace 月}

% ---- 对外设置接口（在 main.tex 中调用）----
\newcommand{\ThesisTitle}     [1]{\renewcommand{\gdt@title}     {#1}}
\newcommand{\ThesisTitleEn}   [1]{\renewcommand{\gdt@titleEn}   {#1}}
\newcommand{\ThesisCollege}   [1]{\renewcommand{\gdt@college}   {#1}}
\newcommand{\ThesisMajor}     [1]{\renewcommand{\gdt@major}     {#1}}
\newcommand{\ThesisYear}      [1]{\renewcommand{\gdt@year}      {#1}}
\newcommand{\ThesisClass}     [1]{\renewcommand{\gdt@class}     {#1}}
\newcommand{\ThesisStudentID} [1]{\renewcommand{\gdt@studentid} {#1}}
\newcommand{\ThesisAuthor}    [1]{\renewcommand{\gdt@author}    {#1}}
\newcommand{\ThesisSupervisor}[1]{\renewcommand{\gdt@supervisor}{#1}}
\newcommand{\ThesisDate}      [1]{\renewcommand{\gdt@date}      {#1}}

% ================================================================
%  页面尺寸
% ================================================================
\RequirePackage{geometry}
\geometry{
  a4paper,
  top    = 30mm,
  bottom = 25mm,
  left   = 30mm,
  right  = 20mm,
}

% ================================================================
%  字体
% ================================================================
\RequirePackage{fontspec}
\setmainfont{Times New Roman}[
  BoldFont       = Times New Roman Bold,
  ItalicFont     = Times New Roman Italic,
  BoldItalicFont = Times New Roman Bold Italic,
]
% 中文字体：在 ctex 完成初始化后覆盖，保留 \songti \heiti 等命令
\AtBeginDocument{
  \setCJKmainfont{SimSun}[AutoFakeBold=2, AutoFakeSlant=0.2]
  \setCJKsansfont{SimHei}[AutoFakeBold=2]
  \setCJKmonofont{SimHei}
}

% ================================================================
%  行距
% ================================================================
\RequirePackage{setspace}

% ================================================================
%  页眉页脚
% ================================================================
\RequirePackage{fancyhdr}

\fancypagestyle{main}{
  \fancyhf{}
  \fancyfoot[R]{\zihao{-5}\rmfamily\thepage}
  \renewcommand{\headrulewidth}{0pt}
  \renewcommand{\footrulewidth}{0pt}
}
\fancypagestyle{plain}{   % 章节首页与正文相同
  \fancyhf{}
  \fancyfoot[R]{\zihao{-5}\rmfamily\thepage}
  \renewcommand{\headrulewidth}{0pt}
  \renewcommand{\footrulewidth}{0pt}
}
\pagestyle{main}

% ================================================================
%  图形
% ================================================================
\RequirePackage{graphicx}
\RequirePackage{float}
\graphicspath{{figures/}}

% ================================================================
%  TikZ（流程图 / 架构图）
% ================================================================
\RequirePackage{tikz}
\usetikzlibrary{shapes.geometric, arrows.meta, positioning, calc, fit, backgrounds}

\RequirePackage{subfig}

% ================================================================
%  图表标题格式
% ================================================================
\RequirePackage{caption}

\DeclareCaptionFont{gdt@figfont}    {\zihao{5}\songti}
\DeclareCaptionFont{gdt@figlabel}   {\zihao{5}\rmfamily}
\DeclareCaptionFont{gdt@tabfont}    {\zihao{5}\heiti\bfseries}
\DeclareCaptionFont{gdt@tablabel}   {\zihao{5}\rmfamily\bfseries}

% 图题：五号宋体，居中
\captionsetup[figure]{
  font          = gdt@figfont,
  labelfont     = gdt@figlabel,
  textfont      = gdt@figfont,
  labelsep      = space,
  justification = centering,
}
% 表题：五号黑体加粗，居中，置于表上方
\captionsetup[table]{
  font          = gdt@tabfont,
  labelfont     = gdt@tablabel,
  textfont      = gdt@tabfont,
  labelsep      = space,
  justification = centering,
  position      = top,
  skip          = 2pt,
}

% 图、表按章编号
\renewcommand{\thefigure}{\thechapter.\arabic{figure}}
\renewcommand{\thetable} {\thechapter.\arabic{table}}

% ================================================================
%  表格
% ================================================================
\RequirePackage{booktabs}
\RequirePackage{array}
\RequirePackage{multirow}
\RequirePackage{longtable}
\RequirePackage{tabularx}    % 自动列宽，X 列类型
\RequirePackage{diagbox}     % 斜线表头 \diagbox{左}{右}
\RequirePackage{etoolbox}

% 扩展列类型
\newcolumntype{Z}{>{\centering\arraybackslash}X}   % 居中 X 列
\newcolumntype{L}{>{\raggedright\arraybackslash}X} % 左对齐 X 列
\newcolumntype{R}{>{\raggedleft\arraybackslash}X}  % 右对齐 X 列

% 行高：arraystretch 控制单元格上下内边距，1.25 与正文行距协调
\renewcommand{\arraystretch}{1.25}

% 表格内字号与行距：五号宋体，单倍行距
% （避免继承全局 1.25 倍行距导致行高过大）
\AtBeginEnvironment{tabular}  {\setstretch{1.0}\zihao{5}\songti}
\AtBeginEnvironment{tabularx} {\setstretch{1.0}\zihao{5}\songti}
\AtBeginEnvironment{longtable}{\setstretch{1.0}\zihao{5}\songti}

% ================================================================
%  数学
% ================================================================
\RequirePackage{amsmath}
\RequirePackage{amssymb}
\RequirePackage{bm}
\numberwithin{equation}{chapter}   % 公式按章编号：(1.1)

% ================================================================
%  列表
% ================================================================
\RequirePackage{enumitem}

% 全局压缩标准列表行距（替代 compactenum/compactitem）
\setlist{
  topsep     = 0pt,
  partopsep  = 0pt,
  parsep     = 0pt,
  itemsep    = 0pt,
}

% 款（1、）
\newlist{kuan}{enumerate}{1}
\setlist[kuan]{
  label      = \zihao{-4}\heiti\bfseries\arabic*、,
  leftmargin = *,
  align      = left,
  topsep     = 0pt,
  itemsep    = 0pt,
}

% 项（(1)）
\newlist{xiang}{enumerate}{1}
\setlist[xiang]{
  label      = (\arabic*),
  leftmargin = 2em,
  align      = left,
  topsep     = 0pt,
  itemsep    = 0pt,
}

% ================================================================
%  代码（附录源程序）
% ================================================================
\RequirePackage{xcolor}
\RequirePackage{listings}
\lstset{
  basicstyle       = \zihao{-5}\ttfamily,
  breaklines       = true,
  numbers          = left,
  numberstyle      = \tiny\color{gray},
  frame            = single,
  captionpos       = b,
  showstringspaces = false,
}

% ================================================================
%  URL
% ================================================================
\RequirePackage{url}
\urlstyle{same}

% ================================================================
%  参考文献（GB/T 7714-2015，gbt7714 + BibTeX）
% ================================================================
\RequirePackage[super, sort&compress]{gbt7714}
% \upcite → 上标 [n]
\newcommand{\upcite}[1]{\cite{#1}}
% \inlinecite → 行内 [n]
\newcommand{\inlinecite}[1]{[\citenum{#1}]}

% ================================================================
%  超链接（最后加载）
% ================================================================
\RequirePackage[
  hidelinks,
  unicode,
  bookmarksnumbered = true,
  bookmarksopen     = true,
]{hyperref}

% ================================================================
%  标题格式
% ================================================================
% 正文采用小四字号、固定 20pt 基线距，放在版式设置附近便于统一维护
\newcommand{\gdtbodyfontsize}{12pt}
\newcommand{\gdtbodyleading}{18pt}
\makeatletter
\renewcommand\normalsize{%
  \@setfontsize\normalsize{\gdtbodyfontsize}{\gdtbodyleading}%
  \abovedisplayskip 10pt plus 2pt minus 5pt
  \abovedisplayshortskip \z@ plus 3pt
  \belowdisplayshortskip 6pt plus 3pt minus 3pt
  \belowdisplayskip \abovedisplayskip
  \let\@listi\@listI
}
\makeatother
\AtBeginDocument{\normalsize}

\ctexset{
  chapter = {
    name       = {,},
    number     = \arabic{chapter},
    format     = \zihao{3}\heiti\bfseries\centering,
    beforeskip = 0pt,
    afterskip  = 21pt,
    fixskip    = true,
  },
  section = {
    format     = \zihao{-4}\heiti\bfseries\raggedright,
    beforeskip = 21pt,   % ≈ 一个正文行距，与行文保持一致
    afterskip  = 21pt,
    fixskip    = true,   % 禁止 ctex 叠加额外间距
  },
  subsection = {
    format     = \zihao{-4}\heiti\raggedright,
    beforeskip = 21pt,
    afterskip  = 21pt,
    fixskip    = true,
  },
}

% ================================================================
%  目录格式
% ================================================================
\RequirePackage{titletoc}

\renewcommand{\contentsname}{目录}
%   \vspace{0bp}%

% 目录整体行距独立控制：保存全局行距 → 切到 1.5 → 输出 TOC → 还原
\makeatletter
\AtBeginDocument{%
  \renewcommand{\tableofcontents}{%
    \edef\gdt@savedstretch{\baselinestretch}%
    \clearpage
    \phantomsection
    \chapter*{\centering\zihao{3}\heiti\bfseries \contentsname}%
    \@mkboth{\contentsname}{\contentsname}%
    \vspace{-20bp}%
    \setstretch{1.65}%
%     \fontsize{\gdtbodyfontsize}{\gdtbodyleading}\selectfont
    \@starttoc{toc}%
    \expandafter\setstretch\expandafter{\gdt@savedstretch}%
  }%
}
\makeatother

% 章（一级）：小四号黑体
% 编号/页码用 \fontfamily\rmdefault\selectfont 只切西文字体，
% 不触发 ctex 对 \rmfamily 的重定义（会连带把 CJK 切回宋体）
\titlecontents{chapter}
  [0em]
  {\zihao{-4}\heiti}
  {\begingroup\fontfamily{\rmdefault}\selectfont\thecontentslabel\endgroup\quad}
  {}
  {\titlerule*[0.25pc]{.}\begingroup\fontfamily{\rmdefault}\selectfont\contentspage\endgroup}

% 节（二级）：小四号宋体
\titlecontents{section}
  [2em]
  {\zihao{-4}\songti}
  {\begingroup\fontfamily{\rmdefault}\selectfont\thecontentslabel\endgroup\quad}
  {}
  {\titlerule*[0.25pc]{.}\begingroup\fontfamily{\rmdefault}\selectfont\contentspage\endgroup}

% 条（三级）：小四号宋体
\titlecontents{subsection}
  [4em]
  {\zihao{-4}\songti}
  {\begingroup\fontfamily{\rmdefault}\selectfont\thecontentslabel\endgroup\quad}
  {}
  {\titlerule*[0.25pc]{.}\begingroup\fontfamily{\rmdefault}\selectfont\contentspage\endgroup}

% ================================================================
%  用户命令
% ================================================================

% ---- 交叉引用快捷命令 ----
\newcommand{\figref}[1]{图~\ref{#1}}
\newcommand{\tabref}[1]{表~\ref{#1}}
\newcommand{\secref}[1]{第~\ref{#1}~节}
% \eqref 由 amsmath 提供，格式为 (1.1)

% ---- 带下划线的固定宽度文本框（封面信息栏使用）----
% 用法：\dlmu{内容}  或  \dlmu[8cm]{内容}
\newcommand{\dlmu}[2][7cm]{%
  \underline{\makebox[#1][c]{#2}}%
}

% ---- 封面（移植自 GraduationDesign/data/cover.tex）----
\makeatletter
\newcommand{\makecover}{%
  \begin{titlepage}%
    % 校徽（左对齐）
    \begin{flushleft}
      \hspace{8.5mm}%
      \IfFileExists{figures/xiaohui.jpg}{%
        \includegraphics[height=2.19cm,width=2.21cm]{figures/xiaohui.jpg}%
      }{%
        \vspace{2.19cm}%   % 图片缺失时占位
      }%
    \end{flushleft}
    % 校名（居中）
    \begin{center}
      \IfFileExists{figures/mingchen.jpg}{%
        \includegraphics[height=2.96cm,width=10.56cm]{figures/mingchen.jpg}%
      }{%
        \vspace{2.96cm}%
      }%
    \end{center}
    \vspace{1mm}
    % 论文类型
    \begin{center}
      {\zihao{1}\heiti\bfseries 本科毕业设计（论文）}
    \end{center}
    \vspace{12mm}
    % 论文题目
    \begin{center}
      {\zihao{2}\heiti\bfseries \gdt@title}
    \end{center}
    \vspace{26mm}
    % 信息栏
    \begin{center}
      {\zihao{3}\heiti 学\hspace{2em}院:\hspace{3mm}\dlmu{\textbf{\gdt@college}}}\\[4mm]
      {\zihao{3}\heiti 专\hspace{2em}业:\hspace{3mm}\dlmu{\textbf{\gdt@major}}}\\[4mm]
      {\zihao{3}\heiti 年级班别:\hspace{3mm}\dlmu{\textbf{\gdt@year 级（\gdt@class）班}}}\\[4mm]
      {\zihao{3}\heiti 学\hspace{2em}号:\hspace{3mm}\dlmu{\textbf{\gdt@studentid}}}\\[4mm]
      {\zihao{3}\heiti 学生姓名:\hspace{3mm}\dlmu{\textbf{\gdt@author}}}\\[4mm]
      {\zihao{3}\heiti 指导教师:\hspace{3mm}\dlmu{\textbf{\gdt@supervisor}}}%
    \end{center}
    \vfill
    % 日期
    \begin{center}
      {\zihao{3}\heiti \gdt@date}
    \end{center}
  \end{titlepage}%
}
\makeatother

% ---- 正文基准字体命令 ----
\newcommand{\gdtbodytext}{%
  \fontsize{\gdtbodyfontsize}{\gdtbodyleading}\selectfont\songti
}

% ---- 中文摘要环境 ----
\newenvironment{cnabstract}{%
  \clearpage\thispagestyle{empty}%
  {\centering\zihao{3}\heiti\bfseries 摘要\par}%
  \vspace{0bp}%
  \gdtbodytext
}{%
  \clearpage%
}

% 中文关键词
\newcommand{\cnkeywords}[1]{%
  \par\vspace{\baselineskip}%
  {\zihao{4}\heiti\bfseries 关键词：}{\gdtbodytext #1}%
}

% ---- 英文摘要环境 ----
\newenvironment{enabstract}{%
  \clearpage\thispagestyle{empty}%
  \begin{center}{\zihao{3}\rmfamily\bfseries Abstract}\end{center}%
  \setstretch{1.4}%   % 英文摘要行距，可单独调整
  \zihao{-4}\rmfamily%
}{%
  \clearpage%
}

% 英文关键词
\newcommand{\enkeywords}[1]{%
  \par\vspace{\baselineskip}%
  {\zihao{4}\rmfamily\bfseries Key words:\enspace}{\zihao{-4}\rmfamily #1}%
}

% ---- 无编号章节（出现在目录）----
\newcommand{\unnumberedchapter}[1]{%
  \chapter*{#1}%
  \addcontentsline{toc}{chapter}{#1}%
  \markboth{#1}{#1}%
}

% ---- 附录目录写入控制 ----
\makeatletter
\newtoggle{gdt@inappendix}
\togglefalse{gdt@inappendix}
\let\gdt@origaddcontentsline\addcontentsline
\renewcommand{\addcontentsline}[3]{%
  \ifboolexpr{
    togl {gdt@inappendix}
    and test {\ifstrequal{#1}{toc}}
    and test {\ifstrequal{#2}{chapter}}
  }%
    {\gdt@origaddcontentsline{#1}{#2}{附录\thechapter}}%
    {\gdt@origaddcontentsline{#1}{#2}{#3}}%
}
\makeatother

% ---- 参考文献输出 ----
\newcommand{\printreferences}{%
  \clearpage
  \phantomsection
  \addcontentsline{toc}{chapter}{参考文献}%
%   \setstretch{1.5}
  {\gdtbodytext\setlength{\bibsep}{0pt}\bibliography{references}}%
}

% ---- 进入附录模式 ----
\newcommand{\startappendix}{%
  \clearpage
  \appendix
  \toggletrue{gdt@inappendix}%
  \ctexset{
    chapter/name   = {附录,},
    chapter/number = \Alph{chapter},
  }%
  \renewcommand{\thefigure}  {\Alph{chapter}\arabic{figure}}%
  \renewcommand{\thetable}   {\Alph{chapter}\arabic{table}}%
  \renewcommand{\theequation}{\Alph{chapter}\arabic{equation}}%
}
