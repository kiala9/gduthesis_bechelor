@echo off
setlocal

set "ROOT=%~dp0"
set "SCRIPTS=%ROOT%scripts"
set "TOOLS=%ROOT%tools"
set "EXPORT_TEX=%ROOT%pandoc_export.tex"
set "OUTPUT_DOCX=%ROOT%thesis.docx"
set "PYTHON_CMD="

cd /d "%ROOT%"

where latexmk >nul 2>nul
if errorlevel 1 (
  echo [ERROR] latexmk not found in PATH.
  echo Install TeX Live or MiKTeX and ensure latexmk is available.
  exit /b 1
)

where xelatex >nul 2>nul
if errorlevel 1 (
  echo [ERROR] xelatex not found in PATH.
  echo Install TeX Live or MiKTeX and ensure xelatex is available.
  exit /b 1
)

where bibtex >nul 2>nul
if errorlevel 1 (
  echo [ERROR] bibtex not found in PATH.
  echo Install TeX Live or MiKTeX and ensure bibtex is available.
  exit /b 1
)

where pandoc >nul 2>nul
if errorlevel 1 (
  echo [ERROR] pandoc not found in PATH.
  echo Install Pandoc and ensure it is available in Command Prompt.
  exit /b 1
)

where py >nul 2>nul
if not errorlevel 1 (
  set "PYTHON_CMD=py -3"
)

if not defined PYTHON_CMD (
  where python >nul 2>nul
  if not errorlevel 1 (
    set "PYTHON_CMD=python"
  )
)

if not defined PYTHON_CMD (
  echo [ERROR] Python 3 not found in PATH.
  echo Install Python 3 and ensure either "py" or "python" is available.
  exit /b 1
)

echo [INFO] Refreshing LaTeX references and bibliography...
latexmk -xelatex -bibtex -interaction=nonstopmode "main.tex"
if errorlevel 1 (
  echo [ERROR] LaTeX build failed. Word export stopped.
  exit /b 1
)

echo [INFO] Building pandoc export source...
%PYTHON_CMD% "%SCRIPTS%\build_pandoc_export.py" "%ROOT%" --output "%EXPORT_TEX%"
if errorlevel 1 (
  echo [ERROR] Failed to generate pandoc_export.tex
  exit /b 1
)

echo [INFO] Converting to Word...
pandoc "%EXPORT_TEX%" ^
  --from=latex ^
  --to=docx ^
  --output="%OUTPUT_DOCX%" ^
  --reference-doc="%TOOLS%\Gthesis_reference.docx" ^
  --lua-filter="%SCRIPTS%\pandoc_keyword_style.lua" ^
  --lua-filter="%SCRIPTS%\pandoc_bibliography_list.lua" ^
  --lua-filter="%SCRIPTS%\pandoc_caption_numbering.lua" ^
  --resource-path="%ROOT%;%ROOT%figures" ^
  --top-level-division=chapter ^
  --number-sections ^
  --toc ^
  -M suppress-bibliography=true
if errorlevel 1 (
  echo [ERROR] Pandoc conversion failed.
  exit /b 1
)

echo [INFO] Post-processing Word document...
%PYTHON_CMD% "%SCRIPTS%\fix_docx_bibliography.py" "%OUTPUT_DOCX%"
if errorlevel 1 (
  echo [ERROR] Word post-processing failed.
  exit /b 1
)

echo [OK] Generated thesis.docx
exit /b 0
