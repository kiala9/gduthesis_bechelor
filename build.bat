@echo off
setlocal

set "ROOT=%~dp0"
cd /d "%ROOT%"

where latexmk >nul 2>nul
if errorlevel 1 (
  echo [ERROR] latexmk not found in PATH.
  echo Install TeX Live or MiKTeX and ensure latexmk is available in Command Prompt.
  exit /b 1
)

where xelatex >nul 2>nul
if errorlevel 1 (
  echo [ERROR] xelatex not found in PATH.
  echo Install TeX Live or MiKTeX and ensure xelatex is available in Command Prompt.
  exit /b 1
)

where bibtex >nul 2>nul
if errorlevel 1 (
  echo [ERROR] bibtex not found in PATH.
  echo Install TeX Live or MiKTeX and ensure bibtex is available in Command Prompt.
  exit /b 1
)

echo [INFO] Building PDF with XeLaTeX...
latexmk -xelatex -bibtex -interaction=nonstopmode "main.tex"
if errorlevel 1 (
  echo [ERROR] PDF build failed.
  exit /b 1
)

echo [OK] Generated main.pdf
exit /b 0
