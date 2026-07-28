@echo off
setlocal

REM ===============================================
REM Polymer MS Studio Project Setup
REM ===============================================

set PROJECT_NAME=Chemical_Manager
set GITHUB_ID=cihkill3

echo.
echo ===============================================
echo Creating project: %PROJECT_NAME%
echo ===============================================
echo.

REM ===============================================
REM GitHub
REM ===============================================

mkdir .github\workflows

REM ===============================================
REM .gitignore
REM ===============================================

(
echo __pycache__/
echo *.pyc
echo .pytest_cache/
echo .mypy_cache/
echo .ruff_cache/
echo .venv/
echo .idea/
echo .vscode/settings.json
echo build/
echo dist/
echo *.egg-info/
echo data/
echo outputs/
echo *.mzML
echo *.mzXML
echo *.raw
echo *.d
echo *.wiff
echo *.tmp
echo *.log
echo ~*.*
) > .gitignore


REM ===============================================
REM Initialize Git
REM ===============================================

git init

git branch -M main

git add .

git commit -m "Initial project structure"

git remote add origin https://github.com/%GITHUB_ID%/%PROJECT_NAME%.git

git push -u origin main

echo.
echo ===============================================
echo Setup Completed.
echo ===============================================
pause