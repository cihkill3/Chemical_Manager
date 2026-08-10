@echo off
setlocal
REM 인수가 없으면 사용법 출력
if "%~1"=="" (
    echo.
    echo git commit하는 배치파일입니다.
    echo.
    echo 사용법:
    echo    %~nx0 코멘트
    echo.
    echo 코멘트는 따옴표 없이 문자열을 입력합니다. 공백도 처리합니다.
    echo 관리자권한으로 실행하세요
    echo.
    exit /b 1
)
echo %*
REM ===============================================
REM Polymer MS Studio Project Setup
REM ===============================================

set PROJECT_NAME=Chemical_Manager
set GITHUB_ID=cihkill3

git add .

git commit -m "%*"

git remote add origin https://github.com/%GITHUB_ID%/%PROJECT_NAME%.git

git push -u origin main

echo.
echo ===============================================
echo Commit Completed.
echo ===============================================
git log --oneline -5
pause

:end