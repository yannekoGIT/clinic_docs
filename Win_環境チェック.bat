@echo off
chcp 65001 >nul
echo.
echo =============================================
echo   ���Z�b�g�A�b�v �� �`�F�b�N
echo =============================================
echo.
cd /d "%~dp0"

:: ---- 1. Python �`�F�b�N ----
where python >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo [NG] Python ��������܂���Bwinget �ŃC���X�g�[�����܂�...
    echo.
    winget install -e --id Python.Python.3.12 --accept-package-agreements --accept-source-agreements
    if %ERRORLEVEL% NEQ 0 (
        echo.
        echo [ERROR] Python �̃C���X�g�[���Ɏ��s���܂����B
        echo         �蓮�ŃC���X�g�[�����Ă�������: https://www.python.org/downloads/
        echo.
        pause
        exit /b 1
    )
    echo.
    echo [OK] Python ���C���X�g�[�����܂����B
    echo [!] PATH �𔽉f���邽�߂ɂ��̃E�B���h�E����āA�ēx���s���Ă��������B
    echo.
    pause
    exit /b 0
)
for /f "tokens=*" %%i in ('python --version 2^>^&1') do echo [OK] %%i

:: ---- 2. Node.js �`�F�b�N ----
where node >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo [NG] Node.js ��������܂���Bwinget �ŃC���X�g�[�����܂�...
    echo.
    winget install -e --id OpenJS.NodeJS.LTS --accept-package-agreements --accept-source-agreements
    if %ERRORLEVEL% NEQ 0 (
        echo.
        echo [ERROR] Node.js �̃C���X�g�[���Ɏ��s���܂����B
        echo         �蓮�ŃC���X�g�[�����Ă�������: https://nodejs.org/
        echo.
        pause
        exit /b 1
    )
    echo.
    echo [OK] Node.js ���C���X�g�[�����܂����B
    echo [!] PATH �𔽉f���邽�߂ɂ��̃E�B���h�E����āA�ēx���s���Ă��������B
    echo.
    pause
    exit /b 0
)
for /f "tokens=*" %%i in ('node --version 2^>^&1') do echo [OK] Node.js %%i

:: ---- 3. Python �̊��`�F�b�N�������C���X�g�[���� ----
echo.
python generate.py --check --auto-install
echo.
pause
