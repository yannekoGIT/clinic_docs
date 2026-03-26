@echo off
chcp 65001 >nul
echo.
echo =============================================
echo   ��Q�N���f�f���i���_�̏�Q�p�j�o�b�`�쐬
echo =============================================
echo.
cd /d "%~dp0"
if "%1"=="--check" (
    python generate.py --check
    echo.
    pause
    exit /b
)
where python >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo [ERROR] Python ��������܂���B
    pause
    exit /b 1
)
echo ��Q�N���f�f���i.xlsx�j�𐶐����܂��B
echo.
python run_batch.py shougai_nenkin %*
echo.
if %ERRORLEVEL% EQU 0 (
    echo �o�͐�: output �t�H���_���m�F���Ă��������B
) else (
    echo �G���[���������܂����B
)
echo.
pause
