@echo off
echo.
echo =============================================
echo   �����x����Ðf�f���i���_�ʉ@��×p�j�쐬
echo   Word�Łi.docx�j�Ő������܂�
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
python run_batch.py jiritsu_shien_word %*
echo.
if %ERRORLEVEL% EQU 0 (
    echo �o�͐�: output �t�H���_���m�F���Ă��������B
) else (
    echo �G���[���������܂����B
)
echo.
pause
