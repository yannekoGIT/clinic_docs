@echo off
chcp 65001 >nul
echo.
echo =============================================
echo   �Љ��i�f�Ï��񋟏��j�쐬
echo   �R�s�y�p�e�L�X�g�i.txt�j�𐶐����܂�
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
python run_batch.py referral %*
echo.
if %ERRORLEVEL% EQU 0 (
    echo �o�͐�: output �t�H���_���m�F���Ă��������B
) else (
    echo �G���[���������܂����B
)
echo.
pause
