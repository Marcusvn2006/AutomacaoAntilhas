@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo ============================================================
echo   PROCESSO 3 - ANTILHAS
echo ============================================================
echo.

python processo3.py

echo.
if %errorlevel% neq 0 (
    echo  ERRO: O script terminou com falha. Veja as mensagens acima.
) else (
    echo  Concluido com sucesso!
)

echo.
pause
