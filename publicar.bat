@echo off
chcp 65001 > nul
echo ============================================================
echo   PUBLICAR — copia codigo para producao
echo ============================================================
echo.

set PROD=P:\Publico\Marcus V\Contagem Antilhas\AutomacaoAntilhas\antilhas

echo Destino: %PROD%
echo.

if not exist "%PROD%" (
    echo ERRO: pasta de producao nao encontrada.
    echo Verifique se o drive P: esta conectado.
    pause
    exit /b 1
)

echo Copiando arquivos de codigo...

copy /Y "processo1.py"         "%PROD%\processo1.py"
copy /Y "processo2.py"         "%PROD%\processo2.py"
copy /Y "rodar.bat"            "%PROD%\rodar.bat"
copy /Y "rodar_p2.bat"         "%PROD%\rodar_p2.bat"

echo.
echo Arquivos NAO copiados (existem so em producao):
echo   - config_p1.yaml
echo   - config_p2.yaml
echo   - 04_PLANILHAS_PAI\
echo   - 00_ENTRADA\
echo.
echo Publicacao concluida com sucesso!
echo ============================================================
pause
