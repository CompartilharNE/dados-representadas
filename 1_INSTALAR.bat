@echo off
chcp 65001 >nul
echo ============================================================
echo   COMPARTILHAR NE - Instalacao do Sistema
echo ============================================================
echo.

python --version >nul 2>&1
if errorlevel 1 (
    echo ERRO: Python nao encontrado.
    echo.
    echo Instale em: https://www.python.org/downloads/
    echo Marque "Add Python to PATH" durante a instalacao.
    echo.
    pause
    exit /b 1
)

echo Instalando dependencias...
echo.
pip install streamlit openpyxl pandas beautifulsoup4 lxml xlrd --quiet

echo.
echo ============================================================
echo   Instalacao concluida!
echo   Execute o arquivo 2_INICIAR_SERVIDOR.bat para iniciar.
echo ============================================================
echo.
pause
