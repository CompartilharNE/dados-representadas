@echo off
chcp 65001 >nul
echo.
echo ============================================================
echo   Instalando servico Windows - Dados Representadas
echo ============================================================
echo.

:: Pasta fixa do app (independente de onde o bat for executado)
set APP_DIR=Z:\COMPARTILHAR VENDAS\CLAUDE\Marcos\Dados Representadas\
set NSSM=%APP_DIR%nssm.exe

cd /d "%APP_DIR%"
echo Pasta do app: %APP_DIR%
echo.

:: Verificar NSSM
if not exist "%NSSM%" (
    echo ERRO: nssm.exe nao encontrado em %APP_DIR%
    echo Coloque o nssm.exe nessa pasta e tente novamente.
    pause
    exit /b 1
)
echo [OK] nssm.exe encontrado

:: Verificar Python
python --version >nul 2>&1
if errorlevel 1 (
    echo ERRO: Python nao encontrado no PATH.
    pause
    exit /b 1
)
echo [OK] Python encontrado

:: Verificar streamlit
python -m streamlit --version >nul 2>&1
if errorlevel 1 (
    echo ERRO: Streamlit nao instalado. Execute 1_INSTALAR.bat primeiro.
    pause
    exit /b 1
)
echo [OK] Streamlit encontrado

:: Descobrir caminho do Python
for /f "tokens=*" %%i in ('where python') do set PYTHON_EXE=%%i
echo [OK] Python em: %PYTHON_EXE%

:: Arquivo do app e logs
set APP_FILE=%APP_DIR%app.py
set LOG_DIR=%APP_DIR%logs

echo [OK] App em: %APP_DIR%
echo.

:: Criar pasta de logs
if not exist "%LOG_DIR%" mkdir "%LOG_DIR%"

:: Remover servico anterior se existir
echo Removendo servico anterior se existir...
"%NSSM%" stop DadosRepresentadas 2>nul
"%NSSM%" remove DadosRepresentadas confirm 2>nul
timeout /t 2 /nobreak >nul

:: Instalar servico usando python -m streamlit
echo Instalando servico...
"%NSSM%" install DadosRepresentadas "%PYTHON_EXE%"
"%NSSM%" set DadosRepresentadas AppParameters "-m streamlit run \"%APP_FILE%\" --server.port 8501 --server.address 0.0.0.0 --server.headless true --browser.gatherUsageStats false"
"%NSSM%" set DadosRepresentadas AppDirectory "%APP_DIR%"
"%NSSM%" set DadosRepresentadas DisplayName "Dados Representadas - Compartilhar NE"
"%NSSM%" set DadosRepresentadas Start SERVICE_AUTO_START
"%NSSM%" set DadosRepresentadas AppStdout "%LOG_DIR%\output.log"
"%NSSM%" set DadosRepresentadas AppStderr "%LOG_DIR%\error.log"
"%NSSM%" set DadosRepresentadas AppRotateFiles 1
"%NSSM%" set DadosRepresentadas AppRotateBytes 10485760

:: Iniciar servico
echo Iniciando servico...
"%NSSM%" start DadosRepresentadas
timeout /t 3 /nobreak >nul

:: Verificar status
"%NSSM%" status DadosRepresentadas

echo.
echo ============================================================
echo   Concluido!
echo   Acesse: http://192.168.18.190:8501
echo   Para gerenciar: services.msc - "Dados Representadas"
echo ============================================================
echo.
pause
