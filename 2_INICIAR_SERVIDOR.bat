@echo off
chcp 65001 >nul
echo ============================================================
echo   COMPARTILHAR NE - Iniciando Servidor
echo ============================================================
echo.
echo  Acesse no navegador: http://localhost:8501
echo  Na rede local:       http://[IP-DO-SERVIDOR]:8501
echo.
echo  Para parar o servidor pressione Ctrl+C
echo ============================================================
echo.

cd /d "%~dp0"
streamlit run app.py --server.port 8501 --server.address 0.0.0.0 --server.headless true --browser.gatherUsageStats false
pause
