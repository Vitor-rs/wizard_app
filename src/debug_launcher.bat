@echo off
REM =======================================================
REM DEBUG LAUNCHER FOR WIZPED
REM =======================================================
set LOGFILE=%~dp0debug_log.txt
echo [%DATE% %TIME%] Iniciando Launcher >> %LOGFILE%
echo [%DATE% %TIME%] Args recebidos: %* >> %LOGFILE%

cd /d %~dp0
echo [%DATE% %TIME%] CWD: %CD% >> %LOGFILE%

REM Tenta rodar com caminho absoluto do UV
set UV_PATH=C:\Users\user\.local\bin\uv.exe

echo [%DATE% %TIME%] Executando python... >> %LOGFILE%
"%UV_PATH%" run python -m wizped.cli %* >> %LOGFILE% 2>&1

IF %ERRORLEVEL% NEQ 0 (
    echo [%DATE% %TIME%] ERRO FATAL: %ERRORLEVEL% >> %LOGFILE%
    exit /b %ERRORLEVEL%
)

echo [%DATE% %TIME%] Sucesso. >> %LOGFILE%
exit /b 0
