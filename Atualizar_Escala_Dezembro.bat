@echo off
chcp 65001 > nul
title Atualizar Escala de Dezembro - Protecbag GRU
cd /d "C:\Users\User\Desktop\escala-gru"

echo =======================================================
echo   PROTECBAG GRU - ATUALIZADOR DE ESCALA DE DEZEMBRO
echo =======================================================
echo.
echo Executando importador...
echo.

"C:\Users\User\AppData\Local\Programs\Python\Python312\python.exe" importar_escala_dezembro.py %*

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo ❌ Ocorreu um erro durante a execucao.
    pause
)
