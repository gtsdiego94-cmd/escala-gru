@echo off
chcp 65001 > nul
cd /d "C:\Users\User\Desktop\escala-gru"
echo Executando extrator e relatorio de vendas OneBoss...
node one-boss-report.js
