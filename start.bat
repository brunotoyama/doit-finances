@echo off
cd /d "%~dp0"
echo ===================================
echo   DOit Finances - Dashboard
echo ===================================
echo.
echo Iniciando servidor em http://localhost:3000/
echo.
npx serve . -p 3000 -o || echo ERRO: Porta 3000 ja esta em uso. Tente: npx serve . -p 3001
