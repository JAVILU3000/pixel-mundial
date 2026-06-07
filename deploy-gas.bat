@echo off
chcp 65001 >nul

echo Copiando gas-code.gs al portapapeles...
type "%~dp0gas-code.gs" | clip

echo Abriendo Google Apps Script en el navegador...
start "" "https://script.google.com/home/projects/16LxznLW90JIfug3aQB59dk8P3ees3EJGwmc_CRgjeMTR1MIEpDryO1nJ/edit"

echo.
echo ============================================================
echo  Codigo copiado al portapapeles.
echo  Ahora en el editor de GAS:
echo    Ctrl+A  (seleccionar todo)
echo    Ctrl+V  (pegar el codigo nuevo)
echo    Ctrl+S  (guardar)
echo    luego: Implementar
echo ============================================================
echo.
pause
