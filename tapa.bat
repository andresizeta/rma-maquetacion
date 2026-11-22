@echo off
chcp 65001 >nul
set PYTHONUTF8=1
cd /d "%~dp0"

.venv\Scripts\python.exe generar_trama.py --vol 20 --num 1 --year 2027 --out trama.svg
if %errorlevel% neq 0 goto :error

.venv\Scripts\python.exe -m weasyprint tapa.html tapa_v20n1.pdf --presentational-hints
if %errorlevel% neq 0 goto :error

.venv\Scripts\python.exe -c "import fitz; doc = fitz.open('tapa_v20n1.pdf'); doc[0].get_pixmap(dpi=300).save('tapa_v20n1.png')"
if %errorlevel% neq 0 goto :error

echo [OK] Portada generada: tapa_v20n1.pdf y tapa_v20n1.png
pause
exit /b 0

:error
echo [ERROR] Fallo la generacion de la portada.
pause
exit /b 1