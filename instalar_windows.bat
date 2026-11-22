@echo off
chcp 65001 > nul
set "PYTHONIOENCODING=utf-8"

if not defined WEASYPRINT_DLL_DIRECTORIES (
    if exist "C:\Program Files\GTK3-Runtime Win64\bin\libgobject-2.0-0.dll" (
        set "WEASYPRINT_DLL_DIRECTORIES=C:\Program Files\GTK3-Runtime Win64\bin"
    ) else if exist "C:\msys64\mingw64\bin\libpango-1.0-0.dll" (
        set "WEASYPRINT_DLL_DIRECTORIES=C:\msys64\mingw64\bin"
    )
)

echo [1/4] Verificando Python...
python --version > nul 2>&1
if errorlevel 1 (
    echo [ERROR] No se encontro Python en el sistema. Instala Python 3.11 o superior ^(64-bit^) marcando "Add python.exe to PATH".
    pause
    exit /b 1
)

echo [2/4] Creando entorno virtual .venv...
if not exist ".venv\Scripts\python.exe" (
    python -m venv .venv
    if errorlevel 1 (
        echo [ERROR] No se pudo crear el entorno virtual .venv.
        pause
        exit /b 1
    )
)

echo [3/4] Instalando dependencias en .venv...
.venv\Scripts\python.exe -m pip install --upgrade pip
if errorlevel 1 (
    echo [ERROR] Se produjo un fallo al actualizar pip.
    pause
    exit /b 1
)

.venv\Scripts\python.exe -m pip install -r requirements.txt
if errorlevel 1 (
    echo [ERROR] Se produjo un fallo durante la instalacion de dependencias.
    pause
    exit /b 1
)

echo [4/4] Verificando motor grafico de WeasyPrint ^(GTK3/Pango^)...
.venv\Scripts\python.exe -c "import weasyprint; print('WeasyPrint', weasyprint.__version__, 'OK')"
if errorlevel 1 (
    echo.
    echo [AVISO] Python y los paquetes se instalaron, pero falta el runtime GTK3/Pango ^(64-bit^) en Windows.
    echo Instala GTK3-Runtime Win64 o MSYS2 ^(mingw-w64-x86_64-pango^) y vuelve a ejecutar este script.
    pause
    exit /b 1
)

echo.
echo [OK] Instalacion completada correctamente.
pause
exit /b 0
