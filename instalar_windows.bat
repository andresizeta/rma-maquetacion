@echo off
echo Creando entorno virtual Python para RMA...
python -m venv .venv
call .venv\Scripts\activate.bat
python -m pip install --upgrade pip
pip install -r requirements.txt
echo.
echo Instalacion completada. Para probar el articulo 4812, ejecuta:
echo .venv\Scripts\python maquetar_rma.py ejemplo_ojs 4812
echo Para probar la paginacion consecutiva de un fasciculo completo:
echo .venv\Scripts\python maquetar_fasciculo.py ejemplo_fasciculo
pause
