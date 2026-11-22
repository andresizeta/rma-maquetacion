#!/usr/bin/env bash
set -e
echo 'Creando entorno virtual Python para RMA...'
python3 -m venv .venv
source .venv/bin/activate
pip install --upgrade pip
pip install -r requirements.txt
echo ''
echo 'Instalación completada. Para probar un artículo individual:'
echo '.venv/bin/python maquetar_rma.py ejemplo_ojs 4812'
echo 'Para probar un fascículo completo con paginación consecutiva:'
echo '.venv/bin/python maquetar_fasciculo.py ejemplo_fasciculo'
