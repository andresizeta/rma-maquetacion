#!/usr/bin/env bash
set -euo pipefail
cd "$(dirname "$0")"

.venv/bin/python generar_trama.py --vol 20 --num 1 --year 2027 --out trama.svg
.venv/bin/python -m weasyprint tapa.html tapa_v20n1.pdf --presentational-hints
.venv/bin/python -c "import fitz; doc = fitz.open('tapa_v20n1.pdf'); doc[0].get_pixmap(dpi=300).save('tapa_v20n1.png')"

echo "[OK] Portada generada: tapa_v20n1.pdf y tapa_v20n1.png"