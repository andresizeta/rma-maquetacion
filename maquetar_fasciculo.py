#!/usr/bin/env python3
"""
Compilador de Número / Fascículo Completo - Revista del Museo de Antropología (RMA)
Asigna números de página consecutivos automáticamente entre todos los artículos
que componen un número, sincroniza sus metadatos YAML y genera PDF, HTML y XML para cada uno.

Uso:
  python maquetar_fasciculo.py ejemplo_fasciculo
  python maquetar_fasciculo.py ruta/al/fasciculo --formato tradicional
  python maquetar_fasciculo.py ruta/al/fasciculo --formato moderno
"""

import sys
from pathlib import Path
from maquetar_rma import cargar_config, compilar_articulo


def procesar_fasciculo(carpeta_str: str = ".", formato_cli: str | None = None):
    carpeta = Path(carpeta_str).resolve()
    cfg = cargar_config(carpeta / "fasciculo", obligatorio=True)

    volumen = str(cfg.get("volumen", "")).strip()
    numero = str(cfg.get("numero", "")).strip()
    anio = str(cfg.get("anio", "")).strip()
    formato_fasc = formato_cli or (str(cfg.get("formato", "")).strip().lower() or None)
    cursor_pagina = int(cfg.get("pagina_inicial_numero", 1))
    actualizar_yaml = bool(cfg.get("actualizar_metadatos_yaml", True))

    articulos = cfg.get("articulos", [])
    if not articulos:
        raise ValueError(f"No hay artículos definidos en {carpeta / 'fasciculo.yaml'}")

    print(f"=== Compilando fascículo Vol. {volumen} N.º {numero} ({anio}) ===")
    if formato_fasc:
        print(f"Perfil de diseño aplicado: {formato_fasc.upper()}")
    print(f"Página inicial del bloque de artículos: {cursor_pagina}\n")

    resumen = []
    for item in articulos:
        if isinstance(item, dict):
            art_id = str(item.get("id") or item.get("id_ojs")).strip()
            overrides = {
                "codigo_publicacion": item.get("codigo_publicacion", ""),
                "seccion": item.get("seccion", ""),
                "volumen": volumen,
                "numero": numero,
                "anio": anio,
            }
        else:
            art_id = str(item).strip()
            overrides = {
                "volumen": volumen,
                "numero": numero,
                "anio": anio,
            }

        info_art = compilar_articulo(
            carpeta=carpeta,
            art_id=art_id,
            pagina_inicio_override=cursor_pagina,
            meta_overrides=overrides,
            actualizar_yaml=actualizar_yaml,
            formato_override=formato_fasc,
        )
        resumen.append(info_art)
        cursor_pagina = info_art["pagina_fin"] + 1
        print()

    carpeta_salida = carpeta / "salida"
    carpeta_salida.mkdir(parents=True, exist_ok=True)
    lineas_tabla = [
        f"Índice de paginación consecutiva — Revista del Museo de Antropología Vol. {volumen} ({numero}), {anio}",
        "=" * 92,
        f"{'Código':<14} {'ID OJS':<8} {'Páginas':<10} {'Folios':<8} {'Formato':<13} {'Sección':<18} Título",
        "-" * 92,
    ]
    for r in resumen:
        lineas_tabla.append(
            f"{r['codigo_publicacion']:<14} {r['id_ojs']:<8} {r['paginas']:<10} "
            f"{r['total_paginas']:<8} {r['formato']:<13} {r['seccion']:<18} {r['titulo']}"
        )
    texto_resumen = "\n".join(lineas_tabla) + "\n"
    ruta_resumen = carpeta_salida / f"paginacion_v{volumen}n{numero}.txt"
    ruta_resumen.write_text(texto_resumen, encoding="utf-8")

    print(texto_resumen)
    print(f"Resumen guardado en: {ruta_resumen}")


if __name__ == "__main__":
    args = sys.argv[1:]
    fmt = None
    pos = []
    idx = 0
    while idx < len(args):
        if args[idx] in ("--formato", "-f") and (idx + 1) < len(args):
            fmt = args[idx + 1].strip().lower()
            idx += 2
        else:
            pos.append(args[idx])
            idx += 1
    destino = pos[0] if pos else "."
    procesar_fasciculo(destino, formato_cli=fmt)
