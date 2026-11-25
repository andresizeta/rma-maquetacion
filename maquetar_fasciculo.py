#!/usr/bin/env python3
"""
Compilador de Número / Fascículo Completo - Revista del Museo de Antropología (RMA)
Busca automáticamente cada artículo en su subcarpeta (ej. volumen19_n3_54854/),
asigna números de página consecutivos, sincroniza los YAML y reúne todos los PDF,
HTML y XML finales en una carpeta consolidada.

Uso:
  .venv/bin/python maquetar_fasciculo.py .
  .venv/bin/python maquetar_fasciculo.py . --formato tradicional
"""

import shutil
import sys
from pathlib import Path
from maquetar_rma import cargar_config, compilar_articulo


def _resolver_carpeta_articulo(raiz: Path, art_id: str, carpeta_explicit: str = "") -> Path:
    if carpeta_explicit:
        cand = Path(carpeta_explicit)
        cand = cand.resolve() if cand.is_absolute() else (raiz / cand).resolve()
        if (cand / f"{art_id}_manuscrito.docx").exists():
            return cand

    # 1. ¿Está directamente en la carpeta raíz?
    if (raiz / f"{art_id}_manuscrito.docx").exists():
        return raiz

    # 2. Buscar subcarpetas tipo volumen19_n3_54854, v19n03_54854 o 54854
    for sub in sorted(raiz.iterdir()):
        if not sub.is_dir() or sub.name.startswith("."):
            continue
        if sub.name == art_id or sub.name.endswith(f"_{art_id}") or sub.name.endswith(f"-{art_id}"):
            if (sub / f"{art_id}_manuscrito.docx").exists():
                return sub.resolve()

    # 3. Cualquier subcarpeta directa que contenga {art_id}_manuscrito.docx
    for sub in sorted(raiz.iterdir()):
        if sub.is_dir() and not sub.name.startswith("."):
            if (sub / f"{art_id}_manuscrito.docx").exists():
                return sub.resolve()

    raise FileNotFoundError(
        f"No se encontró '{art_id}_manuscrito.docx' en {raiz} ni en subcarpetas tipo '*_{art_id}'."
    )


def procesar_fasciculo(carpeta_str: str = ".", formato_cli: str | None = None):
    raiz = Path(carpeta_str).resolve()
    cfg = cargar_config(raiz / "fasciculo", obligatorio=True)

    volumen = str(cfg.get("volumen", "")).strip()
    numero = str(cfg.get("numero", "")).strip()
    anio = str(cfg.get("anio", "")).strip()
    formato_fasc = formato_cli or (str(cfg.get("formato", "")).strip().lower() or None)
    cursor_pagina = int(cfg.get("pagina_inicial_numero", 1))
    actualizar_yaml = bool(cfg.get("actualizar_metadatos_yaml", True))

    articulos = cfg.get("articulos", [])
    if not articulos:
        raise ValueError(f"No hay artículos definidos en {raiz / 'fasciculo.yaml'}")

    print(f"=== Compilando fascículo Vol. {volumen} N.º {numero} ({anio}) ===")
    if formato_fasc:
        print(f"Perfil de diseño aplicado: {formato_fasc.upper()}")
    print(f"Página inicial del bloque de artículos: {cursor_pagina}\n")

    carpeta_salida_global = raiz / f"salida_v{volumen}n{numero}"
    carpeta_salida_global.mkdir(parents=True, exist_ok=True)

    resumen = []
    errores = []

    for item in articulos:
        if isinstance(item, dict):
            art_id = str(item.get("id") or item.get("id_ojs")).strip()
            carp_exp = str(item.get("carpeta", "")).strip()
            overrides = {
                "codigo_publicacion": item.get("codigo_publicacion", ""),
                "seccion": item.get("seccion", ""),
                "volumen": volumen,
                "numero": numero,
                "anio": anio,
            }
        else:
            art_id = str(item).strip()
            carp_exp = ""
            overrides = {
                "volumen": volumen,
                "numero": numero,
                "anio": anio,
            }

        try:
            carpeta_art = _resolver_carpeta_articulo(raiz, art_id, carp_exp)
            info_art = compilar_articulo(
                carpeta=carpeta_art,
                art_id=art_id,
                pagina_inicio_override=cursor_pagina,
                meta_overrides=overrides,
                actualizar_yaml=actualizar_yaml,
                formato_override=formato_fasc,
            )
            resumen.append(info_art)
            cursor_pagina = info_art["pagina_fin"] + 1

            # Copiar los 3 entregables finales a la carpeta consolidada del número
            nom = info_art["nombre_salida"]
            for ext in (".pdf", ".html", ".xml"):
                src = carpeta_art / "salida" / f"{nom}{ext}"
                if src.exists():
                    shutil.copy2(src, carpeta_salida_global / f"{nom}{ext}")
        except Exception as exc:
            print(f"\n[ERROR] Falló el artículo {art_id}: {exc}\n")
            errores.append((art_id, str(exc)))
        print()

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
    ruta_resumen = carpeta_salida_global / f"paginacion_v{volumen}n{numero}.txt"
    ruta_resumen.write_text(texto_resumen, encoding="utf-8")

    print(texto_resumen)
    print(f"Entregables consolidados y resumen en: {carpeta_salida_global}")
    if errores:
        print("\nATENCIÓN — Artículos con error (revisar y volver a correr):")
        for eid, msg in errores:
            print(f"  - ID {eid}: {msg}")


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
