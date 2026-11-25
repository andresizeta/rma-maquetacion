#!/usr/bin/env python3
"""Asistente interactivo completo para configurar metadatos, ORCID/ROR, recursos y compilar artículos RMA."""
from __future__ import annotations

import re
from pathlib import Path
import yaml

SECCIONES = ['Arqueología', 'Antropología Biológica', 'Antropología Social', 'Museología', 'Dossier']

EXT_IMAGEN = (".png", ".jpg", ".jpeg", ".tif", ".tiff", ".webp")
EXT_TABLA = (".docx", ".xlsx", ".csv")


def _pedir(prompt: str, defecto: str = "") -> str:
    suf = f" [{defecto}]" if str(defecto).strip() != "" else ""
    val = input(f"{prompt}{suf}: ").strip()
    return val if val != "" else str(defecto)


def _sugerir_autor_corto(autores: list[dict]) -> str:
    nombres = [str(a.get("nombre", "")).strip() for a in autores if str(a.get("nombre", "")).strip()]
    if not nombres:
        return ""

    def _abreviar(full: str) -> str:
        partes = [p for p in full.split() if p]
        if len(partes) == 1:
            return partes[0]
        ini = f"{partes[0][0].upper()}."
        apellido = " ".join(partes[1:])
        return f"{ini} {apellido}"

    if len(nombres) == 1:
        return _abreviar(nombres[0])
    if len(nombres) == 2:
        return f"{_abreviar(nombres[0])} y {_abreviar(nombres[1])}"
    return f"{_abreviar(nombres[0])} et al."


def _detectar_id_ojs(carpeta: Path) -> str:
    for p in sorted(carpeta.glob("*_manuscrito.docx")):
        if p.name.startswith(("~$", ".~")):
            continue
        return p.name[: -len("_manuscrito.docx")]
    return ""


def _defaults_desde_carpeta(nombre_carpeta: str, rev: dict) -> dict:
    vol = str(rev.get("volumen_actual", "19"))
    num = str(rev.get("numero_actual", "3"))
    anio = str(rev.get("anio_actual", "2026"))
    cod = ""
    m = re.search(r"volumen(\d+)[_-]n(\d+)[_-](a\d+)", nombre_carpeta, re.I)
    if m:
        vol = str(int(m.group(1)))
        num = str(int(m.group(2)))
        cod = f"v{int(vol):02d}n{int(num):02d}{m.group(3).lower()}"
    return {"volumen": vol, "numero": num, "anio": anio, "codigo_publicacion": cod}


def _yaml_dump_limpio(data: dict) -> str:
    return yaml.safe_dump(
        data,
        allow_unicode=True,
        sort_keys=False,
        default_flow_style=False,
        width=1000,
    )


def _buscar_candidatos_recurso(
    carpeta: Path,
    art_id: str,
    num: int,
    palabras: tuple[str, ...],
    extensiones: tuple[str, ...],
    subdirs: tuple[str, ...],
) -> list[str]:
    """Busca archivos candidatos para Figura N o Tabla N, con o sin prefijo art_id."""
    dirs = [carpeta]
    for sub in sorted(carpeta.iterdir(), key=lambda p: p.name.lower()):
        if sub.is_dir() and sub.name.lower() in subdirs:
            dirs.append(sub)

    pal_re = "|".join(re.escape(w) for w in palabras)
    sep = r"(?:[\s_\-\.\+]|%20)*"
    pat_con_id = re.compile(rf"^{re.escape(art_id)}{sep}(?:{pal_re}){sep}0*{num}(?!\d)", re.I)
    pat_sin_id = re.compile(rf"^(?:{pal_re}){sep}0*{num}(?!\d)", re.I)

    exactos: list[str] = []
    flexibles: list[str] = []

    for d in dirs:
        for f in sorted(d.iterdir(), key=lambda p: p.name.lower()):
            if not f.is_file() or f.name.startswith(("~$", ".~")):
                continue
            if f.suffix.lower() not in extensiones:
                continue
            stem = f.stem.strip()
            rel = f.name if d == carpeta else f"{d.name}/{f.name}"
            if pat_con_id.match(stem):
                exactos.append(f.name)
            elif pat_sin_id.match(stem):
                flexibles.append(f.name)

    vistos: set[str] = set()
    ordenados: list[str] = []
    for item in exactos + flexibles:
        if item not in vistos:
            vistos.add(item)
            ordenados.append(item)
    return ordenados


def _listar_archivos_por_ext(carpeta: Path, extensiones: tuple[str, ...], subdirs: tuple[str, ...]) -> list[str]:
    dirs = [carpeta]
    for sub in sorted(carpeta.iterdir(), key=lambda p: p.name.lower()):
        if sub.is_dir() and sub.name.lower() in subdirs:
            dirs.append(sub)
    vistos: set[str] = set()
    res: list[str] = []
    for d in dirs:
        for f in sorted(d.iterdir(), key=lambda p: p.name.lower()):
            if f.is_file() and not f.name.startswith(("~$", ".~")) and f.suffix.lower() in extensiones:
                if f.name not in vistos:
                    vistos.add(f.name)
                    res.append(f.name)
    return res


def ejecutar_flujo_interactivo(carpeta: str | Path, id_ojs: str = "") -> tuple[Path, str]:
    from maquetar_rma import (
        buscar_archivo_epigrafes,
        cargar_config,
        normalizar_orcid,
        normalizar_rors,
        parsear_autores_y_afiliaciones_docx,
        parsear_epigrafes_docx,
        parsear_manuscrito,
        sugerir_titulo_corto,
    )

    carpeta_path = Path(carpeta).resolve()
    carpeta_path.mkdir(parents=True, exist_ok=True)
    dir_script = Path(__file__).resolve().parent
    rev = cargar_config(dir_script / "revista", obligatorio=False)

    id_detectado = id_ojs.strip() or _detectar_id_ojs(carpeta_path)
    id_ojs = _pedir("ID de envío OJS", id_detectado)
    if not id_ojs:
        raise ValueError("Se requiere un ID de envío OJS.")

    ruta_meta = carpeta_path / f"{id_ojs}_metadatos.yaml"
    meta_prev = cargar_config(carpeta_path / f"{id_ojs}_metadatos", obligatorio=False) if ruta_meta.exists() else {}
    defs_carpeta = _defaults_desde_carpeta(carpeta_path.name, rev)

    ruta_ms = carpeta_path / f"{id_ojs}_manuscrito.docx"
    autores_docx: list[dict] = []
    afiliaciones_docx: dict[str, dict] = {}
    titulo_sugerido = ""
    bloques_ms: list[dict] = []

    if ruta_ms.exists():
        print(f"  -> Leyendo manuscrito: {ruta_ms.name} ...")
        extraido = parsear_autores_y_afiliaciones_docx(ruta_ms)
        autores_docx = extraido.get("autores", [])
        for idx, af in enumerate(extraido.get("afiliaciones", []), 1):
            clave = str(af.get("clave") or idx)
            afiliaciones_docx[clave] = {
                "texto": str(af.get("texto", "")).strip(),
                "ror": str(af.get("ror", "")).strip(),
            }
        try:
            art_tmp = parsear_manuscrito(ruta_ms)
            bloques_ms = art_tmp.get("bloques_cuerpo", [])
            titulo_plano = re.sub(r"<[^>]+>", "", art_tmp.get("titulo_es", "")).strip()
            titulo_sugerido = sugerir_titulo_corto(titulo_plano)
        except Exception:
            pass

    # --- 1. Datos editoriales ---
    print("\nSecciones disponibles:")
    for i, sec in enumerate(rev.get("secciones") or SECCIONES, 1):
        print(f"  {i}. {sec}")
    sec_def = str(meta_prev.get("seccion") or "Antropología Biológica")
    sel_sec = _pedir("Sección (número o nombre)", sec_def)
    if sel_sec.isdigit() and 1 <= int(sel_sec) <= len(rev.get("secciones") or SECCIONES):
        seccion = (rev.get("secciones") or SECCIONES)[int(sel_sec) - 1]
    else:
        seccion = sel_sec

    formato = _pedir("Formato (tradicional / moderno)", str(meta_prev.get("formato") or "tradicional"))
    volumen = _pedir("Volumen", str(meta_prev.get("volumen") or defs_carpeta["volumen"]))
    numero = _pedir("Número", str(meta_prev.get("numero") or defs_carpeta["numero"]))
    anio = _pedir("Año", str(meta_prev.get("anio") or defs_carpeta["anio"]))
    pagina_inicio = _pedir("Página inicial", str(meta_prev.get("pagina_inicio") or "01"))

    cod_prev = str(meta_prev.get("codigo_publicacion") or "").strip()
    cod_carpeta = defs_carpeta["codigo_publicacion"]
    cod_def = cod_carpeta if (cod_carpeta and cod_prev != cod_carpeta) else (cod_prev or cod_carpeta or f"v{int(volumen):02d}n{int(numero):02d}a01")
    codigo_pub = _pedir("Código de publicación", cod_def)
    doi = _pedir("DOI", str(meta_prev.get("doi") or "https://doi.org/10.31048/"))

    # --- 2. Autores y ORCID ---
    autores_prev = meta_prev.get("autores") or []
    tiene_autores_prev = any(str(a.get("nombre", "")).strip() for a in autores_prev if isinstance(a, dict))
    autores_base = autores_prev if tiene_autores_prev else autores_docx

    print(f"\n--- Autores y ORCID ({len(autores_base)} detectado/s) ---")
    autores_final: list[dict] = []
    idx_aut = 1
    while True:
        base_a = autores_base[idx_aut - 1] if (idx_aut - 1) < len(autores_base) else {}
        def_nom = str(base_a.get("nombre", "")).strip()
        etiqueta_nom = f"Autor {idx_aut} - Nombre completo (Enter para terminar)" if not def_nom else f"Autor {idx_aut} - Nombre completo"
        nom = _pedir(etiqueta_nom, def_nom)
        if not nom:
            break
        def_orcid = str(base_a.get("orcid", "")).strip()
        raw_orcid = _pedir(f"  ORCID de {nom} (ID o URL, '-' para limpiar)", def_orcid)
        if raw_orcid == "-":
            orcid_norm = ""
        else:
            orcid_norm, _ = normalizar_orcid(raw_orcid)
        def_af = str(base_a.get("afiliacion", "1")).strip() or "1"
        af_num = _pedir(f"  N° de afiliación de {nom} (ej. 1 o 1, 2)", def_af)
        autores_final.append({
            "nombre": nom,
            "orcid": orcid_norm,
            "afiliacion": af_num,
        })
        idx_aut += 1
        if idx_aut > len(autores_base):
            mas = _pedir("¿Agregar otro autor? [s/N]", "N")
            if mas.lower() not in ("s", "si", "sí"):
                break

    if not autores_final:
        autores_final = [{"nombre": "", "orcid": "", "afiliacion": "1"}]

    # --- 3. Afiliaciones y ROR ---
    af_prev = meta_prev.get("afiliaciones") or {}
    tiene_af_prev = isinstance(af_prev, dict) and any(
        str(v.get("texto", "")).strip() for v in af_prev.values() if isinstance(v, dict)
    )
    af_base = af_prev if tiene_af_prev else afiliaciones_docx

    claves_necesarias: list[str] = []
    for a in autores_final:
        for tok in str(a.get("afiliacion", "")).split(","):
            t = tok.strip()
            if t and t not in claves_necesarias:
                claves_necesarias.append(t)
    for k in af_base.keys():
        ks = str(k).strip()
        if ks and ks not in claves_necesarias:
            claves_necesarias.append(ks)
    if not claves_necesarias:
        claves_necesarias = ["1"]

    print(f"\n--- Afiliaciones y ROR ({len(claves_necesarias)} afiliación/es) ---")
    afiliaciones_final: dict[str, dict] = {}
    for clave in claves_necesarias:
        info_af = af_base.get(clave) or af_base.get(int(clave)) if clave.isdigit() else af_base.get(clave)
        info_af = info_af if isinstance(info_af, dict) else {}
        def_txt = str(info_af.get("texto", "")).strip()
        txt_af = _pedir(f"Afiliación [{clave}] - Institución / país / email", def_txt)
        def_ror = str(info_af.get("ror", "")).strip()
        raw_ror = _pedir(f"  ROR de afiliación [{clave}] (ID o URL, ej. 03x524c58, '-' para limpiar)", def_ror)
        if raw_ror == "-":
            ror_norm = ""
        else:
            pares_ror = normalizar_rors(raw_ror)
            ror_norm = ", ".join(r[0] for r in pares_ror) if pares_ror else ""
        afiliaciones_final[str(clave)] = {
            "texto": txt_af,
            "ror": ror_norm,
        }

    # --- 4. Encabezados y fechas ---
    autor_corto_def = _sugerir_autor_corto(autores_final) or str(meta_prev.get("autor_corto") or "")
    autor_corto = _pedir("\nAutor corto (encabezado)", autor_corto_def)
    titulo_corto_def = str(meta_prev.get("titulo_corto") or titulo_sugerido)
    titulo_corto = _pedir("Título corto (encabezado)", titulo_corto_def)

    recibido = _pedir("Fecha recibido (DD-MM-AAAA)", str(meta_prev.get("recibido") or ""))
    correcciones = _pedir("Fecha correcciones (DD-MM-AAAA)", str(meta_prev.get("correcciones") or ""))
    aceptado = _pedir("Fecha aceptado (DD-MM-AAAA)", str(meta_prev.get("aceptado") or ""))
    editor_sec = _pedir("Editor/a de sección", str(meta_prev.get("editor_seccion") or ""))
    editores_inv = _pedir("Editores invitados (si aplica)", str(meta_prev.get("editores_invitados") or ""))

    meta_out = {
        "formato": formato,
        "codigo_publicacion": codigo_pub,
        "seccion": seccion,
        "volumen": str(volumen),
        "numero": str(numero),
        "anio": str(anio),
        "pagina_inicio": str(pagina_inicio),
        "paginas": str(meta_prev.get("paginas") or ""),
        "doi": doi,
        "autor_corto": autor_corto,
        "titulo_corto": titulo_corto,
        "recibido": recibido,
        "correcciones": correcciones,
        "aceptado": aceptado,
        "editor_seccion": editor_sec,
        "editores_invitados": editores_inv,
        "autores": autores_final,
        "afiliaciones": afiliaciones_final,
    }
    ruta_meta.write_text("# Metadatos editoriales del envío OJS\n" + _yaml_dump_limpio(meta_out), encoding="utf-8")
    print(f"  [OK] Guardado: {ruta_meta.name}")

    # --- 5. Recursos: figuras, tablas y epígrafes ---
    ruta_rec = carpeta_path / f"{id_ojs}_recursos.yaml"
    rec_prev = cargar_config(carpeta_path / f"{id_ojs}_recursos", obligatorio=False) if ruta_rec.exists() else {}
    figs_out: dict[int, dict] = {}
    for k, v in (rec_prev.get("figuras") or {}).items():
        if str(k).isdigit():
            figs_out[int(k)] = dict(v or {})
    tabs_out: dict[int, dict] = {}
    for k, v in (rec_prev.get("tablas") or {}).items():
        if str(k).isdigit():
            tabs_out[int(k)] = dict(v or {})

    ruta_epi = buscar_archivo_epigrafes(carpeta_path, id_ojs)
    if ruta_epi and ruta_epi.exists():
        print(f"  -> Leyendo epígrafes desde: {ruta_epi.name}")
        epi = parsear_epigrafes_docx(ruta_epi)
        for num_f, datos_f in (epi.get("figuras") or {}).items():
            nf = int(num_f)
            prev_f = dict(figs_out.get(nf) or {})
            prev_f.setdefault("ancho", "columna")
            if datos_f.get("es"):
                prev_f["es"] = datos_f["es"]
            if datos_f.get("en"):
                prev_f["en"] = datos_f["en"]
            figs_out[nf] = prev_f
        for num_t, datos_t in (epi.get("tablas") or {}).items():
            nt = int(num_t)
            prev_t = dict(tabs_out.get(nt) or {})
            if datos_t.get("es"):
                prev_t["es"] = datos_t["es"]
            if datos_t.get("en"):
                prev_t["en"] = datos_t["en"]
            if datos_t.get("notas"):
                prev_t["notas"] = datos_t["notas"]
            tabs_out[nt] = prev_t

    nums_fig_ms = [int(b["num"]) for b in bloques_ms if b.get("tipo") == "FIGURA"]
    nums_tab_ms = [int(b["num"]) for b in bloques_ms if b.get("tipo") == "TABLA"]
    todas_figs = sorted(set(figs_out.keys()) | set(nums_fig_ms))
    todas_tabs = sorted(set(tabs_out.keys()) | set(nums_tab_ms))

    imgs_disponibles = _listar_archivos_por_ext(carpeta_path, EXT_IMAGEN, ("figuras", "figura", "figures", "imagenes", "imágenes"))
    tabs_disponibles = _listar_archivos_por_ext(carpeta_path, EXT_TABLA, ("tablas", "tabla", "tables", "cuadros"))

    if todas_figs:
        print(f"\n--- Vinculación de Figuras ({len(todas_figs)}) ---")
        if imgs_disponibles:
            print("  Imágenes en carpeta:", ", ".join(imgs_disponibles))
        for nf in todas_figs:
            entry_f = dict(figs_out.get(nf) or {})
            cands = _buscar_candidatos_recurso(
                carpeta_path, id_ojs, nf, ("figura", "fig", "figure"), EXT_IMAGEN,
                ("figuras", "figura", "figures", "imagenes", "imágenes"),
            )
            def_arch = str(entry_f.get("archivo") or (cands[0] if cands else "")).strip()
            arch_sel = _pedir(f"  Figura {nf} - Archivo de imagen", def_arch)
            if arch_sel:
                entry_f["archivo"] = arch_sel
            ancho_sel = _pedir(f"  Figura {nf} - Ancho (columna / completo)", str(entry_f.get("ancho") or "columna"))
            entry_f["ancho"] = ancho_sel
            figs_out[nf] = entry_f

    if todas_tabs:
        print(f"\n--- Vinculación de Tablas ({len(todas_tabs)}) ---")
        if tabs_disponibles:
            print("  Archivos de tabla en carpeta:", ", ".join(tabs_disponibles))
        for nt in todas_tabs:
            entry_t = dict(tabs_out.get(nt) or {})
            cands_t = _buscar_candidatos_recurso(
                carpeta_path, id_ojs, nt, ("tabla", "tab", "table", "cuadro"), EXT_TABLA,
                ("tablas", "tabla", "tables", "cuadros"),
            )
            def_arch_t = str(entry_t.get("archivo") or (cands_t[0] if cands_t else "")).strip()
            arch_sel_t = _pedir(f"  Tabla {nt} - Archivo (.docx/.xlsx/.csv)", def_arch_t)
            if arch_sel_t:
                entry_t["archivo"] = arch_sel_t
            tabs_out[nt] = entry_t

    rec_out = {"figuras": figs_out, "tablas": tabs_out}
    ruta_rec.write_text(_yaml_dump_limpio(rec_out), encoding="utf-8")
    print(f"  [OK] Guardado: {ruta_rec.name} ({len(figs_out)} figuras, {len(tabs_out)} tablas)")

    return carpeta_path, id_ojs


def main():
    from maquetar_rma import compilar_articulo

    base = Path(".").resolve()
    carpetas = sorted(
        [
            d for d in base.iterdir()
            if d.is_dir() and not d.name.startswith((".", "_")) and (list(d.glob("*_manuscrito.docx")) or d.name.startswith("volumen"))
        ],
        key=lambda p: p.name.lower(),
    )

    print("=== Asistente de Maquetación RMA ===")
    if carpetas:
        print("\nCarpetas de artículos disponibles:")
        for idx, c in enumerate(carpetas, 1):
            ms = _detectar_id_ojs(c)
            etiqueta = f" (ID: {ms})" if ms else ""
            print(f"  {idx:2d}. {c.name}{etiqueta}")
        sel = _pedir("\nElegí número de carpeta o escribí la ruta", "1")
        if sel.isdigit() and 1 <= int(sel) <= len(carpetas):
            carpeta_sel = carpetas[int(sel) - 1]
        else:
            carpeta_sel = Path(sel).resolve()
    else:
        carpeta_sel = Path(_pedir("Carpeta del artículo", ".")).resolve()

    carpeta_obj, id_ojs = ejecutar_flujo_interactivo(carpeta_sel, "")
    compilar_ahora = _pedir("\n¿Compilar ahora PDF + HTML OJS + XML SciELO? [S/n]", "S")
    if compilar_ahora.lower() not in ("n", "no"):
        compilar_articulo(carpeta_obj, id_ojs)


if __name__ == "__main__":
    main()
