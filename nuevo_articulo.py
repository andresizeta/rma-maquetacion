#!/usr/bin/env python3
"""
Asistente Interactivo de Artículos OJS - Revista del Museo de Antropología (RMA)

Detecta automáticamente carpetas de artículos creadas ad-hoc (ej. volumen19_n3_54854),
identifica archivos Word (.docx), tablas separadas (.docx / .xlsx / .csv) e imágenes
con cualquier nombre, pregunta qué es cada archivo, los estandariza con el ID de OJS
y compila PDF + HTML (OJS) + XML (SciELO).

Modos de uso:
  python asistente.py
  python nuevo_articulo.py
  python nuevo_articulo.py 54854 volumen19_n3_54854
"""

import re
import shutil
import sys
from pathlib import Path
from docx import Document
from maquetar_rma import (
    _parsear_yaml_basico,
    buscar_archivo_epigrafes,
    es_archivo_epigrafes,
    normalizar_orcid,
    normalizar_ror,
    parsear_autores_y_afiliaciones_docx,
    parsear_epigrafes_docx,
)

try:
    import yaml
except ImportError:
    yaml = None

SECCIONES_PREDEFINIDAS = [
    "Arqueología",
    "Antropología Social",
    "Bioantropología",
    "Dossier",
    "Editorial",
    "Foro",
    "Notas",
    "Reseñas",
]

EXTS_IMAGEN = (".png", ".jpg", ".jpeg", ".tif", ".tiff", ".webp")
EXTS_TABLA = (".docx", ".xlsx", ".csv")
CARPETAS_IGNORADAS = {
    ".venv",
    "venv",
    "env",
    "__pycache__",
    "assets",
    "ejemplo_ojs",
    "ejemplo_fasciculo",
    "salida",
    "figuras",
    "figura",
    "figures",
    "imagenes",
    "imágenes",
    "tablas",
    "tabla",
    "tables",
    "cuadros",
    "epigrafes",
    "epígrafes",
    "epigrafe",
    "leyendas",
    "captions",
}


def _natural_key(texto: str):
    return [int(tok) if tok.isdigit() else tok.lower() for tok in re.split(r"(\d+)", texto)]


def _cargar_yaml_simple(ruta: Path) -> dict:
    if not ruta.exists():
        return {}
    txt = ruta.read_text(encoding="utf-8")
    if yaml is not None:
        return yaml.safe_load(txt) or {}
    return _parsear_yaml_basico(txt)


def _escapar_yaml(valor: str) -> str:
    return str(valor).replace("\\", "\\\\").replace('"', '\\"')


def preguntar(etiqueta: str, defecto: str = "", obligatorio: bool = False) -> str:
    while True:
        sufijo = f" [{defecto}]" if defecto != "" else ""
        try:
            resp = input(f"  {etiqueta}{sufijo}: ").strip()
        except EOFError:
            resp = ""
        if resp == "":
            if defecto != "" or not obligatorio:
                return defecto
            print("    ! Este dato es obligatorio. Ingresá un valor.")
            continue
        if resp == "-":
            return ""
        return resp


def preguntar_si_no(etiqueta: str, defecto_si: bool = True) -> bool:
    hint = "S/n" if defecto_si else "s/N"
    try:
        resp = input(f"\n{etiqueta} [{hint}]: ").strip().lower()
    except EOFError:
        resp = ""
    if not resp:
        return defecto_si
    return resp in ("s", "si", "sí", "y", "yes")


def elegir_seccion(defecto: str = "Arqueología") -> str:
    print("\n  Secciones disponibles:")
    for idx, sec in enumerate(SECCIONES_PREDEFINIDAS, start=1):
        print(f"    {idx}. {sec}")
    resp = preguntar("Sección (número o nombre)", defecto=defecto)
    if resp.isdigit():
        n = int(resp)
        if 1 <= n <= len(SECCIONES_PREDEFINIDAS):
            return SECCIONES_PREDEFINIDAS[n - 1]
    return resp


def elegir_formato(defecto: str = "tradicional") -> str:
    print("\n  Perfil de maquetación:")
    print("    1. tradicional (2 columnas, márgenes angostos 20 mm — número actual)")
    print("    2. moderno     (1 columna, bloques pastel — próximo año)")
    def_num = "2" if defecto == "moderno" else "1"
    resp = preguntar("Formato (1 o 2)", defecto=def_num).lower()
    if resp in ("2", "moderno", "1col"):
        return "moderno"
    return "tradicional"


def detectar_carpetas_articulos(raiz: Path) -> list[dict]:
    """Escanea subcarpetas creadas por el usuario (ej. volumen19_n3_54854) en 'raiz'."""
    encontradas = []
    if not raiz.exists():
        return encontradas
    for sub in sorted(raiz.iterdir(), key=lambda p: _natural_key(p.name)):
        if not sub.is_dir() or sub.name.startswith(".") or sub.name in CARPETAS_IGNORADAS:
            continue
        n_docx = len([p for p in sub.glob("*.docx") if not p.name.startswith("~$")])
        n_imgs = len([p for p in sub.iterdir() if p.is_file() and p.suffix.lower() in EXTS_IMAGEN])
        for sub_fig in sub.iterdir():
            if sub_fig.is_dir() and sub_fig.name.lower() in ("figuras", "figura", "figures", "imagenes", "imágenes"):
                n_imgs += len([p for p in sub_fig.iterdir() if p.is_file() and p.suffix.lower() in EXTS_IMAGEN])
        tiene_salida = (sub / "salida").is_dir()
        encontradas.append({
            "path": sub,
            "nombre": sub.name,
            "n_docx": n_docx,
            "n_imgs": n_imgs,
            "tiene_salida": tiene_salida,
        })
    return encontradas


def extraer_datos_de_nombre_carpeta(carpeta: Path) -> dict:
    """
    Si la carpeta se llama por ej. 'volumen19_n3_54854' o 'v19n03_54854' o '54854',
    deduce automáticamente id_ojs='54854', volumen='19', numero='3'.
    """
    nombre = carpeta.name
    datos = {"id_ojs": "", "volumen": "", "numero": ""}

    m_vol = re.search(r"(?:volumen|vol|v)[\s_-]*(\d{1,3})", nombre, re.I)
    if m_vol:
        datos["volumen"] = str(int(m_vol.group(1)))

    m_num = re.search(r"(?<![a-zA-Z])(?:numero|num|n)[\s_-]*(\d{1,2})", nombre, re.I)
    if m_num:
        datos["numero"] = str(int(m_num.group(1)))

    m_id_end = re.search(r"(?:^|[_\-\s])(\d{4,7})$", nombre)
    if m_id_end:
        datos["id_ojs"] = m_id_end.group(1)
    else:
        nums_largos = re.findall(r"\d{4,7}", nombre)
        for n in nums_largos:
            if not (n.startswith("20") and len(n) == 4):
                datos["id_ojs"] = n
                break

    return datos


def _resumen_docx(ruta_docx: Path) -> tuple[int, int]:
    """Devuelve (cantidad_parrafos_no_vacios, cantidad_tablas) de un archivo Word."""
    try:
        doc = Document(str(ruta_docx))
        n_paras = sum(1 for p in doc.paragraphs if p.text.strip())
        n_tabs = len(doc.tables)
        return n_paras, n_tabs
    except Exception:
        return 0, 0


def _seleccionar_de_lista(respuesta: str, candidatos: list[Path]) -> Path | None:
    r = respuesta.strip()
    if not r or r == "-":
        return None
    if r.isdigit():
        idx = int(r) - 1
        if 0 <= idx < len(candidatos):
            return candidatos[idx]
    for c in candidatos:
        if c.name.lower() == r.lower() or c.stem.lower() == r.lower():
            return c
    return None


def detectar_y_estandarizar_manuscrito(destino: Path, art_id: str) -> tuple[Path, Path, Path | None]:
    """
    Detecta los archivos .docx en 'destino' (con cualquier nombre: texto.docx, paper.docx, etc.),
    pregunta cuál es el manuscrito principal y lo estandariza como {id}_manuscrito.docx.
    Devuelve (destino_final, ruta_manuscrito_std, ruta_archivo_origen_manuscrito).
    """
    destino.mkdir(parents=True, exist_ok=True)
    ruta_ms_std = destino / f"{art_id}_manuscrito.docx"

    def _listar_docx_sueltos(carpeta_b: Path) -> list[Path]:
        return sorted(
            [
                p
                for p in carpeta_b.glob("*.docx")
                if not p.name.startswith("~$")
                and not re.match(r"^\d+_(manuscrito|tablas|tabla\d+|epigrafes)\.docx$", p.name, re.I)
                and not es_archivo_epigrafes(p.name, art_id)
            ],
            key=lambda p: _natural_key(p.name),
        )

    todos_docx = _listar_docx_sueltos(destino)

    if ruta_ms_std.exists() and not todos_docx:
        print(f"  Manuscrito ya estandarizado detectado: {ruta_ms_std.name}")
        return destino, ruta_ms_std, None

    while not ruta_ms_std.exists() and not todos_docx:
        print(f"\n  [Atención] No encontré ningún archivo .docx en: {destino}")
        subcarpetas = detectar_carpetas_articulos(destino)
        if subcarpetas:
            print("  Sin embargo, detecté estas subcarpetas aquí dentro:")
            for idx, sc in enumerate(subcarpetas, start=1):
                print(f"    {idx}. {sc['nombre']}  ({sc['n_docx']} .docx, {sc['n_imgs']} imágenes)")
            resp_alt = preguntar(
                "Ingresá el número o nombre de la subcarpeta (o el nombre del .docx si lo copiaste aquí)",
                defecto="1" if len(subcarpetas) == 1 else "",
            )
            if resp_alt.isdigit() and 1 <= int(resp_alt) <= len(subcarpetas):
                destino = subcarpetas[int(resp_alt) - 1]["path"].resolve()
                ruta_ms_std = destino / f"{art_id}_manuscrito.docx"
                todos_docx = _listar_docx_sueltos(destino)
                print(f"  -> Cambiando carpeta de trabajo a: {destino}")
                break
            elif resp_alt and (destino / resp_alt).is_dir():
                destino = (destino / resp_alt).resolve()
                ruta_ms_std = destino / f"{art_id}_manuscrito.docx"
                todos_docx = _listar_docx_sueltos(destino)
                print(f"  -> Cambiando carpeta de trabajo a: {destino}")
                break
            elif resp_alt and Path(resp_alt).expanduser().is_dir():
                destino = Path(resp_alt).expanduser().resolve()
                ruta_ms_std = destino / f"{art_id}_manuscrito.docx"
                todos_docx = _listar_docx_sueltos(destino)
                print(f"  -> Cambiando carpeta de trabajo a: {destino}")
                break
        else:
            print("  Copiá el archivo Word del artículo en esa carpeta (puede llamarse paper.docx, texto.docx, etc.)")
            resp_alt = preguntar("Nombre del archivo .docx (o ruta a la carpeta donde está)", defecto="")

        if resp_alt:
            cand_dir = Path(resp_alt).expanduser()
            if not cand_dir.is_absolute():
                cand_dir = (destino / resp_alt).resolve()
            if cand_dir.is_dir():
                destino = cand_dir
                ruta_ms_std = destino / f"{art_id}_manuscrito.docx"
                todos_docx = _listar_docx_sueltos(destino)
                continue
            cand_file = destino / resp_alt
            if cand_file.exists() and cand_file.suffix.lower() == ".docx":
                shutil.copy2(cand_file, ruta_ms_std)
                print(f"  Estandarizado como: {ruta_ms_std.name}")
                return destino, ruta_ms_std, cand_file
            print(f"    ! No encuentro '{resp_alt}' en {destino}.")
        else:
            todos_docx = _listar_docx_sueltos(destino)

    if ruta_ms_std.exists() and not todos_docx:
        return destino, ruta_ms_std, None

    print("\n  --- Detección de archivos Word (.docx) en la carpeta ---")
    info_docx = []
    for idx, p in enumerate(todos_docx, start=1):
        np, nt = _resumen_docx(p)
        info_docx.append((p, np, nt))
        print(f"    {idx}. {p.name}  ({np} párrafos, {nt} tablas)")

    origen_ms = None
    if not ruta_ms_std.exists():
        if len(info_docx) == 1:
            unico = info_docx[0][0]
            if preguntar_si_no(f"¿'{unico.name}' es el MANUSCRITO (texto principal) del envío #{art_id}?", defecto_si=True):
                shutil.copy2(unico, ruta_ms_std)
                origen_ms = unico
                print(f"  -> Guardado estándar: {ruta_ms_std.name}")
        else:
            idx_sug_ms = max(range(len(info_docx)), key=lambda i: info_docx[i][1]) + 1
            while not ruta_ms_std.exists():
                resp_ms = preguntar(
                    "¿Cuál archivo es el MANUSCRITO (texto principal)? (número o nombre)",
                    defecto=str(idx_sug_ms),
                    obligatorio=True,
                )
                elegido_ms = _seleccionar_de_lista(resp_ms, todos_docx)
                if elegido_ms is not None:
                    shutil.copy2(elegido_ms, ruta_ms_std)
                    origen_ms = elegido_ms
                    print(f"  -> Manuscrito estandarizado: {ruta_ms_std.name}")
                else:
                    print("    ! Opción no válida. Elegí un número de la lista.")

    return destino, ruta_ms_std, origen_ms


def detectar_y_estandarizar_tablas(
    destino: Path,
    art_id: str,
    tabs_en_texto: int,
    tabs_embebidas: int,
    origen_ms: Path | None,
    formato: str,
    epigrafes_tablas: dict | None = None,
) -> dict:
    """
    Detecta archivos de tablas (.docx, .xlsx, .csv) en 'destino' o 'destino/tablas',
    soportando tanto:
      - Un único archivo con todas las tablas (estandarizado como {id}_tablas.docx)
      - Cada tabla en un archivo distinto (estandarizado como {id}_tabla01.docx / .xlsx / .csv)
      - Tablas ya embebidas dentro del manuscrito principal
    """
    carpeta_tab = destino / "tablas"

    # Revisar archivos ya estandarizados previamente
    ya_std_ind = {}
    auto_tabs = {}
    for carpeta_busq in _carpetas_recursos_ci(destino, ('tablas', 'tabla', 'tables', 'cuadros')):
        if not carpeta_busq.exists():
            continue
        for p in carpeta_busq.iterdir():
            m = re.match(rf"^{re.escape(art_id)}_tabla0*(\d+)\.(docx|xlsx|csv)$", p.name, re.I)
            if m:
                ya_std_ind[int(m.group(1))] = p

    # Buscar archivos candidatos a tablas (excluyendo el manuscrito y archivos ya estandarizados)
    candidatos = []
    for carpeta_busq in _carpetas_recursos_ci(destino, ('tablas', 'tabla', 'tables', 'cuadros')):
        if not carpeta_busq.exists():
            continue
        for p in sorted(carpeta_busq.iterdir(), key=lambda x: _natural_key(x.name)):
            if not p.is_file() or p.name.startswith("~$"):
                continue
            if p.suffix.lower() not in EXTS_TABLA:
                continue
            if origen_ms is not None and p.resolve() == origen_ms.resolve():
                continue
            if re.match(r"^\d+_(manuscrito|tablas|tabla\d+|epigrafes)\.", p.name, re.I):
                continue
            if p.suffix.lower() == ".docx" and es_archivo_epigrafes(p.name, art_id):
                continue
            num_t = _extraer_num_recurso_ojs(p.stem, art_id, ("tabla", "tab", "table", "cuadro"))
            if num_t is not None and num_t not in ya_std_ind:
                nombre_std = f"{art_id}_tabla{num_t:02d}{p.suffix.lower()}"
                shutil.copy2(p, destino / nombre_std)
                ya_std_ind[num_t] = destino / nombre_std
                auto_tabs[num_t] = p.name
            else:
                candidatos.append(p)
    candidatos = sorted(candidatos, key=lambda p: _natural_key(p.name))

    ruta_lote_std = destino / f"{art_id}_tablas.docx"
    tabs_en_lote = 0
    if ruta_lote_std.exists():
        _, tabs_en_lote = _resumen_docx(ruta_lote_std)

    epigrafes_tablas = epigrafes_tablas or {}
    total_objetivo = max(
        tabs_en_texto,
        tabs_embebidas,
        tabs_en_lote,
        len(ya_std_ind),
        max(epigrafes_tablas.keys(), default=0),
    )
    if total_objetivo == 0 and not candidatos:
        return {}

    print("\n  --- Detección y asignación de Tablas ---")
    if tabs_en_texto > 0:
        print(f"  El manuscrito menciona {tabs_en_texto} tabla(s) (TABLA 1 .. TABLA {tabs_en_texto}).")
    if tabs_embebidas > 0:
        print(f"  (El propio manuscrito ya contiene {tabs_embebidas} tabla(s) insertada(s) adentro).")
    if ya_std_ind:
        print("  Tablas individuales ya estandarizadas: " + ", ".join(ya_std_ind[k].name for k in sorted(ya_std_ind)))
    if candidatos:
        print("  Archivos externos de tablas disponibles:")
        for idx, p in enumerate(candidatos, start=1):
            if p.suffix.lower() == ".docx":
                np, nt = _resumen_docx(p)
                det = f"{nt} tabla(s), {np} párrafos"
            else:
                det = p.suffix.lower()
            ubic = "tablas/" if p.parent == carpeta_tab else "carpeta raíz"
            print(f"    {idx}. {p.name}  ({det} — {ubic})")

    if total_objetivo == 0 and candidatos:
        cant_raw = preguntar(
            "¿Cuántas tablas querés incorporar al artículo?",
            defecto=str(len(candidatos)),
        )
        total_objetivo = int(cant_raw) if cant_raw.isdigit() else 0

    if total_objetivo == 0:
        return {}

    if not preguntar_si_no(
        f"¿Querés asignar los archivos y títulos bilingües de las {total_objetivo} tabla(s) ahora?",
        defecto_si=True,
    ):
        return {}

    # Caso especial: 1 solo archivo externo que ya contiene todas las tablas juntas (ej. cuadros.docx con 3 tablas)
    usar_lote_unico = False
    if len(candidatos) == 1 and candidatos[0].suffix.lower() == ".docx" and total_objetivo > 1:
        _, nt_unico = _resumen_docx(candidatos[0])
        if nt_unico >= total_objetivo:
            if preguntar_si_no(
                f"¿Todas las {total_objetivo} tablas están juntas dentro de '{candidatos[0].name}'?",
                defecto_si=True,
            ):
                shutil.copy2(candidatos[0], ruta_lote_std)
                print(f"  -> Archivo conjunto guardado como: {ruta_lote_std.name}")
                usar_lote_unico = True

    if formato == "tradicional":
        prompt_ancho = "Ancho (1 = completo 2 columnas, 2 = 1 columna)"
    else:
        prompt_ancho = "Ancho (1 = completo 100%, 2 = centrado 74%)"

    tablas_cfg = {}
    usados = set()

    for num in range(1, total_objetivo + 1):
        print(f"\n  --- Tabla {num} ---")
        archivo_asignado = ""

        if num in ya_std_ind:
            archivo_asignado = ya_std_ind[num].name
            print(f"  Archivo ya asignado: {archivo_asignado}")
        elif not usar_lote_unico and candidatos:
            sug_idx = ""
            for idx, c in enumerate(candidatos, start=1):
                if c not in usados:
                    sug_idx = str(idx)
                    break
            if tabs_embebidas >= num and not sug_idx:
                sug_idx = "-"
            resp_t = preguntar(
                f"Archivo para Tabla {num} (número de lista, nombre, o '-' si va embebida en el Word)",
                defecto=sug_idx,
            )
            elegido = _seleccionar_de_lista(resp_t, candidatos)
            if elegido is not None:
                usados.add(elegido)
                nombre_std = f"{art_id}_tabla{num:02d}{elegido.suffix.lower()}"
                shutil.copy2(elegido, destino / nombre_std)
                archivo_asignado = nombre_std
                print(f"  -> Copiada como: {nombre_std}")

        ancho_in = preguntar(prompt_ancho, defecto="1")
        if ancho_in in ("2", "columna", "centrado", "74%", "70%"):
            ancho = "columna" if formato == "tradicional" else "centrado"
        else:
            ancho = "completo"
        ep_t = epigrafes_tablas.get(num, {})
        if ep_t.get("es") or ep_t.get("en"):
            print(f"  Epígrafes precargados desde {art_id}_epigrafes.docx")
        es_def = ep_t.get("es") or f"Descripción de la Tabla {num}."
        en_def = ep_t.get("en") or f"Description of Table {num}."
        notas_def = ep_t.get("notas") or ""
        es = preguntar(f"Título/epígrafe en ESPAÑOL (Tabla {num})", defecto=es_def)
        en = preguntar(f"Título/epígrafe en INGLÉS (Table {num})", defecto=en_def)
        notas = preguntar(f"Notas al pie de la Tabla {num} (opcional, Enter para omitir)", defecto=notas_def)

        item_cfg = {"ancho": ancho, "es": es, "en": en, "notas": notas}
        if archivo_asignado:
            item_cfg["archivo"] = archivo_asignado
        tablas_cfg[num] = item_cfg

    return tablas_cfg


def _extraer_num_recurso_ojs(stem: str, art_id: str, palabras: tuple[str, ...]) -> int | None:
    """
    Si el nombre empieza con el ID de OJS y luego dice 'Figura+1', 'figura_1', 'fig 02', 'Tabla+1', etc.,
    devuelve el número entero de recurso (ej. 1, 2).
    """
    pal_re = "|".join(re.escape(w) for w in palabras)
    sep = r"(?:[\s_\-\.\+]|%20)*"
    m = re.match(rf"^{re.escape(art_id)}{sep}(?:{pal_re}){sep}0*(\d+)(?!\d)", stem.strip(), re.I)
    return int(m.group(1)) if m else None


def _carpetas_recursos_ci(destino: Path, nombres_sub: tuple[str, ...]) -> list[Path]:
    """Devuelve subcarpetas coincidentes (sin importar mayúsculas/minúsculas) más la carpeta raíz."""
    dirs = []
    if destino.exists():
        for p in sorted(destino.iterdir(), key=lambda x: x.name.lower()):
            if p.is_dir() and p.name.lower() in nombres_sub:
                dirs.append(p)
        dirs.append(destino)
    return dirs


def detectar_y_estandarizar_imagenes(
    destino: Path,
    art_id: str,
    figs_en_texto: int,
    formato: str,
    epigrafes_figuras: dict | None = None,
) -> dict:
    """
    Busca imágenes en 'destino', 'destino/figuras', 'destino/Figuras', etc.
    Reconoce automáticamente nombres estilo OJS como 54854_Figura+1.tiff o 54854_Figura+2.jpg,
    las estandariza a figuras/{id}_fig01.<ext> y recoge sus anchos y epígrafes bilingües.
    """
    carpeta_fig = destino / "figuras"
    carpeta_fig.mkdir(parents=True, exist_ok=True)
    dirs_busqueda = _carpetas_recursos_ci(destino, ("figuras", "figura", "figures", "imagenes", "imágenes"))

    ya_std = {}
    auto_detectadas = {}
    candidatas = []

    # 1) Revisar figuras ya estandarizadas en figuras/
    for p in sorted(carpeta_fig.iterdir(), key=lambda x: _natural_key(x.name)):
        if p.is_file() and p.suffix.lower() in EXTS_IMAGEN:
            m = re.match(rf"^{re.escape(art_id)}_fig0*(\d+)\.", p.name, re.I)
            if m:
                ya_std[int(m.group(1))] = p

    # 2) Revisar imágenes en figuras/, Figuras/ y la carpeta raíz del artículo
    for carpeta_busq in dirs_busqueda:
        if not carpeta_busq.exists():
            continue
        for p in sorted(carpeta_busq.iterdir(), key=lambda x: _natural_key(x.name)):
            if not p.is_file() or p.suffix.lower() not in EXTS_IMAGEN:
                continue
            if re.match(rf"^{re.escape(art_id)}_fig\d+\.", p.name, re.I):
                continue
            num_fig = _extraer_num_recurso_ojs(p.stem, art_id, ("figura", "fig", "figure"))
            if num_fig is not None and num_fig not in ya_std and num_fig not in auto_detectadas:
                nombre_std = f"{art_id}_fig{num_fig:02d}{p.suffix.lower()}"
                shutil.copy2(p, carpeta_fig / nombre_std)
                ya_std[num_fig] = carpeta_fig / nombre_std
                rel_origen = p.name if p.parent == destino else f"{p.parent.name}/{p.name}"
                auto_detectadas[num_fig] = rel_origen
            else:
                candidatas.append(p)

    epigrafes_figuras = epigrafes_figuras or {}
    total_objetivo = max(
        figs_en_texto,
        max(ya_std.keys(), default=0),
        max(epigrafes_figuras.keys(), default=0),
    )
    if total_objetivo == 0 and not candidatas:
        return {}

    print("\n  --- Detección y asignación de Imágenes / Figuras ---")
    if figs_en_texto > 0:
        print(f"  El manuscrito menciona {figs_en_texto} figura(s) (FIGURA 1 .. FIGURA {figs_en_texto}).")
    if auto_detectadas:
        print("  Figuras emparejadas automáticamente por código OJS:")
        for k in sorted(auto_detectadas):
            print(f"    - Figura {k}: {auto_detectadas[k]}  ->  figuras/{ya_std[k].name}")
    elif ya_std:
        print("  Ya estandarizadas en figuras/: " + ", ".join(ya_std[k].name for k in sorted(ya_std)))
    if candidatas:
        print("  Otras imágenes disponibles para asignar:")
        for idx, img_p in enumerate(candidatas, start=1):
            ubic = "figuras/" if img_p.parent == carpeta_fig else "carpeta raíz"
            print(f"    {idx}. {img_p.name}  ({ubic})")

    if total_objetivo == 0 and candidatas:
        cant_raw = preguntar(
            "¿Cuántas figuras querés incorporar al artículo?",
            defecto=str(len(candidatas)),
        )
        total_objetivo = int(cant_raw) if cant_raw.isdigit() else 0

    if total_objetivo == 0:
        return {}

    if not preguntar_si_no(
        f"¿Querés cargar ahora los anchos y epígrafes bilingües de las {total_objetivo} figura(s)?",
        defecto_si=True,
    ):
        return {}

    if formato == "tradicional":
        prompt_ancho = "Ancho (1 = completo 2 columnas, 2 = 1 columna)"
    else:
        prompt_ancho = "Ancho (1 = completo 100%, 2 = centrado 74%)"

    figuras_cfg = {}
    usadas = set()

    for num in range(1, total_objetivo + 1):
        print(f"\n  --- Figura {num} ---")
        if num in ya_std:
            orig_txt = f" (desde {auto_detectadas[num]})" if num in auto_detectadas else ""
            print(f"  Archivo asignado: figuras/{ya_std[num].name}{orig_txt}")
        elif candidatas:
            sug_idx = ""
            for idx, c in enumerate(candidatas, start=1):
                if c not in usadas:
                    sug_idx = str(idx)
                    break
            resp_img = preguntar(
                f"Archivo de imagen para Figura {num} (número de lista, nombre, o '-' si aún no la tenés)",
                defecto=sug_idx,
            )
            elegida = _seleccionar_de_lista(resp_img, candidatas)
            if elegida is not None:
                usadas.add(elegida)
                nombre_std = f"{art_id}_fig{num:02d}{elegida.suffix.lower()}"
                shutil.copy2(elegida, carpeta_fig / nombre_std)
                print(f"  -> Copiada como: figuras/{nombre_std}")
            else:
                print(f"  -> Sin archivo por ahora (se usará recuadro provisional hasta que agregues {art_id}_figura_{num}.png)")

        defecto_ancho = "1"
        if num in ya_std and ya_std[num].exists():
            try:
                from PIL import Image as _PILImage
                with _PILImage.open(ya_std[num]) as _im:
                    _w, _h = _im.size
                    if _h > 0:
                        _ratio = _w / _h
                        if _ratio < 1.28:
                            defecto_ancho = "2"
                            orient_txt = f"vertical/cuadrada ({_w}x{_h} px)"
                        else:
                            defecto_ancho = "1"
                            orient_txt = f"apaisada ({_w}x{_h} px)"
                        print(f"  Orientación detectada: {orient_txt} -> sugerido: {defecto_ancho}")
            except Exception:
                pass
        ancho_in = preguntar(prompt_ancho, defecto=defecto_ancho)
        if ancho_in in ("2", "columna", "centrado", "74%", "70%"):
            ancho = "columna" if formato == "tradicional" else "centrado"
        else:
            ancho = "completo"
        ep_f = epigrafes_figuras.get(num, {})
        if ep_f.get("es") or ep_f.get("en"):
            print(f"  Epígrafes precargados desde {art_id}_epigrafes.docx")
        es_def = ep_f.get("es") or f"Descripción de la Figura {num}."
        en_def = ep_f.get("en") or f"Description of Figure {num}."
        es = preguntar(f"Epígrafe en ESPAÑOL (Figura {num})", defecto=es_def)
        en = preguntar(f"Epígrafe en INGLÉS (Figure {num})", defecto=en_def)
        figuras_cfg[num] = {"ancho": ancho, "es": es, "en": en}

    return figuras_cfg


def inspeccionar_manuscrito_existente(ruta_docx: Path) -> dict:
    """Lee el Word del autor para detectar cantidad de figuras, tablas, título y autor."""
    info = {"figs": 0, "tabs_texto": 0, "tabs_embebidas": 0, "titulo_sugerido": "", "autor_sugerido": ""}
    if not ruta_docx.exists():
        return info
    try:
        doc = Document(str(ruta_docx))
        parrafos = [p.text.strip() for p in doc.paragraphs if p.text.strip() and p.text.strip() != "."]
        if parrafos:
            t0 = parrafos[0]
            info["titulo_sugerido"] = (t0[:42] + " ...") if len(t0) > 45 else t0
        if len(parrafos) >= 3:
            aut = re.sub(r"[\*\d]+", "", parrafos[2]).strip()
            partes = [a.strip() for a in re.split(r",|\sy\s", aut) if a.strip()]
            if len(partes) == 1:
                tokens = partes[0].split()
                info["autor_sugerido"] = f"{tokens[0][0]}. {' '.join(tokens[1:])}" if len(tokens) > 1 else partes[0]
            elif len(partes) == 2:
                t1, t2 = partes[0].split(), partes[1].split()
                a1 = f"{t1[0][0]}. {' '.join(t1[1:])}" if len(t1) > 1 else partes[0]
                a2 = f"{t2[0][0]}. {' '.join(t2[1:])}" if len(t2) > 1 else partes[1]
                info["autor_sugerido"] = f"{a1} y {a2}"
            elif len(partes) > 2:
                t1 = partes[0].split()
                a1 = f"{t1[0][0]}. {' '.join(t1[1:])}" if len(t1) > 1 else partes[0]
                info["autor_sugerido"] = f"{a1} et al."

        max_fig = 0
        max_tab = 0
        for p in doc.paragraphs:
            t = p.text.strip()
            mf = re.match(r"^FIGURA\s+(\d+)\s*$", t, re.I)
            if mf:
                max_fig = max(max_fig, int(mf.group(1)))
            mt = re.match(r"^TABLA\s+(\d+)\s*$", t, re.I)
            if mt:
                max_tab = max(max_tab, int(mt.group(1)))
        info["figs"] = max_fig
        info["tabs_texto"] = max_tab
        info["tabs_embebidas"] = len(doc.tables)
    except Exception:
        pass
    return info


def inferir_valores_contexto(destino: Path) -> dict:
    """Detecta formato, volumen, número, año, editor y siguiente página disponible."""
    dir_script = Path(__file__).resolve().parent
    rev = _cargar_yaml_simple(destino / "revista.yaml") or _cargar_yaml_simple(dir_script / "revista.yaml")
    info = {
        "formato": str(rev.get("formato") or "tradicional"),
        "volumen": "19",
        "numero": "2",
        "anio": "2026",
        "pagina_inicio": "01",
        "editor_seccion": "",
        "siguiente_orden": 1,
    }

    desde_nombre = extraer_datos_de_nombre_carpeta(destino)
    if desde_nombre["volumen"]:
        info["volumen"] = desde_nombre["volumen"]
    if desde_nombre["numero"]:
        info["numero"] = desde_nombre["numero"]

    fasc = _cargar_yaml_simple(destino / "fasciculo.yaml")
    if fasc:
        if fasc.get("formato"):
            info["formato"] = str(fasc["formato"])
        if fasc.get("volumen"):
            info["volumen"] = str(fasc["volumen"])
        if fasc.get("numero"):
            info["numero"] = str(fasc["numero"])
        if fasc.get("anio"):
            info["anio"] = str(fasc["anio"])
        if fasc.get("pagina_inicial_numero"):
            info["pagina_inicio"] = f"{int(fasc['pagina_inicial_numero']):02d}"

    metas_existentes = sorted(destino.glob("*_metadatos.yaml"))
    max_pag_fin = 0
    for m_path in metas_existentes:
        m = _cargar_yaml_simple(m_path)
        if m.get("volumen"):
            info["volumen"] = str(m["volumen"])
        if m.get("numero"):
            info["numero"] = str(m["numero"])
        if m.get("anio"):
            info["anio"] = str(m["anio"])
        if m.get("editor_seccion"):
            info["editor_seccion"] = str(m["editor_seccion"])
        p_fin = str(m.get("pagina_fin", "")).strip()
        if p_fin.isdigit():
            max_pag_fin = max(max_pag_fin, int(p_fin))
        else:
            pags = str(m.get("paginas", "")).strip()
            match = re.match(r"^\d+\s*[-–]\s*(\d+)$", pags)
            if match:
                max_pag_fin = max(max_pag_fin, int(match.group(1)))

    if max_pag_fin > 0:
        info["pagina_inicio"] = f"{max_pag_fin + 1:02d}"

    info["siguiente_orden"] = len(metas_existentes) + 1
    return info


def escribir_metadatos_yaml(ruta_yaml: Path, meta: dict):
    lineas = [
        "# Metadatos editoriales del envío OJS (generado con asistente interactivo RMA)",
        f'formato: "{_escapar_yaml(meta.get("formato", "tradicional"))}"',
        f'codigo_publicacion: "{_escapar_yaml(meta.get("codigo_publicacion", ""))}"',
        f'seccion: "{_escapar_yaml(meta.get("seccion", "Arqueología"))}"',
        f'volumen: "{_escapar_yaml(meta.get("volumen", ""))}"',
        f'numero: "{_escapar_yaml(meta.get("numero", ""))}"',
        f'anio: "{_escapar_yaml(meta.get("anio", ""))}"',
        f'pagina_inicio: "{_escapar_yaml(meta.get("pagina_inicio", "01"))}"',
        f'paginas: "{_escapar_yaml(meta.get("paginas", ""))}"',
        f'doi: "{_escapar_yaml(meta.get("doi", ""))}"',
        f'autor_corto: "{_escapar_yaml(meta.get("autor_corto", ""))}"',
        f'titulo_corto: "{_escapar_yaml(meta.get("titulo_corto", ""))}"',
        f'recibido: "{_escapar_yaml(meta.get("recibido", ""))}"',
        f'correcciones: "{_escapar_yaml(meta.get("correcciones", ""))}"',
        f'aceptado: "{_escapar_yaml(meta.get("aceptado", ""))}"',
        f'editor_seccion: "{_escapar_yaml(meta.get("editor_seccion", ""))}"',
        f'editores_invitados: "{_escapar_yaml(meta.get("editores_invitados", ""))}"',
    ]

    autores_list = meta.get("autores") or []
    if autores_list:
        lineas.append("")
        lineas.append("autores:")
        for aut in autores_list:
            if not isinstance(aut, dict):
                continue
            nom = _escapar_yaml(aut.get("nombre", ""))
            oid = _escapar_yaml(aut.get("orcid", ""))
            af = _escapar_yaml(aut.get("afiliacion", ""))
            lineas.append(f'  - nombre: "{nom}"')
            lineas.append(f'    orcid: "{oid}"')
            lineas.append(f'    afiliacion: "{af}"')

    afils_cfg = meta.get("afiliaciones") or {}
    if afils_cfg:
        lineas.append("")
        lineas.append("afiliaciones:")
        if isinstance(afils_cfg, dict):
            for clave_af, datos_af in afils_cfg.items():
                if isinstance(datos_af, dict):
                    txt_af = _escapar_yaml(datos_af.get("texto", ""))
                    ror_af = _escapar_yaml(datos_af.get("ror", ""))
                else:
                    txt_af = ""
                    ror_af = _escapar_yaml(datos_af)
                lineas.append(f"  {clave_af}:")
                if txt_af:
                    lineas.append(f'    texto: "{txt_af}"')
                lineas.append(f'    ror: "{ror_af}"')
        elif isinstance(afils_cfg, list):
            for idx_af, datos_af in enumerate(afils_cfg, start=1):
                if isinstance(datos_af, dict):
                    clave_af = datos_af.get("clave") or idx_af
                    txt_af = _escapar_yaml(datos_af.get("texto", ""))
                    ror_af = _escapar_yaml(datos_af.get("ror", ""))
                    lineas.append(f"  {clave_af}:")
                    if txt_af:
                        lineas.append(f'    texto: "{txt_af}"')
                    lineas.append(f'    ror: "{ror_af}"')

    ruta_yaml.write_text("\n".join(lineas) + "\n", encoding="utf-8")


def escribir_recursos_yaml(ruta_yaml: Path, figuras_cfg: dict, tablas_cfg: dict):
    lineas = ["figuras:"]
    if not figuras_cfg:
        lineas.append("  # Sin figuras registradas")
    for num in sorted(figuras_cfg):
        item = figuras_cfg[num]
        lineas.extend([
            f"  {num}:",
            f'    ancho: "{_escapar_yaml(item.get("ancho", "completo"))}"',
            f'    es: "{_escapar_yaml(item.get("es", f"Descripción de la Figura {num}."))}"',
            f'    en: "{_escapar_yaml(item.get("en", f"Description of Figure {num}."))}"',
        ])

    lineas.append("\ntablas:")
    if not tablas_cfg:
        lineas.append("  # Sin tablas registradas")
    for num in sorted(tablas_cfg):
        item = tablas_cfg[num]
        lineas.extend([
            f"  {num}:",
            f'    ancho: "{_escapar_yaml(item.get("ancho", "completo"))}"',
        ])
        if item.get("archivo"):
            lineas.append(f'    archivo: "{_escapar_yaml(item["archivo"])}"')
        lineas.extend([
            f'    es: "{_escapar_yaml(item.get("es", f"Descripción de la Tabla {num}."))}"',
            f'    en: "{_escapar_yaml(item.get("en", f"Description of Table {num}."))}"',
        ])
        if item.get("notas"):
            lineas.append(f'    notas: "{_escapar_yaml(item["notas"])}"')

    ruta_yaml.write_text("\n".join(lineas) + "\n", encoding="utf-8")
    print(f"  Recursos guardados en: {ruta_yaml.name}")


def registrar_en_fasciculo(destino: Path, art_id: str, cod_pub: str, seccion: str):
    ruta_fasc = destino / "fasciculo.yaml"
    if not ruta_fasc.exists():
        return
    txt = ruta_fasc.read_text(encoding="utf-8")
    if re.search(rf'id:\s*["\']?{re.escape(art_id)}["\']?\b', txt):
        return
    bloque = (
        f'  - id: "{_escapar_yaml(art_id)}"\n'
        f'    codigo_publicacion: "{_escapar_yaml(cod_pub)}"\n'
        f'    seccion: "{_escapar_yaml(seccion)}"\n'
    )
    if not txt.endswith("\n"):
        txt += "\n"
    ruta_fasc.write_text(txt + bloque, encoding="utf-8")
    print(f"  Agregado ID '{art_id}' al orden de {ruta_fasc.name}")


def ejecutar_flujo_interactivo(destino: Path, art_id: str) -> tuple[Path, dict]:
    """Flujo completo: detecta y estandariza manuscrito, imágenes y tablas (juntas o por separado)."""
    destino.mkdir(parents=True, exist_ok=True)
    (destino / "figuras").mkdir(exist_ok=True)

    print("\n" + "=" * 68)
    print(f"  ASISTENTE EDITORIAL RMA — ENVÍO OJS #{art_id}")
    print(f"  Carpeta de trabajo: {destino}")
    print("=" * 68)

    # 1. Detectar y estandarizar el manuscrito principal
    destino, ruta_ms, origen_ms = detectar_y_estandarizar_manuscrito(destino, art_id)
    ms_info = inspeccionar_manuscrito_existente(ruta_ms)

    # 2. Inferir contexto editorial
    ruta_meta = destino / f"{art_id}_metadatos.yaml"
    actual = _cargar_yaml_simple(ruta_meta)
    ctx = inferir_valores_contexto(destino)

    fmt_def = str(actual.get("formato") or ctx["formato"])
    vol_def = str(actual.get("volumen") or ctx["volumen"])
    num_def = str(actual.get("numero") or ctx["numero"])
    anio_def = str(actual.get("anio") or ctx["anio"])
    sec_def = str(actual.get("seccion") or "Arqueología")

    cod_def = str(actual.get("codigo_publicacion", "")).strip()
    if not cod_def and vol_def.isdigit() and num_def.isdigit():
        cod_def = f"v{int(vol_def)}n{int(num_def):02d}a{ctx['siguiente_orden']:02d}"

    p_ini_def = str(actual.get("pagina_inicio") or ctx["pagina_inicio"]).strip()
    if p_ini_def.isdigit():
        p_ini_def = f"{int(p_ini_def):02d}"

    autor_def = str(actual.get("autor_corto") or ms_info["autor_sugerido"])
    titulo_def = str(actual.get("titulo_corto") or ms_info["titulo_sugerido"])

    print("\n  --- Metadatos editoriales del artículo ---")
    print("  (Presioná Enter para aceptar el valor entre [corchetes], o '-' para dejar vacío)")

    formato = elegir_formato(defecto=fmt_def)
    seccion = elegir_seccion(defecto=sec_def)
    volumen = preguntar("Volumen", defecto=vol_def)
    numero = preguntar("Número", defecto=num_def)
    anio = preguntar("Año", defecto=anio_def)
    codigo_pub = preguntar("Código de publicación (ej. v19n02a02, '-' para vacío)", defecto=cod_def)
    pagina_inicio = preguntar("Página inicial (ej. 07 o 27)", defecto=p_ini_def)
    if pagina_inicio.isdigit():
        pagina_inicio = f"{int(pagina_inicio):02d}"

    doi_in = preguntar("DOI (URL completa o 10.31048/...)", defecto=str(actual.get("doi") or "https://doi.org/10.31048/"))
    if doi_in and not doi_in.startswith(("http://", "https://")):
        doi_clean = re.sub(r"^doi:\s*", "", doi_in, flags=re.I)
        doi_in = f"https://doi.org/{doi_clean}"

    autor_corto = preguntar("Autor corto para encabezado par (ej. M. Gómez et al.)", defecto=autor_def)
    titulo_corto = preguntar("Título corto para encabezado impar", defecto=titulo_def)
    recibido = preguntar("Fecha de recibido (DD-MM-AAAA)", defecto=str(actual.get("recibido") or ""))
    correcciones = preguntar("Fecha recibido con correcciones (DD-MM-AAAA)", defecto=str(actual.get("correcciones") or ""))
    aceptado = preguntar("Fecha de aceptado (DD-MM-AAAA)", defecto=str(actual.get("aceptado") or ""))
    editor_sec = preguntar("Editor/a de sección (ej. R. Cattáneo)", defecto=str(actual.get("editor_seccion") or ctx["editor_seccion"]))
    editores_inv = preguntar("Editores invitados (opcional, Enter si no aplica)", defecto=str(actual.get("editores_invitados") or ""))

    # 2b. Identificadores persistentes de autoría e instituciones (ORCID y ROR)
    aut_af_docx = parsear_autores_y_afiliaciones_docx(ruta_ms)
    autores_detectados = aut_af_docx.get("autores", [])
    afils_detectadas = aut_af_docx.get("afiliaciones", [])

    # Combinar con valores previos ya guardados en {id}_metadatos.yaml
    autores_prev = actual.get("autores") if isinstance(actual.get("autores"), list) else []
    afils_prev = actual.get("afiliaciones") if isinstance(actual.get("afiliaciones"), dict) else {}

    autores_finales = []
    if autores_detectados:
        print("\n  --- Identificadores ORCID de autores/as ---")
        print("  (Ingresá los 16 dígitos o la URL de orcid.org; Enter para aceptar/omitir, '-' para borrar)")
        for idx_a, aut in enumerate(autores_detectados):
            prev_item = autores_prev[idx_a] if idx_a < len(autores_prev) and isinstance(autores_prev[idx_a], dict) else {}
            nom_a = str(prev_item.get("nombre") or aut.get("nombre") or "").strip()
            af_a = str(prev_item.get("afiliacion") or aut.get("afiliacion") or "").strip()
            oid_def = normalizar_orcid(prev_item.get("orcid") or aut.get("orcid") or "")[0]
            oid_in = preguntar(f"ORCID de {nom_a}", defecto=oid_def)
            oid_norm, _ = normalizar_orcid(oid_in)
            autores_finales.append({
                "nombre": nom_a,
                "orcid": oid_norm,
                "afiliacion": af_a,
            })
    elif autores_prev:
        autores_finales = autores_prev

    afiliaciones_finales = {}
    if afils_detectadas:
        print("\n  --- Identificadores ROR de afiliaciones institucionales ---")
        print("  (Ingresá el código ROR de 9 caracteres o la URL de ror.org; Enter para aceptar/omitir, '-' para borrar)")
        for af in afils_detectadas:
            clave_af = str(af.get("clave", "1")).strip()
            prev_af = afils_prev.get(clave_af) or afils_prev.get(int(clave_af) if clave_af.isdigit() else clave_af) or {}
            if not isinstance(prev_af, dict):
                prev_af = {"ror": str(prev_af)}
            txt_af = str(prev_af.get("texto") or af.get("texto") or "").strip()
            ror_def = normalizar_ror(prev_af.get("ror") or af.get("ror") or "")[0]
            resumen_af = (txt_af[:56] + "...") if len(txt_af) > 59 else txt_af
            ror_in = preguntar(f"ROR afiliación {clave_af} [{resumen_af}]", defecto=ror_def)
            ror_norm, _ = normalizar_ror(ror_in)
            afiliaciones_finales[clave_af] = {
                "texto": txt_af,
                "ror": ror_norm,
            }
    elif afils_prev:
        afiliaciones_finales = afils_prev

    meta_final = {
        "formato": formato,
        "codigo_publicacion": codigo_pub,
        "seccion": seccion,
        "volumen": volumen,
        "numero": numero,
        "anio": anio,
        "pagina_inicio": pagina_inicio,
        "paginas": str(actual.get("paginas") or ""),
        "doi": doi_in,
        "autor_corto": autor_corto,
        "titulo_corto": titulo_corto,
        "recibido": recibido,
        "correcciones": correcciones,
        "aceptado": aceptado,
        "editor_seccion": editor_sec,
        "editores_invitados": editores_inv,
        "autores": autores_finales,
        "afiliaciones": afiliaciones_finales,
    }
    escribir_metadatos_yaml(ruta_meta, meta_final)
    print(f"\n  Metadatos guardados en: {ruta_meta.name}")

    # 3. Detectar y estandarizar archivo Word de epígrafes (ej. 54854_epigrafes.docx) si existe
    epigrafes_doc = {"figuras": {}, "tablas": {}}
    ruta_epig_encontrada = buscar_archivo_epigrafes(destino, art_id)
    if ruta_epig_encontrada is not None:
        ruta_epig_std = destino / f"{art_id}_epigrafes.docx"
        if ruta_epig_encontrada.resolve() != ruta_epig_std.resolve():
            shutil.copy2(ruta_epig_encontrada, ruta_epig_std)
            print(f"\n  Archivo de epígrafes detectado: {ruta_epig_encontrada.name} -> {ruta_epig_std.name}")
        else:
            print(f"\n  Archivo de epígrafes detectado: {ruta_epig_std.name}")
        epigrafes_doc = parsear_epigrafes_docx(ruta_epig_std)
        n_ef = len(epigrafes_doc.get("figuras", {}))
        n_et = len(epigrafes_doc.get("tablas", {}))
        print(f"  -> Epígrafes extraídos automáticamente: {n_ef} de figura(s), {n_et} de tabla(s).")

    # 4. Detectar y estandarizar imágenes y tablas (precargando epígrafes si existen)
    figuras_cfg = detectar_y_estandarizar_imagenes(
        destino,
        art_id,
        ms_info["figs"],
        formato,
        epigrafes_figuras=epigrafes_doc.get("figuras", {}),
    )
    tablas_cfg = detectar_y_estandarizar_tablas(
        destino,
        art_id,
        ms_info["tabs_texto"],
        ms_info["tabs_embebidas"],
        origen_ms,
        formato,
        epigrafes_tablas=epigrafes_doc.get("tablas", {}),
    )

    ruta_rec = destino / f"{art_id}_recursos.yaml"
    if figuras_cfg or tablas_cfg or not ruta_rec.exists():
        escribir_recursos_yaml(ruta_rec, figuras_cfg, tablas_cfg)

    if (destino / "fasciculo.yaml").exists():
        registrar_en_fasciculo(destino, art_id, codigo_pub, seccion)

    return destino, meta_final


def main():
    args = [a for a in sys.argv[1:] if a not in ("-i", "--interactivo")]

    print("=== Asistente de Maquetación Editorial — Revista del Museo de Antropología (RMA) ===")
    if not args:
        cwd = Path(".").resolve()
        subcarpetas = detectar_carpetas_articulos(cwd)
        if subcarpetas:
            print("\n  Carpetas de artículos detectadas aquí:")
            for idx, sc in enumerate(subcarpetas, start=1):
                estado = " [ya compilada]" if sc["tiene_salida"] else ""
                print(f"    {idx}. {sc['nombre']}  ({sc['n_docx']} .docx, {sc['n_imgs']} imágenes){estado}")
            sin_compilar = [i + 1 for i, sc in enumerate(subcarpetas) if not sc["tiene_salida"]]
            def_carp = str(sin_compilar[-1] if sin_compilar else len(subcarpetas))
            resp_c = preguntar(
                "Carpeta del artículo (número de lista, nombre de carpeta, o '.')",
                defecto=def_carp,
            )
            if resp_c.isdigit() and 1 <= int(resp_c) <= len(subcarpetas):
                destino = subcarpetas[int(resp_c) - 1]["path"].resolve()
            else:
                p_cand = Path(resp_c).expanduser()
                destino = p_cand.resolve() if p_cand.is_absolute() else (cwd / resp_c).resolve()
        else:
            destino_str = preguntar("Carpeta donde pusiste los archivos del artículo", defecto=".")
            p_cand = Path(destino_str).expanduser()
            destino = p_cand.resolve() if p_cand.is_absolute() else (cwd / destino_str).resolve()

        datos_carpeta = extraer_datos_de_nombre_carpeta(destino)
        art_id = preguntar(
            "ID numérico del envío en OJS (ej. 54854)",
            defecto=datos_carpeta["id_ojs"],
            obligatorio=True,
        )
    elif len(args) == 1 and Path(args[0]).is_dir():
        destino = Path(args[0]).resolve()
        datos_carpeta = extraer_datos_de_nombre_carpeta(destino)
        art_id = preguntar(
            "ID numérico del envío en OJS (ej. 54854)",
            defecto=datos_carpeta["id_ojs"],
            obligatorio=True,
        )
    else:
        art_id = args[0].strip()
        destino_str = args[1].strip() if len(args) > 1 else "."
        destino = Path(destino_str).expanduser().resolve()

    destino, _ = ejecutar_flujo_interactivo(destino, art_id)

    if preguntar_si_no("¿Querés compilar el PDF, HTML (OJS) y XML (SciELO) ahora mismo?", defecto_si=True):
        from maquetar_rma import compilar_articulo
        compilar_articulo(destino, art_id)


if __name__ == "__main__":
    main()
