#!/usr/bin/env python3
"""
Automatizador de Maquetación Editorial - Revista del Museo de Antropología (RMA)
Soporta dos perfiles de maquetación con los mismos archivos de entrada:
  - "tradicional": diseño clásico a 2 columnas con márgenes angostos (20 mm) para el número actual
  - "moderno": rediseño a 1 columna con bloques pastel y lectura cómoda para el próximo año

Exporta simultáneamente:
  - PDF editorial maquetado con paginación consecutiva
  - Galería HTML5 responsiva para OJS
  - Archivo XML JATS 1.1 / SciELO SPS 1.9 completo y verificado

Modos de uso:
  python maquetar_rma.py 4812
  python maquetar_rma.py carpeta_fasciculo 4812
  python maquetar_rma.py carpeta_fasciculo 4812 --formato tradicional
  python maquetar_rma.py carpeta_fasciculo 4812 --formato moderno --pagina-inicio 7
  python maquetar_rma.py carpeta_fasciculo 4812 -i
"""

import base64
import html
import mimetypes
import json
import re
import shutil
import sys
import unicodedata
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from docx import Document
from PIL import Image, ImageDraw
from weasyprint import HTML

try:
    import yaml
except ImportError:
    yaml = None

from rma_citaciones import MotorCitacionesRMA

W_NS = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def resolver_destino(args: list) -> tuple[Path, str, int | None, str | None, bool]:
    """Determina (carpeta_trabajo, id_ojs, pagina_inicio_cli, formato_cli, interactivo)."""
    pagina_inicio_cli = None
    formato_cli = None
    interactivo = False
    posicionales = []
    i = 0
    while i < len(args):
        if args[i] in ("--pagina-inicio", "-p") and (i + 1) < len(args):
            pagina_inicio_cli = int(args[i + 1])
            i += 2
        elif args[i] in ("--formato", "-f") and (i + 1) < len(args):
            formato_cli = args[i + 1].strip().lower()
            i += 2
        elif args[i] in ("--interactivo", "-i"):
            interactivo = True
            i += 1
        else:
            posicionales.append(args[i])
            i += 1

    if not posicionales:
        raise SystemExit(
            "Uso: python maquetar_rma.py <id_ojs> [--formato tradicional|moderno] [--pagina-inicio N] [-i]\n"
            "  ó: python maquetar_rma.py <carpeta> <id_ojs> [--formato tradicional|moderno] [--pagina-inicio N] [-i]\n"
            "  ó: python maquetar_rma.py <ruta/al/4812_manuscrito.docx>"
        )
    primero = Path(posicionales[0])
    if len(posicionales) == 1:
        if primero.is_file() and primero.name.endswith("_manuscrito.docx"):
            art_id = primero.name[: -len("_manuscrito.docx")]
            return primero.resolve().parent, art_id, pagina_inicio_cli, formato_cli, interactivo
        return Path(".").resolve(), posicionales[0].strip(), pagina_inicio_cli, formato_cli, interactivo

    carpeta = primero.resolve()
    art_id = posicionales[1].strip()
    return carpeta, art_id, pagina_inicio_cli, formato_cli, interactivo


def _parsear_yaml_basico(texto: str) -> dict:
    """Parser ligero de respaldo para archivos YAML cuando PyYAML no está disponible."""
    lineas_utiles = [
        raw.rstrip()
        for raw in texto.splitlines()
        if raw.strip() and not raw.lstrip().startswith("#")
    ]
    raiz = {}
    pila = [(-1, raiz)]

    def _parse_scalar(v: str):
        v = v.strip()
        if (v.startswith('"') and v.endswith('"')) or (v.startswith("'") and v.endswith("'")):
            return v[1:-1]
        if v.isdigit():
            return int(v)
        if v.lower() in ("true", "false"):
            return v.lower() == "true"
        return v

    for idx, line in enumerate(lineas_utiles):
        indent = len(line) - len(line.lstrip(" "))
        stripped = line.strip()
        while len(pila) > 1 and indent <= pila[-1][0]:
            pila.pop()
        contenedor = pila[-1][1]

        if stripped.startswith("- "):
            val = stripped[2:].strip()
            if ":" in val and not val.startswith(("http://", "https://")):
                k_item, v_item = val.split(":", 1)
                k_item = k_item.strip().strip('"').strip("'")
                nuevo_item = {k_item: _parse_scalar(v_item)}
                if isinstance(contenedor, list):
                    contenedor.append(nuevo_item)
                pila.append((indent + 1, nuevo_item))
            else:
                if isinstance(contenedor, list):
                    contenedor.append(_parse_scalar(val))
            continue

        if ":" in stripped:
            k, v = stripped.split(":", 1)
            k = k.strip().strip('"').strip("'")
            v = v.strip()
            if not v:
                siguiente = lineas_utiles[idx + 1].strip() if (idx + 1) < len(lineas_utiles) else ""
                nuevo = [] if siguiente.startswith("- ") else {}
                if isinstance(contenedor, dict):
                    contenedor[k] = nuevo
                pila.append((indent, nuevo))
            else:
                if isinstance(contenedor, dict):
                    contenedor[k] = _parse_scalar(v)
    return raiz


def cargar_config(ruta_sin_ext: Path, obligatorio: bool = True) -> dict:
    """Carga configuración desde .yaml o .json."""
    ruta_yaml = ruta_sin_ext.with_suffix(".yaml")
    ruta_json = ruta_sin_ext.with_suffix(".json")
    if ruta_yaml.exists():
        contenido = ruta_yaml.read_text(encoding="utf-8")
        if yaml is not None:
            return yaml.safe_load(contenido) or {}
        if ruta_json.exists():
            return json.loads(ruta_json.read_text(encoding="utf-8"))
        return _parsear_yaml_basico(contenido)
    if ruta_json.exists():
        return json.loads(ruta_json.read_text(encoding="utf-8"))
    if obligatorio:
        raise FileNotFoundError(
            f"No se encontró {ruta_yaml.name} (ni {ruta_json.name}) en {ruta_sin_ext.parent}"
        )
    return {}


def actualizar_yaml_campos(ruta_yaml: Path, actualizaciones: dict):
    """Actualiza campos clave-valor simples en un archivo YAML preservando comentarios y estructura."""
    if not ruta_yaml.exists():
        return
    lineas = ruta_yaml.read_text(encoding="utf-8").splitlines()
    claves_restantes = dict(actualizaciones)
    nuevas_lineas = []

    for linea in lineas:
        m = re.match(r"^([A-Za-z0-9_]+)\s*:(.*)$", linea)
        if m and m.group(1) in claves_restantes:
            k = m.group(1)
            val = claves_restantes.pop(k)
            nuevas_lineas.append(f'{k}: "{val}"')
        else:
            nuevas_lineas.append(linea)

    for k, val in claves_restantes.items():
        nuevas_lineas.append(f'{k}: "{val}"')

    ruta_yaml.write_text("\n".join(nuevas_lineas) + "\n", encoding="utf-8")


def formatear_rango_paginas(p_ini: int, p_fin: int, digitos: int = 2) -> str:
    """Formatea un rango consecutivo de páginas (ej. 7 y 26 -> '07-26')."""
    return f"{p_ini:0{digitos}d}-{p_fin:0{digitos}d}"



_PALABRAS_VACIAS_FIN_TITULO = {
    "a", "al", "ante", "bajo", "con", "contra", "de", "del", "desde", "durante",
    "en", "entre", "hacia", "hasta", "mediante", "para", "por", "según", "segun",
    "sin", "sobre", "tras", "versus", "vs", "vs.",
    "el", "la", "los", "las", "un", "una", "unos", "unas", "lo",
    "y", "e", "o", "u", "ni", "que", "quien", "quienes", "cual", "cuales",
    "cuyo", "cuya", "cuyos", "cuyas", "donde", "cuando", "como", "si", "sino",
    "pero", "aunque", "mientras",
    "este", "esta", "estos", "estas", "ese", "esa", "esos", "esas",
    "aquel", "aquella", "aquellos", "aquellas", "su", "sus", "nuestro", "nuestra",
    "nuestros", "nuestras",
    "an", "the", "of", "in", "on", "at", "to", "for", "with", "by", "from",
    "into", "during", "among", "towards", "upon", "and", "or", "but", "nor",
    "as", "that", "which", "who", "whose", "where", "when", "while",
}


def sugerir_titulo_corto(titulo: str, max_len: int = 55) -> str:
    """
    Acorta un título largo para encabezados de página respetando palabras completas
    y evitando terminar en preposiciones, artículos o conjunciones antes de '...'.
    """
    t = re.sub(r"\s+", " ", (titulo or "")).strip()
    if not t or len(t) <= max_len:
        return t

    m_sub = re.match(r"^(.{18," + str(max_len) + r"}?)\s*(?::|—|–|\s-\s)\s*.+$", t)
    if m_sub:
        base_sub = m_sub.group(1).strip(" ,;:-–—")
        palabras_sub = base_sub.split()
        while len(palabras_sub) > 2 and palabras_sub[-1].lower().strip(".,;:()[]\"'“”‘’") in _PALABRAS_VACIAS_FIN_TITULO:
            palabras_sub.pop()
        cand_sub = " ".join(palabras_sub).strip(" ,;:-–—")
        if len(cand_sub) >= 18:
            return cand_sub

    limite = max_len - 3
    recorte = t[:limite]
    if len(t) > limite and not t[limite].isspace():
        ultimo_esp = recorte.rfind(" ")
        if ultimo_esp > 12:
            recorte = recorte[:ultimo_esp]

    recorte = recorte.strip(" ,;:-–—([\"“‘")
    palabras = recorte.split()

    while len(palabras) > 2:
        ultima_norm = palabras[-1].lower().strip(".,;:!?()[]\"'“”‘’")
        if ultima_norm in _PALABRAS_VACIAS_FIN_TITULO:
            palabras.pop()
        else:
            break

    resultado = " ".join(palabras).strip(" ,;:-–—([\"“‘")
    return f"{resultado}..." if resultado else (t[:limite].rstrip() + "...")



def normalizar_codigo_idioma(valor: str) -> str:
    v = str(valor or "").strip().lower()
    if v in ("es", "spa", "esp", "español", "espanol", "spanish") or v.startswith(("es-", "es_")): return "es"
    if v in ("en", "eng", "ing", "ingles", "inglés", "english") or v.startswith(("en-", "en_")): return "en"
    if v in ("pt", "por", "portugues", "portugués", "portuguese") or v.startswith(("pt-", "pt_")): return "pt"
    return ""

def detectar_idioma_texto(texto: str) -> str:
    t = f" {(texto or '').lower()} "
    en_h = len(re.findall(r"\b(?:the|and|of|in|with|from|for|that|this|were|was|between|abstract|keywords|references|figure|table|into|during)\b", t))
    es_h = len(re.findall(r"\b(?:el|la|los|las|del|en|con|fueron|fue|resumen|palabras|referencias|tabla|son|una|un|al|como|más|sobre|también)\b", t)) + len(re.findall(r"[ñ¿¡]", t)) * 2
    pt_h = len(re.findall(r"\b(?:o|os|as|do|da|dos|das|no|na|nos|nas|em|com|foram|foi|resumo|palavras|referências|tabela|são|não|uma|um|ao|aos|pelo|pela|também|através)\b", t)) + len(re.findall(r"[ãõç]", t)) * 2
    if en_h > max(es_h, pt_h) * 1.15: return "en"
    return "pt" if pt_h > es_h * 1.10 else "es"


def determinar_pagina_inicio(meta: dict, pagina_inicio_override: int | None = None) -> int:
    """Resuelve el número de página inicial desde CLI/fascículo, 'pagina_inicio' o 'paginas'."""
    if pagina_inicio_override is not None:
        return int(pagina_inicio_override)
    p_ini_raw = str(meta.get("pagina_inicio", "")).strip()
    if p_ini_raw.isdigit():
        return int(p_ini_raw)
    pags_raw = str(meta.get("paginas", "")).strip()
    m = re.match(r"^(\d+)\s*[-–]\s*(\d+)$", pags_raw)
    if m:
        return int(m.group(1))
    if pags_raw.isdigit():
        return int(pags_raw)
    return 1


def resolver_perfil_formato(formato_override: str | None, meta: dict, rev: dict) -> str:
    """Determina si se compila con perfil 'tradicional' (2 cols) o 'moderno' (1 col)."""
    cand = (
        formato_override
        or str(meta.get("formato", "")).strip().lower()
        or str(rev.get("formato", "")).strip().lower()
        or "tradicional"
    )
    if cand in ("moderno", "1col", "1-columna", "nuevo"):
        return "moderno"
    return "tradicional"


def normalizar_orcid(valor: str) -> tuple[str, str]:
    """Devuelve (id_orcid, url_orcid) normalizados (ej. ('0000-0002-0000-0001', 'https://orcid.org/0000-0002-0000-0001'))."""
    raw = str(valor or "").strip()
    if not raw:
        return "", ""
    m = re.search(r"(\d{4}-\d{4}-\d{4}-\d{3}[\dXx])", raw)
    if m:
        oid = m.group(1).upper()
        return oid, f"https://orcid.org/{oid}"
    limpio = re.sub(r"^https?://(www\.)?orcid\.org/", "", raw, flags=re.I).strip("/")
    if not limpio:
        return "", ""
    return limpio, f"https://orcid.org/{limpio}"


def normalizar_rors(valor: str) -> list[tuple[str, str]]:
    """Devuelve lista de tuplas (id_ror, url_ror) normalizadas (soporta múltiples ROR separados por coma/espacio)."""
    raw = str(valor or "").strip()
    if not raw:
        return []
    encontrados: list[tuple[str, str]] = []
    vistos: set[str] = set()
    for m in re.finditer(r"(?:https?://(?:www\.)?ror\.org/)?(0[a-z0-9]{6}\d{2})\b", raw, flags=re.I):
        rid = m.group(1).lower()
        if rid not in vistos:
            vistos.add(rid)
            encontrados.append((rid, f"https://ror.org/{rid}"))
    if encontrados:
        return encontrados
    for trozo in re.split(r"[,;\s]+", raw):
        limpio = re.sub(r"^https?://(www\.)?ror\.org/", "", trozo.strip(), flags=re.I).strip("/")
        if limpio and limpio.lower() not in vistos:
            vistos.add(limpio.lower())
            encontrados.append((limpio, f"https://ror.org/{limpio}"))
    return encontrados


def normalizar_rors(valor: str) -> list[tuple[str, str]]:
    """Devuelve lista de tuplas (id_ror, url_ror) normalizadas (soporta múltiples ROR)."""
    raw = str(valor or "").strip()
    if not raw:
        return []
    res, vistos = [], set()
    for m in re.finditer(r"(?:https?://(?:www\.)?ror\.org/)?(0[a-z0-9]{6}\d{2})\b", raw, flags=re.I):
        rid = m.group(1).lower()
        if rid not in vistos:
            vistos.add(rid)
            res.append((rid, f"https://ror.org/{rid}"))
    if res:
        return res
    for t in re.split(r"[,;\s]+", raw):
        c = re.sub(r"^https?://(www\.)?ror\.org/", "", t.strip(), flags=re.I).strip("/")
        if c and c.lower() not in vistos:
            vistos.add(c.lower())
            res.append((c, f"https://ror.org/{c}"))
    return res


def normalizar_ror(valor: str) -> tuple[str, str]:
    """Devuelve (id_ror, url_ror) normalizados."""
    lista = normalizar_rors(valor)
    return lista[0] if lista else ("", "")


def _limpiar_urls_orcid_ror_en_html(fragmento_html: str) -> tuple[str, list[str], list[str]]:
    """
    Extrae enlaces/URLs de ORCID y ROR presentes en un fragmento HTML de afiliación
    y los elimina del texto visible para reemplazarlos por íconos hipervinculados.
    """
    orcids = []
    rors = []

    def _sub_anchor(m: re.Match) -> str:
        href = m.group(1)
        if re.search(r"orcid\.org/", href, re.I):
            oid, _ = normalizar_orcid(href)
            if oid and oid not in orcids:
                orcids.append(oid)
            return ""
        if re.search(r"ror\.org/", href, re.I):
            rid, _ = normalizar_ror(href)
            if rid and rid not in rors:
                rors.append(rid)
            return ""
        return m.group(0)

    limpio = re.sub(
        r'(?:\b(?:ORCID|ROR)\s*:?\s*)?<a\b[^>]*?\bhref=["\']([^"\']+)["\'][^>]*>.*?</a>\s*\.?',
        _sub_anchor,
        fragmento_html,
        flags=re.I,
    )

    def _sub_url_plana(m: re.Match) -> str:
        url = m.group(1)
        if re.search(r"orcid\.org/", url, re.I):
            oid, _ = normalizar_orcid(url)
            if oid and oid not in orcids:
                orcids.append(oid)
            return ""
        if re.search(r"ror\.org/", url, re.I):
            rid, _ = normalizar_ror(url)
            if rid and rid not in rors:
                rors.append(rid)
            return ""
        return m.group(0)

    limpio = re.sub(
        r'(?:\b(?:ORCID|ROR)\s*:?\s*)?(https?://(?:www\.)?(?:orcid|ror)\.org/[^\s<>\"\'\)\]]+)\s*\.?',
        _sub_url_plana,
        limpio,
        flags=re.I,
    )
    limpio = re.sub(r"\s{2,}", " ", limpio).strip()
    limpio = re.sub(r"\s+([.,;:])", r"\1", limpio)
    return limpio, orcids, rors


def parsear_autores_y_afiliaciones_docx(ruta_docx: Path) -> dict:
    """
    Extrae la lista estructurada de autores (con superíndice de afiliación y ORCID si figura en el Word)
    y de afiliaciones (con número/marca y ROR si figura en el Word) desde el manuscrito .docx.
    """
    resultado = {"autores": [], "afiliaciones": []}
    if not ruta_docx.exists():
        return resultado
    try:
        doc = Document(str(ruta_docx))
    except Exception:
        return resultado

    parrafos = [p for p in doc.paragraphs if p.text.strip() and p.text.strip() != "."]
    if len(parrafos) < 4:
        return resultado

    p_aut = parrafos[2]
    tokens_aut = []
    for run in p_aut.runs:
        txt = run.text or ""
        if not txt:
            continue
        kind = "sup" if run.font.superscript else "txt"
        if tokens_aut and tokens_aut[-1][0] == kind == "sup":
            tokens_aut[-1] = ("sup", tokens_aut[-1][1] + txt)
        else:
            tokens_aut.append((kind, txt))

    sup_blocks, marcada_parts = [], []
    for kind, val in tokens_aut:
        if kind == "sup":
            marcada_parts.append(f"__SUP_{len(sup_blocks)}__")
            sup_blocks.append(val)
        else:
            marcada_parts.append(val)
    marcada = "".join(marcada_parts)

    partes_aut = [p.strip() for p in re.split(r"(?!(?<=\d),\s*\d),|;|\r?\n|\sy\s|\sand\s|\se\s(?=[A-ZÁÉÍÓÚÂÊÔÃÕÇ])", marcada) if p.strip()]
    autores_lista = []
    for idx_a, parte_mask in enumerate(partes_aut, start=1):
        sups = [sup_blocks[int(m.group(1))] for m in re.finditer(r"__SUP_(\d+)__", parte_mask)]
        sin_sup = re.sub(r"__SUP_\d+__", "", parte_mask).strip()
        orcid_det = ""
        m_oid = re.search(r"(?:https?://(?:www\.)?orcid\.org/)?(\d{4}-\d{4}-\d{4}-\d{3}[\dXx])", sin_sup, flags=re.I)
        if m_oid:
            orcid_det = normalizar_orcid(m_oid.group(1))[0]
            sin_sup = re.sub(r"[\(\[\{]?\s*(?:ORCID\s*:?\s*)?(?:https?://(?:www\.)?orcid\.org/)?\d{4}-\d{4}-\d{4}-\d{3}[\dXx]\s*[\)\]\}]?", "", sin_sup, flags=re.I).strip()
        af_claves = []
        for s in sups:
            m_os = re.search(r"(?:https?://(?:www\.)?orcid\.org/)?(\d{4}-\d{4}-\d{4}-\d{3}[\dXx])", s, flags=re.I)
            if m_os and not orcid_det:
                orcid_det = normalizar_orcid(m_os.group(1))[0]
            s_clean = re.sub(r"(?:https?://(?:www\.)?orcid\.org/)?\d{4}-\d{4}-\d{4}-\d{3}[\dXx]", "", s, flags=re.I)
            for tok in re.split(r"[,;\s:]+", s_clean):
                tc = tok.strip("()[]{}.")
                if tc and tc not in af_claves:
                    af_claves.append(tc)
        sin_sup = sin_sup.rstrip(" :;,").strip()
        if not af_claves:
            m_peg = re.match(r"^(.*?)([\d,\*\s]+)$", sin_sup)
            if m_peg and m_peg.group(1).strip():
                sin_sup = m_peg.group(1).strip()
                af_claves = [x.strip() for x in m_peg.group(2).split(",") if x.strip()]
        if sin_sup:
            autores_lista.append({
                "nombre": sin_sup,
                "afiliacion": ", ".join(af_claves) or (str(idx_a) if len(partes_aut) > 1 else "1"),
                "orcid": orcid_det,
            })

    claves_esperadas = {
        c.strip()
        for a in autores_lista
        for c in str(a.get("afiliacion", "")).split(",")
        if c.strip()
    }
    re_fin_afiliaciones = re.compile(
        r"^(resumen|abstract|sumario|summary|palabras\s*clave|keywords|key\s*words|introducci[oó]n|introduction|1[\.\)\s])\b",
        re.I,
    )

    afiliaciones_lista = []
    for idx_af, p in enumerate(parrafos[3:], start=1):
        t = p.text.strip()
        if re_fin_afiliaciones.match(t):
            break

        clave = str(idx_af)
        texto_sin_num = t
        tiene_marca_clave = False
        if p.runs and p.runs[0].font.superscript and p.runs[0].text.strip():
            clave = p.runs[0].text.strip()
            texto_sin_num = "".join(r.text or "" for r in p.runs[1:]).strip()
            tiene_marca_clave = True
        else:
            m_ini = re.match(r"^(\d+|\*+)\s*(.+)$", t)
            if m_ini:
                clave = m_ini.group(1)
                texto_sin_num = m_ini.group(2).strip()
                tiene_marca_clave = True

        if afiliaciones_lista and claves_esperadas and claves_esperadas.issubset({a["clave"] for a in afiliaciones_lista}):
            if not tiene_marca_clave or clave not in claves_esperadas:
                break

        if len(texto_sin_num) > 260 and not tiene_marca_clave and "ror.org" not in texto_sin_num.lower():
            break

        orcids_enc = [
            normalizar_orcid(u)[0]
            for u in re.findall(r"https?://(?:www\.)?orcid\.org/[^\s<>\"\'\)\]]+", texto_sin_num, flags=re.I)
        ]
        rors_enc = [
            normalizar_ror(u)[0]
            for u in re.findall(r"https?://(?:www\.)?ror\.org/[^\s<>\"\'\)\]]+", texto_sin_num, flags=re.I)
        ]

        texto_limpio = re.sub(
            r"(?:\b(?:ORCID|ROR)\s*:?\s*)?https?://(?:www\.)?(?:orcid|ror)\.org/[^\s<>\"\'\)\]]+\s*\.?",
            "",
            texto_sin_num,
            flags=re.I,
        )
        texto_limpio = re.sub(r"\s{2,}", " ", texto_limpio).strip()
        texto_limpio = re.sub(r"\s+([.,;:])", r"\1", texto_limpio)

        if orcids_enc:
            for oid in orcids_enc:
                asignado = False
                for a in autores_lista:
                    refs = [x.strip() for x in str(a.get("afiliacion", "")).split(",")]
                    if clave in refs and not a.get("orcid"):
                        a["orcid"] = oid
                        asignado = True
                        break
                if not asignado and (idx_af - 1) < len(autores_lista) and not autores_lista[idx_af - 1].get("orcid"):
                    autores_lista[idx_af - 1]["orcid"] = oid

        afiliaciones_lista.append({
            "clave": clave,
            "texto": texto_limpio,
            "ror": ", ".join(rors_enc) if rors_enc else "",
        })

    resultado["autores"] = autores_lista
    resultado["afiliaciones"] = afiliaciones_lista
    return resultado


def normalizar_url_doi(valor: str) -> str:
    """Convierte un DOI o URL en una URL https:// válida para enlaces clicables."""
    v = valor.strip()
    if not v:
        return ""
    if v.startswith(("http://", "https://")):
        return v
    v_sin_prefijo = re.sub(r"^doi:\s*", "", v, flags=re.I)
    return f"https://doi.org/{v_sin_prefijo}"


def enlazar_urls_y_correos(fragmento_html: str) -> str:
    """Convierte URLs, DOIs y correos electrónicos en enlaces <a> clicables."""
    if not fragmento_html:
        return fragmento_html

    partes = re.split(r"(<[^>]+>)", fragmento_html)
    dentro_de_a = False
    resultado = []

    patron_tokens = re.compile(
        r"(?P<url>https?://[^\s<>\"'\)\]]+)"
        r"|(?P<doi>\bdoi:\s*10\.\d{4,9}/[^\s<>\"'\)\]]+)"
        r"|(?P<email>\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}\b)",
        flags=re.IGNORECASE,
    )

    def reemplazar_match(m: re.Match) -> str:
        raw = m.group(0)
        cola = ""
        while raw and raw[-1] in ".,;:":
            cola = raw[-1] + cola
            raw = raw[:-1]

        if m.group("url"):
            href = raw
        elif m.group("doi"):
            doi_id = re.sub(r"^doi:\s*", "", raw, flags=re.I)
            href = f"https://doi.org/{doi_id}"
        else:
            href = f"mailto:{raw}"

        return f'<a href="{html.escape(href, quote=True)}">{raw}</a>{cola}'

    for parte in partes:
        if not parte:
            continue
        if parte.startswith("<"):
            tag_lower = parte.lower()
            if tag_lower.startswith("<a ") or tag_lower == "<a>":
                dentro_de_a = True
            elif tag_lower.startswith("</a"):
                dentro_de_a = False
            resultado.append(parte)
        else:
            if dentro_de_a:
                resultado.append(parte)
            else:
                resultado.append(patron_tokens.sub(reemplazar_match, parte))

    return "".join(resultado)


def _quitar_tildes(s: str) -> str:
    return "".join(
        c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn"
    )


def es_archivo_epigrafes(nombre_archivo: str, art_id: str = "") -> bool:
    """
    Detecta si un archivo .docx corresponde a epígrafes, por ejemplo:
      54854_epigrafes.docx, 54854_Epígrafes.docx, 54854-epigrafes_figuras.docx,
      54854 epigrafes.docx, epigrafes.docx, Epígrafes_54854.docx
    """
    p = Path(nombre_archivo)
    if p.suffix.lower() != ".docx" or p.name.startswith("~$"):
        return False
    stem_norm = _quitar_tildes(p.stem).lower()
    if re.search(r"(?:^|[^a-z0-9])(epigrafes?|leyendas?|captions?|pies?(?:[\s_-]*de[\s_-]*figuras?)?)(?:$|[^a-z0-9])", stem_norm):
        return True
    if art_id:
        sep = r"(?:[\s_\-\.\+]|%20)*"
        if re.match(rf"^{re.escape(art_id)}{sep}(epigraf|leyend|caption|pie)", stem_norm):
            return True
    return False


def buscar_archivo_epigrafes(carpeta: Path, art_id: str) -> Path | None:
    """Busca en 'carpeta' (y subcarpetas 'epigrafes' si existieran) el archivo Word de epígrafes."""
    if not carpeta.exists():
        return None
    std = carpeta / f"{art_id}_epigrafes.docx"
    if std.exists():
        return std

    candidatos = []
    dirs_busq = [carpeta]
    for sub in sorted(carpeta.iterdir(), key=lambda x: x.name.lower()):
        if sub.is_dir() and re.search(r"(epigraf|leyend|caption)", _quitar_tildes(sub.name).lower()):
            dirs_busq.append(sub)

    for d in dirs_busq:
        for p in sorted(d.iterdir(), key=lambda x: x.name.lower()):
            if p.is_file() and es_archivo_epigrafes(p.name, art_id):
                candidatos.append(p)

    if not candidatos:
        return None
    con_id = [c for c in candidatos if c.name.startswith(art_id)]
    return con_id[0] if con_id else candidatos[0]


def _consolidar_tags_inline_html(fragmento_html: str) -> str:
    """Fusiona etiquetas inline idénticas contiguas generadas por runs fragmentados de Word."""
    if not fragmento_html:
        return ""
    s = str(fragmento_html)
    for _ in range(3):
        s = re.sub(r"</(strong|em|b|i|sup|sub)>(\s*)<\1>", r"\2", s, flags=re.I)
        s = re.sub(r"<(strong|em|b|i|sup|sub)>(\s*)</\1>", r"\2", s, flags=re.I)
    return s.strip()


def _run_paragraph_epigrafe_a_html(paragraph) -> str:
    """Convierte un párrafo python-docx a HTML conservando cursivas, negritas, super/subíndices y saltos."""
    partes = []
    for child in paragraph._element:
        tag = child.tag
        runs_xml = []
        if tag == f"{W_NS}r":
            runs_xml = [child]
        elif tag == f"{W_NS}hyperlink":
            runs_xml = child.findall(f"{W_NS}r")

        for r_el in runs_xml:
            for sub in r_el:
                if sub.tag == f"{W_NS}br":
                    partes.append("\n")
                elif sub.tag == f"{W_NS}t":
                    texto = html.escape(sub.text or "")
                    if not texto:
                        continue
                    rPr = r_el.find(f"{W_NS}rPr")
                    if rPr is not None:
                        vert = rPr.find(f"{W_NS}vertAlign")
                        if vert is not None:
                            val = vert.attrib.get(f"{W_NS}val", "")
                            if val == "superscript":
                                texto = f"<sup>{texto}</sup>"
                            elif val == "subscript":
                                texto = f"<sub>{texto}</sub>"
                        if rPr.find(f"{W_NS}i") is not None:
                            texto = f"<em>{texto}</em>"
                        if rPr.find(f"{W_NS}b") is not None:
                            texto = f"<strong>{texto}</strong>"
                    partes.append(texto)
    return _consolidar_tags_inline_html("".join(partes))


def _limpiar_prefijo_recurso_html(html_str: str) -> str:
    """
    Quita prefijos iniciales reiterados como 'Figura 1.', 'Figura 1. Figura 1.',
    'Figure 1:', 'Fig. 1 -', 'Tabla 2.', 'Table 2:' incluso si están envueltos en <strong> o <em>.
    """
    s = (html_str or "").strip()
    patron = re.compile(
        r"^(?:<(?:strong|em|b|i)>\s*)*"
        r"(?:figuras?|figures?|figs?\.?|tablas?|tables?|tabs?\.?|cuadros?)"
        r"\s*0*\d+\s*[\.\:\-\–\—]?\s*"
        r"(?:</(?:strong|em|b|i)>\s*)*"
        r"[\.\:\-\–\—]?\s*",
        flags=re.I,
    )
    while True:
        s_nuevo = patron.sub("", s, count=1).strip()
        s_nuevo = re.sub(r"^<(strong|em|b|i)>\s*</\1>\s*", "", s_nuevo, flags=re.I).strip()
        if s_nuevo == s:
            break
        s = s_nuevo
    return s


def parsear_epigrafes_docx(ruta_docx: Path) -> dict:
    """
    Extrae epígrafes bilingües de Figuras y Tablas desde un Word (ej. 54854_epigrafes.docx).
    Devuelve:
      {
        "figuras": {1: {"es": "...", "en": "..."}, ...},
        "tablas":  {1: {"es": "...", "en": "...", "notas": "..."}, ...}
      }
    """
    resultado = {"figuras": {}, "tablas": {}}
    if not ruta_docx or not ruta_docx.exists():
        return resultado

    doc = Document(str(ruta_docx))

    lineas_raw: list[tuple[str, str]] = []
    for p in doc.paragraphs:
        html_p = _run_paragraph_epigrafe_a_html(p)
        if not html_p:
            continue
        sub_htmls = [x.strip() for x in html_p.split("\n") if x.strip()]
        for sh in sub_htmls:
            txt_p = html.unescape(re.sub(r"<[^>]+>", "", sh)).strip()
            if txt_p:
                lineas_raw.append((txt_p, sh))

    for tbl in doc.tables:
        for row in tbl.rows:
            celdas_txt = []
            celdas_html = []
            vistos = set()
            for c in row.cells:
                if id(c._tc) in vistos:
                    continue
                vistos.add(id(c._tc))
                h_celda = " ".join(_run_paragraph_epigrafe_a_html(p) for p in c.paragraphs if p.text.strip()).strip()
                t_celda = html.unescape(re.sub(r"<[^>]+>", "", h_celda)).strip()
                if t_celda:
                    celdas_txt.append(t_celda)
                    celdas_html.append(h_celda)
            if len(celdas_txt) >= 2:
                m_c0 = re.match(r"^(figuras?|figures?|figs?\.?|tablas?|tables?|cuadros?)\s*0*(\d+)\b", celdas_txt[0], re.I)
                if m_c0:
                    etiq = _quitar_tildes(m_c0.group(1)).lower()
                    num = int(m_c0.group(2))
                    es_en = etiq.startswith(("figure", "table"))
                    bucket = "tablas" if etiq.startswith(("tabla", "table", "cuadro")) else "figuras"
                    entry = resultado[bucket].setdefault(num, {})
                    if es_en:
                        entry["en"] = _limpiar_prefijo_recurso_html(celdas_html[1])
                    else:
                        entry["es"] = _limpiar_prefijo_recurso_html(celdas_html[1])
                        if len(celdas_html) >= 3:
                            entry["en"] = _limpiar_prefijo_recurso_html(celdas_html[2])

    re_etiqueta = re.compile(
        r"^(?P<tipo>figuras?|figures?|figs?\.?|tablas?|tables?|tabs?\.?|cuadros?)\s*0*(?P<num>\d+)"
        r"(?:\s*[\.\:\-\–\—]\s*(?P<resto>.*)|\s+(?P<resto2>.*))?$",
        flags=re.I,
    )
    re_titulo_seccion = re.compile(
        r"^(ep[ií]grafes?(?:\s+de\s+(?:figuras|tablas|cuadros))?|figuras|figures|tablas|tables|leyendas|captions)\s*[\.\:]?\s*$",
        flags=re.I,
    )
    re_nota_tabla = re.compile(r"^(notas?|fuentes?|notes?|sources?|referencias?)\s*[\.\:\-\–]\s*(.+)$", flags=re.I)

    actual_bucket: str | None = None
    actual_num: int | None = None
    esperando_idioma: str | None = None

    for txt_p, html_p in lineas_raw:
        if re_titulo_seccion.match(txt_p):
            actual_bucket = None
            actual_num = None
            esperando_idioma = None
            continue

        m = re_etiqueta.match(txt_p)
        if m:
            tipo_raw = _quitar_tildes(m.group("tipo")).lower()
            num = int(m.group("num"))
            resto_txt = (m.group("resto") or m.group("resto2") or "").strip()

            if tipo_raw.startswith(("figure", "table")):
                idioma = "en"
            elif tipo_raw.startswith(("figura", "tabla", "tabela", "cuadro")):
                idioma = "es"
            else:
                bucket_tmp = "tablas" if tipo_raw.startswith("tab") else "figuras"
                idioma = "en" if "es" in resultado[bucket_tmp].get(num, {}) else "es"

            bucket = "tablas" if tipo_raw.startswith(("tabla", "tabela", "table", "tab", "cuadro")) else "figuras"
            actual_bucket = bucket
            actual_num = num
            entry = resultado[bucket].setdefault(num, {})

            if resto_txt:
                contenido_limpio = _limpiar_prefijo_recurso_html(html_p)
                if " / " in contenido_limpio and idioma == "es" and "en" not in entry:
                    p_es, p_en = contenido_limpio.split(" / ", 1)
                    entry["es"] = p_es.strip()
                    entry["en"] = _limpiar_prefijo_recurso_html(p_en.strip())
                    esperando_idioma = None
                else:
                    entry[idioma] = contenido_limpio
                    esperando_idioma = "en" if idioma == "es" else None
            else:
                esperando_idioma = idioma
            continue

        if actual_bucket is not None and actual_num is not None:
            entry = resultado[actual_bucket].setdefault(actual_num, {})

            if actual_bucket == "tablas" and re_nota_tabla.match(txt_p):
                prev_notas = entry.get("notas", "")
                entry["notas"] = f"{prev_notas} {html_p}".strip() if prev_notas else html_p
                continue

            if esperando_idioma == "es":
                entry["es"] = html_p
                esperando_idioma = "en"
            elif esperando_idioma == "en":
                entry["en"] = html_p
                esperando_idioma = None
            else:
                if "en" in entry:
                    entry["en"] = f"{entry['en']} {html_p}".strip()
                elif "es" in entry:
                    entry["es"] = f"{entry['es']} {html_p}".strip()

    return resultado


def crear_imagen_provisional(ruta_destino: Path, art_id: str, num: int):
    """Genera un recuadro provisional limpio cuando la imagen aún no fue copiada a figuras/."""
    ruta_destino.parent.mkdir(parents=True, exist_ok=True)
    img = Image.new("RGB", (900, 480), "#f3f1ed")
    draw = ImageDraw.Draw(img)
    draw.rectangle([18, 18, 882, 462], outline="#7a2e1d", width=4)
    draw.line([(18, 18), (882, 462)], fill="#d6cec2", width=2)
    draw.line([(882, 18), (18, 462)], fill="#d6cec2", width=2)
    draw.text((250, 215), f"[IMAGEN PENDIENTE: {art_id}_fig{num:02d}.png]", fill="#333333")
    draw.text((235, 245), "Coloca el archivo en figuras/ y volve a compilar", fill="#666666")
    img.save(ruta_destino)


def _directorios_recursos(carpeta: Path, nombres_sub: tuple[str, ...]) -> list[Path]:
    """Devuelve subcarpetas coincidentes sin importar mayúsculas/minúsculas (ej. figuras, Figuras, FIGURAS) más la carpeta raíz."""
    dirs = []
    if carpeta.exists():
        for p in sorted(carpeta.iterdir(), key=lambda x: x.name.lower()):
            if p.is_dir() and p.name.lower() in nombres_sub:
                dirs.append(p)
        dirs.append(carpeta)
    return dirs


def _coincide_recurso_ojs(stem: str, art_id: str, num: int, palabras: tuple[str, ...]) -> bool:
    """
    Detecta nombres estilo OJS como:
      54854_fig01, 54854_fig1, 54854_Figura+1, 54854_figura_1, 54854-figura 1, 54854%20Figura%2001
    """
    pal_re = "|".join(re.escape(w) for w in palabras)
    sep = r"(?:[\s_\-\.\+]|%20)*"
    patron = rf"^{re.escape(art_id)}{sep}(?:{pal_re}){sep}0*{num}(?!\d)"
    return re.match(patron, stem.strip(), re.I) is not None


def _copiar_imagen_compatible(origen: Path, carpeta_salida_fig: Path, nombre_destino: str) -> str:
    """
    Copia la imagen a salida/figuras/. Si es .tif/.tiff (que los navegadores web no muestran
    en galerías HTML5 de OJS), la convierte automáticamente a .png de alta calidad.
    """
    ext = origen.suffix.lower()
    if ext in (".tif", ".tiff"):
        stem_dest = Path(nombre_destino).stem
        nombre_png = f"{stem_dest}.png"
        with Image.open(origen) as im:
            if im.mode not in ("RGB", "RGBA"):
                im = im.convert("RGB")
            im.save(carpeta_salida_fig / nombre_png, format="PNG")
        return nombre_png
    shutil.copy2(origen, carpeta_salida_fig / nombre_destino)
    return nombre_destino


def obtener_imagen_figura(
    carpeta: Path, carpeta_salida_fig: Path, art_id: str, num: int, ruta_explicit: str = ""
) -> str:
    """Localiza y copia la figura a salida/figuras/, o crea un marcador provisional si aún falta."""
    dirs_fig = _directorios_recursos(carpeta, ("figuras", "figura", "figures", "imagenes", "imágenes"))

    if ruta_explicit:
        for base_dir in dirs_fig:
            p = base_dir / ruta_explicit
            if p.exists():
                return _copiar_imagen_compatible(p, carpeta_salida_fig, p.name)

    extensiones = (".png", ".jpg", ".jpeg", ".tif", ".tiff", ".webp")
    patrones = [f"{art_id}_fig{num:02d}", f"{art_id}_fig{num}"]
    for sub in dirs_fig:
        # 1) Coincidencia exacta estándar (54854_fig01.png / 54854_fig1.png)
        for pat in patrones:
            for ext in extensiones:
                cand = sub / f"{pat}{ext}"
                if cand.exists():
                    return _copiar_imagen_compatible(cand, carpeta_salida_fig, cand.name)
        # 2) Coincidencia flexible estilo OJS (54854_Figura+1.tiff, 54854_figura_1.png, 54854 Figura 1.jpg)
        for cand in sorted(sub.iterdir(), key=lambda p: p.name.lower()):
            if not cand.is_file() or cand.suffix.lower() not in extensiones:
                continue
            if _coincide_recurso_ojs(cand.stem, art_id, num, ("figura", "fig", "figure")):
                nombre_std = f"{art_id}_fig{num:02d}{cand.suffix.lower()}"
                return _copiar_imagen_compatible(cand, carpeta_salida_fig, nombre_std)

    nombre_prov = f"{art_id}_fig{num:02d}_pendiente.png"
    crear_imagen_provisional(carpeta_salida_fig / nombre_prov, art_id, num)
    print(f"  [Aviso] Falta Figura {num} ({art_id}_Figura+{num}.* o figuras/{art_id}_fig{num:02d}.*) -> se insertó recuadro provisional.")
    return nombre_prov


def extraer_notas_al_pie(ruta_docx: Path) -> dict:
    """Extrae notas al pie formateadas desde word/footnotes.xml."""
    notas = {}
    with zipfile.ZipFile(ruta_docx, "r") as zf:
        if "word/footnotes.xml" not in zf.namelist():
            return notas
        root = ET.fromstring(zf.read("word/footnotes.xml"))
        for fn in root.findall(f"{W_NS}footnote"):
            fn_id = fn.attrib.get(f"{W_NS}id")
            fn_type = fn.attrib.get(f"{W_NS}type", "")
            if fn_type in ("separator", "continuationSeparator") or not fn_id:
                continue
            partes = []
            for r in fn.findall(f".//{W_NS}r"):
                t_elems = r.findall(f"{W_NS}t")
                txt = "".join(t.text or "" for t in t_elems)
                if not txt:
                    continue
                txt = html.escape(txt)
                rPr = r.find(f"{W_NS}rPr")
                if rPr is not None:
                    if rPr.find(f"{W_NS}i") is not None:
                        txt = f"<em>{txt}</em>"
                    if rPr.find(f"{W_NS}b") is not None:
                        txt = f"<strong>{txt}</strong>"
                partes.append(txt)
            texto_fn = _consolidar_tags_inline_html("".join(partes))
            if texto_fn:
                notas[fn_id] = enlazar_urls_y_correos(texto_fn)
    return notas


def paragraph_a_html(paragraph, notas_docx: dict, notas_usadas: list) -> str:
    """Convierte un párrafo de python-docx a HTML preservando cursivas, negritas, superíndices y notas."""
    partes = []
    for child in paragraph._element:
        tag = child.tag
        runs_xml = []
        if tag == f"{W_NS}r":
            runs_xml = [child]
        elif tag == f"{W_NS}hyperlink":
            runs_xml = child.findall(f"{W_NS}r")

        for r_el in runs_xml:
            fn_ref = r_el.find(f"{W_NS}footnoteReference")
            if fn_ref is not None:
                fn_id = fn_ref.attrib.get(f"{W_NS}id")
                if fn_id in notas_docx:
                    notas_usadas.append((fn_id, notas_docx[fn_id]))
                    num_orden = len(notas_usadas)
                    partes.append(
                        f'<a class="fn-ref-link" id="fnref-{num_orden}" href="#fn-{num_orden}">{num_orden}</a>'
                        f'<span class="footnote">{notas_docx[fn_id]}</span>'
                    )
                continue

            t_elems = r_el.findall(f"{W_NS}t")
            texto = "".join(t.text or "" for t in t_elems)
            if not texto:
                continue
            texto = html.escape(texto)

            rPr = r_el.find(f"{W_NS}rPr")
            if rPr is not None:
                vert = rPr.find(f"{W_NS}vertAlign")
                if vert is not None:
                    val = vert.attrib.get(f"{W_NS}val", "")
                    if val == "superscript":
                        texto = f"<sup>{texto}</sup>"
                    elif val == "subscript":
                        texto = f"<sub>{texto}</sub>"
                if rPr.find(f"{W_NS}i") is not None:
                    texto = f"<em>{texto}</em>"
                if rPr.find(f"{W_NS}b") is not None:
                    texto = f"<strong>{texto}</strong>"
            partes.append(texto)

    return _consolidar_tags_inline_html("".join(partes))


def limpiar_enfasis_externo(html_str: str) -> str:
    """Quita <strong> o <em> si envuelven todo el bloque."""
    s = html_str.strip()
    for tag in ("strong", "em"):
        patron = rf"^<{tag}>(.*)</{tag}>$"
        m = re.match(patron, s, flags=re.DOTALL)
        if m and f"</{tag}>" not in m.group(1):
            s = m.group(1).strip()
    return s


def es_subtitulo(paragraph, texto_plano: str) -> bool:
    """Detecta subtítulos de sección por estilo Heading o por estar enteramente en negrita."""
    if len(texto_plano) <= 2 or len(texto_plano) > 160:
        return False
    if re.match(r"^(FIGURA|TABLA)\s+\d+\s*$", texto_plano, re.I):
        return False
    if paragraph.style.name.startswith("Heading"):
        return True
    runs_con_texto = [r for r in paragraph.runs if r.text.strip()]
    if runs_con_texto and all(r.bold for r in runs_con_texto):
        return True
    return False


def extraer_tabla_docx(table, notas_docx: dict, notas_usadas: list) -> str:
    """Convierte una tabla de Word a HTML limpio soportando celdas combinadas horizontales."""
    filas_html = []
    for i, row in enumerate(table.rows):
        tag = "th" if i == 0 else "td"
        celdas = []
        vistos = set()
        for cell in row.cells:
            tc_id = id(cell._tc)
            if tc_id in vistos:
                continue
            vistos.add(tc_id)
            grid_span = (
                cell._tc.tcPr.gridSpan.val
                if (cell._tc.tcPr is not None and cell._tc.tcPr.gridSpan is not None)
                else 1
            )
            colspan_attr = f' colspan="{grid_span}"' if grid_span > 1 else ""
            contenido_celda = "<br>".join(
                enlazar_urls_y_correos(paragraph_a_html(p, notas_docx, notas_usadas))
                for p in cell.paragraphs
                if p.text.strip()
            )
            celdas.append(f"<{tag}{colspan_attr}>{contenido_celda}</{tag}>")
        filas_html.append(f"<tr>{''.join(celdas)}</tr>")

    thead = f"<thead>{filas_html[0]}</thead>" if filas_html else ""
    tbody = f"<tbody>{''.join(filas_html[1:])}</tbody>" if len(filas_html) > 1 else ""
    return f"<table>{thead}{tbody}</table>"



def _leer_xlsx_a_html(ruta_xlsx: Path) -> str:
    """Convierte la primera hoja de un archivo .xlsx a tabla HTML sin requerir dependencias externas."""
    ns = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    with zipfile.ZipFile(ruta_xlsx, "r") as zf:
        shared = []
        if "xl/sharedStrings.xml" in zf.namelist():
            ss_root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
            for si in ss_root.findall("m:si", ns):
                txt = "".join(t.text or "" for t in si.findall(".//m:t", ns))
                shared.append(txt)
        hojas = sorted(n for n in zf.namelist() if n.startswith("xl/worksheets/sheet") and n.endswith(".xml"))
        if not hojas:
            return "<p>[Hoja vacía]</p>"
        sh_root = ET.fromstring(zf.read(hojas[0]))
        filas = []
        for idx_r, row in enumerate(sh_root.findall(".//m:sheetData/m:row", ns)):
            tag = "th" if idx_r == 0 else "td"
            celdas = []
            for c in row.findall("m:c", ns):
                t_attr = c.attrib.get("t", "")
                v_el = c.find("m:v", ns)
                is_el = c.find("m:is", ns)
                val = ""
                if t_attr == "s" and v_el is not None and v_el.text is not None:
                    s_idx = int(v_el.text)
                    val = shared[s_idx] if 0 <= s_idx < len(shared) else v_el.text
                elif t_attr == "inlineStr" and is_el is not None:
                    val = "".join(t.text or "" for t in is_el.findall(".//m:t", ns))
                elif v_el is not None and v_el.text is not None:
                    val = v_el.text
                celdas.append(f"<{tag}>{enlazar_urls_y_correos(html.escape(val.strip()))}</{tag}>")
            if celdas:
                filas.append(f"<tr>{''.join(celdas)}</tr>")
        if not filas:
            return "<p>[Tabla vacía]</p>"
        thead = f"<thead>{filas[0]}</thead>"
        tbody = f"<tbody>{''.join(filas[1:])}</tbody>" if len(filas) > 1 else ""
        return f"<table>{thead}{tbody}</table>"


def _leer_csv_a_html(ruta_csv: Path) -> str:
    """Convierte un archivo .csv (detectando coma o punto y coma) a tabla HTML."""
    import csv
    raw = ruta_csv.read_text(encoding="utf-8-sig", errors="replace")
    delim = ";" if raw.count(";") > raw.count(",") else ","
    reader = csv.reader(raw.splitlines(), delimiter=delim)
    filas = []
    for idx_r, row in enumerate(reader):
        if not any(cell.strip() for cell in row):
            continue
        tag = "th" if idx_r == 0 else "td"
        celdas = [f"<{tag}>{enlazar_urls_y_correos(html.escape(cell.strip()))}</{tag}>" for cell in row]
        filas.append(f"<tr>{''.join(celdas)}</tr>")
    if not filas:
        return "<p>[Tabla vacía]</p>"
    thead = f"<thead>{filas[0]}</thead>"
    tbody = f"<tbody>{''.join(filas[1:])}</tbody>" if len(filas) > 1 else ""
    return f"<table>{thead}{tbody}</table>"


def _archivo_tabla_a_html(ruta_archivo: Path) -> str:
    """Extrae la tabla de un archivo individual (.docx, .xlsx o .csv) y la devuelve como HTML."""
    ext = ruta_archivo.suffix.lower()
    if ext == ".docx":
        doc_t = Document(ruta_archivo)
        if doc_t.tables:
            return extraer_tabla_docx(doc_t.tables[0], {}, [])
        # Si el autor pegó texto tabulado en lugar de una tabla de Word, convertir líneas con tabulaciones
        lineas = [p.text.strip() for p in doc_t.paragraphs if p.text.strip()]
        if lineas and any("\t" in l for l in lineas):
            filas = []
            for idx_r, l in enumerate(lineas):
                tag = "th" if idx_r == 0 else "td"
                celdas = [f"<{tag}>{enlazar_urls_y_correos(html.escape(c.strip()))}</{tag}>" for c in l.split("\t")]
                filas.append(f"<tr>{''.join(celdas)}</tr>")
            thead = f"<thead>{filas[0]}</thead>"
            tbody = f"<tbody>{''.join(filas[1:])}</tbody>" if len(filas) > 1 else ""
            return f"<table>{thead}{tbody}</table>"
        return "<p>[El archivo no contiene una tabla estructurada]</p>"
    elif ext == ".xlsx":
        return _leer_xlsx_a_html(ruta_archivo)
    elif ext == ".csv":
        return _leer_csv_a_html(ruta_archivo)
    return "<p>[Formato de tabla no soportado]</p>"


def resolver_html_tabla(
    carpeta: Path,
    art_id: str,
    num: int,
    info_cfg: dict,
    tablas_lote_html: list[str],
    tablas_embebidas: list[str],
) -> str:
    """
    Resuelve el HTML de la Tabla N buscando en orden:
      1) Archivo explícito en recursos.yaml (ej. archivo: '54854_tabla01.docx')
      2) Archivo individual estandarizado ({id}_tabla01.docx / .xlsx / .csv)
      3) Archivo conjunto ({id}_tablas.docx)
      4) Tablas embebidas dentro de {id}_manuscrito.docx
    """
    archivo_explicit = str(info_cfg.get("archivo", "")).strip()
    dirs_tab = _directorios_recursos(carpeta, ("tablas", "tabla", "tables", "cuadros"))
    if archivo_explicit:
        for base_dir in dirs_tab:
            cand = base_dir / archivo_explicit
            if cand.exists():
                return _archivo_tabla_a_html(cand)

    for base_dir in dirs_tab:
        if not base_dir.exists():
            continue
        for pat in (f"{art_id}_tabla{num:02d}", f"{art_id}_tabla{num}"):
            for ext in (".docx", ".xlsx", ".csv"):
                cand = base_dir / f"{pat}{ext}"
                if cand.exists():
                    return _archivo_tabla_a_html(cand)
        for cand in sorted(base_dir.iterdir(), key=lambda p: p.name.lower()):
            if not cand.is_file() or cand.suffix.lower() not in (".docx", ".xlsx", ".csv"):
                continue
            if _coincide_recurso_ojs(cand.stem, art_id, num, ("tabla", "tab", "table", "cuadro")):
                return _archivo_tabla_a_html(cand)

    if (num - 1) < len(tablas_lote_html):
        return tablas_lote_html[num - 1]
    if (num - 1) < len(tablas_embebidas):
        return tablas_embebidas[num - 1]

    print(f"  [Aviso] Falta archivo para Tabla {num} -> se insertó recuadro provisional.")
    return f"<p><strong>[Tabla {num} pendiente de incorporar]</strong></p>"


def parsear_manuscrito(ruta_docx: Path) -> dict:
    # Extrae metadatos del frente, cuerpo estructurado, notas al pie y bibliografía (ES y EN).
    doc = Document(ruta_docx)
    notas_docx = extraer_notas_al_pie(ruta_docx)
    notas_usadas = []

    parrafos = [p for p in doc.paragraphs if p.text.strip() and p.text.strip() != "."]
    if len(parrafos) < 6:
        raise ValueError(f"El manuscrito {ruta_docx.name} tiene muy pocos párrafos para extraer la estructura RMA.")

    t0_html = limpiar_enfasis_externo(paragraph_a_html(parrafos[0], notas_docx, notas_usadas))
    t1_html = limpiar_enfasis_externo(paragraph_a_html(parrafos[1], notas_docx, notas_usadas))
    if detectar_idioma_texto(parrafos[0].text) == "en" and detectar_idioma_texto(parrafos[1].text) in ("es", "pt"):
        titulo_es, titulo_en = t1_html, t0_html
    else:
        titulo_es, titulo_en = t0_html, t1_html

    data = {
        "titulo_es": titulo_es,
        "titulo_en": titulo_en,
        "autores": limpiar_enfasis_externo(paragraph_a_html(parrafos[2], notas_docx, notas_usadas)),
        "afiliaciones": [],
        "resumen": "",
        "palabras_clave": "",
        "abstract": "",
        "keywords": "",
        "idioma_resumen_local": "",
        "bloques_cuerpo": [],
        "referencias": [],
        "notas_usadas": notas_usadas,
        "tablas_embebidas": [extraer_tabla_docx(t, notas_docx, notas_usadas) for t in doc.tables],
    }

    re_refs = re.compile(
        r"^(?:\d+(?:\.\d+)*\.?\s+)?(?:refer[eê]ncias(?:\s+bibliogr[aá]ficas)?|bibliograf[ií]a|references(?:\s+cited)?|literature\s+cited|bibliography)\s*[.:]?$",
        re.I,
    )
    re_inicio_cuerpo = re.compile(
        r"^(?:(?:\d+(?:\.\d+)*|[ivx]+)\.?\s+)?"
        r"(?:introduction|introducci[oó]n|introdu[cç][aã]o|background|antecedentes|study\s+area|[aá]rea\s+de\s+estudio|"
        r"materials?\s+and\s+methods|materiales\s+y\s+m[eé]todos|materiais?\s+e\s+m[eé]todos|methodology|metodolog[ií]a|metodologia|"
        r"theoretical\s+framework|marco\s+te[oó]rico|results|resultados|discussion|discusi[oó]n|discuss[aã]o|"
        r"conclusions?|conclusiones|conclus[aã]o|conclus[oõ]es|considera[cç][oõ]es\s+finais|acknowledgments?|acknowledgements?|agradecimientos|agradecimentos)\s*[.:]?$",
        re.I,
    )

    estado = "afiliaciones"
    for p in parrafos[3:]:
        texto_plano = p.text.strip()
        low = re.sub(r"\s+", " ", texto_plano).strip().lower()
        html_p = enlazar_urls_y_correos(paragraph_a_html(p, notas_docx, notas_usadas))

        if low in ("resumen", "resumen:", "resumo", "resumo:"):
            estado = "resumen"
            data["idioma_resumen_local"] = "pt" if low.startswith("resumo") else (data["idioma_resumen_local"] or "es")
            continue
        elif re.match(r"^pala[bv]ras[\s-]+cha?ves?\s*:", texto_plano, re.I) or ((low.startswith("palabras clave") or low.startswith("palavras-chave") or low.startswith("palavras chave")) and len(texto_plano) < 300):
            data["palabras_clave"] = html_p
            if low.startswith("palavras"): data["idioma_resumen_local"] = "pt"
            estado = "cuerpo"
            continue
        elif low in ("abstract", "abstract:"):
            estado = "abstract"
            continue
        elif re.match(r"^key\s*-?\s*words?\s*:", texto_plano, re.I) or (low.startswith("keywords") and len(texto_plano) < 300):
            data["keywords"] = html_p
            estado = "cuerpo"
            continue
        elif re_refs.match(texto_plano):
            estado = "referencias"
            continue

        if estado in ("resumen", "abstract") and (re_inicio_cuerpo.match(texto_plano) or es_subtitulo(p, texto_plano)):
            estado = "cuerpo"

        if estado == "afiliaciones":
            data["afiliaciones"].append(html_p)
        elif estado == "resumen":
            data["resumen"] += f"<p>{html_p}</p>"
        elif estado == "abstract":
            data["abstract"] += f"<p>{html_p}</p>"
        elif estado == "referencias":
            data["referencias"].append(html_p)
        elif estado == "cuerpo":
            m_recurso = re.match(r"^(FIGURA|FIGURE|TABLA|TABLE|TABELA|QUADRO)\s+(\d+)\s*$", texto_plano, re.I)
            if m_recurso:
                tipo_rec = "FIGURA" if m_recurso.group(1).upper().startswith("FIG") else "TABLA"
                data["bloques_cuerpo"].append({
                    "tipo": tipo_rec,
                    "num": int(m_recurso.group(2)),
                })
            elif re.match(r"^[A-ZÁÉÍÓÚÑ][a-záéíóúñ]+,\s+\d{1,2}\s+de\s+[a-z]+\s+de\s+\d{4}\.?$", texto_plano):
                data["bloques_cuerpo"].append({"tipo": "fecha", "html": html_p})
            elif es_subtitulo(p, texto_plano):
                data["bloques_cuerpo"].append({
                    "tipo": "h2",
                    "html": limpiar_enfasis_externo(html_p),
                })
            elif p.paragraph_format.left_indent and p.paragraph_format.left_indent.pt > 15:
                data["bloques_cuerpo"].append({"tipo": "blockquote", "html": html_p})
            else:
                data["bloques_cuerpo"].append({"tipo": "p", "html": html_p})

    return data


def _relacion_aspecto_imagen(ruta_img: Path) -> float | None:
    """Devuelve ancho/alto de la imagen, o None si no se puede leer."""
    try:
        if ruta_img.exists():
            with Image.open(ruta_img) as im:
                w, h = im.size
                if h > 0:
                    return w / h
    except Exception:
        pass
    return None


def clase_ancho_recurso(
    info: dict,
    perfil: str,
    ruta_img: Path | None = None,
    num_cols_tabla: int | None = None,
) -> str:
    """
    Determina la clase de ancho según el perfil ('tradicional' 2 cols vs 'moderno' 1 col),
    ajustando automáticamente según orientación de imagen o cantidad de columnas de tabla.
    """
    ancho = str(info.get("ancho", "")).strip().lower()
    explicito_columna = ancho in ("columna", "1", "1col", "centrado", "74%", "70%", "medio") or (
        "columnas" in info and int(info.get("columnas", 2)) == 1
    )
    explicito_completo = ancho in ("completo", "2", "2col", "100%", "full") or (
        "columnas" in info and int(info.get("columnas", 1)) == 2
    )

    # Si es una tabla con 4 o más columnas en formato tradicional y no se forzó 1 columna,
    # promover automáticamente a ancho completo (170 mm) para evitar aplastamiento.
    if num_cols_tabla is not None and num_cols_tabla >= 4 and not explicito_columna:
        return "ancho-completo"

    if explicito_columna:
        return "ancho-columna" if perfil == "tradicional" else "ancho-centrado"
    if explicito_completo:
        return "ancho-completo"

    # Si es una figura sin ancho explícito, decidir por orientación (vertical/cuadrada -> 1 col, apaisada -> completo)
    if ruta_img is not None:
        ratio = _relacion_aspecto_imagen(ruta_img)
        if ratio is not None and ratio < 1.28:
            return "ancho-columna" if perfil == "tradicional" else "ancho-centrado"

    return "ancho-completo"


def _etiquetar_filas_tabla_pdf(tabla_html: str, blk_id: str) -> str:
    """Agrega data-trow='blk_id:idx' a cada fila del cuerpo de la tabla para rastrear saltos de página."""
    m_tbody = re.search(r"(<tbody\b[^>]*>)(.*?)(</tbody>)", tabla_html, flags=re.I | re.DOTALL)
    idx_counter = [0]

    def _add_attr(m):
        tag_open = m.group(1)
        if "data-trow=" in tag_open:
            return tag_open
        i = idx_counter[0]
        idx_counter[0] += 1
        return tag_open[:-1] + f' data-trow="{blk_id}:{i}">'

    if m_tbody:
        cuerpo_etq = re.sub(r"(<tr\b[^>]*>)", _add_attr, m_tbody.group(2), flags=re.I)
        return tabla_html[:m_tbody.start()] + m_tbody.group(1) + cuerpo_etq + m_tbody.group(3) + tabla_html[m_tbody.end():]

    # Si no hay <tbody> explícito, omitir la primera fila <tr> si funciona como encabezado
    filas = list(re.finditer(r"<tr\b[^>]*>.*?</tr>", tabla_html, flags=re.I | re.DOTALL))
    if len(filas) <= 1:
        return tabla_html
    inicio_cuerpo = 1 if "<th" in filas[0].group(0).lower() else 0
    partes = [tabla_html[:filas[inicio_cuerpo].start()]]
    pos = filas[inicio_cuerpo].start()
    for f in filas[inicio_cuerpo:]:
        partes.append(tabla_html[pos:f.start()])
        tr_etq = re.sub(r"^(<tr\b[^>]*>)", _add_attr, f.group(0), count=1, flags=re.I)
        partes.append(tr_etq)
        pos = f.end()
    partes.append(tabla_html[pos:])
    return "".join(partes)


def _html_caption_bilingue(etiq_es: str, etiq_en: str, num, cap_es: str, cap_en: str, idioma_art: str | bool) -> str:
    if (idioma_art == "en") if isinstance(idioma_art, str) else bool(idioma_art):
        return (
            f'    <span class="cap-en"><strong>{etiq_en} {num}.</strong> {cap_en}</span>\n'
            f'    <span class="cap-es"><em>{etiq_es} {num}. {cap_es}</em></span>'
        )
    return (
        f'    <span class="cap-es"><strong>{etiq_es} {num}.</strong> {cap_es}</span>\n'
        f'    <span class="cap-en"><em>{etiq_en} {num}. {cap_en}</em></span>'
    )


def _segmentar_cuerpo_multicolumna(lista_cuerpo: list[str], cierre_final: str, perfil: str) -> list[str]:
    if perfil != "tradicional":
        return [f'<main class="cuerpo-articulo">\n{chr(10).join(lista_cuerpo)}\n{cierre_final}\n    </main>']
    partes, buf, primera = [], [], True
    for item in lista_cuerpo:
        if "rma-tabla" in item and "ancho-completo" in item and ("tabla-divisible" in item or "tabla-fragmento-" in item):
            if buf:
                cls = "cuerpo-articulo" if primera else "cuerpo-articulo cuerpo-articulo-cont"
                partes.append(f'<div class="{cls}">\n{chr(10).join(buf)}\n</div>\n')
                buf, primera = [], False
            partes.append(f'<div class="cuerpo-articulo-tabla-ancha">\n{item}\n</div>\n')
            primera = False
        else:
            buf.append(item)
    cls = "cuerpo-articulo" if primera else "cuerpo-articulo cuerpo-articulo-cont"
    partes.append(f'<div class="{cls}">\n{chr(10).join(buf)}\n{cierre_final}\n</div>')
    return partes


def _construir_fragmentos_tabla_dividida(blk_tab: dict, row_pages: dict[int, int], pagina_inicio: int, idioma_art: str = "es") -> list[str] | None:
    """
    Si una tabla divisible se partió en 2 o más páginas, devuelve la lista de fragmentos HTML
    con 'Continúa en página X' al pie de cada corte y 'Continuación de Tabla N' + encabezado
    repetido en la página siguiente.
    """
    if not row_pages:
        return None
    paginas_unicas = sorted(set(row_pages.values()))
    if len(paginas_unicas) <= 1:
        return None

    tabla_code = blk_tab["tabla_code_pdf"]
    m_table = re.search(r"<table\b([^>]*)>(.*)</table>", tabla_code, flags=re.I | re.DOTALL)
    if not m_table:
        return None
    table_attrs = m_table.group(1)
    inner = m_table.group(2)

    m_col = re.search(r"<colgroup\b[^>]*>.*?</colgroup>", inner, flags=re.I | re.DOTALL)
    colgroup_html = m_col.group(0) if m_col else ""

    m_thead = re.search(r"<thead\b[^>]*>.*?</thead>", inner, flags=re.I | re.DOTALL)
    if m_thead:
        thead_html = m_thead.group(0)
        m_tbody = re.search(r"<tbody\b[^>]*>(.*?)</tbody>", inner, flags=re.I | re.DOTALL)
        cuerpo_src = m_tbody.group(1) if m_tbody else inner
        filas_cuerpo = re.findall(r"<tr\b[^>]*>.*?</tr>", cuerpo_src, flags=re.I | re.DOTALL)
    else:
        todas_filas = re.findall(r"<tr\b[^>]*>.*?</tr>", inner, flags=re.I | re.DOTALL)
        if len(todas_filas) <= 1:
            return None
        if "<th" in todas_filas[0].lower():
            thead_html = f"<thead>{todas_filas[0]}</thead>"
            filas_cuerpo = todas_filas[1:]
        else:
            thead_html = ""
            filas_cuerpo = todas_filas

    if not filas_cuerpo:
        return None

    by_page: dict[int, list[int]] = {}
    for r_idx, p_idx in sorted(row_pages.items()):
        if 0 <= r_idx < len(filas_cuerpo):
            by_page.setdefault(p_idx, []).append(r_idx)

    pages_sorted = sorted(by_page.keys())
    if len(pages_sorted) <= 1:
        return None

    # Reservar el espacio de la última fila en páginas intermedias para que 'Continúa en página X'
    # quede siempre al pie de esa página y no empuje una fila suelta a la siguiente
    chunks = [list(by_page[p]) for p in pages_sorted]
    for i in range(len(chunks) - 1):
        if len(chunks[i]) > 2:
            moved = chunks[i].pop()
            chunks[i + 1].insert(0, moved)

    num = blk_tab["num"]
    clase_ancho = blk_tab["clase_ancho"]
    clase_densidad = blk_tab["clase_densidad"]
    notas_t = blk_tab["notas_html"]
    cap_es = blk_tab["cap_es"]
    cap_en = blk_tab["cap_en"]

    fragmentos_html = []
    for idx_p, r_indices in enumerate(chunks):
        pidx = pages_sorted[idx_p]
        es_primera = idx_p == 0
        es_ultima = idx_p == len(chunks) - 1
        sub_rows = "\n".join(filas_cuerpo[r] for r in r_indices)
        folio_sig = pagina_inicio + pidx + 1

        sup_html = (
            f'<div class="tabla-cont-sup">{"Continuation of Table" if idioma_art == "en" else ("Continuação da Tabela" if idioma_art == "pt" else "Continuación de Tabla")} {num}</div>'
            if not es_primera
            else ""
        )
        inf_html = (
            f'<div class="tabla-cont-inf">{"Continued on page" if idioma_art == "en" else ("Continua na página" if idioma_art == "pt" else "Continúa en página")} {folio_sig}</div>'
            if not es_ultima
            else ""
        )
        pie_final_html = (
            f"{notas_t}\n"
            f'  <div class="tabla-caption">\n'
            f'{_html_caption_bilingue("Tabela" if idioma_art == "pt" else "Tabla", "Table", num, cap_es, cap_en, idioma_art)}\n'
            f"  </div>"
            if es_ultima
            else ""
        )
        cls_frag = (
            (" tabla-fragmento-continua" if not es_ultima else "")
            + (" tabla-fragmento-continuacion" if not es_primera else "")
        )
        id_attr = f' id="tab-{num}"' if es_primera else f' id="tab-{num}-cont-{idx_p + 1}"'
        fragmentos_html.append(
            f'<div class="rma-tabla {clase_ancho} {clase_densidad} tabla-divisible{cls_frag}"{id_attr}>\n'
            f"  {sup_html}\n"
            f"  <table{table_attrs}>{colgroup_html}{thead_html}<tbody>{sub_rows}</tbody></table>\n"
            f"  {inf_html}\n"
            f"  {pie_final_html}\n"
            f"</div>"
        )

    return fragmentos_html


def preparar_tabla_responsive(tabla_html: str) -> tuple[str, int, str, int, int]:
    """
    Analiza una tabla HTML, calcula anchos proporcionales de columna (<colgroup>) para evitar
    desbordes en table-layout: fixed, y devuelve (html_ajustado, num_columnas, clase_densidad).
    """
    m_table = re.search(r"<table\b[^>]*>(.*)</table>", tabla_html, flags=re.I | re.DOTALL)
    if not m_table:
        return (tabla_html, 1, "tabla-normal", 0, 0)

    filas = re.findall(r"<tr\b[^>]*>(.*?)</tr>", m_table.group(1), flags=re.I | re.DOTALL)
    if not filas:
        return (tabla_html, 1, "tabla-normal", 0, 0)

    anchos_col: list[int] = []
    for fila in filas:
        celdas = re.findall(r"<(th|td)\b([^>]*)>(.*?)</(?:th|td)>", fila, flags=re.I | re.DOTALL)
        col_idx = 0
        for _, attrs, contenido in celdas:
            m_span = re.search(r"colspan\s*=\s*['\"]?(\d+)", attrs, flags=re.I)
            span = max(1, int(m_span.group(1))) if m_span else 1
            texto = re.sub(r"<[^>]+>", " ", contenido)
            texto = html.unescape(re.sub(r"\s+", " ", texto)).strip()
            largo = max(3, min(len(texto), 48))
            largo_por_col = max(3, int(largo / span))
            while len(anchos_col) < col_idx + span:
                anchos_col.append(4)
            for k in range(span):
                if largo_por_col > anchos_col[col_idx + k]:
                    anchos_col[col_idx + k] = largo_por_col
            col_idx += span

    num_cols = max(1, len(anchos_col))
    if num_cols <= 5:
        clase_densidad = "tabla-normal"
    elif num_cols <= 8:
        clase_densidad = "tabla-compacta"
    elif num_cols <= 12:
        clase_densidad = "tabla-densa"
    else:
        clase_densidad = "tabla-micro"

    # Calcular porcentajes proporcionales con piso mínimo razonable para que ninguna columna colapse
    pesos = [max(4.0, (w ** 0.72)) for w in anchos_col]
    suma = sum(pesos) or 1.0
    pcts = [round((p / suma) * 100.0, 2) for p in pesos]
    if pcts:
        pcts[-1] = round(100.0 - sum(pcts[:-1]), 2)

    colgroup = "<colgroup>" + "".join(f'<col style="width:{p}%">' for p in pcts) + "</colgroup>"
    if "<colgroup>" not in tabla_html.lower():
        tabla_html = re.sub(r"(<table\b[^>]*>)", lambda m: m.group(1) + colgroup, tabla_html, count=1, flags=re.I)

    m_tbody = re.search(r"<tbody\b[^>]*>(.*?)</tbody>", tabla_html, flags=re.I | re.DOTALL)
    if m_tbody:
        filas_cuerpo = re.findall(r"<tr\b[^>]*>.*?</tr>", m_tbody.group(1), flags=re.I | re.DOTALL)
    else:
        filas_cuerpo = filas[1:] if len(filas) > 1 else filas
    num_filas_cuerpo = len(filas_cuerpo)
    texto_total_tabla = html.unescape(re.sub(r"<[^>]+>", " ", m_table.group(1)))
    largo_texto_tabla = len(re.sub(r"\s+", " ", texto_total_tabla).strip())

    return tabla_html, num_cols, clase_densidad, num_filas_cuerpo, largo_texto_tabla


def empaquetar_html_autocontenido(html_str: str, carpeta_base: Path, ruta_css: Path) -> str:
    """
    Incrusta la hoja de estilos CSS en <style> y convierte todas las imágenes locales
    (logos y figuras) a data URIs Base64 para generar un único archivo HTML autocontenido.
    """
    if ruta_css.exists():
        css_contenido = ruta_css.read_text(encoding="utf-8").strip()
        html_str = re.sub(
            r'<link\s+rel=["\']stylesheet["\']\s+href=["\']rma_ojs\.css["\']\s*/?>',
            lambda _: f"<style>\n{css_contenido}\n    </style>",
            html_str,
            count=1,
            flags=re.I,
        )

    mime_map = {
        ".png": "image/png",
        ".jpg": "image/jpeg",
        ".jpeg": "image/jpeg",
        ".gif": "image/gif",
        ".webp": "image/webp",
        ".svg": "image/svg+xml",
    }

    def _reemplazar_src(match: re.Match) -> str:
        prefix, src_val, suffix = match.group(1), match.group(2).strip(), match.group(3)
        if not src_val or src_val.startswith(("http://", "https://", "data:", "//")):
            return match.group(0)
        ruta_img = (carpeta_base / src_val).resolve()
        if not ruta_img.exists() or not ruta_img.is_file():
            return match.group(0)
        ext = ruta_img.suffix.lower()
        mime = mime_map.get(ext) or mimetypes.guess_type(str(ruta_img))[0] or "image/png"
        b64 = base64.b64encode(ruta_img.read_bytes()).decode("ascii")
        return f'{prefix}data:{mime};base64,{b64}{suffix}'

    html_str = re.sub(
        r'(<img\b[^>]*?\bsrc=["\'])([^"\']+)(["\'][^>]*>)',
        _reemplazar_src,
        html_str,
        flags=re.I,
    )
    return html_str


def _xml_escape_text(valor: str) -> str:
    """Decodifica entidades HTML previas y escapa únicamente caracteres reservados XML (&, <, >)."""
    txt = html.unescape(str(valor or ""))
    return txt.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def _xml_escape_attr(valor: str) -> str:
    """Escapa texto para atributos XML (&, <, >, \")."""
    return _xml_escape_text(valor).replace('"', "&quot;")


def _consolidar_tags_inline_jats(jats_str: str) -> str:
    """Fusiona etiquetas JATS contiguas (<bold>, <italic>, <sup>, <sub>) y elimina etiquetas vacías."""
    if not jats_str:
        return ""
    s = str(jats_str)
    for _ in range(4):
        s = re.sub(r"</(italic|bold|sup|sub)>(\s*)<\1>", r"\2", s)
        s = re.sub(r"<(italic|bold|sup|sub)>(\s*)</\1>", r"\2", s)
    return s.strip()


def _limpiar_titulo_jats(jats_str: str) -> str:
    """
    Aplica las reglas de SciELO SPS 1.9 a títulos (<article-title>, <trans-title>, <sec>/<title>):
    elimina negritas (<bold>) y cursivas globales que envuelven todo el título, conservando
    <italic> únicamente cuando marca nombres científicos o términos puntuales en otro idioma.
    """
    s = _consolidar_tags_inline_jats(jats_str)
    if not s:
        return ""
    s = re.sub(r"</?bold>", "", s)
    s = _consolidar_tags_inline_jats(s)

    txt_total = len(re.sub(r"<[^>]+>", "", s).strip())
    if txt_total > 0:
        italic_spans = re.findall(r"<italic>(.*?)</italic>", s, flags=re.DOTALL)
        italic_chars = sum(len(re.sub(r"<[^>]+>", "", seg).strip()) for seg in italic_spans)
        if italic_chars / txt_total >= 0.85:
            s = re.sub(r"</?italic>", "", s)
    return _consolidar_tags_inline_jats(s)


def _limpiar_resumen_jats(jats_p_str: str) -> str:
    """
    Aplica las reglas de SciELO SPS 1.9 a párrafos de <abstract> y <trans-abstract>:
    elimina negritas y cursivas globales aplicadas a todo el párrafo en Word, conservando
    <italic> únicamente para nombres científicos (ej. Homo sapiens) o términos puntuales.
    """
    s = _consolidar_tags_inline_jats(jats_p_str)
    if not s:
        return ""
    s = re.sub(r"</?bold>", "", s)
    s = _consolidar_tags_inline_jats(s)

    txt_total = len(re.sub(r"<[^>]+>", "", s).strip())
    if txt_total > 0:
        italic_spans = re.findall(r"<italic>(.*?)</italic>", s, flags=re.DOTALL)
        italic_chars = sum(len(re.sub(r"<[^>]+>", "", seg).strip()) for seg in italic_spans)
        if italic_chars / txt_total >= 0.85:
            s = re.sub(r"</?italic>", "", s)
    return _consolidar_tags_inline_jats(s)


def _html_inline_a_jats(fragmento_html: str) -> str:
    """
    Convierte HTML inline (<em>, <strong>, <sup>, <sub>, <a href>, notas al pie)
    a etiquetas JATS 1.1 nativas (<italic>, <bold>, <sup>, <sub>, <ext-link>, <xref>)
    garantizando etiquetas balanceadas, consolidadas y cero doble escape de entidades.
    """
    if not fragmento_html:
        return ""
    s = _consolidar_tags_inline_html(str(fragmento_html))

    s = re.sub(
        r'<a\b[^>]*class=["\'][^"\']*fn-ref-link[^"\']*["\'][^>]*href=["\']#fn-(\d+)["\'][^>]*>(\d+)</a>\s*<span\b[^>]*class=["\'][^"\']*footnote[^"\']*["\'][^>]*>.*?</span>',
        lambda m: f"[[JATS_FN:{m.group(1)}:{m.group(2)}]]",
        s,
        flags=re.I | re.DOTALL,
    )
    s = re.sub(r"<br\s*/?>", " ", s, flags=re.I)

    tokens = re.split(r"(<[^>]+>)", s)
    out = []
    stack = []

    for tok in tokens:
        if not tok:
            continue
        if tok.startswith("<") and tok.endswith(">"):
            t_low = tok.lower().strip()
            if t_low.startswith("<!--"):
                continue
            es_cierre = t_low.startswith("</")
            m_tag = re.match(r"^</?\s*([a-z0-9]+)", t_low)
            if not m_tag:
                continue
            tag_name = m_tag.group(1)

            if not es_cierre:
                if tag_name in ("em", "i"):
                    out.append("<italic>")
                    stack.append("italic")
                elif tag_name in ("strong", "b"):
                    out.append("<bold>")
                    stack.append("bold")
                elif tag_name == "sup":
                    out.append("<sup>")
                    stack.append("sup")
                elif tag_name == "sub":
                    out.append("<sub>")
                    stack.append("sub")
                elif tag_name == "a":
                    m_href = re.search(r'\bhref\s*=\s*["\']([^"\']+)["\']', tok, flags=re.I)
                    href = html.unescape(m_href.group(1)).strip() if m_href else ""
                    if href.startswith("mailto:"):
                        out.append(f'<ext-link ext-link-type="email" xlink:href="{_xml_escape_attr(href)}">')
                        stack.append("ext-link")
                    elif href.startswith(("http://", "https://")):
                        out.append(f'<ext-link ext-link-type="uri" xlink:href="{_xml_escape_attr(href)}">')
                        stack.append("ext-link")
                    else:
                        stack.append(None)
            else:
                target_jats = None
                if tag_name in ("em", "i"):
                    target_jats = "italic"
                elif tag_name in ("strong", "b"):
                    target_jats = "bold"
                elif tag_name == "sup":
                    target_jats = "sup"
                elif tag_name == "sub":
                    target_jats = "sub"
                elif tag_name == "a":
                    if stack and stack[-1] is None:
                        stack.pop()
                        continue
                    target_jats = "ext-link"

                if target_jats and target_jats in stack:
                    while stack:
                        top = stack.pop()
                        if top:
                            out.append(f"</{top}>")
                        if top == target_jats:
                            break
        else:
            esc = _xml_escape_text(tok)
            esc = re.sub(
                r"\[\[JATS_FN:(\d+):(\d+)\]\]",
                lambda m: f'<xref ref-type="fn" rid="fn{m.group(1)}">{m.group(2)}</xref>',
                esc,
            )
            out.append(esc)

    while stack:
        top = stack.pop()
        if top:
            out.append(f"</{top}>")

    return _consolidar_tags_inline_jats("".join(out))


def _inferir_sec_type_jats(titulo_html: str) -> str:
    """Asigna el atributo @sec-type admitido por SciELO SPS 1.9 según el título de sección."""
    plain = _quitar_tildes(html.unescape(re.sub(r"<[^>]+>", "", titulo_html))).lower().strip()
    plain = re.sub(r"^\d+[\.\)\-\s]+", "", plain).strip()
    if any(k in plain for k in ("introduccion", "antecedentes", "marco teorico", "planteo del problema")):
        return "intro"
    if any(k in plain for k in ("materiales y metodos", "material y metodos", "metodologia", "area de estudio", "muestra y metodos", "materiales")):
        return "materials|methods"
    if plain.startswith("resultados"):
        return "results"
    if plain.startswith("discusion"):
        return "discussion"
    if any(k in plain for k in ("conclusion", "consideraciones finales", "palabras finales", "sintesis final")):
        return "conclusions"
    if plain.startswith("agradecimiento"):
        return "ack"
    return ""


def _parsear_fecha_dmy(fecha_str: str) -> tuple[str, str, str] | None:
    """Convierte 'DD-MM-AAAA', 'DD/MM/AAAA' o 'AAAA-MM-DD' en (day, month, year) de 2/2/4 dígitos."""
    s = str(fecha_str or "").strip()
    if not s:
        return None
    m1 = re.match(r"^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})$", s)
    if m1:
        d, mo, y = int(m1.group(1)), int(m1.group(2)), int(m1.group(3))
        if 1 <= d <= 31 and 1 <= mo <= 12:
            return f"{d:02d}", f"{mo:02d}", f"{y:04d}"
    m2 = re.match(r"^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})$", s)
    if m2:
        y, mo, d = int(m2.group(1)), int(m2.group(2)), int(m2.group(3))
        if 1 <= d <= 31 and 1 <= mo <= 12:
            return f"{d:02d}", f"{mo:02d}", f"{y:04d}"
    return None


def _detectar_pais_afiliacion_orig(texto_sin_email: str, emails: list[str], ror_id: str = "") -> tuple[str, str]:
    """Devuelve (codigo_iso2, nombre_pais) a partir del ROR ID, texto institucional o dominio de correo."""
    ror_code, _ = normalizar_ror(ror_id)
    ror_paises = {
        "03cqe8w59": ("AR", "Argentina"),  # CONICET
        "056tb7j80": ("AR", "Argentina"),  # UNC
        "033352052": ("AR", "Argentina"),  # UBA
        "01tjs6929": ("AR", "Argentina"),  # UNLP
        "020454351": ("UY", "Uruguay"),    # Universidad de la República
    }
    if ror_code in ror_paises:
        return ror_paises[ror_code]

    t_norm = _quitar_tildes(texto_sin_email).lower()
    paises = [
        (r"\bargentina\b", "AR", "Argentina"),
        (r"\buruguay\b", "UY", "Uruguay"),
        (r"\bchile\b", "CL", "Chile"),
        (r"\b(?:brasil|brazil)\b", "BR", "Brasil"),
        (r"\bperu\b", "PE", "Perú"),
        (r"\bmexico\b", "MX", "México"),
        (r"\bcolombia\b", "CO", "Colombia"),
        (r"\b(?:espana|spain)\b", "ES", "España"),
        (r"\bbolivia\b", "BO", "Bolivia"),
        (r"\becuador\b", "EC", "Ecuador"),
        (r"\bparaguay\b", "PY", "Paraguay"),
        (r"\bvenezuela\b", "VE", "Venezuela"),
        (r"\b(?:estados unidos|united states|usa|u\.s\.a\.)\b", "US", "Estados Unidos"),
        (r"\b(?:francia|france)\b", "FR", "Francia"),
        (r"\b(?:reino unido|united kingdom|inglaterra|england|uk)\b", "GB", "Reino Unido"),
        (r"\b(?:alemania|germany)\b", "DE", "Alemania"),
        (r"\b(?:italia|italy)\b", "IT", "Italia"),
        (r"\b(?:canada)\b", "CA", "Canadá"),
        (r"\bportugal\b", "PT", "Portugal"),
    ]
    for pat, code, nombre in paises:
        if re.search(pat, t_norm):
            return code, nombre

    tld_map = {
        ".ar": ("AR", "Argentina"),
        ".uy": ("UY", "Uruguay"),
        ".cl": ("CL", "Chile"),
        ".br": ("BR", "Brasil"),
        ".pe": ("PE", "Perú"),
        ".mx": ("MX", "México"),
        ".co": ("CO", "Colombia"),
        ".es": ("ES", "España"),
        ".bo": ("BO", "Bolivia"),
        ".ec": ("EC", "Ecuador"),
        ".py": ("PY", "Paraguay"),
        ".ve": ("VE", "Venezuela"),
        ".fr": ("FR", "Francia"),
        ".uk": ("GB", "Reino Unido"),
        ".de": ("DE", "Alemania"),
        ".it": ("IT", "Italia"),
        ".ca": ("CA", "Canadá"),
        ".pt": ("PT", "Portugal"),
    }
    for em in emails:
        em_low = em.lower().strip()
        for tld, pair in tld_map.items():
            if em_low.endswith(tld):
                return pair

    claves_ar = (
        "conicet",
        "consejo nacional de investigaciones cientificas y tecnicas",
        "inapl",
        "instituto nacional de antropologia",
        "idacor",
        "cicterra",
        "incuapa",
        "imhicihu",
        "icsoh",
        "indes",
        "ipehcs",
        "cigeobio",
        "ciafic",
        "universidad nacional de",
        "universidad de buenos aires",
        "museo etnografico",
        "museo de antropologia",
        "museo de la plata",
        "facultad de filosofia y letras",
        "facultad de filosofia y humanidades",
        "facultad de ciencias naturales y museo",
        "ciudad autonoma de buenos aires",
        "buenos aires",
        "cordoba",
        "rosario",
        "la plata",
        "tucuman",
        "mendoza",
        "olavarria",
        "mar del plata",
        "bahia blanca",
        "bariloche",
        "ushuaia",
        "catamarca",
        "san juan",
        "san luis",
        "jujuy",
        "salta",
        "santa fe",
        "parana",
        "posadas",
        "corrientes",
        "resistencia",
        "neuquen",
        "rio gallegos",
        "comodoro rivadavia",
    )
    if any(k in t_norm for k in claves_ar):
        return "AR", "Argentina"
    if any(k in t_norm for k in ("universidad de la republica", "udelar", "montevideo", "pedeciba")):
        return "UY", "Uruguay"
    if any(k in t_norm for k in ("universidad de chile", "pontificia universidad catolica de chile", "santiago de chile")):
        return "CL", "Chile"
    if any(k in t_norm for k in ("universidade de sao paulo", "rio de janeiro", "cnpq")):
        return "BR", "Brasil"
    if any(k in t_norm for k in ("universidad nacional autonoma de mexico", "unam", "inah", "ciudad de mexico")):
        return "MX", "México"

    return "", ""
_EXTRA_PAISES_RMA = {
    "malta": ("MT", "Malta"), "republic of malta": ("MT", "Malta"), "republica de malta": ("MT", "Malta"),
    "chipre": ("CY", "Chipre"), "cyprus": ("CY", "Cyprus"),
    "irlanda": ("IE", "Irlanda"), "ireland": ("Ireland", "IE"),
    "grecia": ("GR", "Grecia"), "greece": ("Greece", "GR"),
    "belgica": ("BE", "Bélgica"), "belgium": ("Belgium", "BE"),
    "paises bajos": ("NL", "Países Bajos"), "netherlands": ("Netherlands", "NL"),
    "suiza": ("CH", "Suiza"), "switzerland": ("Switzerland", "CH"),
    "austria": ("AT", "Austria"), "suecia": ("Suecia", "SE"), "sweden": ("Sweden", "SE"),
    "noruega": ("Noruega", "NO"), "norway": ("Norway", "NO"),
    "dinamarca": ("Dinamarca", "DK"), "denmark": ("Denmark", "DK"),
    "polonia": ("Polonia", "PL"), "poland": ("Poland", "PL"),
    "portugal": ("Portugal", "PT"), "italia": ("Italia", "IT"), "italy": ("Italy", "IT"),
    "francia": ("Francia", "FR"), "france": ("France", "FR"),
    "alemania": ("Alemania", "DE"), "germany": ("Germany", "DE"),
    "espana": ("España", "ES"), "spain": ("Spain", "ES"),
    "reino unido": ("Reino Unido", "GB"), "united kingdom": ("United Kingdom", "GB"),
    "uk": ("United Kingdom", "GB"), "england": ("United Kingdom", "GB"), "scotland": ("United Kingdom", "GB"),
    "estados unidos": ("Estados Unidos", "US"), "united states": ("United States", "US"),
    "united states of america": ("United States", "US"), "usa": ("United States", "US"),
    "canada": ("Canada", "CA"), "australia": ("Australia", "AU"),
    "nueva zelanda": ("Nueva Zelanda", "NZ"), "new zealand": ("New Zealand", "NZ"),
    "argentina": ("Argentina", "AR"), "brasil": ("Brasil", "BR"), "brazil": ("Brazil", "BR"),
    "chile": ("Chile", "CL"), "uruguay": ("Uruguay", "UY"), "paraguay": ("Paraguay", "PY"),
    "bolivia": ("Bolivia", "BO"), "peru": ("Perú", "PE"), "ecuador": ("Ecuador", "EC"),
    "colombia": ("Colombia", "CO"), "venezuela": ("Venezuela", "VE"), "mexico": ("México", "MX"),
}
_EXTRA_PISTAS_RMA = [
    ("MT", "Malta", ("malta", "university of malta", "msida", "valletta", "gozo", "heritage malta", "superintendence of cultural heritage", "birkirkara", "sliema", "mdina")),
    ("AR", "Argentina", ("conicet", "unc", "uba", "unlp", "unt", "unr", "unsa", "unmdp", "buenos aires", "cordoba", "tucuman", "salta", "mendoza", "santa fe", "la plata", "rosario", "idacor", "inapl")),
    ("US", "United States", ("arizona", "california", "texas", "new york", "new mexico", "colorado", "florida", "massachusetts", "illinois", "washington", "michigan", "pennsylvania", "virginia", "smithsonian", "harvard", "stanford", "berkeley", "ucla", "yale")),
    ("GB", "United Kingdom", ("oxford", "cambridge", "ucl", "london", "edinburgh", "manchester", "bristol", "durham", "sheffield", "york", "british museum")),
]

def _detectar_pais_afiliacion_ext(texto_af: str, emails=None, ror_id: str = "") -> tuple[str, str]:
    iso, nom = _detectar_pais_afiliacion_orig(texto_af, emails, ror_id=ror_id)
    if iso and nom:
        return iso, nom
    norm_full = "".join(c for c in unicodedata.normalize("NFD", (texto_af or "").lower()) if unicodedata.category(c) != "Mn")
    norm_full = re.sub(r"\s+", " ", re.sub(r"[^a-z0-9\s]", " ", norm_full)).strip()
    for p in reversed([x.strip() for x in re.split(r"[,.;()]", texto_af or "") if x.strip()][-5:]):
        pn = "".join(c for c in unicodedata.normalize("NFD", p.lower()) if unicodedata.category(c) != "Mn")
        pn = re.sub(r"\s+", " ", re.sub(r"[^a-z0-9\s]", " ", pn)).strip()
        if pn in _EXTRA_PAISES_RMA:
            return _EXTRA_PAISES_RMA[pn]
    padded = f" {norm_full} "
    for iso_c, nom_c, pistas in _EXTRA_PISTAS_RMA:
        if any(f" {pt} " in padded for pt in pistas):
            return iso_c, nom_c
    for em in (emails or []):
        if str(em).lower().endswith(".mt"):
            return "MT", "Malta"
    return "", ""

_detectar_pais_afiliacion = _detectar_pais_afiliacion_ext


def _desglosar_afiliacion_jats(clave_af: str, txt_af_html: str, ror_id: str, indent: str = "      ") -> str:
    """Construye el bloque <aff> compatible con SciELO SPS 1.9 separando <institution>, <country> y <email>."""
    nl = chr(10)
    limpio_html, _, rors_inline = _limpiar_urls_orcid_ror_en_html(txt_af_html)
    orig_txt = html.unescape(re.sub(r"<[^>]+>", "", limpio_html)).strip()
    orig_txt = re.sub(rf"^{re.escape(clave_af)}\s*", "", orig_txt).strip()

    if not ror_id and rors_inline:
        ror_id = rors_inline[0]

    emails = re.findall(r"[\w.+-]+@[\w-]+(?:\.[\w-]+)+", orig_txt)
    sin_email = re.sub(
        r"(?:[,;.\s]*\b(?:e-?mail|correo(?:\s+electr[oó]nico)?)\s*:?\s*)?[\w.+-]+@[\w-]+(?:\.[\w-]+)+\.?",
        "",
        orig_txt,
        flags=re.I,
    )
    sin_email = re.sub(r"\s{2,}", " ", sin_email).strip(" .,;:-")

    iso_pais, nom_pais = _detectar_pais_afiliacion(sin_email, emails, ror_id=ror_id)
    orgname = sin_email
    if nom_pais:
        orgname = re.sub(
            rf"(?:,\s*|\s*-\s*|\s+de\s+){re.escape(nom_pais)}\.?$",
            "",
            orgname,
            flags=re.I,
        ).strip(" .,;:-")
    if not orgname:
        orgname = sin_email or orig_txt

    partes = [
        f'{indent}<aff id="aff{_xml_escape_attr(clave_af)}">',
        f'{indent}  <label>{_xml_escape_text(clave_af)}</label>',
        f'{indent}  <institution content-type="original">{_xml_escape_text(orig_txt)}</institution>',
    ]
    rors_list = normalizar_rors(ror_id)
    if rors_list:
        partes.append(f'{indent}  <institution-wrap>')
        for _, ror_url in rors_list:
            partes.append(f'{indent}    <institution-id institution-id-type="ror">{_xml_escape_text(ror_url)}</institution-id>')
        partes.append(f'{indent}    <institution content-type="orgname">{_xml_escape_text(orgname)}</institution>')
        partes.append(f'{indent}  </institution-wrap>')
    else:
        partes.append(f'{indent}  <institution content-type="orgname">{_xml_escape_text(orgname)}</institution>')

    if iso_pais and nom_pais:
        partes.append(f'{indent}  <country country="{_xml_escape_attr(iso_pais)}">{_xml_escape_text(nom_pais)}</country>')
    for em in emails:
        partes.append(f'{indent}  <email>{_xml_escape_text(em)}</email>')

    partes.append(f'{indent}</aff>')
    return nl.join(partes)


def _html_tabla_a_jats(tabla_html: str, indent: str = "        ") -> str:
    """Convierte una tabla HTML a <table> JATS 1.1 válida conservando thead, tbody, colspan y estilos inline."""
    nl = chr(10)
    m_table = re.search(r"<table\b[^>]*>(.*?)</table>", tabla_html, flags=re.I | re.DOTALL)
    if not m_table:
        txt_fb = _xml_escape_text(re.sub(r"<[^>]+>", " ", tabla_html).strip() or "Tabla sin datos")
        return f"{indent}<table>{nl}{indent}  <tbody>{nl}{indent}    <tr><td>{txt_fb}</td></tr>{nl}{indent}  </tbody>{nl}{indent}</table>"

    inner = m_table.group(1)

    def _convertir_filas(bloque_html: str) -> list[str]:
        filas_out = []
        for tr_m in re.finditer(r"<tr\b[^>]*>(.*?)</tr>", bloque_html, flags=re.I | re.DOTALL):
            celdas_out = []
            for cell_m in re.finditer(r"<(th|td)\b([^>]*)>(.*?)</\1>", tr_m.group(1), flags=re.I | re.DOTALL):
                ctag = cell_m.group(1).lower()
                cattrs = cell_m.group(2)
                cbody = cell_m.group(3)
                attrs_jats = ""
                m_cs = re.search(r'\bcolspan\s*=\s*["\']?(\d+)', cattrs, flags=re.I)
                if m_cs and int(m_cs.group(1)) > 1:
                    attrs_jats += f' colspan="{int(m_cs.group(1))}"'
                m_rs = re.search(r'\browspan\s*=\s*["\']?(\d+)', cattrs, flags=re.I)
                if m_rs and int(m_rs.group(1)) > 1:
                    attrs_jats += f' rowspan="{int(m_rs.group(1))}"'
                cbody_br = re.sub(r"<br\s*/?>", "[[JATS_BREAK]]", cbody, flags=re.I)
                c_jats = _html_inline_a_jats(cbody_br).replace("[[JATS_BREAK]]", "<break/>")
                celdas_out.append(f"<{ctag}{attrs_jats}>{c_jats}</{ctag}>")
            if celdas_out:
                filas_out.append(f"{indent}    <tr>{''.join(celdas_out)}</tr>")
        return filas_out

    m_thead = re.search(r"<thead\b[^>]*>(.*?)</thead>", inner, flags=re.I | re.DOTALL)
    m_tbody = re.search(r"<tbody\b[^>]*>(.*?)</tbody>", inner, flags=re.I | re.DOTALL)

    partes = [f"{indent}<table>"]
    if m_thead:
        f_head = _convertir_filas(m_thead.group(1))
        if f_head:
            partes.append(f"{indent}  <thead>")
            partes.extend(f_head)
            partes.append(f"{indent}  </thead>")
    if m_tbody:
        f_body = _convertir_filas(m_tbody.group(1))
        if f_body:
            partes.append(f"{indent}  <tbody>")
            partes.extend(f_body)
            partes.append(f"{indent}  </tbody>")
    if not m_thead and not m_tbody:
        f_all = _convertir_filas(inner)
        if f_all:
            partes.append(f"{indent}  <tbody>")
            partes.extend(f_all)
            partes.append(f"{indent}  </tbody>")
    partes.append(f"{indent}</table>")
    return nl.join(partes)


def _parsear_autores_cita_jats(autores_str: str, group_type: str = "author", indent: str = "          ") -> str:
    """Convierte una cadena de autores/editores en <person-group person-group-type='...'> sin texto suelto."""
    nl = chr(10)
    s = autores_str.strip()
    if not s:
        return ""
    if "," not in s and not re.search(r"\b[A-ZÁÉÍÓÚÑ]\.", s):
        return (
            f'{indent}<person-group person-group-type="{group_type}">{nl}'
            f'{indent}  <collab>{_xml_escape_text(s.rstrip("."))}</collab>{nl}'
            f'{indent}</person-group>'
        )

    s_norm = re.sub(r"\s+(?:y|and)\s+(?=[A-ZÁÉÍÓÚÑ][a-záéíóúñ'\-]+)", "; ", s)
    s_norm = re.sub(r"\s*&\s*(?=[A-ZÁÉÍÓÚÑ][a-záéíóúñ'\-]+)", "; ", s_norm)

    if ";" in s_norm:
        chunks = [c.strip() for c in s_norm.split(";") if c.strip()]
    elif group_type == "editor" and "," not in s_norm:
        chunks = [c.strip() for c in re.split(r"\s*,\s*", s_norm) if c.strip()]
    else:
        chunks = [
            c.strip()
            for c in re.split(r"(?<=\.)\s*,\s*(?=[A-ZÁÉÍÓÚÑ][a-záéíóúñ'\-]+(?:\s+[A-Za-zÁÉÍÓÚÑáéíóúñ'\-]+)*,)", s_norm)
            if c.strip()
        ]

    names_xml = []
    for ch in chunks:
        ch_clean = ch.strip().rstrip(",")
        if not ch_clean:
            continue
        if "," in ch_clean:
            ap, nom = ch_clean.split(",", 1)
            ap = ap.strip()
            nom = nom.strip()
            if ap and nom:
                names_xml.append(
                    f"{indent}  <name><surname>{_xml_escape_text(ap)}</surname>"
                    f"<given-names>{_xml_escape_text(nom)}</given-names></name>"
                )
            elif ap:
                names_xml.append(f"{indent}  <collab>{_xml_escape_text(ap)}</collab>")
        else:
            # Formato "N. Apellido" o "Nombre Apellido" (típico de editores en capítulos)
            partes_p = ch_clean.split()
            if len(partes_p) >= 2 and re.match(r"^[A-ZÁÉÍÓÚÑ]\.", partes_p[0]):
                nom = " ".join(partes_p[:-1]).strip()
                ap = partes_p[-1].strip()
                names_xml.append(
                    f"{indent}  <name><surname>{_xml_escape_text(ap)}</surname>"
                    f"<given-names>{_xml_escape_text(nom)}</given-names></name>"
                )
            else:
                names_xml.append(f"{indent}  <collab>{_xml_escape_text(ch_clean)}</collab>")

    if not names_xml:
        return (
            f'{indent}<person-group person-group-type="{group_type}">{nl}'
            f'{indent}  <collab>{_xml_escape_text(s)}</collab>{nl}'
            f'{indent}</person-group>'
        )
    return f'{indent}<person-group person-group-type="{group_type}">{nl}{nl.join(names_xml)}{nl}{indent}</person-group>'


def _formatear_referencia_jats(idx: int, r_html: str) -> str:
    """
    Genera <ref> compatible con SciELO SPS 1.9:
      - <label> únicamente si la referencia original está numerada explícitamente
      - <mixed-citation> con el texto completo formateado (sin atributo publication-type)
      - <element-citation publication-type="..."> puramente estructurado, sin texto ni puntuación suelta
    """
    nl = chr(10)
    mixed_jats = _html_inline_a_jats(r_html)
    txt = html.unescape(re.sub(r"<[^>]+>", "", r_html)).strip()
    txt = re.sub(r"\s+", " ", txt)

    label_xml = ""
    m_num = re.match(r"^(?:\[(\d+)\]|(\d+)[\.\)])\s+(.+)$", txt)
    if m_num:
        num_lbl = m_num.group(1) or m_num.group(2)
        label_xml = f"        <label>{_xml_escape_text(num_lbl)}</label>{nl}"
        txt = m_num.group(3).strip()

    doi_val = ""
    url_val = ""
    m_doi = re.search(r"\s*(?:https?://(?:dx\.)?doi\.org/|doi:\s*)(10\.\d{4,9}/[^\s<>\"\'\)\]]+)\.?\s*$", txt, flags=re.I)
    if m_doi:
        doi_val = m_doi.group(1).rstrip(".")
        txt = txt[: m_doi.start()].strip()
    else:
        m_url = re.search(r"\s*(?:Disponible en:\s*|Recuperado de:\s*)?(https?://[^\s<>\"\'\)\]]+)\.?\s*$", txt, flags=re.I)
        if m_url:
            url_val = m_url.group(1).rstrip(".")
            txt = txt[: m_url.start()].strip()

    ec_lines = []
    pub_type = "other"

    m_year = re.search(r"\(\s*(?:\[(\d{4})\]\s*)?(\d{4}[a-z]?|s\.f\.|en prensa|in press)\s*\)\.?\s*", txt, flags=re.I)
    if not m_year:
        if txt:
            ec_lines.append(f"          <source>{_xml_escape_text(txt.rstrip('.'))}</source>")
        if doi_val:
            ec_lines.append(f'          <pub-id pub-id-type="doi">{_xml_escape_text(doi_val)}</pub-id>')
        if url_val:
            ec_lines.append(f'          <ext-link ext-link-type="uri" xlink:href="{_xml_escape_attr(url_val)}">{_xml_escape_text(url_val)}</ext-link>')
        if not ec_lines:
            ec_lines.append("          <comment>Referencia bibliográfica</comment>")
        return (
            f'      <ref id="B{idx}">{nl}'
            f'{label_xml}'
            f'        <mixed-citation>{mixed_jats}</mixed-citation>{nl}'
            f'        <element-citation publication-type="other">{nl}'
            f'{nl.join(ec_lines)}{nl}'
            f'        </element-citation>{nl}'
            f'      </ref>'
        )

    autores_raw = txt[: m_year.start()].strip()
    pub_yr_raw = m_year.group(2).strip()
    resto = txt[m_year.end() :].strip()

    pg_author = _parsear_autores_cita_jats(autores_raw, "author", indent="          ")
    if pg_author:
        ec_lines.append(pg_author)

    m_yr_digits = re.match(r"^(\d{4})", pub_yr_raw)
    if m_yr_digits:
        ec_lines.append(f"          <year>{m_yr_digits.group(1)}</year>")
    elif pub_yr_raw:
        ec_lines.append(f"          <comment>{_xml_escape_text(pub_yr_raw)}</comment>")

    # 1) Capítulo de libro
    m_cap = re.match(
        r"^(.+?)\.\s+(En|In)\s+(.+?)\s*\((Ed\.|Eds\.|Comp\.|Comps\.|Coord\.|Coords\.|Org\.|Orgs\.)\),?\s*(.+)$",
        resto,
    )
    if m_cap:
        pub_type = "book"
        cap_title = m_cap.group(1).strip()
        eds_raw = m_cap.group(3).strip()
        libro_resto = m_cap.group(5).strip()

        ec_lines.append(f"          <chapter-title>{_xml_escape_text(cap_title)}</chapter-title>")
        pg_ed = _parsear_autores_cita_jats(eds_raw, "editor", indent="          ")
        if pg_ed:
            ec_lines.append(pg_ed)

        m_lib_det = re.match(
            r"^(.+?)(?:\s*\(pp?\.\s*(\d+)\s*[-–—]\s*(\d+)\))?\.\s*([^.:]+?)\s*:\s*([^.]+?)\.?$",
            libro_resto,
        )
        if m_lib_det:
            lib_tit = m_lib_det.group(1).strip()
            fp, lp = m_lib_det.group(2), m_lib_det.group(3)
            loc = m_lib_det.group(4).strip()
            pub_name = m_lib_det.group(5).strip()
            ec_lines.append(f"          <source>{_xml_escape_text(lib_tit)}</source>")
            if fp:
                ec_lines.append(f"          <fpage>{_xml_escape_text(fp)}</fpage>")
            if lp:
                ec_lines.append(f"          <lpage>{_xml_escape_text(lp)}</lpage>")
            if loc:
                ec_lines.append(f"          <publisher-loc>{_xml_escape_text(loc)}</publisher-loc>")
            if pub_name:
                ec_lines.append(f"          <publisher-name>{_xml_escape_text(pub_name)}</publisher-name>")
        else:
            ec_lines.append(f"          <source>{_xml_escape_text(libro_resto.rstrip('.'))}</source>")
    else:
        # 2) Artículo de revista (páginas numéricas o elocation-id)
        m_art = re.match(
            r"^(.+?)\.\s+([^.,]+?),\s*(\d+)(?:\s*\(([^)]+)\))?\s*[:,]\s*(?:(\d+)\s*[-–—]\s*(\d+)|(e\d+))\.?$",
            resto,
        )
        if m_art:
            pub_type = "journal"
            art_tit = m_art.group(1).strip()
            rev_tit = m_art.group(2).strip()
            vol_v = m_art.group(3).strip()
            iss_v = (m_art.group(4) or "").strip()
            fp_v = (m_art.group(5) or "").strip()
            lp_v = (m_art.group(6) or "").strip()
            eloc_v = (m_art.group(7) or "").strip()

            ec_lines.append(f"          <article-title>{_xml_escape_text(art_tit)}</article-title>")
            ec_lines.append(f"          <source>{_xml_escape_text(rev_tit)}</source>")
            if vol_v:
                ec_lines.append(f"          <volume>{_xml_escape_text(vol_v)}</volume>")
            if iss_v:
                ec_lines.append(f"          <issue>{_xml_escape_text(iss_v)}</issue>")
            if fp_v:
                ec_lines.append(f"          <fpage>{_xml_escape_text(fp_v)}</fpage>")
            if lp_v:
                ec_lines.append(f"          <lpage>{_xml_escape_text(lp_v)}</lpage>")
            if eloc_v:
                ec_lines.append(f"          <elocation-id>{_xml_escape_text(eloc_v)}</elocation-id>")
        else:
            # 3) Tesis
            if re.search(r"\b(?:tesis|dissertation|thesis)\b", resto, flags=re.I):
                pub_type = "thesis"
                if ". " in resto:
                    t_part, r_part = resto.split(". ", 1)
                    ec_lines.append(f"          <source>{_xml_escape_text(t_part.strip().rstrip('.'))}</source>")
                    if r_part.strip():
                        ec_lines.append(f"          <publisher-name>{_xml_escape_text(r_part.strip().rstrip('.'))}</publisher-name>")
                else:
                    ec_lines.append(f"          <source>{_xml_escape_text(resto.rstrip('.'))}</source>")
            else:
                # 4) Libro completo
                m_book = re.match(r"^(.+?)\.\s+([^.:]+?)\s*:\s*([^.]+?)\.?$", resto)
                if m_book:
                    pub_type = "book"
                    book_tit = m_book.group(1).strip()
                    loc = m_book.group(2).strip()
                    pub_name = m_book.group(3).strip()
                    ec_lines.append(f"          <source>{_xml_escape_text(book_tit)}</source>")
                    if loc:
                        ec_lines.append(f"          <publisher-loc>{_xml_escape_text(loc)}</publisher-loc>")
                    if pub_name:
                        ec_lines.append(f"          <publisher-name>{_xml_escape_text(pub_name)}</publisher-name>")
                else:
                    if url_val:
                        pub_type = "webpage"
                    if ". " in resto:
                        t_part, r_part = resto.split(". ", 1)
                        ec_lines.append(f"          <source>{_xml_escape_text(t_part.strip().rstrip('.'))}</source>")
                        if r_part.strip():
                            ec_lines.append(f"          <comment>{_xml_escape_text(r_part.strip())}</comment>")
                    elif resto:
                        ec_lines.append(f"          <source>{_xml_escape_text(resto.rstrip('.'))}</source>")

    if doi_val:
        ec_lines.append(f'          <pub-id pub-id-type="doi">{_xml_escape_text(doi_val)}</pub-id>')
    if url_val:
        ec_lines.append(f'          <ext-link ext-link-type="uri" xlink:href="{_xml_escape_attr(url_val)}">{_xml_escape_text(url_val)}</ext-link>')

    if not ec_lines:
        ec_lines.append(f"          <source>{_xml_escape_text(txt)}</source>")

    return (
        f'      <ref id="B{idx}">{nl}'
        f'{label_xml}'
        f'        <mixed-citation>{mixed_jats}</mixed-citation>{nl}'
        f'        <element-citation publication-type="{pub_type}">{nl}'
        f'{nl.join(ec_lines)}{nl}'
        f'        </element-citation>{nl}'
        f'      </ref>'
    )


def validar_xml_scielo(xml_str: str) -> list[str]:
    """Verifica buena formación XML y reglas estructurales clave de SciELO Publishing Schema (SPS 1.9)."""
    avisos = []
    try:
        root = ET.fromstring(xml_str)
    except ET.ParseError as e:
        return [f"Error de sintaxis XML: {e}"]

    if "&amp;#" in xml_str or "&amp;amp;" in xml_str:
        avisos.append("Se detectó doble escape de entidades (&amp;# o &amp;amp;).")
    if "&lt;strong&gt;" in xml_str or "&lt;em&gt;" in xml_str:
        avisos.append("Se detectaron etiquetas HTML escapadas dentro del XML.")

    doi_el = root.find(".//article-meta/article-id[@pub-id-type='doi']")
    if doi_el is None or not (doi_el.text or "").strip():
        avisos.append('Falta <article-id pub-id-type="doi"> en <article-meta>.')
    elif doi_el.text.strip().endswith("/xxxxxx"):
        avisos.append('El DOI en <article-id pub-id-type="doi"> es provisional; actualizalo en metadatos.yaml antes de subir a SciELO.')

    abs_el = root.find(".//article-meta/abstract")
    if abs_el is not None and "{http://www.w3.org/XML/1998/namespace}lang" in abs_el.attrib:
        avisos.append("El elemento <abstract> principal no debe llevar atributo xml:lang en SciELO SPS 1.9.")

    for path_expr, label in (
        (".//title-group/article-title", "<article-title>"),
        (".//title-group//trans-title", "<trans-title>"),
        (".//article-meta/abstract", "<abstract>"),
        (".//article-meta/trans-abstract", "<trans-abstract>"),
    ):
        for el in root.findall(path_expr):
            if el.find(".//bold") is not None:
                avisos.append(f"Se detectó etiqueta <bold> no permitida dentro de {label}.")

    for cont_tag in ("fig", "table-wrap"):
        for cont in root.findall(f".//{cont_tag}"):
            if len(cont.findall("caption")) > 1:
                cid = cont.attrib.get("id", "?")
                avisos.append(
                    f'El elemento <{cont_tag} id="{cid}"> contiene múltiples etiquetas <caption>; SciELO SPS 1.9 admite como máximo una.'
                )

    for cap_el in root.findall(".//caption"):
        title_cap = cap_el.find("title")
        if title_cap is None:
            avisos.append("Un elemento <caption> carece de la etiqueta obligatoria <title>.")
        else:
            txt_cap_title = "".join(title_cap.itertext()).strip()
            if re.match(
                r"^(?:figuras?|figures?|figs?\.?|tablas?|tables?|tabs?\.?|cuadros?)\s*0*\d+[\.\:\-\–\—]?",
                txt_cap_title,
                flags=re.I,
            ):
                avisos.append(
                    f'El título de epígrafe "{txt_cap_title[:45]}..." conserva un prefijo redundante de figura/tabla.'
                )
        for p_cap in cap_el.findall("p"):
            txt_p_cap = "".join(p_cap.itertext()).strip()
            if re.match(r"^Description\s+of\s+(?:Figure|Table)\s+0*\d+\.?$", txt_p_cap, flags=re.I):
                avisos.append(
                    f'El epígrafe contiene texto marcador sin traducir ("{txt_p_cap}").'
                )

    for ref_el in root.findall(".//ref-list/ref"):
        rid = ref_el.attrib.get("id", "?")
        if ref_el.find("element-citation") is None:
            avisos.append(f'Falta <element-citation> en <ref id="{rid}">.')

    for aff_el in root.findall(".//article-meta/aff"):
        if aff_el.find("country") is None:
            aid = aff_el.attrib.get("id", "?")
            avisos.append(f'Falta <country country="XX"> en <aff id="{aid}">.')

    body = root.find("body")
    if body is None or len(list(body)) == 0:
        avisos.append("El elemento <body> está vacío.")
    else:
        txt_body = "".join(body.itertext())
        if "Completar secciones desde el HTML generado" in txt_body:
            avisos.append("El cuerpo del XML todavía contiene texto marcador.")

    fpage_el = root.find(".//article-meta/fpage")
    if fpage_el is not None and fpage_el.text and len(fpage_el.text) > 1 and fpage_el.text.startswith("0"):
        avisos.append(f"La etiqueta <fpage> ({fpage_el.text}) conserva ceros a la izquierda.")

    if root.find(".//article-meta/permissions") is None:
        avisos.append("Falta el bloque obligatorio <permissions>.")

    for inst in root.findall(".//aff//institution[@content-type='orgname']"):
        itxt = "".join(inst.itertext())
        if "@" in itxt:
            avisos.append("Hay un correo electrónico mezclado dentro de <institution content-type='orgname'>.")
            break

    return avisos


def generar_xml_scielo(
    art: dict,
    meta: dict,
    rev: dict,
    figuras_usadas: list,
    autores_est: list | None = None,
    afiliaciones_est: list | None = None,
    cuerpo_bloques: list | None = None,
) -> str:
    """Genera un archivo XML JATS 1.1 completo compatible con SciELO Publishing Schema (SPS 1.9)."""
    idioma_art = normalizar_codigo_idioma((meta or {}).get("idioma", "")) or detectar_idioma_texto(f"{(art or {}).get('titulo_es', '')} {(art or {}).get('titulo_en', '')} {(art or {}).get('resumen', '')} {(art or {}).get('abstract', '')}")
    idioma_local_xml = "pt" if (idioma_art == "pt" or (art or {}).get("idioma_resumen_local") == "pt") else "es"
    nl = chr(10)

    titulo_es_jats = _limpiar_titulo_jats(_html_inline_a_jats(art.get("titulo_es", "")))
    titulo_en_jats = _limpiar_titulo_jats(_html_inline_a_jats(art.get("titulo_en", "")))

    def _parrafos_resumen_jats(html_res: str, indent: str = "        ") -> str:
        pars = re.findall(r"<p\b[^>]*>(.*?)</p>", str(html_res or ""), flags=re.I | re.DOTALL)
        if not pars:
            limpio = _limpiar_resumen_jats(_html_inline_a_jats(html_res))
            return f"{indent}<p>{limpio}</p>" if limpio else ""
        salida_pars = []
        for p in pars:
            p_limpio = _limpiar_resumen_jats(_html_inline_a_jats(p))
            if p_limpio:
                salida_pars.append(f"{indent}<p>{p_limpio}</p>")
        return nl.join(salida_pars)

    resumen_jats = _parrafos_resumen_jats(art.get("resumen", ""))
    abstract_jats = _parrafos_resumen_jats(art.get("abstract", ""))

    def _extraer_keywords_jats(html_kw: str) -> list[str]:
        plain = html.unescape(re.sub(r"<[^>]+>", "", str(html_kw or ""))).strip()
        if ":" in plain:
            plain = plain.split(":", 1)[1].strip()
        sep = ";" if ";" in plain else ","
        items = []
        for k in plain.split(sep):
            k_clean = k.strip().rstrip(".")
            if k_clean:
                items.append(_xml_escape_text(k_clean))
        return items

    kw_es_list = _extraer_keywords_jats(art.get("palabras_clave", ""))
    kw_en_list = _extraer_keywords_jats(art.get("keywords", ""))
    kw_es_xml = nl.join(f"        <kwd>{k}</kwd>" for k in kw_es_list)
    kw_en_xml = nl.join(f"        <kwd>{k}</kwd>" for k in kw_en_list)

    contribs_xml = []
    for aut in (autores_est or []):
        nombre_full = html.unescape(str(aut.get("nombre", ""))).strip()
        if not nombre_full:
            continue
        surname = html.unescape(str(aut.get("apellidos", ""))).strip()
        given = html.unescape(str(aut.get("nombres", ""))).strip()
        if not surname:
            partes_n = nombre_full.split()
            given = " ".join(partes_n[:-1]) if len(partes_n) > 1 else nombre_full
            surname = partes_n[-1] if len(partes_n) > 1 else ""
        _, orcid_url = normalizar_orcid(aut.get("orcid", ""))
        orcid_tag = f'{nl}          <contrib-id contrib-id-type="orcid">{_xml_escape_text(orcid_url)}</contrib-id>' if orcid_url else ""
        refs_af = [x.strip() for x in str(aut.get("afiliacion", "")).split(",") if x.strip()]
        xrefs_tag = "".join(
            f'{nl}          <xref ref-type="aff" rid="aff{_xml_escape_attr(r)}">{_xml_escape_text(r)}</xref>'
            for r in refs_af
        )
        contribs_xml.append(
            f'        <contrib contrib-type="author">{orcid_tag}{nl}'
            f'          <name>{nl}'
            f'            <surname>{_xml_escape_text(surname)}</surname>{nl}'
            f'            <given-names>{_xml_escape_text(given)}</given-names>{nl}'
            f'          </name>{xrefs_tag}{nl}'
            f'        </contrib>'
        )

    affs_xml = []
    for idx_af, af in enumerate(afiliaciones_est or [], start=1):
        clave_af = str(af.get("clave", idx_af)).strip() or str(idx_af)
        txt_af = str(af.get("texto", "")).strip()
        ror_id = str(af.get("ror", "")).strip()
        affs_xml.append(_desglosar_afiliacion_jats(clave_af, txt_af, ror_id, indent="      "))

    contrib_group_block = (
        f'{nl}      <contrib-group>{nl}{nl.join(contribs_xml)}{nl}      </contrib-group>'
        if contribs_xml
        else ""
    )
    affs_block = (nl + nl.join(affs_xml)) if affs_xml else ""

    seccion_str = str(meta.get("seccion", "Artículos")).strip()
    sec_low = _quitar_tildes(seccion_str).lower()
    if "editorial" in sec_low:
        article_type = "editorial"
    elif "resena" in sec_low:
        article_type = "book-review"
    elif sec_low.startswith("nota"):
        article_type = "brief-report"
    elif "foro" in sec_low:
        article_type = "discussion"
    else:
        article_type = "research-article"

    art_id_str = str(meta.get("id_ojs", "")).strip()
    doi_limpio = re.sub(r"^https?://(dx\.)?doi\.org/", "", str(meta.get("doi", "")).strip()).strip()
    doi_limpio = re.sub(r"^doi:\s*", "", doi_limpio, flags=re.I).strip()
    if re.match(r"^10\.\d{4,9}/\S+$", doi_limpio):
        doi_xml_val = doi_limpio
    elif re.match(r"^10\.\d{4,9}/?$", doi_limpio):
        pref = doi_limpio.rstrip("/")
        sufijo_prov = re.sub(r"[^A-Za-z0-9._-]", "", art_id_str) or "xxxxxx"
        doi_xml_val = f"{pref}/{sufijo_prov}"
        print(f"  [Aviso XML SciELO] DOI sin sufijo definitivo ('{doi_limpio}'); se usó '{doi_xml_val}' en <article-id pub-id-type=\"doi\">.")
    else:
        sufijo_prov = re.sub(r"[^A-Za-z0-9._-]", "", art_id_str) or "xxxxxx"
        doi_xml_val = f"10.31048/{sufijo_prov}"
        print(f"  [Aviso XML SciELO] DOI ausente o no estándar; se usó '{doi_xml_val}' en <article-id pub-id-type=\"doi\">.")
    doi_tag = f'{nl}      <article-id pub-id-type="doi">{_xml_escape_text(doi_xml_val)}</article-id>'

    fp_raw = str(meta.get("fpage_str", meta.get("pagina_inicio", "1"))).strip()
    lp_raw = str(meta.get("lpage_str", "")).strip()
    fpage_num = str(int(fp_raw)) if fp_raw.isdigit() and int(fp_raw) > 0 else (fp_raw.lstrip("0") or "1")
    lpage_num = str(int(lp_raw)) if lp_raw.isdigit() and int(lp_raw) > 0 else lp_raw.lstrip("0")
    lpage_tag = f'{nl}      <lpage>{_xml_escape_text(lpage_num)}</lpage>' if lpage_num else ""

    anio_str = str(meta.get("anio", "")).strip() or "2026"
    num_rev_str = str(meta.get("numero", "")).strip()
    pub_dmy = _parsear_fecha_dmy(meta.get("fecha_publicacion", ""))
    if not pub_dmy:
        ac_dmy = _parsear_fecha_dmy(meta.get("aceptado", ""))
        mes_num = {"1": "04", "2": "08", "3": "12"}.get(num_rev_str, ac_dmy[1] if ac_dmy else "06")
        pub_dmy = ("01", mes_num, f"{int(anio_str):04d}" if anio_str.isdigit() else anio_str)

    hist_fechas = []
    for campo_m, dtype in (("recibido", "received"), ("correcciones", "rev-recd"), ("aceptado", "accepted")):
        dmy = _parsear_fecha_dmy(meta.get(campo_m, ""))
        if dmy:
            hist_fechas.append(
                f'        <date date-type="{dtype}">{nl}'
                f'          <day>{dmy[0]}</day>{nl}'
                f'          <month>{dmy[1]}</month>{nl}'
                f'          <year>{dmy[2]}</year>{nl}'
                f'        </date>'
            )
    history_block = f'{nl}      <history>{nl}{nl.join(hist_fechas)}{nl}      </history>' if hist_fechas else ""

    titulo_rev = str(rev.get("titulo_revista", "Revista del Museo de Antropología")).strip()
    licencia_txt = str(rev.get("licencia", "CC BY-NC-SA 4.0")).strip()
    url_lic = str(rev.get("url_licencia", "https://creativecommons.org/licenses/by-nc-sa/4.0/")).strip()
    permissions_block = (
        f'{nl}      <permissions>{nl}'
        f'        <copyright-statement>Copyright © {_xml_escape_text(anio_str)} {_xml_escape_text(titulo_rev)}</copyright-statement>{nl}'
        f'        <copyright-year>{_xml_escape_text(anio_str)}</copyright-year>{nl}'
        f'        <copyright-holder>{_xml_escape_text(titulo_rev)}</copyright-holder>{nl}'
        f'        <license license-type="open-access" xlink:href="{_xml_escape_attr(url_lic)}" xml:lang="{idioma_art}">{nl}'
        f'          <license-p>{"This is an open-access article distributed under the terms of the Creative Commons Attribution License" if idioma_art == "en" else ("Este é um artigo publicado em acesso aberto sob uma licença Creative Commons" if idioma_art == "pt" else "Este es un artículo publicado en acceso abierto bajo una licencia Creative Commons")} ({_xml_escape_text(licencia_txt)}).</license-p>{nl}'
        f'        </license>{nl}'
        f'      </permissions>'
    )

    body_lines = []
    en_seccion = False

    def _construir_captions_jats(cap_es_h: str, cap_en_h: str, indent: str) -> str:
        c_es = _limpiar_titulo_jats(_html_inline_a_jats(_limpiar_prefijo_recurso_html(cap_es_h)))
        txt_en_plano = re.sub(r"<[^>]+>", "", cap_en_h or "").strip()
        if re.match(r"^Description\s+of\s+(?:Figure|Table)\s+0*\d+\.?$", txt_en_plano, flags=re.I):
            c_en = ""
        else:
            c_en = _limpiar_resumen_jats(_html_inline_a_jats(_limpiar_prefijo_recurso_html(cap_en_h)))
            if re.match(r"^Description\s+of\s+(?:Figure|Table)\s+0*\d+\.?$", re.sub(r"<[^>]+>", "", c_en).strip(), flags=re.I):
                c_en = ""

        if not c_es and not c_en:
            return ""
        if c_es:
            lineas_cap = [
                f'{indent}  <caption>',
                f'{indent}    <title>{c_es}</title>',
            ]
            if c_en:
                lineas_cap.append(f'{indent}    <p xml:lang="en">{c_en}</p>')
            lineas_cap.append(f'{indent}  </caption>')
            return nl.join(lineas_cap)
        return (
            f'{indent}  <caption xml:lang="en">{nl}'
            f'{indent}    <title>{c_en}</title>{nl}'
            f'{indent}  </caption>'
        )

    def _render_fig_jats(num_f: int, nombre_img: str, cap_es_h: str, cap_en_h: str, indent: str) -> str:
        cap_xml = _construir_captions_jats(cap_es_h, cap_en_h, indent)
        cap_linea = f"{cap_xml}{nl}" if cap_xml else ""
        return (
            f'{indent}<fig id="f{num_f}">{nl}'
            f'{indent}  <label>{"Figure" if idioma_art == "en" else "Figura"} {num_f}</label>{nl}'
            f'{cap_linea}'
            f'{indent}  <graphic xlink:href="{_xml_escape_attr(nombre_img)}"/>{nl}'
            f'{indent}</fig>'
        )

    def _render_tabla_jats(blk: dict, indent: str) -> str:
        num_t = blk["num"]
        cap_xml = _construir_captions_jats(
            blk.get("cap_es_html") or blk.get("cap_es", ""),
            blk.get("cap_en_html") or blk.get("cap_en", ""),
            indent,
        )
        cap_linea = f"{cap_xml}{nl}" if cap_xml else ""
        tbl_xml = _html_tabla_a_jats(blk.get("tabla_html_raw", ""), indent=f"{indent}  ")
        notas_raw = str(blk.get("notas_raw", "")).strip()
        foot_xml = ""
        if notas_raw:
            foot_xml = (
                f'{nl}{indent}  <table-wrap-foot>{nl}'
                f'{indent}    <fn id="tfn{num_t}"><p>{_html_inline_a_jats(notas_raw)}</p></fn>{nl}'
                f'{indent}  </table-wrap-foot>'
            )
        return (
            f'{indent}<table-wrap id="t{num_t}">{nl}'
            f'{indent}  <label>{"Table" if idioma_art == "en" else ("Tabela" if idioma_art == "pt" else "Tabla")} {num_t}</label>{nl}'
            f'{cap_linea}'
            f'{tbl_xml}{foot_xml}{nl}'
            f'{indent}</table-wrap>'
        )

    if cuerpo_bloques:
        for blk in cuerpo_bloques:
            btipo = blk.get("tipo", "")
            ind = "      " if en_seccion else "    "
            if btipo == "h2":
                if en_seccion:
                    body_lines.append("    </sec>")
                sec_type = _inferir_sec_type_jats(blk.get("html", ""))
                st_attr = f' sec-type="{sec_type}"' if sec_type else ""
                body_lines.append(f"    <sec{st_attr}>")
                body_lines.append(f"      <title>{_limpiar_titulo_jats(_html_inline_a_jats(blk.get('html', '')))}</title>")
                en_seccion = True
            elif btipo in ("p", "fecha"):
                p_jats = _html_inline_a_jats(blk.get("html", ""))
                if p_jats:
                    body_lines.append(f"{ind}<p>{p_jats}</p>")
            elif btipo == "blockquote":
                q_jats = _html_inline_a_jats(blk.get("html", ""))
                if q_jats:
                    body_lines.append(f"{ind}<disp-quote>{nl}{ind}  <p>{q_jats}</p>{nl}{ind}</disp-quote>")
            elif btipo == "FIGURA":
                body_lines.append(
                    _render_fig_jats(
                        blk["num"],
                        blk.get("nombre_img", f"fig{blk['num']:02d}.png"),
                        blk.get("cap_es_html", ""),
                        blk.get("cap_en_html", ""),
                        indent=ind,
                    )
                )
            elif btipo == "TABLA":
                body_lines.append(_render_tabla_jats(blk, indent=ind))
        if en_seccion:
            body_lines.append("    </sec>")

    body_xml = nl.join(body_lines) if body_lines else "    <p/>"

    fn_group_block = ""
    if art.get("notas_usadas"):
        fns_xml = []
        for idx_fn, (_, fn_html) in enumerate(art["notas_usadas"], start=1):
            fns_xml.append(
                f'      <fn id="fn{idx_fn}">{nl}'
                f'        <label>{idx_fn}</label>{nl}'
                f'        <p>{_html_inline_a_jats(fn_html)}</p>{nl}'
                f'      </fn>'
            )
        fn_group_block = f'    <fn-group>{nl}{nl.join(fns_xml)}{nl}    </fn-group>{nl}'

    refs_xml = [
        _formatear_referencia_jats(idx, r_html)
        for idx, r_html in enumerate(art.get("referencias", []), start=1)
    ]

    return f"""<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE article PUBLIC "-//NLM//DTD JATS (Z39.96) Journal Publishing DTD v1.1 20151215//EN" "JATS-journalpublishing1.dtd">
<article article-type="{article_type}" dtd-version="1.1" specific-use="sps-1.9" xml:lang="{idioma_art}" xmlns:xlink="http://www.w3.org/1999/xlink">
  <front>
    <journal-meta>
      <journal-id journal-id-type="publisher-id">rma</journal-id>
      <journal-title-group>
        <journal-title>{_xml_escape_text(titulo_rev)}</journal-title>
        <abbrev-journal-title abbrev-type="publisher">Rev. Mus. Antropol.</abbrev-journal-title>
      </journal-title-group>
      <issn pub-type="ppub">{_xml_escape_text(str(rev.get('issn_impreso', '1852-060X')))}</issn>
      <issn pub-type="epub">{_xml_escape_text(str(rev.get('issn_electronico', '1852-4826')))}</issn>
      <publisher>
        <publisher-name>{_xml_escape_text(str(rev.get('subtitulo_institucional', 'IDACOR-CONICET / FFyH - UNC')))}</publisher-name>
      </publisher>
    </journal-meta>
    <article-meta>
      <article-id pub-id-type="publisher-id">{_xml_escape_text(str(meta.get('id_ojs', '')))}</article-id>{doi_tag}
      <article-categories>
        <subj-group subj-group-type="heading">
          <subject>{_xml_escape_text(seccion_str)}</subject>
        </subj-group>
      </article-categories>
      <title-group>
        <article-title>{titulo_en_jats if idioma_art == "en" else titulo_es_jats}</article-title>
        <trans-title-group xml:lang="{idioma_local_xml if idioma_art == "en" else "en"}">
          <trans-title>{titulo_es_jats if idioma_art == "en" else titulo_en_jats}</trans-title>
        </trans-title-group>
      </title-group>{contrib_group_block}{affs_block}
      <pub-date date-type="pub" publication-format="electronic">
        <day>{pub_dmy[0]}</day>
        <month>{pub_dmy[1]}</month>
        <year>{pub_dmy[2]}</year>
      </pub-date>
      <volume>{_xml_escape_text(str(meta.get('volumen', '')))}</volume>
      <issue>{_xml_escape_text(str(meta.get('numero', '')))}</issue>
      <fpage>{_xml_escape_text(fpage_num)}</fpage>{lpage_tag}{history_block}{permissions_block}
      <abstract>
{abstract_jats if idioma_art == "en" else resumen_jats}
      </abstract>
      <trans-abstract xml:lang="{idioma_local_xml if idioma_art == "en" else "en"}">
{resumen_jats if idioma_art == "en" else abstract_jats}
      </trans-abstract>
      <kwd-group xml:lang="{idioma_local_xml}">
        <title>{"Palavras-chave" if idioma_local_xml == "pt" else "Palabras clave"}</title>
{kw_es_xml}
      </kwd-group>
      <kwd-group xml:lang="en">
        <title>Keywords</title>
{kw_en_xml}
      </kwd-group>
    </article-meta>
  </front>
  <body>
{body_xml}
  </body>
  <back>
{fn_group_block}    <ref-list>
      <title>{"References" if idioma_art == "en" else ("Referências bibliográficas" if idioma_art == "pt" else "Referencias bibliográficas")}</title>
{nl.join(refs_xml)}
    </ref-list>
  </back>
</article>
"""


def compilar_articulo(
    carpeta: Path,
    art_id: str,
    pagina_inicio_override: int | None = None,
    meta_overrides: dict | None = None,
    actualizar_yaml: bool = False,
    formato_override: str | None = None,
) -> dict:
    """Compila un artículo en el perfil elegido ('tradicional' 2 cols o 'moderno' 1 col)."""
    dir_script = Path(__file__).resolve().parent

    if (carpeta / "revista.yaml").exists() or (carpeta / "revista.json").exists():
        rev = cargar_config(carpeta / "revista", obligatorio=False)
    else:
        rev = cargar_config(dir_script / "revista", obligatorio=False)

    ruta_manuscrito = carpeta / f"{art_id}_manuscrito.docx"
    if not ruta_manuscrito.exists():
        raise FileNotFoundError(f"No se encontró el manuscrito: {ruta_manuscrito}")

    ruta_meta_yaml = carpeta / f"{art_id}_metadatos.yaml"
    meta = cargar_config(carpeta / f"{art_id}_metadatos", obligatorio=True)
    if meta_overrides:
        for k, v in meta_overrides.items():
            if v is not None and str(v).strip() != "":
                meta[k] = v

    perfil = resolver_perfil_formato(formato_override, meta, rev)
    css_origen = dir_script / ("rma_print_tradicional.css" if perfil == "tradicional" else "rma_print_moderno.css")

    rec = cargar_config(carpeta / f"{art_id}_recursos", obligatorio=False)
    art = parsear_manuscrito(ruta_manuscrito)

    tablas_html = []
    ruta_tablas = carpeta / f"{art_id}_tablas.docx"
    if ruta_tablas.exists():
        doc_t = Document(ruta_tablas)
        tablas_html = [extraer_tabla_docx(t, {}, []) for t in doc_t.tables]
    elif art["tablas_embebidas"]:
        tablas_html = art["tablas_embebidas"]

    figuras_cfg = dict(rec.get("figuras") or {})
    tablas_cfg = dict(rec.get("tablas") or {})

    ruta_epigrafes = buscar_archivo_epigrafes(carpeta, art_id)
    if ruta_epigrafes is not None:
        epig_doc = parsear_epigrafes_docx(ruta_epigrafes)
        for num_f, datos_f in epig_doc.get("figuras", {}).items():
            actual_f = dict(figuras_cfg.get(num_f) or figuras_cfg.get(str(num_f)) or {})
            es_act = str(actual_f.get("es", "")).strip()
            en_act = str(actual_f.get("en", "")).strip()
            if (not es_act or es_act == f"Descripción de la Figura {num_f}.") and datos_f.get("es"):
                actual_f["es"] = datos_f["es"]
            if (not en_act or en_act == f"Description of Figure {num_f}.") and datos_f.get("en"):
                actual_f["en"] = datos_f["en"]
            figuras_cfg[num_f] = actual_f

        for num_t, datos_t in epig_doc.get("tablas", {}).items():
            actual_t = dict(tablas_cfg.get(num_t) or tablas_cfg.get(str(num_t)) or {})
            es_act = str(actual_t.get("es", "")).strip()
            en_act = str(actual_t.get("en", "")).strip()
            if (not es_act or es_act == f"Descripción de la Tabla {num_t}.") and datos_t.get("es"):
                actual_t["es"] = datos_t["es"]
            if (not en_act or en_act == f"Description of Table {num_t}.") and datos_t.get("en"):
                actual_t["en"] = datos_t["en"]
            if not str(actual_t.get("notas", "")).strip() and datos_t.get("notas"):
                actual_t["notas"] = datos_t["notas"]
            tablas_cfg[num_t] = actual_t

    carpeta_salida = carpeta / "salida"
    carpeta_salida_fig = carpeta_salida / "figuras"
    carpeta_salida_assets = carpeta_salida / "assets"
    carpeta_salida_fig.mkdir(parents=True, exist_ok=True)
    carpeta_salida_assets.mkdir(parents=True, exist_ok=True)

    logo_rel_cfg = str(meta.get("logo") or rev.get("logo") or "assets/logo_rma.png")
    ruta_logo_origen = (carpeta / logo_rel_cfg) if (carpeta / logo_rel_cfg).exists() else (dir_script / logo_rel_cfg)
    logo_html_tag = ""
    if ruta_logo_origen.exists():
        shutil.copy2(ruta_logo_origen, carpeta_salida_assets / ruta_logo_origen.name)
        logo_html_tag = f'<img class="logo-rma-img" src="assets/{html.escape(ruta_logo_origen.name)}" alt="Logo RMA">'

    logo_cc_cfg = str(meta.get("logo_cc") or rev.get("logo_cc") or "assets/cc_by_nc_sa.png")
    ruta_cc_origen = (carpeta / logo_cc_cfg) if (carpeta / logo_cc_cfg).exists() else (dir_script / logo_cc_cfg)
    url_licencia = str(rev.get("url_licencia", "https://creativecommons.org/licenses/by-nc-sa/4.0/")).strip()
    licencia = str(rev.get("licencia", "CC BY-NC-SA 4.0")).strip()
    if ruta_cc_origen.exists():
        shutil.copy2(ruta_cc_origen, carpeta_salida_assets / ruta_cc_origen.name)
        cc_html_tag = (
            f'<a href="{html.escape(url_licencia)}" title="{html.escape(licencia)}">'
            f'<img class="logo-cc-img" src="assets/{html.escape(ruta_cc_origen.name)}" alt="{html.escape(licencia)}">'
            f'</a>'
        )
    else:
        cc_html_tag = f'<a href="{html.escape(url_licencia)}">{html.escape(licencia)}</a>'

    tiene_icono_orcid = False
    ruta_orcid_svg = dir_script / "assets" / "orcid.svg"
    if ruta_orcid_svg.exists():
        shutil.copy2(ruta_orcid_svg, carpeta_salida_assets / "orcid.svg")
        tiene_icono_orcid = True

    tiene_icono_ror = False
    ruta_ror_svg = dir_script / "assets" / "ror.svg"
    if ruta_ror_svg.exists():
        shutil.copy2(ruta_ror_svg, carpeta_salida_assets / "ror.svg")
        tiene_icono_ror = True

    cuerpo_bloques = []
    figuras_usadas = []

    idioma_art = normalizar_codigo_idioma(meta.get("idioma", "")) or detectar_idioma_texto(" ".join(re.sub(r"<[^>]+>", " ", x.get("html", "")) for x in art["bloques_cuerpo"][:8]) or f"{art.get('titulo_es', '')} {art.get('titulo_en', '')}")
    idioma_local_art = "pt" if (idioma_art == "pt" or art.get("idioma_resumen_local") == "pt") else "es"
    es_ingles = (idioma_art == "en")
    etiq_figura = "Figure" if es_ingles else "Figura"
    etiq_tabla = "Table" if es_ingles else ("Tabela" if idioma_art == "pt" else "Tabla")
    etiq_refs = "References" if es_ingles else ("Referências bibliográficas" if idioma_art == "pt" else "Referencias bibliográficas")
    etiq_notas = "Notes" if es_ingles else "Notas"

    motor_citas = MotorCitacionesRMA(art["referencias"], estilo_id="semantico")
    for idx_b, b in enumerate(art["bloques_cuerpo"]):
        blk_id = f"blk-{idx_b}"
        if b["tipo"] == "h2":
            cuerpo_bloques.append({
                "id": blk_id,
                "tipo": "h2",
                "html": b["html"],
                "movible": True,
                "html_ojs": f"<h2>{b['html']}</h2>",
                "html_pdf": f'<h2 data-blk="{blk_id}">{b["html"]}</h2>',
            })
        elif b["tipo"] == "blockquote":
            cuerpo_bloques.append({
                "id": blk_id,
                "tipo": "blockquote",
                "html": b["html"],
                "movible": True,
                "html_ojs": f"<blockquote>{motor_citas.enlazar_fragmento(b['html'])}</blockquote>",
                "html_pdf": f'<blockquote data-blk="{blk_id}">{motor_citas.enlazar_fragmento(b["html"])}</blockquote>',
            })
        elif b["tipo"] == "fecha":
            cuerpo_bloques.append({
                "id": blk_id,
                "tipo": "fecha",
                "html": b["html"],
                "movible": False,
                "html_ojs": f'<p class="fecha-manuscrito">{b["html"]}</p>',
                "html_pdf": f'<p class="fecha-manuscrito" data-blk="{blk_id}">{b["html"]}</p>',
            })
        elif b["tipo"] == "p":
            cuerpo_bloques.append({
                "id": blk_id,
                "tipo": "p",
                "html": b["html"],
                "movible": True,
                "html_ojs": f"<p>{motor_citas.enlazar_fragmento(b['html'])}</p>",
                "html_pdf": f'<p data-blk="{blk_id}">{motor_citas.enlazar_fragmento(b["html"])}</p>',
            })
        elif b["tipo"] == "FIGURA":
            num = b["num"]
            info = figuras_cfg.get(num) or figuras_cfg.get(str(num)) or {}
            nombre_img = obtener_imagen_figura(carpeta, carpeta_salida_fig, art_id, num, info.get("archivo", ""))
            clase_ancho = clase_ancho_recurso(info, perfil, ruta_img=carpeta_salida_fig / nombre_img)
            alto_max_cfg = str(info.get("alto_max") or info.get("max_height") or "").strip()
            ruta_rel_html = f"figuras/{nombre_img}"
            flujo_fijo = str(info.get("flujo", "")).strip().lower() in ("fijo", "fixed", "anclado", "estatico", "estático")

            cap_es_raw = info.get("pt") or info.get("es") or (f"Descrição da Figura {num}." if idioma_art == "pt" else f"Descripción de la Figura {num}.")
            cap_en_raw = info.get("en", f"Description of Figure {num}.")
            cap_es = enlazar_urls_y_correos(cap_es_raw)
            cap_en = enlazar_urls_y_correos(cap_en_raw)
            figuras_usadas.append((num, nombre_img, cap_es_raw, cap_en_raw))

            cuerpo_bloques.append({
                "id": blk_id,
                "tipo": "FIGURA",
                "num": num,
                "nombre_img": nombre_img,
                "cap_es_html": cap_es_raw,
                "cap_en_html": cap_en_raw,
                "clase_ancho": clase_ancho,
                "movible": not flujo_fijo,
                "html_ojs": (
                    f'<figure class="rma-figura {clase_ancho}" id="fig-{num}">\n'
                    f'  <img src="{html.escape(ruta_rel_html)}" alt="{etiq_figura} {num}">\n'
                    f'  <figcaption>\n'
                    f'{_html_caption_bilingue("Figura", "Figure", num, cap_es, cap_en, idioma_art)}\n'
                    f'  </figcaption>\n'
                    f'</figure>'
                ),
                "html_pdf": (
                    f'<figure class="rma-figura {clase_ancho}" id="fig-{num}" data-blk="{blk_id}">\n'
                    f'  <img src="{html.escape(ruta_rel_html)}" alt="{etiq_figura} {num}">\n'
                    f'  <figcaption>\n'
                    f'{_html_caption_bilingue("Figura", "Figure", num, cap_es, cap_en, idioma_art)}\n'
                    f'  </figcaption>\n'
                    f'</figure>'
                ),
            })
        elif b["tipo"] == "TABLA":
            num = b["num"]
            info = tablas_cfg.get(num) or tablas_cfg.get(str(num)) or {}
            tabla_raw = resolver_html_tabla(
                carpeta, art_id, num, info, tablas_html, art["tablas_embebidas"]
            )
            tabla_code, num_cols_t, clase_densidad, num_filas_t, largo_txt_t = preparar_tabla_responsive(tabla_raw)
            clase_ancho = clase_ancho_recurso(info, perfil, num_cols_tabla=num_cols_t)
            flujo_fijo = str(info.get("flujo", "")).strip().lower() in ("fijo", "fixed", "anclado", "estatico", "estático")

            dividir_raw = info.get("dividir", "auto")
            if isinstance(dividir_raw, bool):
                modo_dividir = "divisible" if dividir_raw else "indivisible"
                dividir_forzado = True
            else:
                dv_str = str(dividir_raw).strip().lower()
                if dv_str in ("false", "0", "no", "off", "indivisible", "entera", "fijo"):
                    modo_dividir = "indivisible"
                    dividir_forzado = True
                elif dv_str in ("true", "1", "si", "sí", "on", "divisible", "partir"):
                    modo_dividir = "divisible"
                    dividir_forzado = True
                else:
                    dividir_forzado = False
                    modo_dividir = (
                        "indivisible"
                        if (num_filas_t <= 22 and largo_txt_t <= 2800)
                        else "divisible"
                    )

            cap_es = enlazar_urls_y_correos(info.get("pt") or info.get("es") or (f"Descrição da Tabela {num}." if idioma_art == "pt" else f"Descripción de la Tabla {num}."))
            cap_en = enlazar_urls_y_correos(info.get("en", f"Description of Table {num}."))
            notas_t = (
                f'<div class="tabla-notas">{enlazar_urls_y_correos(info["notas"])}</div>'
                if info.get("notas")
                else ""
            )

            tabla_code_pdf = _etiquetar_filas_tabla_pdf(tabla_code, blk_id)
            clase_div_css = f"tabla-{modo_dividir}"

            cuerpo_bloques.append({
                "id": blk_id,
                "tipo": "TABLA",
                "num": num,
                "clase_ancho": clase_ancho,
                "clase_densidad": clase_densidad,
                "modo_dividir": modo_dividir,
                "dividir_forzado": dividir_forzado,
                "tabla_html_raw": tabla_code,
                "tabla_code_pdf": tabla_code_pdf,
                "notas_raw": str(info.get("notas", "") or "").strip(),
                "notas_html": notas_t,
                "cap_es_html": str(info.get("pt") or info.get("es") or (f"Descrição da Tabela {num}." if idioma_art == "pt" else f"Descripción de la Tabla {num}.")),
                "cap_en_html": str(info.get("en", f"Description of Table {num}.")),
                "cap_es": cap_es,
                "cap_en": cap_en,
                "movible": not flujo_fijo,
                "html_ojs": (
                    f'<div class="rma-tabla {clase_ancho} {clase_densidad}" id="tab-{num}">\n'
                    f'  {tabla_code}\n'
                    f'  {notas_t}\n'
                    f'  <div class="tabla-caption">\n'
                    f'{_html_caption_bilingue("Tabela" if idioma_art == "pt" else "Tabla", "Table", num, cap_es, cap_en, idioma_art)}\n'
                    f'  </div>\n'
                    f'</div>'
                ),
                "html_pdf": (
                    f'<div class="rma-tabla {clase_ancho} {clase_densidad} {clase_div_css}" id="tab-{num}" data-blk="{blk_id}">\n'
                    f'  {tabla_code_pdf}\n'
                    f'  {notas_t}\n'
                    f'  <div class="tabla-caption">\n'
                    f'{_html_caption_bilingue("Tabela" if idioma_art == "pt" else "Tabla", "Table", num, cap_es, cap_en, idioma_art)}\n'
                    f'  </div>\n'
                    f'</div>'
                ),
            })

    cuerpo_render_ojs = [x["html_ojs"] for x in cuerpo_bloques]

    # Resolver autores y afiliaciones estructuradas (combinando manuscrito .docx + metadatos.yaml)
    detectado_aut_af = parsear_autores_y_afiliaciones_docx(ruta_manuscrito)
    autores_est = list(detectado_aut_af.get("autores", []))
    afiliaciones_est = list(detectado_aut_af.get("afiliaciones", []))

    meta_autores = meta.get("autores")
    if isinstance(meta_autores, list) and meta_autores:
        autores_curados = []
        for idx_a, ma in enumerate(meta_autores):
            if not isinstance(ma, dict):
                continue
            base_det = autores_est[idx_a] if idx_a < len(autores_est) else {}
            nom_m = str(ma.get("nombre") or base_det.get("nombre") or "").strip()
            if not nom_m:
                continue
            af_m = str(ma.get("afiliacion") if ma.get("afiliacion") not in (None, "") else base_det.get("afiliacion", "")).strip()
            oid_m = normalizar_orcid(ma.get("orcid") or base_det.get("orcid") or "")[0]
            autores_curados.append({
                "nombre": nom_m,
                "afiliacion": af_m,
                "orcid": oid_m,
            })
        if autores_curados:
            autores_est = autores_curados

    meta_afils = meta.get("afiliaciones")
    ror_por_clave = {}
    if isinstance(meta_afils, dict):
        for k_af, v_af in meta_afils.items():
            clave_s = str(k_af).strip()
            if isinstance(v_af, dict):
                rids_str = ", ".join(r[0] for r in normalizar_rors(v_af.get("ror", "")))
                if rids_str:
                    ror_por_clave[clave_s] = rids_str
                for af_item in afiliaciones_est:
                    if str(af_item.get("clave", "")).strip() == clave_s:
                        if rids_str:
                            af_item["ror"] = rids_str
                        if v_af.get("texto"):
                            af_item["texto"] = str(v_af["texto"]).strip()
            elif isinstance(v_af, str):
                rids_str = ", ".join(r[0] for r in normalizar_rors(v_af))
                if rids_str:
                    ror_por_clave[clave_s] = rids_str
    elif isinstance(meta_afils, list):
        for idx_af, v_af in enumerate(meta_afils, start=1):
            if isinstance(v_af, dict):
                clave_s = str(v_af.get("clave") or idx_af).strip()
                rid, _ = normalizar_ror(v_af.get("ror", ""))
                if rid:
                    ror_por_clave[clave_s] = rid
                if idx_af - 1 < len(afiliaciones_est):
                    if rid:
                        afiliaciones_est[idx_af - 1]["ror"] = rid
                    if v_af.get("texto"):
                        afiliaciones_est[idx_af - 1]["texto"] = str(v_af["texto"]).strip()

    # Limpiar URLs crudas de ORCID/ROR en los párrafos de afiliación y añadir ícono ROR hipervinculado
    lineas_afil_render = []
    for idx_af, a_html in enumerate(art["afiliaciones"], start=1):
        a_limpia, orcids_linea, rors_linea = _limpiar_urls_orcid_ror_en_html(a_html)
        clave_af = (
            str(afiliaciones_est[idx_af - 1].get("clave", idx_af)).strip()
            if (idx_af - 1) < len(afiliaciones_est)
            else str(idx_af)
        )
        # Si había ORCID en la afiliación del Word y no estaba asignado aún al autor, asignarlo
        if orcids_linea:
            for oid in orcids_linea:
                asignado = False
                for aut in autores_est:
                    refs = [x.strip() for x in str(aut.get("afiliacion", "")).split(",")]
                    if clave_af in refs and not aut.get("orcid"):
                        aut["orcid"] = oid
                        asignado = True
                        break
                if not asignado and (idx_af - 1) < len(autores_est) and not autores_est[idx_af - 1].get("orcid"):
                    autores_est[idx_af - 1]["orcid"] = oid

        ror_raw = ror_por_clave.get(clave_af, "")
        if not ror_raw and (idx_af - 1) < len(afiliaciones_est):
            ror_raw = str(afiliaciones_est[idx_af - 1].get("ror", "")).strip()
        rors_list = normalizar_rors(ror_raw)
        if not rors_list and rors_linea:
            rors_list = [normalizar_ror(r) for r in rors_linea if normalizar_ror(r)[0]]
            if rors_list and (idx_af - 1) < len(afiliaciones_est):
                afiliaciones_est[idx_af - 1]["ror"] = ", ".join(r[0] for r in rors_list)

        for rid, ror_url in rors_list:
            if tiene_icono_ror:
                badge_ror = (
                    f' <a class="enlace-id enlace-ror" href="{html.escape(ror_url)}" '
                    f'title="ROR: {html.escape(rid)}">'
                    f'<img class="icono-ror" src="assets/ror.svg" alt="ROR: {html.escape(rid)}"></a>'
                )
            else:
                badge_ror = f' <a class="enlace-id enlace-ror" href="{html.escape(ror_url)}">[ROR]</a>'
            a_limpia = f"{a_limpia}{badge_ror}"

        lineas_afil_render.append(f'      <p class="afil">{a_limpia}</p>')

    afiliaciones_html = "\n".join(lineas_afil_render)

    # Renderizar línea de autores con ícono ORCID hipervinculado cuando exista
    if autores_est:
        partes_autores_html = []
        for idx_a, aut in enumerate(autores_est):
            nom = html.escape(str(aut.get("nombre", "")).strip())
            if not nom:
                continue
            af_ref = str(aut.get("afiliacion", "")).strip()
            sup_html = f"<sup>{html.escape(af_ref)}</sup>" if af_ref else ""
            oid, orcid_url = normalizar_orcid(aut.get("orcid", ""))
            orcid_badge = ""
            if oid:
                if tiene_icono_orcid:
                    orcid_badge = (
                        f'<a class="enlace-id enlace-orcid" href="{html.escape(orcid_url)}" '
                        f'title="ORCID: {html.escape(oid)}">'
                        f'<img class="icono-orcid" src="assets/orcid.svg" alt="ORCID: {html.escape(oid)}"></a>'
                    )
                else:
                    orcid_badge = f' <a class="enlace-id enlace-orcid" href="{html.escape(orcid_url)}">[ORCID]</a>'
            partes_autores_html.append(f"<strong>{nom}</strong>{sup_html}{orcid_badge}")
        autores_html_render = ", ".join(partes_autores_html) if partes_autores_html else art["autores"]
    else:
        autores_html_render = art["autores"]

    meta_autores_head = []
    for aut in autores_est:
        nom = str(aut.get("nombre", "")).strip()
        if nom:
            meta_autores_head.append(f'    <meta name="citation_author" content="{html.escape(nom)}">')
            _, orcid_url = normalizar_orcid(aut.get("orcid", ""))
            if orcid_url:
                meta_autores_head.append(f'    <meta name="citation_author_orcid" content="{html.escape(orcid_url)}">')
    meta_autores_head_str = ("\n" + "\n".join(meta_autores_head)) if meta_autores_head else ""
    refs_con_backlinks = motor_citas.renderizar_referencias_con_backlinks(art["referencias"])
    refs_html = "\n".join(f'      <p class="ref-item">{r}</p>' for r in refs_con_backlinks)

    notas_html_section = ""
    if art["notas_usadas"]:
        items_fn = "\n".join(
            f'      <li id="fn-{idx}">{txt} <a href="#fnref-{idx}" title="Volver al texto">↩</a></li>'
            for idx, (_, txt) in enumerate(art["notas_usadas"], start=1)
        )
        notas_html_section = (
            f'    <section class="notas-al-pie-html">\n'
            f'      <h2>{etiq_notas}</h2>\n'
            f'      <ol>\n{items_fn}\n      </ol>\n'
            f'    </section>\n'
        )

    titulo_plano = re.sub(r"<[^<]+?>", "", art["titulo_es"]).strip()
    autores_planos = re.sub(r"<[^<]+?>", "", art["autores"]).strip()

    autor_corto_val = str(meta.get("autor_corto", "")).strip()
    if not autor_corto_val and autores_planos:
        lista_aut = [a.strip() for a in re.split(r"(?!(?<=\d),\s*\d),|;|\by\b", autores_planos) if a.strip()]
        if lista_aut:
            partes_p1 = lista_aut[0].split()
            inicial = f"{partes_p1[0][0]}. " if len(partes_p1) > 1 else ""
            apellido = partes_p1[-1]
            autor_corto_val = f"{inicial}{apellido}" + (" et al." if len(lista_aut) > 1 else "")

    titulo_corto_val = str(meta.get("titulo_corto", "")).strip()
    if not titulo_corto_val and titulo_plano:
        titulo_corto_val = sugerir_titulo_corto(titulo_plano, 55)

    sigla = str(rev.get("sigla", "RMA")).strip()
    titulo_revista = str(rev.get("titulo_revista", "Revista del Museo de Antropología")).strip()
    subtitulo_inst = str(
        rev.get(
            "subtitulo_institucional",
            "IDACOR-CONICET / Facultad de Filosofía y Humanidades – Universidad Nacional de Córdoba - Argentina",
        )
    ).strip()
    url_revista = str(rev.get("url_revista", "http://revistas.unc.edu.ar/index.php/antropologia/index")).strip()
    issn_imp = str(rev.get("issn_impreso", "1852-060X")).strip()
    issn_ele = str(rev.get("issn_electronico", "1852-4826")).strip()
    digitos_pag = int(rev.get("digitos_pagina", 2))

    doi_raw = str(meta.get("doi", "")).strip() or "https://doi.org/10.31048/xxxxxx"
    doi_href = normalizar_url_doi(doi_raw)
    seccion = str(meta.get("seccion", "Artículos")).strip()
    idioma_art = normalizar_codigo_idioma(meta.get("idioma", "")) or detectar_idioma_texto(" ".join(re.sub(r"<[^>]+>", " ", b.get("html", "")) for b in cuerpo_bloques[:8]) or f"{titulo_plano} {art.get('resumen', '')} {art.get('abstract', '')}")
    idioma_local_art = "pt" if (idioma_art == "pt" or art.get("idioma_resumen_local") == "pt") else "es"
    es_ingles = (idioma_art == "en")
    etiq_figura = "Figure" if es_ingles else "Figura"
    etiq_tabla = "Table" if es_ingles else ("Tabela" if idioma_art == "pt" else "Tabla")
    etiq_refs = "References" if es_ingles else ("Referências bibliográficas" if idioma_art == "pt" else "Referencias bibliográficas")
    etiq_notas = "Notes" if es_ingles else "Notas"
    vol = str(meta.get("volumen", "")).strip()
    num_rev = str(meta.get("numero", "")).strip()
    anio = str(meta.get("anio", "")).strip()

    editor_sec = str(meta.get("editor_seccion") or meta.get("editor") or "").strip()
    editores_inv = str(meta.get("editores_invitados") or "").strip()
    lineas_editores = []
    if editor_sec:
        lineas_editores.append(f'<p class="editor-principal">Editado por {html.escape(editor_sec)}</p>')
    if editores_inv:
        lineas_editores.append(f'<p class="editores-invitados">Editores invitados: {html.escape(editores_inv)}</p>')
    editores_html = "\n                ".join(lineas_editores)

    issn_running = f"ISSN {issn_imp} (impreso) / ISSN {issn_ele} (electrónico)"
    cod_pub = str(meta.get("codigo_publicacion", "")).strip()
    nombre_salida = cod_pub if cod_pub else str(art_id)
    meta["id_ojs"] = art_id
    meta["nombre_salida"] = nombre_salida

    pagina_inicio = determinar_pagina_inicio(meta, pagina_inicio_override)

    def construir_documento_html(
        hoja_css: str,
        pags_rango: str,
        fpage_str: str,
        lpage_str: str,
        estilo_inline_extra: str = "",
        bloques_html_cuerpo: list[str] | None = None,
        incluir_refs_externas: bool = True,
    ) -> str:
        lista_cuerpo = bloques_html_cuerpo if bloques_html_cuerpo is not None else cuerpo_render_ojs
        cita_revista = f"{titulo_revista} {vol} ({num_rev}): {pags_rango} | {anio}"
        edicion_cabezal = f"Vol. {vol}, N.º {num_rev} ({anio}) · pp. {pags_rango}"
        cita_pie_p1 = (
            f"{titulo_revista} {vol} ({num_rev}): {pags_rango} / {anio} / "
            f"ISSN {issn_imp} (impreso) / ISSN {issn_ele} (electrónico)"
        )
        style_tag = f"\n    <style>{estilo_inline_extra}</style>" if estilo_inline_extra else ""
        incluir_bloque_refs = incluir_refs_externas and (bloques_html_cuerpo is None)
        bloque_refs_final = (
            f'\n        <h2>{etiq_refs}</h2>\n'
            f'        <section class="referencias">\n{refs_html}\n        </section>'
            if incluir_bloque_refs
            else ""
        )
        return f"""<!DOCTYPE html>
<html lang="{idioma_art}">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{html.escape(titulo_plano)} | {html.escape(titulo_revista)}</title>
    <meta name="citation_journal_title" content="{html.escape(titulo_revista)}">
    <meta name="citation_issn" content="{html.escape(issn_ele)}">
    <meta name="citation_volume" content="{html.escape(vol)}">
    <meta name="citation_issue" content="{html.escape(num_rev)}">
    <meta name="citation_firstpage" content="{html.escape(fpage_str)}">
    <meta name="citation_lastpage" content="{html.escape(lpage_str)}">
    <meta name="citation_title" content="{html.escape(titulo_plano)}">
    <meta name="citation_doi" content="{html.escape(doi_href)}">{meta_autores_head_str}
    <link rel="stylesheet" href="{hoja_css}">{style_tag}
</head>
<body class="perfil-{perfil}">
    <div class="running-header-left">
        <span class="rh-linea">{html.escape(autor_corto_val)} | {html.escape(cita_revista)}</span>
        <span class="rh-linea">DOI: <a href="{html.escape(doi_href)}">{html.escape(doi_href)}</a></span>
    </div>
    <div class="running-header-right">
        <span class="rh-linea">{html.escape(titulo_corto_val)} | {html.escape(cita_revista)}</span>
        <span class="rh-linea">DOI: <a href="{html.escape(doi_href)}">{html.escape(doi_href)}</a></span>
    </div>
    <span class="meta-running-issn">{html.escape(issn_running)}</span>

    <footer class="pie-primera-pagina">
        <div class="pie-p1-grid">
            <div class="pie-p1-cc">
                {cc_html_tag}
            </div>
            <div class="pie-p1-centro">
                <p>Recibido {html.escape(str(meta.get('recibido', '')))}. Recibido con correcciones {html.escape(str(meta.get('correcciones', '')))}. Aceptado {html.escape(str(meta.get('aceptado', '')))}</p>
                <p class="cita-biblio-p1">{html.escape(cita_pie_p1)}</p>
                <p><a href="{html.escape(url_revista)}">{html.escape(url_revista)}</a></p>
                <p>{html.escape(subtitulo_inst)}</p>
            </div>
            <div class="pie-p1-editores">
                {editores_html}
            </div>
        </div>
    </footer>

    <header class="portada-articulo">
        <div class="cabezal-revista">
            <div class="marca-rma-vertical">
                {logo_html_tag}
                <span class="sigla-rma">{html.escape(sigla)}</span>
                <span class="seccion-rma">{html.escape(seccion)}</span>
            </div>
            <div class="datos-cabezal-der">
                <span class="nombre-revista-cabezal">{html.escape(titulo_revista)}</span>
                <span class="edicion-cabezal">{html.escape(edicion_cabezal)}</span>
                <span class="issn-cabezal">{html.escape(issn_running)}</span>
                <span class="doi-rma">DOI: <a href="{html.escape(doi_href)}">{html.escape(doi_raw)}</a></span>
            </div>
        </div>

        <div class="bloque-titulos">
            <h1 class="titulo-es">{art['titulo_es']}</h1>
            <div class="acento-titulo"></div>
            <p class="titulo-en">{art['titulo_en']}</p>
        </div>

        <div class="bloque-autoria">
            <div class="autores">{autores_html_render}</div>
            <div class="afiliaciones">
{afiliaciones_html}
            </div>
        </div>

        <div class="bloque-resumenes">
            <h3>{'Abstract' if es_ingles else ('Resumo' if idioma_local_art == 'pt' else 'Resumen')}</h3>
            {art['abstract'] if es_ingles else art['resumen']}
            <p class="keywords">{art['keywords'] if es_ingles else art['palabras_clave']}</p>
            <hr class="divisor-resumen">
            <h3>{('Resumo' if idioma_local_art == 'pt' else 'Resumen') if es_ingles else 'Abstract'}</h3>
            {art['resumen'] if es_ingles else art['abstract']}
            <p class="keywords">{art['palabras_clave'] if es_ingles else art['keywords']}</p>
        </div>
    </header>

    <hr class="divisor-portada">

    {"".join(_segmentar_cuerpo_multicolumna(lista_cuerpo, f"{notas_html_section}{bloque_refs_final}", perfil))}
</body>
</html>"""

    # Bloques para PDF incluyendo Referencias bibliográficas al final como bloques fluibles
    bloques_pdf_base = [dict(x) for x in cuerpo_bloques]
    if art["referencias"]:
        bloques_pdf_base.append({
            "id": "blk-refs-h2",
            "tipo": "h2_refs",
            "movible": True,
            "html_pdf": f'<h2 data-blk="blk-refs-h2">{etiq_refs}</h2>',
        })
        for idx_r, r_html in enumerate(art["referencias"]):
            rid = f"blk-ref-{idx_r}"
            bloques_pdf_base.append({
                "id": rid,
                "tipo": "ref",
                "movible": True,
                "html_pdf": f'<div class="referencias" data-blk="{rid}"><p class="ref-item">{r_html}</p></div>',
            })

    def _emparejar_figuras_columna(lista_blks: list[dict]) -> list[dict]:
        """En perfil tradicional (2 cols), agrupa pares consecutivos de figuras de 1 columna."""
        if perfil != "tradicional":
            return [dict(x) for x in lista_blks]
        salida = []
        i = 0
        while i < len(lista_blks):
            act = lista_blks[i]
            if (
                i + 1 < len(lista_blks)
                and act["tipo"] == "FIGURA"
                and act.get("clase_ancho") == "ancho-columna"
                and act.get("movible", True)
                and lista_blks[i + 1]["tipo"] == "FIGURA"
                and lista_blks[i + 1].get("clase_ancho") == "ancho-columna"
                and lista_blks[i + 1].get("movible", True)
            ):
                sig = lista_blks[i + 1]
                par_id = f"{act['id']}_{sig['id']}"
                salida.append({
                    "id": par_id,
                    "tipo": "FIGURA",
                    "clase_ancho": "ancho-completo",
                    "movible": True,
                    "es_par": True,
                    "html_pdf": (
                        f'<div class="par-figuras-columnas" data-blk="{par_id}">\n'
                        f'{act["html_pdf"]}\n{sig["html_pdf"]}\n'
                        f'</div>'
                    ),
                })
                i += 2
            else:
                salida.append(dict(act))
                i += 1
        return salida

    def _renderizar_y_medir_pdf(lista_blks: list[dict], regla_ini: str):
        html_str = construir_documento_html(
            ".rma_print_tmp.css",
            pags_rango=f"{pagina_inicio:0{digitos_pag}d}-00",
            fpage_str=f"{pagina_inicio:0{digitos_pag}d}",
            lpage_str="00",
            estilo_inline_extra=regla_ini,
            bloques_html_cuerpo=[x["html_pdf"] for x in lista_blks],
            incluir_refs_externas=False,
        )
        doc_r = HTML(string=html_str, base_url=str(carpeta_salida)).render()
        px_por_mm = 96.0 / 25.4
        paginas_info = []
        blk_pages: dict[str, list[int]] = {}
        blk_boxes: dict[str, list[tuple[int, float, float, float, float]]] = {}
        trow_pages: dict[str, dict[int, int]] = {}

        for p_idx, page in enumerate(doc_r.pages):
            pbox = page._page_box
            cbox = pbox.children[0] if pbox.children else pbox
            c_top = cbox.position_y
            c_bot = cbox.position_y + cbox.height
            c_left = cbox.position_x
            c_right = cbox.position_x + cbox.width
            mid_x = (c_left + c_right) / 2.0

            items = []
            def _walk(box):
                el = getattr(box, "element", None)
                if el is not None:
                    tr_attr = el.get("data-trow")
                    if tr_attr and ":" in tr_attr:
                        tid, ridx_s = tr_attr.split(":", 1)
                        try:
                            trow_pages.setdefault(tid, {})[int(ridx_s)] = p_idx
                        except ValueError:
                            pass
                    bid = el.get("data-blk")
                    if bid:
                        items.append((bid, box.position_x, box.position_y, box.width, box.height))
                        # Continuar recorriendo hijos si es una tabla para capturar data-trow
                        if not bid.startswith("blk-"):
                            return
                for ch in getattr(box, "children", []):
                    _walk(ch)
            _walk(pbox)

            max_y_l = c_top
            max_y_r = c_top
            for bid, x, y, w, h in items:
                blk_pages.setdefault(bid, []).append(p_idx)
                blk_boxes.setdefault(bid, []).append((p_idx, x, y, w, h))
                bot = y + h
                if perfil == "moderno" or w > (c_right - c_left) * 0.62:
                    if bot > max_y_l:
                        max_y_l = bot
                    if bot > max_y_r:
                        max_y_r = bot
                elif x + w / 2.0 < mid_x:
                    if bot > max_y_l:
                        max_y_l = bot
                else:
                    if bot > max_y_r:
                        max_y_r = bot

            gap_l_mm = max(0.0, (c_bot - max_y_l) / px_por_mm)
            gap_r_mm = max(0.0, (c_bot - max_y_r) / px_por_mm)
            paginas_info.append({
                "page_idx": p_idx,
                "c_top": c_top,
                "c_bot": c_bot,
                "gap_l_mm": gap_l_mm,
                "gap_r_mm": gap_r_mm,
                "gap_avg_mm": (gap_l_mm + gap_r_mm) / 2.0,
                "blks": list(dict.fromkeys(it[0] for it in items)),
            })

        return doc_r, paginas_info, blk_pages, blk_boxes, trow_pages

    def _obtener_paquete_texto(lista_blks: list[dict], j: int) -> list[int] | None:
        """Devuelve los índices consecutivos de prosa que pueden adelantarse sin romper títulos ni citas."""
        if j >= len(lista_blks):
            return None
        sig = lista_blks[j]
        if sig["tipo"] in ("FIGURA", "TABLA", "fecha") or not sig.get("movible", True):
            return None
        if sig["tipo"] in ("h2", "h2_refs"):
            paq = [j]
            k = j + 1
            while k < len(lista_blks) and len(paq) < 3:
                if lista_blks[k]["tipo"] in ("p", "blockquote", "ref"):
                    paq.append(k)
                    k += 1
                else:
                    break
            return paq if len(paq) >= 2 else None
        return [j]

    def _optimizar_flujo_bloques_pdf(lista_inicial: list[dict], regla_ini: str) -> tuple[list[dict], object]:
        """
        Ancla figuras y tablas a la parte superior o inferior de las páginas y deja fluir
        el texto (incluyendo referencias al cierre del artículo) en los espacios sobrantes
        para evitar vacíos al pie de página (como en la anteúltima página).
        """
        trabajo = _emparejar_figuras_columna(lista_inicial)
        doc_act, pags, bp, bbox, tr_pages = _renderizar_y_medir_pdf(trabajo, regla_ini)

        flujo_activo = str(meta.get("flujo_automatico", rev.get("flujo_automatico", "true"))).strip().lower() not in (
            "false", "0", "no", "off", "fijo"
        )
        if not flujo_activo or len(pags) <= 1:
            return trabajo, doc_act

        px_por_mm = 96.0 / 25.4
        ancla_orig = {}
        for b in trabajo:
            if b["tipo"] in ("FIGURA", "TABLA") and b["id"] in bp:
                ancla_orig[b["id"]] = min(bp[b["id"]])

        # Fase 0: Compactar tipografía/relleno de tablas altas para que entren enteras en 1 página;
        # solo si aun así superan 212 mm, permitir salto nativo de WeasyPrint con <thead> repetido
        promovida = False
        for idx, b in enumerate(trabajo):
            if b["tipo"] == "TABLA" and b["id"] in bbox:
                alt_mm = sum(box[4] for box in bbox[b["id"]]) / px_por_mm
                if alt_mm > 185.0:
                    cabe_en_una = False
                    for nivel_c in ("tabla-compacta-1", "tabla-compacta-2", "tabla-compacta-3"):
                        cand_c = [dict(x) for x in trabajo]
                        cand_c[idx]["html_pdf"] = cand_c[idx]["html_pdf"].replace(
                            'class="rma-tabla', f'class="rma-tabla {nivel_c}', 1
                        )
                        _, _, _, bbox_c, _ = _renderizar_y_medir_pdf(cand_c, regla_ini)
                        if b["id"] in bbox_c:
                            alt_c = sum(box[4] for box in bbox_c[b["id"]]) / px_por_mm
                            if alt_c <= 212.0:
                                trabajo[idx] = cand_c[idx]
                                promovida = True
                                cabe_en_una = True
                                break
                    if not cabe_en_una and not b.get("dividir_forzado", False):
                        trabajo[idx] = dict(b)
                        trabajo[idx]["modo_dividir"] = "divisible"
                        trabajo[idx]["html_pdf"] = (
                            trabajo[idx]["html_pdf"]
                            .replace('class="rma-tabla', 'class="rma-tabla tabla-compacta-2', 1)
                            .replace("tabla-indivisible", "tabla-divisible", 1)
                        )
                        promovida = True
        if promovida:
            doc_act, pags, bp, bbox, tr_pages = _renderizar_y_medir_pdf(trabajo, regla_ini)
            ancla_orig.update({b["id"]: min(bp[b["id"]]) for b in trabajo if b["tipo"] in ("FIGURA", "TABLA") and b["id"] in bp})

        def _ajustar_altura_figuras_saltadas(trab_in, doc_in, pags_in, bp_in, bbox_in):
            for idx, b in enumerate(trab_in):
                if b["tipo"] != "FIGURA" or not b.get("movible", True) or b.get("es_par"):
                    continue
                bid = b["id"]
                if bid not in bp_in or bid not in bbox_in:
                    continue
                p_fig = min(bp_in[bid])
                if p_fig == 0:
                    continue
                p_ant = pags_in[p_fig - 1]
                if p_ant["blks"] and p_ant["blks"][-1] != bid and p_ant["gap_avg_mm"] >= 40.0:
                    fig_h_mm = bbox_in[bid][0][4] / px_por_mm
                    deficit_mm = fig_h_mm - p_ant["gap_avg_mm"]
                    if 0 < deficit_mm <= 48.0:
                        base_img_h = (100.0 if b.get("clase_ancho") == "ancho-completo" else 185.0) if perfil == "tradicional" else (114.0 if b.get("clase_ancho") == "ancho-completo" else 98.0)
                        img_h_est = min(base_img_h, max(35.0, fig_h_mm - 24.0))
                        for extra_h in (4.0, 9.0, 15.0, 22.0):
                            nueva_h = max(46.0, img_h_est - (deficit_mm + extra_h))
                            cand = [dict(x) for x in trab_in]
                            if 'style="' in cand[idx]["html_pdf"]:
                                cand[idx]["html_pdf"] = re.sub(r'style="[^"]*"', f'style="max-height:{nueva_h:.1f}mm;"', cand[idx]["html_pdf"], count=1)
                            else:
                                cand[idx]["html_pdf"] = re.sub(r"(<im[^>]*)(>)", lambda m: f'{m.group(1)} style="max-height:{nueva_h:.1f}mm;"{m.group(2)}', cand[idx]["html_pdf"], count=1)
                            doc_c, pags_c, bp_c, bbox_c, _ = _renderizar_y_medir_pdf(cand, regla_ini)
                            if bid in bp_c and min(bp_c[bid]) == p_fig - 1 and len(pags_c) <= len(pags_in):
                                trab_in, doc_in, pags_in, bp_in, bbox_in = cand, doc_c, pags_c, bp_c, bbox_c
                                break
            return trab_in, doc_in, pags_in, bp_in, bbox_in

        # Fase 1: Ajuste elástico inicial de figuras que saltaron por <= 48 mm
        trabajo, doc_act, pags, bp, bbox = _ajustar_altura_figuras_saltadas(trabajo, doc_act, pags, bp, bbox)

        # Fase 2: Rellenar vacíos previos a figuras/tablas desplazadas y anclarlas al tope o pie de página
        for _ in range(110):
            cambio = False
            for idx, b in enumerate(trabajo):
                if b["tipo"] not in ("FIGURA", "TABLA") or not b.get("movible", True):
                    continue
                bid = b["id"]
                if bid not in bp or bid not in bbox:
                    continue
                p_fig = min(bp[bid])
                if p_fig > ancla_orig.get(bid, p_fig) + 1:
                    continue

                paquete = _obtener_paquete_texto(trabajo, idx + 1)
                if not paquete:
                    continue

                necesita_relleno_anterior = False
                if p_fig >= 1:
                    p_ant = pags[p_fig - 1]
                    if (p_ant["gap_avg_mm"] >= 10.0) or (
                        perfil == "tradicional" and max(p_ant["gap_l_mm"], p_ant["gap_r_mm"]) >= 16.0
                    ):
                        necesita_relleno_anterior = True

                _, _, y_fig, _, h_fig = bbox[bid][0]
                c_top_f = pags[p_fig]["c_top"]
                c_bot_f = pags[p_fig]["c_bot"]
                dist_top_mm = (y_fig - c_top_f) / px_por_mm
                dist_bot_mm = (c_bot_f - (y_fig + h_fig)) / px_por_mm
                en_medio_pagina = (dist_top_mm > 14.0) and (dist_bot_mm > 14.0)

                hueco_col = (perfil == "tradicional" and pags[p_fig]["gap_l_mm"] >= 28.0 and dist_top_mm <= 16.0)
                if not (necesita_relleno_anterior or en_medio_pagina or hueco_col):
                    continue

                cand = []
                for k, item in enumerate(trabajo):
                    if k == idx:
                        for pk in paquete:
                            cand.append(trabajo[pk])
                        cand.append(item)
                    elif k in paquete:
                        continue
                    else:
                        cand.append(item)

                doc_c, pags_c, bp_c, bbox_c, _ = _renderizar_y_medir_pdf(cand, regla_ini)
                if bid not in bp_c or bid not in bbox_c or len(pags_c) > len(pags):
                    continue
                p_fig_c = min(bp_c[bid])
                if p_fig_c > ancla_orig.get(bid, p_fig) + 1:
                    continue

                aceptar = False
                if necesita_relleno_anterior and p_fig >= 1:
                    p_ant = pags[p_fig - 1]
                    p_ant_c = pags_c[p_fig - 1] if p_fig - 1 < len(pags_c) else p_ant
                    gap_antes = p_ant["gap_avg_mm"]
                    gap_despues = p_ant_c["gap_avg_mm"]
                    deseq_antes = abs(p_ant["gap_l_mm"] - p_ant["gap_r_mm"])
                    deseq_despues = abs(p_ant_c["gap_l_mm"] - p_ant_c["gap_r_mm"])
                    if (gap_despues < gap_antes - 1.0) or (
                        gap_despues <= gap_antes + 0.5 and deseq_despues < deseq_antes - 3.0
                    ):
                        aceptar = True

                if not aceptar and en_medio_pagina and p_fig_c == p_fig:
                    _, _, y_fig_c, _, h_fig_c = bbox_c[bid][0]
                    dist_bot_c = (pags_c[p_fig_c]["c_bot"] - (y_fig_c + h_fig_c)) / px_por_mm
                    if dist_bot_c < dist_bot_mm - 2.0 and dist_bot_c >= -1.0:
                        aceptar = True

                if not aceptar and hueco_col and p_fig < len(pags_c) and pags_c[p_fig]["gap_l_mm"] < pags[p_fig]["gap_l_mm"] - 5.0:
                    aceptar = True

                if aceptar:
                    trabajo = cand
                    doc_act, pags, bp, bbox = doc_c, pags_c, bp_c, bbox_c
                    cambio = True
                    break

            if not cambio:
                break

        # Fase 1b: Re-ajustar altura de figuras que tras Fase 2 quedaron saltando con >= 40 mm libres arriba
        trabajo, doc_act, pags, bp, bbox = _ajustar_altura_figuras_saltadas(trabajo, doc_act, pags, bp, bbox)

        bloques_finales = trabajo
        return bloques_finales, doc_act

    shutil.copy2(css_origen, carpeta_salida / ".rma_print_tmp.css")
    try:
        regla_contador_pagina = f"@page :first {{ counter-reset: page {pagina_inicio}; }}"
        _k=int(meta.get("compactar_pdf") or 0);regla_contador_pagina+=(f"body{{line-height:{1.35-.03*_k:.2f}}}.cuerpo-articulo p{{margin-bottom:{2.1-.35*_k:.2f}mm}}.referencias .ref-item{{line-height:{1.26-.04*_k:.2f};margin-bottom:{1.4-.35*_k:.2f}mm}}.rma-figura{{margin:{5.5-.6*_k:.1f}mm 0}}" if _k else "")+str(meta.get("css_pdf_extra",""))
        _niv=str(meta.get("compactar_pdf","")).strip();_mp={"1":"body{line-height:1.32}.cuerpo-articulo p{margin-bottom:1.7mm}.referencias .ref-item{line-height:1.21;margin-bottom:1.0mm}","2":"body{line-height:1.29}.cuerpo-articulo p{margin-bottom:1.45mm}.referencias .ref-item{line-height:1.18;margin-bottom:0.8mm}.rma-figura{margin:4.5mm 0 5mm 0}","3":"body{line-height:1.26;font-size:8.85pt}.cuerpo-articulo p{margin-bottom:1.25mm}.referencias .ref-item{line-height:1.15;margin-bottom:0.65mm}.rma-figura{margin:4mm 0 4.5mm 0}"};regla_contador_pagina+=("\n"+_mp[_niv] if _niv in _mp else "")+("\n"+str(meta["css_pdf_extra"]) if meta.get("css_pdf_extra") else "")
        bloques_pdf_opt, doc_medido = _optimizar_flujo_bloques_pdf(bloques_pdf_base, regla_contador_pagina)
        cuerpo_render_pdf = [x["html_pdf"] for x in bloques_pdf_opt]

        total_paginas = len(doc_medido.pages)
        pagina_fin = pagina_inicio + total_paginas - 1

        fpage_str = f"{pagina_inicio:0{digitos_pag}d}"
        lpage_str = f"{pagina_fin:0{digitos_pag}d}"
        rango_paginas = formatear_rango_paginas(pagina_inicio, pagina_fin, digitos_pag)

        html_pdf_final = construir_documento_html(
            ".rma_print_tmp.css",
            pags_rango=rango_paginas,
            fpage_str=fpage_str,
            lpage_str=lpage_str,
            estilo_inline_extra=regla_contador_pagina,
            bloques_html_cuerpo=cuerpo_render_pdf,
        )
        HTML(string=html_pdf_final, base_url=str(carpeta_salida)).write_pdf(
            carpeta_salida / f"{nombre_salida}.pdf"
        )
    finally:
        (carpeta_salida / ".rma_print_tmp.css").unlink(missing_ok=True)

    html_ojs = construir_documento_html(
        "rma_ojs.css",
        pags_rango=rango_paginas,
        fpage_str=fpage_str,
        lpage_str=lpage_str,
    )
    auto_raw = meta.get("html_autocontenido", rev.get("html_autocontenido", True))
    if isinstance(auto_raw, bool):
        html_autocontenido = auto_raw
    else:
        html_autocontenido = str(auto_raw).strip().lower() not in ("false", "0", "no", "off", "linkeado", "externo")

    if html_autocontenido:
        html_ojs = empaquetar_html_autocontenido(html_ojs, carpeta_salida, dir_script / "rma_ojs.css")
        (carpeta_salida / f"{nombre_salida}.html").write_text(html_ojs, encoding="utf-8")
        (carpeta_salida / "rma_ojs.css").unlink(missing_ok=True)
        shutil.rmtree(carpeta_salida_fig, ignore_errors=True)
        shutil.rmtree(carpeta_salida_assets, ignore_errors=True)
    else:
        (carpeta_salida / f"{nombre_salida}.html").write_text(html_ojs, encoding="utf-8")
        shutil.copy2(dir_script / "rma_ojs.css", carpeta_salida / "rma_ojs.css")

    meta["fpage_str"] = fpage_str
    meta["lpage_str"] = lpage_str
    meta["paginas"] = rango_paginas
    xml_scielo = generar_xml_scielo(
        art,
        meta,
        rev,
        figuras_usadas,
        autores_est=autores_est,
        afiliaciones_est=afiliaciones_est,
        cuerpo_bloques=cuerpo_bloques,
    )
    (carpeta_salida / f"{nombre_salida}.xml").write_text(xml_scielo, encoding="utf-8")
    avisos_xml = validar_xml_scielo(xml_scielo)

    if actualizar_yaml and ruta_meta_yaml.exists():
        campos_sync = {
            "paginas": rango_paginas,
            "pagina_inicio": fpage_str,
            "pagina_fin": lpage_str,
        }
        if meta_overrides:
            for k in ("codigo_publicacion", "seccion", "volumen", "numero", "anio", "formato"):
                if k in meta_overrides and str(meta_overrides[k]).strip() != "":
                    campos_sync[k] = str(meta_overrides[k]).strip()
        actualizar_yaml_campos(ruta_meta_yaml, campos_sync)

    print(
        f"Maquetación completada [{perfil.upper()}] para ID OJS '{art_id}' "
        f"(pp. {rango_paginas}, {total_paginas} págs.) en: {carpeta_salida}"
    )
    print(f"  - {nombre_salida}.pdf")
    if html_autocontenido:
        print(f"  - {nombre_salida}.html (autocontenido: CSS + imágenes embebidas en Base64)")
    else:
        print(f"  - {nombre_salida}.html (+ rma_ojs.css + assets/ + figuras/)")
    if avisos_xml:
        print(f"  - {nombre_salida}.xml [Avisos SPS 1.9: {'; '.join(avisos_xml)}]")
    else:
        print(f"  - {nombre_salida}.xml (JATS 1.1 / SciELO SPS 1.9 verificado)")
    sin_ref, no_citadas = motor_citas.diagnosticar()
    if sin_ref:
        print(f"  [!] Citas en texto SIN referencia ({len(sin_ref)}):")
        for item in sin_ref:
            print(f"      * {item}")
    if no_citadas:
        print(f"  [!] Referencias NO citadas en texto ({len(no_citadas)}):")
        for item in no_citadas:
            print(f"      * {item}")
    print(f"  - Citas bidireccionales enlazadas: {motor_citas._contador_citas}")

    return {
        "id_ojs": art_id,
        "codigo_publicacion": cod_pub,
        "nombre_salida": nombre_salida,
        "seccion": seccion,
        "titulo": titulo_plano,
        "formato": perfil,
        "pagina_inicio": pagina_inicio,
        "pagina_fin": pagina_fin,
        "total_paginas": total_paginas,
        "paginas": rango_paginas,
    }



def _norm_clave_autor(s: str) -> str:
    s = unicodedata.normalize("NFD", s.lower())
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    s = re.sub(r"[’'`´]", "", s)
    s = re.sub(r"[^a-z0-9\s\-]", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def _extraer_primer_apellido_ref(autores_raw: str) -> str:
    s = re.sub(r"\s+", " ", autores_raw).strip().rstrip(".,;")
    if not s:
        return ""
    if "," in s:
        ap = s.split(",", 1)[0].strip()
        return _norm_clave_autor(ap)
    # Autor institucional / sigla sin coma (ej. CONICET, INDEC, UNESCO)
    return _norm_clave_autor(s)

def _extraer_apellido_probable_cita(autor_cita: str) -> str:
    """
    Extrae la clave de apellido para cruzar citas en texto con referencias bibliográficas,
    manejando tanto 'Apellido (2026)' como 'Nombre Apellido (2026)' o 'de Apellido (2026)'.
    """
    s = re.sub(r"\s+", " ", autor_cita or "").strip(" ,.;:()[]\"“”‘’")
    s = re.sub(
        r"^(?:vease|véase|ver|cf\.?|cfr\.?|en|in|citado\s+en|segun|según|por|e\.g\.?|por\s+ejemplo)\s+",
        "",
        s,
        flags=re.I,
    ).strip()
    s = re.sub(r"\s+et\.?\s+al\.?$", "", s, flags=re.I).strip()
    primer_bloque = re.split(r"\s*[,;]\s*|\s+(?:y|and|e|&)\s+", s, maxsplit=1)[0].strip()
    partes = primer_bloque.split()
    if not partes:
        return ""
    particulas = {"de", "da", "do", "dos", "das", "del", "la", "las", "los", "van", "von"}
    while len(partes) > 1 and _norm_clave_autor(partes[0]) in particulas:
        partes.pop(0)
    if not partes:
        return ""
    if len(partes) >= 2:
        return _norm_clave_autor(partes[-1])
    return _norm_clave_autor(partes[0])


_PAT_ENTIDAD_NO_AUTOR = re.compile(
    r"^(?:"
    r"brasil|rio\s+de\s+janeiro|sao\s+paulo|são\s+paulo|belo\s+horizonte|brasilia|brasília|salvador|olinda|ouro\s+preto|goias|goiás|paraty|valongo|pantanal|"
    r"iphan|unesco|ibge|funarte|minc|conicet|uba|unc|"
    r"centro\s+historico|centro\s+histórico|conjunto\s+moderno|parque\s+nacional|sitio\s+roberto|sítio\s+roberto|"
    r"salvaguarda\s+da|oral\s+e\s+imaterial|cultural\s+imaterial|cirio\s+de|círio\s+de|roda\s+de|frevo|descobrimento|sudeste|nacionais\s+da|baiano|espirito\s+santo|espírito\s+santo|luis|luís|mar|dia|fernando\s+de\s+noronha"
    r")\b",
    flags=re.I,
)

def _extraer_primer_apellido_cita(autor_cita: str) -> str:
    s = re.sub(r"\s+", " ", autor_cita).strip(" ,.;:()[]")
    s = re.sub(
        r"^(?:véase|vease|ver|cf\.?|cfr\.?|en|in|citado en|según|segun|por|e\.g\.?|por ejemplo)\s+",
        "",
        s,
        flags=re.I,
    ).strip()
    s = re.sub(r"\s+et\s+al\.?$", "", s, flags=re.I).strip()
    s = re.split(r"\s*,\s*|\s+(?:y|and|e)\s+|\s*&\s*", s, maxsplit=1)[0].strip()
    return _norm_clave_autor(s)


def _limpiar_html_citas(raw_h: str) -> str:
    s = re.sub(r"<sup\b[^>]*>.*?</sup>", "", raw_h, flags=re.I | re.S)
    s = re.sub(r"</?(?:p|div|li|tr|td|th|h[1-6]|blockquote|br|hr)\b[^>]*>", " ", raw_h, flags=re.I)
    s = re.sub(r"<[^>]+>", "", s)
    s = html.unescape(s)
    s = s.replace("\u200b", "").replace("\xad", "")
    return re.sub(r"\s+", " ", s).strip()


def _norm_lista_autores_completa(autores_raw: str) -> str:
    s = re.sub(r"\s*(?:&|\band\b|\be\b)\s*", " y ", autores_raw, flags=re.I)
    return _norm_clave_autor(s)


_PAT_YEAR_REF = re.compile(
    r"(?:"
    r"\(\s*([^()]{0,35}?\b(?:\d{4}\s*[a-z]?|s\.f\.|en prensa|in press)\b[^()]{0,15})\s*\)"
    r"|"
    r"(?<=[.\s])\s*(\d{4}\s*[a-z]?|s\.f\.|en prensa|in press)(?=\s*\.)"
    r")",
    flags=re.I,
)


def validar_formato_referencias_rma(referencias_html: list[str]) -> tuple[list[dict], list[str]]:
    """
    Revisa cada entrada de Referencias bibliográficas según las normas editoriales de la RMA:
      1. Año entre paréntesis seguido inmediatamente de punto: '(Año).'
      2. Uso de 'y' en lugar de '&' (o 'and') entre autores (ignorando URLs/DOIs).
      3. Sin espacio entre el año y la letra diferenciadora (ej. '(2020a).' y no '(2020 a).').
      4. Desambiguación con letras 'a', 'b', ... cuando coinciden autor(es) y año.
    """
    hallazgos = []
    grupos_autor_anio = {}

    for idx, r_html in enumerate(referencias_html, 1):
        txt = _limpiar_html_citas(r_html)
        txt_sin_num = re.sub(r"^(?:\[\d+\]|\d+[.)])\s+", "", txt).strip()
        if not txt_sin_num:
            continue

        problemas_entry = []
        m_y = _PAT_YEAR_REF.search(txt_sin_num)

        if not m_y:
            problemas_entry.append("No se encontró un bloque de año válido con formato '(Año).'")
            hallazgos.append({
                "idx": idx,
                "preview": txt_sin_num[:90] + ("..." if len(txt_sin_num) > 90 else ""),
                "problemas": problemas_entry,
            })
            continue

        autores_raw = txt_sin_num[: m_y.start()].strip()
        resto_tras_anio = txt_sin_num[m_y.end() :]
        primer_ap = _extraer_primer_apellido_ref(autores_raw)

        # 1) Uso de 'y' en lugar de '&' o 'and' entre autores
        if "&" in autores_raw:
            problemas_entry.append("Usa '&' en lugar de 'y' entre autores")
        elif re.search(r"\band\b", autores_raw, flags=re.I):
            problemas_entry.append("Usa 'and' en lugar de 'y' entre autores")

        # 2) Año entre paréntesis seguido de punto: '(Año).'
        if m_y.group(2) is not None:
            anio_suelto = m_y.group(2).strip()
            problemas_entry.append(f"Año sin paréntesis: '{anio_suelto}.' (debe ir como '({anio_suelto}).')")
            bloque_anio = anio_suelto
        else:
            bloque_anio = (m_y.group(1) or "").strip()
            if not resto_tras_anio.startswith("."):
                problemas_entry.append(
                    f"Falta punto después del paréntesis del año: '({bloque_anio})' (debe ser '({bloque_anio}).')"
                )

        # 3) Espacio indebido entre año y letra: ej. (2020 a) -> (2020a)
        m_esp_letra = re.search(r"\b(\d{4})\s+([a-z])\b", bloque_anio, flags=re.I)
        if m_esp_letra:
            problemas_entry.append(
                f"Espacio entre año y letra: '({m_esp_letra.group(1)} {m_esp_letra.group(2)})' (debe ser '({m_esp_letra.group(1)}{m_esp_letra.group(2).lower()}).')"
            )

        # 4) Agrupar por primer apellido + año numérico puro (sin letra a/b)
        encontrados = [
            re.sub(r"\s+", "", a).lower()
            for a in re.findall(r"\b(?:\d{4}\s*[a-z]?|s\.f\.|en prensa|in press)\b", bloque_anio, flags=re.I)
        ]
        anio_pub = encontrados[-1] if encontrados else ""
        if primer_ap and anio_pub and re.fullmatch(r"\d{4}", anio_pub):
            autores_norm_completo = _norm_lista_autores_completa(autores_raw)
            es_et_al = (autores_raw.count(",") >= 3) or bool(re.search(r"\bet\s+al\b", autores_raw, flags=re.I))
            grupos_autor_anio.setdefault((primer_ap, anio_pub), []).append({
                "idx": idx,
                "autores_norm": autores_norm_completo,
                "autores_raw": autores_raw,
                "es_et_al": es_et_al,
            })

        if problemas_entry:
            hallazgos.append({
                "idx": idx,
                "preview": txt_sin_num[:90] + ("..." if len(txt_sin_num) > 90 else ""),
                "problemas": problemas_entry,
            })

    colisiones_ab = []
    for (primer_ap, anio_pub), items in sorted(grupos_autor_anio.items(), key=lambda kv: kv[1][0]["idx"]):
        if len(items) < 2:
            continue
    idxs_str = ", ".join(f"[{it['idx']}]" for it in items)
    mismos_autores = len({it["autores_norm"] for it in items}) == 1
    todos_et_al = all(it["es_et_al"] for it in items)
    if mismos_autores:
        colisiones_ab.append(
            f"Referencias {idxs_str} comparten autor(es) y año ({anio_pub}) sin letra a/b -> usar ({anio_pub}a), ({anio_pub}b)..."
        )
    elif todos_et_al:
        ap_vis = items[0]["autores_raw"].split(",")[0].strip()
        colisiones_ab.append(
            f"Referencias {idxs_str} comparten primer autor '{ap_vis}' (3+ autores -> et al.) y año ({anio_pub}) sin letra a/b -> revisar si requiere ({anio_pub}a), ({anio_pub}b)"
        )

    return hallazgos, colisiones_ab


def _clave_autor_flexible(clave: str) -> str:
    import re
    c = (clave or "").strip().lower()
    if not c:
        return ""
    return re.sub(r"([aeiou])z([aeiou])", r"\1s\2", c)


def auditar_citas_y_referencias(art: dict, cuerpo_bloques: list[dict] | None = None) -> dict:
    """
    Cruza las citas detectadas en cuerpo, tablas, epígrafes y notas al pie con
    las entradas de Referencias bibliográficas, y valida el formato RMA de cada entrada.
    """
    refs_raw_list = art.get("referencias", [])
    errores_formato_rma, colisiones_ab = validar_formato_referencias_rma(refs_raw_list)

    refs_parsed = []
    indice_ref_por_clave = {}
    anios_por_autor = {}

    for idx, r_html in enumerate(refs_raw_list, 1):
        txt = _limpiar_html_citas(r_html)
        txt_sin_num = re.sub(r"^(?:\[\d+\]|\d+[.)])\s+", "", txt)

        m_y = _PAT_YEAR_REF.search(txt_sin_num)
        if not m_y:
            refs_parsed.append({
                "idx": idx,
                "texto": txt_sin_num,
                "autor_clave": "",
                "anios": [],
                "citada": False,
            })
            continue

        autores_raw = txt_sin_num[: m_y.start()].strip()
        autor_clave = _extraer_primer_apellido_ref(autores_raw)
        bloque_anio = (m_y.group(1) or m_y.group(2) or "").strip()

        encontrados = [
            re.sub(r"\s+", "", a).lower()
            for a in re.findall(r"\b(?:\d{4}\s*[a-z]?|s\.f\.|en prensa|in press)\b", bloque_anio, flags=re.I)
        ]
        anio_pub = encontrados[-1] if encontrados else ""
        anios = list(dict.fromkeys(([anio_pub] if anio_pub else []) + encontrados))

        info_ref = {
            "idx": idx,
            "texto": txt_sin_num,
            "autor_clave": autor_clave,
            "anios": anios,
            "citada": False,
        }
        refs_parsed.append(info_ref)

        if autor_clave and anios:
            anios_por_autor.setdefault(autor_clave, set()).update(anios)
            for a in anios:
                indice_ref_por_clave.setdefault((autor_clave, a), []).append(info_ref)

    fragmentos_texto = []
    celdas_tabla = []
    fuentes_bloques = cuerpo_bloques if cuerpo_bloques is not None else art.get("bloques_cuerpo", [])
    for b in fuentes_bloques:
        raw_h = b.get("html_pdf") or b.get("html_ojs") or b.get("html") or b.get("texto") or ""
        if raw_h:
            fragmentos_texto.append(_limpiar_html_citas(raw_h))
            for m_td in re.finditer(r"<t[dh]\b[^>]*>(.*?)</t[dh]>", raw_h, flags=re.I | re.S):
                celdas_tabla.append(_limpiar_html_citas(m_td.group(1)))
    if cuerpo_bloques is None:
        for t_html in art.get("tablas_embebidas", []):
            fragmentos_texto.append(_limpiar_html_citas(t_html))
            for m_td in re.finditer(r"<t[dh]\b[^>]*>(.*?)</t[dh]>", t_html, flags=re.I | re.S):
                celdas_tabla.append(_limpiar_html_citas(m_td.group(1)))
    for item_nota in art.get("notas_usadas", []):
        n_html = item_nota[1] if isinstance(item_nota, (tuple, list)) else str(item_nota)
        fragmentos_texto.append(_limpiar_html_citas(n_html))

    texto_completo = "\n".join(fragmentos_texto)

    pat_apellido = (
        r"(?:(?:de|del|de\s+la|de\s+los|van|von|da|dos|das|le|la|Del|De\s+la)\s+)?"
        r"[A-ZÁÉÍÓÚÑÜ][A-Za-zÁÉÍÓÚÑÜáéíóúñü'’\-]+"
        r"(?:\s+(?:(?:de|del|de\s+la|de\s+los|van|von|da|dos|das)\s+)?[A-ZÁÉÍÓÚÑÜ][A-Za-zÁÉÍÓÚÑÜáéíóúñü'’\-]+){0,2}"
    )
    pat_autor_cita = (
        rf"{pat_apellido}"
        rf"(?:"
        rf"\s+et\.?\s+al\s*\.?"
        rf"|(?:,\s*{pat_apellido})*\s*(?:,?\s*(?:y|and|e|&)\s+{pat_apellido})"
        rf")?"
    )
    pat_anio = r"(?:\d{4}[a-z]?|s\.f\.|en prensa|in press)"
    pat_bloque_anios = (
        rf"(?:\[\s*\d{{4}}(?:\s*[-–—]\s*\d{{2,4}})?\s*\]\s*)?"
        rf"(?:{pat_anio})(?:\s*(?:,|;|y|and)\s*(?:\[\s*\d{{4}}(?:\s*[-–—]\s*\d{{2,4}})?\s*\]\s*)?(?:{pat_anio}))*"
    )

    stop_autores = {
        "figura", "tabla", "fig", "tab", "figure", "figures", "table", "tables", "cap", "capitulo", "ano", "anos",
        "en", "para", "desde", "hasta", "entre", "durante", "siglo", "siglos",
    }
    citas_detectadas = []

    for m_par in re.finditer(r"\(([^()]{4,400})\)", texto_completo):
        contenido = m_par.group(1)
        if not re.search(pat_anio, contenido, flags=re.I):
            continue
        for m_c in re.finditer(
            rf"({pat_autor_cita})(?:\s*,\s*|\s+)({pat_bloque_anios})(?=$|[\s:,.\];])",
            contenido,
        ):
            autor_txt = m_c.group(1).strip()
            anios_txt = m_c.group(2)
            autor_cl = _extraer_primer_apellido_cita(autor_txt)
            if not autor_cl or autor_cl in stop_autores:
                continue
            for a in re.findall(pat_anio, anios_txt, flags=re.I):
                citas_detectadas.append((f"{autor_txt} ({a})", autor_cl, a.lower().strip()))

    for m_nar in re.finditer(
        rf"({pat_autor_cita})\s*\(\s*({pat_bloque_anios})(?:\s*[: ,][^)]*)?\)",
        texto_completo,
    ):
        autor_txt = m_nar.group(1).strip()
        anios_txt = m_nar.group(2)
        autor_cl = _extraer_primer_apellido_cita(autor_txt)
        if not autor_cl or autor_cl in stop_autores:
            continue
        for a in re.findall(pat_anio, anios_txt, flags=re.I):
            citas_detectadas.append((f"{autor_txt} ({a})", autor_cl, a.lower().strip()))

    for celda in celdas_tabla:
        for m_td_c in re.finditer(
            rf"({pat_autor_cita})(?:\s*,\s*|\s+)({pat_bloque_anios})(?=$|[\s:,.\];])",
            celda,
        ):
            autor_txt = m_td_c.group(1).strip()
            anios_txt = m_td_c.group(2)
            autor_cl = _extraer_primer_apellido_cita(autor_txt)
            if not autor_cl or autor_cl in stop_autores:
                continue
            for a in re.findall(pat_anio, anios_txt, flags=re.I):
                citas_detectadas.append((f"{autor_txt} ({a})", autor_cl, a.lower().strip()))

    citas_sin_ref = {}
    for etiqueta, autor_cl, a_norm in citas_detectadas:
        clave = (autor_cl, a_norm)
        if clave in indice_ref_por_clave:
            for r_obj in indice_ref_por_clave[clave]:
                r_obj["citada"] = True
        else:
            if clave not in citas_sin_ref:
                otros = sorted(anios_por_autor.get(autor_cl, []))
                sugerencia = f" (en Bibliografía figura con año: {', '.join(otros)})" if otros else ""
                citas_sin_ref[clave] = f"{etiqueta}{sugerencia}"

    refs_no_citadas = []
    for r in refs_parsed:
        if not r["citada"]:
            muestra = ""
            if r["autor_clave"]:
                m_dbg = re.search(rf".{{0,25}}\b{re.escape(r['autor_clave'].split()[0])}\b.{{0,35}}", texto_completo, flags=re.I)
                if m_dbg:
                    muestra = f" [en texto aparece: '{m_dbg.group(0).strip()}']"
            refs_no_citadas.append(
                f"[{r['idx']}] {r['texto'][:90]}{'...' if len(r['texto']) > 90 else ''}{muestra}"
            )

    return {
        "total_citas": len(citas_detectadas),
        "total_refs": len(refs_parsed),
        "citas_sin_ref": list(citas_sin_ref.values()),
        "refs_no_citadas": refs_no_citadas,
        "posibles_duplicados": colisiones_ab,
        "errores_formato_rma": errores_formato_rma,
    }


def imprimir_auditoria_citas(art: dict, cuerpo_bloques: list[dict] | None = None) -> None:
    rep = auditar_citas_y_referencias(art, cuerpo_bloques=cuerpo_bloques)
    print(
        f"  - Control citas/bibliografía: {rep['total_citas']} citas en texto vs. "
        f"{rep['total_refs']} referencias"
    )
    if rep["citas_sin_ref"]:
        print(f"    [!] Citas en texto SIN referencia en bibliografía ({len(rep['citas_sin_ref'])}):")
        for item in rep["citas_sin_ref"]:
            print(f"        * {item}")
    if rep["refs_no_citadas"]:
        print(f"    [!] Referencias en bibliografía NO citadas en texto ({len(rep['refs_no_citadas'])}):")
        for item in rep["refs_no_citadas"]:
            print(f"        * {item}")
    if rep["errores_formato_rma"]:
        print(f"    [!] Formato bibliográfico RMA ({len(rep['errores_formato_rma'])} entrada/s con observaciones):")
        for item in rep["errores_formato_rma"]:
            print(f"        * [{item['idx']}] {item['preview']}")
            for prob in item["problemas"]:
                print(f"            - {prob}")
    if rep["posibles_duplicados"]:
        print(f"    [!] Posibles referencias mismo autor/año sin letra a/b ({len(rep['posibles_duplicados'])}):")
        for item in rep["posibles_duplicados"]:
            print(f"        * {item}")
    if (
        not rep["citas_sin_ref"]
        and not rep["refs_no_citadas"]
        and not rep["errores_formato_rma"]
        and not rep["posibles_duplicados"]
    ):
        print("    [OK] Todas las citas y referencias coinciden y cumplen el formato RMA.")


if __name__ == "__main__":
    if not sys.argv[1:]:
        from nuevo_articulo import main as main_interactivo
        main_interactivo()
    else:
        carpeta_obj, id_ojs, p_ini_cli, fmt_cli, interactivo_cli = resolver_destino(sys.argv[1:])
        ruta_meta_check = carpeta_obj / f"{id_ojs}_metadatos.yaml"
        if interactivo_cli or not ruta_meta_check.exists():
            from nuevo_articulo import ejecutar_flujo_interactivo
            carpeta_obj, _ = ejecutar_flujo_interactivo(carpeta_obj, id_ojs)
        compilar_articulo(carpeta_obj, id_ojs, pagina_inicio_override=p_ini_cli, formato_override=fmt_cli)
