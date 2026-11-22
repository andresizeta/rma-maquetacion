#!/usr/bin/env python3
"""
Motor General de Citación Bidireccional y Normalización Unicode para RMA.
Compatible con Python 3.10+ (solo biblioteca estándar).
"""

from __future__ import annotations

import html
import re
import unicodedata
from dataclasses import dataclass, field

# Caracteres latinos que NFKD no descompone en ASCII base
_MAPA_LATIN_EXTRA = str.maketrans({
    "ß": "ss",
    "ø": "o",
    "Ø": "o",
    "ł": "l",
    "Ł": "l",
    "æ": "ae",
    "Æ": "ae",
    "œ": "oe",
    "Œ": "oe",
    "ð": "d",
    "Ð": "d",
    "þ": "th",
    "Þ": "th",
    "đ": "d",
    "Đ": "d",
    "ı": "i",
    "İ": "i",
    "ħ": "h",
    "Ħ": "h",
    "ŧ": "t",
    "Ŧ": "t",
    "ŋ": "ng",
    "Ŋ": "ng",
})

# Rangos Unicode completos para alfabetos latinos (básico + suplemento + extendido A/B/adicional)
LETRA_MAY = r"[A-Z\u00C0-\u00D6\u00D8-\u00DE\u0100-\u024F\u1E00-\u1EFF]"
LETRA_ALL = r"[A-Za-z\u00C0-\u00D6\u00D8-\u00F6\u00F8-\u024F\u1E00-\u1EFF]"
APOSTROFO = r"['’‘ʼ´`]"
GUION = r"[-‐‑‒–—]"
PARTICULA_SURNAME = (
    r"(?:(?:de\s+la|de\s+las|de\s+los|van\s+der|van\s+den|von\s+der|"
    r"de|del|da|das|dos|di|du|des|la|le|les|van|von|der|den|ten|ter|al|el|bin|ibn|mac|mc|o|st\.?|san)\s+)"
)

# Palabras capitalizadas al inicio de oración que nunca deben absorberse dentro del enlace
_STOPWORDS_INICIO_ORACION = {
    "segun", "como", "para", "entre", "tanto", "mientras", "aunque", "asimismo",
    "ademas", "por", "en", "de", "desde", "hasta", "sobre", "tras", "durante",
    "mediante", "ante", "bajo", "contra", "hacia", "sin", "con", "asi", "donde",
    "cuando", "quien", "quienes", "cual", "cuales", "este", "esta", "estos",
    "estas", "ese", "esa", "esos", "esas", "aquel", "aquella", "otro", "otra",
    "otros", "otras", "algunos", "algunas", "varios", "varias", "muchos", "muchas",
    "diversos", "diversas", "autores", "investigadores", "trabajos", "estudios",
    "vease", "ver", "cfr", "cf", "sensu", "apud", "ibid", "ibidem", "op",
    "according", "see", "as", "in", "for", "both", "while", "however", "whereas",
    "although", "since", "after", "before", "during", "among", "between", "from",
    "by", "with", "without", "following", "unlike", "similarly", "furthermore",
    "moreover", "nevertheless", "nonetheless", "indeed", "thus", "hence", "therefore",
}


def limpiar_texto_unicode(texto: str) -> str:
    """Desescapa entidades HTML y normaliza espacios invisibles y composición Unicode."""
    if not texto:
        return ""
    t = html.unescape(texto)
    t = unicodedata.normalize("NFC", t)
    t = re.sub(r"[\u00A0\u2000-\u200B\u202F\u205F\u3000]", " ", t)
    return t


def normalizar_clave_autor(texto: str) -> str:
    """
    Convierte cualquier apellido (con diéresis, cedillas, apóstrofos, guiones o letras
    del bloque Latino Extendido) a una clave ASCII limpia en minúsculas.
    """
    t = limpiar_texto_unicode(texto).translate(_MAPA_LATIN_EXTRA)
    t = unicodedata.normalize("NFKD", t)
    t = "".join(c for c in t if unicodedata.category(c) != "Mn")
    t = t.lower()
    t = re.sub(r"[^a-z0-9]+", " ", t)
    return re.sub(r"\s+", " ", t).strip()


def normalizar_slug_id(texto: str) -> str:
    """Genera un identificador HTML seguro separado por guiones (ej. 'd-angelo-del-campo')."""
    return normalizar_clave_autor(texto).replace(" ", "-")


@dataclass
class EntradaReferencia:
    idx: int
    ref_id: str
    primer_apellido: str
    clave_primer_apellido: str
    claves_secundarias: list[str] = field(default_factory=list)
    segundo_apellido_clave: str = ""
    anio_norm: str = ""
    anio_visible: str = ""
    texto_plano: str = ""
    citas_ids: list[str] = field(default_factory=list)


@dataclass
class ResultadoCitacion:
    cuerpo_html: str
    referencias_html: list[str]
    total_citas_enlazadas: int
    citas_sin_referencia: list[str]
    referencias_no_citadas: list[str]
    entradas: list[EntradaReferencia]


class MotorCitacionesRMA:
    """Motor reutilizable para indexar bibliografía, enlazar citas y generar backlinks."""

    RE_ANIO = r"(?:(?:19|20)\d{2}[a-z]?|s\s*/\s*f|s\.\s*f\.|n\.\s*d\.|en\s+prensa|in\s+press)"
    RE_PAG = (
        r"(?:\s*:\s*\d+(?:\s*[-–—]\s*\d+)?"
        r"|\s*,\s*(?:pp?|págs?|figs?|láms?|tabs?)\.?\s*\d+(?:\s*[-–—]\s*\d+)?)?"
    )

    def __init__(self, referencias_html: list[str], estilo_id: str = "semantico"):
        """
        estilo_id:
          - 'semantico': id="ref-apellido-anio" (ej. ref-knusel-1997, ref-d-angelo-del-campo-2017)
          - 'numerico':  id="ref-1", id="ref-2", ...
        """
        self.estilo_id = estilo_id
        self.entradas: list[EntradaReferencia] = []
        self.indice_primario: dict[tuple[str, str], list[EntradaReferencia]] = {}
        self.indice_secundario: dict[tuple[str, str], list[EntradaReferencia]] = {}
        self._citas_no_encontradas: list[str] = []
        self._contador_citas = 0
        self._indexar_referencias(referencias_html)

    @staticmethod
    def _limpiar_backlinks_previos(ref_html: str) -> str:
        """Elimina enlaces de retorno previos para garantizar idempotencia."""
        return re.sub(
            r'\s*<a\b[^>]*class=["\'][^"\']*ref-backlink[^"\']*["\'][^>]*>.*?</a>',
            "",
            ref_html,
            flags=re.I | re.S,
        ).strip()

    def _indexar_referencias(self, referencias_html: list[str]) -> None:
        re_anio_b = re.compile(rf"\b({self.RE_ANIO})\b", flags=re.I)
        ids_usados: set[str] = set()

        for idx, r_html in enumerate(referencias_html, start=1):
            r_limpia = self._limpiar_backlinks_previos(r_html)
            plano = limpiar_texto_unicode(re.sub(r"<[^>]+>", " ", r_limpia))
            plano = re.sub(r"\s+", " ", plano).strip()

            m_anio = re_anio_b.search(plano)
            if not m_anio:
                ref_id = f"ref-{idx}"
                ids_usados.add(ref_id)
                self.entradas.append(
                    EntradaReferencia(
                        idx=idx,
                        ref_id=ref_id,
                        primer_apellido="",
                        clave_primer_apellido="",
                        texto_plano=plano,
                    )
                )
                continue

            anio_visible = m_anio.group(1).strip()
            anio_norm = re.sub(r"[\s\.]+", "", anio_visible.lower())
            cabecera = plano[: m_anio.start()].strip(" ,.;:()")

            autores_partes = re.split(r"\s+(?:y|e|and|&)\s+|;\s*", cabecera, flags=re.I)
            primer_bloque = autores_partes[0].strip()
            primer_apellido = primer_bloque.split(",")[0].strip()
            clave_primaria = normalizar_clave_autor(primer_apellido)

            segundo_apellido_clave = ""
            if len(autores_partes) >= 2:
                seg_ap = autores_partes[1].strip().split(",")[0].strip()
                segundo_apellido_clave = normalizar_clave_autor(seg_ap)
            else:
                m_seg = re.search(
                    rf",\s*(?:{LETRA_MAY}\.?\s*)+,\s*({LETRA_MAY}{LETRA_ALL}+(?:[\s'\-]{LETRA_MAY}{LETRA_ALL}+)*)\s*,",
                    cabecera,
                )
                if m_seg:
                    segundo_apellido_clave = normalizar_clave_autor(m_seg.group(1))

            if self.estilo_id == "semantico" and clave_primaria:
                slug_base = f"ref-{normalizar_slug_id(primer_apellido)}-{normalizar_slug_id(anio_visible)}"
                ref_id = slug_base
                suf = 2
                while ref_id in ids_usados:
                    ref_id = f"{slug_base}-{suf}"
                    suf += 1
            else:
                ref_id = f"ref-{idx}"
            ids_usados.add(ref_id)

            claves_sec: list[str] = []
            partes = clave_primaria.split()
            if len(partes) > 1:
                if len(partes[-1]) >= 3:
                    claves_sec.append(partes[-1])
                if len(partes[0]) >= 3 and partes[0] not in {
                    "de", "del", "la", "los", "las", "van", "von", "da", "dos", "san"
                }:
                    claves_sec.append(partes[0])

            entrada = EntradaReferencia(
                idx=idx,
                ref_id=ref_id,
                primer_apellido=primer_apellido,
                clave_primer_apellido=clave_primaria,
                claves_secundarias=claves_sec,
                segundo_apellido_clave=segundo_apellido_clave,
                anio_norm=anio_norm,
                anio_visible=anio_visible,
                texto_plano=plano,
            )
            self.entradas.append(entrada)

            if clave_primaria:
                self.indice_primario.setdefault((clave_primaria, anio_norm), []).append(entrada)
                for cs in claves_sec:
                    self.indice_secundario.setdefault((cs, anio_norm), []).append(entrada)

    def _limpiar_palabras_iniciales_no_autor(self, autor_txt: str) -> tuple[str, int]:
        """Evita que palabras iniciales de oración ('Según', 'Como', 'Para'...) entren al enlace."""
        texto = autor_txt
        offset = 0
        while True:
            m = re.match(rf"^({LETRA_MAY}{LETRA_ALL}+)(\s+(?:de|del|la|los|las)?\s*)", texto)
            if not m:
                break
            primera_norm = normalizar_clave_autor(m.group(1))
            resto = texto[m.end():]
            if primera_norm in _STOPWORDS_INICIO_ORACION and re.match(rf"^{LETRA_MAY}", resto):
                offset += m.end()
                texto = resto
            else:
                break
        return texto, offset

    def _buscar_entrada(self, autor_txt: str, anio_txt: str) -> EntradaReferencia | None:
        anio_norm = re.sub(r"[\s\.]+", "", anio_txt.lower())
        partes_autores = re.split(
            r"\s+(?:y|e|and|&)\s+|,?\s*et\.?\s+al\.?|,?\s*y\s+(?:coautores?|colaboradores?|cols?\.|otros)|,?\s*(?:and|&)\s+(?:colleagues|co-?authors|others)",
            autor_txt,
            flags=re.I,
        )
        primer = partes_autores[0].strip().rstrip(",")
        segundo = partes_autores[1].strip().rstrip(",") if len(partes_autores) > 1 else ""

        clave1 = normalizar_clave_autor(primer)
        clave2 = normalizar_clave_autor(segundo)
        if not clave1:
            return None

        candidatos = self.indice_primario.get((clave1, anio_norm))
        if not candidatos:
            candidatos = self.indice_secundario.get((clave1, anio_norm))
        if not candidatos:
            tokens = clave1.split()
            if len(tokens) > 1:
                for tok in (tokens[-1], tokens[0]):
                    if len(tok) >= 3 and (tok, anio_norm) in self.indice_primario:
                        candidatos = self.indice_primario[(tok, anio_norm)]
                        break
                if not candidatos:
                    for tok in (tokens[-1], tokens[0]):
                        if len(tok) >= 3 and (tok, anio_norm) in self.indice_secundario:
                            candidatos = self.indice_secundario[(tok, anio_norm)]
                            break

        if not candidatos:
            return None

        if len(candidatos) == 1:
            return candidatos[0]

        # Desambiguación por segundo apellido cuando el mismo primer autor tiene >1 obra en el mismo año
        if clave2:
            for c in candidatos:
                if c.segundo_apellido_clave and (
                    clave2 == c.segundo_apellido_clave
                    or clave2 in c.segundo_apellido_clave
                    or c.segundo_apellido_clave in clave2
                ):
                    return c
        return candidatos[0]

    @staticmethod
    def _mapear_texto_plano_html(seg_html: str) -> tuple[str, list[tuple[int, int]], list[bool]]:
        """Proyecta HTML inline a texto plano conservando spans originales y máscara de protección."""
        token_re = re.compile(r"<[^>]+>|&(?:#\d+|#x[0-9a-fA-F]+|[a-zA-Z]+);|.", flags=re.S)
        plain_chars: list[str] = []
        spans: list[tuple[int, int]] = []
        protegido: list[bool] = []
        prof_excl = 0

        for m in token_re.finditer(seg_html):
            tok = m.group(0)
            s, e = m.span()
            if tok.startswith("<"):
                if re.match(r"^<(?:a|sup|sub|code|pre|script|style)\b", tok, flags=re.I) and not tok.endswith("/>"):
                    prof_excl += 1
                elif re.match(r"^<\/(?:a|sup|sub|code|pre|script|style)\b", tok, flags=re.I) and prof_excl > 0:
                    prof_excl -= 1
                elif re.match(r"^<(?:br|p|div|li|tr|td|th)\b", tok, flags=re.I):
                    plain_chars.append(" ")
                    spans.append((s, e))
                    protegido.append(prof_excl > 0)
                continue

            esta_prot = prof_excl > 0
            if tok.startswith("&"):
                dec = limpiar_texto_unicode(tok)
                for ch in dec:
                    plain_chars.append(ch)
                    spans.append((s, e))
                    protegido.append(esta_prot)
            else:
                ch = " " if tok in "\u00A0\u2000\u2001\u2002\u2003\u2004\u2005\u2006\u2007\u2008\u2009\u200A\u202F\u205F\u3000" else tok
                plain_chars.append(ch)
                spans.append((s, e))
                protegido.append(esta_prot)

        return "".join(plain_chars), spans, protegido

    @staticmethod
    def _balancear_etiquetas_inline(fragmento: str) -> str:
        """Garantiza etiquetas inline balanceadas (em, i, strong, b) dentro de cada enlace."""
        tag_re = re.compile(r"<(/?)(em|i|strong|b)\b[^>]*>", flags=re.I)
        pila: list[str] = []
        prefijo = ""
        for m in tag_re.finditer(fragmento):
            es_cierre = bool(m.group(1))
            tag = m.group(2).lower()
            if not es_cierre:
                pila.append(tag)
            else:
                if pila and pila[-1] == tag:
                    pila.pop()
                elif tag in pila:
                    while pila and pila[-1] != tag:
                        pila.pop()
                    if pila:
                        pila.pop()
                else:
                    prefijo = f"<{tag}>" + prefijo
        sufijo = "".join(f"</{t}>" for t in reversed(pila))
        return prefijo + fragmento + sufijo

    @staticmethod
    def _normalizar_puntuacion_et_al(html_str: str) -> str:
        """Corrige comas duplicadas de Word alrededor de '<em>et al.</em>'."""
        return re.sub(
            r"(<(?:em|i)>\s*et\.?\s+al\.?)\s*,\s*(</(?:em|i)>)\s*,",
            r"\1\2,",
            html_str,
            flags=re.I,
        )

    def enlazar_fragmento(self, fragmento_html: str) -> str:
        """Enlaza todas las citas narrativas y parentéticas en un fragmento HTML."""
        if not self.entradas or not fragmento_html:
            return fragmento_html

        fragmento_html = re.sub(
            r'<a\b[^>]*class=["\'][^"\']*cita-enlace[^"\']*["\'][^>]*>(.*?)</a>',
            r"\1",
            fragmento_html,
            flags=re.I | re.S,
        )
        fragmento_html = self._normalizar_puntuacion_et_al(fragmento_html)

        re_word = rf"{LETRA_MAY}{LETRA_ALL}*(?:(?:{APOSTROFO}|{GUION}){LETRA_ALL}+)*"
        re_apellido_comp = rf"(?:{PARTICULA_SURNAME})?{re_word}(?:\s+(?:{PARTICULA_SURNAME})?{re_word}){{0,3}}"
        re_conector = r"\s+(?:y|e|and|&)\s+"
        re_sufijo_autor = (
            r"(?:(?:\s*,\s*|\s+)(?:et\.?\s+al\.?|y\s+(?:coautores?|colaboradores?|cols?\.|otros)|"
            r"e\s+(?:colaboradores?|cols?\.)|(?:and|&)\s+(?:colleagues|co-?authors|others)))?"
        )
        re_autor = (
            rf"{re_apellido_comp}"
            rf"(?:(?:\s*,\s*{re_apellido_comp})*{re_conector}{re_apellido_comp})?"
            rf"{re_sufijo_autor}"
            rf"(?:['’]s)?"
        )
        re_prefijo_paren = (
            r"(?:(?:e\.?\s*g\.?|i\.?\s*e\.?|por\s+ejemplo|véase|ver\s+también|ver|"
            r"cf\.?|cfr\.?|sensu|contra|según|entre\s+otros|citado\s+en|en)\s*,?\s*)*"
        )
        re_lista_anios = (
            rf"{self.RE_ANIO}{self.RE_PAG}"
            rf"(?:\s*,\s*(?:{self.RE_ANIO}|[a-z]\b){self.RE_PAG})*"
        )

        pat_narrativa = re.compile(
            rf"(?<!{LETRA_ALL})({re_autor})\s*\(({re_lista_anios})\)",
            flags=re.UNICODE,
        )
        pat_item_paren = re.compile(
            rf"({re_autor})[\s,]*({re_lista_anios})",
            flags=re.UNICODE,
        )
        pat_anio_item = re.compile(
            rf"({self.RE_ANIO}|(?<![A-Za-z])[a-z](?![A-Za-z]))({self.RE_PAG})",
            flags=re.I,
        )

        plain, spans, protegido = self._mapear_texto_plano_html(fragmento_html)
        if not plain or not re.search(self.RE_ANIO, plain, flags=re.I):
            return fragmento_html

        reemplazos: list[tuple[int, int, EntradaReferencia]] = []
        paren_spans: list[tuple[int, int]] = []

        # 1) Citas parentéticas: (Autor, 2020, 2026; Otro et al., 2019; Figura 1)
        for m_par in re.finditer(r"\(([^()]+)\)", plain):
            paren_spans.append(m_par.span())
            c_start, _ = m_par.span(1)
            contenido = m_par.group(1)
            if not re.search(self.RE_ANIO, contenido, flags=re.I):
                continue

            for mi in pat_item_paren.finditer(contenido):
                autor_raw = mi.group(1)
                autor_sin_pref = re.sub(rf"^{re_prefijo_paren}", "", autor_raw, flags=re.I).strip()
                offset_pref = len(autor_raw) - len(autor_sin_pref)
                autor_limpio, offset_stop = self._limpiar_palabras_iniciales_no_autor(autor_sin_pref)
                offset_total = offset_pref + offset_stop

                anios_bloque = mi.group(2)
                anios_start_rel = mi.start(2)
                anios_matches = list(pat_anio_item.finditer(anios_bloque))
                if not anios_matches:
                    continue

                anio_base_digits = ""
                for idx_a, m_ax in enumerate(anios_matches):
                    tok_anio = m_ax.group(1).strip()
                    m_base = re.match(r"^((?:19|20)\d{2})", tok_anio)
                    if m_base:
                        anio_base_digits = m_base.group(1)
                    elif len(tok_anio) == 1 and tok_anio.isalpha() and anio_base_digits:
                        tok_anio = f"{anio_base_digits}{tok_anio.lower()}"

                    ent = self._buscar_entrada(autor_limpio, tok_anio)
                    if ent:
                        if idx_a == 0:
                            s0 = c_start + mi.start(1) + offset_total
                            e0 = c_start + anios_start_rel + m_ax.end()
                            reemplazos.append((s0, e0, ent))
                        else:
                            sx = c_start + anios_start_rel + m_ax.start()
                            ex = c_start + anios_start_rel + m_ax.end()
                            reemplazos.append((sx, ex, ent))
                    else:
                        cita_perdida = f"{autor_limpio}, {tok_anio}"
                        if cita_perdida not in self._citas_no_encontradas:
                            self._citas_no_encontradas.append(cita_perdida)

        # 2) Citas narrativas: Maruca Sosa (1957: 176), Fibiger y Knüsel (2005), Waldron (1991, 2019)
        for mn in pat_narrativa.finditer(plain):
            ns, ne = mn.span()
            if any(not (ne <= ps or ns >= pe) for ps, pe in paren_spans):
                continue
            autor_raw = mn.group(1)
            autor_limpio, offset_stop = self._limpiar_palabras_iniciales_no_autor(autor_raw.strip())
            ns_adj = ns + (len(autor_raw) - len(autor_raw.lstrip())) + offset_stop

            anios_bloque = mn.group(2)
            anios_start_rel = mn.start(2)
            anios_matches = list(pat_anio_item.finditer(anios_bloque))
            if not anios_matches:
                continue

            anio_base_digits = ""
            for idx_a, m_ax in enumerate(anios_matches):
                tok_anio = m_ax.group(1).strip()
                m_base = re.match(r"^((?:19|20)\d{2})", tok_anio)
                if m_base:
                    anio_base_digits = m_base.group(1)
                elif len(tok_anio) == 1 and tok_anio.isalpha() and anio_base_digits:
                    tok_anio = f"{anio_base_digits}{tok_anio.lower()}"

                ent = self._buscar_entrada(autor_limpio, tok_anio)
                if ent:
                    es_ultimo = idx_a == len(anios_matches) - 1
                    if idx_a == 0:
                        e0 = (ns + anios_start_rel + m_ax.end() + 1) if es_ultimo else (ns + anios_start_rel + m_ax.end())
                        reemplazos.append((ns_adj, min(e0, ne), ent))
                    else:
                        sx = ns + anios_start_rel + m_ax.start()
                        ex = (ns + anios_start_rel + m_ax.end() + 1) if es_ultimo else (ns + anios_start_rel + m_ax.end())
                        reemplazos.append((sx, min(ex, ne), ent))
                else:
                    cita_perdida = f"{autor_limpio} ({tok_anio})"
                    if cita_perdida not in self._citas_no_encontradas:
                        self._citas_no_encontradas.append(cita_perdida)

        if not reemplazos:
            return fragmento_html

        reemplazos.sort(key=lambda x: (x[0], -x[1]))
        filtrados: list[tuple[int, int, EntradaReferencia]] = []
        ultimo_e = -1
        for s, e, ent in reemplazos:
            if s >= ultimo_e and s < e and e <= len(spans):
                if not any(protegido[k] for k in range(s, e)):
                    filtrados.append((s, e, ent))
                    ultimo_e = e

        partes_out: list[str] = []
        cursor_orig = 0
        for s_p, e_p, ent in filtrados:
            s_o = spans[s_p][0]
            e_o = spans[e_p - 1][1]
            if s_o < cursor_orig:
                continue
            self._contador_citas += 1
            cita_id = f"cita-{self._contador_citas}"
            ent.citas_ids.append(cita_id)

            partes_out.append(fragmento_html[cursor_orig:s_o])
            inner = self._balancear_etiquetas_inline(fragmento_html[s_o:e_o])
            titulo_tip = html.escape(f"Ver referencia: {ent.primer_apellido} ({ent.anio_visible})", quote=True)
            partes_out.append(
                f'<a id="{cita_id}" class="cita-enlace" href="#{ent.ref_id}" title="{titulo_tip}">{inner}</a>'
            )
            cursor_orig = e_o

        partes_out.append(fragmento_html[cursor_orig:])
        return "".join(partes_out)

    def renderizar_referencias_con_backlinks(self, referencias_html: list[str]) -> list[str]:
        """Devuelve cada referencia como '<p id="ref-..." class="ref-item">... ↩<sup>k</sup></p>'."""
        salida = []
        for ent, r_html in zip(self.entradas, referencias_html):
            r_limpia = self._limpiar_backlinks_previos(r_html)
            m_p = re.match(r"^<p\b[^>]*>(.*)</p>$", r_limpia, flags=re.I | re.S)
            inner = m_p.group(1).strip() if m_p else r_limpia

            backlinks = ""
            if ent.citas_ids:
                enlaces_ret = [
                    f'<a class="ref-backlink" href="#{cid}" title="Volver a la cita {k} en el texto">↩<sup>{k}</sup></a>'
                    for k, cid in enumerate(ent.citas_ids, start=1)
                ]
                backlinks = " " + " ".join(enlaces_ret)

            salida.append(f'<p id="{ent.ref_id}" class="ref-item">{inner}{backlinks}</p>')
        return salida

    def diagnosticar(self) -> tuple[list[str], list[str]]:
        """Devuelve (citas_sin_referencia_en_bibliografia, referencias_no_citadas_en_texto)."""
        no_citadas = [
            f"{e.primer_apellido} ({e.anio_visible})" if e.primer_apellido else e.texto_plano[:60]
            for e in self.entradas
            if not e.citas_ids
        ]
        return list(self._citas_no_encontradas), no_citadas


def procesar_html_completo(html_documento: str, estilo_id: str = "semantico") -> ResultadoCitacion:
    """
    Reprocesa un archivo HTML5 completo (con <main class='cuerpo-articulo'> y
    <section class='referencias'>), enlazando todas las citas y reconstruyendo los backlinks.
    """
    m_refs_sec = re.search(
        r'(<section\b[^>]*class=["\'][^"\']*referencias[^"\']*["\'][^>]*>)(.*?)(</section>)',
        html_documento,
        flags=re.I | re.S,
    )
    if not m_refs_sec:
        return ResultadoCitacion(html_documento, [], 0, [], [], [])

    refs_raw = re.findall(
        r'<p\b[^>]*class=["\'][^"\']*ref-item[^"\']*["\'][^>]*>.*?</p>',
        m_refs_sec.group(2),
        flags=re.I | re.S,
    )
    if not refs_raw:
        return ResultadoCitacion(html_documento, [], 0, [], [], [])

    motor = MotorCitacionesRMA(refs_raw, estilo_id=estilo_id)

    m_main = re.search(
        r'(<main\b[^>]*class=["\'][^"\']*cuerpo-articulo[^"\']*["\'][^>]*>)(.*?)(</main>)',
        html_documento,
        flags=re.I | re.S,
    )
    if not m_main:
        return ResultadoCitacion(html_documento, refs_raw, 0, [], [], motor.entradas)

    main_inner = m_main.group(2)
    m_refs_en_main = re.search(
        r'(<section\b[^>]*class=["\'][^"\']*referencias[^"\']*["\'][^>]*>)(.*?)(</section>)',
        main_inner,
        flags=re.I | re.S,
    )

    antes_refs = main_inner[: m_refs_en_main.start()] if m_refs_en_main else main_inner
    despues_refs = main_inner[m_refs_en_main.end():] if m_refs_en_main else ""

    def _enlazar_bloque(m_blk: re.Match) -> str:
        tag_open, contenido, tag_close = m_blk.group(1), m_blk.group(2), m_blk.group(3)
        if re.search(r'class=["\'][^"\']*fecha-manuscrito', tag_open, flags=re.I):
            return m_blk.group(0)
        return f"{tag_open}{motor.enlazar_fragmento(contenido)}{tag_close}"

    antes_refs_procesado = re.sub(
        r"(<(?:p|li|blockquote|td|th)\b[^>]*>)(.*?)(</(?:p|li|blockquote|td|th)>)",
        _enlazar_bloque,
        antes_refs,
        flags=re.I | re.S,
    )

    refs_finales = motor.renderizar_referencias_con_backlinks(refs_raw)
    refs_bloque_html = (
        f"{m_refs_sec.group(1)}\n      "
        + "\n      ".join(refs_finales)
        + f"\n    {m_refs_sec.group(3)}"
    )

    if m_refs_en_main:
        nuevo_main = f"{m_main.group(1)}{antes_refs_procesado}{refs_bloque_html}{despues_refs}{m_main.group(3)}"
        nuevo_html = html_documento[: m_main.start()] + nuevo_main + html_documento[m_main.end():]
    else:
        nuevo_html = html_documento[: m_main.start()] + f"{m_main.group(1)}{antes_refs_procesado}{m_main.group(3)}" + html_documento[m_main.end():]
        nuevo_html = re.sub(
            r'<section\b[^>]*class=["\'][^"\']*referencias[^"\']*["\'][^>]*>.*?</section>',
            lambda _: refs_bloque_html,
            nuevo_html,
            count=1,
            flags=re.I | re.S,
        )

    sin_ref, no_citadas = motor.diagnosticar()
    return ResultadoCitacion(
        cuerpo_html=nuevo_html,
        referencias_html=refs_finales,
        total_citas_enlazadas=motor._contador_citas,
        citas_sin_referencia=sin_ref,
        referencias_no_citadas=no_citadas,
        entradas=motor.entradas,
    )


if __name__ == "__main__":
    import sys
    from pathlib import Path

    if len(sys.argv) < 2:
        raise SystemExit("Uso: python rma_citaciones.py <archivo.html> [salida.html]")

    entrada_path = Path(sys.argv[1])
    salida_path = Path(sys.argv[2]) if len(sys.argv) >= 3 else entrada_path
    res = procesar_html_completo(entrada_path.read_text(encoding="utf-8"))
    salida_path.write_text(res.cuerpo_html, encoding="utf-8")

    print(f"Citas enlazadas: {res.total_citas_enlazadas}")
    if res.citas_sin_referencia:
        print("Citas en texto sin entrada en Referencias:", ", ".join(res.citas_sin_referencia))
    if res.referencias_no_citadas:
        print("Referencias no citadas en texto:", ", ".join(res.referencias_no_citadas))