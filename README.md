# Automatización de Maquetación — Revista del Museo de Antropología (RMA)

Este paquete automatiza la generación simultánea del **PDF maquetado** (en formato **tradicional a 2 columnas** para el número actual o **moderno a 1 columna** para el próximo año), la **galería HTML5 para OJS** y el **archivo XML JATS 1.1 / SciELO SPS 1.9 completo y verificado**.

---

## Instalación (una sola vez)

1. Ubicá esta carpeta en `/your_folder/automatizacion_rma_ojs` (o `~/automatizacion_rma_ojs`).
2. Abrí una terminal dentro de esa carpeta y ejecutá el instalador:
   ```bash
   cd /home/your_folder/automatizacion_rma_ojs
   bash instalar_linux_mac.sh
   ```
   *(En Windows: doble clic en `instalar_windows.bat`).*

---

## Preparación previa recomendada antes de compilar

1. **Manuscrito (`_manuscrito.docx`):**
   * Primeros párrafos en orden estricto: Título ES -> Título EN -> Autores -> Afiliaciones -> `Resumen` -> `Palabras clave:` -> `Abstract` -> `Keywords:` -> cuerpo -> `Bibliografía`.
   * Ubicación de recursos en párrafos aislados que digan únicamente `FIGURA 1`, `FIGURA 2`, `TABLA 1`, etc.
   * Subtítulos en **negrita** completa (sin punto final). Eliminar saltos de página manuales.
2. **Epígrafes en archivo Word separado (`{id}_epigrafes.docx`):**
   * Si tenés los epígrafes en un Word cuyo nombre empieza con el código de OJS (por ejemplo `54854_epigrafes.docx`, `54854_Epígrafes.docx` o `54854-epigrafes.docx`), copialo en la carpeta del artículo.
   * El sistema lo detecta automáticamente, conserva cursivas/negritas/superíndices y precarga los epígrafes bilingües de Figuras (`Figura 1.` / `Figure 1.`) y Tablas (`Tabla 1.` / `Table 1.` + `Nota:` / `Fuente:`).
3. **Figuras (`figuras/` o carpeta del artículo):**
   * Archivos sueltos (`.png`, `.jpg`, `.tif`/`.tiff`) a `300 dpi`, sin márgenes blancos excesivos.
   * Nombres estilo OJS (`54854_Figura+1.tiff`, `54854_Figura+2.jpg`) o estándar (`54854_fig01.png`) se detectan solos.
4. **Tablas (`_tablas.docx` o archivos individuales `.docx` / `.xlsx` / `.csv`):**
   * Usar tablas reales (no imágenes ni espacios). Procurar no superar las 10–12 columnas y usar encabezados breves.

---

## Ejecución

```bash
cd /home/andres/ANDRES/automatizacion_rma_ojs
.venv/bin/python asistente.py
```

Para recompilar directamente un artículo ya configurado:
```bash
.venv/bin/python maquetar_rma.py volumen19_n3_54854 54854
```

---

## Flujo flexible de texto y control de espacios en blanco (v11)

* **Relleno automático de vacíos en PDF:** Cuando una figura o tabla no entra al pie de una página y salta a la siguiente dejando un espacio en blanco grande, el compilador deja fluir automáticamente los párrafos siguientes por delante del recurso hasta completar las columnas de la página anterior (manteniendo el orden correlativo de figuras/tablas, sin separar subtítulos de su primer párrafo y conservando el orden original del manuscrito en el HTML de OJS y el XML de SciELO).
* **Ajuste elástico de altura:** Si una imagen no entra en la página por pocos milímetros, su altura máxima se ajusta suavemente dentro de un rango legible para encajar sin saltos innecesarios.
* **Anclaje puntual opcional:** Si necesitás que una figura o tabla específica no sea adelantada por el texto, agregá `flujo: "fijo"` en su entrada dentro de `{id}_recursos.yaml`.

---

## Equilibrio de columnas y separación de epígrafes (v12)

* **Equilibrio automático de dos columnas (`column-fill: balance`):** En el diseño tradicional a 2 columnas, los bloques de texto previos a una figura/tabla de ancho completo y el cierre final del artículo se distribuyen de manera pareja entre ambas columnas, evitando columnas vacías a mitad de página o al final del PDF.
* **Separación visual entre epígrafes y cuerpo de texto:** Se amplió el espaciado vertical (1,5–2 líneas) y se añadió un filete sutil de cierre al pie de figuras y tablas para distinguir con claridad los epígrafes bilingües del texto principal.

---

## Anclaje de figuras/tablas y encabezados pares/impares (v13)

* **Encabezados de páginas pares e impares (`@page :left` / `@page :right`):** Se implementaron cajas `position: running(...)` de ancho completo con filete inferior y respaldo automático de `autor_corto` / `titulo_corto`, además del pie de página con número de folio e ISSN en todas las páginas interiores.
* **Anclaje superior/inferior de figuras y tablas + flujo continuo hasta el cierre:** Las figuras y tablas se ubican preferentemente al tope o al pie de la caja de página, dejando fluir el texto (incluido el inicio de *Referencias bibliográficas* cuando un recurso queda al final del cuerpo) para eliminar vacíos en la anteúltima página.

---

## Tratamiento híbrido de tablas y avisos de continuación (v14)

* **Tablas cortas y medianas indivisibles:** Las tablas de hasta 22 filas (o menos del 78 % del alto útil de página) se comportan como bloques indivisibles (`break-inside: avoid`), evitando cortes incómodos a mitad de página y dejando fluir el texto posterior en el espacio libre.
* **Tablas extensas multipágina con avisos automáticos:** Cuando una tabla supera ese umbral y se distribuye en dos o más páginas, el compilador repite automáticamente el encabezado de columnas e inserta:
  * `Continúa en página X` al pie de cada fragmento intermedio.
  * `Continuación de Tabla N` en el encabezado de la página siguiente.
* **Control manual en `{id}_recursos.yaml`:** Podés agregar `dividir: false` (para forzar que una tabla no se corte) o `dividir: true` (para permitir que una tabla mediana se parta entre páginas) en su entrada dentro de `tablas:`.

---

## HTML autocontenido de archivo único para OJS (v15)

* **`html_autocontenido: true` en `revista.yaml`:** Por defecto, el archivo `{codigo}_{id}.html` incrusta la hoja de estilos (`rma_ojs.css`) dentro de `<style>` y convierte todas las figuras y logos institucionales a URI `data:image/...;base64,...`. Así podés subir **un único archivo `.html`** como galerada en OJS 3 sin cargar archivos dependientes uno por uno.
* Si en algún artículo muy pesado preferís mantener los archivos separados, cambiá `html_autocontenido: false` en `revista.yaml` o en el `{id}_metadatos.yaml` de ese envío.

---

## Identificadores ORCID y ROR con íconos hipervinculados (v16)

* **Detección y carga asistida en `{id}_metadatos.yaml`:** El asistente (`asistente.py`) extrae los autores y afiliaciones desde el manuscrito Word (rescatando cualquier URL de `orcid.org` o `ror.org` ya incluida) y solicita los identificadores faltantes para guardarlos en los bloques `autores:` y `afiliaciones:` de `{id}_metadatos.yaml`.
* **Íconos SVG clicables en PDF y HTML:** En lugar de imprimir URLs largas, el nombre de cada autor muestra el ícono verde **ORCID** hipervinculado y cada institución muestra el ícono **ROR** hipervinculado (embebidos en Base64 cuando `html_autocontenido: true`).
* **Etiquetas normalizadas en XML JATS 1.1:** El archivo `{codigo}_{id}.xml` incorpora `<contrib-id contrib-id-type="orcid">` y `<institution-id institution-id-type="ror">` compatibles con SciELO.

---

## Exportación XML JATS 1.1 / SciELO SPS 1.9 completa y verificada (v19)

* **Cuerpo completo (`<body>`) estructurado en secciones (`<sec>`):** Vuelca automáticamente todo el texto del artículo (`<p>`), citas en bloque (`<disp-quote>`), notas al pie (`<xref ref-type="fn">` + `<fn-group>`), figuras (`<fig>`) y tablas (`<table-wrap>`) con atributos `@sec-type` válidos en SciELO SPS 1.9 (`intro`, `materials|methods`, `results`, `discussion`, `conclusions`, `ack`).
* **Epígrafes JATS nativos y cero doble escape de entidades:** Convierte cursivas, negritas y super/subíndices a `<italic>`, `<bold>`, `<sup>` y `<sub>`, separa los epígrafes bilingües en `<p xml:lang="es">` y `<p xml:lang="en">` sin repetir `Figura N.` dentro de `<caption>` y evita el doble escape (`&amp;#x27;`) en palabras clave y citas.
* **Metadatos completos de `<article-meta>`:** Genera paginación sin ceros a la izquierda (`<fpage>`/`<lpage>`), afiliaciones desglosadas (`<institution content-type="original">`, `<institution content-type="orgname">`, `<institution-id institution-id-type="ror">`, `<country>` y `<email>`), `<pub-date>`, `<history>` editorial (`received`, `rev-recd`, `accepted`) y `<permissions>` Creative Commons.
* **Referencias granulares (`<mixed-citation>`) y verificación automática:** Desglosa autores, años, títulos, revistas, volúmenes, números, páginas y DOIs en etiquetas JATS y ejecuta una verificación automática de conformidad XML al finalizar cada compilación.

---

## Validación estricta SciELO Style Checker SPS 1.9 (v23)

* **`<abstract>` principal sin atributo `xml:lang` redundante:** El resumen en español hereda el idioma principal definido en `<article xml:lang="es">`, reservando `xml:lang="en"` exclusivamente para `<trans-abstract>`.
* **Epígrafe único `<caption>` bilingüe con `<title>` y limpieza de prefijos:** Cada `<fig>` y `<table-wrap>` emite un único `<caption>` conforme a `JATS-journalpublishing1.dtd` y SciELO SPS 1.9, con el epígrafe en español en `<title>` (eliminando prefijos reiterados como `Figura 1. Figura 1.`) y la traducción real al inglés como `<p xml:lang="en">` dentro del mismo `<caption>` (descartando textos marcador `Description of Figure/Table N.`).
* **Referencias duales `<mixed-citation>` + `<element-citation>`:** Cada `<ref>` incluye el texto completo formateado en `<mixed-citation>` junto con su desglose estructurado en `<element-citation publication-type="...">` sin texto ni puntuación suelta entre etiquetas, omitiendo `<label>` cuando la bibliografía no está numerada.
* **Identificador DOI garantizado en `<article-meta>`:** Se asegura la presencia de `<article-id pub-id-type="doi">` (generando un identificador provisional basado en el ID OJS si aún no se asignó sufijo definitivo, con aviso en consola).
