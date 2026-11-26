"""
Genera la presentación del proyecto (Parte 1: merge y limpieza) con formato sencillo:
diseños estándar de PowerPoint (Título, Título y contenido, Dos contenidos, Comparación),
fondo blanco y gráficas nativas. Todo se puede editar directamente en PowerPoint.

Uso:  python presentacion/generar_presentacion.py
Salida: presentacion/Presentacion_Dengue.pptx

Para agregar secciones (EDA, modelo): escribe funciones slide_* como las de abajo
y agrégalas a SECCIONES al final del archivo.

Formato de las viñetas: cada elemento es un texto o (texto, nivel).
El texto entre **dobles asteriscos** sale en negritas.
"""
from pathlib import Path
import base64
import io
import json
import uuid

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION, XL_LEGEND_POSITION
from pptx.oxml.ns import qn
from lxml import etree

SALIDA = Path(__file__).with_name("Presentacion_Dengue.pptx")
RAIZ = Path(__file__).resolve().parent.parent


def imagen_notebook(notebook, texto_en_codigo):
    """Regresa (como archivo en memoria) la primera imagen de la celda cuyo código contiene el texto."""
    nb = json.loads((RAIZ / notebook).read_text(encoding="utf-8"))
    for celda in nb["cells"]:
        if celda["cell_type"] == "code" and texto_en_codigo in "".join(celda["source"]):
            for salida in celda.get("outputs", []):
                png = salida.get("data", {}).get("image/png")
                if png:
                    return io.BytesIO(base64.b64decode("".join(png) if isinstance(png, list) else png))
    raise ValueError(f"No se encontró la imagen '{texto_en_codigo}' en {notebook}")

# Diseños de la plantilla por defecto
L_TITULO, L_CONTENIDO, L_SECCION, L_DOS, L_COMPARACION, L_SOLO_TITULO = 0, 1, 2, 3, 4, 5

TAM_TITULO = 36
TAM_TEXTO = 20
TAM_SUB = 18


# ───────────────────────── Helpers ─────────────────────────
def a_16x9(prs):
    """Cambia a 16:9 y ensancha los placeholders del patrón y los diseños."""
    k = 13.333 / 10
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    for contenedor in [prs.slide_master] + list(prs.slide_layouts):
        for sh in contenedor.shapes:
            if sh._element.find(".//" + qn("a:xfrm")) is not None:
                sh.left, sh.width = int(sh.left * k), int(sh.width * k)


def _runs(p, texto, size):
    for i, trozo in enumerate(texto.split("**")):
        if not trozo:
            continue
        r = p.add_run()
        r.text = trozo
        r.font.size = Pt(size)
        r.font.bold = i % 2 == 1


def vinetas(ph, items, size=TAM_TEXTO, columnas=1):
    tf = ph.text_frame
    tf.clear()
    for i, it in enumerate(items):
        texto, nivel = (it, 0) if isinstance(it, str) else it
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.level = nivel
        p.space_after = Pt(6)
        _runs(p, texto, size if nivel == 0 else TAM_SUB)
    if columnas > 1:
        tf._txBody.find(qn("a:bodyPr")).set("numCol", str(columnas))


def titulo(s, txt):
    s.shapes.title.text = txt
    for p in s.shapes.title.text_frame.paragraphs:
        for r in p.runs:
            r.font.size = Pt(TAM_TITULO)


def ph(s, idx):
    return s.placeholders[idx]


def _quitar(s, idx):
    """Quita el placeholder idx y regresa su posición (x, y, w, h)."""
    lugar = ph(s, idx)
    pos = lugar.left, lugar.top, lugar.width, lugar.height
    lugar._element.getparent().remove(lugar._element)
    return pos


def a_tabla(s, filas, x, y, w, anchos=None, size=16, alto_fila=0.4):
    """Tabla simple; la primera fila es el encabezado (en negritas)."""
    tabla = s.shapes.add_table(len(filas), len(filas[0]), x, y, w, Inches(alto_fila * len(filas))).table
    for fila in tabla.rows:
        fila.height = Inches(alto_fila)
    if anchos:
        total = sum(anchos)
        for c, a in enumerate(anchos):
            tabla.columns[c].width = int(w * a / total)
    for r, fila in enumerate(filas):
        for c, valor in enumerate(fila):
            cel = tabla.cell(r, c)
            cel.text = str(valor)
            for p in cel.text_frame.paragraphs:
                for run in p.runs:
                    run.font.size = Pt(size)
                    run.font.bold = r == 0
    return tabla


def a_grafica(s, idx, categorias, valores, formato="0.0%", horizontal=False, nombre="Serie"):
    """Reemplaza el placeholder idx por una gráfica de barras nativa en la misma posición.
    valores: lista (una serie) o dict {nombre_serie: lista} (varias series, con leyenda)."""
    x, y, w, h = _quitar(s, idx)
    cd = CategoryChartData()
    cd.categories = categorias
    series = valores if isinstance(valores, dict) else {nombre: valores}
    for n, v in series.items():
        cd.add_series(n, v)
    tipo = XL_CHART_TYPE.BAR_CLUSTERED if horizontal else XL_CHART_TYPE.COLUMN_CLUSTERED
    ch = s.shapes.add_chart(tipo, x, y, w, h, cd).chart
    ch.has_legend = len(series) > 1
    if ch.has_legend:
        ch.legend.position = XL_LEGEND_POSITION.BOTTOM
        ch.legend.include_in_layout = False
    ch.font.size = Pt(14)
    plot = ch.plots[0]
    plot.has_data_labels = True
    plot.data_labels.number_format = formato
    plot.data_labels.number_format_is_linked = False
    plot.data_labels.position = XL_LABEL_POSITION.OUTSIDE_END
    ch.value_axis.visible = False
    ch.value_axis.has_major_gridlines = False
    if horizontal:
        ch.category_axis.reverse_order = True
    return ch


def a_grafica_lineas(s, idx, categorias, series, formato="#,##0", minimo=None):
    """Reemplaza el placeholder idx por una gráfica de líneas con marcadores (varias series)."""
    x, y, w, h = _quitar(s, idx)
    cd = CategoryChartData()
    cd.categories = categorias
    for n, v in series.items():
        cd.add_series(n, v)
    ch = s.shapes.add_chart(XL_CHART_TYPE.LINE_MARKERS, x, y, w, h, cd).chart
    ch.font.size = Pt(14)
    ch.has_legend = True
    ch.legend.position = XL_LEGEND_POSITION.BOTTOM
    ch.legend.include_in_layout = False
    va = ch.value_axis
    va.tick_labels.number_format = formato
    va.tick_labels.number_format_is_linked = False
    if minimo is not None:
        va.minimum_scale = minimo
    return ch


def imagen_en(s, img, x, y, w, h):
    """Inserta una imagen dentro del recuadro (x, y, w, h), conservando proporción y centrada."""
    pic = s.shapes.add_picture(img, x, y, width=w)
    if pic.height > h:
        pic.width = int(pic.width * h / pic.height)
        pic.height = h
    pic.left = x + (w - pic.width) // 2
    return pic


def imagen_y_texto(prs, txt_titulo, notebook, clave, items, notas=None, ancho=10, size=16):
    """Imagen del notebook arriba (centrada) y viñetas debajo."""
    s = nueva(prs, L_CONTENIDO, txt_titulo, notas)
    pic = s.shapes.add_picture(imagen_notebook(notebook, clave), 0, Inches(1.5), width=Inches(ancho))
    pic.left = (prs.slide_width - pic.width) // 2
    cuerpo = ph(s, 1)
    cuerpo.top = pic.top + pic.height + Inches(0.2)
    cuerpo.height = Inches(7.1) - cuerpo.top
    vinetas(cuerpo, items, size)
    return s


def nueva(prs, layout, txt_titulo, notas=None):
    s = prs.slides.add_slide(prs.slide_layouts[layout])
    titulo(s, txt_titulo)
    if notas:
        s.notes_slide.notes_text_frame.text = notas
    return s


def contenido(prs, txt_titulo, items, notas=None, size=TAM_TEXTO, columnas=1):
    s = nueva(prs, L_CONTENIDO, txt_titulo, notas)
    vinetas(ph(s, 1), items, size, columnas)
    return s


def seccion(prs, txt_titulo, subtitulo, notas=None):
    s = nueva(prs, L_SECCION, txt_titulo, notas)
    ph(s, 1).text = subtitulo
    return s


def texto_y_grafica(prs, txt_titulo, items, categorias, valores, formato, notas=None,
                    grafica_izq=False, horizontal=False, size=TAM_TEXTO):
    s = nueva(prs, L_DOS, txt_titulo, notas)
    txt_idx, graf_idx = (2, 1) if grafica_izq else (1, 2)
    vinetas(ph(s, txt_idx), items, size)
    a_grafica(s, graf_idx, categorias, valores, formato, horizontal)
    return s


def comparacion(prs, txt_titulo, enc_izq, items_izq, enc_der, items_der, notas=None,
                size=TAM_TEXTO, columnas=1):
    s = nueva(prs, L_COMPARACION, txt_titulo, notas)
    ph(s, 1).text = enc_izq
    vinetas(ph(s, 2), items_izq, size, columnas)
    ph(s, 3).text = enc_der
    vinetas(ph(s, 4), items_der, size, columnas)
    return s


# ───────────────────────── Introducción ─────────────────────────
def slide_portada(prs):
    s = nueva(prs, L_TITULO, "Factores asociados a la defunción por dengue en México",
              "Proyecto final: modelo de regresión logística para identificar qué factores se asocian "
              "a la muerte en pacientes con dengue confirmado. Esta primera parte explica cómo "
              "construimos y limpiamos la base de datos.")
    ph(s, 1).text = ("Regresión logística binaria con datos abiertos de la DGE (SINAVE), 2020–2026\n"
                     "Análisis Estadístico Multivariado · ITESO\n"
                     "Construcción de la base, análisis exploratorio y modelo")


def slide_ruta(prs):
    contenido(prs, "Ruta del proyecto", [
        "**01 · Merge de datos:** unir 294 fotos semanales en una base con un registro por paciente (completado)",
        "**02 · Limpieza:** filtrar casos confirmados, seleccionar variables y recodificar a 1/0 (completado)",
        "**03 · Análisis exploratorio:** describir la mortalidad según cada variable (completado)",
        "**04 · Modelo:** estimar la regresión logística e interpretar los coeficientes (completado)",
    ], "El trabajo está dividido en cuatro notebooks.")


def slide_pregunta(prs):
    contenido(prs, "¿Qué factores se asocian a morir por dengue?", [
        "**Variable objetivo:** DEFUNCION (1 = murió, 0 = no murió)",
        "**Población:** casos confirmados de dengue en México, inicio de síntomas 2020–2026",
        "**Variables independientes:**",
        ("Demográficas: sexo, edad, indígena", 1),
        ("Comorbilidades: diabetes, hipertensión, enfermedad renal, inmunosupresión, cirrosis hepática, embarazo", 1),
        ("Gravedad clínica: manifestaciones hemorrágicas, tipo de paciente (hospitalizado)", 1),
        ("Laboratorio: resultado de PCR (serotipo)", 1),
        ("Tiempo y lugar: mes de inicio de síntomas, entidad de asignación (costera o no)", 1),
    ], "Las variables se eligieron al inicio del proyecto (archivo VARIABLES.rtf).")


# ───────────────────────── 01 · Merge ─────────────────────────
def slide_div_merge(prs):
    seccion(prs, "01 · Merge de datos", "De 294 fotos semanales a una base con un registro por paciente",
            "Primera etapa: notebook 01_merge_datos.ipynb.")


def slide_fuente(prs):
    texto_y_grafica(prs, "Fuente de los datos", [
        "**Datos Abiertos de la DGE:** bases históricas de enfermedades transmitidas por vector (dic-2020 a nov-2026)",
        "Cada archivo es una **foto acumulada** del año: trae todos los casos hasta esa fecha",
        "Los registros **se corrigen** de una foto a otra (ESTATUS_CASO, DEFUNCION, etc.)",
        "Las 294 fotos tienen las mismas 28 columnas",
    ], ["2020", "2021", "2022", "2023", "2024", "2025", "2026"], [4, 50, 53, 51, 50, 48, 38], "0",
        "Gráfica: número de archivos por año. 2020 solo incluye diciembre; 2026 llega al 5 de noviembre. "
        "Como cada foto es acumulada, el mismo paciente aparece muchas veces y no basta con concatenar.",
        size=18)


def slide_estrategia(prs):
    contenido(prs, "Estrategia del merge", [
        "**1. Ordenar** las fotos cronológicamente según la fecha del nombre del archivo",
        "**2. Leer** solo el CSV válido de cada zip, todo como texto para no perder información",
        "**3. Deduplicar:** concatenar foto por foto y quedarse con la última versión de cada ID_REGISTRO",
        "**4. Bitácora:** guardar en qué fotos aparece cada registro y su valor de DEFUNCION",
        "**Resultado:** 19,876,426 filas leídas → 1,276,897 registros únicos",
        ("Se hace foto por foto porque unir ~20 millones de filas de una vez satura la memoria", 1),
        ("Mismo criterio que el repositorio DengueMX; 148 IDs repetidos dentro de una misma foto", 1),
    ])


def slide_trazabilidad(prs):
    contenido(prs, "Metadatos de trazabilidad", [
        "**N_FOTOS:** en cuántas fotos semanales apareció el registro",
        "**PRIMERA_FOTO / ULTIMA_FOTO:** fecha de la primera y la última foto donde aparece",
        "**DEFUNCION_CAMBIO:** si DEFUNCION cambió en algún momento (2,494 registros)",
        "**EN_FOTO_FINAL:** si el registro sigue en la última foto de su año (22,629 dados de baja)",
    ], "DEFUNCION_CAMBIO y EN_FOTO_FINAL se usan en la limpieza para decidir qué registros quitar.")


def slide_resultado_merge(prs):
    texto_y_grafica(prs, "Resultado del merge", [
        "Fechas en dos formatos se unifican; 9999-99-99 se vuelve fecha vacía",
        "Las claves de catálogo pasan a enteros (sin recodificar)",
        "Se agrega ANIO_SINTOMAS",
        "**1,276,897 registros × 36 columnas**",
        "dengue_merge.csv.gz: 15.4 MB",
    ], ["2020", "2021", "2022", "2023", "2024", "2025", "2026"],
        [120671, 37395, 61123, 281916, 555685, 146444, 73663], "#,##0",
        "Gráfica: registros por año de inicio de síntomas. Aquí todavía hay casos probables y descartados.",
        size=18)


# ───────────────────────── 02 · Limpieza ─────────────────────────
def slide_div_limpieza(prs):
    seccion(prs, "02 · Limpieza", "De 1.28 millones de registros a 251,900 casos confirmados",
            "Segunda etapa: notebook 02_limpieza.ipynb.")


def slide_embudo(prs):
    s = nueva(prs, L_SOLO_TITULO, "Filtros de la limpieza",
              "El filtro más fuerte es quedarse solo con confirmados: casi todo lo que se quita son "
              "casos probables (719,278) y descartados (283,089).")
    filas = [
        ("Paso", "Se quitan", "Quedan", "Muertes"),
        ("0 · Base del merge", "–", "1,276,897", "5,296"),
        ("1 · Sigue en la foto final", "22,629 (1.77%)", "1,254,268", "5,207"),
        ("2 · Solo confirmados (ESTATUS_CASO = 2)", "1,002,367 (79.92%)", "251,901", "1,763"),
        ("3 · Entidad de asignación especificada", "1", "251,900", "1,763"),
        ("Total", "1,024,997 (80.27%)", "251,900", "1,763 (0.70%)"),
    ]
    tabla = s.shapes.add_table(len(filas), 4, Inches(0.67), Inches(1.9), Inches(12), Inches(3.6)).table
    for c, w in enumerate([5.4, 2.6, 2.0, 2.0]):
        tabla.columns[c].width = Inches(w)
    for r, fila in enumerate(filas):
        for c, valor in enumerate(fila):
            cel = tabla.cell(r, c)
            cel.text = valor
            for p in cel.text_frame.paragraphs:
                for run in p.runs:
                    run.font.size = Pt(18)
                    run.font.bold = r in (0, len(filas) - 1)


def slide_decision(prs):
    comparacion(prs, "Decisión: se conservan los cambios en DEFUNCION",
                "Opción elegida: conservarlos", [
                    "251,900 casos confirmados",
                    "1,763 muertes (0.70%)",
                    "El modelo ve todas las muertes, reportadas rápido o tarde",
                ],
                "Alternativa descartada: quitarlos", [
                    "251,077 casos confirmados",
                    "962 muertes",
                    "Se pierden 801 de 1,770 muertes confirmadas (45%) → sesgo de selección",
                ],
                "2,494 registros cambiaron DEFUNCION; 2,433 pasaron de No a Sí porque la muerte se "
                "registró después. Su valor final es el correcto. Por eso solo filtramos por foto final.")


def slide_variables(prs):
    comparacion(prs, "Selección de variables: 17 de 35 columnas",
                "Se conservan (17)",
                ["ID_REGISTRO", "DEFUNCION", "SEXO", "EDAD_ANOS", "FECHA_SIGN_SINTOMAS",
                 "ANIO_SINTOMAS", "ENTIDAD_ASIG", "INDIGENA", "HEMORRAGICOS",
                 "DIABETES", "HIPERTENSION", "ENFERMEDAD_RENAL", "INMUNOSUPR",
                 "CIRROSIS_HEPATICA", "EMBARAZO", "TIPO_PACIENTE", "RESULTADO_PCR"],
                "No entran al modelo (18)",
                ["DICTAMEN", "ENFERMEDAD_ULC_PEPTICA", "HABLA_LENGUA_INDIG", "TOMA_MUESTRA",
                 "ENTIDAD_RES", "MUNICIPIO_RES", "ENTIDAD_UM_NOTIF", "MUNICIPIO_UM_NOTIF",
                 "MUNICIPIO_ASIG", "INSTITUCION_UM_NOTIF", "FECHA_ACTUALIZACION", "ARCHIVO_FUENTE",
                 "FECHA_ARCHIVO", "N_FOTOS", "PRIMERA_FOTO", "ULTIMA_FOTO",
                 "DEFUNCION_CAMBIO", "EN_FOTO_FINAL"],
                "DICTAMEN se excluye por fuga de información: solo se asigna a quienes murieron "
                "(250,135 de 250,138 vivos tienen 'No aplica'). Tampoco se calcula el conteo de "
                "muertos por año, porque se derivaría del mismo DEFUNCION.",
                size=13, columnas=2)


def slide_recodificacion(prs):
    texto_y_grafica(prs, "Recodificación a 1/0", [
        "**Catálogo Sí/No:** 1 = Sí, 2 = No → 1 = Sí, 0 = No",
        "**SEXO → HOMBRE:** 1 = hombre (44.8%)",
        "**TIPO_PACIENTE → HOSPITALIZADO:** 1 = hospitalizado (38.9%); 9 vacíos → ambulatorio",
        "**ENTIDAD_ASIG → COSTERA:** 1 = una de las 17 entidades con costa (70.1%)",
        "No se quita ninguna fila",
    ], ["Indígena", "Diabetes", "Embarazo", "Hipertensión", "Hemorrágicos",
        "Enf. renal", "Inmunosupresión", "Cirrosis"],
        [0.03314, 0.02879, 0.02621, 0.02349, 0.00908, 0.00251, 0.00147, 0.00093], "0.00%",
        "Gráfica: proporción de pacientes con cada condición. Las comorbilidades son poco frecuentes; "
        "la más común, ser indígena, es apenas 3.3%. Embarazo vale 0 en todos los hombres.",
        horizontal=True, size=18)


def slide_serotipo(prs):
    texto_y_grafica(prs, "Serotipo (RESULTADO_PCR) → variables dummy", [
        "5 categorías → 4 dummies: DENV2, DENV3, DENV4, SIN_SEROTIPO",
        "**DENV1 es la referencia** (menor mortalidad: 0.29%)",
        "Incluir las 5 columnas causaría colinealidad con el intercepto",
        "**Sin serotipo** (56.4% de los confirmados) es un serotipo desconocido, no un quinto virus",
    ], ["DENV1", "DENV2", "DENV3", "DENV4", "Sin serotipo"],
        [0.0029, 0.0066, 0.0067, 0.0086, 0.0076], "0.00%",
        "Gráfica: mortalidad por serotipo. Casos: DENV1 13,531; DENV2 23,139; DENV3 71,812; DENV4 1,396; "
        "sin serotipo 142,022. Sin serotipo se debe a confirmación por NS1 o IgM, muestra tomada tarde "
        "o confirmación por asociación epidemiológica.",
        grafica_izq=True, size=18)


def slide_mes(prs):
    texto_y_grafica(prs, "Mes de inicio de síntomas → variables dummy", [
        "11 dummies; **octubre es la referencia** (mes con más casos, 21.7%)",
        "Lluvias (jun–oct) concentran el 72% de los casos",
        "Mortalidad: 0.68% en lluvias vs 0.76% en secas",
        "Hospitalización: ~48% en enero–marzo",
    ], ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"],
        [0.036, 0.022, 0.020, 0.020, 0.028, 0.052, 0.080, 0.158, 0.211, 0.217, 0.133, 0.024], "0%",
        "Gráfica: porcentaje de casos confirmados por mes. En secas se estudian sobre todo los casos "
        "graves; el mes capta ese cambio en el perfil de los casos.",
        grafica_izq=True, size=18)


def slide_anio(prs):
    texto_y_grafica(prs, "El año se excluye del modelo", [
        "La mortalidad oscila entre 0.58% y 0.86% sin tendencia clara",
        "Los avances médicos entre años no implican necesariamente menos muertes",
        "**Limitación:** el serotipo dominante cambió (DENV2 en 2020–2021, DENV3 desde 2023); "
        "sin el año, el serotipo absorbe parte del efecto de la época",
    ], ["2020", "2021", "2022", "2023", "2024", "2025", "2026"],
        [0.0058, 0.0076, 0.0066, 0.0073, 0.0069, 0.0086, 0.0058], "0.00%",
        "Gráfica: mortalidad de casos confirmados por año. Confirmados: 2020 24,308; 2021 6,553; "
        "2022 12,327; 2023 54,197; 2024 123,141; 2025 21,918; 2026 9,456.",
        grafica_izq=True, size=18)


def slide_base_final(prs):
    contenido(prs, "Base final", [
        "**251,900 pacientes × 29 columnas**, sin valores faltantes (dengue_limpio.csv)",
        "**1,763 muertes (0.70%)**",
        "Columnas:",
        ("ID_REGISTRO y DEFUNCION (objetivo)", 1),
        ("EDAD_ANOS: media 28.5, mediana 25, rango 0–112", 1),
        ("11 binarias: HOMBRE, COSTERA, INDIGENA, HEMORRAGICOS, comorbilidades, HOSPITALIZADO", 1),
        ("4 de serotipo (ref. DENV1) y 11 de mes (ref. octubre)", 1),
        "**Para el modelo:** clases muy desbalanceadas; la exactitud por sí sola engaña",
    ], "Con 0.7% de muertes, un modelo que dijera 'nadie muere' tendría 99.3% de exactitud.")


# ───────────────────────── 03 · EDA ─────────────────────────
def slide_div_eda(prs):
    seccion(prs, "03 · Análisis exploratorio", "Qué variables se asocian a la muerte antes de ajustar el modelo",
            "Tercera etapa: notebook 03_EDA.ipynb.")


def slide_desbalance(prs):
    contenido(prs, "Variable objetivo: clases muy desbalanceadas", [
        "**250,137 sobrevivieron y 1,763 fallecieron:** mortalidad de 0.70% (1 de cada 142)",
        "Consecuencias para el modelo:",
        ("La exactitud engaña: predecir \"nadie fallece\" acierta más del 99%", 1),
        ("Con umbral de 0.5 casi no se detectaría a ningún fallecido", 1),
        ("Hay que revisar recall, precision, F1 y la curva ROC", 1),
        ("Estratificar por DEFUNCION al dividir en entrenamiento y prueba", 1),
    ], "La base tiene 251,900 casos confirmados, sin faltantes ni IDs duplicados.")


def slide_edad(prs):
    s = nueva(prs, L_SOLO_TITULO, "Edad: los fallecidos son 11 años mayores en promedio",
              "La distribución general tiene asimetría positiva (cola larga de adultos mayores). "
              "En los fallecidos es casi simétrica y aplanada: hay muertes en todas las edades, "
              "con mucho más peso en adultos mayores.")
    filas = [
        ("Medida", "Todos", "Sobrevivió (0)", "Falleció (1)"),
        ("Media", "28.5", "28.4", "39.4"),
        ("Mediana", "25", "25", "37"),
        ("Moda", "13", "13", "10"),
        ("Desviación estándar", "17.8", "17.7", "25.1"),
        ("Q1 – Q3", "14 – 40", "14 – 40", "16 – 59.5"),
        ("Rango intercuartílico", "26", "26", "43.5"),
        ("Asimetría", "0.79", "0.79", "0.25"),
        ("Curtosis", "0.03", "0.03", "-1.09"),
    ]
    a_tabla(s, filas, Inches(0.67), Inches(1.8), Inches(12), anchos=[4, 2.6, 2.7, 2.7], size=18)


def slide_grupo_edad(prs):
    texto_y_grafica(prs, "Mortalidad por grupo de edad: forma de \"U\"", [
        "Alta en menores de 5 años (1.17%)",
        "Mínima entre 15 y 19 años (0.33%)",
        "Crece desde los 40 años hasta **7.09% en mayores de 80**",
        "El log-odds no es una línea recta → **la relación con la edad no es lineal**",
        "En el modelo se probará un término **edad²** (comparando AIC y BIC)",
    ], ["0-4", "5-9", "10-14", "15-19", "20-29", "30-39", "40-49", "50-59", "60-69", "70-79", "80+"],
        [0.0117, 0.0067, 0.0047, 0.0033, 0.0042, 0.0053, 0.0072, 0.0100, 0.0157, 0.0339, 0.0709], "0.0%",
        "La regresión logística supone una relación lineal entre cada variable y el log-odds. "
        "Log-odds por grupo: de -4.43 (0-4) baja a -5.71 (15-19) y sube a -2.57 (80+).",
        grafica_izq=True, size=18)


def slide_binarias_eda(prs):
    texto_y_grafica(prs, "Las comorbilidades son mucho más frecuentes en los fallecidos", [
        "**Enfermedad renal:** 0.25% de todos los casos, 4.4% de los fallecidos (17 veces más)",
        "Cirrosis 7.4×, inmunosupresión 6.2×, hemorrágicos 5.6×, diabetes 5.2×, hipertensión 4.7×",
        "**96% de los fallecidos estuvo hospitalizado** (39% del total)",
        "Sexo, indígena, embarazo, costera y serotipo: proporciones parecidas en ambos grupos",
    ], ["Diabetes", "Hipertensión", "Hemorrágicos", "Enf. renal", "Indígena", "Embarazo",
        "Inmunosupresión", "Cirrosis"],
        {"Todos los casos": [0.0288, 0.0235, 0.0091, 0.0025, 0.0331, 0.0262, 0.0015, 0.0009],
         "Fallecidos": [0.1492, 0.1100, 0.0505, 0.0437, 0.0391, 0.0289, 0.0091, 0.0068]}, "0.0%",
        "Gráfica: porcentaje de pacientes con cada condición. Las variables más raras (cirrosis 233 casos, "
        "inmunosupresión 371, DENV4 1,396) tienen pocos fallecidos: sus coeficientes pueden ser poco precisos.",
        horizontal=True, size=18)


def slide_or_crudo(prs):
    s = nueva(prs, L_DOS, "Odds ratio crudo de cada condición",
              "El OR crudo compara los odds de fallecer con y sin la condición, sin ajustar por las demás "
              "variables. Puede cambiar en el modelo multivariado, porque las variables están relacionadas "
              "(por ejemplo, la diabetes es más común en personas mayores).")
    vinetas(ph(s, 1), [
        "**Hospitalizado (OR ≈ 36):** 1.72% de mortalidad contra 0.05% en ambulatorios",
        "**Enfermedad renal (OR ≈ 20.5):** 12.2% de mortalidad, la más alta",
        "Cirrosis, inmunosupresión, diabetes, hemorrágicos e hipertensión: OR entre 5 y 8",
        "Costera, serotipo, indígena, embarazo y sexo: OR cercano a 1 (asociación débil)",
    ], size=18)
    x, y, w, h = _quitar(s, 2)
    a_tabla(s, [
        ("Variable", "Mort. con", "Mort. sin", "OR"),
        ("Hospitalizado", "1.72%", "0.05%", "36.45"),
        ("Enf. renal", "12.16%", "0.67%", "20.50"),
        ("Cirrosis", "5.15%", "0.70%", "7.75"),
        ("Inmunosupresión", "4.31%", "0.69%", "6.44"),
        ("Diabetes", "3.63%", "0.61%", "6.10"),
        ("Hemorrágicos", "3.89%", "0.67%", "6.00"),
        ("Hipertensión", "3.28%", "0.64%", "5.28"),
        ("Costera", "0.76%", "0.56%", "1.36"),
        ("DENV4", "0.86%", "0.70%", "1.23"),
        ("Sin serotipo", "0.76%", "0.62%", "1.23"),
        ("Indígena", "0.83%", "0.70%", "1.19"),
        ("Embarazo", "0.77%", "0.70%", "1.11"),
        ("Hombre", "0.72%", "0.68%", "1.06"),
        ("DENV2", "0.66%", "0.70%", "0.94"),
        ("DENV3", "0.67%", "0.71%", "0.94"),
    ], x, y - Inches(0.15), w, anchos=[3, 2, 2, 1.5], size=12, alto_fila=0.3)


def slide_serotipo_mes_eda(prs):
    contenido(prs, "Serotipo y mes: asociaciones débiles", [
        "**Serotipo:**",
        ("56% de los casos no tiene serotipo identificado; DENV3 es el más común (28.5%)", 1),
        ("DENV1 tiene la menor mortalidad (0.29%); los demás entre 0.66% y 0.86%", 1),
        "**Mes de inicio de síntomas:**",
        ("Los casos se concentran entre agosto y noviembre, con el máximo en octubre", 1),
        ("Mortalidad algo mayor fuera de temporada (0.76%–1.01% en enero–junio) y ≈0.65% en los meses pico", 1),
        ("En temporada alta se confirman muchos casos leves", 1),
        "En ambos casos las diferencias son pequeñas en términos absolutos",
    ], "Las gráficas de serotipo y mes ya se mostraron en la sección de limpieza.")


def slide_correlacion(prs):
    s = nueva(prs, L_DOS, "Correlación: sin problemas graves de multicolinealidad",
              "Pearson y Spearman dan mapas casi idénticos porque casi todas las variables son binarias; "
              "solo cambia la edad. Para dos variables 0/1, Pearson equivale al coeficiente phi.")
    vinetas(ph(s, 1), [
        "**Con DEFUNCION todas las correlaciones son bajas** (la mayor es 0.10, con HOSPITALIZADO)",
        ("Es esperado con una variable objetivo tan desbalanceada", 1),
        ("El orden coincide: hospitalización, enf. renal, diabetes, edad, hipertensión, hemorrágicos", 1),
        "**Entre predictores** la correlación más alta es entre dummies de la misma variable (serotipo)",
        "Diabetes, hipertensión y edad: correlación moderada con sentido clínico",
    ], size=18)
    x, y, w, h = _quitar(s, 2)
    a_tabla(s, [
        ("Par de predictores", "Correlación"),
        ("DENV3 – SIN_SEROTIPO", "-0.718"),
        ("DIABETES – HIPERTENSION", "0.367"),
        ("DENV2 – SIN_SEROTIPO", "-0.362"),
        ("EDAD_ANOS – DIABETES", "0.254"),
        ("EDAD_ANOS – HIPERTENSION", "0.243"),
        ("HEMORRAGICOS – DIABETES", "0.233"),
    ], x, y, w, anchos=[3, 1.3], size=16)


def slide_distribuciones(prs):
    texto_y_grafica(prs, "Edad: ajuste de 11 distribuciones (MLE + Kolmogorov-Smirnov)", [
        "Parámetros estimados por máxima verosimilitud; menor estadístico KS = mejor ajuste",
        "**La mejor es la gamma** (KS = 0.043), seguida de chi-cuadrada, beta y Weibull",
        "**Todos los p-values son 0:** se rechazan las 11 distribuciones",
        ("Con más de 250 mil datos, KS detecta diferencias muy pequeñas", 1),
        ("La edad está en años enteros; las distribuciones son continuas", 1),
    ], ["Gamma", "Chi-cuadrada", "Beta", "Weibull", "Logística", "Normal", "Triangular",
        "Exponencial", "Poisson", "Uniforme", "Log-normal"],
        [0.0430, 0.0430, 0.0452, 0.0467, 0.0779, 0.0950, 0.1682, 0.1840, 0.3484, 0.4194, 0.5060], "0.000",
        "Gráfica: estadístico KS de cada distribución (menor es mejor). La log-normal falla porque hay "
        "pacientes de 0 años; la Poisson porque supone media = varianza (28.5 vs 317).",
        horizontal=True, size=18)


def slide_kde(prs):
    imagen_y_texto(prs, "Edad: la KDE describe mejor los datos", "03_EDA.ipynb", "KDE de la edad por desenlace", [
        "En una muestra de 5,000 pacientes, KS de la KDE = 0.027 contra 0.036 de la gamma",
        "Sobrevivientes: concentrados entre 10 y 30 años. Fallecidos: densidad plana con mucha masa entre 50 y 90",
        "Desde los ~43 años la densidad de los fallecidos supera a la de los sobrevivientes",
    ], "KDE con kernel gaussiano y ancho de banda de Silverman (factor 0.088).")


def slide_conclusiones_eda(prs):
    contenido(prs, "Conclusiones del EDA", [
        "**Datos:** 251,900 casos confirmados, sin faltantes ni duplicados",
        "**Desbalance:** 0.70% de muertes; estratificar y no usar la exactitud como métrica",
        "**Edad:** relación en forma de \"U\" y no lineal con el log-odds → probar edad²",
        "**Distribución de la edad:** gamma es la mejor paramétrica, pero KS la rechaza; la KDE ajusta mejor",
        "**Factores más asociados:** hospitalización, enf. renal, cirrosis, inmunosupresión, diabetes, "
        "hemorrágicos e hipertensión",
        "**Hospitalización** (96% de los fallecidos) refleja la gravedad más que causar la muerte: "
        "se analizará incluirla y excluirla",
        "**Multicolinealidad:** solo entre dummies de serotipo, lo cual es esperado",
    ], size=18)


# ───────────────────────── 04 · Modelo ─────────────────────────
def slide_div_modelo(prs):
    seccion(prs, "04 · Modelo de regresión logística",
            "Probabilidad de fallecer y odds ratio de cada variable",
            "Cuarta etapa: notebook 04_modelo.ipynb. p = 1 / (1 + e^-(β0 + β1·x1 + … + βk·xk)); "
            "el efecto de cada variable se interpreta con su odds ratio e^β.")


def slide_decisiones_eda(prs):
    s = nueva(prs, L_SOLO_TITULO, "Decisiones del modelo que vienen del EDA")
    a_tabla(s, [
        ("Hallazgo del EDA", "Decisión en el modelo"),
        ("Solo 0.70% de fallecidos", "Undersampling 75/25 y partición estratificada; evaluar con recall, "
                                     "precision, F1 y ROC-AUC"),
        ("La edad no es lineal en el log-odds (forma de \"U\")", "Agregar el término edad² (EDAD_ANOS_2)"),
        ("HOSPITALIZADO refleja la gravedad del paciente", "Comparar el modelo final con y sin esta variable"),
        ("Muchas variables con asociación débil", "Backward selection con p-value, AIC y BIC"),
    ], Inches(0.67), Inches(1.9), Inches(12), anchos=[5, 7], size=18, alto_fila=0.7)


def slide_edad_logodds(prs):
    imagen_y_texto(prs, "La edad entra como parábola en el log-odds", "04_modelo.ipynb", "Log-odds empírico (tamaño", [
        "**Lineal:** AIC 20,427.7 · **Con edad²:** AIC 20,139.9 (≈290 menos)",
        "β₁ = −0.036 y β₂ = 0.00076: forma de \"U\" con mínimo cerca de los 24 años",
        "La recta subestima el riesgo de niños y adultos mayores, y lo sobreestima entre 20 y 50 años",
    ], "Ajustes univariados (solo edad) con todos los datos, antes del undersampling. Los puntos son el "
       "log-odds empírico por intervalos de 5 años. Referencia: log-odds de −4.6 ≈ 1% de mortalidad, "
       "−2.2 ≈ 10%, 0 = 50%. El modelo sigue siendo lineal en los parámetros.", ancho=11)


def slide_undersampling(prs):
    s = nueva(prs, L_DOS, "Undersampling: 75% sobrevivientes / 25% fallecidos",
              "Los odds ratios no cambian porque el muestreo depende solo del desenlace. El intercepto sí: "
              "se corrige con β0_real = β0 − ln[(0.25/0.75) / (π/(1−π))]. El conjunto de prueba también queda "
              "con 25% de fallecidos, así que precision, F1 y accuracy no reflejan la proporción real.")
    vinetas(ph(s, 1), [
        "Se conservan **todos los fallecidos** y se toman al azar 3 sobrevivientes por cada uno",
        "Solo se usa el **2% de los sobrevivientes**",
        "Los OR no cambian; el intercepto sí → hay que corregir las probabilidades",
        "La edad conserva su distribución dentro de cada clase",
    ], size=18)
    x, y, w, h = _quitar(s, 2)
    a_tabla(s, [
        ("", "Original", "Undersampling"),
        ("Casos", "251,900", "7,052"),
        ("Fallecidos", "1,763", "1,763"),
        ("Sobrevivientes", "250,137", "5,289"),
        ("Mortalidad", "0.70%", "25.0%"),
    ], x, y, w, anchos=[2.2, 1.8, 2], size=18)


def slide_modelo_completo(prs):
    contenido(prs, "Partición y modelo completo", [
        "**Partición 70/30 estratificada por DEFUNCION:** entrenamiento 4,936 (1,234 fallecidos) y "
        "prueba 2,116 (529); ambos con 25%",
        "**Modelo completo con 28 variables** (statsmodels, sobre entrenamiento):",
        ("Log-verosimilitud −1,869.8 contra −2,775.7 del modelo nulo; LLR p-value ≈ 0", 1),
        ("Pseudo R² = 0.326 (no comparable con un modelo sobre los datos originales)", 1),
        ("Significativas: edad y edad², hemorrágicos, diabetes, enf. renal, hospitalización, DENV4 e indígena", 1),
        ("No significativas: sexo, costera, hipertensión, inmunosupresión, cirrosis, embarazo, "
         "demás serotipos y todos los meses", 1),
    ], "El conjunto de prueba no se usa ni en el ajuste ni en la selección de variables; solo al final.")


def slide_backward(prs):
    s = nueva(prs, L_DOS, "Backward selection: de 28 a 7 variables",
              "Se eliminaron los 11 meses, costera, cirrosis, sexo, DENV2, embarazo, sin serotipo, DENV3, "
              "hipertensión, inmunosupresión e indígena. Con undersampling se pierde potencia: cirrosis, "
              "inmunosupresión y algunos meses eran significativos con todos los datos.")
    vinetas(ph(s, 1), [
        "En cada paso se elimina la variable con mayor p-value mientras sea > 0.05",
        "**El BIC baja en todos los pasos** (3,986 → 3,833)",
        "El AIC tiene su mínimo con 12 variables (3,776.6); el modelo final queda 3.9 arriba",
        "**Final:** edad, edad², hemorrágicos, diabetes, enf. renal, hospitalización y DENV4",
        "Sin edad² el AIC sube 57 y el BIC 51 → se confirma la \"U\"",
    ], size=17)
    a_grafica_lineas(s, 2, [str(n) for n in range(28, 6, -1)], {
        "AIC": [3797.650, 3795.655, 3793.696, 3791.742, 3789.968, 3788.103, 3786.405, 3784.735, 3783.191,
                3781.981, 3780.870, 3779.961, 3779.180, 3778.604, 3778.394, 3776.640, 3776.626, 3776.636,
                3777.317, 3777.870, 3778.823, 3780.571],
        "BIC": [3986.275, 3977.776, 3969.312, 3960.854, 3952.576, 3944.207, 3936.004, 3927.830, 3919.782,
                3912.068, 3904.452, 3897.038, 3889.753, 3882.673, 3875.959, 3867.701, 3861.182, 3854.688,
                3848.864, 3842.913, 3837.362, 3832.606],
    }, minimo=3750)


def slide_modelo_final(prs):
    s = nueva(prs, L_SOLO_TITULO, "Modelo final",
              "Pseudo R² = 0.322, log-verosimilitud −1,882.3, LLR p-value ≈ 0. El intercepto está en la escala "
              "de la muestra (25% de fallecidos); corregido a la mortalidad real es −7.670.")
    a_tabla(s, [
        ("Variable", "Coeficiente (β)", "Odds ratio (e^β)", "IC 95% del OR", "p-value"),
        ("Hospitalizado", "3.516", "33.66", "25.24 – 44.90", "< 0.001"),
        ("DENV4", "1.616", "5.03", "1.61 – 15.67", "0.005"),
        ("Enfermedad renal", "1.236", "3.44", "1.71 – 6.91", "< 0.001"),
        ("Hemorrágicos", "1.127", "3.09", "1.73 – 5.51", "< 0.001"),
        ("Diabetes", "0.677", "1.97", "1.39 – 2.78", "< 0.001"),
        ("Edad²", "0.0007", "1.0007", "1.0005 – 1.0008", "< 0.001"),
        ("Edad", "−0.031", "0.970", "0.956 – 0.983", "< 0.001"),
        ("Intercepto", "−3.814", "–", "–", "< 0.001"),
    ], Inches(0.67), Inches(1.9), Inches(12), anchos=[3, 2.2, 2.3, 2.6, 1.6], size=18, alto_fila=0.48)


def slide_or_final(prs):
    texto_y_grafica(prs, "Interpretación de los odds ratios", [
        "**Hospitalizado:** odds de fallecer 34 veces los de un ambulatorio",
        "Enf. renal ×3.4, hemorrágicos ×3.1, diabetes ×2.0",
        "Ningún intervalo de confianza incluye al 1",
        "**DENV4 (×5.0) no es confiable:** solo 31 pacientes en entrenamiento (11 fallecidos) e IC de 1.6 a 15.7",
        "Los OR son válidos con undersampling y parecidos al modelo con todos los datos",
    ], ["Hospitalizado", "DENV4", "Enf. renal", "Hemorrágicos", "Diabetes"],
        [33.66, 5.03, 3.44, 3.09, 1.97], "0.0",
        "Los OR se interpretan manteniendo constantes las demás variables. Modelo sin undersampling: "
        "hospitalización 31.4, enf. renal 4.2, diabetes 1.7. En los datos completos la mortalidad con DENV4 "
        "es 0.86%, casi igual a la general, y sin undersampling no era significativo.",
        horizontal=True, size=17)


def slide_efecto_edad(prs):
    s = nueva(prs, L_DOS, "Efecto de la edad en la probabilidad de fallecer",
              "Paciente de referencia: sin comorbilidades, sin hemorragia y con serotipo distinto de DENV4. "
              "Sin la corrección del intercepto, un hospitalizado de 85 años tendría ≈87%, porque el modelo "
              "calcula probabilidades para una población donde 1 de cada 4 fallece.")
    vinetas(ph(s, 1), [
        "Edad y edad² no se interpretan por separado: juntas forman la \"U\"",
        "Se corrige el intercepto (−3.81 → −7.67) para volver a la mortalidad real de 0.70%",
        "El riesgo mínimo está alrededor de los **23 años**",
        "Un hospitalizado de 85 años sin comorbilidades: **12.3%**",
    ], size=18)
    x, y, w, h = _quitar(s, 2)
    a_tabla(s, [
        ("Edad", "Ambulatorio", "Hospitalizado"),
        ("0 años", "0.05%", "1.55%"),
        ("20 años", "0.03%", "1.09%"),
        ("60 años", "0.08%", "2.64%"),
        ("85 años", "0.41%", "12.29%"),
    ], x, y, w, anchos=[1.8, 2, 2], size=18)


def slide_hospitalizado(prs):
    s = nueva(prs, L_DOS, "¿Qué pasa si se quita HOSPITALIZADO?",
              "La hospitalización no la causa la muerte, pero sí es consecuencia de la gravedad. Se conserva "
              "porque se conoce desde la primera atención y mejora mucho el ajuste. El OR de 34 no significa "
              "que hospitalizar cause la muerte: se hospitaliza a los pacientes más graves.")
    vinetas(ph(s, 1), [
        "Refleja la gravedad del paciente (posible endogeneidad)",
        "**Sin ella el modelo empeora mucho:** ROC-AUC de 0.865 a 0.700",
        "Las comorbilidades suben su OR: parte de su efecto ocurre a través de la hospitalización",
        "DENV4 baja de 5.0 a 1.7: resultado frágil",
        "**Decisión: se conserva.** Los demás OR se leen como efecto adicional a la hospitalización",
    ], size=17)
    x, y, w, h = _quitar(s, 2)
    a_tabla(s, [
        ("", "Con", "Sin"),
        ("AIC", "3,781", "5,038"),
        ("BIC", "3,833", "5,083"),
        ("Pseudo R²", "0.322", "0.095"),
        ("ROC-AUC (prueba)", "0.865", "0.700"),
        ("OR enf. renal", "3.44", "6.56"),
        ("OR diabetes", "1.97", "2.64"),
        ("OR hemorrágicos", "3.09", "3.76"),
        ("OR DENV4", "5.03", "1.65"),
    ], x, y, w, anchos=[2.6, 1.5, 1.5], size=16)


def slide_separabilidad(prs):
    imagen_y_texto(prs, "Separabilidad de las probabilidades estimadas", "04_modelo.ipynb",
                   "distribution_prob(y_test", [
        "Sobrevivientes: mediana de 2.5%. Fallecidos: mediana de 40.5% (conjunto de prueba)",
        "Dos grupos: ambulatorios (pico izquierdo) y hospitalizados (pico derecho)",
        "Los sobrevivientes hospitalizados se traslapan con los fallecidos",
    ], "Probabilidades sin corregir (escala de la muestra con 25% de fallecidos), eje en escala log. "
       "En la población real serían menores, pero el orden de los pacientes es el mismo.", ancho=9)


def slide_metricas(prs):
    s = nueva(prs, L_DOS, "Métricas en prueba y curva ROC",
              "El umbral de la curva ROC maximiza TPR − FPR (índice de Youden) y se calcula solo con "
              "entrenamiento. En este problema un falso negativo es más costoso que un falso positivo, "
              "por eso el recall es la métrica principal. Precision, F1 y accuracy dependen de la proporción "
              "de 25% de fallecidos del conjunto de prueba. Nota: en el notebook, la celda de la sección 8.1 "
              "usa umbral = 0.2; las cifras de 0.5 se verificaron reproduciendo el modelo.")
    izq = ph(s, 1)
    x, y, w = izq.left, izq.top, izq.width
    a_tabla(s, [
        ("Métrica (prueba)", "Umbral 0.5", "Umbral 0.143"),
        ("Recall", "40.3%", "96.8%"),
        ("Precision", "74.2%", "44.1%"),
        ("F1", "0.522", "0.606"),
        ("Accuracy", "81.6%", "68.5%"),
        ("Fallecidos detectados", "213 de 529", "512 de 529"),
    ], x, y, w, anchos=[2.6, 1.6, 1.7], size=15)
    izq.top, izq.height = y + Inches(2.6), Inches(2.4)
    vinetas(izq, [
        "**ROC-AUC: 0.860 en entrenamiento y 0.865 en prueba** (buena discriminación, sin sobreajuste)",
        "Con 0.5 se escapa el 60% de los fallecidos",
        "El umbral de 0.143 sale de la curva ROC",
    ], size=16)
    xr, yr, wr, hr = _quitar(s, 2)
    imagen_en(s, imagen_notebook("04_modelo.ipynb", "plt.title('Curva ROC')"), xr, yr, wr, hr)


def slide_proporcion_real(prs):
    s = nueva(prs, L_DOS, "Desempeño con la proporción real de fallecidos",
              "Conjunto armado solo con pacientes que el modelo nunca vio: los 529 fallecidos de prueba y "
              "sobrevivientes que quedaron fuera del undersampling, para llegar a 0.70%. Mismo umbral. "
              "Con un evento tan raro, una precision baja es inevitable con cualquier umbral de recall alto.")
    vinetas(ph(s, 1), [
        "75,584 pacientes nunca vistos, con 0.70% de mortalidad (529 fallecidos)",
        "Recall (96.8%) y ROC-AUC (0.865) no cambian",
        "**La precision cae a 1.7%:** se marca como alto riesgo al 39.2% de los pacientes",
        "**Un paciente de bajo riesgo sobrevive en el 99.96% de los casos**",
        "Sirve para **descartar riesgo**, no para predecir con certeza quién morirá",
    ], size=17)
    x, y, w, h = _quitar(s, 2)
    a_tabla(s, [
        ("", "Pred. sobrevive", "Pred. fallece"),
        ("Sobrevivió", "45,905", "29,150"),
        ("Falleció", "17", "512"),
    ], x, y, w, anchos=[1.8, 2.1, 2], size=18, alto_fila=0.6)


def slide_conclusiones_modelo(prs):
    contenido(prs, "Conclusiones del modelo", [
        "**Buena discriminación:** ROC-AUC de 0.865, igual en entrenamiento y prueba",
        "Con umbral de 0.143 detecta al 96.8% de los fallecidos; útil como herramienta de tamizaje",
        "**Factores de riesgo:** hospitalización (OR 34), enf. renal (3.4), hemorrágicos (3.1), "
        "diabetes (2.0) y edad en forma de \"U\"",
        "**Sin efecto al ajustar:** sexo, embarazo, costera, hipertensión, demás serotipos y meses",
        ("La hipertensión se explica por la edad y la diabetes, con las que está correlacionada", 1),
        "**Undersampling:** no mejora la capacidad predictiva (AUC 0.865 contra 0.863)",
        ("Mantiene los OR, pero pierde potencia estadística y obliga a corregir el intercepto", 1),
    ], size=18)


def slide_recomendaciones(prs):
    contenido(prs, "Recomendaciones y mejoras futuras", [
        "**Uso:** priorizar la vigilancia de hospitalizados, mayores de 60, menores de 5 y pacientes con "
        "enf. renal, diabetes o hemorragias",
        "**Estabilidad:** repetir el undersampling con distintas semillas (DENV4 probablemente no se repita)",
        "**Desbalance:** comparar con pesos por clase o sobremuestreo, que no descartan información",
        "**Más variables clínicas:** signos de alarma, plaquetas, hematocrito, días desde el inicio de síntomas",
        "**Modelo solo con hospitalizados**, donde ocurren casi todas las muertes",
        "**Calidad de los datos:** posible subregistro de comorbilidades; asociación no implica causalidad",
    ], size=18)


# ───────────────────────── Cierre ─────────────────────────
def slide_insights(prs):
    contenido(prs, "Hallazgos principales", [
        "**1 de cada 142** pacientes con dengue confirmado fallece, pero el riesgo está muy concentrado: "
        "el 39% de los pacientes (alto riesgo) reúne el 97% de las muertes",
        "Un **hospitalizado de 85 años** sin comorbilidades tiene 12.3% de probabilidad de fallecer, "
        "contra 1.1% a los 20 años",
        "La **enfermedad renal** es la comorbilidad más peligrosa: 12.2% de mortalidad, 18 veces la de "
        "quienes no la tienen",
        "**Balancear las clases no mejora el modelo:** cambia la escala de las probabilidades y el umbral, "
        "pero no la capacidad de distinguir a los pacientes en riesgo",
    ])


# ───────────────────────── Secciones ─────────────────────────
# Para agregar EDA o modelo: añade aquí ("03 · EDA", [...]) antes de "Siguientes pasos".
SECCIONES = [
    ("Introducción", [slide_portada, slide_ruta, slide_pregunta]),
    ("01 · Merge de datos", [slide_div_merge, slide_fuente, slide_estrategia, slide_trazabilidad,
                             slide_resultado_merge]),
    ("02 · Limpieza", [slide_div_limpieza, slide_embudo, slide_decision, slide_variables,
                       slide_recodificacion, slide_serotipo, slide_mes, slide_anio, slide_base_final]),
    ("03 · Análisis exploratorio", [slide_div_eda, slide_desbalance, slide_edad, slide_grupo_edad,
                                    slide_binarias_eda, slide_or_crudo, slide_serotipo_mes_eda,
                                    slide_correlacion, slide_distribuciones, slide_kde,
                                    slide_conclusiones_eda]),
    ("04 · Modelo", [slide_div_modelo, slide_decisiones_eda, slide_edad_logodds, slide_undersampling,
                     slide_modelo_completo, slide_backward, slide_modelo_final, slide_or_final,
                     slide_efecto_edad, slide_hospitalizado, slide_separabilidad, slide_metricas,
                     slide_proporcion_real, slide_conclusiones_modelo, slide_recomendaciones]),
    ("Cierre", [slide_insights]),
]


def agregar_secciones(prs, secciones_ids):
    """Escribe las secciones de PowerPoint (p14:sectionLst) en presentation.xml."""
    P14 = "http://schemas.microsoft.com/office/powerpoint/2010/main"
    pres = prs.part._element
    ext_lst = pres.find(qn("p:extLst"))
    if ext_lst is None:
        ext_lst = etree.SubElement(pres, qn("p:extLst"))
    ext = etree.SubElement(ext_lst, qn("p:ext"))
    ext.set("uri", "{521415D9-36F7-43E2-AB2F-B90AF26B5E84}")
    sl = etree.SubElement(ext, f"{{{P14}}}sectionLst", nsmap={"p14": P14})
    for nombre, ids in secciones_ids:
        sec = etree.SubElement(sl, f"{{{P14}}}section")
        sec.set("name", nombre)
        sec.set("id", "{" + str(uuid.uuid4()).upper() + "}")
        lst = etree.SubElement(sec, f"{{{P14}}}sldIdLst")
        for i in ids:
            etree.SubElement(lst, f"{{{P14}}}sldId").set("id", str(i))


def main():
    prs = Presentation()
    a_16x9(prs)
    prs.core_properties.title = "Factores asociados a la defunción por dengue en México"
    secciones_ids = []
    for nombre, funciones in SECCIONES:
        ids = []
        for f in funciones:
            f(prs)
            ids.append(prs.slides._sldIdLst[-1].id)
        secciones_ids.append((nombre, ids))
    agregar_secciones(prs, secciones_ids)
    prs.save(SALIDA)
    print(f"Guardado: {SALIDA} ({len(prs.slides)} diapositivas)")


if __name__ == "__main__":
    main()
