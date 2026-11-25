"""
Genera la presentación del proyecto (Parte 1: merge y limpieza).

Uso:  python presentacion/generar_presentacion.py
Salida: presentacion/Presentacion_Dengue.pptx

Para agregar secciones nuevas (EDA, modelo):
  1. Escribe funciones de diapositiva como las de abajo (slide_* ).
  2. Agrega una entrada en SECCIONES al final del archivo.
Los helpers (titulo_contenido, tarjeta, stat, grafica_barras, etc.) mantienen el estilo.
"""
from pathlib import Path
import uuid

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION
from pptx.oxml.ns import qn
from lxml import etree

SALIDA = Path(__file__).with_name("Presentacion_Dengue.pptx")

# ───────────────────────── Paleta y tipografía ─────────────────────────
INK = "13302B"      # verde selva profundo (fondos oscuros, títulos)
INK_2 = "1D4A41"    # círculos decorativos sobre fondo oscuro
PAPER = "FFFFFF"
MIST = "EEF3F1"     # tinte para tarjetas
LINE = "C9D3D0"
TEXT = "1E2A28"
MUTED = "5B6B67"
TEAL = "2E8B7A"     # casos / conteos
RED = "C8384E"      # muertes / mortalidad
AMBER = "E0A030"    # categoría de referencia / acento
AMBER_TXT = "9A6A10"
LIGHT_ON_DARK = "B9CFC9"

HEAD = "Cambria"
BODY = "Calibri"

W, H = 13.333, 7.5
PIE = "Defunción por dengue · Proyecto AEM"


def rgb(h):
    return RGBColor.from_string(h)


# ───────────────────────── Helpers básicos ─────────────────────────
def _sin_estilo(shape):
    st = shape._element.find(qn("p:style"))
    if st is not None:
        shape._element.remove(st)


def fondo(slide, color):
    f = slide.background.fill
    f.solid()
    f.fore_color.rgb = rgb(color)


def forma(slide, tipo, x, y, w, h, fill=None, line=None, radio=None, nombre=None):
    s = slide.shapes.add_shape(tipo, Inches(x), Inches(y), Inches(w), Inches(h))
    _sin_estilo(s)
    if fill:
        s.fill.solid()
        s.fill.fore_color.rgb = rgb(fill)
    else:
        s.fill.background()
    if line:
        s.line.color.rgb = rgb(line)
        s.line.width = Pt(1)
    else:
        s.line.fill.background()
    if radio is not None:
        s.adjustments[0] = radio
    if nombre:
        s.name = nombre
    return s


def tarjeta(slide, x, y, w, h, fill=MIST, line=None, nombre=None):
    return forma(slide, MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h, fill, line, radio=0.06, nombre=nombre)


def _bullet(p, color, nivel_in=0.25):
    pPr = p._p.get_or_add_pPr()
    pPr.set("marL", str(Inches(nivel_in)))
    pPr.set("indent", str(-Inches(nivel_in)))
    buClr = etree.SubElement(pPr, qn("a:buClr"))
    etree.SubElement(buClr, qn("a:srgbClr")).set("val", color)
    etree.SubElement(pPr, qn("a:buFont")).set("typeface", "Arial")
    etree.SubElement(pPr, qn("a:buChar")).set("char", "•")


def texto(slide, x, y, w, h, parrafos, size=16, color=TEXT, bold=False, font=BODY,
          align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, after=4, nombre=None):
    """parrafos: str | lista de str/dict. dict: text|runs, size, color, bold, italic,
    font, bullet, align, after. runs: lista de (texto, {opciones})."""
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    if nombre:
        tb.name = nombre
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = anchor
    if isinstance(parrafos, (str, dict)):
        parrafos = [parrafos]
    for i, item in enumerate(parrafos):
        if isinstance(item, str):
            item = {"text": item}
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = item.get("align", align)
        p.space_after = Pt(item.get("after", after))
        runs = item.get("runs") or [(item.get("text", ""), {})]
        for t, o in runs:
            r = p.add_run()
            r.text = t
            f = r.font
            f.name = o.get("font", item.get("font", font))
            f.size = Pt(o.get("size", item.get("size", size)))
            f.bold = o.get("bold", item.get("bold", bold))
            f.italic = o.get("italic", item.get("italic", False))
            f.color.rgb = rgb(o.get("color", item.get("color", color)))
        if item.get("bullet"):
            _bullet(p, item.get("bullet_color", TEAL))
    return tb


def circulo_num(slide, x, y, d, txt, fill=TEAL, color=PAPER, size=16):
    c = forma(slide, MSO_SHAPE.OVAL, x, y, d, d, fill)
    tf = c.text_frame
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = txt
    r.font.size, r.font.bold, r.font.name = Pt(size), True, BODY
    r.font.color.rgb = rgb(color)
    return c


def stat(slide, x, y, w, numero, etiqueta, color=INK, size=40, label_color=MUTED, label_size=13):
    texto(slide, x, y, w, size / 72 * 1.25, numero, size=size, bold=True, color=color, font=HEAD)
    texto(slide, x, y + size / 72 * 1.3, w, 0.7, etiqueta, size=label_size, color=label_color)


def notas(slide, txt):
    slide.notes_slide.notes_text_frame.text = txt


def pie(slide, prs, oscuro=False):
    n = prs.slides.index(slide) + 1
    c = LIGHT_ON_DARK if oscuro else MUTED
    texto(slide, 0.6, 7.0, 8, 0.3, PIE, size=10, color=c)
    texto(slide, W - 1.6, 7.0, 1.0, 0.3, str(n), size=10, color=c, align=PP_ALIGN.RIGHT)


def nueva(prs):
    return prs.slides.add_slide(prs.slide_layouts[6])  # en blanco


def titulo_contenido(prs, kicker, titulo):
    """Diapositiva clara con etiqueta de sección + título (posición fija en todo el deck)."""
    s = nueva(prs)
    fondo(s, PAPER)
    texto(s, 0.6, 0.42, 12, 0.3, kicker.upper(), size=12, bold=True, color=TEAL)
    texto(s, 0.6, 0.72, 12.1, 0.8, titulo, size=30, bold=True, color=INK, font=HEAD,
          anchor=MSO_ANCHOR.TOP, nombre="Título")
    pie(s, prs)
    return s


def divisor(prs, numero, titulo, subtitulo):
    s = nueva(prs)
    fondo(s, INK)
    forma(s, MSO_SHAPE.OVAL, 8.4, -1.6, 7.0, 7.0, INK_2)
    forma(s, MSO_SHAPE.OVAL, 10.6, 4.6, 3.6, 3.6, INK_2)
    texto(s, 0.9, 1.9, 4, 1.6, numero, size=96, bold=True, color=AMBER, font=HEAD)
    texto(s, 0.9, 3.55, 8.5, 1.0, titulo, size=40, bold=True, color=PAPER, font=HEAD, nombre="Título")
    texto(s, 0.9, 4.55, 8.0, 1.0, subtitulo, size=18, color=LIGHT_ON_DARK)
    pie(s, prs, oscuro=True)
    return s


def grafica_barras(slide, x, y, w, h, categorias, valores, colores, formato="0.0%",
                   horizontal=False, label_size=12, cat_size=12, gap=60, max_val=None):
    cd = CategoryChartData()
    cd.categories = categorias
    cd.add_series("serie", valores)
    tipo = XL_CHART_TYPE.BAR_CLUSTERED if horizontal else XL_CHART_TYPE.COLUMN_CLUSTERED
    ch = slide.shapes.add_chart(tipo, Inches(x), Inches(y), Inches(w), Inches(h), cd).chart
    ch.has_legend = False
    ch.has_title = False
    ch.font.name = BODY
    ch.font.size = Pt(cat_size)
    ch.font.color.rgb = rgb(MUTED)
    plot = ch.plots[0]
    plot.gap_width = gap
    plot.vary_by_categories = False
    plot.has_data_labels = True
    dl = plot.data_labels
    dl.number_format = formato
    dl.number_format_is_linked = False
    dl.position = XL_LABEL_POSITION.OUTSIDE_END
    dl.font.size = Pt(label_size)
    dl.font.bold = True
    dl.font.color.rgb = rgb(TEXT)
    ser = plot.series[0]
    for i, col in enumerate(colores if isinstance(colores, list) else [colores] * len(valores)):
        pt = ser.points[i]
        pt.format.fill.solid()
        pt.format.fill.fore_color.rgb = rgb(col)
    va = ch.value_axis
    va.visible = False
    va.has_major_gridlines = False
    va.minimum_scale = 0
    if max_val:
        va.maximum_scale = max_val
    ca = ch.category_axis
    ca.format.line.color.rgb = rgb(LINE)
    ca.tick_labels.font.size = Pt(cat_size)
    ca.tick_labels.font.color.rgb = rgb(MUTED)
    ca.has_major_gridlines = False
    if horizontal:
        ca.reverse_order = True
    return ch


# ───────────────────────── Introducción ─────────────────────────
def slide_portada(prs):
    s = nueva(prs)
    fondo(s, INK)
    forma(s, MSO_SHAPE.OVAL, 8.9, -2.2, 6.8, 6.8, INK_2)
    texto(s, 0.9, 1.0, 7.5, 0.4, "PROYECTO FINAL · ANÁLISIS ESTADÍSTICO MULTIVARIADO · ITESO",
          size=12, bold=True, color=AMBER)
    texto(s, 0.9, 1.6, 7.6, 2.4, "Factores asociados a la defunción por dengue en México",
          size=44, bold=True, color=PAPER, font=HEAD, nombre="Título")
    texto(s, 0.9, 4.1, 7.2, 0.9,
          "Regresión logística binaria con datos abiertos de la DGE (SINAVE), 2020–2026",
          size=18, color=LIGHT_ON_DARK)
    texto(s, 0.9, 5.7, 7.2, 0.4, "Parte 1 · Construcción y limpieza de la base", size=16,
          bold=True, color=PAPER)

    for i, (num, lab) in enumerate([("294", "fotos semanales de datos abiertos"),
                                     ("1,276,897", "registros únicos tras el merge"),
                                     ("251,900", "casos confirmados para el modelo")]):
        y = 2.0 + i * 1.45
        texto(s, 9.3, y, 3.6, 0.75, num, size=40, bold=True, color=AMBER if i == 2 else PAPER, font=HEAD)
        texto(s, 9.3, y + 0.75, 3.6, 0.4, lab, size=13, color=LIGHT_ON_DARK)
    notas(s, "Presentamos el proyecto final: un modelo de regresión logística para identificar "
             "qué factores se asocian a la muerte en pacientes con dengue confirmado en México. "
             "Esta primera parte explica cómo construimos y limpiamos la base de datos.")


def slide_ruta(prs):
    s = titulo_contenido(prs, "Ruta del proyecto", "El proyecto avanza en cuatro etapas")
    etapas = [
        ("01", "Merge de datos", "Unir 294 fotos semanales en una sola base con un registro por paciente.", True),
        ("02", "Limpieza", "Filtrar casos confirmados, seleccionar variables y recodificar a 1/0.", True),
        ("03", "Análisis exploratorio", "Describir la mortalidad según cada variable.", False),
        ("04", "Modelo", "Estimar la regresión logística e interpretar los coeficientes.", False),
    ]
    cw, gap, x0, y0, ch = 2.75, 0.35, 0.64, 1.95, 3.9
    for i, (n, t, d, hecho) in enumerate(etapas):
        x = x0 + i * (cw + gap)
        tarjeta(s, x, y0, cw, ch, fill=MIST if hecho else PAPER, line=None if hecho else LINE,
                nombre=f"Etapa {n}")
        circulo_num(s, x + 0.3, y0 + 0.35, 0.8, n, fill=TEAL if hecho else LINE,
                    color=PAPER if hecho else MUTED, size=20)
        texto(s, x + 0.3, y0 + 1.4, cw - 0.6, 0.5, t, size=20, bold=True, color=INK, font=HEAD)
        texto(s, x + 0.3, y0 + 1.95, cw - 0.6, 1.2, d, size=14, color=TEXT)
        chip = forma(s, MSO_SHAPE.ROUNDED_RECTANGLE, x + 0.3, y0 + ch - 0.7, 1.55, 0.36,
                     fill=TEAL if hecho else MIST, radio=0.5)
        tf = chip.text_frame
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run()
        r.text = "Completado" if hecho else "Próximamente"
        r.font.size, r.font.bold, r.font.name = Pt(12), True, BODY
        r.font.color.rgb = rgb(PAPER if hecho else MUTED)
    texto(s, 0.64, 6.2, 12, 0.4,
          "Esta presentación cubre las etapas 01 y 02. Las etapas 03 y 04 se agregarán después.",
          size=14, color=MUTED)
    notas(s, "El trabajo está dividido en cuatro notebooks. Hoy cubrimos los dos primeros: "
             "cómo unimos los archivos de datos abiertos y cómo limpiamos la base para el modelo.")


def slide_pregunta(prs):
    s = titulo_contenido(prs, "Planteamiento", "¿Qué factores se asocian a morir por dengue?")
    tarjeta(s, 0.6, 1.8, 4.1, 4.7, fill=INK, nombre="Variable objetivo")
    texto(s, 0.95, 2.15, 3.4, 0.3, "VARIABLE OBJETIVO", size=12, bold=True, color=AMBER)
    texto(s, 0.95, 2.55, 3.4, 0.7, "DEFUNCION", size=32, bold=True, color=PAPER, font=HEAD)
    texto(s, 0.95, 3.4, 3.4, 1.0, ["1 = el paciente murió", "0 = no murió"], size=16, color=PAPER)
    texto(s, 0.95, 4.6, 3.4, 1.6,
          "Población: casos confirmados de dengue en México con inicio de síntomas entre 2020 y 2026.",
          size=14, color=LIGHT_ON_DARK)

    grupos = [
        ("Demográficas", "Sexo · Edad · Indígena"),
        ("Comorbilidades", "Diabetes · Hipertensión · Enfermedad renal · Inmunosupresión · Cirrosis hepática · Embarazo"),
        ("Gravedad clínica", "Manifestaciones hemorrágicas · Tipo de paciente (hospitalizado)"),
        ("Laboratorio", "Resultado de PCR (serotipo del virus)"),
        ("Tiempo y lugar", "Mes de inicio de síntomas · Entidad de asignación (costera o no)"),
    ]
    texto(s, 5.1, 1.8, 7.6, 0.35, "VARIABLES INDEPENDIENTES", size=12, bold=True, color=TEAL)
    for i, (g, v) in enumerate(grupos):
        y = 2.25 + i * 0.86
        tarjeta(s, 5.1, y, 7.6, 0.74, fill=MIST)
        texto(s, 5.35, y, 2.2, 0.74, g, size=15, bold=True, color=INK, anchor=MSO_ANCHOR.MIDDLE)
        texto(s, 7.55, y, 5.0, 0.74, v, size=14, color=TEXT, anchor=MSO_ANCHOR.MIDDLE)
    notas(s, "La variable de respuesta es binaria: si el paciente murió o no. Las variables "
             "independientes se eligieron al inicio del proyecto (archivo VARIABLES.rtf) y se "
             "agrupan en demográficas, comorbilidades, gravedad, laboratorio y tiempo/lugar.")


# ───────────────────────── 01 · Merge ─────────────────────────
def slide_div_merge(prs):
    s = divisor(prs, "01", "Merge de datos", "De 294 fotos semanales a una base con un registro por paciente")
    notas(s, "Primera etapa: notebook 01_merge_datos.ipynb.")


def slide_fuente(prs):
    s = titulo_contenido(prs, "01 · Merge de datos", "Los datos llegan como 294 fotos semanales acumuladas")
    texto(s, 0.6, 1.9, 5.6, 4.4, [
        {"runs": [("Fuente: ", {"bold": True}), ("Datos Abiertos de la DGE, bases históricas de "
                  "enfermedades transmitidas por vector (dic-2020 a nov-2026).", {})], "bullet": True, "after": 14},
        {"runs": [("Cada archivo es acumulado: ", {"bold": True}), ("trae todos los casos del año "
                  "hasta esa fecha, no solo los de esa semana.", {})], "bullet": True, "after": 14},
        {"runs": [("Los registros se corrigen ", {"bold": True}), ("de una foto a otra: cambian "
                  "ESTATUS_CASO, DEFUNCION, etc.", {})], "bullet": True, "after": 14},
        {"runs": [("Mismo formato: ", {"bold": True}), ("las 294 fotos tienen las mismas 28 columnas.", {})],
         "bullet": True},
    ], size=16)
    texto(s, 6.8, 1.9, 5.9, 0.35, "Archivos por año", size=14, bold=True, color=INK)
    grafica_barras(s, 6.7, 2.3, 6.0, 4.1, ["2020", "2021", "2022", "2023", "2024", "2025", "2026"],
                   [4, 50, 53, 51, 50, 48, 38], TEAL, formato="0", label_size=13, cat_size=13)
    texto(s, 6.8, 6.45, 5.9, 0.3, "2020 solo incluye diciembre; 2026 llega hasta el 5 de noviembre.",
          size=11, color=MUTED)
    notas(s, "Lo importante es que cada zip es una foto acumulada del año: el mismo paciente aparece "
             "en muchas fotos, y su información puede cambiar entre ellas. Por eso no basta con "
             "concatenar los archivos.")


def slide_estrategia(prs):
    s = titulo_contenido(prs, "01 · Merge de datos", "El merge conserva la versión más reciente de cada registro")
    pasos = [
        ("1", "Ordenar", "Las fotos se leen en orden cronológico según la fecha del nombre del archivo."),
        ("2", "Leer", "Solo el CSV válido de cada zip; todo como texto para no perder información."),
        ("3", "Deduplicar", "Concatenar foto por foto y quedarse con la última versión de cada ID_REGISTRO."),
        ("4", "Bitácora", "Guardar en qué fotos aparece cada registro y qué valor de DEFUNCION tenía."),
    ]
    cw, gap = 2.8, 0.3
    for i, (n, t, d) in enumerate(pasos):
        x = 0.6 + i * (cw + gap)
        tarjeta(s, x, 1.85, cw, 2.45, fill=MIST)
        circulo_num(s, x + 0.25, 2.1, 0.6, n, size=16)
        texto(s, x + 1.0, 2.1, cw - 1.2, 0.6, t, size=18, bold=True, color=INK, font=HEAD,
              anchor=MSO_ANCHOR.MIDDLE)
        texto(s, x + 0.25, 2.9, cw - 0.5, 1.3, d, size=14, color=TEXT)
    stat(s, 0.6, 4.75, 3.8, "19,876,426", "filas leídas en total", color=MUTED, size=36)
    texto(s, 4.3, 4.85, 0.8, 0.7, "→", size=40, color=TEAL, bold=True, align=PP_ALIGN.CENTER)
    stat(s, 5.1, 4.75, 3.8, "1,276,897", "registros únicos (ID_REGISTRO)", color=TEAL, size=36)
    texto(s, 9.3, 4.8, 3.4, 1.4,
          "Se hace foto por foto porque unir ~20 millones de filas de una vez satura la memoria. "
          "Es el mismo criterio que el repositorio DengueMX.", size=13, color=MUTED)
    notas(s, "El algoritmo acumula foto por foto y en cada paso elimina duplicados conservando el "
             "último. Así, de casi 20 millones de filas leídas quedan 1.28 millones de registros únicos. "
             "También detectamos 148 IDs repetidos dentro de una misma foto.")


def slide_trazabilidad(prs):
    s = titulo_contenido(prs, "01 · Merge de datos", "Cada registro guarda su historial para decidir en la limpieza")
    items = [
        ("N_FOTOS", "En cuántas fotos semanales apareció el registro.", None, None),
        ("PRIMERA_FOTO / ULTIMA_FOTO", "Fecha de la primera y la última foto donde aparece.", None, None),
        ("DEFUNCION_CAMBIO", "Indica si DEFUNCION cambió en algún momento (por ejemplo, de No a Sí).",
         "2,494", "registros cambiaron"),
        ("EN_FOTO_FINAL", "Indica si el registro sigue en la última foto de su año. Si no, la DGE lo dio de baja.",
         "22,629", "registros dados de baja"),
    ]
    cw, chh = 5.9, 2.15
    for i, (t, d, num, lab) in enumerate(items):
        x = 0.6 + (i % 2) * (cw + 0.3)
        y = 1.85 + (i // 2) * (chh + 0.3)
        tarjeta(s, x, y, cw, chh, fill=MIST)
        tw = 3.4 if num else cw - 0.6
        texto(s, x + 0.3, y + 0.3, tw, 0.4, t, size=16, bold=True, color=INK)
        texto(s, x + 0.3, y + 0.8, tw, 1.2, d, size=14, color=TEXT)
        if num:
            texto(s, x + 3.9, y + 0.45, 1.8, 0.7, num, size=32, bold=True, color=RED if i == 2 else TEAL,
                  font=HEAD, align=PP_ALIGN.RIGHT)
            texto(s, x + 3.7, y + 1.2, 2.0, 0.6, lab, size=12, color=MUTED, align=PP_ALIGN.RIGHT)
    notas(s, "Además de la base, el merge genera columnas de trazabilidad. Las dos más importantes "
             "son DEFUNCION_CAMBIO y EN_FOTO_FINAL, porque con ellas decidimos qué registros quitar "
             "en la limpieza.")


def slide_resultado_merge(prs):
    s = titulo_contenido(prs, "01 · Merge de datos", "La base unificada cubre síntomas de 2020 a 2026")
    texto(s, 0.6, 1.9, 5.2, 0.35, "Normalización de tipos", size=14, bold=True, color=INK)
    texto(s, 0.6, 2.35, 5.2, 2.6, [
        {"text": "Fechas en dos formatos (AAAA-MM-DD y DD/MM/AAAA) se unifican.", "bullet": True, "after": 10},
        {"text": "9999-99-99 significa \"sin registro\" y se vuelve fecha vacía.", "bullet": True, "after": 10},
        {"text": "Las claves de catálogo pasan a enteros; todavía no se recodifica nada.", "bullet": True, "after": 10},
        {"text": "Se agrega ANIO_SINTOMAS a partir de la fecha de síntomas.", "bullet": True},
    ], size=15)
    tarjeta(s, 0.6, 5.0, 5.2, 1.5, fill=MIST)
    stat(s, 0.9, 5.15, 2.4, "36", "columnas finales", color=TEAL, size=30)
    stat(s, 2.9, 5.15, 2.8, "15.4 MB", "comprimido (.csv.gz) para GitHub", color=TEAL, size=30)
    texto(s, 6.5, 1.9, 6.2, 0.35, "Registros por año de inicio de síntomas", size=14, bold=True, color=INK)
    grafica_barras(s, 6.4, 2.3, 6.3, 4.2, ["2020", "2021", "2022", "2023", "2024", "2025", "2026"],
                   [120671, 37395, 61123, 281916, 555685, 146444, 73663], TEAL, formato="#,##0",
                   label_size=12, cat_size=13)
    notas(s, "Al final del merge tenemos 1,276,897 registros con 36 columnas: las 28 originales, "
             "el año de síntomas y las columnas de trazabilidad. 2024 fue el año con más registros. "
             "Ojo: aquí todavía hay casos probables y descartados, no solo confirmados.")


# ───────────────────────── 02 · Limpieza ─────────────────────────
def slide_div_limpieza(prs):
    s = divisor(prs, "02", "Limpieza", "De 1.28 millones de registros a 251,900 casos confirmados listos para modelar")
    notas(s, "Segunda etapa: notebook 02_limpieza.ipynb.")


def slide_embudo(prs):
    s = titulo_contenido(prs, "02 · Limpieza", "Tres filtros dejan 251,900 casos confirmados")
    filas = [
        ("Base del merge", None, 1276897, 5296),
        ("1 · Sigue en la foto final", "quita 22,629 (1.77%)", 1254268, 5207),
        ("2 · Solo confirmados", "quita 1,002,367 (79.92%)", 251901, 1763),
        ("3 · Entidad especificada", "quita 1 registro", 251900, 1763),
    ]
    texto(s, 4.05, 1.8, 3, 0.3, "Registros", size=12, bold=True, color=MUTED)
    texto(s, 10.75, 1.8, 1.95, 0.3, "Muertes", size=12, bold=True, color=MUTED, align=PP_ALIGN.RIGHT)
    maxw = 6.0
    for i, (lab, q, n, m) in enumerate(filas):
        y = 2.2 + i * 1.0
        texto(s, 0.6, y, 3.3, 0.45, lab, size=15, bold=True, color=INK)
        if q:
            texto(s, 0.6, y + 0.42, 3.3, 0.35, q, size=12, color=MUTED)
        bw = max(maxw * n / 1276897, 0.05)
        final = i == len(filas) - 1
        forma(s, MSO_SHAPE.RECTANGLE, 4.05, y + 0.05, bw, 0.6, TEAL if final else "9CC7BD")
        texto(s, 4.05 + bw + 0.15, y + 0.05, 2.2, 0.6, f"{n:,}", size=16, bold=True,
              color=TEAL if final else TEXT, anchor=MSO_ANCHOR.MIDDLE)
        texto(s, 10.75, y + 0.05, 1.95, 0.6, f"{m:,}", size=16, bold=True, color=RED,
              anchor=MSO_ANCHOR.MIDDLE, align=PP_ALIGN.RIGHT)
    tarjeta(s, 0.6, 6.25, 12.1, 0.6, fill=MIST)
    texto(s, 0.85, 6.25, 11.6, 0.6,
          [{"runs": [("En total se quita el 80.27% de la base. ", {"bold": True}),
                     ("Casi todo corresponde a casos probables (719,278) y descartados (283,089). "
                      "Mortalidad final: 0.70%.", {})]}],
          size=14, color=TEXT, anchor=MSO_ANCHOR.MIDDLE)
    notas(s, "Paso 1: quitamos registros que la DGE dio de baja (no están en la foto final de su año). "
             "Paso 2: el filtro más fuerte; nos quedamos solo con ESTATUS_CASO = 2 (confirmado). "
             "Paso 3: un solo registro con entidad de asignación 99 (no especificada). "
             "Quedan 251,900 casos con 1,763 muertes.")


def slide_decision(prs):
    s = titulo_contenido(prs, "02 · Limpieza · Decisión del equipo",
                         "Conservamos los cambios en DEFUNCION para no sesgar el modelo")
    texto(s, 0.6, 1.85, 12.1, 0.8,
          [{"runs": [("2,494 registros cambiaron su valor de DEFUNCION. ", {"bold": True}),
                     ("De ellos, 2,433 pasaron de \"No\" a \"Sí\": la muerte se registró después. "
                      "Su valor final es el correcto.", {})]}], size=16)
    # Opción elegida
    tarjeta(s, 0.6, 2.85, 5.9, 3.0, fill=MIST, nombre="Opción elegida")
    texto(s, 0.9, 3.1, 5.3, 0.3, "OPCIÓN ELEGIDA · SE CONSERVAN", size=12, bold=True, color=TEAL)
    stat(s, 0.9, 3.5, 2.6, "251,900", "casos confirmados", color=INK, size=34)
    stat(s, 3.6, 3.5, 2.6, "1,763", "muertes (0.70%)", color=RED, size=34)
    texto(s, 0.9, 4.85, 5.3, 0.9, "El modelo ve todas las muertes, se hayan reportado rápido o tarde.",
          size=14, color=TEXT)
    # Alternativa
    tarjeta(s, 6.8, 2.85, 5.9, 3.0, fill=PAPER, line=LINE, nombre="Alternativa")
    texto(s, 7.1, 3.1, 5.3, 0.3, "ALTERNATIVA DESCARTADA · QUITARLOS", size=12, bold=True, color=MUTED)
    stat(s, 7.1, 3.5, 2.6, "251,077", "casos confirmados", color=MUTED, size=34)
    stat(s, 9.8, 3.5, 2.6, "962", "muertes", color=MUTED, size=34)
    texto(s, 7.1, 4.85, 5.3, 0.9, "Se perderían 801 de 1,770 muertes confirmadas (45%).",
          size=14, color=TEXT)
    texto(s, 0.6, 6.15, 12.1, 0.6,
          [{"runs": [("Riesgo evitado: ", {"bold": True, "color": RED}),
                     ("sesgo de selección; el modelo solo aprendería de las muertes que se reportaron rápido.", {})]}],
          size=15)
    notas(s, "Esta fue la decisión más importante de la limpieza. Si quitábamos los registros cuyo "
             "valor de DEFUNCION cambió, perdíamos casi la mitad de las muertes, y justamente las que "
             "se registraron tarde. Por eso solo filtramos por foto final y conservamos el valor más reciente.")


def slide_variables(prs):
    s = titulo_contenido(prs, "02 · Limpieza", "Seleccionamos 17 de las 35 columnas")
    conservan = ["ID_REGISTRO", "DEFUNCION", "SEXO", "EDAD_ANOS", "FECHA_SIGN_SINTOMAS",
                 "ANIO_SINTOMAS", "ENTIDAD_ASIG", "INDIGENA", "HEMORRAGICOS",
                 "DIABETES", "HIPERTENSION", "ENFERMEDAD_RENAL", "INMUNOSUPR",
                 "CIRROSIS_HEPATICA", "EMBARAZO", "TIPO_PACIENTE", "RESULTADO_PCR"]
    excluidas = ["DICTAMEN", "ENFERMEDAD_ULC_PEPTICA", "HABLA_LENGUA_INDIG", "TOMA_MUESTRA",
                 "ENTIDAD_RES", "MUNICIPIO_RES", "ENTIDAD_UM_NOTIF", "MUNICIPIO_UM_NOTIF",
                 "MUNICIPIO_ASIG", "INSTITUCION_UM_NOTIF", "FECHA_ACTUALIZACION", "ARCHIVO_FUENTE",
                 "FECHA_ARCHIVO", "N_FOTOS", "PRIMERA_FOTO", "ULTIMA_FOTO",
                 "DEFUNCION_CAMBIO", "EN_FOTO_FINAL"]

    def lista(x, titulo, color_titulo, items, bullet_color, fill, line):
        tarjeta(s, x, 1.85, 5.9, 4.95, fill=fill, line=line)
        texto(s, x + 0.3, 2.05, 5.3, 0.3, titulo, size=12, bold=True, color=color_titulo)
        mitad = (len(items) + 1) // 2
        for c, grupo in enumerate([items[:mitad], items[mitad:]]):
            texto(s, x + 0.3 + c * 2.75, 2.5, 2.7, 4.2,
                  [{"text": v, "bullet": True, "bullet_color": bullet_color, "after": 5} for v in grupo],
                  size=13)

    lista(0.6, f"SE CONSERVAN ({len(conservan)})", TEAL, conservan, TEAL, MIST, None)
    lista(6.8, f"NO ENTRAN AL MODELO ({len(excluidas)})", MUTED, excluidas, MUTED, PAPER, LINE)
    notas(s, "Pasamos de 35 a 17 columnas. Lo más importante es lo que dejamos fuera a propósito: "
             "DICTAMEN es consecuencia de la muerte, no una causa, así que incluirlo haría trampa al modelo.")


def slide_recodificacion(prs):
    s = titulo_contenido(prs, "02 · Limpieza", "Recodificamos las variables a 1/0 para la regresión logística")
    texto(s, 0.6, 1.85, 5.6, 0.35, "Reglas de recodificación", size=14, bold=True, color=INK)
    reglas = [
        ("Catálogo Sí/No", "1 = Sí, 2 = No  →  1 = Sí, 0 = No"),
        ("SEXO → HOMBRE", "1 = hombre, 0 = mujer"),
        ("TIPO_PACIENTE → HOSPITALIZADO", "1 = hospitalizado; 9 vacíos se asignan a ambulatorio (0)"),
        ("ENTIDAD_ASIG → COSTERA", "1 = una de las 17 entidades con costa"),
    ]
    for i, (t, d) in enumerate(reglas):
        y = 2.3 + i * 0.95
        tarjeta(s, 0.6, y, 5.6, 0.82, fill=MIST)
        texto(s, 0.85, y + 0.1, 5.1, 0.35, t, size=14, bold=True, color=INK)
        texto(s, 0.85, y + 0.43, 5.1, 0.35, d, size=13, color=TEXT)
    texto(s, 0.6, 6.2, 5.6, 0.5, "Ninguna variable tiene valores fuera de {1, 2}: no se quita ninguna fila.",
          size=13, color=MUTED)

    texto(s, 6.7, 1.85, 6.0, 0.35, "Proporción de pacientes con cada condición", size=14, bold=True, color=INK)
    for i, (num, lab) in enumerate([("70.1%", "costera"), ("44.8%", "hombres"), ("38.9%", "hospitalizados")]):
        stat(s, 6.7 + i * 2.05, 2.25, 2.0, num, lab, color=TEAL, size=28)
    grafica_barras(s, 6.6, 3.45, 6.1, 3.35,
                   ["Indígena", "Diabetes", "Embarazo*", "Hipertensión", "Hemorrágicos",
                    "Enf. renal", "Inmunosupresión", "Cirrosis"],
                   [0.03314, 0.02879, 0.02621, 0.02349, 0.00908, 0.00251, 0.00147, 0.00093],
                   TEAL, formato="0.00%", horizontal=True, label_size=11, cat_size=12, gap=40, max_val=0.042)
    texto(s, 6.7, 6.8, 6.0, 0.25, "* Embarazo vale 0 en todos los hombres.", size=10, color=MUTED)
    notas(s, "El modelo necesita variables 0/1, pero el catálogo oficial usa 1 = Sí y 2 = No. "
             "Las comorbilidades son poco frecuentes: la más común, ser indígena, es apenas 3.3%. "
             "Esto importa para la estabilidad de los coeficientes.")


def slide_serotipo(prs):
    s = titulo_contenido(prs, "02 · Limpieza · Variables dummy",
                         "Serotipo: DENV1 es la referencia y tiene la menor mortalidad")
    texto(s, 0.6, 1.85, 6.8, 0.35, "Mortalidad por serotipo (resultado de PCR)", size=14, bold=True, color=INK)
    grafica_barras(s, 0.5, 2.25, 7.0, 4.0,
                   ["DENV1\n13,531", "DENV2\n23,139", "DENV3\n71,812", "DENV4\n1,396", "Sin serotipo\n142,022"],
                   [0.0029, 0.0066, 0.0067, 0.0086, 0.0076], [AMBER, RED, RED, RED, RED],
                   formato="0.00%", label_size=14, cat_size=12, max_val=0.0105)
    texto(s, 0.6, 6.3, 6.8, 0.3, "Debajo de cada barra: número de casos confirmados.", size=11, color=MUTED)

    texto(s, 7.9, 1.85, 4.8, 4.8, [
        {"runs": [("5 categorías → 4 dummies ", {"bold": True, "color": INK}),
                  ("(DENV2, DENV3, DENV4, SIN_SEROTIPO).", {})], "after": 6},
        {"text": "Incluir las 5 columnas las haría colineales con el intercepto (trampa de las dummies).",
         "size": 14, "color": MUTED, "after": 18},
        {"runs": [("DENV1 es la referencia: ", {"bold": True, "color": AMBER_TXT}),
                  ("cada coeficiente se lee comparado con DENV1.", {})], "after": 18},
        {"runs": [("\"Sin serotipo\" es el 56.4% de los confirmados. ", {"bold": True, "color": INK}),
                  ("Es un serotipo desconocido, no un quinto virus.", {})], "after": 6},
        {"text": "Causas: confirmación por NS1 o IgM, muestra tomada tarde (más de ~5 días) "
                 "o confirmación por asociación epidemiológica.", "size": 14, "color": MUTED},
    ], size=15)
    notas(s, "El resultado de PCR tiene cinco categorías, así que creamos cuatro variables dummy y "
             "dejamos DENV1 como referencia. Más de la mitad de los confirmados no tienen serotipo "
             "identificado, pero siguen siendo casos de dengue confirmados.")


def slide_mes(prs):
    s = titulo_contenido(prs, "02 · Limpieza · Variables dummy",
                         "El mes de síntomas controla la estacionalidad")
    texto(s, 0.6, 1.85, 7.4, 0.35, "Porcentaje de casos confirmados por mes", size=14, bold=True, color=INK)
    meses = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"]
    casos = [0.036, 0.022, 0.020, 0.020, 0.028, 0.052, 0.080, 0.158, 0.211, 0.217, 0.133, 0.024]
    col = ["9CC7BD"] * 5 + [TEAL] * 4 + [AMBER] + ["9CC7BD"] * 2
    grafica_barras(s, 0.5, 2.25, 7.5, 3.9, meses, casos, col, formato="0%", label_size=11, cat_size=12, gap=40)
    texto(s, 0.6, 6.2, 7.4, 0.5,
          [{"runs": [("Lluvias (jun–oct) ", {"bold": True, "color": TEAL}), ("concentran el 72% de los casos. ", {}),
                     ("Octubre ", {"bold": True, "color": AMBER_TXT}), ("es la referencia.", {})]}],
          size=13, color=MUTED)

    tarjeta(s, 8.4, 1.85, 4.3, 2.35, fill=MIST)
    texto(s, 8.7, 2.05, 3.8, 0.3, "MORTALIDAD", size=12, bold=True, color=RED)
    stat(s, 8.7, 2.4, 1.9, "0.68%", "en lluvias", color=RED, size=28)
    stat(s, 10.7, 2.4, 1.9, "0.76%", "en secas", color=RED, size=28)
    texto(s, 8.7, 3.6, 3.8, 0.5, "En temporada alta se confirman muchos casos leves.", size=12, color=TEXT)

    tarjeta(s, 8.4, 4.4, 4.3, 2.3, fill=MIST)
    texto(s, 8.7, 4.6, 3.8, 0.3, "HOSPITALIZACIÓN", size=12, bold=True, color=TEAL)
    stat(s, 8.7, 4.95, 3.8, "~48%", "en enero–marzo, contra ~37% en temporada alta", color=TEAL, size=28)
    notas(s, "El mes de inicio de síntomas entra como 11 dummies con octubre como referencia, por ser "
             "el mes con más casos (21.7%) y dar coeficientes más estables. En secas se estudian sobre "
             "todo los casos graves, por eso la mortalidad y la hospitalización son más altas: el mes "
             "capta ese cambio en el perfil de los casos.")


def slide_anio(prs):
    s = titulo_contenido(prs, "02 · Limpieza · Decisión del equipo",
                         "El año se excluye: su mortalidad no muestra tendencia")
    texto(s, 0.6, 1.85, 7.4, 0.35, "Mortalidad de casos confirmados por año", size=14, bold=True, color=INK)
    grafica_barras(s, 0.5, 2.25, 7.5, 4.0,
                   ["2020\n24,308", "2021\n6,553", "2022\n12,327", "2023\n54,197",
                    "2024\n123,141", "2025\n21,918", "2026\n9,456"],
                   [0.0058, 0.0076, 0.0066, 0.0073, 0.0069, 0.0086, 0.0058], RED,
                   formato="0.00%", label_size=13, cat_size=12, max_val=0.0105)
    texto(s, 0.6, 6.3, 7.4, 0.3, "Debajo de cada barra: número de casos confirmados.", size=11, color=MUTED)

    texto(s, 8.4, 1.85, 4.3, 0.35, "Por qué se excluye", size=14, bold=True, color=INK)
    texto(s, 8.4, 2.3, 4.3, 2.0, [
        {"text": "La mortalidad oscila entre 0.58% y 0.86% sin una dirección clara.", "bullet": True, "after": 10},
        {"text": "Los avances médicos entre años no implican necesariamente menos muertes.", "bullet": True},
    ], size=15)
    tarjeta(s, 8.4, 4.45, 4.3, 1.85, fill="FBF1DE")
    texto(s, 8.7, 4.65, 3.8, 0.3, "LIMITACIÓN", size=12, bold=True, color=AMBER_TXT)
    texto(s, 8.7, 5.0, 3.8, 1.2,
          "El serotipo dominante cambió con los años (DENV2 en 2020–2021, DENV3 desde 2023): "
          "sin el año, el serotipo absorbe parte del efecto de la época.", size=13, color=TEXT)
    notas(s, "Decidimos no incluir el año. Lo dejamos explícito como limitación porque el serotipo "
             "dominante cambió con el tiempo.")


def slide_base_final(prs):
    s = titulo_contenido(prs, "02 · Limpieza · Resultado", "La base final queda lista para modelar")
    stats = [("251,900", "pacientes confirmados", TEAL), ("29", "columnas", TEAL),
             ("1,763", "muertes (0.70%)", RED), ("0", "valores faltantes", TEAL)]
    for i, (n, l, c) in enumerate(stats):
        x = 0.6 + i * 3.1
        tarjeta(s, x, 1.85, 2.8, 1.55, fill=MIST)
        stat(s, x + 0.3, 2.0, 2.3, n, l, color=c, size=32)

    texto(s, 0.6, 3.75, 7.0, 0.35, "Composición de las 29 columnas", size=14, bold=True, color=INK)
    grupos = [("1", "Identificador", "ID_REGISTRO"), ("1", "Objetivo", "DEFUNCION"),
              ("1", "Numérica", "EDAD_ANOS (media 28.5, mediana 25, rango 0–112)"),
              ("11", "Binarias", "HOMBRE, COSTERA, INDIGENA, HEMORRAGICOS, comorbilidades, HOSPITALIZADO"),
              ("4", "Serotipo", "DENV2, DENV3, DENV4, SIN_SEROTIPO (ref. DENV1)"),
              ("11", "Mes", "MES_ENE … MES_DIC (ref. octubre)")]
    for i, (n, g, d) in enumerate(grupos):
        y = 4.2 + i * 0.44
        texto(s, 0.6, y, 0.55, 0.4, n, size=16, bold=True, color=TEAL, align=PP_ALIGN.RIGHT)
        texto(s, 1.35, y, 1.6, 0.4, g, size=14, bold=True, color=INK)
        texto(s, 3.0, y, 5.0, 0.4, d, size=13, color=TEXT)

    tarjeta(s, 8.4, 3.75, 4.3, 2.95, fill="F9E9EC")
    texto(s, 8.7, 3.95, 3.8, 0.3, "PARA EL MODELO", size=12, bold=True, color=RED)
    texto(s, 8.7, 4.35, 3.8, 2.3, [
        {"runs": [("Clases muy desbalanceadas: ", {"bold": True}), ("solo 7 de cada 1,000 pacientes murieron.", {})],
         "after": 10},
        {"text": "Hay que tenerlo en cuenta al evaluar el modelo: la exactitud por sí sola engaña.", "color": MUTED},
    ], size=14)
    notas(s, "Resultado: dengue_limpio.csv, 251,900 filas por 29 columnas, sin faltantes. "
             "El punto clave para la siguiente etapa es el desbalance: con 0.7% de muertes, un modelo "
             "que dijera 'nadie muere' tendría 99.3% de exactitud.")


# ───────────────────────── Cierre ─────────────────────────
def slide_siguientes(prs):
    s = nueva(prs)
    fondo(s, INK)
    forma(s, MSO_SHAPE.OVAL, 9.2, 3.6, 5.6, 5.6, INK_2)
    texto(s, 0.9, 0.9, 8, 0.4, "SIGUIENTES PASOS", size=12, bold=True, color=AMBER)
    texto(s, 0.9, 1.35, 11, 1.0, "Con la base limpia, sigue explorar y modelar", size=40, bold=True,
          color=PAPER, font=HEAD, nombre="Título")
    pasos = [("03", "Análisis exploratorio", "Mortalidad por edad, comorbilidades, serotipo, mes y región."),
             ("04", "Modelo de regresión logística", "Estimación, razones de momios e interpretación de coeficientes.")]
    for i, (n, t, d) in enumerate(pasos):
        y = 3.0 + i * 1.55
        circulo_num(s, 0.9, y, 0.9, n, fill=AMBER, color=INK, size=20)
        texto(s, 2.1, y - 0.02, 8, 0.5, t, size=22, bold=True, color=PAPER, font=HEAD)
        texto(s, 2.1, y + 0.5, 8, 0.5, d, size=16, color=LIGHT_ON_DARK)
    pie(s, prs, oscuro=True)
    notas(s, "Estas dos etapas se agregarán en las siguientes secciones de la presentación.")


# ───────────────────────── Secciones ─────────────────────────
# Para agregar EDA o modelo: añade aquí ("03 · EDA", [slide_div_eda, ...]) antes de "Siguientes pasos".
SECCIONES = [
    ("Introducción", [slide_portada, slide_ruta, slide_pregunta]),
    ("01 · Merge de datos", [slide_div_merge, slide_fuente, slide_estrategia, slide_trazabilidad,
                             slide_resultado_merge]),
    ("02 · Limpieza", [slide_div_limpieza, slide_embudo, slide_decision, slide_variables,
                       slide_recodificacion, slide_serotipo, slide_mes, slide_anio, slide_base_final]),
    ("Siguientes pasos", [slide_siguientes]),
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
    prs.slide_width, prs.slide_height = Inches(W), Inches(H)
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
