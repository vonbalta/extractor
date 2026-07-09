"""
pdf_extractor.py — Extracción de datos desde informes PDF de esterilizado.

Responsabilidades
-----------------
- Parsear metadatos del informe (lote, autoclave, fecha, hora).
- Extraer la tabla de ciclo (INSTANT, FASE, temperaturas, caudal, etc.).
- Escanear una carpeta de PDFs y devolver solo los archivos válidos.

No tiene dependencias de Excel ni de interfaz gráfica.
"""

import os
import re
import logging
from datetime import datetime

import pdfplumber

logger = logging.getLogger(__name__)
logger.addHandler(logging.NullHandler())


# ---------------------------------------------------------------------------
#  CONSTANTES PÚBLICAS
# ---------------------------------------------------------------------------

# Marcas especiales que van al final de las hojas del Excel.
# Si el nombre del PDF contiene alguna de estas palabras, la hoja se
# coloca después de todas las hojas regulares.
MARCAS_AL_FINAL = ("bonamark", "chedraui", "siprecio", "kirkland", "nair")

# Palabras de descarte por defecto: si el nombre de un PDF contiene alguna
# de estas subcadenas (sin distinguir mayusculas/minusculas), el archivo se
# omite del escaneo. Es el valor usado cuando escanear_carpeta_pdfs() se
# llama sin palabras_descarte (compatibilidad con el comportamiento
# anterior, en el que estas tres palabras estaban fijas en el codigo).
PALABRAS_DESCARTE_DEFAULT = ("prueba", "limpieza", "xxx")


# ---------------------------------------------------------------------------
#  EXPRESIONES REGULARES (compiladas a nivel de módulo)
# ---------------------------------------------------------------------------

# Patrón "oro": extrae lote, autoclave, fecha y hora del encabezado del informe.
_PATRON_ORO = re.compile(
    r'(?:MPI)(\d+)[_\s]+(\d+)[-\s]+(?:.*?)\[(\d{2}/\d{2}/\d{4})\s+(\d{2}:\d{2}:\d{2})\]',
    re.IGNORECASE,
)

# Patrón de hora "HH:MM:SS" para identificar filas de datos en el PDF.
_PATRON_HORA = re.compile(r'^\d{2}:\d{2}:\d{2}$')


# ---------------------------------------------------------------------------
#  API PÚBLICA
# ---------------------------------------------------------------------------

def extraer_metadatos_informe(ruta_pdf):
    """
    Extrae los metadatos del informe (autoclave, fecha, hora, minutos_totales)
    del PDF indicado.

    Como efecto secundario, cachea el texto completo extraído en el dict
    devuelto bajo la clave '_texto_completo' para que extraer_tabla_datos
    pueda reutilizarlo sin abrir el PDF otra vez.

    Devuelve un dict con las claves:
        fecha_dt, autoclave, fecha_display, hora_display, minutos_totales
    o None si el PDF no contiene el patrón esperado o falla la lectura.
    """
    metadatos = {
        'fecha_dt':        None,
        'autoclave':       'X',
        'fecha_display':   '',
        'hora_display':    '',
        'minutos_totales': 0,
    }
    try:
        texto_completo = ""
        with pdfplumber.open(ruta_pdf) as pdf:
            for pagina in pdf.pages:
                txt = pagina.extract_text()
                if txt:
                    texto_completo += txt + "\n"

        match_lote = _PATRON_ORO.search(texto_completo)

        if match_lote:
            metadatos['autoclave']       = str(int(match_lote.group(2)))
            fecha_str                    = match_lote.group(3)
            hora_str                     = match_lote.group(4)
            metadatos['fecha_dt']        = datetime.strptime(
                f"{fecha_str} {hora_str}", "%d/%m/%Y %H:%M:%S"
            )
            metadatos['fecha_display']   = fecha_str
            metadatos['hora_display']    = hora_str
            h, m, _s                     = map(int, hora_str.split(':'))
            metadatos['minutos_totales'] = (h * 60) + m
            metadatos['_texto_completo'] = texto_completo
            return metadatos

    except Exception as e:
        logger.exception("Error al extraer metadatos de %s: %s", ruta_pdf, e)

    return None


def extraer_tabla_datos(ruta_pdf, texto_completo=None):
    """
    Extrae las filas de la sección 'LISTA DE DATOS' del PDF.

    Si se proporciona ``texto_completo`` (el texto extraído previamente por
    extraer_metadatos_informe), se evita abrir el PDF otra vez.

    Devuelve una lista de listas de strings (una sublista por fila válida).
    Cada fila válida empieza con una hora 'HH:MM:SS' y tiene ≥ 12 columnas.
    """
    datos_tabla = []

    if texto_completo is not None:
        paginas_texto = [texto_completo]
    else:
        with pdfplumber.open(ruta_pdf) as pdf:
            paginas_texto = [p.extract_text() or "" for p in pdf.pages]

    for texto in paginas_texto:
        if "LISTA DE DATOS" in texto:
            for linea in texto.split('\n'):
                valores = linea.strip().split()
                if len(valores) >= 12 and _PATRON_HORA.match(valores[0]):
                    if valores[1].isdigit() and 1 <= int(valores[1]) <= 8:
                        datos_tabla.append(valores)
    return datos_tabla


def escanear_carpeta_pdfs(carpeta_pdfs, evento_cancelacion=None,
                           verificar_cancelacion_fn=None,
                           palabras_descarte=None):
    """
    Escanea `carpeta_pdfs`, extrae metadatos de cada PDF válido y devuelve
    una lista de dicts (uno por archivo procesable).

    Los PDFs cuyo nombre contiene alguna de las `palabras_descarte` (sin
    distinguir mayúsculas/minúsculas) se omiten.

    Parámetros
    ----------
    carpeta_pdfs             : ruta de la carpeta a escanear.
    evento_cancelacion       : threading.Event opcional para cancelación
                               cooperativa.
    verificar_cancelacion_fn : callable(evento) que lanza OperacionCancelada
                               si el evento está activo. Se suministra desde
                               el orquestador (logica.py) para evitar imports
                               circulares.
    palabras_descarte        : iterable de palabras/subcadenas a descartar
                               por nombre de archivo. Si es None, se usa
                               PALABRAS_DESCARTE_DEFAULT ('prueba',
                               'limpieza', 'xxx' — el comportamiento
                               histórico). Pasar una lista/tupla vacía
                               desactiva el filtro por completo (ningún PDF
                               se descarta por nombre).
    """
    if palabras_descarte is None:
        palabras_descarte = PALABRAS_DESCARTE_DEFAULT
    palabras_descarte = [str(p).strip().lower() for p in palabras_descarte if str(p).strip()]

    rutas_pdfs = [
        os.path.join(carpeta_pdfs, f)
        for f in os.listdir(carpeta_pdfs)
        if f.lower().endswith('.pdf')
    ]
    archivos_validos = []

    for ruta in rutas_pdfs:
        if verificar_cancelacion_fn:
            verificar_cancelacion_fn(evento_cancelacion)

        nombre = os.path.basename(ruta)
        if palabras_descarte and any(x in nombre.lower() for x in palabras_descarte):
            logger.info(f"Omitido: {nombre}")
            continue

        meta = extraer_metadatos_informe(ruta)
        if meta:
            meta['ruta'] = ruta
            meta['tipo'] = 'candidato'
            nombre_lower    = nombre.lower()
            marca_detectada = next(
                (m for m in MARCAS_AL_FINAL if m in nombre_lower), None
            )
            meta['marca'] = marca_detectada
            archivos_validos.append(meta)
        else:
            logger.warning(f"Sin metadatos válidos: {nombre}")

    return archivos_validos