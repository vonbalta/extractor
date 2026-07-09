"""
excel_manager.py — Gestión de hojas Excel para el Procesador de Esterilizado.

Responsabilidades
-----------------
- Escribir datos de un ciclo en una hoja xlwings (plantilla).
- Leer metadatos de hojas existentes (fecha, hora, autoclave, marca).
- Autoajustar el alto de filas de la hoja 'Desviaciones' vía COM.
- Helpers de archivos: backup con marca de tiempo y limpieza del .bak de Excel.
- Nomenclatura canónica de hojas: "3(A2)" y "3(A2) BONAMARK".

Exporta las expresiones regulares de nombres de hoja (PATRON_HOJA_NUM,
PATRON_HOJA_DATOS) porque otros módulos (report_builder, logica) las necesitan.

Depende de: pdf_extractor (para rellenar la tabla de mediciones).
No tiene dependencias de tkinter ni de la interfaz gráfica.
"""

import os
import re
import shutil
import uuid
import logging
from datetime import datetime, date, time as dtime

import xlwings as xw

from pdf_extractor import extraer_tabla_datos

logger = logging.getLogger(__name__)
logger.addHandler(logging.NullHandler())


# ---------------------------------------------------------------------------
#  EXPRESIONES REGULARES PÚBLICAS
# ---------------------------------------------------------------------------

# Nombre de hoja de datos: "12(A3)" o "12(A3) BONAMARK".
PATRON_HOJA_DATOS = re.compile(r'^\d+\(A\d+\)(\s+\S+)?$')

# Captura: grupo 1 = número de orden, grupo 2 = número de autoclave,
# grupo 3 (opcional) = nombre de marca.
PATRON_HOJA_NUM = re.compile(r'^(\d+)\(A(\d+)\)(?:\s+(\S+))?$')


# ---------------------------------------------------------------------------
#  HELPERS DE ARCHIVO
# ---------------------------------------------------------------------------

def eliminar_bak(ruta_archivo):
    """
    Elimina el .bak que Excel genera automáticamente al guardar
    (mismo nombre que `ruta_archivo` con extensión .bak).
    No lanza excepción si el archivo no existe o no se puede borrar.
    """
    ruta_bak = os.path.splitext(ruta_archivo)[0] + ".bak"
    try:
        os.remove(ruta_bak)
    except OSError:
        pass


def crear_backup_si_existe(ruta_archivo):
    """
    Si `ruta_archivo` ya existe, crea una copia de resguardo en el mismo
    directorio con sufijo de marca de tiempo (p. ej. '_backup_20260629_153012').

    Devuelve la ruta del backup creado, o None si no había nada que respaldar
    o si la copia falló. No lanza excepción: la ausencia de backup nunca debe
    impedir que el procesamiento continúe.
    """
    if not os.path.exists(ruta_archivo):
        return None
    base, ext = os.path.splitext(ruta_archivo)
    marca      = datetime.now().strftime("%Y%m%d_%H%M%S")
    ruta_bk    = f"{base}_backup_{marca}{ext}"
    try:
        shutil.copy2(ruta_archivo, ruta_bk)
        logger.info(f"  Copia de seguridad del archivo anterior: {ruta_bk}")
        return ruta_bk
    except OSError as e:
        logger.warning(f"  No se pudo crear copia de seguridad de '{ruta_archivo}': {e}")
        return None


def asegurar_archivo_libre(ruta):
    """
    Intenta abrir el archivo en modo lectura/escritura para comprobar
    que no está bloqueado por otro proceso (ej. Excel).
    Lanza PermissionError con mensaje claro si está bloqueado.
    """
    if not os.path.exists(ruta):
        return  # No existe → no hay problema
    try:
        with open(ruta, 'r+b') as f:
            pass
    except PermissionError:
        raise PermissionError(
            f"El archivo '{os.path.basename(ruta)}' está abierto en otro programa.\n"
            "Ciérralo antes de continuar."
        )


# ---------------------------------------------------------------------------
#  LIMPIEZA DE VÍNCULOS EXTERNOS FANTASMA
# ---------------------------------------------------------------------------

def romper_vinculos_externos(wb):
    """
    Rompe cualquier vínculo a libro externo que Excel crea implícitamente
    al copiar hojas entre workbooks distintos (xlwings/COM).

    Sheets.Copy entre dos objetos Workbook reescribe como referencia externa
    ('[archivo.xlsx]Hoja'!A1) cualquier fórmula de la hoja copiada que apunte
    a otra hoja del libro de origen. Como esos libros de origen suelen ser
    temporales y se borran al terminar, Excel se queda con un vínculo que ya
    no puede resolver, y el archivo final muestra "No pudimos obtener
    valores actualizados de un libro vinculado" al abrirse.

    Llamar siempre sobre `wb` justo antes de wb.save(), en cualquier flujo
    que haga copy() entre dos Workbooks diferentes. No lanza excepción.
    """
    try:
        fuentes = wb.api.LinkSources(1)  # xlExcelLinks = 1
    except Exception:
        fuentes = None

    if not fuentes:
        return

    for origen in fuentes:
        try:
            wb.api.BreakLink(Name=origen, Type=1)
            logger.info(f"  Vínculo externo roto: {origen}")
        except Exception as e:
            logger.warning(f"  No se pudo romper el vínculo '{origen}': {e}")


# ---------------------------------------------------------------------------
#  NOMENCLATURA DE HOJAS
# ---------------------------------------------------------------------------

def nombre_hoja(numero, autoclave, marca):
    """
    Genera el nombre canónico de hoja para un ciclo.

    Ejemplos:
        nombre_hoja(3, "2", None)        → "3(A2)"
        nombre_hoja(3, "2", "bonamark")  → "3(A2) BONAMARK"

    Trunca a 31 caracteres (límite de Excel).
    """
    base = f"{numero}(A{autoclave})"
    if marca:
        base = f"{base} {marca.upper()}"
    return base[:31]


# ---------------------------------------------------------------------------
#  ESCRITURA DE HOJAS (xlwings)
# ---------------------------------------------------------------------------

def llenar_hoja_con_datos(ws, arch):
    """
    Escribe los datos de un ciclo en la hoja xlwings `ws`.

    Rellena:
        C2 — fecha (como texto 'DD/MM/YYYY')
        C3 — número de autoclave
        C4 — hora de inicio (fracción de día, formato hh:mm:ss)
        A7 en adelante — matriz de mediciones extraída del PDF

    `arch` debe contener: fecha_display, autoclave, hora_display, ruta.
    """
    ws.range('C2').api.NumberFormat = '@'
    ws.range('C2').api.Value2       = arch['fecha_display']
    ws.range('C3').value            = int(arch['autoclave'])

    h, m, s      = map(int, arch['hora_display'].split(':'))
    fraccion_hora = (h * 3600 + m * 60 + s) / 86400
    ws.range('C4').value         = fraccion_hora
    ws.range('C4').number_format = 'hh:mm:ss'

    datos = extraer_tabla_datos(arch['ruta'], arch.get('_texto_completo'))
    if datos:
        matriz_excel = [
            [
                d[0],   # A: INSTANT
                d[1],   # B: FASE
                d[2],   # C: T°C SET
                d[3],   # D: CTRL MEASURE
                d[4],   # E: BAR SET
                d[5],   # F: REGIST Bar
                d[6],   # G: REGIST °C
                None,   # H: (vacío — columna oculta)
                d[8],   # I: NIVEL
                None,   # J: (vacío)
                None,   # K: (vacío)
                d[11],  # L: CAUDAL
            ]
            for d in datos
        ]
        ws.range('A7').value = matriz_excel


# ---------------------------------------------------------------------------
#  LECTURA DE HOJAS (xlwings)
# ---------------------------------------------------------------------------

def leer_metadatos_hoja(ws):
    """
    Lee fecha (C2), autoclave (C3) y hora (C4) de la hoja xlwings `ws`.

    Maneja los distintos tipos que xlwings puede devolver para cada celda
    (datetime, date, float fracción-de-día, str) y normaliza a objetos Python.

    Devuelve un dict con:
        fecha_dt, autoclave, fecha_display, hora_display, hoja_obj, tipo, marca
    o None si la lectura falla o los datos son incoherentes.
    """
    try:
        valor_fecha     = ws.range('C2').value
        valor_autoclave = ws.range('C3').value
        valor_hora      = ws.range('C4').value

        # — Fecha —
        if isinstance(valor_fecha, datetime):
            fecha_obj = valor_fecha.date()
        elif isinstance(valor_fecha, date):
            fecha_obj = valor_fecha
        elif isinstance(valor_fecha, str):
            fecha_obj = datetime.strptime(valor_fecha.strip(), "%d/%m/%Y").date()
        else:
            return None

        # — Hora —
        if isinstance(valor_hora, float):
            total_seg = round(valor_hora * 86400)
            hora_obj  = dtime(
                total_seg // 3600,
                (total_seg % 3600) // 60,
                total_seg % 60,
            )
        elif isinstance(valor_hora, datetime):
            hora_obj = valor_hora.time()
        elif isinstance(valor_hora, dtime):
            hora_obj = valor_hora
        elif isinstance(valor_hora, str):
            hora_obj = datetime.strptime(valor_hora.strip(), "%H:%M:%S").time()
        else:
            return None

        autoclave = str(int(valor_autoclave))
        fecha_dt  = datetime.combine(fecha_obj, hora_obj)

        # La marca real está en el grupo 3 del nombre de hoja ("3(A2) BONAMARK"),
        # nunca en el grupo 2 (número de autoclave).
        match_marca = PATRON_HOJA_NUM.match(ws.name)
        marca = (
            match_marca.group(3).lower()
            if (match_marca and match_marca.group(3))
            else None
        )

        return {
            'fecha_dt':      fecha_dt,
            'autoclave':     autoclave,
            'fecha_display': fecha_obj.strftime("%d/%m/%Y"),
            'hora_display':  hora_obj.strftime("%H:%M:%S"),
            'hoja_obj':      ws,
            'tipo':          'existente',
            'marca':         marca,
        }
    except Exception as e:
        logger.exception("Error al leer metadatos de la hoja '%s': %s", ws.name, e)
        return None


# ---------------------------------------------------------------------------
#  AUTOAJUSTE DE FILAS (xlwings / COM)
# ---------------------------------------------------------------------------

def autoajustar_filas(ruta_archivo):
    """
    Abre `ruta_archivo` con Excel (vía xlwings/COM) y aplica AutoFit nativo
    sobre las filas de la hoja 'Desviaciones', omitiendo las que tienen celdas
    combinadas (Excel las trata como vacías y calcularía el alto en 0).

    Trabaja siempre sobre una copia con nombre ASCII puro para esquivar el
    fallo de COM con rutas que contienen caracteres no-ASCII (tildes, ñ, etc.).
    """
    dir_archivo = os.path.dirname(os.path.abspath(ruta_archivo))
    _, ext      = os.path.splitext(ruta_archivo)
    ruta_tmp    = os.path.join(dir_archivo, f"_af_{uuid.uuid4().hex[:10]}{ext}")
    shutil.copy2(ruta_archivo, ruta_tmp)

    with xw.App(visible=False) as app:
        app.display_alerts  = False
        app.screen_updating = False
        try:
            wb = app.books.open(ruta_tmp)
            try:
                ws          = wb.sheets["Desviaciones"]
                ultima_fila = ws.used_range.last_cell.row
                for fila in range(2, ultima_fila + 1):
                    if ws.range(f'A{fila}').api.MergeCells:
                        continue   # fila fusionada → conservar alto manual
                    ws.range(f'A{fila}:I{fila}').api.EntireRow.AutoFit()
                wb.save()
                eliminar_bak(ruta_tmp)
            finally:
                wb.close()
            shutil.copy2(ruta_tmp, ruta_archivo)
        except Exception as e:
            logger.warning(f"  No se pudo autoajustar el alto de filas: {e}")
        # app.quit() se llama automáticamente al salir del 'with'

    try:
        os.remove(ruta_tmp)
    except OSError:
        pass