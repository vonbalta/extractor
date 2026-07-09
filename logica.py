"""
logica.py — Orquestador de procesamiento del Procesador de Esterilizado.

Se importa desde ui.py; no tiene dependencias de tkinter ni de interfaz gráfica.

Responsabilidades (solo orquestación)
--------------------------------------
  - Definir OperacionCancelada y el punto de chequeo cooperativo.
  - Parsear la hora de inicio y calcular el índice de primer archivo.
  - Decidir qué archivos son nuevos (deduplicación).
  - Coordinar los tres flujos principales delegando en los módulos de dominio:
      Modo 1 — Crear reporte nuevo          → ejecutar_extractor_lotes
      Modo 2 — Actualizar reporte existente → ejecutar_actualizador_lotes
      Modo 3 — Revisar desviaciones         → ejecutar_revisor_desviaciones

Módulos de dominio que orquesta
---------------------------------
  pdf_extractor  : extracción de metadatos y tabla desde PDFs, escaneo de carpeta.
  excel_manager  : escritura/lectura de hojas xlwings, backups, nomenclatura.
  deviation_rules: detección vectorizada de desviaciones (pandas/NumPy).
  report_builder : construcción y escritura de la hoja 'Desviaciones'.
"""

import os
import shutil
import tempfile
import uuid
import logging
from datetime import datetime

import xlwings as xw
from openpyxl import load_workbook
from openpyxl import Workbook as _OWB

from pdf_extractor import escanear_carpeta_pdfs, MARCAS_AL_FINAL, PALABRAS_DESCARTE_DEFAULT
from config import parsear_hora
from excel_manager import (
    llenar_hoja_con_datos,
    leer_metadatos_hoja,
    eliminar_bak,
    crear_backup_si_existe,
    nombre_hoja,
    romper_vinculos_externos,
    PATRON_HOJA_NUM,
    asegurar_archivo_libre,
)
from report_builder import (
    recolectar_desviaciones,
    poblar_hoja,
    leer_contenido_existente,
    aplicar_contenido_existente,
)

logger = logging.getLogger(__name__)
logger.addHandler(logging.NullHandler())


# ---------------------------------------------------------------------------
#  CANCELACIÓN COOPERATIVA
# ---------------------------------------------------------------------------

class OperacionCancelada(Exception):
    """
    Señala que el usuario pidió cancelar (botón 'Cancelar' en la UI).
    No es un error de procesamiento: ui.py la distingue de otras excepciones
    para mostrar 'Cancelado' en lugar de un mensaje de fallo.
    """
    pass


def _verificar_cancelacion(evento_cancelacion):
    """
    Punto de chequeo cooperativo a insertar dentro de cualquier bucle que
    pueda recorrer cientos de elementos (PDFs, hojas).

    Si `evento_cancelacion` es None (caller que no soporta cancelación), no
    hace nada. Si está activo, lanza OperacionCancelada.
    """
    if evento_cancelacion is not None and evento_cancelacion.is_set():
        raise OperacionCancelada("Operación cancelada por el usuario.")


# ---------------------------------------------------------------------------
#  HELPERS INTERNOS DEL ORQUESTADOR
# ---------------------------------------------------------------------------

def _parsear_hora(hora_str):
    """Convierte 'HH:MM' o 'HH.MM' a minutos totales. Lanza ValueError si es inválida."""
    return parsear_hora(hora_str)


def _indice_inicio_por_hora(archivos_validos, min_obj):
    """
    Encuentra el índice del archivo cuyo 'minutos_totales' (minuto del día) está
    más cerca de `min_obj`. En caso de empate de distancia se prefiere el índice
    menor, es decir, el primer archivo cronológicamente (la lista viene ordenada
    por 'fecha_dt').

    IMPORTANTE — por qué búsqueda lineal y no bisect
    -------------------------------------------------
    'minutos_totales' es el minuto del día (0-1439) del ciclo, y NO es monótono
    cuando los PDFs abarcan varias fechas o cruzan la medianoche (p. ej. un turno
    nocturno de 19:50 a 06:00). Una búsqueda binaria (bisect) exige una lista
    ordenada; aplicada a 'minutos_totales' —ordenada por fecha, no por minuto—
    devolvería el archivo equivocado. El costo O(n) es irrelevante: n es el número
    de ciclos (típicamente decenas), no de filas.
    """
    n = len(archivos_validos)
    if n == 0:
        return 0
    return min(
        range(n),
        key=lambda i: abs(archivos_validos[i]['minutos_totales'] - min_obj),
    )


def _ya_existe(meta_nuevo, entradas_existentes, tolerancia_seg=60):
    """
    Devuelve True si ya hay una entrada existente para el mismo autoclave
    con una fecha_dt dentro de `tolerancia_seg` segundos de `meta_nuevo`.
    """
    for e in entradas_existentes:
        if e['autoclave'] == meta_nuevo['autoclave']:
            diff = abs((e['fecha_dt'] - meta_nuevo['fecha_dt']).total_seconds())
            if diff <= tolerancia_seg:
                return True
    return False


# ---------------------------------------------------------------------------
#  HELPERS DE ORQUESTACIÓN COMPARTIDOS
# ---------------------------------------------------------------------------

def _ruta_temporal(directorio, extension):
    """
    Genera una ruta única para archivo temporal en el mismo disco que
    ``directorio``, usando un UUID corto como nombre base.

    La extensión incluye el punto (ej. ``.xlsx``).
    """
    return os.path.join(directorio, f"_tmp_{uuid.uuid4().hex[:10]}{extension}")


def _reemplazo_atomico(ruta_temporal, ruta_destino):
    """
    Respalda el destino si existe, verifica que no esté bloqueado y
    reemplaza atómicamente con ``os.replace``.
    """
    crear_backup_si_existe(ruta_destino)
    asegurar_archivo_libre(ruta_destino)
    os.replace(ruta_temporal, ruta_destino)


def _limpiar_temporal(*rutas):
    """
    Elimina los archivos temporales indicados, ignorando errores.
    """
    for r in rutas:
        if r and os.path.exists(r):
            try:
                os.remove(r)
            except OSError:
                pass


# ---------------------------------------------------------------------------
#  MODO 1: Crear reporte nuevo
# ---------------------------------------------------------------------------

def ejecutar_extractor_lotes(
    ruta_plantilla, hora_inicio, carpeta_pdfs, ruta_salida,
    callback_progreso,
    evento_cancelacion=None,
    callback_stats=None,
    palabras_descarte=None,
):
    """
    Modo 1 — Genera un reporte Excel nuevo a partir de una plantilla y los
    PDFs de la carpeta indicada, comenzando por el ciclo más cercano a
    `hora_inicio`.

    Estrategia de escritura atómica: todo el trabajo ocurre sobre un archivo
    temporal (_tmp_*.xlsx/xlsm) en el mismo disco que `ruta_salida`. Solo
    cuando TODO termina con éxito se ejecuta os.replace (atómico) para
    sustituir el destino final. Si algo falla o el usuario cancela, el
    temporal se descarta y `ruta_salida` permanece intacto.

    `palabras_descarte` : iterable opcional de palabras/subcadenas que, si
    aparecen en el nombre de un PDF, hacen que ese archivo se omita del
    escaneo (ver pdf_extractor.escanear_carpeta_pdfs). Si es None, se usa
    PALABRAS_DESCARTE_DEFAULT.

    Devuelve la ruta del archivo generado.
    """
    if os.path.abspath(ruta_salida) == os.path.abspath(ruta_plantilla):
        raise ValueError("La ruta de salida no puede ser la misma que la plantilla.")

    min_obj = _parsear_hora(hora_inicio)

    if os.path.splitext(ruta_plantilla)[1].lower() in (".xlsm", ".xltm"):
        base, ext = os.path.splitext(ruta_salida)
        if ext.lower() != ".xlsm":
            ruta_salida = base + ".xlsm"

    logger.info("Escaneando archivos PDF en la carpeta…")
    callback_progreso(5)
    archivos_validos = escanear_carpeta_pdfs(
        carpeta_pdfs, evento_cancelacion,
        verificar_cancelacion_fn=_verificar_cancelacion,
        palabras_descarte=palabras_descarte,
    )
    if not archivos_validos:
        raise RuntimeError("No se encontraron PDFs válidos para procesar.")

    _verificar_cancelacion(evento_cancelacion)

    archivos_validos.sort(key=lambda x: x['fecha_dt'])
    indice_inicio    = _indice_inicio_por_hora(archivos_validos, min_obj)
    archivos_a_proc  = archivos_validos[indice_inicio:]
    logger.info(f"{len(archivos_a_proc)} informe(s) a procesar desde las {hora_inicio}.")
    callback_progreso(10)


    _dir = os.path.dirname(os.path.abspath(ruta_salida)) or "."
    _ext = os.path.splitext(ruta_salida)[1]
    ruta_trabajo = _ruta_temporal(_dir, _ext)
    shutil.copy2(ruta_plantilla, ruta_trabajo)

    logger.info("Abriendo Excel y procesando hojas…")

    try:
        with xw.App(visible=False) as app:
            app.display_alerts  = False
            app.screen_updating = False
            wb = None
            try:
                wb = app.books.open(ruta_trabajo)
                hoja_maestra = wb.sheets[0]

                regulares    = [a for a in archivos_a_proc if not a.get('marca')]
                con_marca    = [a for a in archivos_a_proc if a.get('marca')]
                lista_ord    = regulares + con_marca
                total        = len(lista_ord)

                for i, arch in enumerate(lista_ord, start=1):
                    _verificar_cancelacion(evento_cancelacion)
                    nh = nombre_hoja(i, arch['autoclave'], arch.get('marca'))
                    hoja_maestra.copy(after=wb.sheets[-1], name=nh)
                    ws = wb.sheets[nh]
                    llenar_hoja_con_datos(ws, arch)
                    logger.info(f"  [{i}/{total}] Procesado: {nh}  —  {arch['fecha_display']} {arch['hora_display']}")
                    callback_progreso(10 + int((i / total) * 80))
                    if callback_stats:
                        callback_stats({"procesados": i, "total": total})

                hoja_maestra.delete()
                romper_vinculos_externos(wb)
                wb.save()
                eliminar_bak(ruta_trabajo)
                wb.close()
                wb = None

                _reemplazo_atomico(ruta_trabajo, ruta_salida)

                logger.info(f"Reporte guardado en: {ruta_salida}")
                callback_progreso(100)
                return ruta_salida

            except Exception as e:
                if wb is not None:
                    try:
                        wb.close()
                    except Exception:
                        pass
                if isinstance(e, OperacionCancelada):
                    logger.info("Operación cancelada por el usuario.")
                raise
            # app.quit() se llama automáticamente al salir del 'with'
    finally:
        _limpiar_temporal(ruta_trabajo)


# ---------------------------------------------------------------------------
#  MODO 2: Actualizar reporte existente
# ---------------------------------------------------------------------------

def ejecutar_actualizador_lotes(
    ruta_existente,
    ruta_plantilla,
    hora_inicio,
    carpeta_pdfs,
    callback_progreso,
    evento_cancelacion=None,
    callback_stats=None,
    palabras_descarte=None,
):
    """
    Modo 2 — Agrega nuevos ciclos a un reporte Excel existente y reordena
    todas las hojas cronológicamente.

    Estrategia de escritura atómica: todo el trabajo ocurre sobre copias
    temporales de `ruta_existente` y `ruta_plantilla`. El archivo real solo
    se sustituye al final con os.replace (atómico). Si algo falla o se cancela,
    `ruta_existente` permanece exactamente como estaba.

    `palabras_descarte` : ver ejecutar_extractor_lotes.
    """
    min_obj = _parsear_hora(hora_inicio)

    # 1. Escanear PDFs
    logger.info("Escaneando archivos PDF en la carpeta…")
    callback_progreso(5)
    archivos_validos = escanear_carpeta_pdfs(
        carpeta_pdfs, evento_cancelacion,
        verificar_cancelacion_fn=_verificar_cancelacion,
        palabras_descarte=palabras_descarte,
    )
    if not archivos_validos:
        raise RuntimeError("No se encontraron PDFs válidos para procesar.")

    archivos_validos.sort(key=lambda x: x['fecha_dt'])
    indice_inicio = _indice_inicio_por_hora(archivos_validos, min_obj)
    candidatos    = archivos_validos[indice_inicio:]
    callback_progreso(10)


    _verificar_cancelacion(evento_cancelacion)

    # 2. Copias temporales de trabajo (nunca se tocan los archivos reales)
    _, ext_plantilla    = os.path.splitext(ruta_plantilla)
    fd, ruta_plt_tmp    = tempfile.mkstemp(suffix=ext_plantilla)
    os.close(fd)
    shutil.copy2(ruta_plantilla, ruta_plt_tmp)

    _dir_ex      = os.path.dirname(os.path.abspath(ruta_existente)) or "."
    _ext_ex      = os.path.splitext(ruta_existente)[1]
    ruta_trabajo = _ruta_temporal(_dir_ex, _ext_ex)
    shutil.copy2(ruta_existente, ruta_trabajo)

    logger.info("Abriendo el reporte existente…")

    # Si Excel/COM no logra iniciar, la limpieza de temporales del 'finally'
    # interno nunca se ejecutaría (está dentro del with). Se cubre ese hueco
    # eliminando los temporales aquí antes de propagar el error.
    try:
        _app_cm = xw.App(visible=False)
    except Exception:
        for _r in (ruta_plt_tmp, ruta_trabajo):
            try:
                os.remove(_r)
            except OSError:
                pass
        raise

    with _app_cm as app:
        app.display_alerts  = False
        app.screen_updating = False
        wb = wb_plantilla = wb_tmp = None
        try:
            wb           = app.books.open(ruta_trabajo)
            wb_plantilla = app.books.open(ruta_plt_tmp)
            hoja_maestra = wb_plantilla.sheets[0]

            # 3. Leer hojas existentes
            entradas_existentes = []
            numero_maximo = 0

            for ws in wb.sheets:
                match_num = PATRON_HOJA_NUM.match(ws.name)
                if match_num:
                    numero_maximo = max(numero_maximo, int(match_num.group(1)))
                    meta = leer_metadatos_hoja(ws)
                    if meta:
                        entradas_existentes.append(meta)
                    else:
                        logger.warning(f"No se pudieron leer metadatos de la hoja '{ws.name}'")

            logger.info(f"{len(entradas_existentes)} hoja(s) ya existentes en el reporte.")
            callback_progreso(20)

            # 4. Filtrar solo PDFs nuevos
            nuevos = [c for c in candidatos if not _ya_existe(c, entradas_existentes)]

            if not nuevos:
                logger.info("Todos los PDFs encontrados ya estaban incluidos. No se agregó nada.")
                wb.close()
                wb = None
                if wb_plantilla:
                    wb_plantilla.close()
                    wb_plantilla = None
                callback_progreso(100)
                return  # Salimos normalmente, el with cierra app

            logger.info(f"{len(nuevos)} informe(s) nuevo(s) encontrado(s). Agregando…")

            # 5. Agregar hojas nuevas (nombre temporal) al libro de trabajo
            contador_tmp = numero_maximo + 1
            total_nuevos = len(nuevos)

            for idx, arch in enumerate(nuevos, start=1):
                _verificar_cancelacion(evento_cancelacion)
                nombre_tmp = f"_new_{contador_tmp}"
                hoja_maestra.copy(after=wb.sheets[-1], name=nombre_tmp)
                ws_nueva = wb.sheets[nombre_tmp]
                llenar_hoja_con_datos(ws_nueva, arch)
                arch['hoja_obj'] = ws_nueva
                arch['tipo']     = 'nuevo'
                contador_tmp    += 1
                logger.info(
                    f"  [{idx}/{total_nuevos}] Preparado: "
                    f"{arch['fecha_display']} {arch['hora_display']} (A{arch['autoclave']})"
                )
                callback_progreso(20 + int((idx / total_nuevos) * 35))
                if callback_stats:
                    callback_stats({"procesados": idx, "total": total_nuevos})

            # 6. Combinar y ordenar: regulares primero, marcas al final
            lista_final = entradas_existentes + nuevos
            lista_final.sort(key=lambda x: (
                1 if (x.get('marca') in MARCAS_AL_FINAL) else 0,
                x['fecha_dt'],
            ))
            logger.info("Reordenando todas las hojas cronológicamente…")
            callback_progreso(60)

            # 7. Reordenar vía libro temporal (xlwings no tiene move nativo)
            wb_tmp = app.books.add()

            for entrada in lista_final:
                _verificar_cancelacion(evento_cancelacion)
                entrada['hoja_obj'].copy(after=wb_tmp.sheets[-1])

            if len(wb_tmp.sheets) > len(lista_final):
                wb_tmp.sheets[0].delete()

            # Renombrar primero con sufijo único para evitar colisiones
            sufijo = uuid.uuid4().hex[:6]
            for i, ws in enumerate(wb_tmp.sheets, start=1):
                ws.name = f"__r{sufijo}_{i}__"
            for i, (ws, entrada) in enumerate(zip(wb_tmp.sheets, lista_final), start=1):
                ws.name = nombre_hoja(i, entrada['autoclave'], entrada.get('marca'))

            callback_progreso(70)

            # 8. Eliminar hojas antiguas del libro real
            hojas_a_borrar = [
                ws.name for ws in wb.sheets
                if PATRON_HOJA_NUM.match(ws.name) or ws.name.startswith('_new_')
            ]

            # Anclar '__placeholder__' explícitamente antes de 'Desviaciones'
            # para preservar su posición final tras el reordenamiento.
            nombres_actuales = [s.name for s in wb.sheets]
            if "Desviaciones" in nombres_actuales:
                wb.sheets.add('__placeholder__', before=wb.sheets['Desviaciones'])
            else:
                wb.sheets.add('__placeholder__', after=wb.sheets[-1])

            for nombre in hojas_a_borrar:
                try:
                    wb.sheets[nombre].delete()
                except Exception as e:
                    logger.warning(f"No se pudo borrar '{nombre}': {e}")

            callback_progreso(80)

            # 9. Mover hojas ordenadas al libro real (anclando en __placeholder__)
            ancla = wb.sheets['__placeholder__']
            for ws in wb_tmp.sheets:
                ws.copy(after=ancla)
                ancla = wb.sheets[ws.name]

            wb.sheets['__placeholder__'].delete()
            wb_tmp.close()
            wb_tmp = None

            # 10. Forzar columna H oculta en cada hoja de datos
            # (el estado 'Hidden' puede perderse en copias múltiples entre libros)
            for ws in wb.sheets:
                if PATRON_HOJA_NUM.match(ws.name):
                    ws.range('H1').api.EntireColumn.Hidden = True

            romper_vinculos_externos(wb)
            wb.save()
            eliminar_bak(ruta_trabajo)
            wb.close()
            wb = None

            # Cerrar plantilla antes de salir
            if wb_plantilla:
                wb_plantilla.close()
                wb_plantilla = None

            _reemplazo_atomico(ruta_trabajo, ruta_existente)

            logger.info(f"Reporte actualizado: {ruta_existente}")
            logger.info(
                f"  Se agregaron {len(nuevos)} hoja(s) y se reordenaron "
                f"{len(lista_final)} en total."
            )
            callback_progreso(100)

        except Exception as e:
            for libro in (wb_tmp, wb):
                if libro is not None:
                    try:
                        libro.close()
                    except Exception:
                        pass
            if wb_plantilla is not None:
                try:
                    wb_plantilla.close()
                except Exception:
                    pass
            if isinstance(e, OperacionCancelada):
                logger.info("Operación cancelada por el usuario.")
            raise
        # Limpieza de temporales de disco (independiente de los objetos COM)
        finally:
            _limpiar_temporal(ruta_plt_tmp, ruta_trabajo)
    # app.quit() se ejecuta al salir del with


# ---------------------------------------------------------------------------
#  MODO 3: Revisar desviaciones
# ---------------------------------------------------------------------------

def ejecutar_revisor_desviaciones(
    ruta_excel,
    callback_progreso,
    ruta_salida=None,
    evento_cancelacion=None,
    callback_stats=None,
):
    """
    Modo 3 — Analiza un reporte Excel existente, detecta desviaciones y
    genera (o actualiza) la hoja 'Desviaciones'.

    Parámetros
    ----------
    ruta_excel        : str  — ruta al .xlsx/.xlsm a analizar.
    callback_progreso : callable(int) — actualiza la barra de progreso (0-100).
    ruta_salida       : str | None — destino del archivo de salida. Si es None,
                        se guarda junto al original con el sufijo
                        '_Revisión desviaciones'.

    Estrategia para preservar drawings y vínculos externos
    -------------------------------------------------------
    openpyxl descarta silenciosamente drawings, gráficos e imágenes al guardar.
    Para evitarlo:
      1. openpyxl solo se usa para LEER (data_only=True) y para escribir la
         hoja 'Desviaciones' en un workbook temporal de una sola hoja.
      2. El archivo de salida se construye copiando el ORIGINAL byte a byte,
         preservando drawings, vínculos y macros.
      3. xlwings (COM nativo) abre ambos archivos, copia la hoja de uno al
         otro y guarda. Excel gestiona la integridad del paquete OOXML.

    Devuelve la ruta real del archivo generado.
    """
    logger.info("Abriendo archivo para revisión de desviaciones…")
    callback_progreso(5)

    _, ext_origen  = os.path.splitext(ruta_excel)
    tiene_macros   = ext_origen.lower() in (".xlsm", ".xltm")

    _verificar_cancelacion(evento_cancelacion)

    # 1. Leer solo valores (nunca se guarda sobre este wb)
    wb_valores = load_workbook(ruta_excel, data_only=True)
    callback_progreso(15)

    _verificar_cancelacion(evento_cancelacion)

    # 2. Preservar datos manuales de la hoja Desviaciones anterior
    contenido_previo = leer_contenido_existente(wb_valores)
    if contenido_previo:
        logger.info(
            f"  Se conservarán datos manuales de {len(contenido_previo)} "
            f"fila(s) de la hoja Desviaciones anterior."
        )

    logger.info("Analizando hojas y detectando desviaciones…")
    registros = recolectar_desviaciones(wb_valores, evento_cancelacion, _verificar_cancelacion)
    callback_progreso(70)

    _verificar_cancelacion(evento_cancelacion)

    logger.info(f"{len(registros)} desviación(es) detectada(s). Generando hoja…")
    if callback_stats:
        callback_stats({"desviaciones": len(registros)})

    # 3. Crear hoja Desviaciones en un workbook temporal (ASCII path para COM)
    #    Los tres temporales se declaran aquí para que el 'finally' pueda
    #    limpiarlos siempre (éxito, error de proceso o fallo de arranque de COM).
    _ruta_tmp_desvs = _ruta_dst_tmp = _ruta_src_tmp = None
    try:
        _wb_tmp = _OWB()
        _ws_tmp = _wb_tmp.active
        _ws_tmp.title = "Desviaciones"
        poblar_hoja(_ws_tmp, registros)
        aplicar_contenido_existente(_ws_tmp, contenido_previo)

        _ruta_tmp_desvs = _ruta_temporal(tempfile.gettempdir(), ".xlsx")
        _wb_tmp.save(_ruta_tmp_desvs)
        callback_progreso(80)

        # 4. Preparar ruta de salida y copiar el original byte a byte
        if ruta_salida:
            base_salida = os.path.splitext(ruta_salida)[0]
        else:
            base_salida = os.path.splitext(ruta_excel)[0] + "_Revisión desviaciones"

        ext_salida  = ".xlsm" if tiene_macros else ".xlsx"
        ruta_salida = base_salida + ext_salida

        # Evitar shutil.SameFileError si la ruta de entrada y salida son la misma.
        if os.path.abspath(ruta_excel) != os.path.abspath(ruta_salida):
            asegurar_archivo_libre(ruta_salida)
            shutil.copy2(ruta_excel, ruta_salida)
        callback_progreso(83)

        # 5. Inyectar la hoja vía xlwings (COM) sobre copias con nombres ASCII
        #    para esquivar el fallo de COM con rutas no-ASCII (tildes, ñ).
        _dir          = os.path.dirname(os.path.abspath(ruta_salida))
        _ruta_dst_tmp = _ruta_temporal(_dir, ext_salida)
        _ruta_src_tmp = _ruta_temporal(_dir, ".xlsx")
        shutil.copy2(ruta_salida,    _ruta_dst_tmp)
        shutil.copy2(_ruta_tmp_desvs, _ruta_src_tmp)

        with xw.App(visible=False) as _app:
            _app.display_alerts  = False
            _app.screen_updating = False
            _wb_dst = _wb_src = None
            try:
                _wb_dst = _app.books.open(_ruta_dst_tmp)
                _wb_src = _app.books.open(_ruta_src_tmp)

                for _sh in list(_wb_dst.sheets):
                    if _sh.name == "Desviaciones":
                        _sh.delete()
                        break

                _wb_src.sheets["Desviaciones"].copy(after=_wb_dst.sheets[-1])

                _wb_src.close()
                _wb_src = None
                romper_vinculos_externos(_wb_dst)
                _wb_dst.save()
                eliminar_bak(_ruta_dst_tmp)
                _wb_dst.close()
                _wb_dst = None

                shutil.copy2(_ruta_dst_tmp, ruta_salida)

            except Exception:
                for _libro in (_wb_src, _wb_dst):
                    if _libro is not None:
                        try:
                            _libro.close()
                        except Exception:
                            pass
                raise
            # _app.quit() al salir del 'with'

        logger.info(f"  Archivo guardado: {ruta_salida}")
        callback_progreso(100)
        return ruta_salida

    finally:
        _limpiar_temporal(_ruta_tmp_desvs, _ruta_dst_tmp, _ruta_src_tmp)