"""
config.py — Configuración persistente del Procesador de Esterilizado

Guarda valores en un archivo JSON ubicado en la carpeta de datos del usuario,
para que se recuerden entre sesiones aunque la app se compile como .exe con
PyInstaller.

Claves gestionadas
------------------
  ruta_plantilla   : ruta al archivo .xlsx/.xlsm que sirve de plantilla
  ruta_existente   : ruta al reporte Excel ya existente (Modo 2 – Actualizar)
  carpeta_pdfs     : carpeta que contiene los PDFs a procesar
  hora_inicio      : hora de inicio en formato "HH:MM"
  palabras_descarte: lista de palabras/subcadenas de nombre de archivo que
                      hacen que un PDF se omita del escaneo (ver
                      pdf_extractor.escanear_carpeta_pdfs)
"""

import json
import os
import tempfile

NOMBRE_APP = "ProcesadorEsterilizado"

# Palabras que se descartan si el usuario nunca configuró nada (mismo
# comportamiento que tenía pdf_extractor.py de forma fija).
_PALABRAS_DESCARTE_DEFAULT = ["prueba", "limpieza", "xxx", "testing", "calibracion"]

# Claves válidas y sus valores por defecto
_DEFAULTS = {
    "ruta_plantilla":    "",
    "ruta_existente":    "",
    "carpeta_pdfs":      "",
    "hora_inicio":       "",
    "tema":              "light",
    "palabras_descarte": list(_PALABRAS_DESCARTE_DEFAULT),
    "param_temp_critica_f4": 122.5,
    "param_dif_temp_max": 1.0,
    "param_dif_temp_min": -0.1,
    "param_presion_min_f3": 1.6,
    "param_presion_min_f4": 1.5,
    "param_caudal_min_123": 230.0,
    "param_caudal_min_4": 210.0,

    "txt_hdr_n_alto": "Diferencial de temperatura entre registrador y digital mayor a 1°C en {nombre_fase}",
    "txt_ora_n_alto": "Diferencial de temperatura entre registrador ({r_txt}°C) y digital ({d_txt}°C) mayor a 1°C en {nombre_fase}",
    "txt_hdr_n_bajo": "Lectura de temperatura de registrador por arriba de digital en {nombre_fase}",
    "txt_ora_n_bajo": "Lectura de temperatura de registrador ({r_txt}°C) por arriba de digital ({d_txt}°C) en {nombre_fase}",
    "txt_hdr_r_alto": "Diferencial de temperatura entre registrador y controlador mayor a 1°C en {nombre_fase}",
    "txt_ora_r_alto": "Diferencial de temperatura entre registrador ({r_txt}°C) y controlador ({c_txt}°C) mayor a 1°C en {nombre_fase}",
    "txt_hdr_r_bajo": "Lectura de temperatura de controlador por arriba de registrador en {nombre_fase}",
    "txt_ora_r_bajo": "Lectura de temperatura de controlador ({c_txt}°C) por arriba de registrador ({r_txt}°C) en {nombre_fase}",
    "txt_hdr_reg_crit": "Temperatura de registrador por debajo de la temperatura programada en holding",
    "txt_ora_reg_crit": "Temperatura de registrador ({r_txt}°C) por debajo de la temperatura programada en holding",
    "txt_hdr_dig_crit": "Lectura de temperatura de digital por debajo de la temperatura programada en holding",
    "txt_ora_dig_crit": "Lectura de temperatura de digital ({d_txt}°C) por debajo de la temperatura programada en holding",
    "txt_bar_f3": "Presión configurada menor a 1.6 Bar en {nombre_fase}",
    "txt_bar_f4": "Presión configurada menor a 1.5 Bar en {nombre_fase}",
    "txt_caudal_123": "Caudal de flujo de agua por debajo del límite crítico 230 m3/h en {nombre_fase} de calentamiento",
    "txt_caudal_4": "Caudal por debajo del límite crítico 210 m3/h en holding",
}


# ---------------------------------------------------------------------------
#  Ruta interna del archivo JSON
# ---------------------------------------------------------------------------

def _ruta_archivo_config():
    """
    Devuelve la ruta de config.json dentro de %APPDATA% (Windows)
    o del home del usuario en otros sistemas.
    """
    base = os.environ.get("APPDATA") or os.path.expanduser("~")
    carpeta = os.path.join(base, NOMBRE_APP)
    os.makedirs(carpeta, exist_ok=True)
    return os.path.join(carpeta, "config.json")


# ---------------------------------------------------------------------------
#  Carga y guardado genéricos
# ---------------------------------------------------------------------------

def cargar_config():
    """
    Lee el archivo de configuración completo.
    Devuelve un dict con los defaults si el archivo no existe o está corrupto.
    """
    ruta = _ruta_archivo_config()
    if os.path.exists(ruta):
        try:
            with open(ruta, "r", encoding="utf-8") as f:
                datos = json.load(f)
            # Completar con defaults las claves que falten
            return {**_DEFAULTS, **datos}
        except Exception:
            pass
    return dict(_DEFAULTS)


def guardar_config(datos_nuevos):
    """
    Combina datos_nuevos con la configuración existente y guarda en disco.
    Solo acepta claves reconocidas (definidas en _DEFAULTS).
    Lanza ValueError para claves inválidas y RuntimeError si no se puede
    escribir el archivo.
    """
    claves_invalidas = set(datos_nuevos) - set(_DEFAULTS)
    if claves_invalidas:
        raise ValueError(f"Claves de configuración desconocidas: {claves_invalidas}")

    actual = cargar_config()
    actual.update(datos_nuevos)
    ruta = _ruta_archivo_config()
    try:
        ruta_dir = os.path.dirname(ruta) or "."
        fd, ruta_tmp = tempfile.mkstemp(suffix=".json", dir=ruta_dir)
        try:
            with os.fdopen(fd, "w", encoding="utf-8") as f:
                json.dump(actual, f, indent=2, ensure_ascii=False)
            os.replace(ruta_tmp, ruta)
        except Exception:
            try:
                os.remove(ruta_tmp)
            except OSError:
                pass
            raise
    except Exception as exc:
        raise RuntimeError(f"No se pudo guardar la configuración en {ruta}: {exc}") from exc


# ---------------------------------------------------------------------------
#  VALIDACIÓN COMPARTIDA DE HORA HH:MM
# ---------------------------------------------------------------------------

def parsear_hora(hora_str):
    """
    Convierte 'HH:MM' o 'HH.MM' a minutos totales (0-1439).
    Lanza ValueError si el formato es inválido o el valor está fuera de rango.
    Función canónica usada por config.py, logica.py y ui.py.
    """
    partes = hora_str.strip().replace('.', ':').split(':')
    if len(partes) != 2:
        raise ValueError(f"Formato de hora inválido: '{hora_str}'. Use HH:MM.")
    h, m = int(partes[0]), int(partes[1])
    if not (0 <= h <= 23 and 0 <= m <= 59):
        raise ValueError(f"Hora fuera de rango: '{hora_str}'.")
    return (h * 60) + m


# ---------------------------------------------------------------------------
#  Accesores por clave — usados desde ui.py y logica.py
# ---------------------------------------------------------------------------

def obtener_ruta_plantilla():
    """Ruta al .xlsx/.xlsm plantilla, o '' si nunca se definió."""
    return cargar_config().get("ruta_plantilla", "")

def guardar_ruta_plantilla(ruta):
    """Guarda (o sobreescribe) la ruta de plantilla."""
    if ruta:
        guardar_config({"ruta_plantilla": ruta})


def obtener_ruta_existente():
    """Ruta al reporte Excel existente (Modo 2), o '' si nunca se definió."""
    return cargar_config().get("ruta_existente", "")

def guardar_ruta_existente(ruta):
    """Guarda (o sobreescribe) la ruta del reporte existente."""
    if ruta:
        guardar_config({"ruta_existente": ruta})


def obtener_carpeta_pdfs():
    """Carpeta de PDFs a procesar, o '' si nunca se definió."""
    return cargar_config().get("carpeta_pdfs", "")

def guardar_carpeta_pdfs(carpeta):
    """Guarda (o sobreescribe) la carpeta de PDFs."""
    if carpeta:
        guardar_config({"carpeta_pdfs": carpeta})


def obtener_hora_inicio():
    """Hora de inicio en formato 'HH:MM', o '' si nunca se definió."""
    return cargar_config().get("hora_inicio", "")

def guardar_hora_inicio(hora):
    """
    Guarda la hora de inicio tras validar el formato ('HH:MM' o 'HH.MM').
    Lanza ValueError si el formato es inválido.
    """
    if hora:
        min_obj = parsear_hora(hora)
        h = min_obj // 60
        m = min_obj % 60
        guardar_config({"hora_inicio": f"{h:02d}:{m:02d}"})

def obtener_tema():
    return cargar_config().get("tema", "light")

def guardar_tema(tema):
    if tema in ("light", "dark"):
        guardar_config({"tema": tema})


# ---------------------------------------------------------------------------
#  PALABRAS DE DESCARTE — nombres de PDF a omitir del escaneo
# ---------------------------------------------------------------------------

def _normalizar_palabras_descarte(valor):
    """
    Acepta una lista de strings o un string separado por comas y devuelve
    una lista de palabras en minúsculas, sin espacios sobrantes ni vacías,
    sin duplicados y preservando el orden de primera aparición.
    """
    if valor is None:
        return []
    if isinstance(valor, str):
        crudos = valor.split(",")
    else:
        crudos = valor

    vistos = set()
    limpio = []
    for item in crudos:
        palabra = str(item).strip().lower()
        if palabra and palabra not in vistos:
            vistos.add(palabra)
            limpio.append(palabra)
    return limpio


def obtener_palabras_descarte():
    """
    Devuelve la lista de palabras/subcadenas (en minúsculas) que hacen que
    un PDF se omita del escaneo por coincidir con su nombre de archivo.
    Si el usuario nunca configuró nada, devuelve los valores por defecto
    ("prueba", "limpieza", "xxx").
    """
    valor = cargar_config().get("palabras_descarte", _PALABRAS_DESCARTE_DEFAULT)
    return _normalizar_palabras_descarte(valor)


def obtener_palabras_descarte_texto():
    """Igual que obtener_palabras_descarte() pero como texto 'a, b, c' para UI."""
    return ", ".join(obtener_palabras_descarte())


def guardar_palabras_descarte(valor):
    """
    Guarda la lista de palabras de descarte. Acepta una lista de strings o
    un string separado por comas (p. ej. lo que escribe el usuario en un
    campo de texto). Una lista vacía es válida: significa "no descartar
    ningún PDF por nombre".
    """
    guardar_config({"palabras_descarte": _normalizar_palabras_descarte(valor)})
# ---------------------------------------------------------------------------
#  PLANTILLAS DE TEXTO DE DESVIACIONES
# ---------------------------------------------------------------------------

def obtener_plantillas_textos():
    conf = cargar_config()
    plantillas = {}
    for k in _DEFAULTS.keys():
        if k.startswith("txt_"):
            plantillas[k] = conf.get(k, _DEFAULTS[k])
    return plantillas

def guardar_plantillas_textos(plantillas_nuevas):
    validas = {}
    for k, v in plantillas_nuevas.items():
        if k in _DEFAULTS and k.startswith("txt_"):
            validas[k] = v
    if validas:
        guardar_config(validas)

# ---------------------------------------------------------------------------
#  PARÁMETROS CRÍTICOS DE DESVIACIONES
# ---------------------------------------------------------------------------

def obtener_parametros_criticos():
    conf = cargar_config()
    params = {}
    for k in _DEFAULTS.keys():
        if k.startswith("param_"):
            val = conf.get(k, _DEFAULTS[k])
            try:
                params[k] = float(val)
            except (ValueError, TypeError):
                params[k] = _DEFAULTS[k]
    return params

def guardar_parametros_criticos(params_nuevos):
    validos = {}
    for k, v in params_nuevos.items():
        if k in _DEFAULTS and k.startswith("param_"):
            try:
                validos[k] = float(v)
            except (ValueError, TypeError):
                pass
    if validos:
        guardar_config(validos)
