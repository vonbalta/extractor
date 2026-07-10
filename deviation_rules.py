"""
deviation_rules.py — Detección vectorizada de desviaciones de parámetros.

Responsabilidades
-----------------
- Leer las filas de datos de una hoja openpyxl en un DataFrame de pandas.
- Evaluar las reglas de desviación (temperatura, presión, caudal) mediante
  máscaras booleanas vectorizadas sobre el DataFrame completo.
- Devolver las desviaciones agrupadas por encabezado/fase, en orden de
  primera aparición, listas para que report_builder las formatee.

No tiene dependencias de xlwings, PDF ni de interfaz gráfica.
"""

import logging
from datetime import datetime, time as dtime

import numpy as np
import pandas as pd
import config

logger = logging.getLogger(__name__)
logger.addHandler(logging.NullHandler())


# ---------------------------------------------------------------------------
#  CONSTANTES: índices de columna (1-based) y configuración
# ---------------------------------------------------------------------------

_DC_A  = 1   # Hora (datetime / time)
_DC_B  = 2   # FASE
_DC_D  = 4   # CONTROLADOR °C (CTRL MEASURE)
_DC_E  = 5   # BAR SET
_DC_G  = 7   # REGISTRADOR °C
_DC_L  = 12  # CAUDAL
_DC_M  = 13  # DIGITAL °C  (vacío → reglas de temp/presión no aplican)
_DC_N  = 14  # DIFERENCIAL °C  Digital - Registrador
_DC_O  = 15  # Indicador presión (vacío → regla de presión no aplica)
_DC_P  = 16  # DIFERENCIAL Bar
_DC_R  = 18  # DIFERENCIAL °C  Registrador - Controlador
_DC_T  = 20  # DIFERENCIAL °C  Digital - Controlador

_FILA_DATOS = 7   # Primera fila con datos reales en la hoja

# Temperatura crítica en holding (fase 4): por debajo de este valor se
# genera desviación tanto para registrador como para digital.


# Columnas de interés leídas de la hoja (en el orden del DataFrame).
_DC_COLS = (
    _DC_A, _DC_B, _DC_D, _DC_E, _DC_G,
    _DC_L, _DC_M, _DC_N, _DC_O, _DC_P, _DC_R, _DC_T,
)

# Nombres de columna internos del DataFrame.
_COL_NOMBRES = {
    _DC_A: "hora",
    _DC_B: "fase",
    _DC_D: "controlador",
    _DC_E: "bar_set",
    _DC_G: "registrador",
    _DC_L: "caudal",
    _DC_M: "digital",
    _DC_N: "n_diff",
    _DC_O: "o_ind",
    _DC_P: "p_diff",
    _DC_R: "r_diff",
    _DC_T: "t_diff",
}


# ---------------------------------------------------------------------------
#  HELPERS DE FORMATO (módulo-privados)
# ---------------------------------------------------------------------------

def _float(val):
    try:
        return float(val)
    except (TypeError, ValueError):
        return None


def _vacio(val):
    return val is None or str(val).strip() == ""


def _nombre_fase(fase):
    return "holding" if fase == 4 else f"fase {fase}"


def _fmt_hora_corta(val):
    """
    Convierte datetime / time / str / float a 'HH:MM' (sin segundos).

    El valor de la celda 'INSTANT' puede llegar como cualquiera de estos
    tipos dependiendo de cómo xlwings/openpyxl lo haya serializado.
    """
    if isinstance(val, datetime):
        return val.strftime("%H:%M")
    if isinstance(val, dtime):
        return val.strftime("%H:%M")
    if isinstance(val, (int, float)):
        total_seg = round(float(val) * 86400) % 86400
        return f"{total_seg // 3600:02d}:{(total_seg % 3600) // 60:02d}"
    if isinstance(val, str):
        partes = val.strip().split(':')
        if len(partes) >= 2:
            return f"{partes[0]:0>2}:{partes[1]:0>2}"
        return val.strip()
    return str(val) if val is not None else ""


def _fmt_temp(val):
    """Formatea un valor de temperatura para líneas de detalle ('N/D' si falta)."""
    if val is None or (isinstance(val, float) and np.isnan(val)):
        return "N/D"
    return str(round(val, 2))


# ---------------------------------------------------------------------------
#  LECTURA DE FILAS (openpyxl → DataFrame)
# ---------------------------------------------------------------------------

def leer_filas_hoja(ws_opx):
    """
    Devuelve un DataFrame con una fila por cada fila de datos válida de `ws_opx`
    (desde `_FILA_DATOS`), con las columnas de interés ya nombradas.

    Optimización: el bloque completo de celdas se vuelca en un DataFrame de
    una sola vez; los filtros de fila vacía y fase no parseable se aplican
    con máscaras booleanas vectorizadas en lugar de iterar fila a fila.
    """
    filas_crudas = list(ws_opx.iter_rows(min_row=_FILA_DATOS, values_only=True))

    if not filas_crudas:
        return pd.DataFrame(columns=list(_COL_NOMBRES.values()) + ["_orden"])

    # Guardia POR COLUMNA: se toma el valor solo si esa columna existe en la
    # fila; una columna ausente se vuelve None sin invalidar el resto de la
    # fila. (Antes se exigía que la fila tuviera al menos `max(_DC_COLS)` = 20
    # celdas y, si no, se anulaba la fila COMPLETA — lo que en una hoja más
    # angosta que la columna T producía cero desviaciones de forma silenciosa.)
    datos = {
        _COL_NOMBRES[col]: [
            (fila[col - 1] if len(fila) >= col else None)
            for fila in filas_crudas
        ]
        for col in _DC_COLS
    }
    df = pd.DataFrame(datos)
    df["_orden"] = np.arange(len(df))

    # Descartar filas completamente vacías.
    cols_valor = list(_COL_NOMBRES.values())

    def _col_vacia(s):
        return s.isna() | s.astype(str).str.strip().eq("")

    mascara_vacia = pd.concat([_col_vacia(df[c]) for c in cols_valor], axis=1).all(axis=1)
    df = df[~mascara_vacia]

    # FASE debe ser parseable como entero; si no, se descarta la fila.
    fase_num = pd.to_numeric(df["fase"], errors="coerce")
    df       = df[fase_num.notna()].copy()
    df["fase"] = fase_num[fase_num.notna()].astype(int)

    return df.reset_index(drop=True)


def fecha_hoja(ws_opx):
    """Lee la fecha del ciclo desde la celda C2 de la hoja openpyxl."""
    from datetime import date as _date
    valor = ws_opx['C2'].value
    if isinstance(valor, datetime):
        return valor.date()
    if isinstance(valor, _date):
        return valor
    if isinstance(valor, str):
        try:
            return datetime.strptime(valor.strip(), "%d/%m/%Y").date()
        except ValueError:
            return None
    return None


# ---------------------------------------------------------------------------
#  EVALUACIÓN DE REGLAS (vectorizada)
# ---------------------------------------------------------------------------

def evaluar_hoja(df):
    """
    Evalúa las reglas de desviación sobre el DataFrame producido por
    `leer_filas_hoja` y devuelve la lista de grupos de desviación en orden
    de primera aparición.

    Cada elemento de la lista es un dict:
        {
          "header":      <texto del encabezado / clave de agrupación>,
          "ocurrencias": [ {"hora_corta": "HH:MM", "lecturas": "...",
                            "oracion": "..."}, ... ],
          "con_lecturas": True | False,
        }

    Las reglas de temperatura producen ocurrencias con "lecturas" y "oracion";
    las reglas de presión y caudal solo incluyen "hora_corta".
    """
    entradas  = []
    idx_grupos = {}

    # Cargar parámetros configurables
    params = config.obtener_parametros_criticos()
    _T_CRIT = params.get("param_temp_critica_f4", 122.5)
    _DIF_MAX = params.get("param_dif_temp_max", 1.0)
    _DIF_MIN = params.get("param_dif_temp_min", -0.1)
    _P_MIN_F3 = params.get("param_presion_min_f3", 1.6)
    _P_MIN_F4 = params.get("param_presion_min_f4", 1.5)
    _C_MIN_123 = params.get("param_caudal_min_123", 230.0)
    _C_MIN_4 = params.get("param_caudal_min_4", 210.0)

    if df.empty:
        return entradas

    # — Conversión numérica en bloque —
    n           = pd.to_numeric(df["n_diff"],      errors="coerce").round(4)
    r           = pd.to_numeric(df["r_diff"],      errors="coerce").round(4)
    digital     = pd.to_numeric(df["digital"],     errors="coerce")
    registrador = pd.to_numeric(df["registrador"], errors="coerce")
    controlador = pd.to_numeric(df["controlador"], errors="coerce")
    e_bar       = pd.to_numeric(df["bar_set"],     errors="coerce")
    l_caudal    = pd.to_numeric(df["caudal"],      errors="coerce")
    fase        = df["fase"]

    # 'M no vacío' habilita las reglas de temperatura diferencial.
    m_no_vacio = ~(df["digital"].isna() | df["digital"].astype(str).str.strip().eq(""))

    nombre_fase = fase.apply(_nombre_fase)
    hora_corta  = df["hora"].apply(_fmt_hora_corta)
    d_txt       = digital.apply(_fmt_temp)
    r_txt       = registrador.apply(_fmt_temp)
    c_txt       = controlador.apply(_fmt_temp)

    # ── Máscaras booleanas por regla ────────────────────────────────────────

    mask_n_alto  = m_no_vacio & n.notna() & (n > _DIF_MAX)
    mask_n_bajo  = m_no_vacio & n.notna() & (n <= _DIF_MIN)
    mask_r_alto  = m_no_vacio & r.notna() & (r > _DIF_MAX)
    mask_r_bajo  = m_no_vacio & r.notna() & (r <= _DIF_MIN)

    mask_bar_f3 = e_bar.notna() & (fase == 3) & (e_bar < _P_MIN_F3)
    mask_bar_f4 = e_bar.notna() & (fase == 4) & (e_bar < _P_MIN_F4)

    mask_caudal_123 = l_caudal.notna() & fase.isin([1, 2, 3]) & (l_caudal < _C_MIN_123)
    mask_caudal_4   = l_caudal.notna() & (fase == 4)          & (l_caudal < _C_MIN_4)

    # Temperatura crítica en holding: registrador NO requiere m_no_vacio
    # (columna G siempre presente desde el PDF); digital sí lo requiere.
    mask_reg_crit_f4 = registrador.notna() & (fase == 4) & (registrador < _T_CRIT)
    mask_dig_crit_f4 = m_no_vacio          & (fase == 4) & (digital     < _T_CRIT)

    # ── Textos de encabezado, lecturas y oración (Series completas) ─────────

    txts = config.obtener_plantillas_textos()

    def _fmt(template, **kwargs):
        result = pd.Series(template, index=df.index) if any(isinstance(v, pd.Series) for v in kwargs.values()) else template
        for k, v in kwargs.items():
            if isinstance(v, pd.Series):
                if isinstance(result, str):
                    result = pd.Series(result, index=df.index)
                res_arr = result.to_numpy(dtype=str)
                v_arr = v.to_numpy(dtype=str)
                placeholder = f"{{{k}}}"
                result = pd.Series([s.replace(placeholder, val) for s, val in zip(res_arr, v_arr)], index=df.index)
            else:
                if isinstance(result, pd.Series):
                    result = result.str.replace(f"{{{k}}}", str(v), regex=False)
                else:
                    result = result.replace(f"{{{k}}}", str(v))
        return result

    header_n_alto   = _fmt(txts['txt_hdr_n_alto'], nombre_fase=nombre_fase)
    lecturas_n_alto = r_txt + "°C R, " + d_txt + "°C D"
    oracion_n_alto  = _fmt(txts['txt_ora_n_alto'], r_txt=r_txt, d_txt=d_txt, nombre_fase=nombre_fase)

    header_n_bajo   = _fmt(txts['txt_hdr_n_bajo'], nombre_fase=nombre_fase)
    lecturas_n_bajo = r_txt + "°C R, " + d_txt + "°C D"
    oracion_n_bajo  = _fmt(txts['txt_ora_n_bajo'], r_txt=r_txt, d_txt=d_txt, nombre_fase=nombre_fase)

    header_r_alto   = _fmt(txts['txt_hdr_r_alto'], nombre_fase=nombre_fase)
    lecturas_r_alto = r_txt + "°C R, " + c_txt + "°C C"
    oracion_r_alto  = _fmt(txts['txt_ora_r_alto'], r_txt=r_txt, c_txt=c_txt, nombre_fase=nombre_fase)

    header_r_bajo   = _fmt(txts['txt_hdr_r_bajo'], nombre_fase=nombre_fase)
    lecturas_r_bajo = c_txt + "°C C, " + r_txt + "°C R"
    oracion_r_bajo  = _fmt(txts['txt_ora_r_bajo'], r_txt=r_txt, c_txt=c_txt, nombre_fase=nombre_fase)

    _HDR_REG_CRIT     = _fmt(txts['txt_hdr_reg_crit'], nombre_fase=nombre_fase)
    lecturas_reg_crit = r_txt + "°C R"
    oracion_reg_crit  = _fmt(txts['txt_ora_reg_crit'], r_txt=r_txt, nombre_fase=nombre_fase)

    _HDR_DIG_CRIT     = _fmt(txts['txt_hdr_dig_crit'], nombre_fase=nombre_fase)
    lecturas_dig_crit = d_txt + "°C D"
    oracion_dig_crit  = _fmt(txts['txt_ora_dig_crit'], d_txt=d_txt, nombre_fase=nombre_fase)

    texto_bar_f3     = _fmt(txts['txt_bar_f3'], nombre_fase=nombre_fase)
    texto_bar_f4     = _fmt(txts['txt_bar_f4'], nombre_fase=nombre_fase)
    texto_caudal_123 = _fmt(txts['txt_caudal_123'], nombre_fase=nombre_fase)
    _TXT_CAUDAL_4    = _fmt(txts['txt_caudal_4'], nombre_fase=nombre_fase)

    # ── Acumulación de ocurrencias por header (orden de primera aparición) ──
    #
    # Cada tupla es (máscara, header, lecturas, oración). 'header'/'lecturas'/
    # 'oración' pueden ser una Series (valor por fila) o un str constante.
    reglas_temperatura = (
        (mask_n_alto,       header_n_alto,   lecturas_n_alto,   oracion_n_alto),
        (mask_n_bajo,       header_n_bajo,   lecturas_n_bajo,   oracion_n_bajo),
        (mask_r_alto,       header_r_alto,   lecturas_r_alto,   oracion_r_alto),
        (mask_r_bajo,       header_r_bajo,   lecturas_r_bajo,   oracion_r_bajo),
        (mask_reg_crit_f4,  _HDR_REG_CRIT,   lecturas_reg_crit, oracion_reg_crit),
        (mask_dig_crit_f4,  _HDR_DIG_CRIT,   lecturas_dig_crit, oracion_dig_crit),
    )
    reglas_simples = (
        (mask_bar_f3,       texto_bar_f3),
        (mask_bar_f4,       texto_bar_f4),
        (mask_caudal_123,   texto_caudal_123),
        (mask_caudal_4,     _TXT_CAUDAL_4),
    )

    def _acumular(clave, header, ocurrencia, con_lecturas):
        if clave in idx_grupos:
            entradas[idx_grupos[clave]]["ocurrencias"].append(ocurrencia)
        else:
            idx_grupos[clave] = len(entradas)
            entradas.append({
                "header":       header,
                "ocurrencias":  [ocurrencia],
                "con_lecturas": con_lecturas,
            })

    hc_list = hora_corta.tolist()

    def _acceso(valor):
        """Devuelve una función pos->valor tanto para Series como para escalares."""
        if isinstance(valor, pd.Series):
            vals = valor.tolist()
            return lambda pos: vals[pos]
        return lambda _pos: valor

    # En lugar de recorrer fila-a-fila comprobando cada regla con .iat (miles de
    # accesos escalares aunque casi ninguna fila desvíe), se localizan de golpe
    # las filas que disparan cada máscara con np.flatnonzero (nivel C) y se
    # ordenan los eventos por (fila, orden_de_regla). Esto reproduce EXACTAMENTE
    # el orden de 'primera aparición' del bucle original —temperatura antes que
    # presión/caudal, y dentro de cada fila el orden de definición de las reglas—
    # visitando solo las filas realmente desviadas (típicamente muy pocas).

    eventos_temp = []
    for ridx, (mask, hdr, lec, ora) in enumerate(reglas_temperatura):
        get_hdr, get_lec, get_ora = _acceso(hdr), _acceso(lec), _acceso(ora)
        for pos in np.flatnonzero(mask.to_numpy()):
            pos = int(pos)
            eventos_temp.append((pos, ridx, get_hdr(pos), {
                "hora_corta": hc_list[pos],
                "lecturas":   get_lec(pos),
                "oracion":    get_ora(pos),
            }))
    eventos_temp.sort(key=lambda e: (e[0], e[1]))
    for _pos, _ridx, header, ocurrencia in eventos_temp:
        _acumular(header, header, ocurrencia, con_lecturas=True)

    eventos_simples = []
    for ridx, (mask, texto) in enumerate(reglas_simples):
        get_txt = _acceso(texto)
        for pos in np.flatnonzero(mask.to_numpy()):
            pos = int(pos)
            eventos_simples.append((pos, ridx, get_txt(pos)))
    eventos_simples.sort(key=lambda e: (e[0], e[1]))
    for _pos, _ridx, texto in eventos_simples:
        _acumular(texto, texto, {"hora_corta": hc_list[_pos]}, con_lecturas=False)

    return entradas