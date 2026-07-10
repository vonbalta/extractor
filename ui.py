"""
ui.py — Interfaz gráfica para el Procesador de Esterilizado.
Diseño original Windows 11 + modo oscuro + notificaciones emergentes
+ protecciones contra saturación + ventana redimensionable
+ persistencia de configuración (carpeta, hora, tema) con parche de seguridad.
"""

import sys
import os
import html
import shutil
import threading
import logging
import logging.handlers

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QLineEdit, QFileDialog, QTextEdit,
    QProgressBar, QFrame, QStackedWidget, QGraphicsDropShadowEffect,
    QCheckBox,QDialog, QDialogButtonBox, QFormLayout
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QTimer, QVariantAnimation
from PyQt6.QtGui import QFont, QColor, QPalette, QTextCursor, QKeySequence, QShortcut

from qfluentwidgets import ScrollArea, SmoothScrollArea, InfoBar, InfoBarPosition, FluentWindow, NavigationItemPosition, SubtitleLabel, setTheme, Theme, NavigationInterface, PushButton, PrimaryPushButton, LineEdit, TextEdit, ProgressBar, CheckBox, CardWidget, BodyLabel, CaptionLabel, IconWidget, FluentIcon, ToolTipFilter

import config

# ─── Parche de seguridad ─────────────────────────────────────────────────
# Si el módulo config no tiene alguna de estas funciones, creamos unas vacías.
for func_name in [
    "obtener_ruta_plantilla", "guardar_ruta_plantilla",
    "obtener_ruta_existente", "guardar_ruta_existente",
    "obtener_carpeta_pdfs", "guardar_carpeta_pdfs",
    "obtener_hora_inicio", "guardar_hora_inicio",
    "obtener_tema", "guardar_tema",
    "obtener_palabras_descarte_texto", "guardar_palabras_descarte",
]:
    if not hasattr(config, func_name):
        if func_name.startswith("obtener"):
            setattr(config, func_name, lambda: "")
        else:
            setattr(config, func_name, lambda v: None)

# ─── Paletas de colores ─────────────────────────────────────────────────
C_LIGHT = {
    "bg":          "#F3F3F3",
    "surface":     "#FFFFFF",
    "surface_2":   "#FAFAFA",
    "overlay":     "#E5E5E5",
    "muted":       "#5D5D5D",
    "muted_light": "#8A8A8A",
    "text":        "#1A1A1A",
    "text_soft":   "#3B3B3B",
    "accent":      "#005FB8",
    "accent_h":    "#0054A6",
    "accent_light":"#F3F9FF",
    "success":     "#0F7B0F",
    "success_bg":  "#DFF6DD",
    "warning":     "#9D5D00",
    "error":       "#C42B1C",
    "border":      "#E5E5E5",
    "border_soft": "#F0F0F0",
    "border_input":"#D1D1D1",
    "log_bg":      "#FAFAFA",
    "shadow":      "rgba(0,0,0,0.05)",
}

C_DARK = {
    "bg":          "#1E1E1E",
    "surface":     "#2B2B2B",
    "surface_2":   "#252525",
    "overlay":     "#383838",
    "muted":       "#AAAAAA",
    "muted_light": "#777777",
    "text":        "#E5E5E5",
    "text_soft":   "#CCCCCC",
    "accent":      "#60CDFF",
    "accent_h":    "#4DB8E8",
    "accent_light":"#1A3A4A",
    "success":     "#6FDD6F",
    "success_bg":  "#1A3A1A",
    "warning":     "#FFB84D",
    "error":       "#FF6B6B",
    "border":      "#404040",
    "border_soft": "#333333",
    "border_input":"#505050",
    "log_bg":      "#252525",
    "shadow":      "rgba(0,0,0,0.3)",
}

C = C_LIGHT.copy()
current_theme = "light"

# ─── Estilos dinámicos con clases CSS para etiquetas ────────────────────
def generar_estilos(C):
    return f"""
    QMainWindow, QWidget {{
        background-color: {C["bg"]};
        color: {C["text"]};
        font-family: "Segoe UI Variable Text", "Segoe UI Variable Display", "Segoe UI", sans-serif;
        font-size: 14px;
    }}
    QLabel {{
        color: {C["text"]};
        background: transparent;
    }}
    QLabel[class="text-soft"] {{
        color: {C["text_soft"]};
        background: transparent;
        font-weight: 600;
    }}
    QLabel[class="text-muted"] {{
        color: {C["muted"]};
        background: transparent;
    }}
    QLabel[class="text-muted-light"] {{
        color: {C["muted_light"]};
        background: transparent;
        font-weight: 600;
        letter-spacing: 0.8px;
    }}
    QLabel[class="badge"] {{
        color: {C["accent"]};
        background-color: {C["accent_light"]};
        border-radius: 4px;
        padding: 3px 8px;
        font-size: 10px;
        font-weight: 700;
        letter-spacing: 1px;
    }}
    QLineEdit {{
        background-color: {C["surface"]};
        color: {C["text"]};
        border: 1px solid {C["border"]};
        border-bottom: 1px solid {C["border_input"]};
        border-radius: 4px;
        padding: 6px 12px;
        font-size: 14px;
        selection-background-color: {C["accent"]};
        selection-color: #ffffff;
    }}
    QLineEdit:focus {{
        border: 1px solid {C["accent"]};
        border-bottom: 2px solid {C["accent"]};
        background-color: {C["surface"]};
    }}
    QLineEdit:read-only {{
        background-color: {C["surface_2"]};
        color: {C["muted"]};
        border: 1px solid {C["border_soft"]};
    }}
    QTextEdit {{
        background-color: {C["log_bg"]};
        color: {C["text"]};
        border: 1px solid {C["border"]};
        border-radius: 4px;
        padding: 8px;
        font-family: "Cascadia Mono", "Consolas", monospace;
        font-size: 13px;
        selection-background-color: {C["accent"]};
        selection-color: #ffffff;
    }}
    QScrollBar:vertical {{
        background: transparent;
        width: 4px;
        margin: 2px;
        border-radius: 2px;
    }}
    QScrollBar::handle:vertical {{
        background: #8A8A8A;
        border-radius: 2px;
        min-height: 24px;
    }}
    QScrollBar::handle:vertical:hover {{
        background: #5D5D5D;
    }}
    QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{ height: 0px; }}
    QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical {{ background: transparent; }}
    QProgressBar {{
        background-color: {C["border"]};
        border: none;
        border-radius: 2px;
        height: 4px;
        text-align: center;
        color: transparent;
    }}
    QProgressBar::chunk {{
        background-color: {C["accent"]};
        border-radius: 2px;
    }}
    QCheckBox {{
        color: {C["text"]};
        background: transparent;
        font-size: 13px;
        spacing: 8px;
    }}
    QCheckBox::indicator {{
        width: 16px;
        height: 16px;
        border: 1px solid {C["border_input"]};
        border-radius: 3px;
        background-color: {C["surface"]};
    }}
    QCheckBox::indicator:hover {{
        border: 1px solid {C["accent"]};
    }}
    QCheckBox::indicator:checked {{
        background-color: {C["accent"]};
        border: 1px solid {C["accent"]};
    }}
    QCheckBox:disabled {{
        color: {C["muted_light"]};
    }}

    /* Botones */
    QPushButton[class="primary"] {{
        background-color: {C["accent"]};
        color: #ffffff;
        border: 1px solid transparent;
        border-radius: 4px;
        font-size: 14px;
        font-weight: 600;
        padding: 8px 16px;
    }}
    QPushButton[class="primary"]:hover {{
        background-color: {C["accent_h"]};
    }}
    QPushButton[class="primary"]:pressed {{
        background-color: #004482;
        color: rgba(255, 255, 255, 0.7);
    }}
    QPushButton[class="primary"]:disabled {{
        background-color: rgba(0, 0, 0, 0.04);
        color: rgba(0, 0, 0, 0.36);
        border: 1px solid transparent;
    }}

    QPushButton[class="ghost"] {{
        background-color: {C["surface"]};
        color: {C["text"]};
        border: 1px solid {C["border"]};
        border-radius: 4px;
        font-size: 13px;
        font-weight: 400;
        padding: 6px 16px;
    }}
    QPushButton[class="ghost"]:hover {{
        background-color: {C["overlay"]};
    }}
    QPushButton[class="ghost"]:pressed {{
        color: {C["muted"]};
        background-color: #E5E5E5;
    }}

    QPushButton[class="success"] {{
        background-color: {C["success_bg"]};
        color: {C["success"]};
        border: 1px solid transparent;
        border-radius: 4px;
        font-size: 14px;
        font-weight: 600;
        padding: 8px 16px;
    }}
    QPushButton[class="success"]:hover {{
        background-color: #C1EFC1;
    }}
    QPushButton[class="success"]:pressed {{
        background-color: #A3E4A3;
    }}

    /* Foco de teclado visible (accesibilidad): un anillo de acento en el
       control enfocado para que el orden de tabulación sea perceptible. */
    QPushButton:focus {{
        border: 2px solid {C["accent"]};
        padding: 6px 15px;
    }}
    QCheckBox:focus {{
        color: {C["accent"]};
    }}

    /* Tarjetas */
    QFrame#card {{
        background-color: {C["surface"]};
        border: 1px solid {C["border"]};
        border-radius: 8px;
    }}

    QFrame#menuCard {{
        background-color: {C["surface"]};
        border: 1px solid {C["border"]};
        border-radius: 8px;
    }}
    """

MAX_LOG_LINES = 500
_toast_refs = {}

# ─── Manejador de logging ───────────────────────────────────────────────
class UILogHandler(logging.Handler):
    def __init__(self, signal_fn):
        super().__init__()
        self._signal_fn = signal_fn
        self.setFormatter(logging.Formatter("%(message)s"))

    def emit(self, record):
        try:
            self._signal_fn(self.format(record))
        except Exception:
            self.handleError(record)

# ─── Worker Thread ──────────────────────────────────────────────────────
class WorkerThread(QThread):
    progreso  = pyqtSignal(int)
    log       = pyqtSignal(str)
    terminado = pyqtSignal(str)
    error     = pyqtSignal(str)
    cancelado = pyqtSignal()
    stats     = pyqtSignal(dict)

    _LOG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "procesador.log")

    def __init__(self, modo, params):
        super().__init__()
        self.modo   = modo
        self.params = params
        self.evento_cancelacion = threading.Event()
        self._error_emitted = False

    def _configure_logging(self):
        lg = logging.getLogger("logica")
        lg.setLevel(logging.DEBUG)
        ui_handler = UILogHandler(self.log.emit)
        ui_handler.setLevel(logging.DEBUG)
        file_handler = logging.handlers.RotatingFileHandler(
            self._LOG_FILE, maxBytes=2*1024*1024, backupCount=3, encoding="utf-8"
        )
        file_handler.setLevel(logging.DEBUG)
        file_handler.setFormatter(
            logging.Formatter("%(asctime)s [%(levelname)-8s] %(message)s", datefmt="%Y-%m-%d %H:%M:%S")
        )
        lg.addHandler(ui_handler)
        lg.addHandler(file_handler)
        return [ui_handler, file_handler]

    def _teardown_logging(self, handlers):
        lg = logging.getLogger()
        for h in handlers:
            lg.removeHandler(h)
            h.close()

    def solicitar_cancelacion(self):
        self.evento_cancelacion.set()

    def _stats_proxy(self, data):
        self.stats.emit(data)

    def run(self):
        handlers = self._configure_logging()
        cb_stats = self._stats_proxy
        try:
            import logica
            if self.modo == "nuevo":
                ruta_final = logica.ejecutar_extractor_lotes(
                    ruta_plantilla     = self.params["ruta_plantilla"],
                    hora_inicio        = self.params["hora_inicio"],
                    carpeta_pdfs       = self.params["carpeta_pdfs"],
                    ruta_salida        = self.params["ruta_salida"],
                    callback_progreso  = lambda v: self.progreso.emit(v),
                    evento_cancelacion = self.evento_cancelacion,
                    callback_stats     = cb_stats,
                    palabras_descarte  = self.params.get("palabras_descarte"),
                )
                if not self.isInterruptionRequested():
                    self.terminado.emit(ruta_final or self.params["ruta_salida"])
            elif self.modo == "actualizar":
                logica.ejecutar_actualizador_lotes(
                    ruta_existente     = self.params["ruta_existente"],
                    ruta_plantilla     = self.params["ruta_plantilla"],
                    hora_inicio        = self.params["hora_inicio"],
                    carpeta_pdfs       = self.params["carpeta_pdfs"],
                    callback_progreso  = lambda v: self.progreso.emit(v),
                    evento_cancelacion = self.evento_cancelacion,
                    callback_stats     = cb_stats,
                    palabras_descarte  = self.params.get("palabras_descarte"),
                )
                if not self.isInterruptionRequested():
                    self.terminado.emit(self.params["ruta_existente"])
            elif self.modo == "revisar":
                if self.params.get("actualizar_in_situ"):
                    ruta_origen = self.params["ruta_excel"]
                    try:
                        ruta_bak = ruta_origen + ".bak"
                        shutil.copy2(ruta_origen, ruta_bak)
                        logging.getLogger().info(f"  Respaldo creado: {os.path.basename(ruta_bak)}")
                    except Exception as e:
                        logging.getLogger().warning(f"  Advertencia: no se pudo crear respaldo ({e})")
                ruta_final = logica.ejecutar_revisor_desviaciones(
                    ruta_excel        = self.params["ruta_excel"],
                    ruta_salida       = self.params.get("ruta_salida"),
                    callback_progreso = lambda v: self.progreso.emit(v),
                    evento_cancelacion = self.evento_cancelacion,
                    callback_stats     = cb_stats,
                )
                if not self.isInterruptionRequested():
                    self.terminado.emit(ruta_final or self.params["ruta_excel"])
        except logica.OperacionCancelada:
            self.cancelado.emit()
        except Exception as e:
            if not self._error_emitted and not self.isInterruptionRequested():
                self._error_emitted = True
                self.error.emit(str(e))
        finally:
            self._teardown_logging(handlers)

# ─── Notificaciones emergentes ──────────────────────────────────────────
def show_toast(parent, title, content, level="info"):
    position = InfoBarPosition.TOP_RIGHT
    if level in _toast_refs:
        try:
            _toast_refs[level].close()
        except Exception:
            pass
    if level == "success":
        bar = InfoBar.success(title, content, duration=3000, parent=parent, position=position)
    elif level == "error":
        bar = InfoBar.error(title, content, duration=5000, parent=parent, position=position)
    elif level == "warning":
        bar = InfoBar.warning(title, content, duration=3000, parent=parent, position=position)
    else:
        bar = InfoBar.info(title, content, duration=2000, parent=parent, position=position)
    _toast_refs[level] = bar

# ─── Componentes reutilizables ──────────────────────────────────────────
def make_label(text, size=12, weight=400, css_class=None):
    lbl = QLabel(text)
    font = QFont("Segoe UI", size)
    font.setWeight(QFont.Weight(weight))
    lbl.setFont(font)
    if css_class:
        lbl.setProperty("class", css_class)
        lbl.style().unpolish(lbl)
        lbl.style().polish(lbl)
    return lbl

def make_divider():
    line = QFrame()
    line.setFrameShape(QFrame.Shape.HLine)
    line.setStyleSheet(f"border: none; border-top: 1px solid {C['border']};")
    line.setFixedHeight(1)
    return line

class FileSelector(QWidget):
    def __init__(self, label, mode="file", filters="Excel (*.xlsx *.xlsm);;Todos (*)",
                 valor_inicial="", on_select=None, sugerir_nombre=None):
        super().__init__()
        self.mode           = mode
        self.filters        = filters
        self.on_select      = on_select
        self.sugerir_nombre = sugerir_nombre

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(6)

        self.lbl = make_label(label, size=10, css_class="text-soft")
        layout.addWidget(self.lbl)

        row = QHBoxLayout()
        row.setSpacing(8)

        self.field = LineEdit()
        self.field.setPlaceholderText("Sin seleccionar...")
        self.field.setReadOnly(True)
        # Accesibilidad: nombre accesible para lectores de pantalla y tooltip que
        # muestra la ruta completa (el campo suele truncar rutas largas).
        self.field.setAccessibleName(label)
        self.field.setToolTip("Ruta seleccionada. Usa «Examinar» para cambiarla.")
        if valor_inicial:
            self.field.setText(valor_inicial)
            self.field.setToolTip(valor_inicial)
        row.addWidget(self.field, stretch=1)

        _tips = {
            "file": "Seleccionar un archivo",
            "dir":  "Seleccionar una carpeta",
            "save": "Elegir dónde guardar el archivo",
        }
        btn = PushButton("Examinar")
        btn.setProperty("class", "ghost")
        btn.setCursor(Qt.CursorShape.PointingHandCursor)
        btn.setToolTip(_tips.get(self.mode, "Examinar"))
        btn.setAccessibleName(f"Examinar {label}")
        btn.clicked.connect(self._browse)
        row.addWidget(btn)

        layout.addLayout(row)

    def _browse(self):
        if self.mode == "file":
            path, _ = QFileDialog.getOpenFileName(self, "Seleccionar archivo", "", self.filters)
        elif self.mode == "dir":
            path = QFileDialog.getExistingDirectory(self, "Seleccionar carpeta")
        elif self.mode == "save":
            sugerido = ""
            if self.sugerir_nombre:
                try:
                    sugerido = self.sugerir_nombre() or ""
                except Exception:
                    sugerido = ""
            path, _ = QFileDialog.getSaveFileName(self, "Guardar como", sugerido, self.filters)
        else:
            path = ""
        if path:
            self.field.setText(path)
            self.field.setToolTip(path)
            if self.on_select:
                self.on_select(path)

    def set_value(self, ruta):
        self.field.setText(ruta or "")
        self.field.setToolTip(ruta or "Ruta seleccionada. Usa «Examinar» para cambiarla.")

    def value(self):
        return self.field.text().strip()

class Card(CardWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("card")
        shadow = QGraphicsDropShadowEffect()
        shadow.setBlurRadius(16)
        shadow.setOffset(0, 2)
        shadow.setColor(QColor(0, 0, 0, 10))
        self.setGraphicsEffect(shadow)

class MenuCard(CardWidget):
    _instances = []

    def __init__(self, title, description, on_click, parent=None):
        super().__init__(parent)
        self.setObjectName("menuCard")
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self.on_click = on_click
        self.setMinimumHeight(78)

        self.color_normal = QColor(C["surface"])
        self.color_hover = QColor(C["overlay"])

        shadow = QGraphicsDropShadowEffect()
        shadow.setBlurRadius(12)
        shadow.setOffset(0, 2)
        shadow.setColor(QColor(0, 0, 0, 8))
        self.setGraphicsEffect(shadow)

        layout = QHBoxLayout(self)
        layout.setContentsMargins(20, 16, 20, 16)
        layout.setSpacing(16)

        text_col = QVBoxLayout()
        text_col.setSpacing(3)
        self.title_lbl = make_label(title, size=14, weight=600)
        self.desc_lbl = make_label(description, size=12, css_class="text-muted")
        self.desc_lbl.setWordWrap(True)
        text_col.addWidget(self.title_lbl)
        text_col.addWidget(self.desc_lbl)
        layout.addLayout(text_col, stretch=1)

        arrow_lbl = QLabel("›")
        arrow_lbl.setStyleSheet(f"color: {C['muted_light']}; background: transparent;")
        arrow_lbl.setFont(QFont("Segoe UI", 18))
        layout.addWidget(arrow_lbl)

        self.anim = QVariantAnimation(self)
        self.anim.setDuration(150)
        self.anim.valueChanged.connect(self._update_bg_color)

        MenuCard._instances.append(self)

    def _update_bg_color(self, color):
        color_hex = color.name() if isinstance(color, QColor) else color.name()
        self.setStyleSheet(f"QFrame#menuCard {{ background-color: {color_hex}; border: 1px solid {C['border']}; border-radius: 8px; }}")

    def update_theme(self):
        self.color_normal = QColor(C["surface"])
        self.color_hover = QColor(C["overlay"])
        self._update_bg_color(self.color_normal)
        self.title_lbl.style().unpolish(self.title_lbl)
        self.title_lbl.style().polish(self.title_lbl)
        self.desc_lbl.style().unpolish(self.desc_lbl)
        self.desc_lbl.style().polish(self.desc_lbl)

    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            self.on_click()

    def enterEvent(self, event):
        self.anim.stop()
        self.anim.setStartValue(self.color_normal)
        self.anim.setEndValue(self.color_hover)
        self.anim.start()
        super().enterEvent(event)

    def leaveEvent(self, event):
        self.anim.stop()
        self.anim.setStartValue(self.color_hover)
        self.anim.setEndValue(self.color_normal)
        self.anim.start()
        super().leaveEvent(event)

# ─── Pantalla principal (menú) ────────────────────────────────────────
class PantallaMenu(QWidget):

    def __init__(self, on_nuevo, on_actualizar, on_revisar, toggle_tema):
        super().__init__()
        self.toggle_tema = toggle_tema
        root = QVBoxLayout(self)
        root.setContentsMargins(44, 38, 44, 36)
        root.setSpacing(0)

        header_row = QHBoxLayout()
        title_col = QVBoxLayout()
        title_col.setSpacing(4)

        badge = QLabel("PROCESADOR DE ESTERILIZADO")
        badge.setProperty("class", "badge")
        badge.setFixedHeight(22)
        badge_wrap = QHBoxLayout()
        badge_wrap.addWidget(badge)
        badge_wrap.addStretch()

        self.title = make_label("Procesador de Autoclaves", size=20, weight=700)
        self.subtitle = make_label("Genera, actualiza y revisa reportes de ciclos de esterilización.",
                                   size=12, css_class="text-muted")

        title_col.addLayout(badge_wrap)
        title_col.addSpacing(6)
        title_col.addWidget(self.title)
        title_col.addWidget(self.subtitle)

        header_row.addLayout(title_col)
        header_row.addStretch()

        self.btn_tema = PushButton("Light" if current_theme == "light" else "Night")
        self.btn_tema.setProperty("class", "ghost")
        self.btn_tema.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_tema.clicked.connect(self.toggle_tema)
        header_row.addWidget(self.btn_tema)



        root.addLayout(header_row)
        root.addSpacing(28)

        self._divisor = QFrame()
        self._divisor.setFrameShape(QFrame.Shape.HLine)
        self._divisor.setStyleSheet(f"border: none; border-top: 1px solid {C['border_soft']};")
        self._divisor.setFixedHeight(1)
        root.addWidget(self._divisor)
        root.addSpacing(22)

        # Dashboard Overview Section
        dash_lbl = make_label("Resumen del Sistema", size=14, weight=600)
        root.addWidget(dash_lbl)
        root.addSpacing(16)

        stats_row = QHBoxLayout()
        stats_row.setSpacing(16)

        # Stat Card 1
        stat1 = CardWidget()
        stat1.setFixedHeight(110)
        stat1_layout = QVBoxLayout(stat1)
        stat1_layout.addWidget(make_label("Estado del Sistema", size=11, css_class="text-muted"))
        stat1_layout.addWidget(make_label("En línea", size=24, weight=700, css_class="text-success"))
        stat1_layout.addWidget(make_label("Listo para procesar PDFs", size=10, css_class="text-muted-light"))
        stats_row.addWidget(stat1)

        # Stat Card 2
        stat2 = CardWidget()
        stat2.setFixedHeight(110)
        stat2_layout = QVBoxLayout(stat2)
        stat2_layout.addWidget(make_label("Plantilla Base", size=11, css_class="text-muted"))

        has_template = bool(config.obtener_ruta_plantilla())
        t_text = "Configurada" if has_template else "Faltante"
        t_color = "text-success" if has_template else "text-warning"

        stat2_layout.addWidget(make_label(t_text, size=24, weight=700, css_class=t_color))
        stat2_layout.addWidget(make_label("Verifica 'Configuración' para cambiar", size=10, css_class="text-muted-light"))
        stats_row.addWidget(stat2)

        # Stat Card 3
        stat3 = CardWidget()
        stat3.setFixedHeight(110)
        stat3_layout = QVBoxLayout(stat3)
        stat3_layout.addWidget(make_label("Reglas de Desviación", size=11, css_class="text-muted"))
        stat3_layout.addWidget(make_label("Activas", size=24, weight=700, css_class="text-accent"))
        stat3_layout.addWidget(make_label("Motor de revisión cargado", size=10, css_class="text-muted-light"))
        stats_row.addWidget(stat3)

        root.addLayout(stats_row)
        root.addSpacing(32)

        # Quick Guide Section
        guide_lbl = make_label("Guía Rápida", size=14, weight=600)
        root.addWidget(guide_lbl)
        root.addSpacing(12)

        guide_card = CardWidget()
        guide_layout = QVBoxLayout(guide_card)
        guide_layout.setContentsMargins(20, 20, 20, 20)
        guide_layout.setSpacing(12)

        # Now using the callbacks dynamically!
        btn_nuevo = PushButton("Nuevo Reporte")
        if on_nuevo: btn_nuevo.clicked.connect(on_nuevo)

        btn_act = PushButton("Actualizar Reporte")
        if on_actualizar: btn_act.clicked.connect(on_actualizar)

        btn_rev = PushButton("Revisar Desviaciones")
        if on_revisar: btn_rev.clicked.connect(on_revisar)

        row1 = QHBoxLayout()
        row1.addWidget(make_label("1. Genera un Excel en blanco basado en la plantilla.", size=12))
        row1.addStretch()
        row1.addWidget(btn_nuevo)

        row2 = QHBoxLayout()
        row2.addWidget(make_label("2. Añade nuevos ciclos a un archivo existente.", size=12))
        row2.addStretch()
        row2.addWidget(btn_act)

        row3 = QHBoxLayout()
        row3.addWidget(make_label("3. Detecta parámetros fuera de rango en el Excel.", size=12))
        row3.addStretch()
        row3.addWidget(btn_rev)

        guide_layout.addLayout(row1)
        guide_layout.addLayout(row2)
        guide_layout.addLayout(row3)

        root.addWidget(guide_card)
        root.addStretch()

    def update_theme(self):
        self.title.style().unpolish(self.title)
        self.title.style().polish(self.title)
        self.subtitle.style().unpolish(self.subtitle)
        self.subtitle.style().polish(self.subtitle)
        self._divisor.setStyleSheet(
            f"border: none; border-top: 1px solid {C['border_soft']};"
        )
        self.btn_tema.setText("Light" if current_theme == "light" else "Night")

# ─── Pantalla de flujo de trabajo ────────────────────────────────────
class PantallaFlujo(QWidget):
    def __init__(self, modo, on_volver):
        super().__init__()
        self.modo           = modo
        self.on_volver      = on_volver
        self.worker         = None
        self.ruta_resultado = None
        self._executing     = False

        root = QVBoxLayout(self)
        root.setContentsMargins(36, 22, 36, 24)
        root.setSpacing(0)

        top = QHBoxLayout()
        btn_back = PushButton("← Volver")
        btn_back.setProperty("class", "ghost")
        btn_back.setCursor(Qt.CursorShape.PointingHandCursor)
        btn_back.clicked.connect(self._volver)
        top.addWidget(btn_back)
        top.addSpacing(14)

        titulos = {"nuevo": "Crear reporte nuevo",
                   "actualizar": "Actualizar reporte existente",
                   "revisar": "Revisar desviaciones"}
        self.titulo = make_label(titulos.get(modo, modo), size=14, weight=700)
        top.addWidget(self.titulo)
        top.addStretch()
        root.addLayout(top)
        root.addSpacing(14)

        btn_back.setToolTip("Volver al menú (Esc)")
        btn_back.setAccessibleName("Volver al menú")

        self._divisor_top = QFrame()
        self._divisor_top.setFrameShape(QFrame.Shape.HLine)
        self._divisor_top.setStyleSheet(f"border: none; border-top: 1px solid {C['border_soft']};")
        self._divisor_top.setFixedHeight(1)
        root.addWidget(self._divisor_top)
        root.addSpacing(16)

        card = Card()
        card_layout = QVBoxLayout(card)
        card_layout.setContentsMargins(22, 18, 22, 18)
        card_layout.setSpacing(14)
        self._build_inputs(card_layout)
        root.addWidget(card)
        root.addSpacing(14)

        self.btn_ejecutar = PrimaryPushButton("Ejecutar")
        self.btn_ejecutar.setProperty("class", "primary")
        self.btn_ejecutar.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_ejecutar.setToolTip("Iniciar el procesamiento (Ctrl+Enter)")
        self.btn_ejecutar.setAccessibleName("Ejecutar procesamiento")
        self.btn_ejecutar.clicked.connect(self._ejecutar)

        self.btn_cancelar = PushButton("Cancelar")
        self.btn_cancelar.setProperty("class", "ghost")
        self.btn_cancelar.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_cancelar.setVisible(False)
        self.btn_cancelar.setToolTip("Detener la operación en curso de forma segura")
        self.btn_cancelar.setAccessibleName("Cancelar operación")
        self.btn_cancelar.clicked.connect(self._cancelar)

        fila_botones = QHBoxLayout()
        fila_botones.setSpacing(10)
        fila_botones.addWidget(self.btn_ejecutar, stretch=1)
        fila_botones.addWidget(self.btn_cancelar)
        root.addLayout(fila_botones)
        root.addSpacing(10)

        # Indicador de etapa (stepper textual): comunica el paso actual del flujo
        # de forma más clara que solo la barra de progreso.
        self.lbl_estado = make_label("", size=11, css_class="text-muted-light")
        self.lbl_estado.setVisible(False)
        root.addWidget(self.lbl_estado)
        root.addSpacing(4)

        self.progress = ProgressBar()
        self.progress.setRange(0, 100)
        self.progress.setValue(0)
        self.progress.setFixedHeight(4)
        self.progress.setVisible(False)

        self.progress.setAccessibleName("Progreso del procesamiento")
        root.addWidget(self.progress)
        root.addSpacing(4)

        self.log_area = TextEdit()
        self.log_area.setReadOnly(True)
        self.log_area.setPlaceholderText("El registro de procesamiento aparecerá aquí...")
        self.log_area.setMinimumHeight(108)
        self.log_area.setVisible(False)
        self.log_area.setAccessibleName("Registro de procesamiento")
        root.addWidget(self.log_area, stretch=1)

        # Panel de resumen final (además del log/toasts): totales legibles de un
        # vistazo al terminar (ciclos procesados / desviaciones encontradas).
        self.resumen_lbl = make_label("", size=12, css_class="text-soft")
        self.resumen_lbl.setWordWrap(True)
        self.resumen_lbl.setVisible(False)
        root.addWidget(self.resumen_lbl)

        label_abrir = "Abrir archivo revisado" if modo == "revisar" else "Abrir archivo resultado"
        self.btn_abrir = PushButton(label_abrir)
        self.btn_abrir.setProperty("class", "success")
        self.btn_abrir.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_abrir.setVisible(False)
        self.btn_abrir.setToolTip("Abrir el archivo generado en Excel")
        self.btn_abrir.setAccessibleName("Abrir archivo resultado")
        self.btn_abrir.clicked.connect(self._abrir_resultado)
        root.addWidget(self.btn_abrir)

        # Contadores derivados del log para el resumen final.
        self._stats = {"procesados": 0, "desviaciones": None, "hojas": None}

    def _build_inputs(self, layout):
        if self.modo == "nuevo":
            self.sel_plantilla = FileSelector(
                "Plantilla Excel", mode="file",
                valor_inicial=config.obtener_ruta_plantilla(),
                on_select=config.guardar_ruta_plantilla
            )
            layout.addWidget(self.sel_plantilla)

            row = QHBoxLayout()
            row.setSpacing(12)
            self.sel_carpeta = FileSelector(
                "Carpeta de PDFs", mode="dir", filters="",
                valor_inicial=config.obtener_carpeta_pdfs(),
                on_select=config.guardar_carpeta_pdfs
            )
            row.addWidget(self.sel_carpeta, stretch=1)

            hora_col = QVBoxLayout()
            hora_col.setSpacing(6)
            lbl_hora = make_label("Hora inicio (HH:MM)", size=10, css_class="text-soft")
            hora_col.addWidget(lbl_hora)
            self.hora_input = LineEdit()
            self.hora_input.setPlaceholderText("ej. 19:50")
            self.hora_input.setMaxLength(5)
            self.hora_input.setFixedWidth(104)
            self.hora_input.setText(config.obtener_hora_inicio())
            self.hora_input.textChanged.connect(self._guardar_hora)
            hora_col.addWidget(self.hora_input)
            row.addLayout(hora_col)

            layout.addLayout(row)


            self.sel_salida = FileSelector(
                "Guardar reporte como", mode="save",
                filters="Excel (*.xlsx *.xlsm);;Todos (*)",
                sugerir_nombre=self._sugerir_nombre_salida
            )
            layout.addWidget(self.sel_salida)

        elif self.modo == "actualizar":
            self.sel_existente = FileSelector(
                "Reporte Excel existente", mode="file",
                valor_inicial=config.obtener_ruta_existente(),
                on_select=config.guardar_ruta_existente
            )
            layout.addWidget(self.sel_existente)

            self.sel_plantilla = FileSelector(
                "Plantilla Excel (hoja en blanco)", mode="file",
                valor_inicial=config.obtener_ruta_plantilla(),
                on_select=config.guardar_ruta_plantilla
            )
            layout.addWidget(self.sel_plantilla)

            row = QHBoxLayout()
            row.setSpacing(12)
            self.sel_carpeta = FileSelector(
                "Carpeta con todos los PDFs", mode="dir", filters="",
                valor_inicial=config.obtener_carpeta_pdfs(),
                on_select=config.guardar_carpeta_pdfs
            )
            row.addWidget(self.sel_carpeta, stretch=1)

            hora_col = QVBoxLayout()
            hora_col.setSpacing(6)
            lbl_hora2 = make_label("Hora inicio (HH:MM)", size=10, css_class="text-soft")
            hora_col.addWidget(lbl_hora2)
            self.hora_input = LineEdit()
            self.hora_input.setPlaceholderText("ej. 19:50")
            self.hora_input.setMaxLength(5)
            self.hora_input.setFixedWidth(104)
            self.hora_input.setText(config.obtener_hora_inicio())
            self.hora_input.textChanged.connect(self._guardar_hora)
            hora_col.addWidget(self.hora_input)
            row.addLayout(hora_col)

            layout.addLayout(row)


        else:  # revisar
            self.sel_excel = FileSelector(
                "Reporte Excel a revisar", mode="file",
                filters="Excel (*.xlsx *.xlsm);;Todos (*)",
                valor_inicial=config.obtener_ruta_existente(),
                on_select=config.guardar_ruta_existente
            )
            layout.addWidget(self.sel_excel)

            self.chk_actualizar = CheckBox('Actualizar la hoja Desviaciones')
            self.chk_actualizar.setCursor(Qt.CursorShape.PointingHandCursor)
            self.chk_actualizar.toggled.connect(self._toggle_actualizar_in_situ)
            layout.addWidget(self.chk_actualizar)

            self.sel_salida = FileSelector(
                "Guardar resultado como", mode="save",
                filters="Excel (*.xlsx *.xlsm);;Todos (*)",
                sugerir_nombre=self._sugerir_nombre_revision
            )
            layout.addWidget(self.sel_salida)

            self.lbl_info_revisar = make_label(
                'Si dejas la ruta vacía, se guarda junto al original con el sufijo "_Revisión desviaciones".',
                size=11, css_class="text-muted")
            self.lbl_info_revisar.setWordWrap(True)
            layout.addWidget(self.lbl_info_revisar)

    def _guardar_hora(self, texto):
        texto = texto.strip()
        if not texto:
            return
        try:
            config.guardar_hora_inicio(texto)
        except ValueError:
            pass

    def _toggle_actualizar_in_situ(self, checked):
        self.sel_salida.setEnabled(not checked)
        if checked:
            self.lbl_info_revisar.setText('Se sobrescribirá la hoja "Desviaciones" en el archivo seleccionado '
                                          'arriba (se reemplaza la anterior si ya existía).')
        else:
            self.lbl_info_revisar.setText('Si dejas la ruta vacía, se guarda junto al original con el '
                                          'sufijo "_Revisión desviaciones".')

    def _sugerir_nombre_salida(self):
        ruta_plantilla = self.sel_plantilla.value()
        if not ruta_plantilla:
            return ""
        base_nombre, ext = os.path.splitext(os.path.basename(ruta_plantilla))
        ext_salida = ".xlsm" if ext.lower() in (".xlsm", ".xltm") else ".xlsx"
        nombre = f"{base_nombre}{ext_salida}"
        carpeta = self.sel_carpeta.value()
        return os.path.join(carpeta, nombre) if carpeta else nombre

    def _sugerir_nombre_revision(self):
        ruta_excel = self.sel_excel.value()
        if not ruta_excel:
            return ""
        base, ext = os.path.splitext(ruta_excel)
        ext_salida = ".xlsm" if ext.lower() in (".xlsm", ".xltm") else ".xlsx"
        return f"{base}_Revisión esterilizado{ext_salida}"

    # ── Validación ──────────────────────────────────────────────────
    def _validar(self):
        errores = []
        if self.modo == "revisar":
            if not self.sel_excel.value():
                errores.append(("Selecciona el reporte Excel a revisar", self.sel_excel.field))
        else:
            hora = self.hora_input.text().strip()
            if not hora:
                errores.append(("Ingresa la hora de inicio", self.hora_input))
            else:
                try:
                    config.parsear_hora(hora)
                except ValueError:
                    errores.append(("Formato inválido. Usa HH:MM (ej. 08:30)", self.hora_input))

            if self.modo == "nuevo":
                for sel, nombre in [(self.sel_plantilla, "plantilla"),
                                    (self.sel_carpeta, "carpeta de PDFs"),
                                    (self.sel_salida, "ruta de salida")]:
                    if not sel.value():
                        errores.append((f"Selecciona la {nombre}", sel.field))
            else:
                for sel, nombre in [(self.sel_existente, "reporte existente"),
                                    (self.sel_plantilla, "plantilla"),
                                    (self.sel_carpeta, "carpeta de PDFs")]:
                    if not sel.value():
                        errores.append((f"Selecciona el/la {nombre}", sel.field))
        if errores:
            for msg, widget in errores:
                original = widget.styleSheet()
                widget.setStyleSheet(original + f" border-color: {C['error']};")
                QTimer.singleShot(2000, lambda w=widget, orig=original: w.setStyleSheet(orig))
            mensaje = "\n".join(msg for msg, _ in errores[:3])
            if len(errores) > 3:
                mensaje += f"\n...y {len(errores)-3} errores más."
            show_toast(self, "Errores de validación", mensaje, "error")
            return False
        return True

    def _ejecutar(self):
        if self._executing:
            return
        if not self._validar():
            return

        try:
            config.guardar_hora_inicio(self.hora_input.text().strip())
        except ValueError:
            pass

        if self.modo == "revisar":
            actualizar_in_situ = self.chk_actualizar.isChecked()
            ruta_salida = (self.sel_excel.value() if actualizar_in_situ else (self.sel_salida.value() or None))
            params = {
                "ruta_excel": self.sel_excel.value(),
                "ruta_salida": ruta_salida,
                "actualizar_in_situ": actualizar_in_situ,
            }
        else:
            hora_raw = self.hora_input.text().strip().replace('.', ':')
            palabras_descarte = config.obtener_palabras_descarte()
            if self.modo == "nuevo":
                params = {
                    "ruta_plantilla": self.sel_plantilla.value(),
                    "hora_inicio": hora_raw,
                    "carpeta_pdfs": self.sel_carpeta.value(),
                    "ruta_salida": self.sel_salida.value(),
                    "palabras_descarte": palabras_descarte,
                }
            else:
                params = {
                    "ruta_existente": self.sel_existente.value(),
                    "ruta_plantilla": self.sel_plantilla.value(),
                    "hora_inicio": hora_raw,
                    "carpeta_pdfs": self.sel_carpeta.value(),
                    "palabras_descarte": palabras_descarte,
                }

        self._executing = True
        self._stats = {"procesados": 0, "desviaciones": None, "hojas": None}
        self.log_area.clear()
        self.log_area.setVisible(True)
        self.resumen_lbl.setVisible(False)
        self.lbl_estado.setVisible(True)
        self.lbl_estado.setText("●  Iniciando…")
        self.progress.setVisible(True)
        self.progress.setValue(0)
        self.btn_abrir.setVisible(False)
        self.btn_ejecutar.setEnabled(False)
        self.btn_ejecutar.setText("Procesando...")
        self.btn_cancelar.setVisible(True)
        self.btn_cancelar.setEnabled(True)
        self.btn_cancelar.setText("Cancelar")
        self._log_msg("Iniciando procesamiento...", color=C["accent"])

        self.worker = WorkerThread(self.modo, params)
        self.worker.log.connect(self._on_log)
        self.worker.progreso.connect(self._on_progreso)
        self.worker.terminado.connect(self._on_terminado)
        self.worker.error.connect(self._on_error)
        self.worker.cancelado.connect(self._on_cancelado)
        self.worker.stats.connect(self._on_stats)
        self.worker.start()

    def _volver(self):
        if self._executing:
            return
        self.on_volver()

    def _cancelar(self):
        if self.worker and self.worker.isRunning():
            self.worker.solicitar_cancelacion()
            self.btn_cancelar.setEnabled(False)
            self.btn_cancelar.setText("Cancelando...")
            self._log_msg("Cancelando… esperando a que termine la operación.", color=C["warning"])
            show_toast(self, "Cancelación", "Solicitud enviada. Finalizando tarea actual...", "warning")

    def _on_log(self, msg):
        if any(x in msg for x in ["[OK]", "OK", "completado", "guardado", "Preparado", "Procesado"]):
            color = C["success"]
        elif any(x in msg for x in ["Error", "error", "No se pudo", "invalido"]):
            color = C["error"]
        elif any(x in msg for x in ["Omitido", "Sin metadatos", "advertencia"]):
            color = C["warning"]
        else:
            color = C["text"]
        self._log_msg(msg, color=color)

    def _on_stats(self, data):
        if "procesados" in data:
            self._stats["procesados"] = data["procesados"]
        if "desviaciones" in data:
            self._stats["desviaciones"] = data["desviaciones"]

    def _texto_etapa(self, valor):
        """Etiqueta de etapa (stepper textual) según el porcentaje de avance."""
        if self.modo == "revisar":
            if valor < 15:  return "Abriendo archivo…"
            if valor < 70:  return "Analizando hojas y detectando desviaciones…"
            if valor < 100: return "Generando hoja de desviaciones…"
            return "Completado"
        if valor < 10:  return "Escaneando PDFs…"
        if valor < 60:  return "Procesando ciclos…"
        if valor < 85:  return "Reordenando y guardando…"
        if valor < 100: return "Finalizando…"
        return "Completado"

    def _on_progreso(self, valor):
        valor = min(max(valor, 0), 100)
        self.progress.setValue(valor)
        self.lbl_estado.setText("●  " + self._texto_etapa(valor))

    def _on_terminado(self, ruta):
        self.ruta_resultado = ruta
        self.progress.setValue(100)
        self.lbl_estado.setText("●  Completado")
        self._log_msg("")
        self._log_msg("Procesamiento completado exitosamente.", color=C["success"])
        self._mostrar_resumen()
        self._reset_ui_on_finish()
        self.btn_abrir.setVisible(True)
        show_toast(self, "Éxito", "El reporte ha sido generado correctamente.", "success")

    def _mostrar_resumen(self):
        partes = []
        if self.modo in ("nuevo", "actualizar") and self._stats["procesados"]:
            partes.append(f"{self._stats['procesados']} ciclo(s) procesado(s)")
        if self._stats["desviaciones"] is not None:
            partes.append(f"{self._stats['desviaciones']} desviación(es) detectada(s)")
        if partes:
            self.resumen_lbl.setText("Resumen:   " + "     •     ".join(partes))
            self.resumen_lbl.setVisible(True)

    def _on_error(self, msg):
        self.lbl_estado.setText("●  Error")
        self._log_msg(f"Error: {msg}", color=C["error"])
        self._reset_ui_on_finish()
        show_toast(self, "Error", msg, "error")

    def _on_cancelado(self):
        self.lbl_estado.setText("●  Cancelado")
        self._log_msg("Operación cancelada por el usuario.", color=C["warning"])
        self.progress.setValue(0)
        self._reset_ui_on_finish()
        show_toast(self, "Cancelado", "La operación fue cancelada.", "warning")

    def _reset_ui_on_finish(self):
        self._executing = False
        self.btn_ejecutar.setEnabled(True)
        self.btn_ejecutar.setText("Ejecutar")
        self.btn_cancelar.setVisible(False)
        self.btn_cancelar.setEnabled(True)
        self.btn_cancelar.setText("Cancelar")

    def _log_msg(self, msg, color=None):
        c = color if color else C["text"]
        doc = self.log_area.document()
        if doc.blockCount() > MAX_LOG_LINES:
            cursor = QTextCursor(doc.begin())
            for _ in range(doc.blockCount() - MAX_LOG_LINES):
                cursor.movePosition(QTextCursor.MoveOperation.Down, QTextCursor.MoveMode.KeepAnchor)
            cursor.removeSelectedText()
            cursor.movePosition(QTextCursor.MoveOperation.End)
        # Escapar el texto: los mensajes de log (rutas, errores de Excel/COM)
        # pueden contener '<', '>' o '&' que romperían el render HTML del QTextEdit
        # o inyectarían marcado no deseado. Los espacios iniciales se preservan.
        seguro = html.escape(msg)
        self.log_area.insertHtml(
            f'<span style="color:{c}; font-family: Consolas, monospace; font-size:13px;">'
            f'{seguro}</span><br>'
        )
        self.log_area.moveCursor(QTextCursor.MoveOperation.End)

    def _abrir_resultado(self):
        if self.ruta_resultado and os.path.exists(self.ruta_resultado):
            os.startfile(self.ruta_resultado)
        else:
            self._log_msg("No se pudo encontrar el archivo resultado.", color=C["warning"])

    def reset(self):
        self._executing = False
        self._stats = {"procesados": 0, "desviaciones": None, "hojas": None}
        self.log_area.clear()
        self.log_area.setVisible(False)
        self.lbl_estado.setVisible(False)
        self.lbl_estado.setText("")
        self.resumen_lbl.setVisible(False)
        self.resumen_lbl.setText("")
        self.progress.setValue(0)
        self.progress.setVisible(False)
        self.btn_abrir.setVisible(False)
        self.btn_ejecutar.setEnabled(True)
        self.btn_ejecutar.setText("Ejecutar")
        self.btn_cancelar.setVisible(False)
        self.btn_cancelar.setEnabled(True)
        self.btn_cancelar.setText("Cancelar")
        self.ruta_resultado = None

    def update_theme(self):
        """Refresca los elementos con estilo embebido (no cubiertos por el
        stylesheet global) tras un cambio de tema: el divisor superior toma el
        color de borde del tema activo."""
        self._divisor_top.setStyleSheet(
            f"border: none; border-top: 1px solid {C['border_soft']};"
        )
        for lbl in (self.titulo, self.lbl_estado, self.resumen_lbl):
            lbl.style().unpolish(lbl)
            lbl.style().polish(lbl)

# ─── Ventana principal ───────────────────────────────────────────────
class VentanaPrincipal(FluentWindow):
    def __init__(self, on_nuevo=None, on_actualizar=None, on_revisar=None, toggle_tema=None):
        super().__init__()
        self.toggle_tema = toggle_tema
        self.setWindowTitle("Procesador de Esterilizado")
        self.setMinimumSize(720, 520)
        self.resize(720, 520)
        self._centrar()

        self.pantalla_menu = PantallaMenu(
            on_nuevo=lambda: self.switchTo(self.pantalla_nuevo),
            on_actualizar=lambda: self.switchTo(self.pantalla_actualizar),
            on_revisar=lambda: self.switchTo(self.pantalla_revisar),
            toggle_tema=self.alternar_tema
        )
        self.pantalla_menu.setObjectName("PantallaMenu")
        self.pantalla_nuevo = PantallaFlujo("nuevo", on_volver=lambda: self.switchTo(self.pantalla_menu))
        self.pantalla_nuevo.setObjectName("PantallaNuevo")
        self.pantalla_actualizar = PantallaFlujo("actualizar", on_volver=lambda: self.switchTo(self.pantalla_menu))
        self.pantalla_actualizar.setObjectName("PantallaActualizar")
        self.pantalla_revisar = PantallaFlujo("revisar", on_volver=lambda: self.switchTo(self.pantalla_menu))
        self.pantalla_revisar.setObjectName("PantallaRevisar")

        self.pantalla_config = PantallaConfiguracion(toggle_tema_fn=self.alternar_tema)

        self.addSubInterface(self.pantalla_menu, FluentIcon.HOME, 'Inicio')
        self.addSubInterface(self.pantalla_nuevo, FluentIcon.ADD, 'Nuevo Reporte')
        self.addSubInterface(self.pantalla_actualizar, FluentIcon.UPDATE, 'Actualizar Reporte')
        self.addSubInterface(self.pantalla_revisar, FluentIcon.SEARCH, 'Revisar Desviaciones')
        self.addSubInterface(self.pantalla_config, FluentIcon.SETTING, 'Configuración', position=NavigationItemPosition.BOTTOM)

        # Atajos de teclado
        QShortcut(QKeySequence("Ctrl+1"), self, lambda: self._ir_a("nuevo"))
        QShortcut(QKeySequence("Ctrl+2"), self, lambda: self._ir_a("actualizar"))
        QShortcut(QKeySequence("Ctrl+3"), self, lambda: self._ir_a("revisar"))
        QShortcut(QKeySequence("Ctrl+Return"), self, self._ejecutar_actual)
        QShortcut(QKeySequence("Escape"), self, self._ir_a_menu)

    def _ejecutar_actual(self):
        actual = self.stackedWidget.currentWidget()
        if hasattr(actual, "_ejecutar"):
            actual._ejecutar()

    def _centrar(self):
        screen = QApplication.primaryScreen().availableGeometry()
        x = (screen.width() - 720) // 2
        y = (screen.height() - 520) // 2
        self.move(x, y)

    def _ir_a(self, modo):
        if modo == "nuevo":
            self.pantalla_nuevo.reset()
            self.switchTo(self.pantalla_nuevo)
        elif modo == "actualizar":
            self.pantalla_actualizar.reset()
            self.switchTo(self.pantalla_actualizar)
        elif modo == "revisar":
            self.pantalla_revisar.reset()
            self.switchTo(self.pantalla_revisar)

    def switchTo(self, interface):
        self.stackedWidget.setCurrentWidget(interface)
        # Update the navigation interface active item
        self.navigationInterface.setCurrentItem(interface.objectName())

    def _ir_a_menu(self):
        self.switchTo(self.pantalla_menu)

    def alternar_tema(self):
        global current_theme, C
        current_theme = "dark" if current_theme == "light" else "light"
        C = C_DARK.copy() if current_theme == "dark" else C_LIGHT.copy()

        setTheme(Theme.DARK if current_theme == "dark" else Theme.LIGHT)

        try:
            config.guardar_tema(current_theme)
        except Exception:
            pass

        app = QApplication.instance()
        palette = QPalette()
        palette.setColor(QPalette.ColorRole.Window, QColor(C["bg"]))
        palette.setColor(QPalette.ColorRole.WindowText, QColor(C["text"]))
        palette.setColor(QPalette.ColorRole.Base, QColor(C["surface"]))
        palette.setColor(QPalette.ColorRole.AlternateBase, QColor(C["overlay"]))
        palette.setColor(QPalette.ColorRole.ToolTipBase, QColor(C["surface"]))
        palette.setColor(QPalette.ColorRole.ToolTipText, QColor(C["text"]))
        palette.setColor(QPalette.ColorRole.Text, QColor(C["text"]))
        palette.setColor(QPalette.ColorRole.Button, QColor(C["surface"]))
        palette.setColor(QPalette.ColorRole.ButtonText, QColor(C["text"]))
        palette.setColor(QPalette.ColorRole.BrightText, QColor(C["accent"]))
        palette.setColor(QPalette.ColorRole.Highlight, QColor(C["accent"]))
        palette.setColor(QPalette.ColorRole.HighlightedText, QColor("#ffffff"))
        app.setPalette(palette)
        app.setStyleSheet(generar_estilos(C))

        for card in MenuCard._instances:
            card.update_theme()
        self.pantalla_menu.update_theme()
        for pantalla in (self.pantalla_nuevo, self.pantalla_actualizar, self.pantalla_revisar):
            pantalla.update_theme()

# ─── Punto de entrada ────────────────────────────────────────────────
def main():
    app = QApplication(sys.argv)
    app.setStyle("Fusion")

    global current_theme, C
    saved_theme = config.obtener_tema()
    if saved_theme == "dark":
        current_theme = "dark"
        C = C_DARK.copy()
        setTheme(Theme.DARK)
    else:
        current_theme = "light"
        C = C_LIGHT.copy()
        setTheme(Theme.LIGHT)

    app.setStyleSheet(generar_estilos(C))
    ventana = VentanaPrincipal()
    ventana.show()
    sys.exit(app.exec())

class PantallaConfiguracion(SmoothScrollArea):
    """Pantalla para ajustes globales de la aplicación."""
    def __init__(self, toggle_tema_fn=None):
        super().__init__()
        self.setObjectName("PantallaConfiguracion")
        self.setWidgetResizable(True)
        self.setStyleSheet("QScrollArea {background: transparent; border: none;}")
        self.toggle_tema_fn = toggle_tema_fn

        container = QWidget()
        container.setStyleSheet("background: transparent;")

        # We use an QVBoxLayout to hold collapsible cards
        self.expandLayout = QVBoxLayout(container)
        self.expandLayout.setContentsMargins(44, 38, 44, 36)
        self.expandLayout.setSpacing(24)

        titulo = make_label("Configuración", size=24, weight=700)
        self.expandLayout.addWidget(titulo)

        # 1. Ajustes de Tema
        self.card_tema = CardWidget()
        tema_layout = QVBoxLayout(self.card_tema)
        tema_layout.setContentsMargins(24, 24, 24, 24)
        tema_layout.addWidget(make_label("Apariencia y Tema", size=14, weight=600))
        tema_layout.addWidget(make_label("Configura el modo oscuro o claro de la aplicación", size=11, css_class="text-muted"))

        self.btn_tema = PushButton("Cambiar Tema")
        self.btn_tema.setFixedWidth(160)
        self.btn_tema.setCursor(Qt.CursorShape.PointingHandCursor)
        if self.toggle_tema_fn:
            self.btn_tema.clicked.connect(self.toggle_tema_fn)
        tema_layout.addSpacing(8)
        tema_layout.addWidget(self.btn_tema)

        self.expandLayout.addWidget(self.card_tema)

        # 2. Descarte de PDFs
        self.card_filtros = CardWidget()
        filtros_layout = QVBoxLayout(self.card_filtros)
        filtros_layout.setContentsMargins(24, 24, 24, 24)
        filtros_layout.addWidget(make_label("Filtros de PDFs", size=14, weight=600))
        filtros_layout.addWidget(make_label("Palabras a descartar en nombres de PDF (separadas por coma)", size=11, css_class="text-muted"))

        self.descarte_input = LineEdit()
        self.descarte_input.setText(config.obtener_palabras_descarte_texto())
        self.descarte_input.setPlaceholderText("ej. prueba, limpieza, xxx")

        filtros_layout.addSpacing(8)
        filtros_layout.addWidget(self.descarte_input)

        self.expandLayout.addWidget(self.card_filtros)

        # 3. Parámetros Críticos
        self.card_params = CardWidget()
        c_params_layout = QVBoxLayout(self.card_params)
        c_params_layout.setContentsMargins(24, 24, 24, 24)
        c_params_layout.addWidget(make_label("Parámetros Críticos y Umbrales", size=14, weight=600))
        c_params_layout.addWidget(make_label("Modifica los valores numéricos bajo los cuales se disparan las desviaciones", size=11, css_class="text-muted"))
        c_params_layout.addSpacing(8)

        params_layout = QFormLayout()
        params_layout.setSpacing(12)

        self.inputs_params = {}
        parametros = config.obtener_parametros_criticos()

        campos_prm = [
            ("param_temp_critica_f4", "Temp. crítica Holding (°C)", "Desviación si temp < valor"),
            ("param_dif_temp_max", "Dif. máximo temp. (°C)", "Diferencia > valor"),
            ("param_dif_temp_min", "Dif. mínimo temp. (°C)", "Diferencia <= valor (usar negativos)"),
            ("param_presion_min_f3", "Presión mín. Fase 3 (Bar)", "Límite calentamiento"),
            ("param_presion_min_f4", "Presión mín. Fase 4 (Bar)", "Límite holding"),
            ("param_caudal_min_123", "Caudal mín. Fases 1,2,3 (m3/h)", "Límite calentamiento"),
            ("param_caudal_min_4", "Caudal mín. Fase 4 (m3/h)", "Límite holding"),
        ]

        for k, label_text, tooltip in campos_prm:
            inp = LineEdit()
            inp.setText(str(parametros.get(k, "")))
            inp.setMinimumWidth(80)
            inp.setMaximumWidth(150)
            self.inputs_params[k] = inp

            lbl_widget = QWidget()
            lbl_layout = QVBoxLayout(lbl_widget)
            lbl_layout.setContentsMargins(0, 0, 0, 0)
            lbl_layout.setSpacing(2)

            lbl_layout.addWidget(make_label(label_text, size=12))
            lbl_layout.addWidget(make_label(tooltip, size=10, css_class="text-muted"))
            params_layout.addRow(lbl_widget, inp)

        c_params_layout.addLayout(params_layout)
        self.expandLayout.addWidget(self.card_params)

        # 4. Textos de Desviaciones
        self.card_textos = CardWidget()
        c_textos_layout = QVBoxLayout(self.card_textos)
        c_textos_layout.setContentsMargins(24, 24, 24, 24)
        c_textos_layout.addWidget(make_label("Plantillas de Textos", size=14, weight=600))
        c_textos_layout.addWidget(make_label("Personaliza los textos generados en el reporte", size=11, css_class="text-muted"))
        c_textos_layout.addSpacing(8)

        textos_layout = QFormLayout()
        textos_layout.setSpacing(12)

        self.inputs_texto = {}
        plantillas = config.obtener_plantillas_textos()

        campos_txt = [
            ("txt_hdr_n_alto", "Temp Registrador > Digital (> Max) [Header]"),
            ("txt_ora_n_alto", "Temp Registrador > Digital (> Max) [Oración]"),
            ("txt_hdr_n_bajo", "Temp Registrador < Digital (< Min) [Header]"),
            ("txt_ora_n_bajo", "Temp Registrador < Digital (< Min) [Oración]"),
            ("txt_hdr_r_alto", "Temp Registrador > Controlador (> Max) [Header]"),
            ("txt_ora_r_alto", "Temp Registrador > Controlador (> Max) [Oración]"),
            ("txt_hdr_r_bajo", "Temp Registrador < Controlador (< Min) [Header]"),
            ("txt_ora_r_bajo", "Temp Registrador < Controlador (< Min) [Oración]"),
            ("txt_hdr_reg_crit", "Temp Registrador Crítica (Holding) [Header]"),
            ("txt_ora_reg_crit", "Temp Registrador Crítica (Holding) [Oración]"),
            ("txt_hdr_dig_crit", "Temp Digital Crítica (Holding) [Header]"),
            ("txt_ora_dig_crit", "Temp Digital Crítica (Holding) [Oración]"),
            ("txt_bar_f3", "Presión Fase 3"),
            ("txt_bar_f4", "Presión Fase 4"),
            ("txt_caudal_123", "Caudal Fases 1,2,3"),
            ("txt_caudal_4", "Caudal Fase 4"),
        ]

        for k, label in campos_txt:
            inp = LineEdit()
            inp.setText(plantillas.get(k, ""))
            self.inputs_texto[k] = inp
            textos_layout.addRow(make_label(label, size=11), inp)

        c_textos_layout.addLayout(textos_layout)
        self.expandLayout.addWidget(self.card_textos)

        # Botones de acción
        btn_layout = QHBoxLayout()
        btn_layout.setContentsMargins(0, 0, 0, 0)

        self.btn_guardar = PrimaryPushButton("Guardar Cambios")
        self.btn_guardar.setFixedWidth(160)
        self.btn_guardar.clicked.connect(self.guardar_cambios)

        self.btn_restablecer = PushButton("Restablecer Valores por Defecto")
        self.btn_restablecer.setFixedWidth(240)
        self.btn_restablecer.clicked.connect(self.restablecer_valores)

        btn_layout.addWidget(self.btn_guardar)
        btn_layout.addSpacing(12)
        btn_layout.addWidget(self.btn_restablecer)
        btn_layout.addStretch()

        self.expandLayout.addSpacing(16)
        self.expandLayout.addLayout(btn_layout)

        self.setWidget(container)

    def restablecer_valores(self):
        config.restablecer_configuracion()

        # Recargar UI con defaults
        self.descarte_input.setText(config.obtener_palabras_descarte_texto())

        parametros = config.obtener_parametros_criticos()
        for k, inp in self.inputs_params.items():
            inp.setText(str(parametros.get(k, "")))

        plantillas = config.obtener_plantillas_textos()
        for k, inp in self.inputs_texto.items():
            inp.setText(plantillas.get(k, ""))

        InfoBar.success(
            title='Valores restablecidos',
            content="La configuración ha vuelto a sus valores por defecto originales.",
            orient=Qt.Orientation.Horizontal,
            isClosable=True,
            position=InfoBarPosition.TOP_RIGHT,
            duration=3000,
            parent=self.window()
        )

    def guardar_cambios(self):
        config.guardar_palabras_descarte(self.descarte_input.text())

        nuevas_plantillas = {k: inp.text() for k, inp in self.inputs_texto.items()}
        config.guardar_plantillas_textos(nuevas_plantillas)

        nuevos_params = {k: inp.text() for k, inp in self.inputs_params.items()}
        config.guardar_parametros_criticos(nuevos_params)

        InfoBar.success(
            title='Guardado exitoso',
            content="La configuración ha sido actualizada.",
            orient=Qt.Orientation.Horizontal,
            isClosable=True,
            position=InfoBarPosition.TOP_RIGHT,
            duration=3000,
            parent=self.window()
        )
if __name__ == '__main__':
    main()
