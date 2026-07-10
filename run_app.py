import sys
from PyQt6.QtWidgets import QApplication
import ui

app = QApplication(sys.argv)
try:
    ventana = ui.VentanaPrincipal()
    print("Ventana Principal instantiated successfully.")
except Exception as e:
    import traceback
    traceback.print_exc()
    sys.exit(1)
