import re

with open("ui.py", "r") as f:
    content = f.read()

# Fix PantallaMenu init signature to properly accept all expected args passed by VentanaPrincipal
# And gracefully handle them by storing them or ignoring them without crashing
old_menu_init = """    def __init__(self):
        super().__init__()"""

new_menu_init = """    def __init__(self, on_nuevo=None, on_actualizar=None, on_revisar=None, toggle_tema=None):
        super().__init__()
        self.toggle_tema = toggle_tema"""

content = content.replace(old_menu_init, new_menu_init)

with open("ui.py", "w") as f:
    f.write(content)
