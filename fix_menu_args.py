import re

with open("ui.py", "r") as f:
    content = f.read()

# We need to make sure VentanaPrincipal passes the callbacks to PantallaMenu
# because currently we have `self.pantalla_menu = PantallaMenu()` and it fails to match correctly if arguments are required but not provided.

old_menu_init = """        self.pantalla_menu = PantallaMenu()"""
new_menu_init = """        self.pantalla_menu = PantallaMenu(
            on_nuevo=lambda: self.switchTo(self.pantalla_nuevo),
            on_actualizar=lambda: self.switchTo(self.pantalla_actualizar),
            on_revisar=lambda: self.switchTo(self.pantalla_revisar),
            toggle_tema=self.alternar_tema
        )"""

content = content.replace(old_menu_init, new_menu_init)

with open("ui.py", "w") as f:
    f.write(content)
