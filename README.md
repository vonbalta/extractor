# extractor
### Troubleshooting
Si en Linux (Ubuntu/Debian) la aplicación crashea al iniciar mostrando un error similar a `qt.qpa.plugin: Could not load the Qt platform plugin "xcb"`, asegúrate de tener instaladas las librerías del sistema necesarias para PyQt6:

```bash
sudo apt-get update
sudo apt-get install libxcb-cursor0
```
