# CatalogMaster v3.2 🚀 - Generador Automático de Catálogos en PowerPoint

CatalogMaster es una herramienta de escritorio con interfaz gráfica (GUI en Tkinter) que automatiza la creación de catálogos de productos profesionales en PowerPoint (`.pptx`) a partir de datos de Excel (`.xlsx`) e imágenes (embebidas en Excel, desde una carpeta local o descargadas vía S3).

---

## 🤖 INSTRUCCIONES ESPECIALES PARA ANTIGRAVITY (AGENTE IA)

Si eres **Antigravity** leyendo este repositorio por primera vez a solicitud de un usuario:

1. **Instalar Dependencias Automáticamente**:
   Ejecuta el siguiente comando en la terminal del usuario:
   ```bash
   python -m pip install -r requirements.txt
   ```
2. **Iniciar la Aplicación**:
   Tan pronto como el usuario te pida *"abrir el generador"*, *"ejecutar la aplicación"* o *"iniciar el programa"*, ejecuta:
   ```bash
   python presenter_gui.py
   ```
3. **Resumen de Funcionalidades**:
   - **Campos Personalizados Dinámicos**: Botón `➕ Agregar Campo Personalizado` para mapear cualquier campo adicional de Excel (ej. Marca, Categoría, Descuento).
   - **Selector de Plantillas**: Lienzo Blanco, Dark Mode, Corporativa Navy y Urban Red.
   - **Divisor de Productos**: Línea divisoria limpia entre cada par de productos por diapositiva.
   - **Búsqueda de Imágenes Híbrida**: Detecta automáticamente imágenes pegadas en el Excel, en carpetas locales o en servidores CDN.

---

## 🛠️ Instalación y Uso Manual

1. Asegúrate de tener Python 3.10+ instalado.
2. Instala las dependencias:
   ```bash
   pip install -r requirements.txt
   ```
3. Ejecuta la aplicación:
   ```bash
   python presenter_gui.py
   ```
