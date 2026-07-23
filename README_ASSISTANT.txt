# CatalogMaster v2.2 - Guía de Instalación y Uso 🚀

Esta herramienta automatiza la creación de catálogos profesionales en PowerPoint (.pptx) utilizando datos de Excel (.xlsx) e imágenes (ya sean embebidas en el archivo o alojadas en el servidor S3).

## 🛠️ Proceso de Instalación Paso a Paso

Sigue estos pasos para configurar el entorno y ejecutar la aplicación:

### 1. Requisitos Previos
Asegúrate de tener instalado **Python 3.10 o superior**. Puedes verificarlo ejecutando:
```bash
python3 --version
```

### 2. Preparación del Entorno (Virtual Env)
Es recomendable usar un entorno virtual para evitar conflictos con otras librerías de tu sistema.

**En Linux/macOS:**
```bash
# Crear el entorno virtual
python3 -m venv venv

# Activar el entorno
source venv/bin/activate
```

**En Windows:**
```bash
# Crear el entorno virtual
python -m venv venv

# Activar el entorno
.\venv\Scripts\activate
```

### 3. Instalación de Dependencias
Una vez activado el entorno, instala las librerías necesarias utilizando el archivo `requirements.txt`:
```bash
pip install -r requirements.txt
```

*Librerías instaladas:*
- `openpyxl`: Lectura de archivos Excel.
- `python-pptx`: Generación de diapositivas de PowerPoint.
- `requests`: Descarga de imágenes desde la web.
- `Pillow`: Procesamiento y redimensionamiento de imágenes.

### 4. Ejecución del Programa
Para iniciar la aplicación con la interfaz gráfica, ejecuta el siguiente comando:
```bash
python presenter_gui.py
```

---

## 📋 Guía de Uso
1. **Excel File**: Selecciona el archivo `.xlsx`. El programa detectará automáticamente las columnas de SKU, Nombre y Precio.
2. **Save As**: Elige la carpeta y el nombre del archivo `.pptx` final.
3. **Generar**: Haz clic en el botón azul. El programa procesará los productos en pares, buscando imágenes embebidas en el Excel primero y, si no existen, las descargará automáticamente del servidor S3.
4. **Progreso**: Sigue la barra de progreso y los logs en la parte inferior para ver el avance en tiempo real.

## 💡 Notas Técnicas
- **Soporte de Imágenes**: El sistema es híbrido. Soporta imágenes pegadas en las celdas de Excel y enlaces automáticos vía SKU.
- **Diseño**: El layout es de cuadrantes alternados (Checkerboard) para un flujo visual equilibrado.
