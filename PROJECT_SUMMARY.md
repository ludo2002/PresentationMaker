# Project: Adidas Presentation Automator 👟

## 🎯 Objective
Automate the creation of professional PowerPoint presentations (`.pptx`) using data from an Excel file (`.xlsx`) and product images hosted online.

## 🛠️ Current Technical Stack
- **Data Source**: `.xlsx` files (e.g., `ADIDASQ227CLZV_copilot.xlsx`, `RBQ127VVCLZ.xlsx`)
- **Image Mapping**: 
    - Automatic generation via S3 CDN (`https://s3.amazonaws.com/cdn.grabasaweb.mx/Imagenes/Imagenes_Preventa/{SKU}.jpg`)
    - Manual mapping via `.txt` file.
- **Tool**: **CatalogMaster v2.2** (GUI Application)
- **Engine**: Python 3.12
- **Libraries**: `openpyxl`, `python-pptx`, `requests`, `Pillow`, `tkinter`
- **Main App**: `adidas_presentation_project/CatalogMaster/presenter_gui.py`

## 📌 Final Status (v2.2)
- [x] **Universal Column Mapping**: Automatically detects headers regardless of name (Modelo, SKU, Nombre, Precio, etc.) and regardless of the header row position (scans first 10 rows).
- [x] **Auto-URL Generation**: Integrated CDN link generator based on product SKU.
- [x] **High-Fidelity Design**: 
    - No borders (clean look).
    - Aspect-ratio correction for images (no distortion).
    - Dual-column symmetrical layout.
- [x] **Portable GUI**: User-friendly interface for assistants to run the tool without using the terminal.

## 🚀 Execution Command
```bash
./admin/venv/bin/python3 adidas_presentation_project/CatalogMaster/presenter_gui.py
```

## 🕒 Next Steps
- The tool is now stable and universal. 
- Future expansion: Ability to load custom PPTX templates via GUI.
