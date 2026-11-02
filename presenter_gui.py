import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk, simpledialog
import openpyxl
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
import requests
from io import BytesIO
import zipfile
import xml.etree.ElementTree as ET
import posixpath
import os
import copy
from PIL import Image
import threading

class PresenterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("CatalogMaster v3.2 🚀 - Campos Personalizados & Plantillas")
        self.root.geometry("780x980") 
        self.root.configure(bg="#f0f0f0")
        
        self.excel_path = tk.StringVar()
        self.sheet_var = tk.StringVar()
        self.output_path = tk.StringVar()
        self.template_path = tk.StringVar()
        self.template_select_var = tk.StringVar(value="🎨 Predeterminada (Lienzo Blanco)")
        self.local_img_folder = tk.StringVar()
        self.progress_var = tk.DoubleVar(value=0)
        
        # Mapeo de plantillas predefinidas
        self.base_dir = os.path.dirname(os.path.abspath(__file__))
        self.templates_dir = os.path.join(self.base_dir, "templates")
        
        self.builtin_templates = {
            "🎨 Predeterminada (Lienzo Blanco)": "",
            "🌙 Dark Mode Elegante": os.path.join(self.templates_dir, "plantilla_dark_mode.pptx"),
            "👔 Corporativa Azul Navy": os.path.join(self.templates_dir, "plantilla_corporativa_navy.pptx"),
            "🔥 Deportiva Urban Red": os.path.join(self.templates_dir, "plantilla_deportiva_urban.pptx")
        }

        self.map_vars = {
            'sku': tk.StringVar(),
            'name': tk.StringVar(),
            'color': tk.StringVar(),
            'price': tk.StringVar(),
            'sizes': tk.StringVar(),
            'temp': tk.StringVar(),
            'gender': tk.StringVar(),
            'silhouette': tk.StringVar()
        }
        self.excel_headers = []
        self.header_row_idx = 1
        
        self.display_names = {
            'sku': 'SKU / Modelo',
            'name': 'Nombre',
            'temp': 'Temporada',
            'color': 'Color',
            'sizes': 'Tallas',
            'gender': 'Género',
            'silhouette': 'Silueta',
            'price': 'Precio'
        }
        self.display_to_key = {v: k for k, v in self.display_names.items()}
        self.custom_keys = set()
        self.custom_field_counter = 0
        self.comboboxes = {}
        self.current_order = []
        
        self.setup_ui()
        self.update_preview()

    def setup_ui(self):
        main_canvas = tk.Canvas(self.root, bg="#f0f0f0", highlightthickness=0)
        scrollbar = ttk.Scrollbar(self.root, orient="vertical", command=main_canvas.yview)
        
        scrollable_frame = tk.Frame(main_canvas, bg="#f0f0f0", padx=15, pady=15)
        scrollable_frame.bind(
            "<Configure>",
            lambda e: main_canvas.configure(scrollregion=main_canvas.bbox("all"))
        )
        
        main_canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        main_canvas.configure(yscrollcommand=scrollbar.set)
        
        main_canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        header = tk.Label(scrollable_frame, text="CatalogMaster v3.2", font=("Arial", 20, "bold"), bg="#f0f0f0", fg="#333")
        header.pack(pady=(0, 2))
        sub_header = tk.Label(scrollable_frame, text="Campos Dinámicos Personalizados • Plantillas • Divisor Claro entre Productos", font=("Arial", 9, "italic"), bg="#f0f0f0", fg="#555")
        sub_header.pack(pady=(0, 10))
        
        # --- SECCIÓN 1: Selección de Plantilla y Rutas ---
        files_frame = tk.LabelFrame(scrollable_frame, text="Plantilla & Archivos de Entrada", bg="#f0f0f0", font=("Arial", 10, "bold"), padx=10, pady=10)
        files_frame.pack(fill="x", pady=(0, 10))

        # Menú Desplegable de Plantillas
        tpl_row = tk.Frame(files_frame, bg="#f0f0f0", pady=4)
        tpl_row.pack(fill="x")
        lbl_tpl = tk.Label(tpl_row, text="Plantilla Visual:", font=("Arial", 9, "bold"), bg="#f0f0f0", width=22, anchor="w", fg="#007bff")
        lbl_tpl.pack(side="left")
        
        tpl_options = list(self.builtin_templates.keys()) + ["📁 Buscar otra plantilla (.pptx)..."]
        self.cb_template = ttk.Combobox(tpl_row, textvariable=self.template_select_var, values=tpl_options, state="readonly", font=("Arial", 9))
        self.cb_template.pack(side="left", expand=True, fill="x", padx=5)
        self.cb_template.bind("<<ComboboxSelected>>", self.on_template_selected)

        self.create_file_row(files_frame, "Excel File (.xlsx):", self.excel_path, self.browse_excel)

        sheet_row = tk.Frame(files_frame, bg="#f0f0f0", pady=3)
        sheet_row.pack(fill="x")
        tk.Label(sheet_row, text="Hoja del Excel:", font=("Arial", 9, "bold"), bg="#f0f0f0", width=22, anchor="w", fg="#007bff").pack(side="left")
        self.cb_sheet = ttk.Combobox(sheet_row, textvariable=self.sheet_var, values=[], state="disabled", font=("Arial", 9))
        self.cb_sheet.pack(side="left", expand=True, fill="x", padx=5)
        self.cb_sheet.bind("<<ComboboxSelected>>", self.on_sheet_selected)

        self.create_file_row(files_frame, "Save As (.pptx):", self.output_path, self.browse_output)
        self.create_file_row(files_frame, "Carpeta Imágenes (Opcional):", self.local_img_folder, self.browse_local_img_folder)

        # --- SECCIÓN 2: Mapeo de Columnas ---
        self.mapping_frame = tk.LabelFrame(scrollable_frame, text="Asignación de Columnas de Excel", bg="#f0f0f0", font=("Arial", 10, "bold"), padx=10, pady=10)
        self.mapping_frame.pack(fill="x", pady=(0, 10))

        # Barra superior con botón de agregar campo personalizado
        top_map_bar = tk.Frame(self.mapping_frame, bg="#f0f0f0")
        top_map_bar.pack(fill="x", pady=(0, 5))
        
        btn_add_field = tk.Button(top_map_bar, text="➕ Agregar Campo Personalizado", command=self.add_custom_field_dialog, bg="#28a745", fg="white", font=("Arial", 9, "bold"), relief="flat", cursor="hand2", padx=8, pady=2)
        btn_add_field.pack(side="right")
        
        lbl_map_hint = tk.Label(top_map_bar, text="Selecciona la columna de Excel para cada campo:", font=("Arial", 9, "italic"), bg="#f0f0f0", fg="#666")
        lbl_map_hint.pack(side="left")

        self.fields_grid_frame = tk.Frame(self.mapping_frame, bg="#f0f0f0")
        self.fields_grid_frame.pack(fill="x")

        self.render_mapping_fields()

        # --- SECCIÓN 3: Orden de Visualización y Vista Previa ---
        middle_frame = tk.Frame(scrollable_frame, bg="#f0f0f0")
        middle_frame.pack(fill="x", pady=(0, 10))

        # Izquierda: Orden de campos
        self.order_frame = tk.LabelFrame(middle_frame, text="Orden de Campos", bg="#f0f0f0", font=("Arial", 10, "bold"), padx=10, pady=10)
        self.order_frame.pack(side="left", fill="both", expand=True, padx=(0, 5))

        list_container = tk.Frame(self.order_frame, bg="#f0f0f0")
        list_container.pack(fill="both", expand=True)

        self.order_listbox = tk.Listbox(list_container, height=8, font=("Arial", 9), selectmode=tk.SINGLE, activestyle="none")
        self.order_listbox.pack(side="left", expand=True, fill="both", padx=(0, 5))

        for visual_name in self.display_names.values():
            self.order_listbox.insert(tk.END, visual_name)

        btn_container = tk.Frame(list_container, bg="#f0f0f0")
        btn_container.pack(side="right")

        btn_up = tk.Button(btn_container, text="🔼 Subir", command=self.move_up, width=8, cursor="hand2")
        btn_up.pack(pady=3)
        btn_down = tk.Button(btn_container, text="🔽 Bajar", command=self.move_down, width=8, cursor="hand2")
        btn_down.pack(pady=3)

        # Derecha: Panel de Vista Previa (Live Preview con Divisor)
        preview_frame = tk.LabelFrame(middle_frame, text="👁️ Vista Previa en Tiempo Real", bg="#f0f0f0", font=("Arial", 10, "bold"), padx=10, pady=10)
        preview_frame.pack(side="right", fill="both", expand=True, padx=(5, 0))

        self.preview_canvas = tk.Canvas(preview_frame, width=320, height=210, bg="#ffffff", highlightthickness=1, highlightbackground="#bbb")
        self.preview_canvas.pack(anchor="center")

        btn_refresh = tk.Button(preview_frame, text="🔄 Actualizar Vista Previa", command=self.update_preview, font=("Arial", 8), cursor="hand2")
        btn_refresh.pack(pady=(5, 0))

        # --- Botón de Generar y Logs ---
        self.btn_generate = tk.Button(scrollable_frame, text="🚀 GENERATE PRESENTATION", command=self.start_generation, bg="#007bff", fg="white", font=("Arial", 12, "bold"), padx=20, pady=10, cursor="hand2", relief="flat")
        self.btn_generate.pack(fill="x", pady=5)
        
        self.progress_bar = ttk.Progressbar(scrollable_frame, variable=self.progress_var, maximum=100)
        self.progress_bar.pack(fill="x", pady=(0, 10))
        
        self.log_area = scrolledtext.ScrolledText(scrollable_frame, height=6, font=("Consolas", 9), state='disabled', bg="#fff", fg="#555")
        self.log_area.pack(fill="x")

    def render_mapping_fields(self):
        """Redibuja dinámicamente el grid de asignación de columnas (incluyendo campos personalizados)."""
        for child in self.fields_grid_frame.winfo_children():
            child.destroy()
            
        self.comboboxes.clear()
        opciones_desplegable = ["--- Ninguno ---"] + self.excel_headers if self.excel_headers else ["--- Ninguno ---"]

        row_idx = 0
        col_idx = 0

        for key, var in self.map_vars.items():
            visual_name = self.display_names.get(key, key)
            is_custom = key in self.custom_keys

            frame = tk.Frame(self.fields_grid_frame, bg="#f0f0f0")
            frame.grid(row=row_idx, column=col_idx, padx=4, pady=3, sticky="ew")

            label_text = f"✨ {visual_name}:" if is_custom else f"{visual_name}:"
            lbl = tk.Label(frame, text=label_text, bg="#f0f0f0", width=14, anchor="w", font=("Arial", 9, "bold" if is_custom else "normal"))
            lbl.pack(side="left")

            cb = ttk.Combobox(frame, textvariable=var, state="readonly", width=15, values=opciones_desplegable)
            cb.pack(side="left", padx=2)
            cb.bind("<<ComboboxSelected>>", lambda e: self.update_preview())
            self.comboboxes[key] = cb

            if is_custom:
                btn_del = tk.Button(frame, text="❌", command=lambda k=key: self.remove_custom_field(k), bg="#dc3545", fg="white", font=("Arial", 7, "bold"), width=2, relief="flat", cursor="hand2")
                btn_del.pack(side="left", padx=2)

            col_idx += 1
            if col_idx > 1:
                col_idx = 0
                row_idx += 1

    def add_custom_field_dialog(self):
        field_name = simpledialog.askstring("Nuevo Campo Personalizado", "Ingresa el nombre del nuevo campo (ej. Marca, Categoría, Descuento):", parent=self.root)
        if field_name and field_name.strip():
            self.add_custom_field(field_name.strip())

    def add_custom_field(self, name):
        name_clean = name.strip()
        if name_clean in self.display_to_key:
            messagebox.showwarning("Campo Existente", f"El campo '{name_clean}' ya existe.")
            return

        self.custom_field_counter += 1
        key = f"custom_{self.custom_field_counter}"
        self.custom_keys.add(key)

        self.map_vars[key] = tk.StringVar()
        self.display_names[key] = name_clean
        self.display_to_key[name_clean] = key

        self.render_mapping_fields()

        # Agregar a la lista de ordenamiento
        self.order_listbox.insert(tk.END, name_clean)

        # Autoadivinar coincidencia si hay encabezados de Excel cargados
        if self.excel_headers:
            for col_name in self.excel_headers:
                if col_name and name_clean.lower() in col_name.lower():
                    self.map_vars[key].set(col_name)
                    break
            else:
                self.map_vars[key].set("--- Ninguno ---")

        self.log(f"➕ Campo personalizado agregado: '{name_clean}'")
        self.update_preview()

    def remove_custom_field(self, key):
        display_name = self.display_names.get(key)
        if not display_name:
            return

        if messagebox.askyesno("Eliminar Campo", f"¿Deseas eliminar el campo personalizado '{display_name}'?"):
            del self.map_vars[key]
            del self.display_names[key]
            del self.display_to_key[display_name]
            self.custom_keys.remove(key)

            items = self.order_listbox.get(0, tk.END)
            if display_name in items:
                idx = items.index(display_name)
                self.order_listbox.delete(idx)

            self.render_mapping_fields()
            self.log(f"🗑️ Campo personalizado eliminado: '{display_name}'")
            self.update_preview()

    def on_template_selected(self, event=None):
        selected = self.template_select_var.get()
        if selected == "📁 Buscar otra plantilla (.pptx)...":
            self.browse_template()
        elif selected in self.builtin_templates:
            tpl_path = self.builtin_templates[selected]
            self.template_path.set(tpl_path)
            if tpl_path:
                self.log(f"🎨 Plantilla seleccionada del menú: {selected}")
            else:
                self.log("🎨 Seleccionada plantilla predeterminada (Lienzo blanco).")
        self.update_preview()

    def move_up(self):
        try:
            selected_idx = self.order_listbox.curselection()[0]
            if selected_idx > 0:
                text = self.order_listbox.get(selected_idx)
                self.order_listbox.delete(selected_idx)
                self.order_listbox.insert(selected_idx - 1, text)
                self.order_listbox.selection_set(selected_idx - 1)
                self.update_preview()
        except IndexError:
            pass 

    def move_down(self):
        try:
            selected_idx = self.order_listbox.curselection()[0]
            if selected_idx < self.order_listbox.size() - 1:
                text = self.order_listbox.get(selected_idx)
                self.order_listbox.delete(selected_idx)
                self.order_listbox.insert(selected_idx + 1, text)
                self.order_listbox.selection_set(selected_idx + 1)
                self.update_preview()
        except IndexError:
            pass

    def create_file_row(self, parent, label_text, var, browse_func):
        row = tk.Frame(parent, bg="#f0f0f0", pady=3)
        row.pack(fill="x")
        lbl = tk.Label(row, text=label_text, font=("Arial", 9), bg="#f0f0f0", width=22, anchor="w")
        lbl.pack(side="left")
        ent = tk.Entry(row, textvariable=var, font=("Arial", 9))
        ent.pack(side="left", expand=True, fill="x", padx=5)
        btn = tk.Button(row, text="Browse...", command=browse_func, font=("Arial", 8), cursor="hand2")
        btn.pack(side="right")

    def browse_excel(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if path:
            self.excel_path.set(path)
            self.sheet_var.set("")
            self.load_excel_headers()

    def on_sheet_selected(self, event=None):
        if self.excel_path.get() and self.sheet_var.get():
            self.log(f"📄 Hoja seleccionada: {self.sheet_var.get()}")
            self.load_excel_headers()

    def browse_output(self):
        path = filedialog.asksaveasfilename(defaultextension=".pptx", filetypes=[("PowerPoint files", "*.pptx")])
        if path: self.output_path.set(path)

    def browse_template(self):
        path = filedialog.askopenfilename(filetypes=[("PowerPoint files", "*.pptx")])
        if path: 
            self.template_path.set(path)
            self.template_select_var.set(f"📁 Custom: {os.path.basename(path)}")
            self.log(f"🎨 Plantilla cargada externamente: {os.path.basename(path)}")
            self.update_preview()

    def browse_local_img_folder(self):
        path = filedialog.askdirectory()
        if path:
            self.local_img_folder.set(path)
            self.log(f"📁 Carpeta local de imágenes seleccionada: {path}")
        
    @staticmethod
    def select_worksheet(workbook, sheet_name=None):
        """Devuelve la hoja solicitada o la hoja activa si no se indicó una."""
        if sheet_name:
            if sheet_name not in workbook.sheetnames:
                raise ValueError(f"La hoja '{sheet_name}' no existe en el archivo Excel.")
            return workbook[sheet_name]
        return workbook.active

    def load_excel_headers(self):
        try:
            self.log("Cargando encabezados del Excel...")
            wb = openpyxl.load_workbook(self.excel_path.get(), data_only=True, read_only=True)
            sheet_names = wb.sheetnames
            self.cb_sheet.configure(values=sheet_names, state="readonly")
            selected_name = self.sheet_var.get()
            if selected_name not in sheet_names:
                active_sheet = wb.active
                if active_sheet is None:
                    raise ValueError("El archivo Excel no contiene una hoja activa.")
                selected_name = active_sheet.title
                self.sheet_var.set(selected_name)
            sheet = self.select_worksheet(wb, selected_name)
            
            self.excel_headers = []
            self.header_row_idx = 1
            
            for i, row in enumerate(sheet.iter_rows(min_row=1, max_row=10, values_only=True), 1):
                if any(val and any(k in str(val).lower() for k in ['modelo', 'nombre', 'sku', 'precio', 'talla']) for val in row):
                    self.excel_headers = [str(cell).strip() if cell is not None else "" for cell in row]
                    self.header_row_idx = i
                    break
                    
            wb.close()
            
            if not self.excel_headers:
                self.log("⚠️ No se encontraron encabezados claros. Revisa tu Excel.")
                return

            self.render_mapping_fields()
            
            def guess(keywords):
                for col_name in self.excel_headers:
                    if col_name and any(k.lower() in col_name.lower() for k in keywords): 
                        return col_name
                return "--- Ninguno ---" 

            self.map_vars['sku'].set(guess(['modelo', 'sku', 'id', 'codigo', 'artículo']))
            self.map_vars['name'].set(guess(['nombre', 'name', 'titulo', 'product', 'descripcion']))
            self.map_vars['color'].set(guess(['color']))
            self.map_vars['price'].set(guess(['p/p', 'precio', 'price', 'cost']))
            self.map_vars['sizes'].set(guess(['tallas', 'size', 'talla']))
            self.map_vars['temp'].set(guess(['temp', 'temporada', 'season', 'año']))
            self.map_vars['gender'].set(guess(['genero', 'gender', 'sexo']))
            self.map_vars['silhouette'].set(guess(['silueta', 'silhouette', 'estilo', 'corte']))
            
            for key in self.custom_keys:
                display_name = self.display_names.get(key, '')
                self.map_vars[key].set(guess([display_name.lower()]))

            self.log(f"✅ Encabezados cargados de la hoja '{selected_name}'.")
            self.update_preview()

        except Exception as e:
            self.log(f"❌ Error al cargar encabezados: {e}")

    def update_preview(self):
        """Renderiza una vista previa gráfica interactiva de la diapositiva en el Canvas con el Divisor Claro."""
        cv = self.preview_canvas
        cv.delete("all")
        
        W, H = 320, 210
        tpl_name = self.template_select_var.get()

        bg_color = "#ffffff"
        card_bg = "#ffffff"
        text_main = "#000000"
        text_accent = "#d9534f"
        divider_color = "#007bff"
        header_bg = None
        side_stripe = None

        if "Dark Mode" in tpl_name:
            bg_color = "#12141A"
            card_bg = "#1C202A"
            text_main = "#ffffff"
            text_accent = "#ffd700"
            divider_color = "#ffd700"
            header_bg = "#ffd700"
        elif "Corporativa Azul Navy" in tpl_name:
            bg_color = "#f5f7fa"
            card_bg = "#ffffff"
            text_main = "#102540"
            text_accent = "#007aff"
            divider_color = "#007aff"
            header_bg = "#0a2540"
        elif "Deportiva Urban Red" in tpl_name:
            bg_color = "#fafafa"
            card_bg = "#ffffff"
            text_main = "#111111"
            text_accent = "#ff4500"
            divider_color = "#ff4500"
            side_stripe = "#ff4500"

        cv.create_rectangle(2, 2, W-2, H-2, fill=bg_color, outline="#333333", width=1)

        # Adornos visuales de la plantilla en Canvas
        if header_bg:
            cv.create_rectangle(2, 2, W-2, 14, fill=header_bg, outline="")
        if side_stripe:
            cv.create_rectangle(2, 2, 12, H-2, fill=side_stripe, outline="")

        # --- DIVISOR CLARO ENTRE PRODUCTO 1 Y PRODUCTO 2 ---
        cv.create_line(15, H/2, W-15, H/2, fill=divider_color, width=3)
        cv.create_text(W/2, H/2, text=" DIVISOR DE PRODUCTO ", fill=divider_color, font=("Arial", 6, "bold"))

        sample_order = list(self.order_listbox.get(0, tk.END))

        # --- PRODUCTO 1 (Arriba) ---
        img_x1 = 20 if not side_stripe else 24
        cv.create_rectangle(img_x1, 20, W/2 - 10, H/2 - 10, fill=card_bg, outline=divider_color, width=1, dash=(2, 2))
        cv.create_text((img_x1 + W/2 - 10)/2, (20 + H/2 - 10)/2, text="📷 Producto 1", fill=divider_color, font=("Arial", 7, "bold"))

        y_text = 20
        for visual_name in sample_order[:5]:
            key = self.display_to_key.get(visual_name, '')
            val_col = self.map_vars.get(key, type('obj', (), {'get': lambda: ''})()).get()
            if val_col == "--- Ninguno ---" or not val_col:
                continue
            
            prefix = visual_name + ":"
            txt = f"{prefix} Muestra"
            if key == 'price': txt = "Precio: $1,299.00"
            elif key == 'sku': txt = "SKU: 100266540"
            elif key == 'name': txt = "REEBOK VIVA SPEED"
            
            color = text_main if key != 'price' else text_accent
            font_w = "bold" if key in ['name', 'price', 'sku'] else "normal"
            cv.create_text(W/2 + 10, y_text, text=txt[:24], anchor="nw", fill=color, font=("Arial", 7, font_w))
            y_text += 12

        # --- PRODUCTO 2 (Abajo) ---
        y_text2 = H/2 + 15
        for visual_name in sample_order[:5]:
            key = self.display_to_key.get(visual_name, '')
            val_col = self.map_vars.get(key, type('obj', (), {'get': lambda: ''})()).get()
            if val_col == "--- Ninguno ---" or not val_col:
                continue
            
            prefix = visual_name + ":"
            txt = f"{prefix} Muestra"
            if key == 'price': txt = "Precio: $1,499.00"
            elif key == 'sku': txt = "SKU: 100272105"
            elif key == 'name': txt = "NFX TRAINER 2"
            
            color = text_main if key != 'price' else text_accent
            font_w = "bold" if key in ['name', 'price', 'sku'] else "normal"
            cv.create_text(img_x1, y_text2, text=txt[:24], anchor="nw", fill=color, font=("Arial", 7, font_w))
            y_text2 += 12

        cv.create_rectangle(W/2 + 10, H/2 + 15, W - 15, H - 15, fill=card_bg, outline=divider_color, width=1, dash=(2, 2))
        cv.create_text((W/2 + 10 + W - 15)/2, (H/2 + 15 + H - 15)/2, text="📷 Producto 2", fill=divider_color, font=("Arial", 7, "bold"))

    def log(self, message):
        self.log_area.configure(state='normal')
        self.log_area.insert(tk.END, message + "\n")
        self.log_area.see(tk.END)
        self.log_area.configure(state='disabled')
        self.root.update_idletasks()

    def start_generation(self):
        if not self.excel_path.get() or not self.output_path.get():
            messagebox.showerror("Error", "Por favor selecciona al menos el archivo Excel y la ubicación de salida.")
            return
        
        if not self.map_vars['sku'].get() or self.map_vars['sku'].get() == "--- Ninguno ---":
            messagebox.showerror("Error", "Por favor asigna al menos la columna para 'SKU / Modelo'.")
            return
            
        self.current_order = list(self.order_listbox.get(0, tk.END))
            
        self.progress_var.set(0)
        self.btn_generate.config(state='disabled', text="Processing...")
        threading.Thread(target=self.generate_logic, daemon=True).start()

    def extract_embedded_images(self, excel_path, sheet_name=None):
        embedded_images = {}
        try:
            self.image_links = {}
            with zipfile.ZipFile(excel_path, 'r') as z:
                # Follow workbook -> active sheet -> drawing relationships. Sheet
                # numbers are not guaranteed to match their position in the file.
                main_ns = {'m': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
                rel_attr = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
                pkg_ns = {'p': 'http://schemas.openxmlformats.org/package/2006/relationships'}

                def relations(part):
                    path = posixpath.join(posixpath.dirname(part), '_rels', posixpath.basename(part) + '.rels')
                    if path not in z.namelist():
                        return {}
                    return {r.get('Id'): (r.get('Target'), r.get('TargetMode', 'Internal'))
                            for r in ET.fromstring(z.read(path)).findall('p:Relationship', pkg_ns)}

                def resolve(part, target):
                    return target.lstrip('/') if target.startswith('/') else posixpath.normpath(posixpath.join(posixpath.dirname(part), target))

                workbook = ET.fromstring(z.read('xl/workbook.xml'))
                view = workbook.find('m:bookViews/m:workbookView', main_ns)
                active_index = int(view.get('activeTab', '0')) if view is not None else 0
                sheets = workbook.findall('m:sheets/m:sheet', main_ns)
                selected_sheet = None
                if sheet_name:
                    selected_sheet = next((item for item in sheets if item.get('name') == sheet_name), None)
                    if selected_sheet is None:
                        raise ValueError(f"La hoja '{sheet_name}' no existe en el archivo Excel.")
                else:
                    selected_sheet = sheets[active_index]
                sheet_rid = selected_sheet.get(rel_attr + 'id')
                sheet_part = resolve('xl/workbook.xml', relations('xl/workbook.xml')[sheet_rid][0])
                sheet = ET.fromstring(z.read(sheet_part))
                sheet_rels = relations(sheet_part)
                drawing_files = []
                for drawing in sheet.findall('m:drawing', main_ns):
                    rid = drawing.get(rel_attr + 'id')
                    if rid in sheet_rels:
                        drawing_files.append(resolve(sheet_part, sheet_rels[rid][0]))
                for drawing_file in drawing_files:
                    xml_content = z.read(drawing_file)
                    root = ET.fromstring(xml_content)
                    ns = {
                        'xdr': 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
                        'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'
                    }
                    
                    rels_path = posixpath.join(posixpath.dirname(drawing_file), '_rels', posixpath.basename(drawing_file) + '.rels')
                    rels_map = {}
                    if rels_path in z.namelist():
                        rels_content = z.read(rels_path)
                        rels_root = ET.fromstring(rels_content)
                        rel_ns = {'rel': 'http://schemas.openxmlformats.org/package/2006/relationships'}
                        for r in rels_root.findall('.//rel:Relationship', rel_ns):
                            rid = r.get('Id')
                            target = r.get('Target')
                            mode = r.get('TargetMode', 'Internal')
                            rels_map[rid] = (target, mode)

                    anchors = root.findall('.//xdr:twoCellAnchor', ns) + root.findall('.//xdr:oneCellAnchor', ns)
                    for anchor in anchors:
                        from_node = anchor.find('xdr:from', ns)
                        if from_node is not None:
                            row = int(from_node.find('xdr:row', ns).text)
                            col = int(from_node.find('xdr:col', ns).text)
                            pic = anchor.find('.//xdr:pic', ns)
                            if pic is not None:
                                blip = pic.find('.//a:blip', ns)
                                if blip is not None:
                                    rid = blip.get(rel_attr + 'embed')
                                    if rid and rid in rels_map:
                                        target, mode = rels_map[rid]
                                        if mode == 'Internal':
                                            full_path = resolve(drawing_file, target)
                                            if full_path in z.namelist():
                                                embedded_images[(row, col)] = z.read(full_path)
                                    links = []
                                    click = pic.find('.//a:hlinkClick', ns)
                                    for link_id in (click.get(rel_attr + 'id') if click is not None else None,
                                                    blip.get(rel_attr + 'link')):
                                        if link_id in rels_map:
                                            target, mode = rels_map[link_id]
                                            if mode == 'External' and target.startswith(('https://', 'http://')) and target not in links:
                                                links.append(target)
                                    if links:
                                        self.image_links[(row, col)] = links
            self.log(f"🖼️ Se extrajeron {len(embedded_images)} imágenes embebidas del archivo Excel.")
        except Exception as e:
            self.log(f"⚠️ Error extrayendo imágenes embebidas: {e}")
        return embedded_images

    def resolve_product_image_bytes(self, product, embedded_map, row_idx):
        """Elige la imagen válida con más píxeles entre Excel, carpeta y CDN."""
        candidates = []

        def consider(data, source):
            try:
                with Image.open(BytesIO(data)) as image:
                    width, height = image.size
                    image.verify()
                if width > 0 and height > 0:
                    candidates.append((width * height, data, source, width, height))
            except Exception:
                pass

        # Los índices de dibujo son base cero; nunca usar la fila de otro producto.
        for (row, _col), data in embedded_map.items():
            if row + 1 == row_idx:
                consider(data, "Embebida en Excel")

        sku = product['sku']
        extensions = ['jpg', 'png', 'jpeg', 'webp']
        local_dir = self.local_img_folder.get()
        if local_dir and os.path.isdir(local_dir):
            for ext in extensions + [ext.upper() for ext in extensions]:
                local_file = os.path.join(local_dir, f"{sku}.{ext}")
                if os.path.isfile(local_file):
                    try:
                        with open(local_file, 'rb') as f:
                            consider(f.read(), f"Carpeta Local ({ext})")
                    except OSError:
                        pass

        urls = []
        for (row, _col), links in getattr(self, 'image_links', {}).items():
            if row + 1 == row_idx:
                urls.extend((url, "Enlace de Excel") for url in links)
        urls.extend((f"https://s3.amazonaws.com/cdn.grabasaweb.mx/Imagenes/Imagenes_Preventa/{sku}.{ext}",
                     f"CDN S3 ({ext})") for ext in extensions)
        visited = set()
        for url, source in urls:
            if url in visited:
                continue
            visited.add(url)
            try:
                resp = requests.get(url, timeout=4)
                if resp.status_code == 200:
                    consider(resp.content, source)
            except requests.RequestException:
                pass

        if not candidates:
            return None, "No encontrada"
        _, data, source, width, height = max(candidates, key=lambda candidate: candidate[0])
        self.log(f"🖼️ {sku}: {source}, {width} × {height} píxeles")
        return data, source

    def build_dynamic_data(self, product):
        lines_data = []
        
        for visual_name in self.current_order:
            key = self.display_to_key.get(visual_name)
            if not key:
                continue
            val = product.get(key)
            
            if not val or (val == "N/A" and key == 'sku'):
                continue
                
            if key == 'name':
                lines_data.append(('NOMBRE:', f" {val.upper()}", False))
            elif key == 'price':
                lines_data.append(('', '', True))
                lines_data.append(('Precio:', f" ${val}", False))
            elif key == 'sku':
                lines_data.append(('SKU / Modelo:', f" {val}", False))
            elif key == 'sizes':
                lines_data.append(('Talla:', f"\n{val}", False))
            else:
                label_prefixes = {
                    'temp': 'Temp:',
                    'color': 'Color:',
                    'gender': 'Genero:',
                    'silhouette': 'Silueta:'
                }
                prefix = label_prefixes.get(key, f"{visual_name}:")
                lines_data.append((prefix, f" {val}", False))
                
        return lines_data

    def generate_logic(self):
        try:
            self.log("🚀 Iniciando generación de presentación...")
            selected_sheet_name = self.sheet_var.get()
            self.log(f"📄 Procesando hoja: {selected_sheet_name}")
            embedded_map = self.extract_embedded_images(self.excel_path.get(), sheet_name=selected_sheet_name)
            
            wb = openpyxl.load_workbook(self.excel_path.get(), data_only=True)
            sheet = self.select_worksheet(wb, selected_sheet_name)
            
            m_indices = {}
            for key, var in self.map_vars.items():
                col_name = var.get()
                if col_name and col_name != "--- Ninguno ---" and col_name in self.excel_headers:
                    m_indices[key] = self.excel_headers.index(col_name)
                else:
                    m_indices[key] = None

            has_template_slide = False
            template_bg_elements = []
            template_file = self.template_path.get()
            if template_file and os.path.isfile(template_file):
                self.log(f"🎨 Cargando plantilla personalizada: {os.path.basename(template_file)}")
                prs = Presentation(template_file)
                if len(prs.slides) > 0:
                    has_template_slide = True
                    for shape in prs.slides[0].shapes:
                        template_bg_elements.append(copy.deepcopy(shape.element))
            else:
                prs = Presentation()
                prs.slide_width, prs.slide_height = Inches(10), Inches(7.5)

            products = []
            for row_idx, row in enumerate(sheet.iter_rows(min_row=self.header_row_idx + 1, values_only=True), start=self.header_row_idx + 1):
                sku_idx = m_indices['sku']
                sku_val = row[sku_idx] if sku_idx is not None and sku_idx < len(row) else None
                if not sku_val or str(sku_val).strip() == "": continue
                
                clean_sku = str(sku_val).strip()
                if clean_sku.endswith('.0'):
                    clean_sku = clean_sku[:-2]
                
                def get_val(k):
                    idx = m_indices[k]
                    if idx is not None:
                        if idx < len(row) and row[idx] is not None:
                            return str(row[idx]).strip()
                        else:
                            return "N/A"
                    return ""

                prod_dict = {
                    'sku': clean_sku,
                    'row_idx': row_idx
                }
                for k in self.map_vars.keys():
                    if k != 'sku':
                        prod_dict[k] = get_val(k)

                products.append(prod_dict)

            wb.close()
            self.log(f"📦 Total de productos encontrados en la hoja '{selected_sheet_name}': {len(products)}")
            
            total_slides = (len(products) + 1) // 2
            layout_idx = 6 if len(prs.slide_layouts) > 6 else 0

            # Determinar color y grosor de línea divisoria según plantilla
            tpl_name = self.template_select_var.get()
            div_color = RGBColor(230, 230, 230)
            line_thickness = Inches(0.01)
            
            if "Dark Mode" in tpl_name:
                div_color = RGBColor(255, 215, 0)
                line_thickness = Inches(0.03)
            elif "Urban Red" in tpl_name:
                div_color = RGBColor(255, 69, 0)
                line_thickness = Inches(0.03)
            elif "Corporativa" in tpl_name:
                div_color = RGBColor(0, 122, 255)
                line_thickness = Inches(0.03)

            for i in range(0, len(products), 2):
                slide_num = i // 2
                progress = (slide_num / total_slides) * 100
                self.progress_var.set(progress)
                self.log(f"Generando Diapositiva {slide_num + 1}/{total_slides}...")
                
                if has_template_slide:
                    if slide_num == 0:
                        slide = prs.slides[0]
                    else:
                        slide = prs.slides.add_slide(prs.slides[0].slide_layout)
                        for bg_el in template_bg_elements:
                            slide.shapes._spTree.insert_element_before(copy.deepcopy(bg_el), 'p:extLst')
                else:
                    slide = prs.slides.add_slide(prs.slide_layouts[layout_idx])
                
                sw, sh = prs.slide_width, prs.slide_height
                qw, qh = sw / 2, sh / 2
                pad = Inches(0.4)

                # --- AGREGAR DIVISOR VISUAL ENTRE PRODUCTO 1 Y PRODUCTO 2 ---
                if i + 1 < len(products):
                    divider = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.3), qh, sw - Inches(0.6), line_thickness)
                    divider.fill.solid()
                    divider.fill.fore_color.rgb = div_color
                    divider.line.color.rgb = div_color

                # --- PRODUCTO 1 ---
                p1 = products[i]
                img_bytes_1, src_name1 = self.resolve_product_image_bytes(p1, embedded_map, p1['row_idx'])
                if img_bytes_1:
                    self.add_image_bytes_to_quadrant(slide, img_bytes_1, 0, 0, qw, qh)
                else:
                    self.log(f"⚠️ Sin imagen para SKU {p1['sku']}")
                
                p1_data = self.build_dynamic_data(p1)
                self.create_box(slide, qw + pad, pad, qw - 2*pad, qh - 2*pad, p1_data, 15)

                # --- PRODUCTO 2 ---
                if i + 1 < len(products):
                    p2 = products[i+1]
                    p2_data = self.build_dynamic_data(p2)
                    self.create_box(slide, pad, qh + pad, qw - 2*pad, qh - 2*pad, p2_data, 15)
                    
                    img_bytes_2, src_name2 = self.resolve_product_image_bytes(p2, embedded_map, p2['row_idx'])
                    if img_bytes_2:
                        self.add_image_bytes_to_quadrant(slide, img_bytes_2, qw, qh, qw, qh)
                    else:
                        self.log(f"⚠️ Sin imagen para SKU {p2['sku']}")

            self.progress_var.set(100)
            prs.save(self.output_path.get())
            self.log("\n✅ ¡ÉXITO! Presentación generada y guardada correctamente.")
            messagebox.showinfo("Finalizado", "¡Presentación generada con éxito!")
        except Exception as e:
            self.log(f"❌ ERROR: {str(e)}")
            messagebox.showerror("Error", f"Ocurrió un error inesperado: {e}")
        finally:
            self.btn_generate.config(state='normal', text="🚀 GENERATE PRESENTATION")

    def add_image_bytes_to_quadrant(self, slide, img_bytes, left, top, width, height):
        try:
            with Image.open(BytesIO(img_bytes)) as img:
                img_w, img_h = img.size
                if img.format == 'WEBP':
                    converted = BytesIO()
                    img.save(converted, format='PNG')
                    img_bytes = converted.getvalue()
                aspect = img_w / img_h
                img_pad = Inches(0.1)
                max_w = width - 2*img_pad
                max_h = height - 2*img_pad
                if aspect > (max_w / max_h):
                    final_w = max_w; final_h = max_w / aspect
                else:
                    final_h = max_h; final_w = max_h * aspect
                off_x = left + (width - final_w) / 2
                off_y = top + (height - final_h) / 2
                slide.shapes.add_picture(BytesIO(img_bytes), off_x, off_y, final_w, final_h)
        except Exception as e:
            self.log(f"⚠️ Error al insertar imagen en diapositiva: {e}")

    def create_box(self, slide, left, top, width, height, lines_data, font_size=15):
        txBox = slide.shapes.add_textbox(left, top, width, height)
        tf = txBox.text_frame
        tf.word_wrap = True

        first = True
        for prefix, value, is_spacer in lines_data:
            if first:
                p = tf.paragraphs[0]
                first = False
            else:
                p = tf.add_paragraph()

            p.alignment = PP_ALIGN.LEFT

            if is_spacer:
                run = p.add_run()
                run.text = " "
                run.font.size = Pt(font_size * 0.8)
                continue

            # Prefijo en Negrita
            run_prefix = p.add_run()
            run_prefix.text = prefix
            run_prefix.font.bold = True
            run_prefix.font.name = 'Arial'
            run_prefix.font.color.rgb = RGBColor(0, 0, 0)

            # Valor en Regular
            run_val = p.add_run()
            run_val.text = value
            run_val.font.name = 'Arial'
            run_val.font.color.rgb = RGBColor(0, 0, 0)

            if "Precio" in prefix:
                run_val.font.bold = True
                run_prefix.font.size = Pt(font_size + 4)
                run_val.font.size = Pt(font_size + 4)
            else:
                run_val.font.bold = False
                run_prefix.font.size = Pt(font_size)
                run_val.font.size = Pt(font_size)

if __name__ == "__main__":
    root = tk.Tk()
    app = PresenterApp(root)
    root.mainloop()