import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk
import openpyxl
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
import requests
from io import BytesIO
import zipfile
import xml.etree.ElementTree as ET
import posixpath
import os
from PIL import Image
import threading

class PresenterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("CatalogMaster v3.0 🚀 - Plantillas, Multifuente & Vista Previa")
        self.root.geometry("750x950") 
        self.root.configure(bg="#f0f0f0")
        
        self.excel_path = tk.StringVar()
        self.output_path = tk.StringVar()
        self.template_path = tk.StringVar()
        self.local_img_folder = tk.StringVar()
        self.progress_var = tk.DoubleVar(value=0)
        
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

        header = tk.Label(scrollable_frame, text="CatalogMaster v3.0", font=("Arial", 20, "bold"), bg="#f0f0f0", fg="#333")
        header.pack(pady=(0, 2))
        sub_header = tk.Label(scrollable_frame, text="Plantillas Personalizadas • Búsqueda Multifuente • Vista Previa en Tiempo Real", font=("Arial", 9, "italic"), bg="#f0f0f0", fg="#555")
        sub_header.pack(pady=(0, 10))
        
        # --- SECCIÓN 1: Archivos y Rutas ---
        files_frame = tk.LabelFrame(scrollable_frame, text="Archivos y Rutas de Entrada", bg="#f0f0f0", font=("Arial", 10, "bold"), padx=10, pady=10)
        files_frame.pack(fill="x", pady=(0, 10))

        self.create_file_row(files_frame, "Excel File (.xlsx):", self.excel_path, self.browse_excel)
        self.create_file_row(files_frame, "Save As (.pptx):", self.output_path, self.browse_output)
        self.create_file_row(files_frame, "Plantilla PPTX (Opcional):", self.template_path, self.browse_template)
        self.create_file_row(files_frame, "Carpeta Imágenes (Opcional):", self.local_img_folder, self.browse_local_img_folder)

        # --- SECCIÓN 2: Mapeo de Columnas ---
        self.mapping_frame = tk.LabelFrame(scrollable_frame, text="Asignación de Columnas de Excel", bg="#f0f0f0", font=("Arial", 10, "bold"), padx=10, pady=10)
        self.mapping_frame.pack(fill="x", pady=(0, 10))
        
        labels = {
            'sku': 'SKU / Modelo:',
            'name': 'Nombre:',
            'color': 'Color:',
            'price': 'Precio:',
            'sizes': 'Tallas:',
            'temp': 'Temporada:',
            'gender': 'Género:',
            'silhouette': 'Silueta:'
        }
        
        row_idx = 0
        col_idx = 0
        self.comboboxes = {}
        for key, text_label in labels.items():
            frame = tk.Frame(self.mapping_frame, bg="#f0f0f0")
            frame.grid(row=row_idx, column=col_idx, padx=5, pady=4, sticky="ew")
            
            lbl = tk.Label(frame, text=text_label, bg="#f0f0f0", width=12, anchor="w", font=("Arial", 9))
            lbl.pack(side="left")
            
            cb = ttk.Combobox(frame, textvariable=self.map_vars[key], state="readonly", width=16)
            cb.pack(side="left", padx=5)
            cb.bind("<<ComboboxSelected>>", lambda e: self.update_preview())
            self.comboboxes[key] = cb
            
            col_idx += 1
            if col_idx > 1:
                col_idx = 0
                row_idx += 1

        # --- SECCIÓN 3: Orden de Visualización y Vista Previa ---
        middle_frame = tk.Frame(scrollable_frame, bg="#f0f0f0")
        middle_frame.pack(fill="x", pady=(0, 10))

        # Izquierda: Orden de campos
        self.order_frame = tk.LabelFrame(middle_frame, text="Orden de Visualización", bg="#f0f0f0", font=("Arial", 10, "bold"), padx=10, pady=10)
        self.order_frame.pack(side="left", fill="both", expand=True, padx=(0, 5))

        list_container = tk.Frame(self.order_frame, bg="#f0f0f0")
        list_container.pack(fill="both", expand=True)

        self.order_listbox = tk.Listbox(list_container, height=7, font=("Arial", 9), selectmode=tk.SINGLE, activestyle="none")
        self.order_listbox.pack(side="left", expand=True, fill="both", padx=(0, 5))

        for visual_name in self.display_names.values():
            self.order_listbox.insert(tk.END, visual_name)

        btn_container = tk.Frame(list_container, bg="#f0f0f0")
        btn_container.pack(side="right")

        btn_up = tk.Button(btn_container, text="🔼 Subir", command=self.move_up, width=8, cursor="hand2")
        btn_up.pack(pady=3)
        btn_down = tk.Button(btn_container, text="🔽 Bajar", command=self.move_down, width=8, cursor="hand2")
        btn_down.pack(pady=3)

        # Derecha: Panel de Vista Previa (Live Preview)
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
            self.load_excel_headers()

    def browse_output(self):
        path = filedialog.asksaveasfilename(defaultextension=".pptx", filetypes=[("PowerPoint files", "*.pptx")])
        if path: self.output_path.set(path)

    def browse_template(self):
        path = filedialog.askopenfilename(filetypes=[("PowerPoint files", "*.pptx")])
        if path: 
            self.template_path.set(path)
            self.log(f"🎨 Plantilla seleccionada: {os.path.basename(path)}")

    def browse_local_img_folder(self):
        path = filedialog.askdirectory()
        if path:
            self.local_img_folder.set(path)
            self.log(f"📁 Carpeta local de imágenes seleccionada: {path}")
        
    def load_excel_headers(self):
        try:
            self.log("Cargando encabezados del Excel...")
            wb = openpyxl.load_workbook(self.excel_path.get(), data_only=True, read_only=True)
            sheet = wb.active
            
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

            opciones_desplegable = ["--- Ninguno ---"] + self.excel_headers

            for key, cb in self.comboboxes.items():
                cb['values'] = opciones_desplegable
            
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
            
            self.log("✅ Encabezados cargados correctamente.")
            self.update_preview()

        except Exception as e:
            self.log(f"❌ Error al cargar encabezados: {e}")

    def update_preview(self):
        """Renderiza una vista previa gráfica de la diapositiva en el Canvas."""
        cv = self.preview_canvas
        cv.delete("all")
        
        W, H = 320, 210
        cv.create_rectangle(2, 2, W-2, H-2, fill="#ffffff", outline="#333333", width=1)
        cv.create_line(W/2, 2, W/2, H-2, fill="#e0e0e0", dash=(3, 3))
        cv.create_line(2, H/2, W-2, H/2, fill="#e0e0e0", dash=(3, 3))

        # Muestras simuladas
        sample_order = list(self.order_listbox.get(0, tk.END))
        
        # Cuadrante 1: Imagen Izquierda, Texto Derecha
        cv.create_rectangle(10, 10, W/2 - 10, H/2 - 10, fill="#f8f9fa", outline="#007bff", width=1, dash=(2, 2))
        cv.create_text(W/4, H/4, text="📷 Imagen P1", fill="#007bff", font=("Arial", 8, "bold"))
        
        # Texto P1
        y_text = 12
        for visual_name in sample_order[:5]:
            key = self.display_to_key.get(visual_name, '')
            val_col = self.map_vars.get(key, type('obj', (), {'get': lambda: ''})()).get()
            if val_col == "--- Ninguno ---" or not val_col:
                continue
            
            prefix = visual_name + ":"
            txt = f"{prefix} Muestra"
            if key == 'price': txt = "Precio: $1,299.00"
            elif key == 'sku': txt = "SKU: 100266540"
            elif key == 'name': txt = "NOMBRE EJEMPLO"
            
            color = "#000000" if key != 'price' else "#d9534f"
            font_w = "bold" if key in ['name', 'price', 'sku'] else "normal"
            cv.create_text(W/2 + 10, y_text, text=txt[:22], anchor="nw", fill=color, font=("Arial", 7, font_w))
            y_text += 14
            if y_text > H/2 - 15: break

        # Cuadrante 2: Texto Izquierda, Imagen Derecha
        y_text2 = H/2 + 12
        for visual_name in sample_order[:5]:
            key = self.display_to_key.get(visual_name, '')
            val_col = self.map_vars.get(key, type('obj', (), {'get': lambda: ''})()).get()
            if val_col == "--- Ninguno ---" or not val_col:
                continue
            
            prefix = visual_name + ":"
            txt = f"{prefix} Muestra"
            if key == 'price': txt = "Precio: $1,499.00"
            elif key == 'sku': txt = "SKU: 100272105"
            elif key == 'name': txt = "NOMBRE EJEMPLO 2"
            
            color = "#000000" if key != 'price' else "#d9534f"
            font_w = "bold" if key in ['name', 'price', 'sku'] else "normal"
            cv.create_text(10, y_text2, text=txt[:22], anchor="nw", fill=color, font=("Arial", 7, font_w))
            y_text2 += 14
            if y_text2 > H - 15: break

        cv.create_rectangle(W/2 + 10, H/2 + 10, W - 10, H - 10, fill="#f8f9fa", outline="#28a745", width=1, dash=(2, 2))
        cv.create_text(3*W/4, 3*H/4, text="📷 Imagen P2", fill="#28a745", font=("Arial", 8, "bold"))

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

    def extract_embedded_images(self, excel_path):
        embedded_images = {}
        try:
            with zipfile.ZipFile(excel_path, 'r') as z:
                drawing_files = [f for f in z.namelist() if f.startswith('xl/drawings/drawing') and f.endswith('.xml')]
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
                                    rid = None
                                    for attr_k, attr_v in blip.attrib.items():
                                        if attr_k.endswith('embed') or attr_k.endswith('id') or attr_k.endswith('link'):
                                            rid = attr_v
                                            break
                                    if rid and rid in rels_map:
                                        target, mode = rels_map[rid]
                                        if mode == 'Internal':
                                            full_path = posixpath.normpath(posixpath.join(posixpath.dirname(drawing_file), target))
                                            if full_path in z.namelist():
                                                embedded_images[(row, col)] = z.read(full_path)
            self.log(f"🖼️ Se extrajeron {len(embedded_images)} imágenes embebidas del archivo Excel.")
        except Exception as e:
            self.log(f"⚠️ Error extrayendo imágenes embebidas: {e}")
        return embedded_images

    def resolve_product_image_bytes(self, product, embedded_map, row_idx):
        """Búsqueda multifuente de imágenes: 1) Embebida en Excel, 2) Carpeta Local, 3) CDN S3."""
        # 1. Embebida en Excel
        for diff in [0, -1, 1]:
            for (r, c), data in embedded_map.items():
                if (r + 1) == (row_idx + diff):
                    return data, "Embebida en Excel"

        sku = product['sku']
        extensions = ['jpg', 'png', 'jpeg', 'webp', 'JPG', 'PNG', 'JPEG', 'WEBP']

        # 2. Carpeta local de imágenes
        local_dir = self.local_img_folder.get()
        if local_dir and os.path.isdir(local_dir):
            for ext in extensions:
                local_file = os.path.join(local_dir, f"{sku}.{ext}")
                if os.path.isfile(local_file):
                    try:
                        with open(local_file, 'rb') as f:
                            return f.read(), f"Carpeta Local ({ext})"
                    except Exception:
                        pass

        # 3. Servidor CDN S3 (Probando múltiples extensiones)
        for ext in ['jpg', 'png', 'jpeg', 'webp']:
            url = f"https://s3.amazonaws.com/cdn.grabasaweb.mx/Imagenes/Imagenes_Preventa/{sku}.{ext}"
            try:
                resp = requests.get(url, timeout=4)
                if resp.status_code == 200 and len(resp.content) > 100:
                    return resp.content, f"CDN S3 ({ext})"
            except Exception:
                pass

        return None, "No encontrada"

    def build_dynamic_data(self, product):
        """Construye los datos estructurados para texto con formato enriquecido."""
        lines_data = []
        
        for visual_name in self.current_order:
            key = self.display_to_key[visual_name] 
            val = product.get(key)
            
            if not val or (val == "N/A" and key == 'sku'):
                continue
                
            if key == 'name':
                lines_data.append(('NOMBRE:', f" {val.upper()}", False))
            elif key == 'price':
                lines_data.append(('', '', True)) # Agrega una línea en blanco antes del precio
                lines_data.append(('Precio:', f" ${val}", False))
            elif key == 'sku':
                lines_data.append(('SKU / Modelo:', f" {val}", False))
            elif key == 'sizes':
                lines_data.append(('Talla:', f"\n{val}", False)) # Obliga a las tallas a bajar de renglón
            else:
                label_prefixes = {
                    'temp': 'Temp:',
                    'color': 'Color:',
                    'gender': 'Genero:',
                    'silhouette': 'Silueta:'
                }
                prefix = label_prefixes.get(key, visual_name + ":")
                lines_data.append((prefix, f" {val}", False))
                
        return lines_data

    def generate_logic(self):
        try:
            self.log("🚀 Iniciando generación de presentación...")
            embedded_map = self.extract_embedded_images(self.excel_path.get())
            
            wb = openpyxl.load_workbook(self.excel_path.get(), data_only=True)
            sheet = wb.active
            
            m_indices = {}
            for key, var in self.map_vars.items():
                col_name = var.get()
                if col_name and col_name != "--- Ninguno ---" and col_name in self.excel_headers:
                    m_indices[key] = self.excel_headers.index(col_name)
                else:
                    m_indices[key] = None

            # --- MEJORA 1: Cargar plantilla PPTX si está seleccionada ---
            template_file = self.template_path.get()
            if template_file and os.path.isfile(template_file):
                self.log(f"🎨 Cargando plantilla personalizada: {os.path.basename(template_file)}")
                prs = Presentation(template_file)
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
                
                def get_val(key):
                    idx = m_indices[key]
                    if idx is not None:
                        if idx < len(row) and row[idx] is not None:
                            return str(row[idx]).strip()
                        else:
                            return "N/A"
                    return ""

                products.append({
                    'sku': clean_sku, 
                    'name': get_val('name'), 
                    'color': get_val('color'), 
                    'price': get_val('price'), 
                    'sizes': get_val('sizes'), 
                    'temp': get_val('temp'), 
                    'gender': get_val('gender'),
                    'silhouette': get_val('silhouette'),
                    'row_idx': row_idx
                })

            self.log(f"📦 Total de productos encontrados en Excel: {len(products)}")
            
            total_slides = (len(products) + 1) // 2
            layout_idx = 6 if len(prs.slide_layouts) > 6 else 0

            for i in range(0, len(products), 2):
                progress = ((i / 2) / total_slides) * 100
                self.progress_var.set(progress)
                self.log(f"Generando Diapositiva {i//2 + 1}/{total_slides}...")
                
                slide = prs.slides.add_slide(prs.slide_layouts[layout_idx])
                sw, sh = prs.slide_width, prs.slide_height
                qw, qh = sw / 2, sh / 2
                pad = Inches(0.4)

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
                aspect = img_w / img_h
                img_pad = Inches(0.3)
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