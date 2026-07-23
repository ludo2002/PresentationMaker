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
from PIL import Image
import threading

class PresenterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("CatalogMaster v2.5 🚀 - Mapeo Dinámico y Silueta")
        self.root.geometry("650x750")
        self.root.configure(bg="#f0f0f0")
        self.excel_path = tk.StringVar()
        self.output_path = tk.StringVar()
        self.progress_var = tk.DoubleVar(value=0)
        
        # Agregamos 'silhouette' al diccionario de variables
        self.map_vars = {
            'sku': tk.StringVar(),
            'name': tk.StringVar(),
            'color': tk.StringVar(),
            'price': tk.StringVar(),
            'sizes': tk.StringVar(),
            'temp': tk.StringVar(),
            'gender': tk.StringVar(),
            'silhouette': tk.StringVar() # NUEVO
        }
        self.excel_headers = []
        self.header_row_idx = 1
        
        self.setup_ui()

    def setup_ui(self):
        main_frame = tk.Frame(self.root, bg="#f0f0f0", padx=20, pady=20)
        main_frame.pack(expand=True, fill="both")
        
        header = tk.Label(main_frame, text="CatalogMaster", font=("Arial", 20, "bold"), bg="#f0f0f0", fg="#333")
        header.pack(pady=(0, 5))
        sub_header = tk.Label(main_frame, text="Professional Product Presentation Generator", font=("Arial", 10), bg="#f0f0f0", fg="#777")
        sub_header.pack(pady=(0, 15))
        
        self.create_file_row(main_frame, "Excel File (.xlsx):", self.excel_path, self.browse_excel)
        self.create_file_row(main_frame, "Save As (.pptx):", self.output_path, self.browse_output)
        
        self.mapping_frame = tk.LabelFrame(main_frame, text="Asignación de Columnas (Revisa y edita)", bg="#f0f0f0", font=("Arial", 10, "bold"), padx=10, pady=10)
        self.mapping_frame.pack(fill="x", pady=15)
        
        # Agregamos 'silhouette' a las etiquetas visuales
        labels = {
            'sku': 'SKU / Modelo:',
            'name': 'Nombre:',
            'color': 'Color:',
            'price': 'Precio:',
            'sizes': 'Tallas:',
            'temp': 'Temporada:',
            'gender': 'Género:',
            'silhouette': 'Silueta:' # NUEVO
        }
        
        row_idx = 0
        col_idx = 0
        self.comboboxes = {}
        for key, text_label in labels.items():
            frame = tk.Frame(self.mapping_frame, bg="#f0f0f0")
            frame.grid(row=row_idx, column=col_idx, padx=5, pady=5, sticky="ew")
            
            lbl = tk.Label(frame, text=text_label, bg="#f0f0f0", width=12, anchor="w", font=("Arial", 9))
            lbl.pack(side="left")
            
            cb = ttk.Combobox(frame, textvariable=self.map_vars[key], state="readonly", width=15)
            cb.pack(side="left", padx=5)
            self.comboboxes[key] = cb
            
            col_idx += 1
            if col_idx > 1:
                col_idx = 0
                row_idx += 1

        self.btn_generate = tk.Button(main_frame, text="GENERATE PRESENTATION", command=self.start_generation, bg="#007bff", fg="white", font=("Arial", 12, "bold"), padx=20, pady=10, cursor="hand2", relief="flat")
        self.btn_generate.pack(pady=10)
        self.progress_bar = ttk.Progressbar(main_frame, variable=self.progress_var, maximum=100)
        self.progress_bar.pack(fill="x", pady=(0, 10))
        self.log_area = scrolledtext.ScrolledText(main_frame, height=8, font=("Consolas", 9), state='disabled', bg="#fff", fg="#555")
        self.log_area.pack(fill="x")

    def create_file_row(self, parent, label_text, var, browse_func):
        row = tk.Frame(parent, bg="#f0f0f0", pady=5)
        row.pack(fill="x")
        lbl = tk.Label(row, text=label_text, font=("Arial", 10), bg="#f0f0f0", width=15, anchor="w")
        lbl.pack(side="left")
        ent = tk.Entry(row, textvariable=var, font=("Arial", 10))
        ent.pack(side="left", expand=True, fill="x", padx=5)
        btn = tk.Button(row, text="Browse...", command=browse_func, font=("Arial", 9), cursor="hand2")
        btn.pack(side="right")

    def browse_excel(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if path: 
            self.excel_path.set(path)
            self.load_excel_headers()

    def browse_output(self):
        path = filedialog.asksaveasfilename(defaultextension=".pptx", filetypes=[("PowerPoint files", "*.pptx")])
        if path: self.output_path.set(path)
        
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
            # Adivinanza para la nueva variable
            self.map_vars['silhouette'].set(guess(['silueta', 'silhouette', 'estilo', 'corte'])) # NUEVO
            
            self.log("✅ Encabezados cargados. Revisa la sección de Asignación de Columnas.")

        except Exception as e:
            self.log(f"❌ Error al cargar encabezados: {e}")

    def log(self, message):
        self.log_area.configure(state='normal')
        self.log_area.insert(tk.END, message + "\n")
        self.log_area.see(tk.END)
        self.log_area.configure(state='disabled')
        self.root.update_idletasks()

    def start_generation(self):
        if not self.excel_path.get() or not self.output_path.get():
            messagebox.showerror("Error", "Please select at least the Excel file and the output location.")
            return
        
        if not self.map_vars['sku'].get() or self.map_vars['sku'].get() == "--- Ninguno ---":
            messagebox.showerror("Error", "Por favor asigna al menos la columna para 'SKU / Modelo'.")
            return
            
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
                        'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
                        'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
                    }
                    for anchor in root.findall('.//xdr:twoCellAnchor', ns) + root.findall('.//xdr:oneCellAnchor', ns):
                        from_node = anchor.find('xdr:from', ns)
                        if from_node is not None:
                            row = int(from_node.find('xdr:row', ns).text)
                            col = int(from_node.find('xdr:col', ns).text)
                            pic = anchor.find('.//xdr:pic', ns)
                            if pic is not None:
                                blip = pic.find('.//a:blip', ns)
                                if blip is not None:
                                    rid = blip.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
                                    rels_path = drawing_file.replace('.xml', '.xml.rels')
                                    if rels_path in z.namelist():
                                        rels_content = z.read(rels_path)
                                        rels_root = ET.fromstring(rels_content)
                                        rel = rels_root.find(f".//r:Relationship[@Id='{rid}']", ns)
                                        if rel is not None:
                                            target = rel.get('Target')
                                            full_img_path = 'xl/' + target if not target.startswith('xl/') else target
                                            if full_img_path in z.namelist():
                                                embedded_images[(row, col)] = z.read(full_img_path)
        except Exception as e:
            self.log(f"⚠️ Error extrayendo imágenes embebidas: {e}")
        return embedded_images

    def build_dynamic_text(self, product):
        """Construye el texto de la diapositiva omitiendo las líneas que no están mapeadas."""
        lines = []
        
        if product['name']: 
            lines.append(f"NOMBRE: {product['name'].upper()}\n")
            
        if product['temp']: lines.append(f"Temp: {product['temp']}")
        if product['color']: lines.append(f"Color: {product['color']}")
        if product['sizes']: lines.append(f"Talla: {product['sizes']}")
        if product['gender']: lines.append(f"Genero: {product['gender']}")
        # Agregamos la línea de silueta a la generación dinámica de texto
        if product['silhouette']: lines.append(f"Silueta: {product['silhouette']}") # NUEVO
        
        if product['price']: 
            lines.append(f"\nPrecio: ${product['price']}")
            
        return "\n".join(lines)

    def generate_logic(self):
        try:
            self.log("🚀 Starting generation...")
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

            prs = Presentation()
            prs.slide_width, prs.slide_height = Inches(10), Inches(7.5)
            products = []
            
            for row_idx, row in enumerate(sheet.iter_rows(min_row=self.header_row_idx + 1, values_only=True), start=self.header_row_idx + 1):
                sku_idx = m_indices['sku']
                sku_val = row[sku_idx] if sku_idx is not None and sku_idx < len(row) else None
                if not sku_val: continue
                
                img_bytes = None
                for (r, c), data in embedded_map.items():
                    if r + 1 == row_idx:
                        img_bytes = data
                        break
                
                final_url = f"https://s3.amazonaws.com/cdn.grabasaweb.mx/Imagenes/Imagenes_Preventa/{sku_val}.jpg"
                
                def get_val(key):
                    idx = m_indices[key]
                    if idx is not None:
                        if idx < len(row) and row[idx] is not None:
                            return str(row[idx])
                        else:
                            return "N/A"
                    return ""

                # Pasamos el dato extraído al diccionario del producto
                products.append({
                    'sku': str(sku_val), 
                    'name': get_val('name'), 
                    'color': get_val('color'), 
                    'price': get_val('price'), 
                    'sizes': get_val('sizes'), 
                    'temp': get_val('temp'), 
                    'gender': get_val('gender'),
                    'silhouette': get_val('silhouette'), # NUEVO
                    'url': final_url,
                    'embedded_img': img_bytes
                })
            
            total_slides = (len(products) + 1) // 2
            for i in range(0, len(products), 2):
                progress = ((i / 2) / total_slides) * 100
                self.progress_var.set(progress)
                self.log(f"Slide {i//2 + 1}/{total_slides}...")
                slide = prs.slides.add_slide(prs.slide_layouts[6])
                
                sw, sh = prs.slide_width, prs.slide_height
                qw, qh = sw / 2, sh / 2
                pad = Inches(0.4)

                # --- PRODUCTO 1 ---
                p1 = products[i]
                if p1['embedded_img']:
                    try:
                        with Image.open(BytesIO(p1['embedded_img'])) as img:
                            img_w, img_h = img.size
                            aspect = img_w / img_h
                            img_pad = Inches(0.3)
                            max_w = qw - 2*img_pad
                            max_h = qh - 2*img_pad
                            if aspect > (max_w / max_h):
                                final_w = max_w; final_h = max_w / aspect
                            else:
                                final_h = max_h; final_w = max_h * aspect
                            off_x = 0 + (qw - final_w) / 2
                            off_y = 0 + (qh - final_h) / 2
                            slide.shapes.add_picture(BytesIO(p1['embedded_img']), off_x, off_y, final_w, final_h)
                    except Exception as e:
                        self.log(f"⚠️ Error procesando imagen embebida p1: {e}")
                        if p1['url']: self.add_image_to_quadrant(slide, p1['url'], 0, 0, qw, qh)
                elif p1['url']:
                    self.add_image_to_quadrant(slide, p1['url'], 0, 0, qw, qh)
                
                p1_text = self.build_dynamic_text(p1)
                self.create_box(slide, qw + pad, pad, qw - 2*pad, qh - 2*pad, p1_text, 18)

                # --- PRODUCTO 2 ---
                if i + 1 < len(products):
                    p2 = products[i+1]
                    
                    p2_text = self.build_dynamic_text(p2)
                    self.create_box(slide, pad, qh + pad, qw - 2*pad, qh - 2*pad, p2_text, 18)
                    
                    if p2['embedded_img']:
                        try:
                            with Image.open(BytesIO(p2['embedded_img'])) as img:
                                img_w, img_h = img.size
                                aspect = img_w / img_h
                                img_pad = Inches(0.3)
                                max_w = qw - 2*img_pad
                                max_h = qh - 2*img_pad
                                if aspect > (max_w / max_h):
                                    final_w = max_w; final_h = max_w / aspect
                                else:
                                    final_h = max_h; final_w = max_h * aspect
                                off_x = qw + (qw - final_w) / 2
                                off_y = qh + (qh - final_h) / 2
                                slide.shapes.add_picture(BytesIO(p2['embedded_img']), off_x, off_y, final_w, final_h)
                        except Exception as e:
                            self.log(f"⚠️ Error procesando imagen embebida p2: {e}")
                            if p2['url']: self.add_image_to_quadrant(slide, p2['url'], qw, qh, qw, qh)
                    elif p2['url']:
                        self.add_image_to_quadrant(slide, p2['url'], qw, qh, qw, qh)

            self.progress_var.set(100)
            prs.save(self.output_path.get())
            self.log("\n✅ SUCCESS! File saved.")
            messagebox.showinfo("Finished", "Presentation generated successfully!")
        except Exception as e:
            self.log(f"❌ ERROR: {str(e)}")
            messagebox.showerror("Error", f"An unexpected error occurred: {e}")
        finally:
            self.btn_generate.config(state='normal', text="GENERATE PRESENTATION")

    def add_image_to_quadrant(self, slide, url, left, top, width, height):
        try:
            resp = requests.get(url, timeout=10)
            if resp.status_code == 200:
                with Image.open(BytesIO(resp.content)) as img:
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
                    slide.shapes.add_picture(BytesIO(resp.content), off_x, off_y, final_w, final_h)
        except Exception as e:
            self.log(f"⚠️ Image error ({url}): {e}")

    def create_box(self, slide, left, top, width, height, text, font_size, alignment=PP_ALIGN.LEFT, bold=False):
        txBox = slide.shapes.add_textbox(left, top, width, height)
        tf = txBox.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = str(text) if text else "N/A"
        p.alignment = alignment
        p.font.size = Pt(font_size)
        p.font.bold = bold
        p.font.name = 'Arial'
        p.font.color.rgb = RGBColor(0, 0, 0)

if __name__ == "__main__":
    root = tk.Tk()
    app = PresenterApp(root)
    root.mainloop()