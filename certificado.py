import tkinter as tk    # importa la librería tkinter, sirve para crear interfaces gráficas
from tkinter import filedialog, messagebox, ttk # importa las librerías tkinter, filedialog, messagebox y ttk
import pandas as pd # importa la librería pandas, sirve para leer archivos de Excel
#from reportlab.pdfgen import canvas # importa la librería canvas, sirve para crear un PDF
from reportlab.lib.pagesizes import A4  # importa la librería A4, sirve para establecer el tamaño de la página
#from reportlab.pdfbase import pdfmetrics    # importa la librería pdfmetrics, sirve para trabajar con métricas de PDF
from reportlab.pdfbase.pdfmetrics import stringWidth    # importa la librería stringWidth, sirve para calcular el ancho de un texto      
#from reportlab.pdfbase.ttfonts import TTFont    # importa la librería TTFont, sirve para trabajar con fuentes TrueType
#from PyPDF2 import PdfReader, PdfWriter   # importa la librería PyPDF2, sirve para leer y escribir archivos PDF
import fitz  # PyMuPDF
from PIL import Image, ImageTk
import io
import os   # importa la librería os, sirve para interactuar con el sistema operativo
import logging # importa la librería logging, sirve para guardar errores en un archivo de texto
import json # importa la librería json, sirve para leer y escribir archivos JSON
import threading
#from functools import wraps

def debounce(wait):
    def decorator(fn):
        def debounced(*args, **kwargs):
            def call_it():
                fn(*args, **kwargs)
            try:
                debounced.t.cancel()
            except(AttributeError):
                pass
            debounced.t = threading.Timer(wait, call_it)
            debounced.t.start()
        return debounced
    return decorator


# python -m venv env  # Crear un entorno virtual
# \env\Scripts> Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser 
# \env\Scripts> .\Activate.ps1 
# pip install pandas reportlab pillow openpyxl PyMuPDF
# https://github.com/oschwartz10612/poppler-windows/releases 
# Extraer en C:\poppler\Library\bin 
# Agregar C:\poppler\Library\bin al PATH 
# Agregar pyinstaller con:
# pip install pyinstaller
# 
# pyinstaller --onefile -w .\certificado.py --icon=certificado.png
#
# Generar instalador con nullsoft installer
#  





# Configurar logging
logging.basicConfig(filename='errores_certificados.log', level=logging.WARNING)

def dividir_texto_en_lineas(texto, fuente, tamaño, ancho_maximo):
    """
    Divide un texto en múltiples líneas si excede el ancho máximo.
    Retorna una lista de líneas.
    """
    if ancho_maximo <= 0:  # Si no hay límite, devolver el texto completo
        return [texto]

    palabras = texto.split()
    lineas = []
    linea_actual = []

    for palabra in palabras:
        linea_prueba = ' '.join(linea_actual + [palabra])
        ancho_prueba = stringWidth(linea_prueba, fuente, tamaño)

        if ancho_prueba <= ancho_maximo:
            linea_actual.append(palabra)
        else:
            if linea_actual:
                lineas.append(' '.join(linea_actual))
                linea_actual = [palabra]
            else:
                # Si una sola palabra es más ancha que el máximo, agregarla de todos modos
                lineas.append(palabra)

    if linea_actual:
        lineas.append(' '.join(linea_actual))

    return lineas if lineas else [texto]

def generar_certificados(plantilla_pdf, datos_excel, configuracion, directorio_salida, nombre_curso):
    df = pd.read_excel(datos_excel)

    # Verificar si existe la columna 'dni' en el Excel
    if 'dni' not in df.columns:
        # Si no existe 'dni', preguntar al usuario qué columna usar
        columnas_disponibles = list(df.columns)

        # Crear una ventana de diálogo para seleccionar la columna
        ventana_columna = tk.Toplevel()
        ventana_columna.title("Seleccionar Columna de Índice")
        ventana_columna.geometry("400x200")
        ventana_columna.transient()
        ventana_columna.grab_set()

        columna_seleccionada = tk.StringVar()

        ttk.Label(ventana_columna, text="No se encontró la columna 'dni'.\nSeleccione una columna para usar como índice:").pack(pady=10)

        combo_columnas = ttk.Combobox(ventana_columna, textvariable=columna_seleccionada, values=columnas_disponibles, width=30)
        combo_columnas.pack(pady=10)
        if columnas_disponibles:
            combo_columnas.set(columnas_disponibles[0])

        def confirmar():
            ventana_columna.destroy()

        def cancelar():
            columna_seleccionada.set("")
            ventana_columna.destroy()

        frame_botones = ttk.Frame(ventana_columna)
        frame_botones.pack(pady=10)
        ttk.Button(frame_botones, text="Confirmar", command=confirmar).pack(side=tk.LEFT, padx=5)
        ttk.Button(frame_botones, text="Cancelar", command=cancelar).pack(side=tk.LEFT, padx=5)

        # Cerrar con la X equivale a cancelar
        ventana_columna.protocol("WM_DELETE_WINDOW", cancelar)

        # Bloquea hasta que se cierre el diálogo, sin anidar otro mainloop
        ventana_columna.wait_window()

        if not columna_seleccionada.get():
            messagebox.showerror("Error", "Debe seleccionar una columna de índice o cargar un archivo Excel con la columna 'dni'.")
            return

        columna_indice = columna_seleccionada.get()
    else:
        columna_indice = 'dni'

    for index, row in df.iterrows():
        try:
            # Open the template PDF
            doc = fitz.open(plantilla_pdf)
            page = doc[0]  # Assume we're working with the first page

            # Create a new PDF for this certificate
            output_pdf = fitz.open()
            output_page = output_pdf.new_page(width=page.rect.width, height=page.rect.height)

            # Copy the content of the template to the new page
            output_page.show_pdf_page(page.rect, doc, 0)

            # Add text fields
            for campo, config in configuracion.items():
                if campo in row:
                    text = str(row[campo])
                    font_name = config['fuente']
                    font_size = config['tamaño']
                    is_bold = config.get('negrita', False)
                    is_italic = config.get('cursiva', False)
                    ancho_maximo = config.get('ancho_maximo', 0)  # 0 significa sin límite

                    match font_name:
                        case "Times-Roman":
                            font_name = "tiro"
                            if is_bold and is_italic:
                                font_name = "tibi"
                            elif is_bold:
                                font_name = "tibo"
                            elif is_italic:
                                font_name = "tiit"
                        case "Helvetica":
                            font_name = "helv"
                            if is_bold and is_italic:
                                font_name = "hebi"
                            elif is_bold:
                                font_name = "hebo"
                            elif is_italic:
                                font_name = "heit"
                        case "Courier":
                            font_name = "courier"
                            if is_bold and is_italic:
                                font_name = "cobi"
                            elif is_bold:
                                font_name = "cobo"
                            elif is_italic:
                                font_name = "coit"

                    # Dividir el texto en múltiples líneas si es necesario
                    lineas = dividir_texto_en_lineas(text, config['fuente'], config['tamaño'], ancho_maximo)

                    # Calcular la posición Y inicial (ajustada para centrar verticalmente si hay múltiples líneas)
                    y_base = config['y']
                    interlineado = config.get('interlineado') or font_size * 1.2  # 0 o ausente = automático

                    # Insertar cada línea
                    for i, linea in enumerate(lineas):
                        ancho_linea = stringWidth(linea, config['fuente'], config['tamaño'])

                        if config['alineacion'] == 'Centro':
                                x = config['x'] - (ancho_linea / 2)
                        elif config['alineacion'] == 'Derecha':
                                x = config['x'] - ancho_linea
                        else:  # 'Izquierda'
                                x = config['x']

                        y = y_base + (i * interlineado)

                        # Add text to the page
                        output_page.insert_text((x, y), linea, fontname=font_name, fontsize=font_size)


            # Save the certificate

            output_filename = f"{nombre_curso}_{str(row[columna_indice])}.pdf"
            output_path = os.path.join(directorio_salida, output_filename)
            output_pdf.save(output_path)
            output_pdf.close()

        except Exception as e:
            logging.error(f"Error al generar certificado para {columna_indice} {str(row[columna_indice])}: {str(e)}")

    messagebox.showinfo("Proceso Completado", "Todos los certificados han sido generados.")
    
    
    """_summary_def get_font_name(base_font, is_bold, is_italic):
    if base_font == "Times-Roman":
        if is_bold and is_italic:
            return "Times-BoldItalic" 
        elif is_bold:
            return "Times-Bold" 
        elif is_italic: 
            return "Times-Italic"
        else:
            return "Times-Roman"
    # Manejar otras fuentes de manera similar
    # Por ejemplo, para Helvetica:
    elif base_font == "Helvetica":
        if is_bold and is_italic:
            return "Helvetica-BoldOblique"
        elif is_bold:
            return "Helvetica-Bold"
        elif is_italic:
            return "Helvetica-Oblique"
        else:
            return "Helvetica"
    # Si no es una fuente conocida, devolver la base
    return base_font
    """

class VistaPreviaAvanzada(ttk.Frame):
    def __init__(self, master, app=None, *args, **kwargs):
        super().__init__(master, *args, **kwargs)
        self.app = app  # Referencia a la aplicación principal

        # Canvas con proporciones A4 apaisado (842x595 puntos)
        # Escala 0.70: 842×0.70=590, 595×0.70=417
        self.canvas = tk.Canvas(self, width=600, height=430, bg='white')
        self.canvas.pack(fill=tk.BOTH, expand=True)

        # Variables para interactividad
        self.pdf_image = None  # Imagen del PDF de fondo
        self.pdf_width = 0  # Ancho real del PDF en puntos
        self.pdf_height = 0  # Alto real del PDF en puntos
        self.scale = 1.0  # Escala del PDF mostrado
        self.offset_x = 0  # Offset X para centrar la imagen
        self.offset_y = 0  # Offset Y para centrar la imagen
        self.campos_rect = {}  # Diccionario de recuadros por campo
        self.campo_seleccionado = None  # Campo actualmente seleccionado
        self.drag_data = {"x": 0, "y": 0, "item": None, "mode": None}  # Datos para arrastre

        # Eventos del canvas
        self.canvas.bind("<Button-1>", self.on_canvas_click)
        self.canvas.bind("<B1-Motion>", self.on_canvas_drag)
        self.canvas.bind("<ButtonRelease-1>", self.on_canvas_release)

    def actualizar_vista_previa(self, plantilla_pdf, campos, nombre_curso, datos_registro=None):
        try:
            # Open the PDF
            doc = fitz.open(plantilla_pdf)
            page = doc[0]  # Get the first page

            # Guardar dimensiones reales del PDF
            self.pdf_width = page.rect.width
            self.pdf_height = page.rect.height

            # Create a new PDF with the fields
            output_pdf = fitz.open()
            output_page = output_pdf.new_page(width=self.pdf_width, height=self.pdf_height)

            # Copy the content of the template to the new page
            output_page.show_pdf_page(page.rect, doc, 0)

            # Limpiar campos anteriores
            self.campos_rect = {}

            # Add text fields
            for campo, config in campos.items():
                # Obtener texto del Excel o del campo texto_muestra
                if datos_registro is not None and campo in datos_registro:
                    text = str(datos_registro[campo])
                else:
                    text = config.get('texto_muestra', campo)

                font_name = config['fuente']
                font_size = config['tamaño']
                is_bold = config.get('negrita', False)
                is_italic = config.get('cursiva', False)
                ancho_maximo = config.get('ancho_maximo', 0)

                # Convertir nombre de fuente para PyMuPDF
                font_name_pdf = font_name
                match font_name:
                    case "Times-Roman":
                        font_name_pdf = "tiro"
                        if is_bold and is_italic:
                            font_name_pdf = "tibi"
                        elif is_bold:
                            font_name_pdf = "tibo"
                        elif is_italic:
                            font_name_pdf = "tiit"
                    case "Helvetica":
                        font_name_pdf = "helv"
                        if is_bold and is_italic:
                            font_name_pdf = "hebi"
                        elif is_bold:
                            font_name_pdf = "hebo"
                        elif is_italic:
                            font_name_pdf = "heit"
                    case "Courier":
                        font_name_pdf = "courier"
                        if is_bold and is_italic:
                            font_name_pdf = "cobi"
                        elif is_bold:
                            font_name_pdf = "cobo"
                        elif is_italic:
                            font_name_pdf = "coit"

                # Dividir el texto en múltiples líneas si es necesario
                lineas = dividir_texto_en_lineas(text, font_name, font_size, ancho_maximo)

                # Calcular la posición Y inicial
                y_base = config['y']
                interlineado = config.get('interlineado') or font_size * 1.2  # 0 o ausente = automático

                # Calcular dimensiones del recuadro
                max_ancho = max([stringWidth(linea, font_name, font_size) for linea in lineas]) if lineas else 0
                # Altura: desde la primera línea hasta el final de la última línea
                # Si hay N líneas, la última está en (N-1)*interlineado + altura de una línea
                altura_total = ((len(lineas) - 1) * interlineado + font_size) if lineas else font_size

                # Calcular la posición X donde realmente empieza el recuadro según la alineación
                # Usar el ancho máximo configurado si existe, sino el ancho real del texto
                ancho_rect = ancho_maximo if ancho_maximo > 0 else max_ancho

                if config['alineacion'] == 'Centro':
                    x_inicio_rect = config['x'] - (ancho_rect / 2)
                elif config['alineacion'] == 'Derecha':
                    x_inicio_rect = config['x'] - ancho_rect
                else:  # 'Izquierda'
                    x_inicio_rect = config['x']

                # PyMuPDF usa Y como baseline del texto, no como top
                # El recuadro debe empezar más arriba: baseline - ascent
                # Típicamente ascent es ~0.75-0.8 del font_size
                y_top_rect = y_base - (font_size * 0.75)

                # Guardar información del recuadro para este campo
                self.campos_rect[campo] = {
                    'x_ref': config['x'],  # Punto de referencia (centro, derecha o izquierda según alineación)
                    'x_inicio': x_inicio_rect,  # Donde realmente empieza el recuadro
                    'y': y_top_rect,  # Parte superior del recuadro (ajustado desde baseline)
                    'y_baseline': y_base,  # Guardar también el baseline para referencia
                    'ancho': ancho_rect,
                    'altura': altura_total,
                    'config': config,
                    'texto': text,
                    'lineas': lineas,  # Guardar las líneas para referencia
                    'font_size': font_size  # Guardar font_size para recalcular
                }

                # Insertar cada línea en el PDF
                for i, linea in enumerate(lineas):
                    ancho_linea = stringWidth(linea, font_name, font_size)

                    if config['alineacion'] == 'Centro':
                        x = config['x'] - (ancho_linea / 2)
                    elif config['alineacion'] == 'Derecha':
                        x = config['x'] - ancho_linea
                    else:  # 'Izquierda'
                        x = config['x']

                    y = y_base + (i * interlineado)

                    # Add text to the page
                    output_page.insert_text((x, y), linea, fontname=font_name_pdf, fontsize=font_size)

            # Convert PDF to image sin escalar (matriz identidad)
            pix = output_page.get_pixmap()
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

            # Calculate scale to fit canvas (basada en dimensiones REALES del PDF)
            canvas_width = self.canvas.winfo_width() if self.canvas.winfo_width() > 1 else 600
            canvas_height = self.canvas.winfo_height() if self.canvas.winfo_height() > 1 else 430

            scale_x = canvas_width / self.pdf_width
            scale_y = canvas_height / self.pdf_height
            self.scale = min(scale_x, scale_y)  # Usar la menor escala para que entre completo

            # Tamaño de la imagen escalada
            new_width = int(self.pdf_width * self.scale)
            new_height = int(self.pdf_height * self.scale)
            img = img.resize((new_width, new_height), Image.Resampling.LANCZOS)

            # Calcular offset para centrar la imagen
            self.offset_x = (canvas_width - new_width) // 2
            self.offset_y = (canvas_height - new_height) // 2

            # Display in canvas centrado
            self.photo = ImageTk.PhotoImage(img)
            self.canvas.delete("all")
            self.canvas.create_image(self.offset_x, self.offset_y, anchor="nw", image=self.photo)

            # Dibujar recuadros sobre los campos
            self.dibujar_recuadros()

            output_pdf.close()
            doc.close()

        except Exception as e:
            print(f"Error updating preview: {e}")
            import traceback
            traceback.print_exc()

    def dibujar_recuadros(self):
        """Dibuja los recuadros interactivos sobre cada campo"""
        # Borrar todos los recuadros anteriores para evitar estela
        self.canvas.delete("recuadro")
        self.canvas.delete("handle")

        for campo, rect_info in self.campos_rect.items():
            # Convertir coordenadas del PDF a coordenadas del canvas (con escala y offset)
            x1 = (rect_info['x_inicio'] * self.scale) + self.offset_x
            y1 = (rect_info['y'] * self.scale) + self.offset_y
            ancho = rect_info['ancho'] * self.scale
            altura = rect_info['altura'] * self.scale

            x2 = x1 + ancho
            y2 = y1 + altura

            # Color del recuadro (azul normal, verde seleccionado)
            color = '#00ff00' if campo == self.campo_seleccionado else '#0080ff'

            # Dibujar recuadro principal (usando tags genéricos para poder borrar)
            rect_id = self.canvas.create_rectangle(x1, y1, x2, y2, outline=color, width=2, tags=("recuadro", f"rect_{campo}"))

            # Dibujar handles de redimensionamiento en las esquinas
            handle_size = 6
            # Esquina inferior derecha (para redimensionar ancho)
            handle_id = self.canvas.create_rectangle(
                x2 - handle_size, y2 - handle_size, x2 + handle_size, y2 + handle_size,
                fill=color, outline=color, tags=("handle", f"handle_{campo}")
            )

            # Guardar IDs de los elementos del canvas
            rect_info['rect_id'] = rect_id
            rect_info['handle_id'] = handle_id

    def on_canvas_click(self, event):
        """Maneja el clic en el canvas"""
        # Buscar si se hizo clic en un handle
        for campo, rect_info in self.campos_rect.items():
            handle_id = rect_info.get('handle_id')
            if handle_id and 'current' in self.canvas.gettags(handle_id):
                items = self.canvas.find_overlapping(event.x - 5, event.y - 5, event.x + 5, event.y + 5)
                if handle_id in items:
                    self.campo_seleccionado = campo
                    self.drag_data = {"x": event.x, "y": event.y, "item": handle_id, "mode": "resize", "campo": campo}
                    self.dibujar_recuadros()  # Redibujar para resaltar
                    return

        # Buscar si se hizo clic en un recuadro
        for campo, rect_info in self.campos_rect.items():
            rect_id = rect_info.get('rect_id')
            if rect_id:
                items = self.canvas.find_overlapping(event.x, event.y, event.x, event.y)
                if rect_id in items:
                    self.campo_seleccionado = campo
                    self.drag_data = {"x": event.x, "y": event.y, "item": rect_id, "mode": "move", "campo": campo}
                    self.dibujar_recuadros()  # Redibujar para resaltar
                    return

        # Si no se hizo clic en ningún recuadro, deseleccionar
        self.campo_seleccionado = None
        self.dibujar_recuadros()

    def on_canvas_drag(self, event):
        """Maneja el arrastre en el canvas"""
        if self.drag_data["mode"] == "move":
            # Mover el campo (convertir delta de canvas a coordenadas PDF)
            dx = (event.x - self.drag_data["x"]) / self.scale
            dy = (event.y - self.drag_data["y"]) / self.scale

            campo = self.drag_data["campo"]
            rect_info = self.campos_rect[campo]
            config = rect_info['config']

            # Actualizar el punto de referencia (en coordenadas PDF)
            rect_info['x_ref'] += dx
            rect_info['y'] += dy  # y es la parte superior del recuadro
            rect_info['y_baseline'] += dy  # También actualizar el baseline

            # Recalcular x_inicio según alineación
            if config['alineacion'] == 'Centro':
                rect_info['x_inicio'] = rect_info['x_ref'] - (rect_info['ancho'] / 2)
            elif config['alineacion'] == 'Derecha':
                rect_info['x_inicio'] = rect_info['x_ref'] - rect_info['ancho']
            else:  # 'Izquierda'
                rect_info['x_inicio'] = rect_info['x_ref']

            self.drag_data["x"] = event.x
            self.drag_data["y"] = event.y

            # Redibujar
            self.dibujar_recuadros()

        elif self.drag_data["mode"] == "resize":
            # Redimensionar el ancho del campo (convertir delta de canvas a coordenadas PDF)
            dx = (event.x - self.drag_data["x"]) / self.scale

            campo = self.drag_data["campo"]
            rect_info = self.campos_rect[campo]
            config = rect_info['config']

            # Calcular nuevo ancho (en coordenadas PDF)
            nuevo_ancho = rect_info['ancho'] + dx

            if nuevo_ancho > 10:  # Ancho mínimo
                rect_info['ancho'] = nuevo_ancho

                # Recalcular x_inicio según alineación (importante para Centro y Derecha)
                if config['alineacion'] == 'Centro':
                    rect_info['x_inicio'] = rect_info['x_ref'] - (nuevo_ancho / 2)
                elif config['alineacion'] == 'Derecha':
                    rect_info['x_inicio'] = rect_info['x_ref'] - nuevo_ancho
                else:  # 'Izquierda'
                    rect_info['x_inicio'] = rect_info['x_ref']

                self.drag_data["x"] = event.x

                # Redibujar
                self.dibujar_recuadros()

    def on_canvas_release(self, event):
        """Maneja cuando se suelta el mouse"""
        if self.drag_data["mode"] in ["move", "resize"] and self.app:
            campo = self.drag_data["campo"]
            rect_info = self.campos_rect[campo]

            # Actualizar los valores en la configuración de la aplicación
            if self.drag_data["mode"] == "move":
                # Actualizar X (punto de referencia) e Y (baseline, no top)
                self.actualizar_campo_en_ui(campo, 'x', int(rect_info['x_ref']))
                self.actualizar_campo_en_ui(campo, 'y', int(rect_info['y_baseline']))

            elif self.drag_data["mode"] == "resize":
                # Actualizar ancho_maximo
                self.actualizar_campo_en_ui(campo, 'ancho_maximo', int(rect_info['ancho']))

            # Actualizar la vista previa para mostrar el texto en la nueva posición
            self.app.actualizar_vista_previa()

        # Limpiar datos de arrastre
        self.drag_data = {"x": 0, "y": 0, "item": None, "mode": None}

    def actualizar_campo_en_ui(self, campo_nombre, propiedad, valor):
        """Actualiza un campo en la interfaz de usuario"""
        if not self.app:
            return

        # Buscar el frame del campo en la configuración
        for frame in self.app.campos_frame.winfo_children():
            entries = [widget for widget in frame.winfo_children() if isinstance(widget, (ttk.Entry, ttk.Combobox))]
            if len(entries) >= 9:
                campo_entry = entries[0]
                if campo_entry.get() == campo_nombre:
                    # Encontramos el campo correcto, usar StringVar
                    if propiedad == 'x':
                        frame.x_var.set(str(valor))
                    elif propiedad == 'y':
                        frame.y_var.set(str(valor))
                    elif propiedad == 'ancho_maximo':
                        frame.ancho_maximo_var.set(str(valor))
                    break

class Application(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        self.master.title("Generador de Certificados")
        self.pack(fill=tk.BOTH, expand=True)
        self.excel_columns = []
        self.df_excel = None  # DataFrame del Excel cargado
        self.registro_actual = 0  # Índice del registro actual en la vista previa
        self.create_widgets()
        self.campo_config = {}
        self.cargar_config('ultima_config.json')  # Cargar la última configuración al iniciar
        # print(self.datos_entry.get())
        if self.datos_entry.get() != '':
            filename = self.datos_entry.get()
            self.cargar_columnas_excel(filename)


    def create_widgets(self):
        # Frame principal
        main_frame = ttk.Frame(self, padding="10 10 10 10")
        main_frame.grid(column=0, row=0, sticky=(tk.N, tk.W, tk.E, tk.S)) # type: ignore
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)

        # Plantilla PDF
        ttk.Label(main_frame, text="Plantilla PDF:").grid(column=0, row=0, sticky=tk.W)
        self.plantilla_entry = ttk.Entry(main_frame, width=50)
        self.plantilla_entry.grid(column=1, row=0, sticky=(tk.W, tk.E))
        ttk.Button(main_frame, text="Buscar", command=self.buscar_plantilla).grid(column=2, row=0)

        # Datos Excel
        ttk.Label(main_frame, text="Datos Excel:").grid(column=0, row=1, sticky=tk.W)
        self.datos_entry = ttk.Entry(main_frame, width=50)
        self.datos_entry.grid(column=1, row=1, sticky=(tk.W, tk.E))
        ttk.Button(main_frame, text="Buscar", command=self.buscar_datos).grid(column=2, row=1)

        # Nombre del Curso
        ttk.Label(main_frame, text="Nombre del Curso:").grid(column=0, row=2, sticky=tk.W)
        self.curso_entry = ttk.Entry(main_frame, width=50)
        self.curso_entry.grid(column=1, row=2, sticky=(tk.W, tk.E))

        # Directorio de Salida
        ttk.Label(main_frame, text="Directorio de Salida:").grid(column=0, row=3, sticky=tk.W)
        self.salida_entry = ttk.Entry(main_frame, width=50)
        self.salida_entry.grid(column=1, row=3, sticky=(tk.W, tk.E))
        ttk.Button(main_frame, text="Buscar", command=self.buscar_salida).grid(column=2, row=3)

        # Configuración de campos
        ttk.Label(main_frame, text="Configuración de Campos:").grid(column=0, row=4, sticky=tk.W, columnspan=3)
        self.campos_frame = ttk.Frame(main_frame)
        self.campos_frame.grid(column=0, row=5, columnspan=3, sticky=(tk.W, tk.E))
        self.agregar_campo()

        # Botones
        ttk.Button(main_frame, text="Agregar Campo", command=self.agregar_campo).grid(column=0, row=6)
        ttk.Button(main_frame, text="Generar Certificados", command=self.generar).grid(column=1, row=6)
        ttk.Button(main_frame, text="Guardar Configuración", command=self.guardar_config).grid(column=2, row=6)
        ttk.Button(main_frame, text="Cargar Configuración", command=self.cargar_config).grid(column=0, row=7)
        

        # Vista previa y navegador
        preview_frame = ttk.Frame(main_frame)
        preview_frame.grid(column=3, row=0, rowspan=15, padx=10, pady=10, sticky=(tk.N, tk.S, tk.E, tk.W))

        self.vista_previa = VistaPreviaAvanzada(preview_frame, app=self)
        self.vista_previa.pack(fill=tk.BOTH, expand=True)

        # Navegador de registros
        nav_frame = ttk.Frame(preview_frame)
        nav_frame.pack(fill=tk.X, pady=5)

        self.btn_anterior = ttk.Button(nav_frame, text="◄ Anterior", command=self.registro_anterior)
        self.btn_anterior.pack(side=tk.LEFT, padx=5)

        self.label_registro = ttk.Label(nav_frame, text="Sin datos")
        self.label_registro.pack(side=tk.LEFT, padx=10)

        self.btn_siguiente = ttk.Button(nav_frame, text="Siguiente ►", command=self.registro_siguiente)
        self.btn_siguiente.pack(side=tk.LEFT, padx=5)

        self.actualizar_navegador()

        for child in main_frame.winfo_children():
            child.grid_configure(padx=5, pady=5) # Add padding to all widgets
        self.curso_entry.bind('<KeyRelease>', lambda e: self.actualizar_vista_previa())

    def actualizar_navegador(self):
        """Actualiza el estado del navegador de registros"""
        if self.df_excel is not None and len(self.df_excel) > 0:
            total = len(self.df_excel)
            self.label_registro.config(text=f"Registro {self.registro_actual + 1} de {total}")
            self.btn_anterior.config(state='normal' if self.registro_actual > 0 else 'disabled')
            self.btn_siguiente.config(state='normal' if self.registro_actual < total - 1 else 'disabled')
        else:
            self.label_registro.config(text="Sin datos")
            self.btn_anterior.config(state='disabled')
            self.btn_siguiente.config(state='disabled')

    def registro_anterior(self):
        """Navega al registro anterior"""
        if self.df_excel is not None and self.registro_actual > 0:
            self.registro_actual -= 1
            self.actualizar_navegador()
            self.actualizar_vista_previa()

    def registro_siguiente(self):
        """Navega al registro siguiente"""
        if self.df_excel is not None and self.registro_actual < len(self.df_excel) - 1:
            self.registro_actual += 1
            self.actualizar_navegador()
            self.actualizar_vista_previa()

    @debounce(1)  # Debounce with a 0.5 second delay
    def actualizar_vista_previa(self):
        plantilla = self.plantilla_entry.get()
        campos = self.obtener_config_campos()
        nombre_curso = self.curso_entry.get()
        if plantilla and campos and nombre_curso:
            try:
                # Obtener datos del registro actual si hay Excel cargado
                datos_registro = None
                if self.df_excel is not None and len(self.df_excel) > 0:
                    datos_registro = self.df_excel.iloc[self.registro_actual]

                self.vista_previa.actualizar_vista_previa(plantilla, campos, nombre_curso, datos_registro)

            except Exception as e:
                logging.error(f"Error updating preview: {e}")
    



    def agregar_campo(self):
        frame = ttk.Frame(self.campos_frame)
        frame.grid(sticky='ew')  # Use grid instead of pack
        
        row = len(self.campos_frame.grid_slaves())  # Get the number of existing rows
        
        ttk.Label(frame, text="Campo:").grid(row=0, column=0, sticky='w')
        campo = ttk.Combobox(frame, values=self.excel_columns, width=10)
        campo.grid(row=0, column=1, padx=4, sticky='w')

        ttk.Label(frame, text="Muestra:").grid(row=0, column=2, sticky='w')
        texto_muestra_var = tk.StringVar(value="Texto de ejemplo")
        texto_muestra = ttk.Entry(frame, width=15, textvariable=texto_muestra_var)
        texto_muestra.grid(row=0, column=3, padx=4,pady=4, sticky='w')
        frame.texto_muestra_var = texto_muestra_var  # Guardar referencia

        ttk.Label(frame, text="X: (420-A4)").grid(row=0, column=4, sticky='w')
        x_var = tk.StringVar(value="420")
        x = ttk.Entry(frame, width=5, textvariable=x_var)
        x.grid(row=0, column=5, padx=4, sticky='w')
        frame.x_var = x_var  # Guardar referencia

        ttk.Label(frame, text="Y:").grid(row=0, column=6, sticky='w')
        y_var = tk.StringVar(value="270")
        y = ttk.Entry(frame, width=5, textvariable=y_var)
        y.grid(row=0, column=7, padx=4, sticky='w')
        frame.y_var = y_var  # Guardar referencia

        ttk.Label(frame, text="Fuente:").grid(row=0, column=8, sticky='w')
        fuentes = ["Helvetica", "Times-Roman", "Courier"]
        fuente = ttk.Combobox(frame, values=fuentes, width=10)
        fuente.set(fuentes[1])
        fuente.grid(row=0, column=9, padx=4, sticky='w')

        ttk.Label(frame, text="Tamaño:").grid(row=1, column=0, sticky='w')
        tamanos = list(range(8, 73, 2))
        tamano = ttk.Combobox(frame, values=tamanos, width=5)
        tamano.set(20)
        tamano.grid(row=1, column=1, padx=2, sticky='w')

        ttk.Label(frame, text="Alineación:").grid(row=1, column=2, sticky='w')
        alineaciones = ["Izquierda", "Centro", "Derecha"]
        alineacion = ttk.Combobox(frame, values=alineaciones, width=8)
        alineacion.set("Centro")
        alineacion.grid(row=1, column=3, padx=4, sticky='w')

        ttk.Label(frame, text="Estilo:").grid(row=1, column=4, sticky='w')
        bold_var = tk.BooleanVar(value=True)  # Negrita activada por defecto
        bold_check = ttk.Checkbutton(frame, text="Negrita", variable=bold_var)
        bold_check.grid(row=1, column=5, padx=4, sticky='w')
        italic_var = tk.BooleanVar(value=False)
        italic_check = ttk.Checkbutton(frame, text="Cursiva", variable=italic_var)
        italic_check.grid(row=1, column=6, padx=4, sticky='w')
        frame.bold_var = bold_var
        frame.italic_var = italic_var

        # Nuevo campo: Ancho máximo para dividir texto en líneas
        ttk.Label(frame, text="Ancho máx:").grid(row=2, column=0, sticky='w')
        ancho_maximo_var = tk.StringVar(value="600")
        ancho_maximo = ttk.Entry(frame, width=5, textvariable=ancho_maximo_var)
        ancho_maximo.grid(row=2, column=1, padx=4, sticky='w')
        ttk.Label(frame, text="(0=sin límite)").grid(row=2, column=2, sticky='w')
        frame.ancho_maximo_var = ancho_maximo_var  # Guardar referencia

        # Nuevo campo: Interlineado
        ttk.Label(frame, text="Interlineado:").grid(row=2, column=3, sticky='w')
        interlineado_var = tk.StringVar(value="25")
        interlineado = ttk.Entry(frame, width=5, textvariable=interlineado_var)
        interlineado.grid(row=2, column=4, padx=4, sticky='w')
        ttk.Label(frame, text="(0=auto)").grid(row=2, column=5, columnspan=2, sticky='w')
        frame.interlineado_var = interlineado_var  # Guardar referencia

        ttk.Button(frame, text="Eliminar", command=lambda: frame.destroy()).grid(row=2, column=7, padx=2, sticky='e')
            # Add a separator after the "Eliminar" button
        separator = ttk.Separator(frame, orient='horizontal')
        separator.grid(row=3, column=0, columnspan=10, sticky='ew', pady=5)

        # Store the separator reference in the frame for later use
        frame.separator = separator
        

        # Add callbacks for updating preview
        self.plantilla_entry.bind('<KeyRelease>', lambda e: self.actualizar_vista_previa())
        self.datos_entry.bind('<KeyRelease>', lambda e: self.actualizar_vista_previa())
        self.curso_entry.bind('<KeyRelease>', lambda e: self.actualizar_vista_previa())
        self.salida_entry.bind('<KeyRelease>', lambda e: self.actualizar_vista_previa())
        campo.bind('<KeyRelease>', lambda e: self.actualizar_vista_previa())
        texto_muestra.bind('<KeyRelease>', lambda e: self.actualizar_vista_previa())
        x.bind('<KeyRelease>', lambda e: self.actualizar_vista_previa())
        y.bind('<KeyRelease>', lambda e: self.actualizar_vista_previa())
        fuente.bind('<<ComboboxSelected>>', lambda e: self.actualizar_vista_previa())
        tamano.bind('<<ComboboxSelected>>', lambda e: self.actualizar_vista_previa())
        alineacion.bind('<<ComboboxSelected>>', lambda e: self.actualizar_vista_previa())
        bold_var.trace_add('write', lambda *args: self.actualizar_vista_previa())
        italic_var.trace_add('write', lambda *args: self.actualizar_vista_previa())
        ancho_maximo.bind('<KeyRelease>', lambda e: self.actualizar_vista_previa())
        interlineado.bind('<KeyRelease>', lambda e: self.actualizar_vista_previa())

        return frame

    def buscar_plantilla(self):
        filename = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        self.plantilla_entry.delete(0, tk.END)
        self.plantilla_entry.insert(0, filename)
        self.actualizar_vista_previa()

    def buscar_datos(self):
        filename = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        self.datos_entry.delete(0, tk.END)
        self.datos_entry.insert(0, filename)
        self.cargar_columnas_excel(filename)

    def cargar_columnas_excel(self, filename):
        try:
            self.df_excel = pd.read_excel(filename)
            self.excel_columns = list(self.df_excel.columns)
            self.registro_actual = 0  # Resetear al primer registro

            for frame in self.campos_frame.winfo_children():
                campo = frame.winfo_children()[1]  # El combobox de campo está en la segunda posición
                campo['values'] = self.excel_columns

            # Actualizar el navegador de registros y vista previa
            self.actualizar_navegador()
            self.actualizar_vista_previa()
        except Exception as e:
            messagebox.showerror("Error", f"Error al cargar columnas del archivo Excel: {str(e)}")


    def buscar_salida(self):
        directory = filedialog.askdirectory()
        self.salida_entry.delete(0, tk.END)
        self.salida_entry.insert(0, directory)

    def obtener_config_campos(self):
        config = {}

        for frame in self.campos_frame.winfo_children():
            entries = [widget for widget in frame.winfo_children() if isinstance(widget, (ttk.Entry, ttk.Combobox))]
            if len(entries) >= 9:  # Ahora tenemos más campos (incluyendo ancho_maximo e interlineado)
                campo, texto_muestra, x, y, fuente, tamano, alineacion, ancho_maximo, interlineado = entries[:9]

                # Obtener ancho máximo, validar que sea número
                ancho_max_val = ancho_maximo.get() if ancho_maximo.get().replace('.', '', 1).isdigit() else '0'
                ancho_max_val = float(ancho_max_val) if ancho_max_val else 0

                # Obtener interlineado, validar que sea número
                interlineado_val = interlineado.get() if interlineado.get().replace('.', '', 1).isdigit() else '0'
                interlineado_val = float(interlineado_val) if interlineado_val else 0

                config[campo.get()] = {
                    'texto_muestra': texto_muestra.get(),
                    'x': int(x.get()) if x.get().isdigit() else 0,
                    'y': int(y.get()) if y.get().isdigit() else 0,
                    'fuente': fuente.get(),
                    'tamaño': int(tamano.get()),
                    'alineacion': alineacion.get(),
                    'negrita': frame.bold_var.get(),
                    'cursiva': frame.italic_var.get(),
                    'ancho_maximo': ancho_max_val,
                    'interlineado': interlineado_val
                }
        return config

    def generar(self):
        plantilla = self.plantilla_entry.get()
        datos = self.datos_entry.get()
        curso = self.curso_entry.get()
        salida = self.salida_entry.get()
        config = self.obtener_config_campos()

        if not all([plantilla, datos, curso, salida, config]):
            messagebox.showerror("Error", "Todos los campos son obligatorios")
            return

        self.guardar_config()  # Guardar la configuración actual antes de generar
        generar_certificados(plantilla, datos, config, salida, curso)
        #self.actualizar_vista_previa()



    def guardar_config(self):
        config = {
            'plantilla': self.plantilla_entry.get(),
            'datos': self.datos_entry.get(),
            'curso': self.curso_entry.get(),
            'salida': self.salida_entry.get(),
            'campos': self.obtener_config_campos()
        }
        
        # Guardar la última configuración
        with open('ultima_config.json', 'w') as f:
            json.dump(config, f)
        
        # Guardar una copia en la ubicación seleccionada
        directorio_salida = self.salida_entry.get()
        if directorio_salida:
            nombre_archivo = f"config_{self.curso_entry.get().replace(' ', '_')}.json"
            ruta_completa = os.path.join(directorio_salida, nombre_archivo)
            with open(ruta_completa, 'w') as f:
                json.dump(config, f)
            messagebox.showinfo("Éxito", f"Configuración guardada en:\n{ruta_completa}")
        else:
            messagebox.showwarning("Advertencia", "No se ha seleccionado un directorio de salida. Solo se ha guardado la última configuración.")




    def cargar_config(self, filename=None):
        if not filename:
            filename = filedialog.askopenfilename(filetypes=[("JSON files", "*.json")])
        if filename and os.path.exists(filename):
            with open(filename, 'r') as f:
                config = json.load(f)
            
            self.plantilla_entry.delete(0, tk.END)
            self.plantilla_entry.insert(0, config.get('plantilla', ''))
            
            self.datos_entry.delete(0, tk.END)
            self.datos_entry.insert(0, config.get('datos', ''))
            
            self.curso_entry.delete(0, tk.END)
            self.curso_entry.insert(0, config.get('curso', ''))
            
            self.salida_entry.delete(0, tk.END)
            self.salida_entry.insert(0, config.get('salida', ''))
            
            # Limpiar campos existentes
            for widget in self.campos_frame.winfo_children():
                widget.destroy()

            # Cargar configuración de campos
            for campo, valores in config.get('campos', {}).items():
                frame = self.agregar_campo()
                entries = [widget for widget in frame.winfo_children() if isinstance(widget, (ttk.Entry, ttk.Combobox))]

                if len(entries) >= 9:
                    campo_entry, texto_muestra_entry, x_entry, y_entry, fuente_combo, tamano_combo, alineacion_combo, ancho_maximo_entry, interlineado_entry = entries[:9]

                    # Cargar valores usando los StringVar del frame
                    campo_entry.set(campo)  # Combobox usa .set()
                    frame.texto_muestra_var.set(valores.get('texto_muestra', ''))
                    frame.x_var.set(str(valores.get('x', 420)))
                    frame.y_var.set(str(valores.get('y', 270)))
                    fuente_combo.set(valores.get('fuente', 'Times-Roman'))
                    tamano_combo.set(valores.get('tamaño', 20))
                    alineacion_combo.set(valores.get('alineacion', 'Centro'))

                    # Cargar ancho máximo e interlineado usando StringVar
                    frame.ancho_maximo_var.set(str(valores.get('ancho_maximo', 600)))
                    frame.interlineado_var.set(str(valores.get('interlineado', 25)))

                    frame.bold_var.set(valores.get('negrita', True))
                    frame.italic_var.set(valores.get('cursiva', False))



            if filename != 'ultima_config.json':
                messagebox.showinfo("Éxito", f"Configuración cargada desde:\n{filename}")
        self.actualizar_vista_previa()


root = tk.Tk()  # Crear la ventana principal
root.geometry("1366x768")    # Establecer el tamaño de la ventana
root.resizable(False, False)    # La ventana no se puede redimensionar
# root.iconbitmap('Certificado.ico')    # Establecer el icono de la ventana
try:
    root.iconphoto(False, tk.PhotoImage(file='Certificado.png'))    # Establecer el icono de la ventana
except Exception as e:
    logging.warning(f"No se pudo cargar el icono: {e}")
app = Application(master=root) # Crear la aplicación
app.mainloop() # Iniciar la aplicación