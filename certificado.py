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
logging.basicConfig(filename='errores_certificados.log', level=logging.ERROR)

def generar_certificados(plantilla_pdf, datos_excel, configuracion, directorio_salida, nombre_curso):
    df = pd.read_excel(datos_excel)

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

                    ancho_texto = stringWidth(text, config['fuente'], config['tamaño'])
                        
                    if config['alineacion'] == 'Centro':
                            x = config['x'] - (ancho_texto / 2)
                    elif config['alineacion'] == 'Derecha':
                            x = config['x'] - ancho_texto
                    else:  # 'Izquierda'
                            x = config['x']
                        
                        #can.drawString(x, A4[vertical] - config['y'], texto)

                    # PyMuPDF uses top-left origin, so we need to adjust y-coordinate
                                    #y = page.rect.height - config['y']
                    y = config['y']

                    # Add text to the page
                    output_page.insert_text((x, y), text, fontname=font_name, fontsize=font_size)


            # Save the certificate
            
            output_filename = f"{nombre_curso}_{row['dni']}.pdf"
            output_path = os.path.join(directorio_salida, output_filename)
            output_pdf.save(output_path)
            output_pdf.close()

        except Exception as e:
            logging.error(f"Error al generar certificado para DNI {row['dni']}: {str(e)}")

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
    def __init__(self, master, *args, **kwargs):
        super().__init__(master, *args, **kwargs)
        self.canvas = tk.Canvas(self, width=600, height=600)  # Proporción A4
        self.canvas.pack(fill=tk.BOTH, expand=True)

    def actualizar_vista_previa(self, plantilla_pdf, campos, nombre_curso):
        try:
            # Open the PDF
            doc = fitz.open(plantilla_pdf)
            page = doc[0]  # Get the first page

            # Create a new PDF with the fields
            output_buffer = io.BytesIO()
            output_pdf = fitz.open()
            output_page = output_pdf.new_page(width=page.rect.width, height=page.rect.height)

            # Copy the content of the template to the new page
            output_page.show_pdf_page(page.rect, doc, 0)

            # Add text fields
            for campo, config in campos.items():
                text = config.get('texto_muestra', campo)
                font_name = config['fuente']
                font_size = config['tamaño']
                is_bold = config.get('negrita', False)
                is_italic = config.get('cursiva', False)
                    
                # PyMuPDF uses different font names
                
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
                            
                ancho_texto = stringWidth(text, config['fuente'], config['tamaño'])
                    
                if config['alineacion'] == 'Centro':
                        x = config['x'] - (ancho_texto / 2)
                elif config['alineacion'] == 'Derecha':
                        x = config['x'] - ancho_texto
                else:  # 'Izquierda'
                        x = config['x']
                    
                    #can.drawString(x, A4[vertical] - config['y'], texto)

                # PyMuPDF uses top-left origin, so we need to adjust y-coordinate
                                #y = page.rect.height - config['y']
                y = config['y']
                

                # Add text to the page
                output_page.insert_text((x, y), text, fontname=font_name, fontsize=font_size)

            # Save the preview PDF
            output_pdf.save(output_buffer)
            output_buffer.seek(0)

            # Convert PDF to image
            pix = output_page.get_pixmap(matrix=fitz.Matrix(2, 2))  # Increase resolution
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

            # Resize image to fit canvas
            img.thumbnail((600, 600))  # Maintain aspect ratio

            # Display in canvas
            self.photo = ImageTk.PhotoImage(img)
            self.canvas.delete("all")
            self.canvas.create_image(0, 0, anchor="nw", image=self.photo)

        except Exception as e:
            print(f"Error updating preview: {e}")

class Application(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        self.master.title("Generador de Certificados")
        self.pack(fill=tk.BOTH, expand=True)
        self.excel_columns = []
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
        
        
        self.vista_previa = VistaPreviaAvanzada(main_frame)
        self.vista_previa.grid(column=3, row=0, rowspan=15, padx=10, pady=10, sticky=(tk.N, tk.S, tk.E, tk.W))

        for child in main_frame.winfo_children():
            child.grid_configure(padx=5, pady=5) # Add padding to all widgets
        self.curso_entry.bind('<KeyRelease>', lambda e: self.actualizar_vista_previa())


    @debounce(1)  # Debounce with a 0.5 second delay
    def actualizar_vista_previa(self):
        plantilla = self.plantilla_entry.get()
        campos = self.obtener_config_campos()
        nombre_curso = self.curso_entry.get()
        if plantilla and campos and nombre_curso:
            try:
                plantilla = self.plantilla_entry.get()
                campos = self.obtener_config_campos()
                nombre_curso = self.curso_entry.get()
                if plantilla and campos and nombre_curso:
                    self.vista_previa.actualizar_vista_previa(plantilla, campos, nombre_curso)
                    
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
        texto_muestra = ttk.Entry(frame, width=15)
        texto_muestra.grid(row=0, column=3, padx=4,pady=4, sticky='w')

        ttk.Label(frame, text="X:").grid(row=0, column=4, sticky='w')
        x_var = tk.StringVar(value="200")
        x = ttk.Entry(frame, width=5, textvariable=x_var)
        x.grid(row=0, column=5, padx=4, sticky='w')

        ttk.Label(frame, text="Y:").grid(row=0, column=6, sticky='w')
        y_var = tk.StringVar(value="200")
        y = ttk.Entry(frame, width=5, textvariable=y_var)
        y.grid(row=0, column=7, padx=4, sticky='w')

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
        bold_var = tk.BooleanVar()
        bold_check = ttk.Checkbutton(frame, text="Negrita", variable=bold_var)
        bold_check.grid(row=1, column=5, padx=4, sticky='w')
        italic_var = tk.BooleanVar()
        italic_check = ttk.Checkbutton(frame, text="Cursiva", variable=italic_var)
        italic_check.grid(row=1, column=6, padx=4, sticky='w')
        frame.bold_var = bold_var
        frame.italic_var = italic_var

        ttk.Button(frame, text="Eliminar", command=lambda: frame.destroy()).grid(row=2, column=4, padx=2, sticky='e')
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
            df = pd.read_excel(filename)
            self.excel_columns = list(df.columns)
            for frame in self.campos_frame.winfo_children():
                campo = frame.winfo_children()[1]  # El combobox de campo está en la segunda posición
                campo['values'] = self.excel_columns
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
            if len(entries) >= 7:  # Ahora tenemos un campo más
                campo, texto_muestra, x, y, fuente, tamano, alineacion = entries[:7]
                
                config[campo.get()] = {
                    'texto_muestra': texto_muestra.get(),
                    'x': int(x.get()) if x.get().isdigit() else 0,
                    'y': int(y.get()) if y.get().isdigit() else 0,
                    'fuente': fuente.get(),
                    'tamaño': int(tamano.get()),
                    'alineacion': alineacion.get(),
                    'negrita': frame.bold_var.get(),
                    'cursiva': frame.italic_var.get()
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
                
                campo_entry, texto_muestra_entry, x_entry, y_entry, fuente_combo, tamano_combo, alineacion_combo = entries[:7]
                
                campo_entry.insert(0, campo)
                texto_muestra_entry.insert(0, valores.get('texto_muestra', ''))
                x_entry.insert(0, valores['x'])
                y_entry.insert(0, valores['y'])
                fuente_combo.set(valores['fuente'])
                tamano_combo.set(valores['tamaño'])
                alineacion_combo.set(valores['alineacion'])
                
                frame.bold_var.set(valores.get('negrita', False))
                frame.italic_var.set(valores.get('cursiva', False))



            if filename != 'ultima_config.json':
                messagebox.showinfo("Éxito", f"Configuración cargada desde:\n{filename}")
        self.actualizar_vista_previa()


root = tk.Tk()  # Crear la ventana principal
root.geometry("1366x768")    # Establecer el tamaño de la ventana
root.resizable(False, False)    # La ventana no se puede redimensionar
# root.iconbitmap('Certificado.ico')    # Establecer el icono de la ventana
root.iconphoto(False, tk.PhotoImage(file='Certificado.png'))    # Establecer el icono de la ventana
app = Application(master=root) # Crear la aplicación
app.mainloop() # Iniciar la aplicación