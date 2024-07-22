import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd # importa la librería pandas, sirve para leer archivos de Excel
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.pdfmetrics import stringWidth
from reportlab.pdfbase.ttfonts import TTFont
from PyPDF2 import PdfReader, PdfWriter
from pdf2image import convert_from_bytes
from PIL import Image, ImageTk
import io
import os
import logging
import json

# Configurar logging
logging.basicConfig(filename='errores_certificados.log', level=logging.ERROR)

def generar_certificados(plantilla_pdf, datos_excel, configuracion, directorio_salida, nombre_curso):
    # Leer datos del Excel
    df = pd.read_excel(datos_excel) # lee el archivo Excel y lo guarda en la variable df
    
    # Leer la plantilla PDF

    for index, row in df.iterrows(): # itera sobre cada fila del archivo Excel

        try:

            # Crear un nuevo PDF
            output = PdfWriter()
            plantilla = PdfReader(plantilla_pdf)
            packet = io.BytesIO() # crea un archivo temporal en memoria
            can = canvas.Canvas(packet, pagesize=A4) # crea un objeto canvas para dibujar en el PDF, con tamaño A4. 

            page = plantilla.pages[0]
            pagina=page.mediabox
            if pagina.right-pagina.left < pagina.top-pagina.bottom:
                vertical=1
            else:
                vertical=0

            # Añadir texto al PDF
            for campo, config in configuracion.items():
                if campo in row:

                    can.setFont(config['fuente'], config['tamaño'])
                    texto = str(row[campo])
                    ancho_texto = stringWidth(texto, config['fuente'], config['tamaño'])
                    
                    if config['alineacion'] == 'Centro':
                        x = config['x'] - (ancho_texto / 2)
                    elif config['alineacion'] == 'Derecha':
                        x = config['x'] - ancho_texto
                    else:  # 'Izquierda'
                        x = config['x']
                    
                    can.drawString(x, A4[vertical] - config['y'], texto)

      
            can.save()
            
            # Mover al inicio del BytesIO
            packet.seek(0)
            nuevo_pdf = PdfReader(packet)
            
            # Fusionar con la plantilla

            
            page.merge_page(nuevo_pdf.pages[0])
            output.add_page(page)
            
            # Crear el nombre del archivo
            nombre_archivo = f"{nombre_curso}_{row['dni']}.pdf"
            ruta_completa = os.path.join(directorio_salida, nombre_archivo)
            
            # Guardar el nuevo certificado
            with open(ruta_completa, "wb") as output_stream:
                output.write(output_stream)
            
            print(f"Certificado generado: {nombre_archivo}")
        
        except Exception as e:
            logging.error(f"Error al generar certificado para DNI {row['dni']}: {str(e)}")

    messagebox.showinfo("Proceso Completado", "Todos los certificados han sido generados.")

class VistaPreviaAvanzada(ttk.Frame):
    def __init__(self, master, *args, **kwargs):
        super().__init__(master, *args, **kwargs)
        self.canvas = tk.Canvas(self, width=400, height=566)  # Proporción A4
        self.canvas.pack(fill=tk.BOTH, expand=True)

    def actualizar_vista_previa(self, plantilla_pdf, campos, nombre_curso):
        # Generar PDF de vista previa
        packet = io.BytesIO()
        can = canvas.Canvas(packet, pagesize=A4)
        existing_pdf = PdfReader(plantilla_pdf)
        page = existing_pdf.pages[0]
        pagina=page.mediabox
        if pagina.right-pagina.left < pagina.top-pagina.bottom:
            vertical=1
        else:
            vertical=0
                # Dibujar campos
        for campo, config in campos.items():
            can.setFont(config['fuente'], config['tamaño'])
            texto = str(campo)
            ancho_texto = stringWidth(texto, config['fuente'], config['tamaño'])
            if config['alineacion'] == 'Centro':
                x = config['x'] - (ancho_texto / 2)
            elif config['alineacion'] == 'Derecha':
                x = config['x'] - ancho_texto
            else:  # 'Izquierda'
                x = config['x']

            
            can.drawString(x, A4[vertical] - config['y'], texto)
        
        
        can.save()
        
        # Combinar con la plantilla
        packet.seek(0)
        new_pdf = PdfReader(packet)
        
        output = PdfWriter()
        

        page.merge_page(new_pdf.pages[0])
        output.add_page(page)
        
        # Convertir PDF a imagen
        pdf_bytes = io.BytesIO()
        output.write(pdf_bytes)
        pdf_bytes.seek(0)
        
        images = convert_from_bytes(pdf_bytes.getvalue())
        img = images[0]
        
        # Redimensionar imagen
        img.thumbnail((400, 566))  # Mantener proporción A4
        
        # Mostrar en el canvas
        self.photo = ImageTk.PhotoImage(img)
        self.canvas.delete("all")
        self.canvas.create_image(0, 0, anchor="nw", image=self.photo)

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
        print(self.datos_entry.get())
        if self.datos_entry.get() != '':
            filename = self.datos_entry.get()
            self.cargar_columnas_excel(filename)


    def create_widgets(self):
        # Frame principal
        main_frame = ttk.Frame(self, padding="3 3 12 12")
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
        self.vista_previa.grid(column=3, row=0, rowspan=8, padx=10, pady=10, sticky=(tk.N, tk.S, tk.E, tk.W))

        for child in main_frame.winfo_children():
            child.grid_configure(padx=5, pady=5)
        self.curso_entry.bind('<KeyRelease>', lambda e: self.actualizar_vista_previa())



    def actualizar_vista_previa(self):
        plantilla = self.plantilla_entry.get()
        campos = self.obtener_config_campos()
        nombre_curso = self.curso_entry.get()
        if plantilla and campos and nombre_curso:
            self.vista_previa.actualizar_vista_previa(plantilla, campos, nombre_curso)


    def agregar_campo(self):
        frame = ttk.Frame(self.campos_frame)
        frame.pack(fill=tk.X, expand=True, pady=2)

        ttk.Label(frame, text="Campo:").pack(side=tk.LEFT)
        campo = ttk.Combobox(frame, values=self.excel_columns, width=10)
        campo.pack(side=tk.LEFT, padx=2)

        ttk.Label(frame, text="X:").pack(side=tk.LEFT)
        x_var = tk.StringVar(value="200")  # Valor por defecto
        x = ttk.Entry(frame, width=5, textvariable=x_var)
        x.pack(side=tk.LEFT, padx=2)

        ttk.Label(frame, text="Y:").pack(side=tk.LEFT)
        y_var = tk.StringVar(value="200")  # Valor por defecto
        y = ttk.Entry(frame, width=5, textvariable=y_var)
        y.pack(side=tk.LEFT, padx=2)

        ttk.Label(frame, text="Fuente:").pack(side=tk.LEFT)
        fuentes = ["Helvetica", "Times-Roman", "Courier"]
        fuente = ttk.Combobox(frame, values=fuentes, width=10)
        fuente.set(fuentes[0])
        fuente.pack(side=tk.LEFT, padx=2)

        ttk.Label(frame, text="Tamaño:").pack(side=tk.LEFT)
        tamanos = list(range(8, 73, 2))
        tamano = ttk.Combobox(frame, values=tamanos, width=5)
        tamano.set(12)
        tamano.pack(side=tk.LEFT, padx=2)

        ttk.Label(frame, text="Alineación:").pack(side=tk.LEFT)
        alineaciones = ["Izquierda", "Centro", "Derecha"]
        alineacion = ttk.Combobox(frame, values=alineaciones, width=8)
        alineacion.set("Centro")
        alineacion.pack(side=tk.LEFT, padx=2)

        ttk.Button(frame, text="Eliminar", command=lambda: frame.destroy()).pack(side=tk.RIGHT)

        # Añadir callback para actualizar la vista previa
        campo.bind('<KeyRelease>', lambda e: self.actualizar_vista_previa())
        x.bind('<KeyRelease>', lambda e: self.actualizar_vista_previa())
        y.bind('<KeyRelease>', lambda e: self.actualizar_vista_previa())
        fuente.bind('<<ComboboxSelected>>', lambda e: self.actualizar_vista_previa())
        tamano.bind('<<ComboboxSelected>>', lambda e: self.actualizar_vista_previa())
        alineacion.bind('<<ComboboxSelected>>', lambda e: self.actualizar_vista_previa())




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
            print(f"Columnas del Excel: {self.excel_columns}")
            for frame in self.campos_frame.winfo_children():
                campo = frame.winfo_children()[1]  # El combobox de campo está en la segunda posición
                campo['values'] = self.excel_columns
        except Exception as e:
            print(f"Error al cargar columnas del Excel: {str(e)}")
            messagebox.showerror("Error", f"Error al cargar columnas del archivo Excel: {str(e)}")


    def buscar_salida(self):
        directory = filedialog.askdirectory()
        self.salida_entry.delete(0, tk.END)
        self.salida_entry.insert(0, directory)

    def obtener_config_campos(self):
        config = {}
        for frame in self.campos_frame.winfo_children():
            entries = [widget for widget in frame.winfo_children() if isinstance(widget, (ttk.Entry, ttk.Combobox))]
            if len(entries) == 6:
                campo, x, y, fuente, tamano, alineacion = entries
                config[campo.get()] = {
                    'x': int(x.get()),
                    'y': int(y.get()),
                    'fuente': fuente.get(),
                    'tamaño': int(tamano.get()),
                    'alineacion': alineacion.get()
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
        self.actualizar_vista_previa()



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
                self.agregar_campo()
                frame = self.campos_frame.winfo_children()[-1]
                entries = [widget for widget in frame.winfo_children() if isinstance(widget, (ttk.Entry, ttk.Combobox))]
                entries[0].insert(0, campo)
                entries[1].insert(0, valores['x'])
                entries[2].insert(0, valores['y'])
                entries[3].set(valores['fuente'])
                entries[4].set(valores['tamaño'])

            if filename != 'ultima_config.json':
                messagebox.showinfo("Éxito", f"Configuración cargada desde:\n{filename}")
        self.actualizar_vista_previa()


root = tk.Tk()  # Crear la ventana principal
app = Application(master=root) # Crear la aplicación
app.mainloop() # Iniciar la aplicación