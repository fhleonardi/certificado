"""
Generador de Certificados en PDF.

Genera un certificado por fila de un Excel, escribiendo los campos elegidos
sobre una plantilla PDF. La interfaz usa CustomTkinter.

Entorno (crearlo FUERA de OneDrive: OneDrive libera archivos del venv y rompe pip):
    python -m venv C:\\venvs\\certificado
    C:\\venvs\\certificado\\Scripts\\activate
    pip install pandas openpyxl reportlab pillow PyMuPDF customtkinter pyinstaller

Ejecutable:
    pyinstaller --onefile -w certificado.py --icon=certificado.png --add-data "certificado.png:."
"""
import json
import logging
import os
import queue
import re
import sys
import tempfile
import threading
import tkinter as tk
from tkinter import filedialog, messagebox

import customtkinter as ctk
import pandas as pd
import pymupdf
from PIL import Image, ImageTk
from reportlab.pdfbase.pdfmetrics import stringWidth

logging.basicConfig(filename='errores_certificados.log', level=logging.WARNING)

CONFIG_ULTIMA = 'ultima_config.json'  # Se lee y escribe en la carpeta de trabajo (al lado del .exe)
ICONO = 'certificado.png'  # Recurso de solo lectura: va empaquetado dentro del .exe


def ruta_recurso(nombre):
    """Ruta a un recurso empaquetado: dentro del .exe usa la carpeta temporal de PyInstaller,
    en desarrollo la carpeta del script."""
    base = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, nombre)

FUENTES = ["Helvetica", "Times-Roman", "Courier"]
TAMANOS = [str(t) for t in range(8, 73, 2)]
ALINEACIONES = ["Izquierda", "Centro", "Derecha"]

CAMPO_POR_DEFECTO = {
    'nombre': '',
    'texto_muestra': 'Texto de ejemplo',
    'x': 420,  # Centro horizontal de un A4 apaisado (842 pt)
    'y': 270,
    'fuente': 'Times-Roman',
    'tamaño': 20,
    'alineacion': 'Centro',
    'negrita': True,
    'cursiva': False,
    'ancho_maximo': 600,  # 0 = sin límite
    'interlineado': 25,   # 0 = automático (1.2 × tamaño)
}

# Colores (claro, oscuro)
COLOR_CAMPO = '#3b82f6'
COLOR_SELECCION = '#22c55e'
FONDO_CANVAS = ('#d4d4d8', '#18181b')
SOMBRA_PAGINA = ('#a1a1aa', '#09090b')
BORDE_TARJETA = ('gray80', 'gray28')
COLOR_ESTADO = {
    'info': ('gray30', 'gray70'),
    'ok': ('#15803d', '#4ade80'),
    'error': ('#b91c1c', '#f87171'),
}

# Variantes por fuente: (normal, negrita, cursiva, negrita+cursiva).
# Cada una es (nombre para PyMuPDF, nombre para medir con reportlab).
# Medir con la variante correcta importa: la negrita es ~5% más ancha.
_VARIANTES_FUENTE = {
    "Times-Roman": (("tiro", "Times-Roman"), ("tibo", "Times-Bold"),
                    ("tiit", "Times-Italic"), ("tibi", "Times-BoldItalic")),
    "Helvetica": (("helv", "Helvetica"), ("hebo", "Helvetica-Bold"),
                  ("heit", "Helvetica-Oblique"), ("hebi", "Helvetica-BoldOblique")),
    "Courier": (("cour", "Courier"), ("cobo", "Courier-Bold"),
                ("coit", "Courier-Oblique"), ("cobi", "Courier-BoldOblique")),
}


# ---------------------------------------------------------------------------
# Lógica de PDF (sin interfaz)
# ---------------------------------------------------------------------------

def resolver_fuente(fuente, negrita, cursiva):
    """Devuelve (nombre PyMuPDF, nombre reportlab) para la fuente y el estilo."""
    variantes = _VARIANTES_FUENTE.get(fuente, _VARIANTES_FUENTE["Times-Roman"])
    return variantes[int(bool(negrita)) + 2 * int(bool(cursiva))]


def valor_texto(valor):
    """Convierte un valor del Excel a texto: vacío para celdas vacías y sin '.0' en enteros."""
    if pd.isna(valor):
        return ''
    if isinstance(valor, float) and valor.is_integer():
        return str(int(valor))
    return str(valor)


def nombre_archivo_seguro(texto):
    """Reemplaza los caracteres que Windows no admite en nombres de archivo."""
    return re.sub(r'[\\/:*?"<>|]+', '_', str(texto)).strip() or 'sin_nombre'


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
        if stringWidth(linea_prueba, fuente, tamaño) <= ancho_maximo:
            linea_actual.append(palabra)
        elif linea_actual:
            lineas.append(' '.join(linea_actual))
            linea_actual = [palabra]
        else:
            # Si una sola palabra es más ancha que el máximo, agregarla de todos modos
            lineas.append(palabra)

    if linea_actual:
        lineas.append(' '.join(linea_actual))

    return lineas if lineas else [texto]


def x_inicio(alineacion, x_ref, ancho):
    """X donde empieza un bloque de ese ancho, según su punto de referencia y alineación."""
    if alineacion == 'Centro':
        return x_ref - ancho / 2
    if alineacion == 'Derecha':
        return x_ref - ancho
    return x_ref


def maquetar_campo(texto, config):
    """Calcula las líneas de un campo y la posición de cada una (Y = línea base)."""
    fuente_pdf, fuente_rl = resolver_fuente(config['fuente'], config.get('negrita'), config.get('cursiva'))
    tamaño = config['tamaño']
    lineas = dividir_texto_en_lineas(texto, fuente_rl, tamaño, config.get('ancho_maximo', 0))
    interlineado = config.get('interlineado') or tamaño * 1.2  # 0 o ausente = automático

    posiciones = []
    for i, linea in enumerate(lineas):
        ancho = stringWidth(linea, fuente_rl, tamaño)
        x = x_inicio(config['alineacion'], config['x'], ancho)
        posiciones.append((x, config['y'] + i * interlineado, linea))

    return {
        'fuente_pdf': fuente_pdf,
        'fuente_rl': fuente_rl,
        'lineas': lineas,
        'interlineado': interlineado,
        'posiciones': posiciones,
    }


def dibujar_campo(page, maquetado, tamaño):
    for x, y, linea in maquetado['posiciones']:
        page.insert_text((x, y), linea, fontname=maquetado['fuente_pdf'], fontsize=tamaño)


def generar_certificados(plantilla_pdf, df, configuracion, directorio_salida, nombre_curso,
                         columna_indice, on_progreso=None):
    """
    Genera un PDF por fila del DataFrame. No toca la interfaz, así que puede
    correr en un hilo aparte. Devuelve (cantidad_generados, lista_de_errores).
    """
    generados, errores = 0, []
    total = len(df)
    prefijo = nombre_archivo_seguro(nombre_curso)

    with pymupdf.open(plantilla_pdf) as plantilla:  # Se abre una sola vez, no por fila
        rect = plantilla[0].rect

        for n, (_, row) in enumerate(df.iterrows(), start=1):
            identificador = valor_texto(row[columna_indice]) if columna_indice in row else str(n)
            try:
                salida = pymupdf.open()
                page = salida.new_page(width=rect.width, height=rect.height)
                page.show_pdf_page(rect, plantilla, 0)

                for campo, config in configuracion.items():
                    if campo in row:
                        dibujar_campo(page, maquetar_campo(valor_texto(row[campo]), config), config['tamaño'])

                nombre = f"{prefijo}_{nombre_archivo_seguro(identificador)}.pdf"
                salida.save(os.path.join(directorio_salida, nombre))
                salida.close()
                generados += 1
            except Exception as e:
                errores.append(f"{columna_indice} {identificador}: {e}")
                logging.error(f"Error al generar certificado para {columna_indice} {identificador}: {e}")

            if on_progreso:
                on_progreso(n, total)

    return generados, errores


def _numero(texto, por_defecto=0.0):
    try:
        return float(str(texto).replace(',', '.'))
    except ValueError:
        return por_defecto


def _formatear(valor):
    """600.0 -> '600'; 12.5 -> '12.5'."""
    numero = _numero(valor)
    return str(int(numero)) if numero.is_integer() else str(numero)


# ---------------------------------------------------------------------------
# Interfaz
# ---------------------------------------------------------------------------

class VistaPrevia(ctk.CTkFrame):
    """
    Muestra la plantilla con los campos dibujados y permite moverlos,
    cambiarles el ancho y ajustarlos con las flechas.
    No conoce a la aplicación: avisa los cambios mediante callbacks.
    """
    HANDLE = 6  # Medio lado del tirador de redimensionado, en píxeles

    def __init__(self, master, on_mover, on_redimensionar, on_seleccionar, **kwargs):
        super().__init__(master, **kwargs)
        self.on_mover = on_mover
        self.on_redimensionar = on_redimensionar
        self.on_seleccionar = on_seleccionar

        self.canvas = tk.Canvas(self, highlightthickness=0, bd=0)
        self.canvas.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        self._aplicar_fondo()

        self.photo = None
        self.pdf_w = self.pdf_h = 0
        self.scale = 1.0
        self.offset_x = self.offset_y = 0
        self.campos_rect = {}
        self.seleccionado = None
        self.drag = None
        self.bloqueada = False  # True mientras se generan certificados (PyMuPDF no es thread-safe)
        self._args_render = None
        self._mensaje = None
        self._resize_after = None
        self._tamano_canvas = (0, 0)

        self.canvas.bind("<Button-1>", self._on_click)
        self.canvas.bind("<B1-Motion>", self._on_drag)
        self.canvas.bind("<ButtonRelease-1>", self._on_release)
        self.canvas.bind("<Motion>", self._on_motion)
        self.canvas.bind("<Configure>", self._on_configure)
        for tecla, (dx, dy) in {"<Left>": (-1, 0), "<Right>": (1, 0), "<Up>": (0, -1), "<Down>": (0, 1)}.items():
            self.canvas.bind(tecla, lambda e, dx=dx, dy=dy: self._on_tecla(e, dx, dy))

    # --- Tema -------------------------------------------------------------

    def _set_appearance_mode(self, mode_string):
        super()._set_appearance_mode(mode_string)
        self._aplicar_fondo()

    def _aplicar_fondo(self):
        if hasattr(self, 'canvas'):
            self.canvas.configure(bg=self._apply_appearance_mode(FONDO_CANVAS))

    # --- Render -------------------------------------------------------------

    def mostrar_mensaje(self, texto):
        self._args_render = None
        self._mensaje = texto
        self.campos_rect = {}
        self.canvas.delete("all")
        w, h = self.canvas.winfo_width(), self.canvas.winfo_height()
        self.canvas.create_text(w / 2, h / 2, text=texto, fill='gray55', font=('Segoe UI', 12), justify='center')

    def renderizar(self, plantilla_pdf, campos, datos_registro=None):
        if self.bloqueada:
            return
        self._args_render = (plantilla_pdf, campos, datos_registro)
        self._mensaje = None

        cw = max(self.canvas.winfo_width(), 100)
        ch = max(self.canvas.winfo_height(), 100)
        margen = 24

        with pymupdf.open(plantilla_pdf) as doc:
            rect = doc[0].rect
            self.pdf_w, self.pdf_h = rect.width, rect.height
            self.scale = max(min((cw - margen) / self.pdf_w, (ch - margen) / self.pdf_h), 0.05)

            salida = pymupdf.open()
            page = salida.new_page(width=self.pdf_w, height=self.pdf_h)
            page.show_pdf_page(rect, doc, 0)

            self.campos_rect = {}
            for campo, config in campos.items():
                if datos_registro is not None and campo in datos_registro:
                    texto = valor_texto(datos_registro[campo])
                else:
                    texto = config.get('texto_muestra') or campo
                maquetado = maquetar_campo(texto, config)
                dibujar_campo(page, maquetado, config['tamaño'])
                self.campos_rect[campo] = self._rect_de(config, maquetado)

            # Renderizar directo a la escala final: más nítido y rápido que achicar después
            pix = page.get_pixmap(matrix=pymupdf.Matrix(self.scale, self.scale))
            salida.close()

        img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
        self.offset_x = (cw - pix.width) // 2
        self.offset_y = (ch - pix.height) // 2
        self.photo = ImageTk.PhotoImage(img)

        if self.seleccionado not in self.campos_rect:
            self.seleccionado = None

        self.canvas.delete("all")
        self.canvas.create_rectangle(self.offset_x + 3, self.offset_y + 3,
                                     self.offset_x + pix.width + 3, self.offset_y + pix.height + 3,
                                     fill=self._apply_appearance_mode(SOMBRA_PAGINA), outline='')
        self.canvas.create_image(self.offset_x, self.offset_y, anchor="nw", image=self.photo)
        self.dibujar_recuadros()

    @staticmethod
    def _rect_de(config, maquetado):
        tamaño = config['tamaño']
        lineas = maquetado['lineas']
        ancho_texto = max((stringWidth(l, maquetado['fuente_rl'], tamaño) for l in lineas), default=0)
        return {
            'x_ref': config['x'],
            'y_base': config['y'],
            'ancho': config.get('ancho_maximo') or ancho_texto,
            'altura': (len(lineas) - 1) * maquetado['interlineado'] + tamaño,
            'tamaño': tamaño,
            'alineacion': config['alineacion'],
        }

    def dibujar_recuadros(self):
        self.canvas.delete("overlay")
        s, h = self.scale, self.HANDLE
        for campo, r in self.campos_rect.items():
            # PyMuPDF usa Y como línea base; el ascendente es ~0.75 del tamaño
            x1 = x_inicio(r['alineacion'], r['x_ref'], r['ancho']) * s + self.offset_x
            y1 = (r['y_base'] - 0.75 * r['tamaño']) * s + self.offset_y
            x2 = x1 + r['ancho'] * s
            y2 = y1 + r['altura'] * s
            r['bbox'] = (x1, y1, x2, y2)
            # El tirador va en el borde que se puede mover según la alineación
            r['handle'] = (x1 if r['alineacion'] == 'Derecha' else x2, y2)

            seleccionado = campo == self.seleccionado
            color = COLOR_SELECCION if seleccionado else COLOR_CAMPO
            self.canvas.create_rectangle(x1, y1, x2, y2, outline=color, width=2 if seleccionado else 1,
                                         dash=() if seleccionado else (4, 2), tags="overlay")
            self.canvas.create_text(x1, y1 - 2, anchor='sw', text=campo, fill=color,
                                    font=('Segoe UI', 9, 'bold'), tags="overlay")
            hx, hy = r['handle']
            self.canvas.create_rectangle(hx - h, hy - h, hx + h, hy + h, fill=color, outline='white', tags="overlay")

    # --- Interacción ------------------------------------------------------

    def _hit(self, x, y):
        """Devuelve (campo, 'resize'|'move') bajo el cursor. Prioriza el campo seleccionado."""
        orden = sorted(self.campos_rect, key=lambda c: c != self.seleccionado)
        tolerancia = self.HANDLE + 2
        for campo in orden:
            hx, hy = self.campos_rect[campo].get('handle', (None, None))
            if hx is not None and abs(x - hx) <= tolerancia and abs(y - hy) <= tolerancia:
                return campo, 'resize'
        for campo in orden:
            bbox = self.campos_rect[campo].get('bbox')
            if bbox and bbox[0] - 3 <= x <= bbox[2] + 3 and bbox[1] - 3 <= y <= bbox[3] + 3:
                return campo, 'move'
        return None, None

    def _seleccionar(self, campo):
        self.seleccionado = campo
        self.on_seleccionar(campo)
        self.dibujar_recuadros()

    def _on_click(self, event):
        self.canvas.focus_set()  # Para recibir las flechas
        campo, modo = self._hit(event.x, event.y)
        self.drag = {'campo': campo, 'modo': modo, 'x': event.x, 'y': event.y, 'cambio': False} if campo else None
        self._seleccionar(campo)

    def _on_drag(self, event):
        if not self.drag:
            return
        r = self.campos_rect[self.drag['campo']]
        dx = (event.x - self.drag['x']) / self.scale
        dy = (event.y - self.drag['y']) / self.scale

        if self.drag['modo'] == 'move':
            r['x_ref'] += dx
            r['y_base'] += dy
        else:
            # Centro crece hacia ambos lados; Derecha tiene el tirador a la izquierda
            factor = {'Centro': 2, 'Derecha': -1}.get(r['alineacion'], 1)
            r['ancho'] = max(10, r['ancho'] + factor * dx)

        self.drag.update(x=event.x, y=event.y, cambio=True)
        self.dibujar_recuadros()

    def _on_release(self, event):
        if self.drag and self.drag['cambio']:
            campo = self.drag['campo']
            r = self.campos_rect[campo]
            if self.drag['modo'] == 'move':
                self.on_mover(campo, r['x_ref'], r['y_base'])
            else:
                self.on_redimensionar(campo, r['ancho'])
        self.drag = None

    def _on_motion(self, event):
        if self.drag:
            return
        _, modo = self._hit(event.x, event.y)
        self.canvas.configure(cursor={'resize': 'sb_h_double_arrow', 'move': 'fleur'}.get(modo, ''))

    def _on_tecla(self, event, dx, dy):
        if not self.seleccionado:
            return None
        paso = 10 if event.state & 0x0001 else 1  # Shift = pasos de 10 pt
        r = self.campos_rect[self.seleccionado]
        r['x_ref'] += dx * paso
        r['y_base'] += dy * paso
        self.dibujar_recuadros()
        self.on_mover(self.seleccionado, r['x_ref'], r['y_base'])
        return "break"

    def _on_configure(self, event):
        tamano = (event.width, event.height)
        if tamano == self._tamano_canvas:
            return
        self._tamano_canvas = tamano
        if self._resize_after:
            self.after_cancel(self._resize_after)
        self._resize_after = self.after(150, self._rerender)

    def _rerender(self):
        self._resize_after = None
        if self._args_render:
            try:
                self.renderizar(*self._args_render)
            except Exception as e:
                logging.error(f"Error al redibujar la vista previa: {e}")
        elif self._mensaje:
            self.mostrar_mensaje(self._mensaje)


class CampoCard(ctk.CTkFrame):
    """Tarjeta con la configuración de un campo. Guarda sus propias variables."""

    def __init__(self, master, columnas, on_cambio, on_eliminar, valores=None):
        super().__init__(master, corner_radius=10, border_width=2, border_color=BORDE_TARJETA)
        v = {**CAMPO_POR_DEFECTO, **(valores or {})}

        self.var_nombre = tk.StringVar(value=v['nombre'])
        self.var_muestra = tk.StringVar(value=v['texto_muestra'])
        self.var_x = tk.StringVar(value=_formatear(v['x']))
        self.var_y = tk.StringVar(value=_formatear(v['y']))
        self.var_fuente = tk.StringVar(value=v['fuente'])
        self.var_tamano = tk.StringVar(value=_formatear(v['tamaño']))
        self.var_alineacion = tk.StringVar(value=v['alineacion'])
        self.var_negrita = tk.BooleanVar(value=v['negrita'])
        self.var_cursiva = tk.BooleanVar(value=v['cursiva'])
        self.var_ancho = tk.StringVar(value=_formatear(v['ancho_maximo']))
        self.var_interlineado = tk.StringVar(value=_formatear(v['interlineado']))

        self.grid_columnconfigure((0, 1, 2, 3), weight=1, uniform='col')
        pad = {'padx': 6, 'pady': 3}
        chico = ctk.CTkFont(size=11)

        # Fila 0: columna del Excel + eliminar
        self.combo_nombre = ctk.CTkComboBox(self, values=columnas, variable=self.var_nombre,
                                            font=ctk.CTkFont(weight='bold'))
        self.combo_nombre.grid(row=0, column=0, columnspan=3, sticky='ew', padx=(10, 6), pady=(10, 3))
        ctk.CTkButton(self, text="✕", width=30, fg_color='transparent', text_color=('gray30', 'gray70'),
                      hover_color=('#fecaca', '#7f1d1d'), command=lambda: on_eliminar(self)
                      ).grid(row=0, column=3, sticky='e', padx=(6, 10), pady=(10, 3))

        # Fila 1: texto de muestra
        ctk.CTkLabel(self, text="Texto de muestra (si la columna no está en el Excel)", font=chico,
                     text_color=('gray40', 'gray60')).grid(row=1, column=0, columnspan=4, sticky='w', padx=10)
        ctk.CTkEntry(self, textvariable=self.var_muestra).grid(row=2, column=0, columnspan=4, sticky='ew',
                                                               padx=10, pady=(0, 3))

        # Fila 2: fuente, tamaño, estilo
        ctk.CTkOptionMenu(self, values=FUENTES, variable=self.var_fuente).grid(
            row=3, column=0, columnspan=2, sticky='ew', padx=(10, 6), pady=3)
        ctk.CTkComboBox(self, values=TAMANOS, variable=self.var_tamano).grid(row=3, column=2, sticky='ew', **pad)
        estilo = ctk.CTkFrame(self, fg_color='transparent')
        estilo.grid(row=3, column=3, sticky='ew', padx=(6, 10), pady=3)
        ctk.CTkCheckBox(estilo, text="N", width=20, checkbox_width=20, checkbox_height=20,
                        font=ctk.CTkFont(weight='bold'), variable=self.var_negrita,
                        onvalue=True, offvalue=False).pack(side='left', padx=(0, 6))
        ctk.CTkCheckBox(estilo, text="C", width=20, checkbox_width=20, checkbox_height=20,
                        font=ctk.CTkFont(slant='italic'), variable=self.var_cursiva,
                        onvalue=True, offvalue=False).pack(side='left')

        # Fila 3: alineación
        ctk.CTkSegmentedButton(self, values=ALINEACIONES, variable=self.var_alineacion).grid(
            row=4, column=0, columnspan=4, sticky='ew', padx=10, pady=3)

        # Fila 4: posición y tamaño del bloque
        etiquetas = ["X", "Y", "Ancho (0 = libre)", "Interl. (0 = auto)"]
        variables = [self.var_x, self.var_y, self.var_ancho, self.var_interlineado]
        for col, (texto, var) in enumerate(zip(etiquetas, variables)):
            px = (10 if col == 0 else 6, 10 if col == 3 else 6)
            ctk.CTkLabel(self, text=texto, font=chico, text_color=('gray40', 'gray60')).grid(
                row=5, column=col, sticky='w', padx=px)
            ctk.CTkEntry(self, textvariable=var, justify='center').grid(
                row=6, column=col, sticky='ew', padx=px, pady=(0, 10))

        for var in (self.var_nombre, self.var_muestra, self.var_x, self.var_y, self.var_fuente,
                    self.var_tamano, self.var_alineacion, self.var_negrita, self.var_cursiva,
                    self.var_ancho, self.var_interlineado):
            var.trace_add('write', lambda *_: on_cambio())

    def nombre(self):
        return self.var_nombre.get().strip()

    def get_config(self):
        return {
            'texto_muestra': self.var_muestra.get(),
            'x': int(round(_numero(self.var_x.get()))),
            'y': int(round(_numero(self.var_y.get()))),
            'fuente': self.var_fuente.get(),
            'tamaño': max(1, int(round(_numero(self.var_tamano.get(), 20)))),
            'alineacion': self.var_alineacion.get(),
            'negrita': self.var_negrita.get(),
            'cursiva': self.var_cursiva.get(),
            'ancho_maximo': max(0.0, _numero(self.var_ancho.get())),
            'interlineado': max(0.0, _numero(self.var_interlineado.get())),
        }

    def set_columnas(self, columnas):
        self.combo_nombre.configure(values=columnas)

    def set_posicion(self, x, y):
        self.var_x.set(str(int(round(x))))
        self.var_y.set(str(int(round(y))))

    def set_ancho(self, ancho):
        self.var_ancho.set(str(int(round(ancho))))

    def set_seleccionado(self, seleccionado):
        self.configure(border_color=COLOR_SELECCION if seleccionado else BORDE_TARJETA)


class Application(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Generador de Certificados")
        self.geometry("1366x768")
        self.minsize(1100, 650)
        self._configurar_icono()

        self.df_excel = None
        self.excel_columns = []
        self.registro_actual = 0
        self.cards = []
        self._preview_after = None
        self._generando = False
        self._cola = queue.Queue()

        self.var_plantilla = tk.StringVar()
        self.var_datos = tk.StringVar()
        self.var_curso = tk.StringVar()
        self.var_salida = tk.StringVar()
        self.var_plantilla.trace_add('write', lambda *_: self.programar_vista_previa())

        self._crear_widgets()
        self.cargar_config(CONFIG_ULTIMA, silencioso=True)
        if not self.cards:
            self.agregar_campo()

    def _configurar_icono(self):
        # CustomTkinter pone su propio ícono con iconbitmap salvo que ya se haya llamado,
        # por eso se usa iconbitmap con un .ico generado desde el .png (iconphoto no alcanza).
        try:
            ico = os.path.join(tempfile.gettempdir(), 'certificado_icono.ico')
            Image.open(ruta_recurso(ICONO)).save(ico, sizes=[(16, 16), (32, 32), (48, 48), (64, 64)])
            self.iconbitmap(ico)
        except Exception as e:
            logging.warning(f"No se pudo cargar el icono: {e}")

    # --- Construcción de la interfaz ---------------------------------------

    def _crear_widgets(self):
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self._crear_panel_lateral()
        self._crear_panel_vista_previa()

    def _boton_secundario(self, master, **kwargs):
        return ctk.CTkButton(master, fg_color='transparent', border_width=1,
                             border_color=('gray60', 'gray40'), text_color=('gray10', 'gray90'),
                             hover_color=('gray85', 'gray25'), **kwargs)

    def _crear_panel_lateral(self):
        lateral = ctk.CTkFrame(self, corner_radius=0)
        lateral.grid(row=0, column=0, sticky='nsew')
        lateral.grid_columnconfigure(0, weight=1)
        lateral.grid_rowconfigure(3, weight=1)

        # Encabezado
        encabezado = ctk.CTkFrame(lateral, fg_color='transparent')
        encabezado.grid(row=0, column=0, sticky='ew', padx=16, pady=(16, 8))
        encabezado.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(encabezado, text="Generador de Certificados",
                     font=ctk.CTkFont(size=20, weight='bold')).grid(row=0, column=0, sticky='w')
        modos = {"Sistema": "system", "Claro": "light", "Oscuro": "dark"}
        tema = ctk.CTkOptionMenu(encabezado, values=list(modos), width=100,
                                 command=lambda m: ctk.set_appearance_mode(modos[m]))
        tema.set("Sistema")
        tema.grid(row=0, column=1, sticky='e')

        # Archivos
        archivos = ctk.CTkFrame(lateral)
        archivos.grid(row=1, column=0, sticky='ew', padx=16, pady=8)
        archivos.grid_columnconfigure(1, weight=1)
        filas = [
            ("Plantilla PDF", self.var_plantilla, self.buscar_plantilla),
            ("Datos Excel", self.var_datos, self.buscar_datos),
            ("Nombre del curso", self.var_curso, None),
            ("Carpeta de salida", self.var_salida, self.buscar_salida),
        ]
        for fila, (texto, var, comando) in enumerate(filas):
            pady = (10 if fila == 0 else 4, 10 if fila == len(filas) - 1 else 4)
            ctk.CTkLabel(archivos, text=texto).grid(row=fila, column=0, sticky='w', padx=(12, 8), pady=pady)
            ctk.CTkEntry(archivos, textvariable=var).grid(
                row=fila, column=1, columnspan=1 if comando else 2, sticky='ew',
                padx=(0, 4 if comando else 12), pady=pady)
            if comando:
                self._boton_secundario(archivos, text="…", width=34, command=comando).grid(
                    row=fila, column=2, padx=(4, 12), pady=pady)

        # Campos
        cabecera = ctk.CTkFrame(lateral, fg_color='transparent')
        cabecera.grid(row=2, column=0, sticky='ew', padx=16, pady=(8, 0))
        cabecera.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(cabecera, text="Campos", font=ctk.CTkFont(size=15, weight='bold')).grid(
            row=0, column=0, sticky='w')
        ctk.CTkButton(cabecera, text="+ Agregar campo", width=130,
                      command=self.agregar_campo).grid(row=0, column=1, sticky='e')

        self.lista_campos = ctk.CTkScrollableFrame(lateral, width=440, fg_color='transparent')
        self.lista_campos.grid(row=3, column=0, sticky='nsew', padx=10, pady=4)

        # Configuración
        pie = ctk.CTkFrame(lateral, fg_color='transparent')
        pie.grid(row=4, column=0, sticky='ew', padx=16, pady=(4, 16))
        pie.grid_columnconfigure((0, 1), weight=1)
        self._boton_secundario(pie, text="Guardar configuración", command=self.guardar_config).grid(
            row=0, column=0, sticky='ew', padx=(0, 4))
        self._boton_secundario(pie, text="Cargar configuración", command=self.cargar_config).grid(
            row=0, column=1, sticky='ew', padx=(4, 0))

    def _crear_panel_vista_previa(self):
        panel = ctk.CTkFrame(self, fg_color='transparent')
        panel.grid(row=0, column=1, sticky='nsew', padx=16, pady=16)
        panel.grid_columnconfigure(0, weight=1)
        panel.grid_rowconfigure(1, weight=1)

        # Barra superior: título, ayuda y navegador de registros
        barra = ctk.CTkFrame(panel, fg_color='transparent')
        barra.grid(row=0, column=0, sticky='ew', pady=(0, 8))
        barra.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(barra, text="Vista previa", font=ctk.CTkFont(size=15, weight='bold')).grid(
            row=0, column=0, sticky='w')
        ctk.CTkLabel(barra, text="Arrastrá para mover · tirador para el ancho · flechas para ajustar (Shift ×10)",
                     font=ctk.CTkFont(size=11), text_color=('gray45', 'gray55')).grid(
            row=0, column=1, sticky='w', padx=12)

        navegador = ctk.CTkFrame(barra, fg_color='transparent')
        navegador.grid(row=0, column=2, sticky='e')
        self.btn_anterior = self._boton_secundario(navegador, text="◀", width=34, command=self.registro_anterior)
        self.btn_anterior.pack(side='left')
        self.label_registro = ctk.CTkLabel(navegador, text="Sin datos", width=140)
        self.label_registro.pack(side='left', padx=6)
        self.btn_siguiente = self._boton_secundario(navegador, text="▶", width=34, command=self.registro_siguiente)
        self.btn_siguiente.pack(side='left')

        self.vista = VistaPrevia(panel, on_mover=self._on_campo_movido,
                                 on_redimensionar=self._on_campo_redimensionado,
                                 on_seleccionar=self._on_campo_seleccionado)
        self.vista.grid(row=1, column=0, sticky='nsew')

        # Barra inferior: estado, progreso y acción principal
        inferior = ctk.CTkFrame(panel, fg_color='transparent')
        inferior.grid(row=2, column=0, sticky='ew', pady=(10, 0))
        inferior.grid_columnconfigure(0, weight=1)
        self.label_estado = ctk.CTkLabel(inferior, text="", anchor='w', text_color=COLOR_ESTADO['info'])
        self.label_estado.grid(row=0, column=0, sticky='ew')
        self.progreso = ctk.CTkProgressBar(inferior)
        self.progreso.set(0)
        self.btn_generar = ctk.CTkButton(inferior, text="Generar certificados", height=40, width=200,
                                         font=ctk.CTkFont(size=14, weight='bold'), command=self.generar)
        self.btn_generar.grid(row=0, column=2, sticky='e', padx=(12, 0))

        self.actualizar_navegador()

    # --- Estado y vista previa ---------------------------------------------

    def _estado(self, texto, tipo='info'):
        self.label_estado.configure(text=texto, text_color=COLOR_ESTADO[tipo])

    def programar_vista_previa(self, demora=300):
        """Debounce con after(): todo corre en el hilo de Tkinter."""
        if self._preview_after:
            self.after_cancel(self._preview_after)
        self._preview_after = self.after(demora, self._refrescar_vista_previa)

    def _refrescar_vista_previa(self):
        self._preview_after = None
        if self._generando:
            return  # Se refresca al terminar
        plantilla = self.var_plantilla.get().strip()
        if not plantilla or not os.path.isfile(plantilla):
            self.vista.mostrar_mensaje("Elegí una plantilla PDF para ver la vista previa")
            return
        datos = None
        if self.df_excel is not None and len(self.df_excel) > 0:
            datos = self.df_excel.iloc[self.registro_actual]
        try:
            self.vista.renderizar(plantilla, self.obtener_config_campos(), datos)
        except Exception as e:
            logging.error(f"Error al actualizar la vista previa: {e}")
            self.vista.mostrar_mensaje(f"No se pudo generar la vista previa:\n{e}")

    def _card_de(self, campo):
        return next((c for c in self.cards if c.nombre() == campo), None)

    def _on_campo_movido(self, campo, x, y):
        card = self._card_de(campo)
        if card:
            card.set_posicion(x, y)

    def _on_campo_redimensionado(self, campo, ancho):
        card = self._card_de(campo)
        if card:
            card.set_ancho(ancho)

    def _on_campo_seleccionado(self, campo):
        for card in self.cards:
            card.set_seleccionado(campo is not None and card.nombre() == campo)

    # --- Navegador de registros ---------------------------------------------

    def actualizar_navegador(self):
        if self.df_excel is not None and len(self.df_excel) > 0:
            total = len(self.df_excel)
            self.label_registro.configure(text=f"Registro {self.registro_actual + 1} de {total}")
            self.btn_anterior.configure(state='normal' if self.registro_actual > 0 else 'disabled')
            self.btn_siguiente.configure(state='normal' if self.registro_actual < total - 1 else 'disabled')
        else:
            self.label_registro.configure(text="Sin datos")
            self.btn_anterior.configure(state='disabled')
            self.btn_siguiente.configure(state='disabled')

    def registro_anterior(self):
        if self.df_excel is not None and self.registro_actual > 0:
            self.registro_actual -= 1
            self.actualizar_navegador()
            self.programar_vista_previa(0)

    def registro_siguiente(self):
        if self.df_excel is not None and self.registro_actual < len(self.df_excel) - 1:
            self.registro_actual += 1
            self.actualizar_navegador()
            self.programar_vista_previa(0)

    # --- Campos ---------------------------------------------------------------

    def agregar_campo(self, valores=None):
        card = CampoCard(self.lista_campos, self.excel_columns, self.programar_vista_previa,
                         self.eliminar_campo, valores)
        card.pack(fill='x', padx=4, pady=6)
        self.cards.append(card)
        self.programar_vista_previa()
        return card

    def eliminar_campo(self, card):
        self.cards.remove(card)
        card.destroy()
        self.programar_vista_previa()

    def obtener_config_campos(self):
        return {card.nombre(): card.get_config() for card in self.cards if card.nombre()}

    # --- Archivos ---------------------------------------------------------------

    def buscar_plantilla(self):
        filename = filedialog.askopenfilename(filetypes=[("Archivos PDF", "*.pdf")])
        if filename:  # Si se cancela, se conserva la ruta anterior
            self.var_plantilla.set(filename)

    def buscar_datos(self):
        filename = filedialog.askopenfilename(filetypes=[("Archivos Excel", "*.xlsx")])
        if filename:
            self.var_datos.set(filename)
            self.cargar_columnas_excel(filename)

    def buscar_salida(self):
        directory = filedialog.askdirectory()
        if directory:
            self.var_salida.set(directory)

    def cargar_columnas_excel(self, filename, silencioso=False):
        try:
            self.df_excel = pd.read_excel(filename)
        except Exception as e:
            self.df_excel = None
            self.excel_columns = []
            logging.error(f"Error al cargar el Excel {filename}: {e}")
            if silencioso:
                self._estado(f"No se pudo leer el Excel: {e}", 'error')
            else:
                messagebox.showerror("Error", f"Error al cargar el archivo Excel:\n{e}")
        else:
            self.excel_columns = [str(c) for c in self.df_excel.columns]
            self._estado(f"Excel cargado: {len(self.df_excel)} registros, {len(self.excel_columns)} columnas")

        self.registro_actual = 0
        for card in self.cards:
            card.set_columnas(self.excel_columns)
        self.actualizar_navegador()
        self.programar_vista_previa()

    # --- Configuración ------------------------------------------------------

    def _config_actual(self):
        return {
            'plantilla': self.var_plantilla.get(),
            'datos': self.var_datos.get(),
            'curso': self.var_curso.get(),
            'salida': self.var_salida.get(),
            'campos': self.obtener_config_campos(),
        }

    def _escribir_config(self, ruta):
        with open(ruta, 'w', encoding='utf-8') as f:
            json.dump(self._config_actual(), f, ensure_ascii=False, indent=2)

    def guardar_config(self):
        try:
            self._escribir_config(CONFIG_ULTIMA)
        except OSError as e:
            logging.error(f"No se pudo guardar {CONFIG_ULTIMA}: {e}")

        salida = self.var_salida.get()
        curso = nombre_archivo_seguro(self.var_curso.get().strip().replace(' ', '_') or 'certificados')
        ruta = filedialog.asksaveasfilename(
            title="Guardar configuración", defaultextension='.json',
            filetypes=[("Configuración JSON", "*.json")],
            initialdir=salida if os.path.isdir(salida) else None,
            initialfile=f"config_{curso}.json")
        if not ruta:
            return
        try:
            self._escribir_config(ruta)
            self._estado(f"Configuración guardada en {ruta}", 'ok')
        except OSError as e:
            messagebox.showerror("Error", f"No se pudo guardar la configuración:\n{e}")

    def cargar_config(self, filename=None, silencioso=False):
        if not filename:
            filename = filedialog.askopenfilename(filetypes=[("Configuración JSON", "*.json")])
        if not filename or not os.path.exists(filename):
            return
        try:
            with open(filename, 'r', encoding='utf-8') as f:
                config = json.load(f)
        except Exception as e:
            logging.error(f"No se pudo leer la configuración {filename}: {e}")
            if not silencioso:
                messagebox.showerror("Error", f"No se pudo leer la configuración:\n{e}")
            return

        self.var_plantilla.set(config.get('plantilla', ''))
        self.var_datos.set(config.get('datos', ''))
        self.var_curso.set(config.get('curso', ''))
        self.var_salida.set(config.get('salida', ''))

        for card in self.cards:
            card.destroy()
        self.cards = []
        for nombre, valores in config.get('campos', {}).items():
            self.agregar_campo({**valores, 'nombre': nombre})

        datos = config.get('datos', '')
        if datos and os.path.isfile(datos):
            self.cargar_columnas_excel(datos, silencioso=True)

        if not silencioso:
            self._estado(f"Configuración cargada desde {filename}", 'ok')
        self.programar_vista_previa()

    # --- Generación ---------------------------------------------------------

    def elegir_columna_indice(self, columnas):
        """Diálogo modal para elegir la columna que nombra los archivos. Devuelve None si se cancela."""
        dialogo = ctk.CTkToplevel(self)
        dialogo.title("Columna para nombrar los archivos")
        dialogo.geometry("420x200")
        dialogo.resizable(False, False)
        dialogo.transient(self)

        resultado = {'columna': None}
        var = tk.StringVar(value=columnas[0] if columnas else '')

        ctk.CTkLabel(dialogo, text="El Excel no tiene la columna 'dni'.\n"
                                   "Elegí qué columna usar para nombrar los archivos:",
                     justify='left').pack(padx=20, pady=(20, 10), anchor='w')
        ctk.CTkOptionMenu(dialogo, values=columnas or [''], variable=var).pack(padx=20, fill='x')

        def aceptar():
            resultado['columna'] = var.get() or None
            dialogo.destroy()

        botones = ctk.CTkFrame(dialogo, fg_color='transparent')
        botones.pack(pady=20)
        self._boton_secundario(botones, text="Cancelar", command=dialogo.destroy).pack(side='left', padx=6)
        ctk.CTkButton(botones, text="Aceptar", command=aceptar).pack(side='left', padx=6)
        dialogo.protocol("WM_DELETE_WINDOW", dialogo.destroy)  # La X equivale a cancelar

        def modal():
            # grab_set falla si la ventana todavía no es visible
            try:
                dialogo.lift()
                dialogo.focus_force()
                dialogo.grab_set()
            except tk.TclError:
                pass

        dialogo.after(100, modal)
        self.wait_window(dialogo)
        return resultado['columna']

    def generar(self):
        if self._generando:
            return
        plantilla = self.var_plantilla.get().strip()
        datos = self.var_datos.get().strip()
        curso = self.var_curso.get().strip()
        salida = self.var_salida.get().strip()
        config = self.obtener_config_campos()

        faltantes = [nombre for nombre, valor in (("plantilla PDF", plantilla), ("datos Excel", datos),
                                                  ("nombre del curso", curso), ("carpeta de salida", salida),
                                                  ("al menos un campo", config)) if not valor]
        if faltantes:
            messagebox.showerror("Faltan datos", "Completá: " + ", ".join(faltantes) + ".")
            return
        if not os.path.isfile(plantilla):
            messagebox.showerror("Error", f"No se encuentra la plantilla:\n{plantilla}")
            return
        if not os.path.isdir(salida):
            messagebox.showerror("Error", f"La carpeta de salida no existe:\n{salida}")
            return

        try:
            df = pd.read_excel(datos)
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo leer el Excel:\n{e}")
            return

        columna = 'dni' if 'dni' in df.columns else self.elegir_columna_indice([str(c) for c in df.columns])
        if not columna:
            self._estado("Generación cancelada: no se eligió una columna para nombrar los archivos.", 'error')
            return

        try:
            self._escribir_config(CONFIG_ULTIMA)
        except OSError as e:
            logging.error(f"No se pudo guardar {CONFIG_ULTIMA}: {e}")

        self._generando = True
        self.vista.bloqueada = True
        self._salida_actual = salida
        self.btn_generar.configure(state='disabled', text="Generando…")
        self.progreso.set(0)
        self.progreso.grid(row=0, column=1, sticky='ew', padx=(12, 0))
        self._estado(f"Generando 0 de {len(df)}…")

        threading.Thread(target=self._trabajo_generar, daemon=True,
                         args=(plantilla, df, config, salida, curso, columna)).start()
        self.after(100, self._revisar_cola)

    def _trabajo_generar(self, *args):
        """Corre en un hilo aparte: solo se comunica con la interfaz mediante la cola."""
        try:
            generados, errores = generar_certificados(
                *args, on_progreso=lambda n, total: self._cola.put(('progreso', n, total)))
            self._cola.put(('fin', generados, errores))
        except Exception as e:
            logging.exception("Error al generar certificados")
            self._cola.put(('fin', 0, [str(e)]))

    def _revisar_cola(self):
        try:
            while True:
                mensaje = self._cola.get_nowait()
                if mensaje[0] == 'progreso':
                    _, n, total = mensaje
                    self.progreso.set(n / total if total else 1)
                    self._estado(f"Generando {n} de {total}…")
                else:
                    self._terminar_generacion(mensaje[1], mensaje[2])
                    return
        except queue.Empty:
            pass
        self.after(100, self._revisar_cola)

    def _terminar_generacion(self, generados, errores):
        self._generando = False
        self.vista.bloqueada = False
        self.btn_generar.configure(state='normal', text="Generar certificados")
        self.progreso.grid_remove()

        if errores:
            self._estado(f"{generados} certificados generados, {len(errores)} con errores "
                         f"(ver errores_certificados.log).", 'error')
            detalle = "\n".join(errores[:10]) + ("\n…" if len(errores) > 10 else "")
            messagebox.showwarning("Generación con errores",
                                   f"Se generaron {generados} certificados.\n"
                                   f"Fallaron {len(errores)}:\n\n{detalle}")
        else:
            self._estado(f"✓ {generados} certificados generados en {self._salida_actual}", 'ok')
        self.programar_vista_previa(0)


if __name__ == '__main__':
    ctk.set_appearance_mode('system')
    ctk.set_default_color_theme('blue')
    app = Application()
    app.mainloop()
