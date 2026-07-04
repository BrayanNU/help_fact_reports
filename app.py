import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
from datetime import datetime
import pandas as pd
import unicodedata
from PIL import Image, ImageTk
logo_img_tk = None

# ==========================
# ESTADO GLOBAL: FILTRO DE MESES
# ==========================
# None = todos los meses (Opción 1)
# lista de ints = meses específicos (Opción 2)
meses_seleccionados = None  # e.g. [4, 5] para abril y mayo

from excel_processorRappi_DetallesVentas import procesar_finanzas_rappi
from excel_processorPedidosYa_DetallesVentas import procesar_finanzas_pedidosya
from excel_processorRappi_DetallesPedidos import procesar_excel_rappi as procesar_rappi
from excel_processorPedidosYa_DetallesPedidos import procesar_excel_pedidosya as procesar_pedidosya
from pdf_generator import generar_pdf
from config_manager import (
    load_footer_config,
    save_footer_config,
    load_invoice_config,
    save_invoice_config
)


fecha_actual = datetime.now().strftime("%d/%m/%Y")
# ==========================
# FUNCIONES PRINCIPALES
# ==========================
def normalizar_texto(texto):
    texto = texto.lower().strip()
    texto = unicodedata.normalize("NFD", texto)
    texto = "".join(c for c in texto if unicodedata.category(c) != "Mn")
    return texto

def procesar():
    global meses_seleccionados

    rutas = filedialog.askopenfilenames(
        title="Selecciona reportes",
        filetypes=[("Archivos Excel", "*.xlsx *.xls"), ("Todos", "*.*")]
    )
    if not rutas:
        return

    dfs = []
    fechas_inicio = []
    fechas_fin = []

    try:
        for ruta in rutas:
            df = fi = ff = None

            # 1️⃣ Intentar PedidosYa
            try:
                df, fi, ff = procesar_pedidosya([ruta], meses_seleccionados=meses_seleccionados)
            except Exception:
                df = None

            # 2️⃣ Si no fue PedidosYa → intentar Rappi
            if df is None:
                try:
                    df, fi, ff = procesar_rappi([ruta], meses_seleccionados=meses_seleccionados)
                except Exception:
                    df = None

            # 3️⃣ El archivo puede no tener pedidos, pero igual es válido
            if df is None:
                continue

            dfs.append(df)
            fechas_inicio.append(fi)
            fechas_fin.append(ff)

        if not dfs:
            raise Exception("No se pudo procesar ningún archivo válido.")

        df_final = pd.concat(dfs, ignore_index=True)
        df_final = df_final.loc[:, ~df_final.columns.duplicated()]
        df_final.reset_index(drop=True, inplace=True)

        # ==========================
        # UNIFICAR FECHAS (ANTI NaT)
        # ==========================

        # ==========================
        # UNIFICAR FECHAS (ULTRA BLINDADO)
        # ==========================

        def fecha_valida(f):
            return isinstance(f, pd.Timestamp)

        fechas_inicio_validas = [f for f in fechas_inicio if fecha_valida(f)]
        fechas_fin_validas = [f for f in fechas_fin if fecha_valida(f)]

        fecha_inicio = min(fechas_inicio_validas) if fechas_inicio_validas else None
        fecha_fin = max(fechas_fin_validas) if fechas_fin_validas else None

        fecha_inicio = min(fechas_inicio_validas) if fechas_inicio_validas else None
        fecha_fin = max(fechas_fin_validas) if fechas_fin_validas else None

        finanzas_pedidosya = procesar_finanzas_pedidosya(
            rutas=rutas,
            fecha_inicio=fecha_inicio,
            fecha_fin=fecha_fin,
            meses_seleccionados=meses_seleccionados
        )

        finanzas_rappi = procesar_finanzas_rappi(
            rutas=rutas,
            fecha_inicio=fecha_inicio,
            fecha_fin=fecha_fin,
            meses_seleccionados=meses_seleccionados
        )

        ruta_pdf = filedialog.asksaveasfilename(
        title="Guardar reporte PDF",
        defaultextension=".pdf",
        filetypes=[("Archivo PDF", "*.pdf")],
        initialfile="Reporte_Delivery.pdf"
        )

        if not ruta_pdf:
            return

        generar_pdf(
            df_final,
            fecha_inicio,
            fecha_fin,
            finanzas_pedidosya,
            finanzas_rappi,
            ruta_pdf
        )

        messagebox.showinfo("Éxito", "Reporte PDF generado correctamente.")

    except Exception as e:
        messagebox.showerror("Error", str(e))



# ==========================
# INTERFAZ
# ==========================
root = tk.Tk()
# ==========================
# ESTILOS GENERALES
# ==========================
style = ttk.Style(root)
style.theme_use("clam")  # importante para que acepte colores

BG_COLOR = "#F4F7FB"     # azul que ya usas
BTN_COLOR = "#FFFFFF"    # amarillo
BTN_TEXT = "#000000"

# Fondo general
style.configure(
    "TFrame",
    background=BG_COLOR
)

style.configure(
    "TLabel",
    background=BG_COLOR,
    foreground=BTN_TEXT
)

# Botones
style.configure(
    "TButton",
    background=BTN_COLOR,
    foreground=BTN_TEXT,
    font=("Arial", 10, "bold"),
    padding=8
)

style.map(
    "TButton",
    background=[("active", "#F4F7FA")]
)

root.title("Reportes Delivery (Rappi & PedidosYa)")
root.geometry("900x700")
root.resizable(False, False)


notebook = ttk.Notebook(root)
notebook.pack(fill="both", expand=True)

tab_reportes = ttk.Frame(notebook)
tab_config_header = ttk.Frame(notebook)
tab_config_footer = ttk.Frame(notebook)

notebook.add(tab_reportes, text="Generar Reporte")
notebook.add(tab_config_header, text="Configuración Factura")
notebook.add(tab_config_footer, text="Configuración Footer")

# ---------- TAB 1: REPORTES ----------

ttk.Label(
    tab_reportes,
    text="Generador de Reportes Delivery",
    font=("Arial", 14, "bold")
).pack(pady=10)

# ==========================
# SELECTOR DE MESES
# ==========================
frame_meses = ttk.LabelFrame(
    tab_reportes,
    text="Filtro de meses a procesar",
    padding=(15, 10)
)
frame_meses.pack(pady=10, padx=30, fill="x")

NOMBRES_MESES = [
    (1, "Enero"), (2, "Febrero"), (3, "Marzo"),
    (4, "Abril"), (5, "Mayo"), (6, "Junio"),
    (7, "Julio"), (8, "Agosto"), (9, "Septiembre"),
    (10, "Octubre"), (11, "Noviembre"), (12, "Diciembre")
]

# Variable para la opción ("todos" o "especificos")
var_modo_meses = tk.StringVar(value="todos")

# Diccionario de checkboxes de meses: {num_mes: BooleanVar}
checkbox_vars_meses = {num: tk.BooleanVar(value=False) for num, _ in NOMBRES_MESES}

frame_radio = ttk.Frame(frame_meses)
frame_radio.pack(anchor="w", pady=(0, 8))

ttk.Radiobutton(
    frame_radio,
    text="Procesar todos los meses detectados en los archivos subidos",
    variable=var_modo_meses,
    value="todos",
    command=lambda: _actualizar_estado_checkboxes()
).pack(anchor="w")

ttk.Radiobutton(
    frame_radio,
    text="Procesar solo meses específicos",
    variable=var_modo_meses,
    value="especificos",
    command=lambda: _actualizar_estado_checkboxes()
).pack(anchor="w", pady=(4, 0))

# Grid de checkboxes (4 por fila)
frame_checks = ttk.Frame(frame_meses)
frame_checks.pack(anchor="w", padx=20, pady=(4, 0))

_checkboxes_widgets = {}
for idx, (num, nombre) in enumerate(NOMBRES_MESES):
    fila = idx // 4
    col = idx % 4
    cb = ttk.Checkbutton(
        frame_checks,
        text=nombre,
        variable=checkbox_vars_meses[num],
        state="disabled"
    )
    cb.grid(row=fila, column=col, sticky="w", padx=10, pady=2)
    _checkboxes_widgets[num] = cb


def _actualizar_estado_checkboxes():
    """Habilita o deshabilita los checkboxes según el modo seleccionado."""
    estado = "normal" if var_modo_meses.get() == "especificos" else "disabled"
    for cb in _checkboxes_widgets.values():
        cb.configure(state=estado)
    if estado == "disabled":
        for v in checkbox_vars_meses.values():
            v.set(False)


def _obtener_meses_seleccionados():
    """Retorna lista de ints con los meses marcados, o None si es 'todos'."""
    global meses_seleccionados
    if var_modo_meses.get() == "todos":
        meses_seleccionados = None
        return None
    seleccionados = [num for num, var in checkbox_vars_meses.items() if var.get()]
    if not seleccionados:
        messagebox.showwarning(
            "Sin meses seleccionados",
            "Seleccionaste 'meses específicos' pero no marcaste ninguno.\n"
            "Se procesarán todos los meses."
        )
        meses_seleccionados = None
        return None
    meses_seleccionados = sorted(seleccionados)
    return meses_seleccionados


def procesar_con_filtro():
    """Wrapper que captura la selección de meses y luego llama a procesar()."""
    _obtener_meses_seleccionados()
    procesar()


ttk.Button(
    tab_reportes,
    text="Seleccionar Excel y generar reporte PDF",
    command=procesar_con_filtro
).pack(pady=20)


ttk.Label(
    tab_reportes,
    text="Nota: Si ocurre algún error al generar el reporte, se recomienda abrir los archivos Excel descargados de Rappi o PedidosYa y guardarlos nuevamente antes de subirlos al sistema, ya que en algunos casos pueden descargarse con inconsistencias desde la plataforma original.",
    wraplength=600,
    justify="center",
    font=("Arial", 9, "italic")
).pack(pady=10)

# Console Output Area
"""
console_frame = ttk.Frame(tab_reportes)
console_frame.pack(fill="both", expand=True, padx=10, pady=10)

ttk.Label(console_frame, text="Estado del proceso:", font=("Arial", 10, "bold")).pack(anchor="w")

console_text = tk.Text(console_frame, height=15, state="disabled", bg="#000000", fg="#00FF00", font=("Consolas", 9))
console_text.pack(fill="both", expand=True)

import sys

class RedirectText(object):
    def __init__(self, text_ctrl):
        self.output = text_ctrl

    def write(self, string):
        self.output.config(state="normal")
        self.output.insert("end", string)
        self.output.see("end")
        self.output.config(state="disabled")
        # Force GUI update
        self.output.update_idletasks()

    def flush(self):
        pass

sys.stdout = RedirectText(console_text)
sys.stderr = RedirectText(console_text)

"""

# ---------- TAB 2: CONFIG FACTURA ----------
config_invoice = load_invoice_config()

campos_invoice = {
    "Ciudad": "ciudad",
    "A la atención de": "atencion",
    "Dirección": "direccion",
    "Ubigeo": "ubigeo",
    "Fecha (DD-MM-YYYY)": "fecha",
    "RUC": "ruc",
    "Nº de factura": "numero_factura",
    "Customer ID": "customer_id",
    "Mensaje": "mensaje",
    "Logo (Ruta del archivo)": "logo_path"
}

entradas_invoice = {}

ttk.Label(
    tab_config_header,
    text="Configuración de Factura",
    font=("Arial", 13, "bold")
).pack(pady=15)

container_invoice = ttk.Frame(tab_config_header)
container_invoice.pack(pady=10, fill="x")

# columna izquierda: formulario
form_invoice = ttk.Frame(container_invoice)
form_invoice.grid(row=0, column=0, padx=10, sticky="n")

# columna derecha: logo
logo_frame = ttk.Frame(container_invoice)
logo_frame.grid(row=0, column=1, padx=20, sticky="n")

ttk.Label(
    logo_frame,
    text="Logo",
    font=("Arial", 11, "bold")
).pack(pady=(0, 10))

logo_preview = tk.Label(
    logo_frame,
    text="Sin logo",
    width=300,
    height=200,
    relief="solid",
    anchor="center"
)
logo_preview.pack()

def mostrar_logo(ruta):
    global logo_img_tk
    try:
        img = Image.open(ruta)
        img.thumbnail((220, 140))  # tamaño del preview
        logo_img_tk = ImageTk.PhotoImage(img)
        logo_preview.config(image=logo_img_tk, text="")
    except Exception as e:
        logo_preview.config(text="Error al cargar imagen", image="")


def seleccionar_logo(entry_widget):
    ruta = filedialog.askopenfilename(
        title="Seleccionar logo",
        filetypes=[("Imagen", "*.png *.jpg *.jpeg *.bmp")]
    )
    if ruta:
        entry_widget.delete(0, tk.END)
        entry_widget.insert(0, ruta)
        mostrar_logo(ruta)


for i, (label, key) in enumerate(campos_invoice.items()):
    ttk.Label(form_invoice, text=label, width=22, anchor="w").grid(row=i, column=0, pady=6)
    entry = tk.Entry(form_invoice, width=40)
    if key == "fecha":
        valor = fecha_actual
    else:
        valor = config_invoice.get(key, "")
    entry.insert(0, valor)
    entry.grid(row=i, column=1, pady=6)
    if key == "logo_path":
        tk.Button(form_invoice, text="Seleccionar archivo", command=lambda e=entry: seleccionar_logo(e)).grid(row=i, column=2, padx=5, pady=6)
    entradas_invoice[key] = entry

logo_guardado = config_invoice.get("logo_path", "")
if logo_guardado and os.path.exists(logo_guardado):
    mostrar_logo(logo_guardado)


def guardar_configuracion_factura():
    nueva_config = {k: e.get().strip() for k, e in entradas_invoice.items()}
    logo_path = nueva_config.get("logo_path", "")
    if logo_path and not os.path.exists(logo_path):
        messagebox.showwarning("Advertencia", f"El archivo de logo no existe:\n{logo_path}")
    save_invoice_config(nueva_config)
    messagebox.showinfo("Guardado", "Configuración de factura actualizada.")

tk.Button(
    tab_config_header,
    text="Guardar configuración",
    command=guardar_configuracion_factura,
    width=30,
    height=2
).pack(pady=25)

# ---------- TAB 3: CONFIG FOOTER ----------
config_footer = load_footer_config()

campos_footer = {
    "Razón social": "company_name",
    "Dirección": "address",
    "RUC": "ruc",
    "CCI": "cci",
    "IBAN": "iban"
}

entradas_footer = {}

ttk.Label(
    tab_config_footer,
    text="Configuración del pie de página",
    font=("Arial", 13, "bold")
).pack(pady=15)

form_footer = ttk.Frame(tab_config_footer)
form_footer.pack(pady=10)

for i, (label, key) in enumerate(campos_footer.items()):
    ttk.Label(form_footer, text=label, width=18, anchor="w").grid(row=i, column=0, pady=6)
    entry = tk.Entry(form_footer, width=42)
    entry.insert(0, config_footer.get(key, ""))
    entry.grid(row=i, column=1, pady=6)
    entradas_footer[key] = entry

def guardar_configuracion_footer():
    nueva_config = {k: e.get().strip() for k, e in entradas_footer.items()}
    save_footer_config(nueva_config)
    messagebox.showinfo("Guardado", "Configuración del pie de página actualizada.")

tk.Button(
    tab_config_footer,
    text="Guardar configuración",
    command=guardar_configuracion_footer,
    width=30,
    height=2
).pack(pady=25)

# ==========================
# INICIAR INTERFAZ
# ==========================
root.mainloop()
