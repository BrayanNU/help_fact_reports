import pandas as pd
import unicodedata
from collections import defaultdict

# ==========================================
# HELPERS BÁSICOS
# ==========================================
def leer_excel_seguro(ruta, sheet_name, header=0):
    e_openpyxl = None
    e_xlrd = None

    # 1️⃣ Intentar como xlsx (openpyxl)
    try:
        return pd.read_excel(
            ruta,
            sheet_name=sheet_name,
            header=header,
            engine="openpyxl"
        )
    except Exception as e:
        e_openpyxl = e

    # 2️⃣ Intentar como xls real (xlrd)
    try:
        return pd.read_excel(
            ruta,
            sheet_name=sheet_name,
            header=header,
            engine="xlrd"
        )
    except Exception as e:
        e_xlrd = e

    print(
        f"[RECLAMOS] No se pudo leer {ruta} → "
        f"openpyxl: {e_openpyxl} | xlrd: {e_xlrd}"
    )
    return None



def normalizar_texto(texto):
    if pd.isna(texto):
        return ""
    texto = str(texto).lower().strip()
    texto = unicodedata.normalize("NFD", texto)
    texto = "".join(c for c in texto if unicodedata.category(c) != "Mn")
    return texto

def aplanar_columnas(df):
    if isinstance(df.columns, pd.MultiIndex):
        df.columns = [
            " ".join(str(x).strip() for x in col if x and str(x) != "nan").lower()
            for col in df.columns
        ]
    else:
        df.columns = [str(c).strip().lower() for c in df.columns]

def encontrar_columna_por_texto(df, texto):
    texto = normalizar_texto(texto)
    for col in df.columns:
        if texto in normalizar_texto(col):
            return col
    return None

def obtener_valor_seguro(valor):
    try:
        return float(pd.to_numeric(valor, errors="coerce")) if pd.notna(valor) else 0.0
    except Exception:
        return 0.0

# ==========================================
# SUCURSAL BASE EXACTA
# ==========================================
def obtener_sucursal_base(nombre):
    nombre = normalizar_texto(nombre)
    SUCURSALES_VALIDAS = {
        "astrobuns smash burgers": "Astrobuns | Smashburgers",
        "incheon - comida coreana": "Incheon | Korean Fried Chicken"
    }
    return SUCURSALES_VALIDAS.get(nombre, None)

def formatear_fecha_segura(fecha):
    try:
        fecha = pd.to_datetime(fecha, errors="coerce")
        return fecha.strftime("%d-%m-%Y") if not pd.isna(fecha) else "N/A"
    except Exception:
        return "N/A"

# ==========================================
# FUNCIÓN CENTRAL
# ==========================================
def procesar_finanzas_pedidosya(rutas, fecha_inicio=None, fecha_fin=None):
    SUCURSALES_BASE = ["Astrobuns | Smashburgers", "Incheon | Korean Fried Chicken"]

    resultados = defaultdict(lambda: {
        "total": 0.0,
        "cantidad_pedidos": 0,  # 👈 NUEVO
        "promociones_articulos": 0.0,
        "descuentos_fugaces": 0.0,
        "descuentos_cancelaciones": 0.0,
        "descuentos_reclamos": 0.0,
        "reintegros_tienda": 0.0
    })

    for ruta in rutas:
        try:
            hojas = pd.read_excel(ruta, sheet_name=None, header=[0,1], engine="openpyxl")
        except Exception:
            continue

        # ==============================
        # HOJAS NORMALES (PedidosYa)
        # ==============================
        for nombre_hoja, df in hojas.items():
            if df is None or df.empty:
                continue

            if nombre_hoja.strip().lower() == "cargos por cancelaciones":
                # NO procesar cancelaciones aquí, se hará aparte
                continue

            aplanar_columnas(df)

            col_estado = encontrar_columna_por_texto(df, "order status")
            col_sucursal = encontrar_columna_por_texto(df, "restaurant name")
            if not col_sucursal:
                continue

            # Filtrar solo delivered
            if col_estado:
                df[col_estado] = df[col_estado].astype(str).str.lower().str.strip()
                df = df[df[col_estado] == "delivered"]

            # TOTAL
            col_total = encontrar_columna_por_texto(df, "subtotal")
            if col_total:
                for _, fila in df.iterrows():
                    sucursal = obtener_sucursal_base(fila[col_sucursal])
                    if not sucursal:
                        continue

                    if resultados[sucursal]["total"] is None:
                        resultados[sucursal]["total"] = 0.0

                    resultados[sucursal]["total"] += obtener_valor_seguro(fila[col_total])
                    resultados[sucursal]["cantidad_pedidos"] += 1  # 👈 CONTAMOS PEDIDO

            # PROMOCIONES
            col_desc_art = encontrar_columna_por_texto(df, "discount funded by vendor")
            if col_desc_art:
                for _, fila in df.iterrows():
                    sucursal = obtener_sucursal_base(fila[col_sucursal])
                    if not sucursal: continue
                    if resultados[sucursal]["promociones_articulos"] is None: resultados[sucursal]["promociones_articulos"] = 0.0
                    resultados[sucursal]["promociones_articulos"] += obtener_valor_seguro(fila[col_desc_art])

            # DESCUENTOS FUGACES
            col_desc_fug = encontrar_columna_por_texto(df, "ads fee")
            if col_desc_fug:
                for _, fila in df.iterrows():
                    sucursal = obtener_sucursal_base(fila[col_sucursal])
                    if not sucursal: continue
                    if resultados[sucursal]["descuentos_fugaces"] is None: resultados[sucursal]["descuentos_fugaces"] = 0.0
                    resultados[sucursal]["descuentos_fugaces"] += obtener_valor_seguro(fila[col_desc_fug])

        # ==============================
        # HOJA CANCELACIONES (por nombre)
        # ==============================
        try:
            df_cancel = pd.read_excel(ruta, sheet_name="Cargos por cancelaciones", header=0, engine="openpyxl")
        except Exception:
            continue

        if df_cancel is not None and not df_cancel.empty:
            aplanar_columnas(df_cancel)
            col_sucursal = encontrar_columna_por_texto(df_cancel, "sucursal")
            col_estado_cancel = encontrar_columna_por_texto(df_cancel, "estado del pedido")
            col_cancel = encontrar_columna_por_texto(df_cancel, "comisión peya")

            if col_sucursal and col_estado_cancel and col_cancel:
                for _, fila in df_cancel.iterrows():
                    sucursal = obtener_sucursal_base(fila[col_sucursal])
                    if not sucursal: continue
                    estado = str(fila[col_estado_cancel]).lower().strip()
                    if estado != "rechazado":  # Solo cancelaciones
                        continue
                    if resultados[sucursal]["descuentos_cancelaciones"] is None:
                        resultados[sucursal]["descuentos_cancelaciones"] = 0.0
                    resultados[sucursal]["descuentos_cancelaciones"] += obtener_valor_seguro(fila[col_cancel])

        # ==============================
        # DESCUENTOS POR RECLAMOS (BUSCA EN TODAS LAS HOJAS)
        # ==============================
        try:
            hojas_reclamos = pd.read_excel(
                ruta,
                sheet_name=None,
                header=0,
                engine="openpyxl"
            )
        except Exception:
            hojas_reclamos = {}

        for nombre_hoja, df_reclamos in hojas_reclamos.items():
            if df_reclamos is None or df_reclamos.empty:
                continue

            aplanar_columnas(df_reclamos)

            col_reintegro = encontrar_columna_por_texto(
                df_reclamos, "reintegro al usuario por reclamo"
            )
            col_estado = encontrar_columna_por_texto(
                df_reclamos, "estado del pedido"
            )
            col_sucursal = encontrar_columna_por_texto(
                df_reclamos, "sucursal"
            )

            # 👉 Si no tiene estas columnas, NO es hoja de reclamos
            if not (col_reintegro and col_estado and col_sucursal):
                continue

            print(f"[RECLAMOS] Procesando hoja '{nombre_hoja}' en {ruta}")

            for _, fila in df_reclamos.iterrows():
                sucursal = obtener_sucursal_base(fila[col_sucursal])
                if not sucursal:
                    continue

                estado = str(fila[col_estado]).lower().strip()
                if estado != "confirmado":
                    continue

                if resultados[sucursal]["descuentos_reclamos"] is None:
                    resultados[sucursal]["descuentos_reclamos"] = 0.0

                valor = obtener_valor_seguro(fila[col_reintegro])
                resultados[sucursal]["descuentos_reclamos"] += valor

                print(
                    f"[RECLAMOS OK] {sucursal} | Hoja={nombre_hoja} | +S/.{valor:.2f}"
                )

        # ==============================
        # REINTEGRO A FAVOR DE LA TIENDA
        # ==============================
        try:
            hojas_reintegros = pd.read_excel(
                ruta,
                sheet_name=None,
                header=0,
                engine="openpyxl"
            )
        except Exception:
            hojas_reintegros = {}

        for nombre_hoja, df_reintegros in hojas_reintegros.items():

            if nombre_hoja.strip().lower() != "reintegros":
                continue

            if df_reintegros is None or df_reintegros.empty:
                continue

            print(f"[REINTEGROS] Procesando archivo: {ruta}")

            aplanar_columnas(df_reintegros)

            col_sucursal = encontrar_columna_por_texto(df_reintegros, "sucursal")
            col_estado = encontrar_columna_por_texto(df_reintegros, "estado del pedido")
            col_monto = encontrar_columna_por_texto(
                df_reintegros, "monto neto a reintegrar"
            )

            if not (col_sucursal and col_estado and col_monto):
                print(f"[REINTEGROS] Columnas incompletas en {ruta}")
                continue

            for _, fila in df_reintegros.iterrows():
                sucursal = obtener_sucursal_base(fila[col_sucursal])
                if not sucursal:
                    continue

                estado = str(fila[col_estado]).lower().strip()
                if estado != "rechazado":
                    continue

                if resultados[sucursal]["reintegros_tienda"] is None:
                    resultados[sucursal]["reintegros_tienda"] = 0.0

                valor = obtener_valor_seguro(fila[col_monto])
                resultados[sucursal]["reintegros_tienda"] += valor

                print(
                    f"[REINTEGROS OK] {sucursal} | Estado={estado} | +S/.{valor:.2f}"
                )

    # ======================================
    # SALIDA CONSOLA
    # ======================================
    # Normalizar null → 0.0 para PDF
    for suc in resultados:
        for k in resultados[suc]:
            if resultados[suc][k] is None:
                resultados[suc][k] = 0.0

    return resultados
