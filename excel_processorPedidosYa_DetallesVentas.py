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
# LECTURA DE EXCEL CORRUPTOS (STYLES.XML ROTOS)
# ==========================================
def read_broken_xlsx(ruta_archivo):
    import zipfile
    import xml.etree.ElementTree as ET

    try:
        with zipfile.ZipFile(ruta_archivo, 'r') as z:
            # 1. Leer Shared Strings (si existe)
            shared_strings = []
            if "xl/sharedStrings.xml" in z.namelist():
                with z.open("xl/sharedStrings.xml") as f:
                    tree = ET.parse(f)
                    # hack simple para namespace
                    for event, elem in ET.iterparse(f, events=("end",)):
                        if elem.tag.endswith("t"):
                            if elem.text:
                                shared_strings.append(elem.text)
            
            # Re-leer strings de forma robusta con namespace
            shared_strings = []
            has_shared = "xl/sharedStrings.xml" in z.namelist()
            
            if has_shared:
                with z.open("xl/sharedStrings.xml") as f:
                    tree = ET.parse(f)
                    root = tree.getroot()
                    ns = root.tag.split("}")[0] + "}"
                    for si in root.findall(f".//{ns}si"):
                        parts = []
                        for t in si.findall(f".//{ns}t"):
                            if t.text: parts.append(t.text)
                        shared_strings.append("".join(parts))

            # 2. Leer Sheet1
            if "xl/worksheets/sheet1.xml" not in z.namelist():
                return None
            
            data = []
            with z.open("xl/worksheets/sheet1.xml") as f:
                tree = ET.parse(f)
                root = tree.getroot()
                ns = root.tag.split("}")[0] + "}"

                for row in root.findall(f".//{ns}row"):
                    row_data = []
                    cells = row.findall(f"{ns}c")
                    for cell in cells:
                        val = ""
                        ctype = cell.get("t")
                        v_tag = cell.find(f"{ns}v")
                        is_tag = cell.find(f"{ns}is")
                        
                        if is_tag is not None:
                             t_tag = is_tag.find(f"{ns}t")
                             if t_tag is not None and t_tag.text:
                                 val = t_tag.text
                        elif v_tag is not None and v_tag.text:
                            v_text = v_tag.text
                            if ctype == "s" and has_shared:
                                try:
                                    idx = int(v_text)
                                    val = shared_strings[idx] if idx < len(shared_strings) else v_text
                                except:
                                    val = v_text
                            else:
                                val = v_text
                        row_data.append(val)
                    if row_data:
                        data.append(row_data)

            return pd.DataFrame(data) if data else None

    except Exception as e:
        print(f"Error recuperando Excel corrupto {ruta_archivo}: {e}")
        return None


# ==========================================
# BLOQUE 1: LOGICA ORIGINAL (INGLES)
# ==========================================
def procesar_bloque_ingles(ruta, hojas, resultados):
    processed_count = 0
    
    # ==============================
    # HOJAS NORMALES (PedidosYa)
    # ==============================
    for nombre_hoja, df in hojas.items():
        if df is None or df.empty:
            continue

        if "cargos por cancelaciones" in nombre_hoja.strip().lower():
            continue

        aplanar_columnas(df)

        col_estado = encontrar_columna_por_texto(df, "order status")
        col_sucursal = encontrar_columna_por_texto(df, "restaurant name")
        if not col_sucursal:
            continue
        
        # Marcamos que encontramos columnas validas
        processed_count += 1

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
                resultados[sucursal]["cantidad_pedidos"] += 1
        # PRINT DE TOTALES
        for sucursal, datos in resultados.items():
            if datos["total"] is not None:
                print(f"*****BRAYANNN TOTAL ZONA {sucursal.upper()}*****")
                print(f"Total: {datos['total']:.2f}")
                print()

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
        df_cancel = None

    if df_cancel is not None and not df_cancel.empty:
        aplanar_columnas(df_cancel)
        col_sucursal = encontrar_columna_por_texto(df_cancel, "sucursal")
        col_estado_cancel = encontrar_columna_por_texto(df_cancel, "estado del pedido")
        col_cancel = encontrar_columna_por_texto(df_cancel, "comisión peya")

        if col_sucursal and col_estado_cancel and col_cancel:
            processed_count += 1
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
    # DESCUENTOS POR RECLAMOS
    # ==============================
    try:
        hojas_reclamos = pd.read_excel(ruta, sheet_name=None, header=0, engine="openpyxl")
    except Exception:
        hojas_reclamos = {}

    for nombre_hoja, df_reclamos in hojas_reclamos.items():
        if df_reclamos is None or df_reclamos.empty:
            continue

        aplanar_columnas(df_reclamos)

        col_reintegro = encontrar_columna_por_texto(df_reclamos, "reintegro al usuario por reclamo")
        col_estado = encontrar_columna_por_texto(df_reclamos, "estado del pedido")
        col_sucursal = encontrar_columna_por_texto(df_reclamos, "sucursal")

        if not (col_reintegro and col_estado and col_sucursal):
            continue
        
        processed_count += 1
        print(f"[RECLAMOS] Procesando hoja '{nombre_hoja}' en {ruta}")

        for _, fila in df_reclamos.iterrows():
            sucursal = obtener_sucursal_base(fila[col_sucursal])
            if not sucursal: continue

            estado = str(fila[col_estado]).lower().strip()
            if estado != "confirmado": continue

            if resultados[sucursal]["descuentos_reclamos"] is None:
                resultados[sucursal]["descuentos_reclamos"] = 0.0

            valor = obtener_valor_seguro(fila[col_reintegro])
            resultados[sucursal]["descuentos_reclamos"] += valor
            print(f"[RECLAMOS OK] {sucursal} | Hoja={nombre_hoja} | +S/.{valor:.2f}")

    # ==============================
    # REINTEGRO A FAVOR DE LA TIENDA
    # ==============================
    try:
        hojas_reintegros = pd.read_excel(ruta, sheet_name=None, header=0, engine="openpyxl")
    except Exception:
        hojas_reintegros = {}

    for nombre_hoja, df_reintegros in hojas_reintegros.items():
        if "reintegros" not in nombre_hoja.strip().lower():
            continue
        if df_reintegros is None or df_reintegros.empty:
            continue

        print(f"[REINTEGROS] Procesando archivo: {ruta}")
        aplanar_columnas(df_reintegros)

        col_sucursal = encontrar_columna_por_texto(df_reintegros, "sucursal")
        col_estado = encontrar_columna_por_texto(df_reintegros, "estado del pedido")
        col_monto = encontrar_columna_por_texto(df_reintegros, "monto neto a reintegrar")
        col_orden_entregado = encontrar_columna_por_texto(df_reintegros, "orden entregada al repartidor")

        if not (col_sucursal and col_estado and col_monto and col_orden_entregado):
            print(f"[REINTEGROS] Columnas incompletas en {ruta}")
            continue

        processed_count += 1

        for _, fila in df_reintegros.iterrows():
            sucursal = obtener_sucursal_base(fila[col_sucursal])
            if not sucursal: continue

            estado = str(fila[col_estado]).lower().strip()
            if estado != "rechazado": continue

            tipo_reintegro = str(fila[col_orden_entregado]).lower().strip()
            if tipo_reintegro != "no": continue

            if resultados[sucursal]["reintegros_tienda"] is None:
                resultados[sucursal]["reintegros_tienda"] = 0.0

            valor = obtener_valor_seguro(fila[col_monto])
            resultados[sucursal]["reintegros_tienda"] += valor
            print(f"[REINTEGROS OK] {sucursal} | Estado={estado} | +S/.{valor:.2f}")

    return processed_count > 0


# ==========================================================================================================
# BLOQUE 2: COPIA EXACTA CON CAMBIOS STRING
# ==========================================================================================================
def procesar_bloque_espanol(ruta, hojas, resultados):
    processed_count = 0
    
    # ==============================
    # HOJAS NORMALES (PedidosYa)
    # ==============================
    for nombre_hoja, df in hojas.items():
        if df is None or df.empty:
            continue

        if "cargos por cancelaciones" in nombre_hoja.strip().lower():
            continue

        aplanar_columnas(df)
        print("==== COLUMNAS ESPAÑOL ====")
        for c in df.columns:
            print(repr(c))

        col_estado = encontrar_columna_por_texto(df, "estado del pedido")
        col_sucursal = encontrar_columna_por_texto(df, "nombre del local")
        if not col_sucursal:
            continue
        
        # Marcamos que encontramos columnas validas
        processed_count += 1

        # Filtrar solo delivered
        if col_estado:
            df[col_estado] = df[col_estado].astype(str).str.lower().str.strip()
            df = df[df[col_estado] == "entregado"]

        # TOTAL
        col_total = encontrar_columna_por_texto(df, "total del pedido")
        if col_total:
            for _, fila in df.iterrows():
                sucursal = obtener_sucursal_base(fila[col_sucursal])
                if not sucursal:
                    continue

                if resultados[sucursal]["total"] is None:
                    resultados[sucursal]["total"] = 0.0

                resultados[sucursal]["total"] += obtener_valor_seguro(fila[col_total])
                resultados[sucursal]["cantidad_pedidos"] += 1

        # PROMOCIONES - use actual Spanish column names
        col_desc_art = encontrar_columna_por_texto(df, "descuentos subsidiados por la tienda")
        if col_desc_art:
            for _, fila in df.iterrows():
                sucursal = obtener_sucursal_base(fila[col_sucursal])
                if not sucursal: continue
                if resultados[sucursal]["promociones_articulos"] is None: resultados[sucursal]["promociones_articulos"] = 0.0
                resultados[sucursal]["promociones_articulos"] += obtener_valor_seguro(fila[col_desc_art])

        # DESCUENTOS COMERCIALES - use actual Spanish column names
        col_desc_fug = encontrar_columna_por_texto(df, "cargos por descuentos fugaces")
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
        df_cancel = None

    if df_cancel is not None and not df_cancel.empty:
        aplanar_columnas(df_cancel)
        col_sucursal = encontrar_columna_por_texto(df_cancel, "sucursal")
        col_estado_cancel = encontrar_columna_por_texto(df_cancel, "estado del pedido")
        col_cancel = encontrar_columna_por_texto(df_cancel, "comisión peya")

        if col_sucursal and col_estado_cancel and col_cancel:
            processed_count += 1
            for _, fila in df_cancel.iterrows():
                sucursal = obtener_sucursal_base(fila[col_sucursal])
                if not sucursal: continue
                estado = str(fila[col_estado_cancel]).lower().strip()
                if estado != "rechazado":
                    continue
                if resultados[sucursal]["descuentos_cancelaciones"] is None:
                    resultados[sucursal]["descuentos_cancelaciones"] = 0.0
                resultados[sucursal]["descuentos_cancelaciones"] += obtener_valor_seguro(fila[col_cancel])

    # ==============================
    # DESCUENTOS POR RECLAMOS
    # ==============================
    try:
        hojas_reclamos = pd.read_excel(ruta, sheet_name=None, header=0, engine="openpyxl")
    except Exception:
        hojas_reclamos = {}

    for nombre_hoja, df_reclamos in hojas_reclamos.items():
        if df_reclamos is None or df_reclamos.empty:
            continue

        aplanar_columnas(df_reclamos)

        col_reintegro = encontrar_columna_por_texto(df_reclamos, "reintegro al usuario por reclamo")
        col_estado = encontrar_columna_por_texto(df_reclamos, "estado del pedido")
        col_sucursal = encontrar_columna_por_texto(df_reclamos, "sucursal")

        if not (col_reintegro and col_estado and col_sucursal):
            continue
        
        processed_count += 1
        print(f"[RECLAMOS] Procesando hoja '{nombre_hoja}' en {ruta}")

        for _, fila in df_reclamos.iterrows():
            sucursal = obtener_sucursal_base(fila[col_sucursal])
            if not sucursal: continue

            estado = str(fila[col_estado]).lower().strip()
            if estado != "confirmado": continue

            if resultados[sucursal]["descuentos_reclamos"] is None:
                resultados[sucursal]["descuentos_reclamos"] = 0.0

            valor = obtener_valor_seguro(fila[col_reintegro])
            resultados[sucursal]["descuentos_reclamos"] += valor
            print(f"[RECLAMOS OK] {sucursal} | Hoja={nombre_hoja} | +S/.{valor:.2f}")

    # ==============================
    # REINTEGRO A FAVOR DE LA TIENDA
    # ==============================
    try:
        hojas_reintegros = pd.read_excel(ruta, sheet_name=None, header=0, engine="openpyxl")
    except Exception:
        hojas_reintegros = {}

    for nombre_hoja, df_reintegros in hojas_reintegros.items():
        if "reintegros" not in nombre_hoja.strip().lower():
            continue
        if df_reintegros is None or df_reintegros.empty:
            continue

        print(f"[REINTEGROS] Procesando archivo: {ruta}")
        aplanar_columnas(df_reintegros)

        col_sucursal = encontrar_columna_por_texto(df_reintegros, "sucursal")
        col_estado = encontrar_columna_por_texto(df_reintegros, "estado del pedido")
        col_monto = encontrar_columna_por_texto(df_reintegros, "monto neto a reintegrar")
        col_orden_entregado = encontrar_columna_por_texto(df_reintegros, "orden entregada al repartidor")

        if not (col_sucursal and col_estado and col_monto and col_orden_entregado):
            print(f"[REINTEGROS] Columnas incompletas en {ruta}")
            continue
        
        processed_count += 1

        for _, fila in df_reintegros.iterrows():
            sucursal = obtener_sucursal_base(fila[col_sucursal])
            if not sucursal: continue

            estado = str(fila[col_estado]).lower().strip()
            if estado != "rechazado": continue

            tipo_reintegro = str(fila[col_orden_entregado]).lower().strip()
            if tipo_reintegro != "no": continue

            if resultados[sucursal]["reintegros_tienda"] is None:
                resultados[sucursal]["reintegros_tienda"] = 0.0

            valor = obtener_valor_seguro(fila[col_monto])
            resultados[sucursal]["reintegros_tienda"] += valor
            print(f"[REINTEGROS OK] {sucursal} | Estado={estado} | +S/.{valor:.2f}")

    return processed_count > 0


# ==========================================
# FUNCIÓN CENTRAL
# ==========================================
def procesar_finanzas_pedidosya(rutas, fecha_inicio=None, fecha_fin=None):
    SUCURSALES_BASE = ["Astrobuns | Smashburgers", "Incheon | Korean Fried Chicken"]

    resultados = defaultdict(lambda: {
        "total": 0.0,
        "cantidad_pedidos": 0,
        "promociones_articulos": 0.0,
        "descuentos_fugaces": 0.0,
        "descuentos_cancelaciones": 0.0,
        "descuentos_reclamos": 0.0,
        "reintegros_tienda": 0.0
    })

    for ruta in rutas:
        hojas = None
        origen_recuperado = False
        
        # 1. Intentar lectura standard
        try:
            hojas = pd.read_excel(ruta, sheet_name=None, header=[0,1], engine="openpyxl")
        except Exception:
            # 2. Intentar recuperacion
            print(f"Advertencia: Falló lectura standard de {ruta}. Intentando modo recuperación...")
            df_recuperado = read_broken_xlsx(ruta)
            if df_recuperado is not None:
                # Si viene de recuperado, es data cruda sin header multi-index
                # Asignamos primera fila como header para simular estructura compatible
                new_header = df_recuperado.iloc[0]
                df_recuperado = df_recuperado[1:]
                df_recuperado.columns = new_header
                hojas = {"Sheet1": df_recuperado}
                origen_recuperado = True
        
        if hojas is None:
            continue
            
        # 3. Probar Bloque 1 (Ingles)
        if procesar_bloque_ingles(ruta, hojas, resultados):
            continue
            
        # 4. Probar Bloque 2 (Espanol - Fallback)
        procesar_bloque_espanol(ruta, hojas, resultados)

    # ======================================
    # SALIDA CONSOLA
    # ======================================
    for suc in resultados:
        for k in resultados[suc]:
            if resultados[suc][k] is None:
                resultados[suc][k] = 0.0

    return resultados
