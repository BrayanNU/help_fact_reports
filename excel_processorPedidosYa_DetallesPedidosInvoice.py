import pandas as pd
import unicodedata

DEBUG = False

# ==========================================
# DEFINICIONES DE COLUMNAS
# ==========================================

HEADERS_MAP = {
    "Numero de Pedido": ["Numero de Pedido"],
    "Fecha de Pedido": ["Fecha de Pedido"],
    "Monto de la Venta ($)": ["Monto de la Venta ($)"],
    "Sucursal": ["Sucursal"]
}

CAMPOS_OBLIGATORIOS = [
    "Numero de Pedido",
    "Fecha de Pedido",
    "Monto de la Venta ($)",
    "Sucursal"
]

# sucursales que no quieres considerar
SUCURSALES_EXCLUIDAS = {
    "astrobuns-san miguel",
    "astrobuns-miraflores 2"
}


# ==========================================
# NORMALIZAR TEXTO
# ==========================================

def normalizar_texto(texto):

    if pd.isna(texto):
        return ""

    texto = str(texto).lower().strip()

    texto = unicodedata.normalize("NFD", texto)
    texto = "".join(
        c for c in texto
        if unicodedata.category(c) != "Mn"
    )

    return texto


# ==========================================
# LECTURA ROBUSTA DE EXCEL
# ==========================================

def leer_excel_seguro(ruta):

    try:
        return pd.read_excel(
            ruta,
            sheet_name=None,
            header=None,
            engine="openpyxl"
        )
    except Exception:

        try:
            return pd.read_excel(
                ruta,
                sheet_name=None,
                header=None,
                engine="xlrd"
            )
        except Exception:

            print(f"[ERROR] No se pudo leer {ruta}")
            return None


# ==========================================
# BUSCAR HEADER EN CUALQUIER FILA
# ==========================================

def encontrar_fila_header(df_raw):

    headers_norm = {
        normalizar_texto(h): h
        for h in CAMPOS_OBLIGATORIOS
    }

    for i, fila in df_raw.iterrows():

        fila_norm = [normalizar_texto(x) for x in fila]

        encontrados = []

        for celda in fila_norm:

            if celda in headers_norm:
                encontrados.append(headers_norm[celda])

        if all(h in encontrados for h in CAMPOS_OBLIGATORIOS):

            print(f"[HEADER ENCONTRADO] fila {i}")
            print(f"[HEADERS DETECTADOS] {encontrados}")

            return i

    return None
# ==========================================
# MAPEAR NOMBRES DE COLUMNAS
# ==========================================

def mapear_columnas(df):

    nuevas = {}

    headers_norm = {
        normalizar_texto(h): h
        for h in CAMPOS_OBLIGATORIOS
    }

    for col in df.columns:

        col_norm = normalizar_texto(col)

        if col_norm in headers_norm:

            nuevas[col] = headers_norm[col_norm]

            print(f"[COLUMNA MAPPEADA] '{col}' -> '{headers_norm[col_norm]}'")

    df = df.rename(columns=nuevas)

    return df

# ==========================================
# PROCESADOR PRINCIPAL
# ==========================================

def procesar_excel_pedidosya_invocie(rutas):

    dataframes = []

    for ruta in rutas:

        hojas = leer_excel_seguro(ruta)

        if hojas is None:
            continue

        print(f"\n[ARCHIVO] {ruta}")

        for nombre_hoja, df_raw in hojas.items():

            if df_raw is None or df_raw.empty:
                continue

            fila_header = encontrar_fila_header(df_raw)

            if fila_header is None:
                continue

            try:

                df = pd.read_excel(
                    ruta,
                    sheet_name=nombre_hoja,
                    header=fila_header,
                    engine="openpyxl"
                )

            except Exception:
                continue

            if df.empty:
                continue

            df = df.loc[:, ~df.columns.duplicated()]

            df = mapear_columnas(df)

            if not all(c in df.columns for c in CAMPOS_OBLIGATORIOS):
                continue

            # ==========================
            # LIMPIEZA
            # ==========================

            df = df.dropna(subset=[
                "Numero de Pedido",
                "Fecha de Pedido",
                "Monto de la Venta ($)"
            ])

            df["Monto de la Venta ($)"] = pd.to_numeric(
                df["Monto de la Venta ($)"],
                errors="coerce"
            )

            df = df[df["Monto de la Venta ($)"] > 0]

            # ==========================
            # FILTRAR SUCURSALES
            # ==========================

            df["Sucursal_norm"] = df["Sucursal"].apply(normalizar_texto)

            df = df[
                ~df["Sucursal_norm"].isin(SUCURSALES_EXCLUIDAS)
            ]

            df.drop(columns=["Sucursal_norm"], inplace=True)

            # ==========================
            # CAMPOS EXTRA
            # ==========================

            df["Cobrado por"] = "PedidosYa"
            df["Aplicativo"] = "PedidosYa"

            df.reset_index(drop=True, inplace=True)

            dataframes.append(df)

    if not dataframes:
        return None, None, None

    # ==========================
    # UNIR DATAFRAMES
    # ==========================

    df_final = pd.concat(dataframes, ignore_index=True)

    df_final = df_final.loc[:, ~df_final.columns.duplicated()]

    # ==========================
    # NORMALIZAR FECHAS
    # ==========================

    df_final["Fecha de Pedido"] = pd.to_datetime(
        df_final["Fecha de Pedido"],
        errors="coerce"
    )

    df_final = df_final[df_final["Fecha de Pedido"].notna()]

    df_final["Fecha de Pedido"] = df_final["Fecha de Pedido"].dt.floor("s")

    fecha_inicio = df_final["Fecha de Pedido"].min()
    fecha_fin = df_final["Fecha de Pedido"].max()

    return df_final, fecha_inicio, fecha_fin