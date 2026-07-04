import pandas as pd
import unicodedata

DEBUG = False

# ==========================
# DEFINICIONES POR IDIOMA (STRICT)
# ==========================
DEF_INGLES = {
    "headers": {
        "Numero de Pedido": ["order id"],
        "Fecha de Pedido": ["accepted at"],
        "Monto de la Venta ($)": ["subtotal"],
        "Sucursal": ["restaurant name"],
        "Estado de la Orden": ["order status"]
    },
    "status_valido": ["delivered"]
}

DEF_ESPANOL = {
    "headers": {
        "Numero de Pedido": ["nro de pedido"],
        "Fecha de Pedido": ["fecha del pedido"],
        "Monto de la Venta ($)": ["total del pedido"],
        "Sucursal": ["nombre del local"],
        "Estado de la Orden": ["estado del pedido"]
    },
    "status_valido": ["entregado"]
}

SUCURSALES_EXCLUIDAS = {
    "astrobuns-san miguel",
    "astrobuns-miraflores 2"
}


def normalizar_texto(texto):
    if pd.isna(texto):
        return ""
    texto = str(texto).lower().strip()
    texto = unicodedata.normalize("NFD", texto)
    texto = "".join(c for c in texto if unicodedata.category(c) != "Mn")
    return texto


# ==========================
# LECTURA DE EXCEL CORRUPTOS (STYLES.XML ROTOS)
# ==========================
def read_broken_xlsx(ruta_archivo):
    import zipfile
    import xml.etree.ElementTree as ET

    """
    Intenta leer un Excel (.xlsx) ignorando por completo styles.xml.
    Extrae sheet1.xml y sharedStrings.xml directamente del ZIP.
    Retorna un DataFrame de pandas limpio.
    """
    try:
        with zipfile.ZipFile(ruta_archivo, 'r') as z:
            # 1. Leer Shared Strings (si existe)
            shared_strings = []
            if "xl/sharedStrings.xml" in z.namelist():
                with z.open("xl/sharedStrings.xml") as f:
                    tree = ET.parse(f)
                    namespace = ""
                    # hack simple para namespace
                    for event, elem in ET.iterparse(f, events=("end",)):
                        if elem.tag.endswith("t"):
                            if elem.text:
                                shared_strings.append(elem.text)
            
            # Re-leer strings de forma robusta con namespace
            # (El hack anterior puede fallar si no se abre bien, mejor ir por full parse)
            shared_strings = []
            has_shared = "xl/sharedStrings.xml" in z.namelist()
            
            if has_shared:
                with z.open("xl/sharedStrings.xml") as f:
                    tree = ET.parse(f)
                    root = tree.getroot()
                    # extraer namespace del root tag
                    ns = root.tag.split("}")[0] + "}"
                    
                    for si in root.findall(f".//{ns}si"):
                        parts = []
                        for t in si.findall(f".//{ns}t"):
                            if t.text:
                                parts.append(t.text)
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
                    
                    # Logica simple para extraer valores en orden
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
                                    if idx < len(shared_strings):
                                        val = shared_strings[idx]
                                    else:
                                        val = v_text
                                except:
                                    val = v_text
                            else:
                                val = v_text
                        
                        row_data.append(val)
                    
                    if row_data:
                        data.append(row_data)

            if not data:
                return None
            
            return pd.DataFrame(data)

    except Exception as e:
        print(f"Error recuperando Excel corrupto {ruta_archivo}: {e}")
        return None


def read_xml_spreadsheet(ruta_archivo):
    """
    Lee archivos Excel-XML (SpreadsheetML) que Rappi/PedidosYa
    generan y que tienen extensión .xlsx pero no son ZIP reales.
    """
    import xml.etree.ElementTree as ET
    try:
        tree = ET.parse(ruta_archivo)
        root = tree.getroot()

        # Detectar namespace del documento
        ns_raw = root.tag.split("}")[0].lstrip("{") if "}" in root.tag else ""
        ns = f"{{{ns_raw}}}" if ns_raw else ""

        # Intentar localizar Worksheet y Table
        worksheets = root.findall(f".//{ns}Worksheet")
        if not worksheets:
            # Intentar sin namespace
            worksheets = root.findall(".//Worksheet")

        all_data = []
        for ws in worksheets:
            table = ws.find(f"{ns}Table") or ws.find("Table")
            if table is None:
                continue
            for row in table.findall(f"{ns}Row") or table.findall("Row"):
                row_data = []
                for cell in row.findall(f"{ns}Cell") or row.findall("Cell"):
                    data_tag = cell.find(f"{ns}Data") or cell.find("Data")
                    val = data_tag.text if data_tag is not None and data_tag.text else ""
                    row_data.append(val)
                if any(v.strip() for v in row_data):
                    all_data.append(row_data)

        if not all_data:
            return None

        return pd.DataFrame(all_data)

    except Exception as e:
        print(f"[XML-Excel] No se pudo leer {ruta_archivo}: {e}")
        return None


def read_html_mhtml(ruta_archivo):
    """
    Lee archivos MHTML/HTML que PedidosYa y Rappi generan al
    descargar reportes directamente desde el navegador.
    Tienen extensión .xls o .xlsx pero son HTML por dentro.
    """
    try:
        # pandas.read_html puede parsear tablas HTML incluyendo MHTML
        tablas = pd.read_html(ruta_archivo, header=None, flavor="lxml")
        for tabla in tablas:
            if tabla is not None and not tabla.empty and len(tabla.columns) > 2:
                return tabla
        return None
    except Exception:
        pass

    # Segundo intento: leer como texto y buscar tablas HTML dentro
    try:
        encodings = ["utf-8", "latin-1", "utf-16"]
        contenido = None
        for enc in encodings:
            try:
                with open(ruta_archivo, "r", encoding=enc, errors="replace") as f:
                    contenido = f.read()
                break
            except Exception:
                continue

        if contenido:
            import io
            tablas = pd.read_html(io.StringIO(contenido), header=None)
            for tabla in tablas:
                if tabla is not None and not tabla.empty and len(tabla.columns) > 2:
                    return tabla
    except Exception as e:
        print(f"[HTML-MHTML] No se pudo leer {ruta_archivo}: {e}")

    return None


# ==========================
# HEADER FLEXIBLE POR BLOQUE
# ==========================
def encontrar_fila_header(df_raw, diccionario_headers):

    obligatorias = [
        "Numero de Pedido",
        "Fecha de Pedido",
        "Monto de la Venta ($)",
        "Sucursal",
        "Estado de la Orden"
    ]

    for i, fila in df_raw.iterrows():

        fila_norm = [normalizar_texto(c) for c in fila]

        encontrados = set()

        for campo, sinonimos in diccionario_headers.items():
            for s in sinonimos:
                if any(s in celda for celda in fila_norm):
                    encontrados.add(campo)

        if all(c in encontrados for c in obligatorias):
            return i

    return None


def procesar_excel_pedidosya(rutas, meses_seleccionados=None):

    dataframes = []

    for ruta in rutas:

        hojas_raw = None
        
        # INTENTO 1: LECTURA STANDARD (xlsx)
        try:
            hojas_raw = pd.read_excel(
                ruta,
                sheet_name=None,
                header=None,
                engine="openpyxl"
            )
        except Exception:
            pass

        # INTENTO 2: LECTURA XLS ANTIGUO (xlrd)
        if hojas_raw is None:
            try:
                hojas_raw = pd.read_excel(
                    ruta,
                    sheet_name=None,
                    header=None,
                    engine="xlrd"
                )
            except Exception:
                pass

        # INTENTO 3: LECTURA MODO RECUPERACION (xlsx ZIP corrupto)
        if hojas_raw is None:
            print(f"Advertencia: Falló lectura standard de {ruta}. Intentando modo recuperación...")
            df_recuperado = read_broken_xlsx(ruta)
            if df_recuperado is not None:
                hojas_raw = {"Sheet1": df_recuperado}

        # INTENTO 4: LECTURA COMO XML-SPREADSHEET (formato Rappi/PedidosYa descarga directa)
        if hojas_raw is None:
            df_recuperado = read_xml_spreadsheet(ruta)
            if df_recuperado is not None:
                hojas_raw = {"Sheet1": df_recuperado}

        # INTENTO 5: LECTURA HTML/MHTML (archivo descargado desde navegador)
        if hojas_raw is None:
            df_recuperado = read_html_mhtml(ruta)
            if df_recuperado is not None:
                hojas_raw = {"Sheet1": df_recuperado}

        if hojas_raw is None:
            print(f"⚠️ No se pudo leer: {ruta}")
            continue



        for nombre_hoja, df_raw in hojas_raw.items():

            if df_raw is None or df_raw.empty:
                continue

            # ==========================
            # BUSQUEDA SECUENCIAL (PROBAR INGLES -> SI FALLA -> ESPANOL)
            # ==========================
            definicion_usada = None
            fila_header = None
            
            for definicion in [DEF_INGLES, DEF_ESPANOL]:
                fh = encontrar_fila_header(df_raw, definicion["headers"])
                if fh is not None:
                    fila_header = fh
                    definicion_usada = definicion
                    break
            
            if fila_header is None:
                continue
            
            # Si leimos con pandas normal, hay que recargar con header correcto
            # Si leimos con recuperacion, ya tenemos el df "crudo" y hay que ajustar columnas manualmente
            
            if "read_broken_xlsx" in str(type(hojas_raw)): # hack conceptual, en realidad chequeamos si vino de recovery
               pass 

            # DUAL PATH: RELOAD VS RESTRUCTURE
            try:
                # Si el df_raw ya viene de read_broken_xlsx, NO PODEMOS usar read_excel de nuevo con header=X
                # tenemos que promover la fila X a header manualmente.
                
                # Como saber si vino de recovery? simple: si intentamos read_excel y falla, usamos el raw
                try:
                    df = pd.read_excel(
                        ruta,
                        sheet_name=nombre_hoja,
                        header=fila_header,
                        engine="openpyxl"
                    )
                except Exception:
                    # Fallback para cuando read_excel falla pero read_broken_xlsx funcionó antes
                     # df_raw contiene todo. Promovemos fila_header
                    df = df_raw.copy()
                    df.columns = df.iloc[fila_header]
                    df = df.iloc[fila_header + 1:].reset_index(drop=True)

            except Exception:
                continue

            if df.empty:
                continue

            df = df.loc[:, ~df.columns.duplicated()]

            # ==========================
            # RENOMBRADO EXACTO
            # ==========================
            nuevas = {}
            mapping_headers = definicion_usada["headers"]

            for col in df.columns:
                col_norm = normalizar_texto(col)

                for nombre_final, sinonimos in mapping_headers.items():
                    if any(s in col_norm for s in sinonimos):
                        nuevas[col] = nombre_final

            if not nuevas:
                continue

            df = df.rename(columns=nuevas)

            # ==========================
            # LIMPIEZA
            # ==========================
            obligatorias_limpieza = [
                "Numero de Pedido",
                "Fecha de Pedido",
                "Monto de la Venta ($)"
            ]

            df = df.dropna(subset=obligatorias_limpieza)

            df["Monto de la Venta ($)"] = pd.to_numeric(
                df["Monto de la Venta ($)"], errors="coerce"
            )

            df = df[df["Monto de la Venta ($)"] > 0]

            # ==========================
            # FILTRO ESTADO SEGUN IDIOMA DETECTADO
            # ==========================
            if "Estado de la Orden" in df.columns:
                estados_validos = definicion_usada["status_valido"]

                df["Estado de la Orden"] = (
                    df["Estado de la Orden"]
                    .astype(str)
                    .str.lower()
                    .str.strip()
                )

                df = df[df["Estado de la Orden"].isin(estados_validos)]

            # ==========================
            # FILTRO SUCURSAL
            # ==========================
            if "Sucursal" in df.columns:
                df["Sucursal_norm"] = df["Sucursal"].apply(normalizar_texto)
                df = df[~df["Sucursal_norm"].isin(SUCURSALES_EXCLUIDAS)]
                df.drop(columns=["Sucursal_norm"], inplace=True)

            df["Cobrado por"] = "PedidosYa"
            df["Aplicativo"] = "PedidosYa"

            df.reset_index(drop=True, inplace=True)
            dataframes.append(df)

    if not dataframes:
        return None, None, None

    df_final = pd.concat(dataframes, ignore_index=True)
    df_final = df_final.loc[:, ~df_final.columns.duplicated()]
    df_final.reset_index(drop=True, inplace=True)

    # ==========================
    # NORMALIZACIÓN TOTAL FECHAS
    # ==========================

    df_final["Fecha de Pedido"] = pd.to_datetime(
        df_final["Fecha de Pedido"].astype(str),
        errors="coerce"
    )

    df_final = df_final[df_final["Fecha de Pedido"].notna()].copy()

    # ==========================
    # FILTRO POR MESES SELECCIONADOS
    # ==========================
    if meses_seleccionados:
        df_final = df_final[
            df_final["Fecha de Pedido"].dt.month.isin(meses_seleccionados)
        ].copy()

    if df_final.empty:
        return None, None, None

    # forzar mismo dtype en todo
    df_final["Fecha de Pedido"] = df_final["Fecha de Pedido"].dt.floor("s")

    fecha_inicio = pd.Timestamp(df_final["Fecha de Pedido"].min())
    fecha_fin = pd.Timestamp(df_final["Fecha de Pedido"].max())

    print("dtype final:", df_final["Fecha de Pedido"].dtype)

    print("inicio:", fecha_inicio, type(fecha_inicio))
    print("fin:", fecha_fin, type(fecha_fin))

    return df_final, fecha_inicio, fecha_fin

