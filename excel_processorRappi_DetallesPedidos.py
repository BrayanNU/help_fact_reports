import pandas as pd
import unicodedata

COLUMNAS_OBJETIVO = {
    "id de la orden": "ID de la orden",
    "fecha de creacion orden": "Fecha de creación orden",
    "venta bruta": "Venta Bruta",
    "nombre de la tienda": "Nombre de la tienda",
    "estado de la orden": "Estado de la orden"
}


def normalizar_texto(texto):
    if pd.isna(texto):
        return ""
    texto = str(texto).lower().strip()
    texto = unicodedata.normalize("NFD", texto)
    texto = "".join(c for c in texto if unicodedata.category(c) != "Mn")
    return texto


def encontrar_fila_header(df_raw):
    for i, fila in df_raw.iterrows():
        fila_norm = [normalizar_texto(c) for c in fila]
        if all(
            any(clave in celda for celda in fila_norm)
            for clave in COLUMNAS_OBJETIVO.keys()
        ):
            return i
    return None


def obtener_fechas_resumen(ruta):
    try:
        df = pd.read_excel(
            ruta,
            sheet_name="Resumen",
            header=None,
            engine="openpyxl"
        )
        return df.iloc[2, 3], df.iloc[3, 3]
    except Exception:
        return "N/A", "N/A"


def read_xml_spreadsheet(ruta_archivo):
    """
    Lee archivos Excel-XML (SpreadsheetML) que Rappi genera
    con extensión .xlsx pero que no son archivos ZIP reales.
    """
    import xml.etree.ElementTree as ET
    try:
        tree = ET.parse(ruta_archivo)
        root = tree.getroot()

        ns_raw = root.tag.split("}")[0].lstrip("{") if "}" in root.tag else ""
        ns = f"{{{ns_raw}}}" if ns_raw else ""

        worksheets = root.findall(f".//{ns}Worksheet") or root.findall(".//Worksheet")

        all_data = []
        for ws in worksheets:
            table = ws.find(f"{ns}Table") or ws.find("Table")
            if table is None:
                continue
            rows = table.findall(f"{ns}Row") or table.findall("Row")
            for row in rows:
                row_data = []
                cells = row.findall(f"{ns}Cell") or row.findall("Cell")
                for cell in cells:
                    data_tag = cell.find(f"{ns}Data") or cell.find("Data")
                    val = data_tag.text if data_tag is not None and data_tag.text else ""
                    row_data.append(val)
                if any(v.strip() for v in row_data):
                    all_data.append(row_data)

        if not all_data:
            return None
        return pd.DataFrame(all_data)

    except Exception as e:
        print(f"[XML-Rappi] No se pudo leer {ruta_archivo}: {e}")
        return None


def read_html_mhtml(ruta_archivo):
    """
    Lee archivos MHTML/HTML que Rappi genera al descargar
    reportes directamente desde el navegador.
    """
    try:
        tablas = pd.read_html(ruta_archivo, header=None, flavor="lxml")
        for tabla in tablas:
            if tabla is not None and not tabla.empty and len(tabla.columns) > 2:
                return tabla
        return None
    except Exception:
        pass

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
        print(f"[HTML-Rappi] No se pudo leer {ruta_archivo}: {e}")

    return None


def procesar_excel_rappi(rutas, meses_seleccionados=None):
    dataframes = []

    for ruta in rutas:
        # 🔒 LECTURA ROBUSTA: intenta openpyxl (xlsx) y luego xlrd (xls)
        hojas = None
        try:
            hojas = pd.read_excel(
                ruta,
                sheet_name=None,
                header=None,
                engine="openpyxl"
            )
        except Exception:
            pass

        if hojas is None:
            try:
                hojas = pd.read_excel(
                    ruta,
                    sheet_name=None,
                    header=None,
                    engine="xlrd"
                )
            except Exception:
                pass

        # INTENTO 3: XML-SPREADSHEET (Rappi descarga directa)
        if hojas is None:
            df_xml = read_xml_spreadsheet(ruta)
            if df_xml is not None:
                hojas = {"Sheet1": df_xml}

        # INTENTO 4: HTML/MHTML (Rappi descarga desde navegador)
        if hojas is None:
            df_html = read_html_mhtml(ruta)
            if df_html is not None:
                hojas = {"Sheet1": df_html}

        if hojas is None:
            print(f"⚠️ No se pudo leer el archivo Rappi: {ruta}")
            continue

        for nombre_hoja, df_raw in hojas.items():

            if df_raw is None or df_raw.empty:
                continue

            fila_header = encontrar_fila_header(df_raw)
            if fila_header is None:
                continue

            hoja_df = None
            try:
                hoja_df = pd.read_excel(
                    ruta,
                    sheet_name=nombre_hoja,
                    header=fila_header,
                    engine="openpyxl"
                )
            except Exception:
                pass

            if hoja_df is None:
                try:
                    hoja_df = pd.read_excel(
                        ruta,
                        sheet_name=nombre_hoja,
                        header=fila_header,
                        engine="xlrd"
                    )
                except Exception:
                    pass

            if hoja_df is None:
                continue

            # 🔒 ELIMINAR DUPLICADOS INVISIBLES
            hoja_df = hoja_df.loc[:, ~hoja_df.columns.duplicated()]

            # ==========================
            # RENOMBRADO
            # ==========================
            nuevas_columnas = {}
            for col in hoja_df.columns:
                col_norm = normalizar_texto(col)
                for clave, nombre_final in COLUMNAS_OBJETIVO.items():
                    if clave in col_norm:
                        nuevas_columnas[col] = nombre_final

            hoja_df = hoja_df.rename(columns=nuevas_columnas)

            # ==========================
            # COLUMNAS PRESENTES
            # ==========================
            columnas_presentes = [
                c for c in COLUMNAS_OBJETIVO.values()
                if c in hoja_df.columns
            ]

            if not columnas_presentes:
                continue

            hoja_df = hoja_df[columnas_presentes].copy()

            # ==========================
            # LIMPIEZA
            # ==========================
            if "Estado de la orden" in hoja_df.columns:
                hoja_df["Estado de la orden"] = (
                    hoja_df["Estado de la orden"]
                    .astype(str)
                    .str.lower()
                    .str.strip()
                )

                hoja_df = hoja_df[
                    hoja_df["Estado de la orden"].isin(
                        ["pending_review", "finished", "confirmed"]
                    )
                ]

            if "Venta Bruta" in hoja_df.columns:
                hoja_df["Venta Bruta"] = pd.to_numeric(
                    hoja_df["Venta Bruta"], errors="coerce"
                )
                hoja_df = hoja_df[hoja_df["Venta Bruta"] > 0]

            hoja_df["Aplicativo"] = "Rappi"

            hoja_df.reset_index(drop=True, inplace=True)
            dataframes.append(hoja_df)

    if not dataframes:
        raise Exception("No se encontraron datos válidos en Rappi")

    df_final = pd.concat(dataframes, ignore_index=True)

    # 🔒 BLINDAJE FINAL
    df_final = df_final.loc[:, ~df_final.columns.duplicated()]
    df_final.reset_index(drop=True, inplace=True)

    # ==========================
    # FILTRO POR MESES SELECCIONADOS
    # ==========================
    if meses_seleccionados and "Fecha de creación orden" in df_final.columns:
        fechas_parsed = pd.to_datetime(
            df_final["Fecha de creación orden"], errors="coerce"
        )
        mask_mes = fechas_parsed.dt.month.isin(meses_seleccionados)
        # Conservar filas sin fecha válida (no las descartamos)
        df_final = df_final[mask_mes | fechas_parsed.isna()].copy()
        df_final.reset_index(drop=True, inplace=True)

    fecha_inicio, fecha_fin = obtener_fechas_resumen(rutas[0])

    return df_final, fecha_inicio, fecha_fin
