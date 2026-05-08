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


def procesar_excel_rappi(rutas):
    dataframes = []

    for ruta in rutas:
        try:
            # 🔒 LECTURA ROBUSTA (evita error stylesheet/XML)
            hojas = pd.read_excel(
                ruta,
                sheet_name=None,
                header=None,
                engine="openpyxl"
            )
        except Exception:
            continue

        for nombre_hoja, df_raw in hojas.items():

            if df_raw is None or df_raw.empty:
                continue

            fila_header = encontrar_fila_header(df_raw)
            if fila_header is None:
                continue

            try:
                hoja_df = pd.read_excel(
                    ruta,
                    sheet_name=nombre_hoja,
                    header=fila_header,
                    engine="openpyxl"
                )
            except Exception:
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

    fecha_inicio, fecha_fin = obtener_fechas_resumen(rutas[0])

    return df_final, fecha_inicio, fecha_fin
