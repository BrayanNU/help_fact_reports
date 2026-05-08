import pandas as pd
import unicodedata
from collections import defaultdict

# ==========================================
# HELPERS BÁSICOS
# ==========================================

def normalizar_texto(texto):
    if pd.isna(texto):
        return ""
    texto = str(texto).lower().strip()
    texto = unicodedata.normalize("NFD", texto)
    return "".join(c for c in texto if unicodedata.category(c) != "Mn")


def aplanar_columnas(df):
    """
    Convierte columnas con múltiples encabezados de Excel
    en una sola fila de texto plano y en minúsculas.
    """
    if isinstance(df.columns, pd.MultiIndex):
        df.columns = [
            " ".join(
                str(x).strip()
                for x in col
                if x and str(x).lower() != "nan"
            ).lower()
            for col in df.columns
        ]
    else:
        df.columns = [str(c).strip().lower() for c in df.columns]


def encontrar_columna_por_texto(df, texto):
    """
    Busca una columna cuyo nombre contenga el texto indicado,
    ignorando mayúsculas, tildes y espacios.
    """
    texto = normalizar_texto(texto)
    for col in df.columns:
        if texto in normalizar_texto(col):
            return col
    return None


def obtener_valor_seguro(valor):
    """
    Convierte cualquier valor a float.
    Si no se puede convertir, devuelve 0.0
    """
    try:
        return float(pd.to_numeric(valor, errors="coerce")) if pd.notna(valor) else 0.0
    except Exception:
        return 0.0


# ==========================================
# SUCURSALES RAPPI
# ==========================================

def obtener_sucursal_base(nombre):
    """
    Normaliza el nombre de la tienda del Excel
    y lo traduce a un nombre estándar interno.
    """
    nombre = normalizar_texto(nombre)
    SUCURSALES = {
        "astrobuns - comas": "Astrobuns | Smashburgers",
        "incheon comas": "Incheon | Korean Fried Chicken"
    }
    return SUCURSALES.get(nombre)


# ==========================================
# FUNCIÓN CENTRAL RAPPI
# ==========================================

def procesar_finanzas_rappi(rutas, fecha_inicio=None, fecha_fin=None):

    resultados = defaultdict(lambda: {
        "total": 0.0,                   # Venta bruta de pedidos
        "cantidad_pedidos": 0,          # Número de pedidos válidos
        "promociones_articulos": 0.0,   # Descuentos por producto
        "descuentos_fugaces": 0.0,
        "descuentos_cancelaciones": 0.0,
        "descuentos_reclamos": 0.0,
        "reintegros_tienda": 0.0,
        "compensaciones": 0.0           # Compensaciones Rappi
    })

    for ruta in rutas:
        print(f"\n📂 Procesando archivo: {ruta}")

        try:
            df = pd.read_excel(
                ruta,
                sheet_name="Detalle",
                header=[0, 1],
                engine="openpyxl"
            )
        except Exception as e:
            print(f"❌ No se pudo leer el archivo: {e}")
            continue

        if df is None or df.empty:
            print("⚠️ Archivo vacío, se omite")
            continue

        aplanar_columnas(df)

        # ==========================
        # BÚSQUEDA DE COLUMNAS
        # ==========================
        col_estado_orden = encontrar_columna_por_texto(df, "estado de la orden")
        col_tipo_transaccion = encontrar_columna_por_texto(df, "tipo de transaccion")
        col_sucursal = encontrar_columna_por_texto(df, "nombre de la tienda")
        col_total = encontrar_columna_por_texto(df, "venta bruta")
        col_promos = encontrar_columna_por_texto(df, "descuento de producto")
        col_compensaciones = encontrar_columna_por_texto(df, "compensaciones")

        print("🔍 Columnas detectadas:")
        print(f"   - Estado orden: {col_estado_orden}")
        print(f"   - Tipo transacción: {col_tipo_transaccion}")
        print(f"   - Sucursal: {col_sucursal}")
        print(f"   - Venta bruta: {col_total}")
        print(f"   - Promos: {col_promos}")
        print(f"   - Compensaciones: {col_compensaciones}")

        if not (col_sucursal and col_total):
            print("❌ Faltan columnas críticas, se omite archivo")
            continue

        # ==========================
        # RECORRIDO FILA POR FILA
        # ==========================
        for _, fila in df.iterrows():

            sucursal = obtener_sucursal_base(fila[col_sucursal])
            if not sucursal:
                continue

            # --------------------------
            # PEDIDOS / PROMOCIONES
            # --------------------------
            if col_estado_orden:
                estado = str(fila[col_estado_orden]).lower().strip()

                if estado in ["pending_review", "finished"]:
                    monto = obtener_valor_seguro(fila[col_total])

                    resultados[sucursal]["cantidad_pedidos"] += 1
                    resultados[sucursal]["total"] += monto

                    print(
                        f"🧾 PEDIDO | {sucursal} | "
                        f"Estado={estado} | Venta={monto}"
                    )

                    if col_promos:
                        promo = obtener_valor_seguro(fila[col_promos])
                        resultados[sucursal]["promociones_articulos"] += promo

                        if promo != 0:
                            print(
                                f"   🎁 Promo aplicada: {promo}"
                            )

            # --------------------------
            # COMPENSACIONES
            # --------------------------
            if col_tipo_transaccion and col_compensaciones:
                tipo_transaccion = str(fila[col_tipo_transaccion]).strip().upper()

                if tipo_transaccion == "COMPENSACIÓN":
                    comp = obtener_valor_seguro(fila[col_compensaciones])
                    resultados[sucursal]["compensaciones"] += comp

                    print(
                        f"💸 COMPENSACIÓN | {sucursal} | Monto={comp}"
                    )

    # ==========================
    # BLINDAJE FINAL
    # ==========================
    for suc in resultados:
        for k in resultados[suc]:
            if resultados[suc][k] is None:
                resultados[suc][k] = 0.0

    print("\n✅ Procesamiento finalizado")
    return resultados
