import sys
sys.path.insert(0, r"c:/Users/User/Documents/RappiReportes_version1")

import pandas as pd
from excel_processorPedidosYa_DetallesVentas import normalizar_texto, obtener_sucursal_base, aplanar_columnas, encontrar_columna_por_texto

ruta = r"c:/Users/User/Documents/RappiReportes_version1/archivoExcel_PedidosYa_Español/invoice-9601395690.xls"

with open("c:/Users/User/Documents/RappiReportes_version1/debug_detailed.txt", "w", encoding="utf-8") as f:
    f.write("DETAILED DIAGNOSTIC\n")
    f.write("=" * 80 + "\n\n")
    
    # Read file
    df = pd.read_excel(ruta, engine="openpyxl", header=[0,1])
    aplanar_columnas(df)
    
    f.write(f"Total rows: {len(df)}\n\n")
    
    # Check columns
    col_estado = encontrar_columna_por_texto(df, "order status") or encontrar_columna_por_texto(df, "estado")
    col_sucursal = encontrar_columna_por_texto(df, "sucursal")
    col_total = encontrar_columna_por_texto(df, "monto de la venta")
    
    f.write(f"Column 'order status' or 'estado': {col_estado}\n")
    f.write(f"Column 'sucursal': {col_sucursal}\n")
    f.write(f"Column 'monto de la venta': {col_total}\n\n")
    
    if col_sucursal:
        f.write("First 5 sucursal values:\n")
        for i, val in enumerate(df[col_sucursal].head(5)):
            matched = obtener_sucursal_base(val)
            f.write(f"  {i}: '{val}' -> {matched}\n")
        f.write("\n")
    
    if col_estado:
        f.write("First 5 estado values:\n")
        for i, val in enumerate(df[col_estado].head(5)):
            f.write(f"  {i}: '{val}'\n")
        f.write("\n")
        
        # Try filtering
        df_copy = df.copy()
        df_copy[col_estado] = df_copy[col_estado].astype(str).str.lower().str.strip()
        df_filtered = df_copy[df_copy[col_estado] == "delivered"]
        f.write(f"Rows after filtering for 'delivered': {len(df_filtered)}\n\n")
    else:
        f.write("No 'order status' column found - will process all rows\n\n")
    
    if col_total:
        f.write("First 5 'monto de la venta' values:\n")
        for i, val in enumerate(df[col_total].head(5)):
            f.write(f"  {i}: {val} (type: {type(val).__name__})\n")

print("✓ Diagnostic written to: debug_detailed.txt")
