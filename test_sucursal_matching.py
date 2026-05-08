import sys
sys.path.insert(0, r"c:/Users/User/Documents/RappiReportes_version1")

import pandas as pd
from excel_processorPedidosYa_DetallesVentas import normalizar_texto, obtener_sucursal_base, aplanar_columnas, encontrar_columna_por_texto

# Test the sucursal matching
test_names = [
    "Incheon - Comida Coreana",
    "Astrobuns Smash Burgers",
    "incheon - comida coreana",
    "astrobuns smash burgers"
]

print("Testing sucursal name matching:")
print("=" * 80)
for name in test_names:
    normalized = normalizar_texto(name)
    result = obtener_sucursal_base(name)
    print(f"Input: '{name}'")
    print(f"  Normalized: '{normalized}'")
    print(f"  Result: {result}")
    print()

# Test reading the Spanish file
print("\n" + "=" * 80)
print("Testing file reading:")
print("=" * 80)

ruta = r"c:/Users/User/Documents/RappiReportes_version1/archivoExcel_PedidosYa_Español/invoice-9601395690.xls"
df = pd.read_excel(ruta, engine="openpyxl", header=[0,1])
aplanar_columnas(df)

print(f"Rows in dataframe: {len(df)}")
print(f"\nColumns: {list(df.columns)[:5]}...")

col_sucursal = encontrar_columna_por_texto(df, "sucursal")
print(f"\nFound 'sucursal' column: {col_sucursal}")

if col_sucursal:
    print(f"\nFirst 5 sucursal values:")
    for i, val in enumerate(df[col_sucursal].head(5)):
        normalized = normalizar_texto(val)
        matched = obtener_sucursal_base(val)
        print(f"  {i}: '{val}' -> normalized: '{normalized}' -> matched: {matched}")
