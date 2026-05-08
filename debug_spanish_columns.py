import pandas as pd

def aplanar_columnas(df):
    if isinstance(df.columns, pd.MultiIndex):
        df.columns = [
            " ".join(str(x).strip() for x in col if x and str(x) != "nan").lower()
            for col in df.columns
        ]
    else:
        df.columns = [str(c).strip().lower() for c in df.columns]

def inspect_file(path, output_file):
    output_file.write(f"\n{'='*80}\n")
    output_file.write(f"INSPECTING: {path}\n")
    output_file.write(f"{'='*80}\n\n")
    
    try:
        # Mimic production code: header=[0,1]
        df = pd.read_excel(path, engine="openpyxl", header=[0,1])
        output_file.write("✓ Read via OPENPYXL (header=[0,1])\n\n")
    except Exception as e:
        output_file.write(f"✗ OPENPYXL failed: {e}\n")
        return
        
    aplanar_columnas(df)
    
    output_file.write("FLATTENED COLUMNS:\n")
    output_file.write("-" * 80 + "\n")
    for i, col in enumerate(df.columns):
        output_file.write(f"{i:3d}: {col}\n")
    
    output_file.write("\n" + "=" * 80 + "\n")
    output_file.write("FIRST 3 DATA ROWS:\n")
    output_file.write("=" * 80 + "\n")
    for idx, row in df.head(3).iterrows():
        output_file.write(f"\nRow {idx}:\n")
        for col in df.columns:
            val = row[col]
            output_file.write(f"  {col}: {val} (type: {type(val).__name__})\n")

# Write to file
with open("c:/Users/User/Documents/RappiReportes_version1/spanish_columns_debug.txt", "w", encoding="utf-8") as f:
    inspect_file(r"c:/Users/User/Documents/RappiReportes_version1/archivoExcel_PedidosYa_Español/invoice-9601395690.xls", f)
    inspect_file(r"c:/Users/User/Documents/RappiReportes_version1/archivoExcel_PedidosYa_Español/orderDetails (5).xlsx", f)

print("✓ Debug output written to: spanish_columns_debug.txt")
