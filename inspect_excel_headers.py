import pandas as pd
import unicodedata
import os

def normalizar_texto(texto):
    if pd.isna(texto):
        return ""
    texto = str(texto).lower().strip()
    texto = unicodedata.normalize("NFD", texto)
    texto = "".join(c for c in texto if unicodedata.category(c) != "Mn")
    return texto

def check_magic_bytes(filepath):
    try:
        with open(filepath, 'rb') as f:
            header = f.read(10)
        print(f"Magic bytes for {os.path.basename(filepath)}: {header}")
    except Exception as e:
        print(f"Error reading bytes: {e}")

def inspect_file(filepath, engine=None):
    print(f"\n--- INSPECTING: {os.path.basename(filepath)} ---")
    
    check_magic_bytes(filepath)

    try:
        # Try reading with pandas
        if engine:
            df = pd.read_excel(filepath, header=None, nrows=10, engine=engine)
        else:
            df = pd.read_excel(filepath, header=None, nrows=10)
            
        print("RAW CONTENT (First 10 rows):")
        print(df)
        
        print("\nNORMALIZED CONTENT:")
        for i, row in df.iterrows():
            norm_row = [normalizar_texto(c) for c in row]
            print(f"Row {i}: {norm_row}")
    except Exception as e:
        print(f"ERROR reading file with pandas (engine={engine}): {e}")
        
    # If pandas fails, try reading as text (maybe it's HTML or CSV)
    try:
        with open(filepath, 'r', encoding='utf-8', errors='ignore') as f:
            head = [next(f) for _ in range(5)]
        print("\nTEXT CONTENT (First 5 lines):")
        for line in head:
            print(line.strip())
    except Exception as e:
        print(f"Error reading as text: {e}")

# Paths
path_span_xlsx = r"c:/Users/User/Documents/RappiReportes_version1/archivoExcel_PedidosYa_Español/orderDetails (4).xlsx"
path_span_xls = r"c:/Users/User/Documents/RappiReportes_version1/archivoExcel_PedidosYa_Español/invoice-9601391391.xls"

print("Checking Spanish XLSX:")
inspect_file(path_span_xlsx, engine='openpyxl')

print("\nChecking Spanish XLS:")
# standard pandas behavior for .xls uses xlrd usually
inspect_file(path_span_xls) 
