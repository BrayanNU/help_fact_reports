import pandas as pd
import zipfile
import xml.etree.ElementTree as ET

# Force pandas to show all columns
pd.set_option('display.max_columns', None)
pd.set_option('display.max_rows', None)
pd.set_option('display.width', 1000)

def aplanar_columnas(df):
    if isinstance(df.columns, pd.MultiIndex):
        df.columns = [
            " ".join(str(x).strip() for x in col if x and str(x) != "nan").lower()
            for col in df.columns
        ]
    else:
        df.columns = [str(c).strip().lower() for c in df.columns]

def inspect_file(path):
    print(f"--- INSPECTING {path} ---")
    try:
        # Mimic production code: header=[0,1]
        df = pd.read_excel(path, engine="openpyxl", header=[0,1])
        print("Read via OPENPYXL success (header=[0,1])")
    except:
        print("OPENPYXL failed...")
        return
        
    if df is not None:
        aplanar_columnas(df)
        print("\nFLATTENED COLUMNS:")
        for i, col in enumerate(df.columns):
            print(f"{i}: {col}")
            
        print("\nFirst 5 rows with flattened columns:")
        print(df.head(5))

# Inspect one of the Spanish files
inspect_file(r"c:/Users/User/Documents/RappiReportes_version1/archivoExcel_PedidosYa_Español/invoice-9601395690.xls")
inspect_file(r"c:/Users/User/Documents/RappiReportes_version1/archivoExcel_PedidosYa_Español/orderDetails (5).xlsx")
