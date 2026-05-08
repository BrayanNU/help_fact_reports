import pandas as pd
import unicodedata
import warnings
import os

warnings.simplefilter(action='ignore', category=UserWarning)

def inspect_file(filepath, description, engine=None, try_html=False):
    print(f"\n{'-'*20}\n{description}: {os.path.basename(filepath)}\n{'-'*20}")
    
    if not os.path.exists(filepath):
        print("File does not exist!")
        return

    try:
        if try_html:
            print("Attempting read_html...")
            dfs = pd.read_html(filepath)
            df = dfs[0]
        else:
            print(f"Attempting read_excel (engine={engine})...")
            df = pd.read_excel(filepath, header=None, nrows=10, engine=engine)
            
        print("SUCCESS! First 5 rows normalized:")
        for i, row in df.head(5).iterrows():
             print([str(c).lower().strip() for c in row if pd.notna(c)])
             
    except Exception as e:
        print(f"FAILED: {e}")

# Paths
path_eng = r"c:/Users/User/Documents/RappiReportes_version1/archivoExcel_PedidosYa_Ingles/orderDetails (3).xlsx"
path_span_xls = r"c:/Users/User/Documents/RappiReportes_version1/archivoExcel_PedidosYa_Español/invoice-9601391391.xls"
path_span_xlsx = r"c:/Users/User/Documents/RappiReportes_version1/archivoExcel_PedidosYa_Español/orderDetails (4).xlsx"

# 1. Baseline
inspect_file(path_eng, "ENGLISH XLSX (Baseline)")

# 2. Spanish XLS - Try default, then xlrd, then HTML
inspect_file(path_span_xls, "SPANISH XLS (Default)")
inspect_file(path_span_xls, "SPANISH XLS (HTML)", try_html=True)

# 3. Spanish XLSX - Try default
inspect_file(path_span_xlsx, "SPANISH XLSX (Default)")
