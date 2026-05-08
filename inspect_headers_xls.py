import pandas as pd
import unicodedata

def inspect_xls():
    path = r"c:/Users/User/Documents/RappiReportes_version1/archivoExcel_PedidosYa_Español/invoice-9601391391.xls"
    print(f"--- INSPECTING {path} ---")
    try:
        df = pd.read_excel(path, header=None, nrows=20)
        for i, row in df.iterrows():
            vals = [str(c).strip() for c in row if pd.notna(c) and str(c).strip()]
            print(f"ROW {i}: {vals}")
    except Exception as e:
        print(f"ERROR: {e}")

inspect_xls()
