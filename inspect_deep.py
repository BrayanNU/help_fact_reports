import pandas as pd
import unicodedata

def inspect_file(path, desc):
    print(f"\n--- INSPECTING {desc}: {path} ---")
    try:
        # Read with default engine
        df = pd.read_excel(path, header=None, nrows=20)
        
        # Print specific rows raw
        print("ROWS 10-15 RAW:")
        for i in range(10, min(16, len(df))):
            row = df.iloc[i].tolist()
            # simple print
            print(f"ROW {i}: {row}")
            
    except Exception as e:
        print(f"ERROR: {e}")

f1 = r"c:/Users/User/Documents/RappiReportes_version1/archivoExcel_PedidosYa_Español/invoice-9601391391.xls"
f2 = r"c:/Users/User/Documents/RappiReportes_version1/archivoExcel_PedidosYa_Español/orderDetails (5).xlsx"

inspect_file(f1, "INVOICE XLS")
inspect_file(f2, "ORDER DETAILS 5 XLSX")
