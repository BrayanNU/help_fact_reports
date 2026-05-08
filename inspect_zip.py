import zipfile
import os

zip_path = r"c:/Users/User/Documents/RappiReportes_version1/archivoExcel_PedidosYa_Español/orderDetails (4).xlsx"
extract_path = r"c:/Users/User/Documents/RappiReportes_version1/temp_extract"

print(f"--- EXTRACTING {zip_path} ---")

if not os.path.exists(extract_path):
    os.makedirs(extract_path)

try:
    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
        zip_ref.extractall(extract_path)
    print("SUCCESS: Extracted zip.")
    print("Contents:")
    for root, dirs, files in os.walk(extract_path):
        for name in files:
            print(os.path.join(root, name))
            
    # Check sharedStrings.xml and sheet1.xml
    sheet1 = os.path.join(extract_path, "xl", "worksheets", "sheet1.xml")
    if os.path.exists(sheet1):
        print("\n--- HEADER OF sheet1.xml ---")
        with open(sheet1, "r", encoding="utf-8") as f:
            print(f.read(500)) # First 500 chars
            
    # Check styles.xml (the culprit?)
    styles = os.path.join(extract_path, "xl", "styles.xml")
    if os.path.exists(styles):
         print("\n--- Styles exists (likely broken) ---")
         
except Exception as e:
    print(f"FAILED to unzip: {e}")
