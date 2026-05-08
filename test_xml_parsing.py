import xml.etree.ElementTree as ET
import os

EXTRACT_PATH = r"c:/Users/User/Documents/RappiReportes_version1/temp_extract"
SHARED_STRINGS = os.path.join(EXTRACT_PATH, "xl", "sharedStrings.xml")
SHEET1 = os.path.join(EXTRACT_PATH, "xl", "worksheets", "sheet1.xml")

def parse_shared_strings():
    strings = []
    if not os.path.exists(SHARED_STRINGS):
        return []
    
    tree = ET.parse(SHARED_STRINGS)
    root = tree.getroot()
    # Namespace handling is tricky, simpler to ignore or use wildcards depending on implementation
    # Generally: {http://schemas.openxmlformats.org/spreadsheetml/2006/main}sst
    
    # Iterate all <t> elements inside <si>
    # Note: sometimes <t> is nested in <r> for rich text.
    namespace = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
    
    for si in root.findall(f".//{namespace}si"):
        # Text can be in <t> or <r><t>
        text_parts = []
        for t in si.findall(f".//{namespace}t"):
            if t.text:
                text_parts.append(t.text)
        strings.append("".join(text_parts))
        
    return strings

def parse_sheet(shared_strings):
    tree = ET.parse(SHEET1)
    root = tree.getroot()
    namespace = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
    
    data = []
    
    rows = root.findall(f".//{namespace}row")
    print(f"Found {len(rows)} rows.")
    
    for i, row in enumerate(rows[:20]): # First 20 rows
        row_data = []
        cells = row.findall(f"{namespace}c")
        
        # Sort by column index (r attribute) to be safe? 
        # Actually usually they are in order but might be sparse.
        # For this test, just appending is fine to see content.
        
        for cell in cells:
            val = ""
            cell_type = cell.get("t")
            v_tag = cell.find(f"{namespace}v")
            
            if v_tag is not None:
                v_text = v_tag.text
                if cell_type == "s": # Shared string
                    idx = int(v_text)
                    if idx < len(shared_strings):
                        val = shared_strings[idx]
                    else:
                        val = f"ERR:{idx}"
                else:
                    val = v_text
            
            # Check for inlineStr
            is_tag = cell.find(f"{namespace}is")
            if is_tag:
                 t_tag = is_tag.find(f"{namespace}t")
                 if t_tag is not None:
                     val = t_tag.text

            row_data.append(val)
        data.append(row_data)
        
    return data

try:
    print("Parsing shared strings...")
    strs = parse_shared_strings()
    print(f"Loaded {len(strs)} shared strings.")
    
    print("Parsing sheet...")
    rows = parse_sheet(strs)
    
    print("\n--- RECONSTRUCTED DATA ---")
    for r in rows:
        print(r)
        
except Exception as e:
    print(f"ERROR: {e}")
