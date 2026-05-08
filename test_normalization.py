import sys
sys.path.insert(0, r"c:/Users/User/Documents/RappiReportes_version1")

from excel_processorPedidosYa_DetallesVentas import normalizar_texto

test_name = "Incheon - Comida Coreana"
normalized = normalizar_texto(test_name)

print(f"Original: '{test_name}'")
print(f"Normalized: '{normalized}'")
print(f"Length: {len(normalized)}")
print(f"Repr: {repr(normalized)}")
print(f"\nCharacter breakdown:")
for i, char in enumerate(normalized):
    print(f"  {i}: '{char}' (ord={ord(char)})")
