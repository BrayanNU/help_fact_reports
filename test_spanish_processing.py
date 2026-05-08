import sys
sys.path.insert(0, r"c:/Users/User/Documents/RappiReportes_version1")

from excel_processorPedidosYa_DetallesVentas import procesar_finanzas_pedidosya

# Test with Spanish files
rutas = [
    r"c:/Users/User/Documents/RappiReportes_version1/archivoExcel_PedidosYa_Español/invoice-9601395690.xls"
]

print("Testing Spanish file processing...")
print("=" * 80)

resultados = procesar_finanzas_pedidosya(rutas)

# Write to file
with open("c:/Users/User/Documents/RappiReportes_version1/test_results.txt", "w", encoding="utf-8") as f:
    f.write("RESULTADOS DEL PROCESAMIENTO:\n")
    f.write("=" * 80 + "\n\n")
    
    for sucursal, datos in resultados.items():
        f.write(f"{sucursal}:\n")
        f.write(f"  Total: S/.{datos['total']:.2f}\n")
        f.write(f"  Cantidad pedidos: {datos['cantidad_pedidos']}\n")
        f.write(f"  Promociones: -S/.{datos['promociones_articulos']:.2f}\n")
        f.write(f"  Descuentos fugaces: -S/.{datos['descuentos_fugaces']:.2f}\n")
        f.write(f"  Descuentos cancelaciones: -S/.{datos['descuentos_cancelaciones']:.2f}\n")
        f.write(f"  Descuentos reclamos: -S/.{datos['descuentos_reclamos']:.2f}\n")
        f.write(f"  Reintegros tienda: +S/.{datos['reintegros_tienda']:.2f}\n")
        f.write("\n")

print("\n✓ Results written to: test_results.txt")
