"""
test_financial_parser.py
Tests unitarios para el parser financiero y calcular_saldo_total.

Ejecutar con: python test_financial_parser.py
"""

import sys
import os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from decimal import Decimal
from financial_parser import parse_monto, sumar_montos

# ─────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────
passed = 0
failed = 0

def test(nombre, resultado, esperado):
    global passed, failed
    if resultado == esperado:
        print(f"  PASS  {nombre}")
        passed += 1
    else:
        print(f"  FAIL  {nombre}")
        print(f"         Esperado: {esperado}")
        print(f"         Obtenido: {resultado}")
        failed += 1

# ─────────────────────────────────────────────────────────────
# GRUPO 1: Casos del enunciado
# ─────────────────────────────────────────────────────────────
print("\n[Grupo 1] Casos del enunciado")

test("doble negacion: -S/.-248.82",
     parse_monto("-S/.-248.82"), Decimal("-248.82"))

test("negativo simple: -S/.6.23",
     parse_monto("-S/.6.23"), Decimal("-6.23"))

test("positivo explicito: +S/.13.95",
     parse_monto("+S/.13.95"), Decimal("13.95"))

test("sin signo: S/.664.70",
     parse_monto("S/.664.70"), Decimal("664.70"))

test("cero: S/.0.00",
     parse_monto("S/.0.00"), Decimal("0.00"))

# ─────────────────────────────────────────────────────────────
# GRUPO 2: Tipos numericos nativos
# ─────────────────────────────────────────────────────────────
print("\n[Grupo 2] Tipos numericos nativos")

test("float negativo: -248.82",
     parse_monto(-248.82), Decimal("-248.82"))

test("float positivo: 13.95",
     parse_monto(13.95), Decimal("13.95"))

test("int: 100",
     parse_monto(100), Decimal("100.00"))

test("float cero: 0.0",
     parse_monto(0.0), Decimal("0.00"))

# ─────────────────────────────────────────────────────────────
# GRUPO 3: Casos edge
# ─────────────────────────────────────────────────────────────
print("\n[Grupo 3] Casos edge")

test("nan string",
     parse_monto("nan"), Decimal("0.00"))

test("None",
     parse_monto(None), Decimal("0.00"))

test("vacio",
     parse_monto(""), Decimal("0.00"))

test("float nan",
     parse_monto(float("nan")), Decimal("0.00"))

test("formato europeo: 1.248,82",
     parse_monto("1.248,82"), Decimal("1248.82"))

test("numero grande: S/.12345.67",
     parse_monto("S/.12345.67"), Decimal("12345.67"))

# ─────────────────────────────────────────────────────────────
# GRUPO 4: Test de regresion - suma total del enunciado
# ─────────────────────────────────────────────────────────────
print("\n[Grupo 4] Regresion - calculo completo del enunciado")

montos = [
    parse_monto("828.90"),    # Astrobuns Rappi total
    parse_monto("-248.82"),   # Astrobuns Rappi promos
    parse_monto("-31.17"),    # Astrobuns Rappi compensaciones
    parse_monto("664.70"),    # Astrobuns PedidosYa total
    parse_monto("232.50"),    # Incheon Rappi total
    parse_monto("-61.28"),    # Incheon Rappi promos
    parse_monto("-26.22"),    # Incheon Rappi compensaciones
    parse_monto("194.20"),    # Incheon PedidosYa total
    parse_monto("-6.23"),     # Incheon PedidosYa cancelaciones
    parse_monto("287.30"),    # Chickibuns PedidosYa total
    parse_monto("-38.85"),    # Chickibuns PedidosYa fugaces
    parse_monto("-10.43"),    # Chickibuns PedidosYa reclamos
    parse_monto("+13.95"),    # Chickibuns PedidosYa reintegro
]
total = sum(montos, Decimal("0"))

test("suma total exacta = 1798.55",
     total, Decimal("1798.55"))

# ─────────────────────────────────────────────────────────────
# GRUPO 5: Test calcular_saldo_total con datos simulados
# ─────────────────────────────────────────────────────────────
print("\n[Grupo 5] calcular_saldo_total simulado")

from decimal import Decimal, ROUND_HALF_UP

def calcular_saldo_total_fixed(*diccionarios):
    saldo = Decimal("0")
    for grupo in diccionarios:
        if not grupo:
            continue
        for data in grupo.values():
            saldo += Decimal(str(data.get("total", 0)))
            saldo -= abs(Decimal(str(data.get("promociones_articulos", 0))))
            saldo -= abs(Decimal(str(data.get("descuentos_fugaces", 0))))
            saldo -= abs(Decimal(str(data.get("descuentos_cancelaciones", 0))))
            saldo -= abs(Decimal(str(data.get("descuentos_reclamos", 0))))
            saldo -= abs(Decimal(str(data.get("compensaciones", 0))))
            saldo += abs(Decimal(str(data.get("reintegros_tienda", 0))))
    return float(saldo.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))

finanzas_rappi = {
    "Astrobuns | Smashburgers": {
        "total": 828.90, "promociones_articulos": -248.82,
        "compensaciones": -31.17, "descuentos_fugaces": 0,
        "descuentos_cancelaciones": 0, "descuentos_reclamos": 0,
        "reintegros_tienda": 0
    },
    "Incheon | Korean Fried Chicken": {
        "total": 232.50, "promociones_articulos": -61.28,
        "compensaciones": -26.22, "descuentos_fugaces": 0,
        "descuentos_cancelaciones": 0, "descuentos_reclamos": 0,
        "reintegros_tienda": 0
    },
}

finanzas_pedidosya = {
    "Astrobuns | Smashburgers": {
        "total": 664.70, "promociones_articulos": 0,
        "compensaciones": 0, "descuentos_fugaces": 0,
        "descuentos_cancelaciones": 0, "descuentos_reclamos": 0,
        "reintegros_tienda": 0
    },
    "Incheon | Korean Fried Chicken": {
        "total": 194.20, "promociones_articulos": 0,
        "compensaciones": 0, "descuentos_fugaces": 0,
        "descuentos_cancelaciones": -6.23, "descuentos_reclamos": 0,
        "reintegros_tienda": 0
    },
    "Chickibuns": {
        "total": 287.30, "promociones_articulos": 0,
        "compensaciones": 0, "descuentos_fugaces": -38.85,
        "descuentos_cancelaciones": 0, "descuentos_reclamos": -10.43,
        "reintegros_tienda": 13.95
    },
}

resultado = calcular_saldo_total_fixed(finanzas_rappi, finanzas_pedidosya)
test("calcular_saldo_total = 1798.55",
     resultado, 1798.55)

# ─────────────────────────────────────────────────────────────
# RESUMEN
# ─────────────────────────────────────────────────────────────
print(f"\n{'='*50}")
print(f"  RESULTADOS: {passed} pasaron / {failed} fallaron")
print(f"{'='*50}")
if failed > 0:
    sys.exit(1)
