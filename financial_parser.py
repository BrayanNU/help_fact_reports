"""
financial_parser.py
Parser centralizado y robusto para montos monetarios financieros.

REGLAS:
    "-S/.-248.82"  =>  Decimal("-248.82")
    "-S/.6.23"     =>  Decimal("-6.23")
    "+S/.13.95"    =>  Decimal("13.95")
    "S/.664.70"    =>  Decimal("664.70")
    "S/.0.00"      =>  Decimal("0.00")
    -248.82        =>  Decimal("-248.82")

PRINCIPIO:
    1. Detectar signo PRIMERO.
    2. Limpiar simbolos monetarios.
    3. Convertir a Decimal (no float).
"""

import re
from decimal import Decimal, ROUND_HALF_UP, InvalidOperation
import math


def parse_monto(valor) -> Decimal:
    """Parsea un monto financiero a Decimal con signo correcto."""

    if valor is None:
        return Decimal("0.00")

    # Caso 1: numerico nativo (int / float)
    if isinstance(valor, (int, float)):
        if isinstance(valor, float) and math.isnan(valor):
            return Decimal("0.00")
        try:
            return Decimal(str(valor)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
        except InvalidOperation:
            return Decimal("0.00")

    # Caso 2: string
    texto = str(valor).strip()

    if not texto or texto.lower() in ("nan", "none", "null", ""):
        return Decimal("0.00")

    # PASO 1: Detectar signo desde el TEXTO ORIGINAL (antes de limpiar)
    signo = Decimal("1")
    for ch in texto:
        if ch == "-":
            signo = Decimal("-1")
            break
        elif ch == "+":
            signo = Decimal("1")
            break
        elif ch.isdigit():
            break

    # PASO 2: Eliminar prefijos monetarios conocidos PRIMERO
    # Esto evita que "S/.248.82" quede como ".248.82" al limpiar
    prefijos = ["S/.", "S/", "$", "EUR", "PEN", "USD", "MXN"]
    texto_limpio = texto
    for pref in prefijos:
        texto_limpio = texto_limpio.replace(pref, " ")

    # PASO 3: Eliminar todo excepto digitos, punto y coma
    limpio = re.sub(r"[^\d.,]", "", texto_limpio).strip()

    # Eliminar puntos/comas iniciales residuales
    limpio = limpio.lstrip(".,")

    if not limpio:
        return Decimal("0.00")

    # PASO 4: Normalizar separadores de decimales
    if "," in limpio and "." in limpio:
        pos_coma = limpio.rindex(",")
        pos_punto = limpio.rindex(".")
        if pos_coma > pos_punto:
            # Formato europeo: 1.234,56 => 1234.56
            limpio = limpio.replace(".", "").replace(",", ".")
        else:
            # Formato US: 1,234.56 => 1234.56
            limpio = limpio.replace(",", "")
    elif "," in limpio:
        partes = limpio.split(",")
        if len(partes) == 2 and len(partes[1]) <= 2:
            limpio = limpio.replace(",", ".")
        else:
            limpio = limpio.replace(",", "")

    # PASO 5: Convertir a Decimal con el signo detectado
    try:
        valor_abs = Decimal(limpio)
        return (valor_abs * signo).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    except InvalidOperation:
        return Decimal("0.00")


def parse_monto_seguro(valor, nombre_campo: str = "campo") -> Decimal:
    """Wrapper con logging de errores."""
    try:
        return parse_monto(valor)
    except Exception as e:
        print(f"[PARSER ERROR] {nombre_campo}='{valor}' => {e}")
        return Decimal("0.00")


def sumar_montos(*valores) -> Decimal:
    """Suma montos como Decimal sin perdida de precision."""
    total = Decimal("0.00")
    for v in valores:
        total += v if isinstance(v, Decimal) else parse_monto(v)
    return total.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)


def validar_finanzas(finanzas_rappi: dict, finanzas_pedidosya: dict) -> bool:
    """Valida consistencia de datos financieros antes de generar el PDF."""
    SUCURSALES = [
        "Astrobuns | Smashburgers",
        "Incheon | Korean Fried Chicken",
        "Chickibuns"
    ]
    ok = True

    for suc in SUCURSALES:
        for app, grupo in [("Rappi", finanzas_rappi), ("PedidosYa", finanzas_pedidosya)]:
            if not grupo or suc not in grupo:
                continue
            data = grupo[suc]

            total = float(data.get("total", 0))
            cantidad = data.get("cantidad_pedidos", 0)
            comp = float(data.get("compensaciones", 0))
            promos = float(data.get("promociones_articulos", 0))

            if total < 0:
                print(f"[VALIDACION ERR] [{suc} {app}] total={total:.2f} es NEGATIVO.")
                ok = False

            if total > 0 and cantidad == 0:
                print(f"[VALIDACION WARN] [{suc} {app}] total={total:.2f} pero 0 pedidos.")

            if comp > 0:
                print(f"[VALIDACION WARN] [{suc} {app}] compensaciones={comp:.2f} es POSITIVO.")

            if abs(promos) > total > 0:
                print(f"[VALIDACION WARN] [{suc} {app}] |promos| > total.")

    if ok:
        print("[VALIDACION OK] Datos financieros consistentes.")
    return ok
