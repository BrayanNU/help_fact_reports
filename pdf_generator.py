from reportlab.lib.pagesizes import A4
from reportlab.platypus import (
    Image, SimpleDocTemplate, Paragraph, Spacer,
    Table, TableStyle, PageBreak
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.lib.utils import simpleSplit
from datetime import datetime
import pandas as pd
import os
import math
from typing import Optional, Dict, Any, List, Union

from config_manager import load_footer_config, load_invoice_config


def safe_get(fila: pd.Series, columna: str, default: Any = "N/A"):
    if columna in fila and not pd.isna(fila[columna]):
        return fila[columna]
    return default


def sanitize_cell(value: Any) -> str:
    """Evita errores de ReportLab al renderizar tablas (como valores None)."""
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return ""
    return str(value)


def format_fecha_pdf(fecha: Any) -> str:
    """Formatea fecha para PDF, removiendo la hora si existe."""
    if fecha is None or pd.isna(fecha):
        return ""
    if isinstance(fecha, str):
        # Si es string con formato datetime, extraer solo la fecha
        if ' ' in fecha:
            return fecha.split(' ')[0]
        return fecha
    elif hasattr(fecha, 'strftime'):
        # Si es objeto datetime, formatear solo la fecha
        return fecha.strftime('%d/%m/%Y')
    return str(fecha)


def generar_pdf(
    df: pd.DataFrame,
    fecha_inicio: Any,
    fecha_fin: Any,
    finanzas_pedidosya: Optional[Dict],
    finanzas_rappi: Optional[Dict],
    nombre_pdf: str = "reporte_delivery.pdf"
):

    finanzas_pedidosya = finanzas_pedidosya or {}
    finanzas_rappi = finanzas_rappi or {}

    if df is None or df.empty:
        raise ValueError("El DataFrame está vacío o no contiene datos")

    df = df.loc[:, ~df.columns.duplicated()].reset_index(drop=True)

    # 🔑 respeta el orden real del archivo
    df["_orden"] = range(len(df))

    doc = SimpleDocTemplate(
        nombre_pdf,
        pagesize=A4,
        rightMargin=1 * cm,
        leftMargin=1 * cm,
        topMargin=1 * cm,
        bottomMargin=2.5 * cm
    )

    styles = getSampleStyleSheet()
    elementos = []
    invoice = load_invoice_config()

    fecha = invoice.get("fecha", datetime.now().strftime("%d/%m/%Y"))

    # ==========================
    # TÍTULO
    # ==========================
    elementos.append(
        Paragraph(
            "DETALLES FACTURA SERVICIO",
            ParagraphStyle(
                name="TituloFactura",
                fontSize=17,
                alignment=1,
                spaceAfter=15,
                fontName="Helvetica-Bold"
            )
        )
    )

    # ==========================
    # BLOQUE FACTURA
    # ==========================
    factura_texto = f"""
<b>{invoice.get('ciudad', '')}</b><br/>
A la atención de <b>{invoice.get('atencion', '')}</b><br/>
{invoice.get('direccion', '')}<br/>
{invoice.get('ubigeo', '')}<br/><br/>
<b>Fecha:</b> {fecha}<br/>
<b>N° RUC:</b> {invoice.get('ruc', '')}<br/>
<b>Nº de factura:</b> {invoice.get('numero_factura', '')}<br/>
<b>Customer ID:</b> {invoice.get('customer_id', '')}<br/><br/>
{invoice.get('mensaje', '')}
"""

    bloque_factura = Paragraph(
        factura_texto,
        ParagraphStyle(name="BloqueFactura", fontSize=10, leading=14)
    )

    # Reemplazar string por Spacer para evitar problemas de tipos de Pyre2/Reportlab
    logo: Union[Image, Spacer] = Spacer(1, 1)
    if invoice.get("logo_path") and str(invoice.get("logo_path")).strip() and os.path.exists(invoice["logo_path"]):
        logo = Image(invoice["logo_path"], width=5 * cm, height=2.5 * cm)

    header = Table([[bloque_factura, logo]], colWidths=[15 * cm, 3 * cm])
    header.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("ALIGN", (1, 0), (1, 0), "RIGHT"),
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 0),
        ("TOPPADDING", (0, 0), (-1, -1), 0),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
    ]))

    elementos.append(header)
    elementos.append(Spacer(1, 20))

    # ==========================
    # DETALLES FACTURAS
    # ==========================

    elementos.append(
        Paragraph(
            "DETALLES FACTURAS",
            ParagraphStyle(
                name="TituloFacturaDetalle",
                fontSize=15,
                fontName="Helvetica-Bold",
                spaceAfter=12
            )
        )
    )

    COMISION: float = 0.26

    total_comisiones: float = 0.0
    filas_factura: List[List[Union[str, Paragraph]]] = []

    SUCURSALES = [
        "Astrobuns | Smashburgers",
        "Incheon | Korean Fried Chicken"
    ]
    for sucursal in SUCURSALES:

        for app, grupo in [("Rappi", finanzas_rappi), ("PedidosYa", finanzas_pedidosya)]:

            if not grupo or sucursal not in grupo:
                continue

            data = grupo[sucursal]

            neto = round(
                float(data.get("total", 0))
                - abs(float(data.get("promociones_articulos", 0)))
                - abs(float(data.get("descuentos_fugaces", 0))),
                2
            )

            comision = round(neto * COMISION, 2)

            total_comisiones += comision

            filas_factura.append([
                f"{sucursal} {app} ({format_fecha_pdf(fecha_inicio)} en {format_fecha_pdf(fecha_fin)} incluye)"
            ])

            filas_factura.append([
                f"{data.get('cantidad_pedidos', 0)} pedidos por un total de S/.{neto:.2f}"
            ])

            filas_factura.append([
                "Comisión de entrega de la plataforma: 26.00%",
                f"S/.{comision:.2f}"
            ])

            filas_factura.append(["", ""])

    # ==========================
    # SUBTOTALES
    # ==========================

    subtotal = round(total_comisiones, 2)
    igv = round(subtotal * 0.18, 2)
    importe_total = round(subtotal + igv, 2)

    filas_factura += [
        ["", ""],
        ["Subtotal", f"S/.{subtotal:.2f}"],
        ["IGV (18%)", f"S/.{igv:.2f}"],
        ["Importe total de esta factura", f"S/.{importe_total:.2f}"],
        ["Importe a retener por el servicio en línea", f"S/.{importe_total:.2f}"],
        ["Importe pendiente", "S/.0.00"],
        ["", ""]
    ]

    # Sanitizar None values a str para evitar problemas en Reportlab Table
    filas_factura = [[sanitize_cell(c) for c in f] for f in filas_factura]

    tabla_factura = Table(filas_factura, colWidths=[11 * cm, 5 * cm])

    tabla_factura.setStyle(TableStyle([

        ("ALIGN", (1, 0), (1, -1), "RIGHT"),

        ("LINEABOVE", (0, -6), (-1, -6), 0.8, colors.black),
        ("LINEBELOW", (0, -1), (-1, -1), 0.8, colors.black),

        ("FONT", (0, -3), (-1, -3), "Helvetica-Bold"),
        ("FONT", (0, -2), (-1, -2), "Helvetica-Bold"),

        ("LEFTPADDING", (0, 0), (-1, -1), 2),
        ("RIGHTPADDING", (0, 0), (-1, -1), 2),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),

    ]))

    elementos.append(tabla_factura)
    elementos.append(Spacer(1, 12))

    # ==========================
    # DETALLES VENTAS
    # ==========================
    elementos.append(PageBreak()) 
    elementos.append(
        Paragraph(
            "DETALLES VENTAS",
            ParagraphStyle(
                name="TituloPedidos",
                fontSize=15,
                alignment=0,
                spaceAfter=15,
                fontName="Helvetica-Bold"
            )
        )
    )

    styles_small = ParagraphStyle(
        name="Small",
        fontSize=10,
        leading=14
    )

    def bloque_finanzas(nombre, app, data):

        fi = fecha_inicio
        ff = fecha_fin

        elementos.append(
            Paragraph(
                f"<b>{nombre} {app} ({format_fecha_pdf(fi)} en {format_fecha_pdf(ff)} incluye)</b>",
                styles_small
            )
        )

        filas = []

        filas.append([
            "Total",
            f"{data.get('cantidad_pedidos', 0)} pedidos por un total de S/.{data.get('total', 0):.2f}"
        ])

        filas.append([
            "Promociones en artículos",
            f"una cantidad de -S/.{data.get('promociones_articulos', 0):.2f}"
        ])

        filas.append([
            "Descuentos por compensaciones",
            f"una cantidad de -S/.{data.get('compensaciones', 0):.2f}"
        ])

        if data.get("descuentos_fugaces", 0):
            filas.append(
                ["Promociones por descuentos fugaces", f"una cantidad de -S/.{data['descuentos_fugaces']:.2f}"])

        if data.get("descuentos_cancelaciones", 0):
            filas.append(["Descuento por cancelaciones", f"una cantidad de -S/.{data['descuentos_cancelaciones']:.2f}"])

        if data.get("descuentos_reclamos", 0):
            filas.append(
                ["Descuento por reclamos de usuarios", f"una cantidad de -S/.{data['descuentos_reclamos']:.2f}"])

        if data.get("reintegros_tienda", 0):
            filas.append(["Reintegro a favor de la tienda", f"una cantidad de +S/.{data['reintegros_tienda']:.2f}"])

        filas = [[sanitize_cell(c) for c in f] for f in filas]

        tabla = Table(filas, colWidths=[9 * cm, 7 * cm])

        tabla.setStyle(TableStyle([
            ("ALIGN", (1, 0), (1, -1), "RIGHT"),
            ("LEFTPADDING", (0, 0), (-1, -1), 2),
            ("RIGHTPADDING", (0, 0), (-1, -1), 2),
            ("TOPPADDING", (0, 0), (-1, -1), 3),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
            ("FONT", (0, 0), (-1, -1), "Helvetica", 10)
        ]))

        elementos.append(tabla)
        elementos.append(Spacer(1, 10))

    # ORDEN FIJO
    SUCURSALES = [
        "Astrobuns | Smashburgers",
        "Incheon | Korean Fried Chicken"
    ]

    for sucursal in SUCURSALES:

        if sucursal in finanzas_rappi:
            bloque_finanzas(sucursal, "Rappi", finanzas_rappi[sucursal])

        if sucursal in finanzas_pedidosya:
            bloque_finanzas(sucursal, "PedidosYa", finanzas_pedidosya[sucursal])


    def calcular_saldo_total(*diccionarios: Any) -> float:

        saldo: float = 0.0

        for grupo in diccionarios:
            if not grupo:
                continue
            for data in grupo.values():
                saldo += float(data.get("total", 0))

                saldo -= abs(float(data.get("promociones_articulos", 0)))
                saldo -= abs(float(data.get("descuentos_fugaces", 0)))
                saldo -= abs(float(data.get("descuentos_cancelaciones", 0)))
                saldo -= abs(float(data.get("descuentos_reclamos", 0)))

                saldo += float(data.get("reintegros_tienda", 0))

        return round(saldo, 2)

    saldo_restante = calcular_saldo_total(
        finanzas_rappi,
        finanzas_pedidosya
    )

    pago_factura = importe_total

    pago_cci = round(saldo_restante - pago_factura, 2)

    elementos.append(Spacer(1, 10))

    tabla_resumen = Table([
        [f"Saldo restante de los pagos en línea al {format_fecha_pdf(fecha_fin)}", f"S/.{saldo_restante:.2f}"],
        [f"Pago de factura {invoice.get('numero_factura', '')}", f"S/.{pago_factura:.2f}"],
        ["Pago a cuenta CCI 00313801335213862452", ""],
        ["a la atención de GLOBALINK PERU SACS", f"TOTAL: S/.{pago_cci:.2f}"],
    ], colWidths=[12 * cm, 4 * cm])

    tabla_resumen.setStyle(TableStyle([

        ("ALIGN", (1, 0), (1, -1), "RIGHT"),

        ("LINEABOVE", (0, 0), (-1, 0), 0.8, colors.black),
        ("LINEBELOW", (0, -1), (-1, -1), 0.8, colors.black),

        ("SPAN", (0, 2), (1, 2)),

        # 👉 BOLD SOLO bloque CCI + TOTAL
        ("FONT", (0, 2), (-1, 3), "Helvetica-Bold"),

        ("LEFTPADDING", (0, 0), (-1, -1), 2),
        ("RIGHTPADDING", (0, 0), (-1, -1), 2),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
    ]))

    elementos.append(tabla_resumen)

    # ==========================
    # DETALLES PEDIDOS
    # ==========================
    elementos.append(PageBreak())
    elementos.append(
        Paragraph(
            "DETALLES PEDIDOS",
            ParagraphStyle(
                name="TituloPedidos",
                fontSize=15,
                alignment=0,
                spaceAfter=15,
                fontName="Helvetica-Bold"
            )
        )
    )

    # ==========================
    # NEGOCIO
    # ==========================
    # ==========================
    # NEGOCIO (MARCA | CONCEPTO)
    # ==========================

    def separar_marca_concepto(texto: Any):
        if not isinstance(texto, str):
            return "Sin nombre", ""

        partes = texto.replace("-", " ").split()
        if len(partes) <= 1:
            return partes[0] if partes else "Sin nombre", ""

        marca = partes[0]
        concepto = " ".join([partes[i] for i in range(1, len(partes))])
        return marca.strip(), concepto.strip()

    df["NegocioRaw"] = ""

    mask_rappi = df["Aplicativo"] == "Rappi"
    mask_pedidos = df["Aplicativo"] == "PedidosYa"

    if "Nombre de la tienda" in df.columns:
        df.loc[mask_rappi, "NegocioRaw"] = df.loc[mask_rappi, "Nombre de la tienda"]

    if "Sucursal" in df.columns:
        df.loc[mask_pedidos, "NegocioRaw"] = df.loc[mask_pedidos, "Sucursal"]

    # separar
    df[["Marca", "Concepto"]] = df["NegocioRaw"].apply(
        lambda x: pd.Series(separar_marca_concepto(x))
    )
    # clave base ÚNICA por negocio (no depende del app)
    df["MarcaKey"] = df["Marca"].str.lower().str.strip()

    # nombre final a imprimir
    df["Negocio"] = df.apply(
        lambda r: f"{r['Marca']} | {r['Concepto']}" if r["Concepto"] else r["Marca"],
        axis=1
    )

    # ==========================
    # TABLAS (ESCALABLE)
    # ==========================
    # ==========================
    # TABLAS (ORDEN CORRECTO POR NEGOCIO)
    # ==========================

    # preservar orden real de aparición
    df["_orden"] = range(len(df))

    # obtener marcas en orden de aparición REAL
    marcas = (
        df.sort_values("_orden")["MarcaKey"]
        .dropna()
        .unique()
    )

    for idx_n, marca_key in enumerate(marcas):

        df_marca = df[df["MarcaKey"] == marca_key]

        negocio_print = df_marca.iloc[0]["Negocio"]

        # aplicativos de ESTE negocio, en orden de aparición
        aplicativos = (
            df_marca.sort_values("_orden")["Aplicativo"]
            .dropna()
            .unique()
        )

        for app in aplicativos:

            df_grupo = df_marca[df_marca["Aplicativo"] == app]

            if df_grupo.empty:
                continue

            elementos.append(
                Paragraph(
                    f"<b>{negocio_print} {app} ({format_fecha_pdf(fecha_inicio)} al {format_fecha_pdf(fecha_fin)} incluye)</b>",
                    ParagraphStyle(
                        name="TituloTabla",
                        fontSize=13,
                        spaceAfter=10
                    )
                )
            )

            tabla_data = [["#", "ID del pedido", "Fecha del pedido", "Importe del pedido"]]
            df_grupo = df_grupo.reset_index(drop=True)

            def format_id(v: Any) -> str:
                if pd.isna(v) or v is None:
                    return ""
                try:
                    return str(int(float(v))) if isinstance(v, (float, int)) or (isinstance(v, str) and v.replace('.','',1).isdigit()) else str(v)
                except Exception:
                    return str(v)

            for i, fila in df_grupo.iterrows():
                if app == "Rappi":
                    pedido_id = format_id(safe_get(fila, "ID de la orden"))
                    fecha_pedido = safe_get(fila, "Fecha de creación orden", "")
                    importe = safe_get(fila, "Venta Bruta", 0)
                else:
                    pedido_id = format_id(safe_get(fila, "Numero de Pedido"))
                    fecha_pedido = safe_get(fila, "Fecha de Pedido", "")
                    importe = safe_get(fila, "Monto de la Venta ($)", 0)

                try:
                    importe = float(importe)
                except Exception:
                    importe = 0.0

                tabla_data.append([
                    i + 1,
                    sanitize_cell(pedido_id),
                    sanitize_cell(fecha_pedido),
                    f"S/ {importe:.2f}"
                ])

            tabla = Table(tabla_data, colWidths=[1.5 * cm, 4.5 * cm, 6 * cm, 4 * cm])
            tabla.setStyle(TableStyle([
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#2C3E50")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
                ("ALIGN", (0, 1), (0, -1), "CENTER"),
                ("ALIGN", (3, 1), (3, -1), "RIGHT"),
                ("FONT", (0, 0), (-1, 0), "Helvetica-Bold"),
            ]))

            elementos.append(tabla)
            elementos.append(Spacer(1, 10))

            total = 0.0
            if app == "Rappi" and "Venta Bruta" in df_grupo.columns:
                total = pd.to_numeric(df_grupo["Venta Bruta"], errors="coerce").sum()
            elif app == "PedidosYa" and "Monto de la Venta ($)" in df_grupo.columns:
                total = pd.to_numeric(df_grupo["Monto de la Venta ($)"], errors="coerce").sum()

            elementos.append(
                Paragraph(f"<b>Total {app}:</b> S/ {total:.2f}", styles["Normal"])
            )

        # salto de página por negocio
        if idx_n < len(marcas) - 1:
            elementos.append(PageBreak())

    # ==========================
    # FOOTER
    # ==========================
    def footer_canvas(canvas, doc):
        footer = load_footer_config()
        canvas.saveState()

        ancho = A4[0] - doc.leftMargin - doc.rightMargin
        canvas.setFont("Helvetica", 8)
        canvas.setFillColor(colors.grey)

        texto = (
            f"{footer['company_name']} | {footer['address']}\n"
            f"RUC: {footer['ruc']} | CCI: {footer['cci']} | IBAN: {footer['iban']}"
        )

        y = doc.bottomMargin - 0.6 * cm
        for linea in simpleSplit(texto, "Helvetica", 8, ancho):
            canvas.drawString(doc.leftMargin, y, linea)
            y -= 9

        canvas.restoreState()

    doc.build(elementos, onFirstPage=footer_canvas, onLaterPages=footer_canvas)
