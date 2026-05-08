import json
import os

# ==========================
# FOOTER
# ==========================
FOOTER_CONFIG_PATH = "config/footer_config.json"

DEFAULT_FOOTER_CONFIG = {
    "company_name": "RAZÓN SOCIAL",
    "address": "DIRECCIÓN",
    "ruc": "RUC",
    "cci": "CCI",
    "iban": "IBAN"
}


def load_footer_config():
    if not os.path.exists(FOOTER_CONFIG_PATH):
        save_footer_config(DEFAULT_FOOTER_CONFIG)
        return DEFAULT_FOOTER_CONFIG

    with open(FOOTER_CONFIG_PATH, "r", encoding="utf-8") as f:
        return json.load(f)


def save_footer_config(data):
    os.makedirs(os.path.dirname(FOOTER_CONFIG_PATH), exist_ok=True)
    with open(FOOTER_CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=4, ensure_ascii=False)


# ==========================
# HEADER (NUEVO)
# ==========================
HEADER_CONFIG_PATH = "config/header_config.json"

DEFAULT_HEADER_CONFIG = {
    "company_name": "RESTAURACION GLOBAL EIRL",
    "address": "JR. ARGENTINA NRO. 201 URB. EL PARRAL ET. UNO COMAS – LIMA – LIMA",
    "ruc": "20613119923",
    "cci": "92250610000000163110",
    "iban": "GB31TCCL00997962062989"
}


def load_header_config():
    if not os.path.exists(HEADER_CONFIG_PATH):
        save_header_config(DEFAULT_HEADER_CONFIG)
        return DEFAULT_HEADER_CONFIG

    with open(HEADER_CONFIG_PATH, "r", encoding="utf-8") as f:
        return json.load(f)


def save_header_config(data):
    os.makedirs(os.path.dirname(HEADER_CONFIG_PATH), exist_ok=True)
    with open(HEADER_CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=4, ensure_ascii=False)



# ==========================
# INVOICE / FACTURA
# ==========================
INVOICE_CONFIG_PATH = "config/invoice_config.json"

DEFAULT_INVOICE_CONFIG = {
    "ciudad": "Lima",
    "atencion": "GLOBALINK PERU SACS",
    "direccion": "Jr. Argentina Nro. 199, Urb. El Parral, Comas",
    "ubigeo": "15311 Lima",
    "fecha": "",
    "ruc": "20615111903",
    "numero_factura": "E001-50",
    "customer_id": "353189496000",
    "mensaje": "Se le facturarán los siguientes servicios:",
    "logo_path": ""
}


def load_invoice_config():
    if not os.path.exists(INVOICE_CONFIG_PATH):
        save_invoice_config(DEFAULT_INVOICE_CONFIG)
        return DEFAULT_INVOICE_CONFIG

    with open(INVOICE_CONFIG_PATH, "r", encoding="utf-8") as f:
        return json.load(f)


def save_invoice_config(data):
    os.makedirs(os.path.dirname(INVOICE_CONFIG_PATH), exist_ok=True)
    with open(INVOICE_CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=4, ensure_ascii=False)
