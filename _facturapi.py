"""
_facturapi.py — Cliente mínimo de la API de Facturapi (PAC) para emitir y
timbrar CFDI 4.0.

Docs: https://docs.facturapi.io

Las llaves NUNCA van en el repositorio. Se leen de los Secrets de Streamlit:

    # Una sola empresa
    [facturapi]
    secret_key = "sk_test_xxxxxxxxxxxxxxxx"

    # o varias empresas (una llave secreta por organización de Facturapi)
    [facturapi.empresas.sanitaria]
    nombre     = "La Sanitaria"
    secret_key = "sk_live_xxxxxxxxxxxxxxxx"

    [facturapi.empresas.fortez]
    nombre     = "Fortez"
    secret_key = "sk_test_xxxxxxxxxxxxxxxx"

Los CSD (.cer/.key/contraseña) se suben UNA vez en el panel de Facturapi
(Configuración → Certificados). Esta app nunca los toca.
"""
from __future__ import annotations

import requests

BASE_URL = "https://www.facturapi.io/v2"
TIMEOUT = 60


class FacturapiError(Exception):
    """Error devuelto por Facturapi (mensaje legible para el usuario)."""

    def __init__(self, message: str, status: int | None = None, path: str | None = None):
        super().__init__(message)
        self.status = status
        self.path = path


class Facturapi:
    def __init__(self, secret_key: str):
        if not secret_key:
            raise FacturapiError("Falta la llave secreta de Facturapi.")
        self.key = secret_key.strip()
        self.s = requests.Session()
        self.s.headers.update({"Authorization": f"Bearer {self.key}"})

    @property
    def modo_pruebas(self) -> bool:
        return self.key.startswith("sk_test")

    # ── núcleo ────────────────────────────────────────────────────────────
    def _req(self, method: str, path: str, *, raw: bool = False, **kw):
        try:
            r = self.s.request(method, f"{BASE_URL}{path}", timeout=TIMEOUT, **kw)
        except requests.RequestException as e:
            raise FacturapiError(f"No se pudo conectar con Facturapi: {e}") from e
        if r.status_code >= 400:
            try:
                j = r.json()
                msg = j.get("message") or str(j)
                path_err = j.get("path")
            except ValueError:
                msg, path_err = r.text[:300] or f"HTTP {r.status_code}", None
            if path_err:
                msg = f"{msg} (campo: {path_err})"
            raise FacturapiError(msg, r.status_code, path_err)
        if raw:
            return r.content
        return r.json() if r.content else {}

    # ── organización ──────────────────────────────────────────────────────
    def probar(self) -> dict:
        """Valida la llave consultando 1 factura. Devuelve el total registrado."""
        return self._req("GET", "/invoices", params={"limit": 1})

    # ── clientes ──────────────────────────────────────────────────────────
    def buscar_clientes(self, q: str = "", limit: int = 20) -> list[dict]:
        params = {"limit": limit}
        if q:
            params["q"] = q
        return self._req("GET", "/customers", params=params).get("data", [])

    def crear_cliente(self, datos: dict) -> dict:
        return self._req("POST", "/customers", json=datos)

    # ── catálogos SAT ─────────────────────────────────────────────────────
    def buscar_producto_sat(self, q: str) -> list[dict]:
        return self._req("GET", "/catalogs/products", params={"q": q}).get("data", [])

    def buscar_unidad_sat(self, q: str) -> list[dict]:
        return self._req("GET", "/catalogs/units", params={"q": q}).get("data", [])

    # ── facturas ──────────────────────────────────────────────────────────
    def crear_factura(self, datos: dict) -> dict:
        """Crea Y TIMBRA el CFDI. Devuelve la factura con uuid, folio, total."""
        return self._req("POST", "/invoices", json=datos)

    def listar_facturas(self, q: str = "", page: int = 1, limit: int = 50,
                        desde: str | None = None, hasta: str | None = None) -> dict:
        params: dict = {"page": page, "limit": limit}
        if q:
            params["q"] = q
        if desde:
            params["date[gte]"] = desde
        if hasta:
            params["date[lte]"] = hasta
        return self._req("GET", "/invoices", params=params)

    def descargar(self, invoice_id: str, formato: str = "pdf") -> bytes:
        """formato: 'pdf' | 'xml' | 'zip'"""
        return self._req("GET", f"/invoices/{invoice_id}/{formato}", raw=True)

    def enviar_correo(self, invoice_id: str, email: str | list[str] | None = None) -> dict:
        body = {"email": email} if email else {}
        return self._req("POST", f"/invoices/{invoice_id}/email", json=body)

    def cancelar(self, invoice_id: str, motivo: str, sustitucion: str | None = None) -> dict:
        """motivo SAT: 01 (con relación, requiere UUID sustituto), 02, 03, 04."""
        params = {"motive": motivo}
        if motivo == "01" and sustitucion:
            params["substitution"] = sustitucion
        return self._req("DELETE", f"/invoices/{invoice_id}", params=params)


# ── Catálogos SAT más usados (para selects) ───────────────────────────────────
REGIMENES = {
    "601": "General de Ley Personas Morales",
    "603": "Personas Morales con Fines no Lucrativos",
    "605": "Sueldos y Salarios",
    "606": "Arrendamiento",
    "608": "Demás ingresos",
    "612": "Personas Físicas con Actividades Empresariales y Profesionales",
    "616": "Sin obligaciones fiscales",
    "621": "Incorporación Fiscal",
    "625": "Plataformas Tecnológicas",
    "626": "Régimen Simplificado de Confianza (RESICO)",
}

USOS_CFDI = {
    "G01": "Adquisición de mercancías",
    "G02": "Devoluciones, descuentos o bonificaciones",
    "G03": "Gastos en general",
    "I01": "Construcciones",
    "I02": "Mobiliario y equipo de oficina",
    "I04": "Equipo de cómputo",
    "I08": "Otra maquinaria y equipo",
    "D01": "Honorarios médicos y dentales",
    "S01": "Sin efectos fiscales",
    "CP01": "Pagos",
}

FORMAS_PAGO = {
    "01": "Efectivo",
    "02": "Cheque nominativo",
    "03": "Transferencia electrónica",
    "04": "Tarjeta de crédito",
    "05": "Monedero electrónico",
    "28": "Tarjeta de débito",
    "99": "Por definir",
}

METODOS_PAGO = {
    "PUE": "Pago en una sola exhibición",
    "PPD": "Pago en parcialidades o diferido",
}

MOTIVOS_CANCELACION = {
    "02": "Comprobante emitido con errores sin relación",
    "01": "Comprobante emitido con errores con relación (requiere sustituto)",
    "03": "No se llevó a cabo la operación",
    "04": "Operación nominativa relacionada en factura global",
}

UNIDADES_COMUNES = {
    "H87": "Pieza",
    "E48": "Unidad de servicio",
    "ACT": "Actividad",
    "KGM": "Kilogramo",
    "LTR": "Litro",
    "XBX": "Caja",
    "MTR": "Metro",
}

IVA_OPCIONES = {"IVA 16%": 0.16, "IVA 8% (frontera)": 0.08, "IVA 0%": 0.0, "Exento": None}


def impuestos_concepto(iva_label: str, ret_iva: bool = False, ret_isr: bool = False) -> list[dict]:
    """Arma la lista `taxes` de Facturapi para un concepto."""
    taxes: list[dict] = []
    tasa = IVA_OPCIONES.get(iva_label, 0.16)
    if tasa is None:
        taxes.append({"type": "IVA", "factor": "Exento"})
    else:
        taxes.append({"type": "IVA", "rate": tasa})
    if ret_iva:
        # 2/3 partes del IVA (servicios profesionales / arrendamiento PF→PM)
        taxes.append({"type": "IVA", "rate": 0.106667, "withholding": True})
    if ret_isr:
        taxes.append({"type": "ISR", "rate": 0.10, "withholding": True})
    return taxes
