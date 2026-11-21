"""
_app/tickets_factura.py — Motor del módulo FACTURAS DE TICKETS.

Lee fotos de tickets de consumo (gasolina, casetas, restaurantes, tiendas…) con la
API de Claude y deja todo listo para solicitar la factura (CFDI) en el portal del comercio:

  1. Receptor: datos fiscales del cliente, capturados o leídos de la Constancia de
     Situación Fiscal (PDF).
  2. Extracción: de cada foto saca comercio, RFC emisor, fecha, total, forma de pago y
     los datos que pide el portal (ticket, folio, Web ID, estación, tienda…).
  3. Portal y plazo: liga al portal conocido del comercio y fecha límite para facturar.
  4. Avisos fiscales: gasolina pagada en efectivo, efectivo > $2,000, propinas,
     consumos en restaurante, recargas TeleVía, monederos / gift cards.
  5. Salidas: correo de solicitud listo para enviar y control en Excel.

Nota: la factura la timbra siempre el comercio. Este módulo no la emite; prepara la
solicitud. Los portales con captcha o con cuenta los termina el usuario.
"""
from __future__ import annotations

import base64
import calendar
import io
import json
import re
from dataclasses import dataclass, field, asdict
from datetime import date, datetime, timedelta

MODELO = "claude-sonnet-5-5"

REGIMENES = {
    "601": "General de Ley Personas Morales",
    "603": "Personas Morales con Fines no Lucrativos",
    "605": "Sueldos y Salarios e Ingresos Asimilados a Salarios",
    "606": "Arrendamiento",
    "608": "Demás ingresos",
    "612": "Personas Físicas con Actividades Empresariales y Profesionales",
    "616": "Sin obligaciones fiscales",
    "621": "Incorporación Fiscal",
    "625": "Actividades Empresariales con ingresos a través de Plataformas Tecnológicas",
    "626": "Régimen Simplificado de Confianza",
}

USOS_CFDI = {
    "G01": "Adquisición de mercancías",
    "G03": "Gastos en general",
    "I03": "Equipo de transporte",
    "S01": "Sin efectos fiscales",
    "CP01": "Pagos",
}

CATEGORIAS = ["Gasolina", "Caseta / TeleVía", "Restaurante", "Cafetería", "Tienda", "Hotel", "Otro"]
ESTATUS = ["Pendiente", "Solicitada", "Recibida", "No facturable", "Vencida"]

# ── Portales conocidos ────────────────────────────────────────────────────────
# clave: palabras que identifican al comercio → (portal, plazo por defecto, nota)
#   plazo: ("dias", n) días naturales desde el consumo · ("horas", n) · ("mes", n) mismo
#   mes + n días del siguiente
PORTALES = [
    (("sanborns",), "https://facturaelectronica.sanborns.com.mx", ("dias", 30),
     "Pide ticket e importe; al final hay un captcha."),
    (("starbucks", "cafe sirena", "alsea", "vips", "domino", "burger king", "chili",
      "italianni", "cheesecake", "p.f. chang", "pf chang"),
     "https://alsea.interfactura.com", ("dias", 30),
     "Pide RFC, ticket (9 dígitos), tienda (5 dígitos) y fecha. Recargas de tarjeta no se facturan."),
    (("oxxo gas",), "https://www.oxxogas.com", ("mes", 0),
     "También se factura escaneando el QR con la app OXXO Gas."),
    (("g500", "gasolineras interlomas"), "https://g500network.com/facturacion-en-linea/", ("horas", 24),
     "Se busca la estación por permiso CRE; el portal pide cuenta de usuario."),
    (("7 eleven", "7-eleven", "seven eleven"), "https://www.7-eleven.com.mx", ("mes", 5),
     "Dentro del mes y los primeros 5 días del siguiente."),
    (("televia", "tele via"), "https://www.televia.com.mx", ("mes", 0),
     "Se facturan los cruces con el número de tag, no la recarga."),
    (("pemex",), "https://www.pemex.com", ("mes", 0), ""),
    (("bp ", "bp mexico"), "https://www.bpmexico.com.mx", ("mes", 0), ""),
    (("shell",), "https://www.shell.com.mx", ("mes", 0), ""),
    (("mobil",), "https://www.mobil.com.mx", ("mes", 0), ""),
    (("capufe",), "https://facturacion.capufe.gob.mx", ("mes", 0), "Casetas CAPUFE: ticket o tag."),
    (("walmart", "bodega aurrera", "sams", "sam's"), "https://facturacion.walmartmexico.com.mx", ("mes", 0), ""),
    (("oxxo",), "https://www3.oxxo.com:9443", ("mes", 0), ""),
    (("liverpool",), "https://facturacionelectronica.liverpool.com.mx", ("mes", 0), ""),
    (("costco",), "https://www3.costco.com.mx/facturacion", ("mes", 0), ""),
    (("home depot",), "https://www.homedepot.com.mx/facturacion", ("mes", 0), ""),
]


@dataclass
class Receptor:
    rfc: str = ""
    nombre: str = ""
    regimen: str = "612"
    cp: str = ""
    uso_cfdi: str = "G03"
    correo: str = ""

    def completo(self) -> bool:
        return all([self.rfc, self.nombre, self.regimen, self.cp, self.uso_cfdi, self.correo])

    def lineas(self) -> list[str]:
        return [
            f"RFC: {self.rfc}",
            f"Nombre: {self.nombre}",
            f"Régimen fiscal: {self.regimen} – {REGIMENES.get(self.regimen, '')}",
            f"Código postal: {self.cp}",
            f"Uso de CFDI: {self.uso_cfdi} – {USOS_CFDI.get(self.uso_cfdi, '')}",
        ]


@dataclass
class Ticket:
    archivo: str = ""
    comercio: str = ""
    razon_social: str = ""
    rfc_emisor: str = ""
    categoria: str = "Otro"
    fecha: str = ""                 # AAAA-MM-DD
    hora: str = ""
    total: float = 0.0
    propina: float = 0.0
    forma_pago: str = ""            # Efectivo / Tarjeta de crédito / Tarjeta de débito / Monedero / Otro
    portal: str = ""
    plazo_texto: str = ""
    correo_facturacion: str = ""
    telefono: str = ""
    datos_facturar: dict = field(default_factory=dict)   # {"Ticket": "...", "Web ID": "..."}
    notas: str = ""
    # calculados
    fecha_limite: str = ""
    avisos: list = field(default_factory=list)
    nota_portal: str = ""
    estatus: str = "Pendiente"


# ── Constancia de Situación Fiscal ───────────────────────────────────────────
def receptor_de_constancia(pdf_bytes: bytes) -> Receptor:
    """Lee RFC, nombre, CP y régimen de una Constancia de Situación Fiscal del SAT."""
    import pdfplumber
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        # x_tolerance bajo: las tablas del SAT pegan las palabras con el valor por defecto
        texto = "\n".join((p.extract_text(x_tolerance=1) or "") for p in pdf.pages[:3])
    r = Receptor()
    m = re.search(r"RFC:\s*([A-ZÑ&]{3,4}\d{6}[A-Z0-9]{3})", texto)
    r.rfc = m.group(1) if m else ""
    nom = re.search(r"Nombre\s*\(s\):\s*(.+)", texto)
    ap1 = re.search(r"Primer\s*Apellido:\s*(.+)", texto)
    ap2 = re.search(r"Segundo\s*Apellido:\s*(.*)", texto)
    if nom:
        r.nombre = " ".join(x.group(1).strip() for x in (nom, ap1, ap2) if x and x.group(1).strip())
    else:   # persona moral
        rs = re.search(r"Denominaci[oó]n\s*/\s*Raz[oó]n\s*Social:\s*(.+)", texto)
        r.nombre = rs.group(1).strip() if rs else ""
    cp = re.search(r"C[oó]digo\s*Postal:\s*(\d{5})", texto)
    r.cp = cp.group(1) if cp else ""
    regs = texto.lower().split("regímenes:", 1)[-1]
    for clave, desc in REGIMENES.items():
        if desc.lower()[:35] in regs:
            r.regimen = clave
            break
    return r


# ── Extracción con Claude ─────────────────────────────────────────────────────
_PROMPT = """Eres un auxiliar contable mexicano. La imagen es la foto de uno o varios tickets de compra
(gasolina, caseta, restaurante, tienda, cafetería…). Puede venir junto con el voucher de la terminal bancaria:
ese voucher NO es un ticket aparte; úsalo para saber la forma de pago y la propina.

Devuelve ÚNICAMENTE JSON válido con esta forma:
{"tickets": [{
  "comercio": "nombre comercial (Sanborns, OXXO Gas, Starbucks…)",
  "razon_social": "razón social del emisor si aparece",
  "rfc_emisor": "RFC del emisor si aparece, si no vacío",
  "categoria": "Gasolina | Caseta / TeleVía | Restaurante | Cafetería | Tienda | Hotel | Otro",
  "fecha": "AAAA-MM-DD",
  "hora": "HH:MM",
  "total": 0.0,
  "propina": 0.0,
  "forma_pago": "Efectivo | Tarjeta de crédito | Tarjeta de débito | Monedero | Otro",
  "portal": "URL de facturación impresa en el ticket, si aparece",
  "plazo_texto": "texto del ticket sobre el plazo para facturar, si aparece",
  "correo_facturacion": "correo para pedir factura si el ticket lo indica",
  "telefono": "teléfono del comercio si aparece",
  "datos_facturar": {"Etiqueta tal como la pide el ticket": "valor"},
  "notas": "cualquier cosa relevante: recarga, monedero, propina, ticket ilegible…"
}]}

En datos_facturar pon TODOS los identificadores que un portal podría pedir, copiados exactos:
ticket, folio, nota, Web ID, número de estación, permiso CRE, tienda, terminal, transacción,
autorización, referencia, importe. Usa punto decimal en los importes. Las fechas en tickets mexicanos
suelen ser DD/MM/AAAA; si son ambiguas, prefiere la que coincida con el voucher. No inventes datos:
si algo no se lee, déjalo vacío."""


def _media_type(nombre: str, data: bytes) -> str:
    if data[:8].startswith(b"\x89PNG"):
        return "image/png"
    if data[:4] == b"RIFF" and data[8:12] == b"WEBP":
        return "image/webp"
    return "image/jpeg"


def _reducir(data: bytes, max_lado: int = 1800) -> bytes:
    """Achica fotos muy grandes para no rebasar el límite de la API."""
    try:
        from PIL import Image, ImageOps
        im = ImageOps.exif_transpose(Image.open(io.BytesIO(data)))
        if max(im.size) <= max_lado and len(data) < 4_500_000:
            return data
        im.thumbnail((max_lado, max_lado))
        out = io.BytesIO()
        im.convert("RGB").save(out, "JPEG", quality=85)
        return out.getvalue()
    except Exception:
        return data


def extraer_tickets(nombre: str, data: bytes, api_key: str, modelo: str = MODELO) -> list[Ticket]:
    import anthropic
    data = _reducir(data)
    client = anthropic.Anthropic(api_key=api_key)
    resp = client.messages.create(
        model=modelo,
        max_tokens=3000,
        messages=[{"role": "user", "content": [
            {"type": "image", "source": {"type": "base64", "media_type": _media_type(nombre, data),
                                         "data": base64.b64encode(data).decode()}},
            {"type": "text", "text": _PROMPT},
        ]}],
    )
    raw = "".join(b.text for b in resp.content if getattr(b, "type", "") == "text")
    m = re.search(r"\{.*\}", raw, re.DOTALL)
    if not m:
        raise ValueError("La respuesta no trae JSON")
    out = []
    for d in json.loads(m.group()).get("tickets", []):
        t = Ticket(archivo=nombre)
        for k, v in d.items():
            if hasattr(t, k) and v is not None:
                setattr(t, k, v)
        t.total = _num(t.total)
        t.propina = _num(t.propina)
        t.datos_facturar = {str(k): str(v) for k, v in (t.datos_facturar or {}).items() if str(v).strip()}
        out.append(completar(t))
    return out


def _num(v) -> float:
    try:
        return round(float(str(v).replace(",", "").replace("$", "").strip() or 0), 2)
    except ValueError:
        return 0.0


# ── Portal, plazo y avisos ────────────────────────────────────────────────────
def _norm(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "").lower())


def _portal_conocido(t: Ticket):
    texto = _norm(f"{t.comercio} {t.razon_social} {t.portal}")
    for claves, url, plazo, nota in PORTALES:
        if any(c in texto for c in claves):
            return url, plazo, nota
    return None, None, ""


def _plazo_de_texto(texto: str):
    s = _norm(texto)
    m = re.search(r"(\d+)\s*(hrs|horas)", s)
    if m:
        return ("horas", int(m.group(1)))
    m = re.search(r"(\d+)\s*(?:primeros\s*)?d[ií]as", s)
    if m and "mes" in s and ("siguiente" in s or "sigue" in s):
        return ("mes", int(m.group(1)))
    if m:
        return ("dias", int(m.group(1)))
    if "mismo mes" in s or "dentro del mes" in s:
        return ("mes", 0)
    return None


def _fecha(t: Ticket) -> date | None:
    try:
        return datetime.strptime(t.fecha[:10], "%Y-%m-%d").date()
    except (ValueError, TypeError):
        return None


def fecha_limite(t: Ticket, plazo) -> date | None:
    f = _fecha(t)
    if not f or not plazo:
        return None
    tipo, n = plazo
    if tipo == "horas":
        return f + timedelta(days=max(1, round(n / 24)))
    if tipo == "dias":
        return f + timedelta(days=n)
    fin_mes = date(f.year, f.month, calendar.monthrange(f.year, f.month)[1])
    return fin_mes + timedelta(days=n)


def completar(t: Ticket, hoy: date | None = None) -> Ticket:
    hoy = hoy or date.today()
    url, plazo_def, nota = _portal_conocido(t)
    if not t.portal and url:
        t.portal = url
    elif t.portal and not t.portal.startswith("http"):
        t.portal = "https://" + t.portal.strip().lstrip("/")
    t.nota_portal = nota
    plazo = _plazo_de_texto(t.plazo_texto) or plazo_def or ("mes", 0)
    lim = fecha_limite(t, plazo)
    t.fecha_limite = lim.isoformat() if lim else ""
    if lim and lim < hoy and t.estatus == "Pendiente":
        t.estatus = "Vencida"
    t.avisos = avisos(t)
    return t


def avisos(t: Ticket) -> list[str]:
    a = []
    pago = _norm(t.forma_pago)
    cat = t.categoria
    texto = _norm(f"{t.comercio} {t.notas} {' '.join(t.datos_facturar.keys())}")
    if cat == "Gasolina" and "efectivo" in pago:
        a.append("Gasolina pagada en efectivo: no es deducible ni se acredita el IVA, sin importar el monto.")
    elif "efectivo" in pago and t.total > 2000:
        a.append("Pago en efectivo mayor a $2,000: no es deducible.")
    if t.propina:
        a.append(f"La propina (${t.propina:,.2f}) no se incluye en la factura.")
    if cat == "Restaurante":
        a.append("Consumo en restaurante: deducible solo 8.5% y pagado con tarjeta.")
    if "recarga" in texto and ("televia" in texto or "tag" in texto):
        a.append("Recarga TeleVía: la recarga no se factura; se facturan los cruces con el tag. "
                 "La comisión sí se pide al correo que indica el ticket.")
    if "monedero" in pago or "gift" in texto or "starbucks card" in texto:
        a.append("Pagado con monedero o tarjeta de regalo: algunos comercios no lo facturan.")
    if not t.portal and not t.correo_facturacion:
        a.append("El ticket no trae portal ni correo: pedir la factura directo al comercio.")
    return a


# ── Salidas ───────────────────────────────────────────────────────────────────
def texto_correo(t: Ticket, r: Receptor) -> tuple[str, str]:
    asunto = f"Solicitud de factura – {t.comercio or 'consumo'} {t.fecha} ${t.total:,.2f}"
    datos = [f"- Comercio: {t.comercio}" + (f" ({t.razon_social})" if t.razon_social else ""),
             f"- Fecha: {t.fecha} {t.hora}".rstrip(),
             f"- Total: ${t.total:,.2f}",
             f"- Forma de pago: {t.forma_pago}"]
    datos += [f"- {k}: {v}" for k, v in t.datos_facturar.items()]
    cuerpo = ("Buen día:\n\nLes solicito de favor la factura del siguiente consumo:\n\n"
              "Datos del ticket\n" + "\n".join(datos) +
              "\n\nDatos fiscales del receptor\n" + "\n".join(f"- {x}" for x in r.lineas()) +
              f"\n\nFavor de enviar el XML y el PDF a: {r.correo}\n\nMuchas gracias.\nSaludos.")
    return asunto, cuerpo


def mailto(t: Ticket, r: Receptor) -> str:
    from urllib.parse import quote
    asunto, cuerpo = texto_correo(t, r)
    return f"mailto:{t.correo_facturacion}?subject={quote(asunto)}&body={quote(cuerpo)}"


def control_excel(tickets: list[Ticket], r: Receptor) -> bytes:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter
    wb = Workbook()
    ws = wb.active
    ws.title = "Tickets"
    ws["A1"] = "CONTROL DE FACTURAS DE TICKETS"
    ws["A1"].font = Font(bold=True, size=13)
    ws["A2"] = f"Receptor: {r.rfc} · {r.nombre} · Régimen {r.regimen} · CP {r.cp} · {r.uso_cfdi} · {r.correo}"
    enc = ["Fecha", "Comercio", "RFC emisor", "Categoría", "Total", "Propina", "Forma de pago",
           "Datos para facturar", "Portal / correo", "Fecha límite", "Estatus", "Avisos", "Foto"]
    ws.append([])
    ws.append(enc)
    for c in ws[4]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="1F4E79")
    colores = {"Pendiente": "FFF2CC", "Solicitada": "DDEBF7", "Recibida": "E2EFDA",
               "Vencida": "F8CBAD", "No facturable": "D9D9D9"}
    for t in sorted(tickets, key=lambda x: x.fecha):
        ws.append([t.fecha, t.comercio, t.rfc_emisor, t.categoria, t.total, t.propina or None,
                   t.forma_pago, "\n".join(f"{k}: {v}" for k, v in t.datos_facturar.items()),
                   t.portal or t.correo_facturacion, t.fecha_limite, t.estatus,
                   "\n".join(t.avisos), t.archivo])
        fila = ws.max_row
        ws.cell(fila, 11).fill = PatternFill("solid", fgColor=colores.get(t.estatus, "FFFFFF"))
        for col in (5, 6):
            ws.cell(fila, col).number_format = '"$"#,##0.00'
        for col in (8, 12):
            ws.cell(fila, col).alignment = Alignment(wrap_text=True, vertical="top")
    fin = ws.max_row
    ws.append(["", "TOTAL", "", "", f"=SUM(E5:E{fin})"])
    ws.cell(fin + 1, 2).font = Font(bold=True)
    ws.cell(fin + 1, 5).font = Font(bold=True)
    ws.cell(fin + 1, 5).number_format = '"$"#,##0.00'
    for i, w in enumerate([11, 22, 15, 15, 12, 10, 16, 38, 38, 12, 13, 50, 22], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A5"
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


def a_dict(t: Ticket) -> dict:
    return asdict(t)


def de_dict(d: dict) -> Ticket:
    t = Ticket()
    for k, v in d.items():
        if hasattr(t, k):
            setattr(t, k, v)
    return t
