"""
Página de inicio — Auxiliar de Registros
Incluye ticker animado de novedades fiscales SAT/ISR/IVA con auto-refresh cada hora.
"""
import streamlit as st
import streamlit.components.v1 as components
import _theme
import urllib.request
import xml.etree.ElementTree as ET
import re
from datetime import datetime

st.set_page_config(
    page_title="Auxiliar de Registros · Inicio",
    page_icon="🧾",
    layout="wide",
)

_theme.aplicar_header("🧾 AUXILIAR DE REGISTROS",
                       "Sistema contable de La Sanitaria — Selecciona un módulo")

# ── Estilos de los botones de módulo ─────────────────────────────────────────
st.markdown("""
<style>
div[data-testid="column"] .stButton > button {
    background: #EFF6FF;
    border: 2px solid #1E3A8A;
    border-radius: 10px;
    padding: 18px 12px;
    width: 100%;
    min-height: 175px;
    text-align: center;
    white-space: pre-wrap;
    line-height: 1.6;
    color: #1E3A8A;
    font-size: 0.88rem;
    cursor: pointer;
    transition: border-color .2s, background .2s, box-shadow .2s;
    margin-bottom: 4px;
}
div[data-testid="column"] .stButton > button:hover {
    background: #DBEAFE;
    border-color: #2563EB;
    box-shadow: 0 4px 12px rgba(30,58,138,.18);
}
</style>
""", unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════════════════════
# TICKER — Novedades fiscales
# ═══════════════════════════════════════════════════════════════════════════════

KEYWORDS_ISR  = ["ISR","IMPUESTO SOBRE LA RENTA","RETENCIÓN","RETENCI","TARIFA","PTU","AGUINALDO",
                 "ASIMILABLE","PERSONA MORAL","PERSONA FÍSICA","PERSONA FISICA","DEDUCCIÓN","DEDUCCION"]
KEYWORDS_IVA  = ["IVA","IMPUESTO AL VALOR","TASA CERO","EXENTO","TRASLADO","ACREDITABLE","IEPS",
                 "FRONTERA","FRONTERIZA"]
KEYWORDS_SAT  = ["SAT","CFDI","COMPROBANTE FISCAL","TIMBRE","BUZÓN","BUZON","CFF","CSD","E.FIRMA",
                 "E-FIRMA","RFC","CONSTANCIA","DECLARACIÓN","DECLARACION","OPINIÓN","OPINION",
                 "CUMPLIMIENTO","MISCELÁNEA","MISCELANEA"]


@st.cache_data(ttl=3600, show_spinner=False)
def _fetch_novedades():
    """Obtiene noticias fiscales del DOF (RSS) y SAT (HTML). Fallback estático si falla."""
    noticias = []

    # ── 1. DOF RSS ────────────────────────────────────────────────────────────
    try:
        req = urllib.request.Request(
            "https://www.dof.gob.mx/rss/index.php",
            headers={"User-Agent": "Mozilla/5.0 (compatible; AuxiliarRegistros/1.0)"},
        )
        with urllib.request.urlopen(req, timeout=8) as resp:
            xml_data = resp.read()

        root = ET.fromstring(xml_data)
        for item in root.iter("item"):
            titulo = (item.findtext("title") or "").strip()
            link   = (item.findtext("link")  or "").strip()
            fecha  = (item.findtext("pubDate") or "")[:16].strip()
            upper  = titulo.upper()

            if any(k in upper for k in KEYWORDS_ISR + KEYWORDS_IVA + KEYWORDS_SAT):
                cats = []
                if any(k in upper for k in KEYWORDS_ISR): cats.append("ISR")
                if any(k in upper for k in KEYWORDS_IVA): cats.append("IVA")
                if any(k in upper for k in KEYWORDS_SAT): cats.append("SAT")
                noticias.append({
                    "fuente": "DOF",
                    "titulo": titulo[:140],
                    "link":   link,
                    "fecha":  fecha,
                    "cats":   cats,
                })
            if len(noticias) >= 12:
                break
    except Exception:
        pass

    # ── 2. SAT Noticias (HTML scraping ligero) ────────────────────────────────
    try:
        req2 = urllib.request.Request(
            "https://www.sat.gob.mx/noticias",
            headers={"User-Agent": "Mozilla/5.0 (compatible; AuxiliarRegistros/1.0)"},
        )
        with urllib.request.urlopen(req2, timeout=8) as resp:
            html = resp.read().decode("utf-8", errors="replace")

        # Buscar títulos: SAT usa <h2>, <h3>, y links con clase "titulo"
        patrones = [
            r'class="[^"]*(?:titulo|title|noticia)[^"]*"[^>]*>\s*<a[^>]*>([^<]{20,160})</a>',
            r'<h[23][^>]*>\s*<a[^>]*>([^<]{20,160})</a>\s*</h[23]>',
            r'<h[23][^>]*>([^<]{20,160})</h[23]>',
        ]
        vistos = set()
        for pat in patrones:
            for m in re.finditer(pat, html, re.I | re.S):
                t = re.sub(r'\s+', ' ', m.group(1)).strip()
                t = re.sub(r'<[^>]+>', '', t).strip()  # quitar etiquetas residuales
                if len(t) < 20 or t in vistos:
                    continue
                vistos.add(t)
                upper = t.upper()
                cats = []
                if any(k in upper for k in KEYWORDS_ISR): cats.append("ISR")
                if any(k in upper for k in KEYWORDS_IVA): cats.append("IVA")
                # todas las de SAT se etiquetan como SAT
                cats.append("SAT")
                noticias.append({
                    "fuente": "SAT",
                    "titulo": t[:140],
                    "link":   "https://www.sat.gob.mx/noticias",
                    "fecha":  "",
                    "cats":   list(dict.fromkeys(cats)),
                })
                if len(noticias) >= 20:
                    break
            if len(noticias) >= 20:
                break
    except Exception:
        pass

    # ── 3. Fallback estático si no se obtuvo nada ─────────────────────────────
    if not noticias:
        noticias = [
            {"fuente": "SAT", "titulo": "Consulta sat.gob.mx para las últimas disposiciones fiscales",
             "link": "https://www.sat.gob.mx/noticias", "fecha": "", "cats": ["SAT"]},
            {"fuente": "ISR", "titulo": "Tarifa del ISR 2025 — verifica retenciones y límites de deducción personal",
             "link": "https://www.sat.gob.mx", "fecha": "", "cats": ["ISR"]},
            {"fuente": "IVA", "titulo": "IVA general 16% · zona fronteriza 8% — revisa exenciones aplicables",
             "link": "https://www.sat.gob.mx", "fecha": "", "cats": ["IVA"]},
            {"fuente": "CFDI", "titulo": "CFDI v4.0 obligatorio — valida complementos de nómina 1.2 vigentes",
             "link": "https://www.sat.gob.mx", "fecha": "", "cats": ["SAT"]},
            {"fuente": "DOF",  "titulo": "Revisa el Diario Oficial para cambios en leyes fiscales 2025",
             "link": "https://www.dof.gob.mx", "fecha": "", "cats": ["SAT", "ISR", "IVA"]},
        ]

    ts = datetime.now().strftime("%d/%m/%Y %H:%M")
    return noticias, ts


# ── Fetch data ────────────────────────────────────────────────────────────────
col_hdr, col_btn = st.columns([6, 1])
with col_hdr:
    st.markdown("##### 📡 Novedades Fiscales — SAT · ISR · IVA")
with col_btn:
    if st.button("↻ Actualizar", key="refresh_news", help="Forzar recarga de noticias"):
        _fetch_novedades.clear()
        st.rerun()

noticias, ultima_act = _fetch_novedades()

# ── Ticker HTML ───────────────────────────────────────────────────────────────
def _esc(s):
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;")

ticker_items = []
for n in noticias:
    icon = "📌" if n["fuente"] == "DOF" else "🏛️"
    etiq = " · ".join(n["cats"]) if n["cats"] else n["fuente"]
    href = n["link"] if n["link"] else "#"
    href_esc = href.replace("&", "&amp;").replace("'", "%27").replace('"', "%22")
    ticker_items.append(
        f'<a href="{href_esc}" target="_blank" style="color:inherit;text-decoration:none;">'
        f'{icon} <span style="color:#FBCFE8;font-weight:700">[{_esc(etiq)}]</span>'
        f' {_esc(n["titulo"])}'
        f'</a>'
    )

ticker_html = '&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span style="color:#93C5FD;opacity:.7">◆</span>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;'.join(ticker_items)

components.html(f"""
<!DOCTYPE html>
<html>
<head>
<style>
  *{{box-sizing:border-box;margin:0;padding:0;}}
  body{{background:transparent;font-family:'Segoe UI',Arial,sans-serif;overflow:hidden;}}
  .wrap{{
    background:linear-gradient(135deg,#0F2167 0%,#1E40AF 60%,#2563EB 100%);
    border-radius:12px;
    overflow:hidden;
    display:flex;
    align-items:center;
    height:64px;
    box-shadow:0 4px 18px rgba(30,58,138,.45);
  }}
  .badge{{
    background:rgba(251,207,232,.15);
    border-right:2px solid rgba(255,255,255,.18);
    padding:0 20px;
    height:100%;
    display:flex;
    flex-direction:column;
    align-items:center;
    justify-content:center;
    gap:2px;
    flex-shrink:0;
    min-width:90px;
  }}
  .badge-icon{{font-size:1.3rem;line-height:1;}}
  .badge-label{{font-size:.6rem;font-weight:800;color:#FBCFE8;letter-spacing:1.5px;text-transform:uppercase;}}
  .scroll-area{{flex:1;overflow:hidden;height:100%;display:flex;align-items:center;}}
  .scroll-inner{{
    white-space:nowrap;
    display:inline-block;
    animation:ticker {max(50, len(noticias)*9)}s linear infinite;
    color:#F0F9FF;
    font-size:.95rem;
    font-weight:500;
    padding-left:100%;
    letter-spacing:.2px;
  }}
  .scroll-inner:hover{{animation-play-state:paused;}}
  .scroll-inner a{{color:inherit;text-decoration:none;cursor:pointer;}}
  .scroll-inner a:hover{{text-decoration:underline;text-underline-offset:3px;}}
  @keyframes ticker{{
    from{{transform:translateX(0);}}
    to{{transform:translateX(-100%);}}
  }}
  .ts-wrap{{
    display:flex;flex-direction:column;align-items:flex-end;
    padding:0 14px;flex-shrink:0;gap:3px;
  }}
  .ts{{font-size:.62rem;color:rgba(224,242,254,.6);white-space:nowrap;}}
  .live-dot{{
    display:inline-block;width:7px;height:7px;border-radius:50%;
    background:#34D399;box-shadow:0 0 6px #34D399;
    animation:pulse 1.8s ease-in-out infinite;
  }}
  @keyframes pulse{{0%,100%{{opacity:1;transform:scale(1);}}50%{{opacity:.5;transform:scale(.75);}}}}
</style>
</head>
<body>
  <div class="wrap">
    <div class="badge">
      <span class="badge-icon">📡</span>
      <span class="badge-label">FISCAL</span>
    </div>
    <div class="scroll-area">
      <span class="scroll-inner">{ticker_html}</span>
    </div>
    <div class="ts-wrap">
      <span class="live-dot"></span>
      <span class="ts">↻ {_esc(ultima_act)}</span>
    </div>
  </div>
  <script>
    setTimeout(function(){{ window.parent.location.reload(); }}, 3600000);
  </script>
</body>
</html>
""", height=72, scrolling=False)

# ── Tarjetas por categoría ────────────────────────────────────────────────────
CARD_CSS = """
<style>
  *{box-sizing:border-box;margin:0;padding:0;}
  body{background:transparent;font-family:'Segoe UI',Arial,sans-serif;}
  .grid{display:grid;grid-template-columns:repeat(3,1fr);gap:12px;padding:4px 2px 8px;}
  .card{
    background:#fff;border:1.5px solid #BFDBFE;border-radius:12px;
    padding:16px 16px 12px;min-height:130px;
    box-shadow:0 2px 8px rgba(30,58,138,.08);
    transition:box-shadow .2s,border-color .2s,transform .15s;
    display:flex;flex-direction:column;gap:8px;cursor:pointer;
  }
  .card:hover{
    box-shadow:0 6px 20px rgba(30,58,138,.18);
    border-color:#2563EB;transform:translateY(-2px);
  }
  .badge{
    display:inline-block;background:#DBEAFE;color:#1E40AF;
    font-size:.63rem;font-weight:800;border-radius:5px;
    padding:3px 8px;letter-spacing:.5px;width:fit-content;
  }
  .titulo{
    font-size:.86rem;color:#1E293B;line-height:1.5;flex:1;
    text-decoration:none;
  }
  .titulo:hover{color:#1D4ED8;text-decoration:underline;}
  .footer{display:flex;justify-content:space-between;align-items:center;margin-top:4px;}
  .fecha{font-size:.67rem;color:#9CA3AF;}
  .btn{
    display:inline-flex;align-items:center;gap:4px;
    background:#EFF6FF;color:#1D4ED8;
    font-size:.72rem;font-weight:600;
    border:1.5px solid #BFDBFE;border-radius:6px;
    padding:4px 10px;text-decoration:none;
    transition:background .15s,border-color .15s;
  }
  .btn:hover{background:#DBEAFE;border-color:#2563EB;}
  @media(max-width:700px){.grid{grid-template-columns:1fr;}}
</style>
"""

def _tarjetas_html(filtro: str) -> str:
    items = [n for n in noticias if filtro in n["cats"]]
    if not items:
        return "<p style='color:#6B7280;font-size:.85rem;padding:12px'>Sin novedades disponibles en este momento.</p>"
    cards = ""
    for n in items[:9]:
        href  = _esc(n["link"]) if n["link"] else "#"
        fecha = f'<span class="fecha">{_esc(n["fecha"])}</span>' if n["fecha"] else '<span></span>'
        cards += f"""
<div class="card" onclick="window.open('{href}','_blank')">
  <span class="badge">{_esc(n["fuente"])}</span>
  <a class="titulo" href="{href}" target="_blank">{_esc(n["titulo"])}</a>
  <div class="footer">
    {fecha}
    <a class="btn" href="{href}" target="_blank">🔗 Ver fuente</a>
  </div>
</div>"""
    rows = max(1, (len(items[:9]) + 2) // 3)
    return cards, rows

# ══════════════════════════════════════════════════════════════════════════════
# DATOS ESTÁTICOS — LISR y LIVA
# ══════════════════════════════════════════════════════════════════════════════
_ISR_RESUMEN = (
    "La <strong>Ley del Impuesto sobre la Renta (LISR)</strong> regula el gravamen sobre "
    "ingresos percibidos por personas físicas y morales residentes en México o con fuente "
    "de riqueza en territorio nacional. Las personas morales aplican una tasa fija del "
    "<strong>30&nbsp;%</strong> sobre la utilidad fiscal del ejercicio. Las personas físicas "
    "tributan con tarifa progresiva del <strong>0&nbsp;% al 35&nbsp;%</strong>. "
    "Los patrones deben retener, calcular y enterar el ISR de sus trabajadores mensual y "
    "anualmente (arts.&nbsp;96 y&nbsp;97). Las deducciones autorizadas deben ser estrictamente "
    "indispensables y estar amparadas con CFDI (art.&nbsp;27). Los dividendos distribuidos "
    "causan una retención adicional del <strong>10&nbsp;%</strong> (art.&nbsp;10)."
)
_ISR_URL = "https://www.diputados.gob.mx/LeyesBiblio/pdf/LISR.pdf"
_ISR_ARTS = [
    ("Art. 1",   "Sujetos del ISR",
     "Residentes en México y extranjeros con ingresos de fuente de riqueza nacional están obligados al pago del impuesto."),
    ("Art. 9",   "Tasa 30 % — Personas Morales",
     "Las personas morales calcularán el ISR aplicando el 30&nbsp;% sobre su utilidad fiscal; el resultado es el impuesto del ejercicio."),
    ("Art. 10",  "Dividendos — retención 10 %",
     "Los dividendos o utilidades distribuidos por personas morales están sujetos a una tasa adicional del 10&nbsp;% a cargo de la empresa pagadora."),
    ("Art. 27",  "Requisitos de deducciones",
     "Las deducciones deben ser estrictamente indispensables, estar amparadas con CFDI y cumplir requisitos de pago y registro contable."),
    ("Art. 28",  "Gastos no deducibles",
     "Pagos en efectivo mayores a $2,000; el 91.5&nbsp;% de consumos en restaurantes; viáticos sin requisitos; intereses en exceso de mercado, entre otros."),
    ("Art. 40",  "Tasas de depreciación",
     "Porcentajes máximos: edificios 5&nbsp;%, mobiliario 10&nbsp;%, equipo de cómputo 30&nbsp;%, automóviles 25&nbsp;%, activo fijo industrial 10–35&nbsp;%."),
    ("Art. 76",  "Obligaciones de Personas Morales",
     "Expedir CFDI, llevar contabilidad electrónica, presentar declaraciones anuales y retener ISR a trabajadores y proveedores de servicios profesionales."),
    ("Art. 86",  "Partes relacionadas",
     "Las operaciones con partes relacionadas deben pactarse a precios de mercado (arm's length) y documentarse con un estudio de precios de transferencia."),
    ("Art. 94",  "Salarios y asimilados",
     "Son ingresos por salarios los derivados de una relación laboral; los asimilados (honorarios a CA, anticipos de SC) se gravan con el mismo tratamiento."),
    ("Art. 96",  "Retención mensual de trabajadores",
     "El patrón retiene el ISR mensual aplicando la tabla del art.&nbsp;96 sobre el ingreso gravable de cada trabajador, considerando subsidio al empleo."),
    ("Art. 150", "Declaración anual — Personas Físicas",
     "Las PF con ingresos superiores a $400,000 o con dos o más patrones están obligadas a presentar declaración anual en el mes de abril."),
    ("Art. 152", "Tarifa anual — Personas Físicas",
     "Tarifa progresiva de 0&nbsp;% a 35&nbsp;%: incluye límite inferior, cuota fija y porcentaje a aplicar sobre el excedente del límite inferior."),
]

_IVA_RESUMEN = (
    "La <strong>Ley del Impuesto al Valor Agregado (LIVA)</strong> grava la enajenación de "
    "bienes, la prestación de servicios independientes, el uso o goce temporal de bienes y "
    "la importación. La tasa general es del <strong>16&nbsp;%</strong>; en la franja "
    "fronteriza norte se aplica una tasa reducida del <strong>8&nbsp;%</strong>. "
    "Existe tasa del <strong>0&nbsp;%</strong> para alimentos básicos sin proceso industrial, "
    "medicamentos de patente y exportaciones de bienes. "
    "El impuesto es <strong>trasladable en cadena</strong>: el contribuyente puede acreditar "
    "el IVA pagado en sus compras contra el causado en sus ventas; si el acreditable excede "
    "al causado se obtiene un <strong>saldo a favor</strong> solicitado en devolución o "
    "compensado en períodos siguientes (art.&nbsp;6)."
)
_IVA_URL = "https://www.diputados.gob.mx/LeyesBiblio/pdf/LIVA.pdf"
_IVA_ARTS = [
    ("Art. 1",   "Tasa general 16 %",
     "Personas físicas y morales en México que enajenen bienes, presten servicios, otorguen uso temporal de bienes o importen están obligadas al pago del IVA."),
    ("Art. 2-A", "Tasa 0 %",
     "Alimentos sin proceso industrial, medicamentos de patente, agua no gaseosa en envases mayores de 10 litros, libros y revistas, y servicios de exportación digital."),
    ("Art. 4",   "Acreditamiento",
     "El IVA trasladado al contribuyente y el pagado en importación es acreditable contra el IVA causado del mismo período de declaración."),
    ("Art. 5",   "Requisitos para acreditamiento",
     "El bien o servicio debe ser estrictamente indispensable; el IVA debe estar expresamente trasladado en CFDI y haber sido efectivamente pagado."),
    ("Art. 6",   "Saldo a favor",
     "Cuando el IVA acreditable excede al causado se obtiene saldo a favor, que puede solicitarse en devolución o compensarse contra obligaciones propias en períodos siguientes."),
    ("Art. 9",   "Enajenaciones exentas",
     "Suelo, construcciones destinadas a habitación, libros y revistas con contenido editorial, bienes muebles usados por personas físicas sin actividad empresarial y lingotes de oro."),
    ("Art. 14",  "Prestación de servicios",
     "Se considera prestación de servicios independientes toda obligación de hacer que no constituya relación laboral; incluye comisiones, agencia, representación y mandato."),
    ("Art. 15",  "Servicios exentos",
     "Comisiones de créditos hipotecarios, seguros de vida y gastos médicos, transporte terrestre de personas, enseñanza (con validez oficial) y espectáculos públicos."),
    ("Art. 24",  "Importación de bienes",
     "Se considera importación: introducción de bienes al país, uso temporal de bienes extranjeros y adquisición de servicios de residentes en el extranjero aprovechados en México."),
    ("Art. 29",  "Exportación — tasa 0 %",
     "La exportación definitiva de bienes tangibles, servicios aprovechados en el extranjero y el transporte internacional de bienes y personas se grava a tasa 0&nbsp;%."),
    ("Art. 32",  "Obligaciones generales",
     "Trasladar el IVA en forma expresa en CFDI, expedir comprobantes por todos los actos gravados y presentar declaraciones mensuales a más tardar el día 17 del mes siguiente."),
    ("Art. 33",  "Actos accidentales",
     "Las personas que realicen actos o actividades accidentales por los que deban pagar IVA presentarán declaración dentro de los 15 días hábiles siguientes a la obtención del ingreso."),
]


def _render_ley_html(titulo: str, resumen: str, articulos: list, url_ley: str, color: str) -> None:
    """Renderiza tarjeta de resumen de ley + grid de artículos clave + botón ley completa."""
    arts_html = ""
    for codigo, nombre, desc in articulos:
        arts_html += f"""
<div class="art-card">
  <div class="art-header">
    <span class="art-num">{codigo}</span>
    <span class="art-name">{nombre}</span>
  </div>
  <p class="art-desc">{desc}</p>
</div>"""

    rows = (len(articulos) + 2) // 3
    h = 80 + 100 + rows * 130 + 70  # resumen + grid rows + botón

    components.html(f"""<!DOCTYPE html><html><head>
<style>
  *{{box-sizing:border-box;margin:0;padding:0;}}
  body{{background:transparent;font-family:'Segoe UI',Arial,sans-serif;padding:6px 2px;}}

  /* ── Resumen ─────────────────────────────── */
  .resumen{{
    background:linear-gradient(135deg,{color}18 0%,{color}08 100%);
    border:1.5px solid {color}55;border-radius:12px;
    padding:18px 20px;margin-bottom:16px;
    font-size:.88rem;color:#1E293B;line-height:1.65;
  }}
  .resumen strong{{color:{color};}}

  /* ── Subtítulo ───────────────────────────── */
  .subtitulo{{
    font-size:.78rem;font-weight:700;color:{color};
    letter-spacing:.5px;text-transform:uppercase;
    margin-bottom:10px;padding-left:2px;
  }}

  /* ── Grid artículos ──────────────────────── */
  .grid{{display:grid;grid-template-columns:repeat(3,1fr);gap:10px;margin-bottom:16px;}}
  .art-card{{
    background:#fff;border:1.5px solid {color}33;border-radius:10px;
    padding:12px 14px;min-height:115px;
    box-shadow:0 2px 6px {color}14;
    transition:box-shadow .2s,transform .15s,border-color .2s;
  }}
  .art-card:hover{{
    box-shadow:0 5px 16px {color}28;border-color:{color};
    transform:translateY(-2px);
  }}
  .art-header{{display:flex;align-items:center;gap:8px;margin-bottom:6px;}}
  .art-num{{
    background:{color};color:#fff;
    font-size:.63rem;font-weight:800;border-radius:5px;
    padding:3px 8px;white-space:nowrap;flex-shrink:0;
  }}
  .art-name{{font-size:.82rem;font-weight:700;color:#1E293B;}}
  .art-desc{{font-size:.76rem;color:#475569;line-height:1.5;}}

  /* ── Botón ───────────────────────────────── */
  .btn-ley{{
    display:inline-flex;align-items:center;gap:8px;
    background:{color};color:#fff;
    font-size:.85rem;font-weight:700;
    border:none;border-radius:10px;
    padding:12px 24px;text-decoration:none;
    box-shadow:0 4px 14px {color}44;
    transition:opacity .2s,transform .15s,box-shadow .2s;
    cursor:pointer;
  }}
  .btn-ley:hover{{opacity:.88;transform:translateY(-1px);box-shadow:0 6px 20px {color}55;}}
  .btn-wrap{{text-align:center;padding-top:4px;}}

  @media(max-width:700px){{.grid{{grid-template-columns:1fr;}}}}
</style>
</head><body>
  <div class="resumen">{resumen}</div>
  <div class="subtitulo">📌 Artículos más importantes</div>
  <div class="grid">{arts_html}</div>
  <div class="btn-wrap">
    <a class="btn-ley" href="{url_ley}" target="_blank">
      📄 Ver {titulo} completa en Cámara de Diputados
    </a>
  </div>
</body></html>""", height=h, scrolling=False)


tab_sat, tab_isr, tab_iva = st.tabs(["🏛️ SAT", "📊 ISR", "💵 IVA"])

def _render_tab(filtro: str):
    result = _tarjetas_html(filtro)
    if isinstance(result, str):
        st.markdown(result, unsafe_allow_html=True)
        return
    cards_html, rows = result
    h = 160 * rows + 20
    components.html(f"<!DOCTYPE html><html><head>{CARD_CSS}</head>"
                    f"<body><div class='grid'>{cards_html}</div></body></html>",
                    height=h, scrolling=False)

with tab_sat:
    _render_tab("SAT")

with tab_isr:
    _render_tab("ISR")
    st.markdown("---")
    _render_ley_html(
        titulo    = "LISR",
        resumen   = _ISR_RESUMEN,
        articulos = _ISR_ARTS,
        url_ley   = _ISR_URL,
        color     = "#1D4ED8",   # azul
    )

with tab_iva:
    _render_tab("IVA")
    st.markdown("---")
    _render_ley_html(
        titulo    = "LIVA",
        resumen   = _IVA_RESUMEN,
        articulos = _IVA_ARTS,
        url_ley   = _IVA_URL,
        color     = "#059669",   # verde esmeralda
    )

st.markdown("---")
st.caption("AUXILIAR DE REGISTROS · La Sanitaria · v2.0")
