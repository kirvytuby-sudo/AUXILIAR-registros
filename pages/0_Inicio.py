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
    ticker_items.append(f'{icon} <span style="color:#FBCFE8;font-weight:700">[{_esc(etiq)}]</span> {_esc(n["titulo"])}')

ticker_html = "     ◆     ".join(ticker_items)

components.html(f"""
<!DOCTYPE html>
<html>
<head>
<style>
  *{{box-sizing:border-box;margin:0;padding:0;}}
  body{{background:transparent;font-family:'Segoe UI',Arial,sans-serif;overflow:hidden;}}
  .wrap{{
    background:linear-gradient(135deg,#1E3A8A 0%,#2563EB 100%);
    border-radius:10px;overflow:hidden;display:flex;align-items:center;height:44px;
  }}
  .badge{{
    background:rgba(251,207,232,.18);border-right:1px solid rgba(255,255,255,.2);
    padding:0 14px;height:100%;display:flex;align-items:center;
    font-size:.68rem;font-weight:700;color:#FBCFE8;white-space:nowrap;flex-shrink:0;
  }}
  .scroll-area{{flex:1;overflow:hidden;height:100%;display:flex;align-items:center;}}
  .scroll-inner{{
    white-space:nowrap;display:inline-block;
    animation:ticker {max(40, len(noticias)*8)}s linear infinite;
    color:#E0F2FE;font-size:.82rem;padding-left:100%;
  }}
  .scroll-inner:hover{{animation-play-state:paused;cursor:default;}}
  @keyframes ticker{{
    from{{transform:translateX(0);}}
    to{{transform:translateX(-100%);}}
  }}
  .ts{{font-size:.6rem;color:rgba(224,242,254,.55);padding:0 12px;flex-shrink:0;white-space:nowrap;}}
</style>
</head>
<body>
  <div class="wrap">
    <div class="badge">📡 FISCAL</div>
    <div class="scroll-area">
      <span class="scroll-inner">{ticker_html}</span>
    </div>
    <div class="ts">↻ {_esc(ultima_act)}</div>
  </div>
  <script>
    // Auto-reload la página completa cada hora
    setTimeout(function(){{ window.parent.location.reload(); }}, 3600000);
  </script>
</body>
</html>
""", height=52, scrolling=False)

# ── Tarjetas por categoría ────────────────────────────────────────────────────
tab_sat, tab_isr, tab_iva = st.tabs(["🏛️ SAT", "📊 ISR", "💵 IVA"])

def _tarjetas(filtro: str):
    items = [n for n in noticias if filtro in n["cats"]]
    if not items:
        st.caption("Sin novedades disponibles en este momento.")
        return
    cols = st.columns(min(3, len(items)))
    for i, n in enumerate(items[:9]):
        with cols[i % 3]:
            fecha_str = f"<br><span style='font-size:.68rem;color:#6B7280'>{n['fecha']}</span>" if n["fecha"] else ""
            link_str  = f"<a href='{n['link']}' target='_blank' style='font-size:.7rem;color:#2563EB;text-decoration:none'>🔗 Ver fuente</a>" if n["link"] else ""
            st.markdown(f"""
<div style="background:#fff;border:1.5px solid #BFDBFE;border-radius:10px;padding:14px 14px 10px;
            margin-bottom:8px;min-height:110px;transition:box-shadow .2s;
            box-shadow:0 2px 6px rgba(30,58,138,.07)">
  <span style="background:#DBEAFE;color:#1E40AF;font-size:.65rem;font-weight:700;
               border-radius:4px;padding:2px 7px;">{_esc(n['fuente'])}</span>
  <p style="font-size:.82rem;color:#1E293B;margin:8px 0 6px;line-height:1.45">{_esc(n['titulo'])}</p>
  {fecha_str}{link_str}
</div>""", unsafe_allow_html=True)

with tab_sat:
    _tarjetas("SAT")
with tab_isr:
    _tarjetas("ISR")
with tab_iva:
    _tarjetas("IVA")

st.markdown("<br>", unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════════════════════
# MÓDULOS
# ═══════════════════════════════════════════════════════════════════════════════
st.markdown("---")
st.markdown("##### 🗂️ Módulos disponibles")

MODULOS = [
    {"icon": "💼", "title": "Pagos Bancarios",
     "desc": "Conciliación de nómina BBVA Net Cash — PDF → Excel",
     "pagina": "pages/1_Pagos_Bancarios.py"},
    {"icon": "📋", "title": "Provisión de Nómina",
     "desc": "XML (CFDI) → Plantilla SINUBE con columnas dinámicas",
     "pagina": "pages/2_Provision_Nomina.py"},
    {"icon": "💳", "title": "Préstamos",
     "desc": "PDFs de préstamos → Excel con catálogo de cuentas",
     "pagina": "pages/3_Prestamos.py"},
    {"icon": "⛽", "title": "Ventas del Día",
     "desc": "Reporte de ventas diarias — póliza contable",
     "pagina": "pages/4_Ventas_del_Dia.py"},
    {"icon": "📊", "title": "Control Despacho vs Ventas",
     "desc": "Concilia despachos contra póliza — UUID, IVA, IEPS",
     "pagina": "pages/10_Control_Despacho_vs_Ventas.py"},
    {"icon": "🏦", "title": "Depósitos Bancarios",
     "desc": "BBVA, Banorte e Inbursa → póliza de depósitos",
     "pagina": "pages/11_Depositos_Bancarios.py"},
    {"icon": "📈", "title": "Estado de Cuenta",
     "desc": "Análisis y conciliación de estados de cuenta.",
     "pagina": "pages/9_Estado_de_Cuenta.py"},
    {"icon": "📑", "title": "Reconciliación",
     "desc": "Reconciliación contable con plantilla SINUBE",
     "pagina": "pages/6_Reconciliacion.py"},
    {"icon": "🔗", "title": "Conciliación SAT",
     "desc": "Conciliación de CFDIs contra registros contables",
     "pagina": "pages/5_Conciliacion_SAT.py"},
    {"icon": "🔀", "title": "Conciliación Banco vs Auxiliar",
     "desc": "Compara movimientos bancarios contra el auxiliar",
     "pagina": "pages/8_Conciliacion_Banco_Auxiliar.py"},
    {"icon": "🏛️", "title": "Constancia y Opinión SAT",
     "desc": "Genera Constancia Fiscal y Opinión 32-D con e.firma",
     "pagina": "pages/12_Constancia_y_Opinion_SAT.py"},
    {"icon": "📄", "title": "Póliza de Nómina",
     "desc": "Pagos Bancarios Excel → póliza matriz empleados × semana",
     "pagina": "pages/14_Poliza_Nomina.py"},
]

cols = st.columns(4)
for i, mod in enumerate(MODULOS):
    with cols[i % 4]:
        label = f"{mod['icon']}\n\n{mod['title']}\n\n{mod['desc']}\n\n✅ Disponible"
        if st.button(label, key=f"mod_{i}", use_container_width=True):
            st.switch_page(mod["pagina"])

st.markdown("---")
st.caption("AUXILIAR DE REGISTROS · La Sanitaria · v2.0")
