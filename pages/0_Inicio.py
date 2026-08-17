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
  .scroll-inner:hover{{animation-play-state:paused;cursor:default;}}
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

st.markdown("---")
st.caption("AUXILIAR DE REGISTROS · La Sanitaria · v2.0")
