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
import json
from datetime import datetime
from urllib.request import urlopen

st.set_page_config(
    page_title="Auxiliar de Registros · Inicio",
    page_icon="🧾",
    layout="wide",
)

_theme.aplicar_header("🧾 AUXILIAR DE REGISTROS",
                       "Sistema contable de La Sanitaria — Selecciona un módulo")

# ── Estilos globales ──────────────────────────────────────────────────────────
st.markdown("""
<style>
/* ── Tabs ────────────────────────────────────── */
button[data-baseweb="tab"] {
    font-weight: 700 !important;
    font-size: 0.92rem !important;
    letter-spacing: .2px;
    padding: 10px 22px !important;
}
button[data-baseweb="tab"][aria-selected="true"] {
    color: #1E3A8A !important;
}
div[data-baseweb="tab-highlight"] {
    background-color: #1E3A8A !important;
    height: 3px !important;
    border-radius: 3px 3px 0 0 !important;
}
div[data-baseweb="tab-list"] {
    gap: 4px;
    background: #F8FAFC;
    border-radius: 12px 12px 0 0;
    padding: 4px 4px 0;
    border-bottom: 2px solid #E2E8F0;
}
div[data-baseweb="tab-panel"] {
    padding-top: 14px !important;
}
/* ── Scrollbar ───────────────────────────────── */
::-webkit-scrollbar { width: 6px; height: 6px; }
::-webkit-scrollbar-track { background: #F1F5F9; }
::-webkit-scrollbar-thumb { background: #CBD5E1; border-radius: 3px; }
::-webkit-scrollbar-thumb:hover { background: #94A3B8; }
/* ── Refresh button ──────────────────────────── */
button[data-testid="baseButton-secondary"] {
    border-radius: 20px !important;
    font-weight: 600 !important;
    font-size: .8rem !important;
}
/* ── st.caption ──────────────────────────────── */
div[data-testid="stCaptionContainer"] p {
    color: #94A3B8;
    font-size: .72rem;
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

    # ── 2. SAT Noticias — multi-estrategia ───────────────────────────────────
    def _sat_html(url: str, timeout: int = 9) -> str:
        try:
            req = urllib.request.Request(url, headers={
                "User-Agent": (
                    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                    "AppleWebKit/537.36 (KHTML, like Gecko) "
                    "Chrome/124.0 Safari/537.36"
                ),
                "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
                "Accept-Language": "es-MX,es;q=0.9,en;q=0.8",
            })
            with urllib.request.urlopen(req, timeout=timeout) as r:
                return r.read().decode("utf-8", errors="replace")
        except Exception:
            return ""

    def _clean(t: str) -> str:
        t = re.sub(r'<[^>]+>', '', t)
        t = re.sub(r'&[a-z]+;', ' ', t)
        return re.sub(r'\s+', ' ', t).strip()

    _sat_vistos: set = set()

    def _add_sat(titulo: str, link: str = "", fecha: str = ""):
        t = _clean(titulo)[:160]
        if len(t) < 20 or t.lower() in _sat_vistos:
            return
        _sat_vistos.add(t.lower())
        upper = t.upper()
        cats: list[str] = []
        if any(k in upper for k in KEYWORDS_ISR): cats.append("ISR")
        if any(k in upper for k in KEYWORDS_IVA): cats.append("IVA")
        cats.append("SAT")
        noticias.append({
            "fuente": "SAT",
            "titulo": t,
            "link":   link or "https://www.sat.gob.mx/noticias",
            "fecha":  fecha,
            "cats":   list(dict.fromkeys(cats)),
        })

    # ── 2a. Página principal de noticias SAT ──────────────────────────────────
    html_n = _sat_html("https://www.sat.gob.mx/noticias")
    if html_n:
        # Estrategia 1: JSON-LD structured data
        for ld_m in re.finditer(
            r'<script[^>]+type=["\']application/ld\+json["\'][^>]*>(.*?)</script>',
            html_n, re.S | re.I,
        ):
            try:
                data = json.loads(ld_m.group(1))
                if isinstance(data, list):
                    data = data[0] if data else {}
                nombre = data.get("name") or data.get("headline") or ""
                url_art = data.get("url") or ""
                if nombre:
                    _add_sat(nombre, url_art)
                for item in data.get("itemListElement", []):
                    if isinstance(item, dict):
                        n2 = item.get("name") or item.get("headline") or ""
                        u2 = item.get("url") or item.get("@id") or ""
                        if n2:
                            _add_sat(n2, u2)
            except Exception:
                pass

        # Estrategia 2: Angular transfer-state (datos pre-renderizados en JSON)
        for st_m in re.finditer(
            r'<script[^>]+(?:id=["\']server-app-state["\']|type=["\']application/json["\'])[^>]*>'
            r'(.*?)</script>',
            html_n, re.S | re.I,
        ):
            try:
                blob = json.loads(st_m.group(1))
                raw = json.dumps(blob, ensure_ascii=False)
                for m2 in re.finditer(
                    r'"(?:titulo|title|headline|nombre|descripcion)"\s*:\s*"([^"\\]{20,200})"',
                    raw,
                ):
                    _add_sat(m2.group(1))
            except Exception:
                pass

        # Estrategia 3: og:title y twitter:title meta tags
        for m in re.finditer(
            r'<meta[^>]+(?:property|name)=["\'](?:og|twitter):title["\'][^>]+content=["\']([^"\']{15,200})["\']',
            html_n, re.I,
        ):
            _add_sat(m.group(1), "https://www.sat.gob.mx/noticias")

        # Estrategia 4: <h1>–<h3> con links
        for m in re.finditer(
            r'<h[123][^>]*>.*?<a[^>]+href=["\']([^"\']*)["\'][^>]*>([^<]{20,160})</a>.*?</h[123]>',
            html_n, re.I | re.S,
        ):
            href = m.group(1)
            if not href.startswith("http"):
                href = "https://www.sat.gob.mx" + href.lstrip("/")
            _add_sat(m.group(2), href)

        # Estrategia 5: <h2>/<h3> sin link
        for m in re.finditer(r'<h[23][^>]*>([^<]{20,160})</h[23]>', html_n, re.I):
            _add_sat(m.group(1), "https://www.sat.gob.mx/noticias")

    # ── 2b. Sección de comunicados SAT ────────────────────────────────────────
    html_c = _sat_html("https://www.sat.gob.mx/consultas/comunicados")
    if html_c:
        for m in re.finditer(
            r'<a[^>]+href=["\']([^"\']*comunicado[^"\']*)["\'][^>]*>([^<]{15,160})</a>',
            html_c, re.I,
        ):
            href = m.group(1)
            if not href.startswith("http"):
                href = "https://www.sat.gob.mx" + href.lstrip("/")
            _add_sat(m.group(2), href)

    # ── 2c. Boletines de prensa ───────────────────────────────────────────────
    html_b = _sat_html("https://www.sat.gob.mx/consultas/boletines-prensa")
    if html_b:
        for m in re.finditer(
            r'<a[^>]+href=["\']([^"\']*)["\'][^>]*>([^<]{20,160})</a>',
            html_b, re.I,
        ):
            href = m.group(1)
            titulo = m.group(2).strip()
            if not re.search(r'(?:boletin|prensa|comunicado|aviso|noticia)', href, re.I):
                continue
            if not href.startswith("http"):
                href = "https://www.sat.gob.mx" + href.lstrip("/")
            _add_sat(titulo, href)

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


# ══════════════════════════════════════════════════════════════════════════════
# NOTICIAS SAT — feed dedicado (scraping profundo + fallback estático)
# ══════════════════════════════════════════════════════════════════════════════

# Items estáticos con links directos a secciones clave del SAT (siempre disponibles)
_SAT_URL   = "https://www.sat.gob.mx"
_DOF_URL   = "https://www.dof.gob.mx"
_DIPU_URL  = "https://www.diputados.gob.mx/LeyesBiblio/"

_SAT_FALLBACK = [
    {"titulo": "Resolución Miscelánea Fiscal — última versión vigente",
     "link": _DOF_URL, "fecha": ""},
    {"titulo": "CFDI 4.0 — Complementos, esquemas XSD y guías de llenado",
     "link": _SAT_URL, "fecha": ""},
    {"titulo": "Declaración Anual — Personas físicas y morales",
     "link": _SAT_URL, "fecha": ""},
    {"titulo": "Catálogos SAT para CFDI — c_ClaveProdServ, c_Impuesto, c_RegimenFiscal",
     "link": _SAT_URL, "fecha": ""},
    {"titulo": "Buzón Tributario — notificaciones y trámites electrónicos",
     "link": _SAT_URL, "fecha": ""},
    {"titulo": "Constancia de Situación Fiscal — descarga con RFC y e.firma",
     "link": _SAT_URL, "fecha": ""},
    {"titulo": "Opinión de Cumplimiento 32-D — verifica obligaciones fiscales",
     "link": _SAT_URL, "fecha": ""},
    {"titulo": "Declaraciones y Pagos (DyP) — presentación mensual de impuestos",
     "link": _SAT_URL, "fecha": ""},
    {"titulo": "e.firma — renovación y obtención en línea o en módulo SAT",
     "link": _SAT_URL, "fecha": ""},
    {"titulo": "Lista 69-B LISC — contribuyentes con operaciones inexistentes (EFOS)",
     "link": _SAT_URL, "fecha": ""},
    {"titulo": "Tarifas y tablas ISR vigentes — subsidio al empleo 2025",
     "link": _SAT_URL, "fecha": ""},
    {"titulo": "Pagos provisionales ISR — personas morales y físicas con actividad empresarial",
     "link": _SAT_URL, "fecha": ""},
]


@st.cache_data(ttl=3600, show_spinner=False)
def _fetch_sat_noticias():
    """Scraping dedicado de noticias SAT con múltiples fuentes y fallback estático."""
    items: list[dict] = []
    vistos: set[str] = set()

    def _add(titulo: str, link: str = "", fecha: str = ""):
        t = re.sub(r'<[^>]+>', '', titulo)
        t = re.sub(r'&[a-z]+;', ' ', t)
        t = re.sub(r'\s+', ' ', t).strip()[:160]
        if len(t) < 15 or t.lower() in vistos:
            return
        vistos.add(t.lower())
        items.append({"titulo": t, "link": link or "https://www.sat.gob.mx/noticias", "fecha": fecha})

    def _fetch_html(url: str) -> str:
        try:
            req = urllib.request.Request(url, headers={
                "User-Agent": (
                    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"
                ),
                "Accept": "text/html,application/xhtml+xml,*/*;q=0.8",
                "Accept-Language": "es-MX,es;q=0.9",
            })
            with urllib.request.urlopen(req, timeout=10) as r:
                return r.read().decode("utf-8", errors="replace")
        except Exception:
            return ""

    def _parse_html(html: str, base_url: str = "https://www.sat.gob.mx"):
        if not html:
            return

        # 1. JSON-LD structured data
        for m in re.finditer(
            r'<script[^>]+type=["\']application/ld\+json["\'][^>]*>(.*?)</script>',
            html, re.S | re.I,
        ):
            try:
                data = json.loads(m.group(1))
                if isinstance(data, list):
                    data = data[0] if data else {}
                for field in ("name", "headline", "title", "description"):
                    val = data.get(field, "")
                    if val and len(val) >= 15:
                        _add(val, data.get("url", base_url))
                        break
                for el in data.get("itemListElement", []):
                    if isinstance(el, dict):
                        n = el.get("name") or el.get("headline") or ""
                        u = el.get("url") or el.get("@id") or base_url
                        if n:
                            _add(n, u)
            except Exception:
                pass

        # 2. Angular server-app-state
        for m in re.finditer(
            r'<script[^>]+(?:id=["\']server-app-state["\']|id=["\']__NEXT_DATA__["\'])[^>]*>'
            r'(.*?)</script>',
            html, re.S | re.I,
        ):
            try:
                blob = json.loads(m.group(1))
                raw = json.dumps(blob, ensure_ascii=False)
                for m2 in re.finditer(
                    r'"(?:titulo|title|headline|nombre|descripcion|subject)"\s*:\s*"([^"\\]{15,200})"',
                    raw,
                ):
                    # Buscar URL cercana
                    snippet = raw[max(0, m2.start()-200): m2.end()+200]
                    url_m = re.search(r'"(?:url|link|href)"\s*:\s*"([^"\\]{10,200})"', snippet)
                    href = url_m.group(1) if url_m else base_url
                    if not href.startswith("http"):
                        href = "https://www.sat.gob.mx" + href.lstrip("/")
                    _add(m2.group(1), href)
            except Exception:
                pass

        # 3. og: / twitter: meta
        for m in re.finditer(
            r'<meta[^>]+(?:property|name)=["\'](?:og|twitter):title["\'][^>]+content=["\']([^"\']{10,200})["\']',
            html, re.I,
        ):
            _add(m.group(1), base_url)

        # 4. <h1>–<h3> con <a href>
        for m in re.finditer(
            r'<h[123][^>]*>.*?<a[^>]+href=["\']([^"\']{3,200})["\'][^>]*>\s*([^<]{15,160})\s*</a>.*?</h[123]>',
            html, re.I | re.S,
        ):
            href = m.group(1)
            if not href.startswith("http"):
                href = "https://www.sat.gob.mx" + href.lstrip("/")
            _add(m.group(2), href)

        # 5. <h2>/<h3> texto libre
        for m in re.finditer(r'<h[23][^>]*>([^<]{15,160})</h[23]>', html, re.I):
            _add(m.group(1), base_url)

        # 6. Links con texto descriptivo (clases típicas de Angular Material)
        for m in re.finditer(
            r'<a[^>]+href=["\']([^"\']*(?:noticia|comunicado|boletin|aviso|tramite)[^"\']*)["\']'
            r'[^>]*>\s*([^<]{15,160})\s*</a>',
            html, re.I,
        ):
            href = m.group(1)
            if not href.startswith("http"):
                href = "https://www.sat.gob.mx" + href.lstrip("/")
            _add(m.group(2), href)

    # ── Fuentes a raspar ──────────────────────────────────────────────────────
    fuentes = [
        "https://www.sat.gob.mx/noticias",
        "https://www.sat.gob.mx/consultas/comunicados",
        "https://www.sat.gob.mx/consultas/boletines-prensa",
        "https://www.sat.gob.mx/consultas/miscelanea-fiscal",
    ]
    for url in fuentes:
        html = _fetch_html(url)
        _parse_html(html, url)
        if len(items) >= 20:
            break

    # ── Fallback estático (siempre se añade al final) ─────────────────────────
    for fb in _SAT_FALLBACK:
        _add(fb["titulo"], fb["link"], fb["fecha"])

    ts = datetime.now().strftime("%d/%m/%Y %H:%M")
    return items, ts


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

components.html(f"""<!DOCTYPE html><html><head>
<style>
  *, *::before, *::after {{box-sizing:border-box;margin:0;padding:0;}}
  body {{
    background:transparent;
    font-family:'Segoe UI',system-ui,-apple-system,sans-serif;
    overflow:hidden;
    -webkit-font-smoothing:antialiased;
  }}
  .wrap {{
    background:linear-gradient(120deg,#0B1D5E 0%,#1E3A8A 45%,#1D4ED8 80%,#3B82F6 100%);
    border-radius:14px;
    overflow:hidden;
    display:flex;
    align-items:center;
    height:86px;
    box-shadow:0 6px 28px rgba(30,58,138,.5), 0 2px 8px rgba(30,58,138,.3);
    position:relative;
  }}
  /* Subtle shimmer overlay */
  .wrap::before {{
    content:'';
    position:absolute;
    inset:0;
    background:linear-gradient(180deg,rgba(255,255,255,.07) 0%,transparent 50%,rgba(0,0,0,.08) 100%);
    pointer-events:none;
  }}
  /* Left badge */
  .badge {{
    background:rgba(255,255,255,.08);
    border-right:1px solid rgba(255,255,255,.14);
    padding:0 22px;
    height:100%;
    display:flex;
    flex-direction:column;
    align-items:center;
    justify-content:center;
    gap:3px;
    flex-shrink:0;
    min-width:100px;
    position:relative;
    z-index:1;
  }}
  .badge-icon {{font-size:1.6rem;line-height:1;filter:drop-shadow(0 1px 3px rgba(0,0,0,.2));}}
  .badge-label {{font-size:.58rem;font-weight:900;color:#FBCFE8;letter-spacing:2px;text-transform:uppercase;}}
  .badge-sub {{font-size:.52rem;color:rgba(251,207,232,.55);letter-spacing:.5px;}}
  /* Scroll zone with fade masks */
  .scroll-area {{
    flex:1;overflow:hidden;height:100%;display:flex;align-items:center;
    -webkit-mask-image:linear-gradient(90deg,transparent 0%,#000 5%,#000 95%,transparent 100%);
    mask-image:linear-gradient(90deg,transparent 0%,#000 5%,#000 95%,transparent 100%);
    position:relative;z-index:1;
  }}
  .scroll-inner {{
    white-space:nowrap;
    display:inline-block;
    animation:ticker {max(60, len(noticias)*10)}s linear infinite;
    color:#E0F2FE;
    font-size:1rem;
    font-weight:500;
    padding-left:100%;
    letter-spacing:.15px;
    line-height:1;
  }}
  .scroll-inner:hover {{animation-play-state:paused;cursor:pointer;}}
  .scroll-inner a {{
    color:inherit;text-decoration:none;cursor:pointer;
    transition:color .15s;
  }}
  .scroll-inner a:hover {{
    color:#FFFFFF;
    text-decoration:underline;
    text-underline-offset:4px;
    text-decoration-color:rgba(251,207,232,.6);
  }}
  @keyframes ticker {{
    from {{transform:translateX(0);}}
    to   {{transform:translateX(-100%);}}
  }}
  /* Right timestamp */
  .ts-wrap {{
    display:flex;flex-direction:column;align-items:flex-end;
    padding:0 18px;flex-shrink:0;gap:4px;position:relative;z-index:1;
  }}
  .live-row {{display:flex;align-items:center;gap:5px;}}
  .live-dot {{
    width:8px;height:8px;border-radius:50%;
    background:#34D399;
    box-shadow:0 0 0 2px rgba(52,211,153,.3), 0 0 8px #34D399;
    animation:pulse 2s ease-in-out infinite;
  }}
  .live-label {{font-size:.58rem;font-weight:700;color:#34D399;letter-spacing:.8px;}}
  .ts {{font-size:.6rem;color:rgba(224,242,254,.5);white-space:nowrap;}}
  @keyframes pulse {{
    0%,100% {{opacity:1;transform:scale(1);box-shadow:0 0 0 2px rgba(52,211,153,.3),0 0 8px #34D399;}}
    50%      {{opacity:.6;transform:scale(.8);box-shadow:0 0 0 3px rgba(52,211,153,.15),0 0 4px #34D399;}}
  }}
</style>
</head>
<body>
  <div class="wrap">
    <div class="badge">
      <span class="badge-icon">📡</span>
      <span class="badge-label">FISCAL</span>
      <span class="badge-sub">EN VIVO</span>
    </div>
    <div class="scroll-area">
      <span class="scroll-inner">{ticker_html}</span>
    </div>
    <div class="ts-wrap">
      <div class="live-row">
        <span class="live-dot"></span>
        <span class="live-label">EN VIVO</span>
      </div>
      <span class="ts">↻ {_esc(ultima_act)}</span>
    </div>
  </div>
  <script>setTimeout(function(){{ window.parent.location.reload(); }}, 3600000);</script>
</body></html>""", height=94, scrolling=False)

# ── Tarjetas por categoría ────────────────────────────────────────────────────
# Paleta de acento por fuente
_ACCENT = {
    "DOF": "#B45309",   # ámbar oscuro
    "SAT": "#B91C1C",   # rojo
    "ISR": "#1D4ED8",   # azul
    "IVA": "#047857",   # verde
}
_ACCENT_BG = {
    "DOF": "#FEF3C7",
    "SAT": "#FEE2E2",
    "ISR": "#DBEAFE",
    "IVA": "#D1FAE5",
}
_ACCENT_DEFAULT = "#1E3A8A"

CARD_CSS = """<style>
  *, *::before, *::after {box-sizing:border-box;margin:0;padding:0;}
  body {
    background:transparent;
    font-family:'Segoe UI',system-ui,-apple-system,sans-serif;
    -webkit-font-smoothing:antialiased;
  }
  .grid {
    display:grid;
    grid-template-columns:repeat(3,1fr);
    gap:14px;
    padding:4px 2px 10px;
  }
  .card {
    background:#FFFFFF;
    border-radius:14px;
    padding:0;
    min-height:138px;
    box-shadow:0 1px 3px rgba(15,23,42,.06),0 6px 22px rgba(15,23,42,.05);
    border:1px solid rgba(226,232,240,.9);
    border-left:4px solid var(--accent,#1D4ED8);
    display:flex;
    flex-direction:column;
    cursor:pointer;
    transition:box-shadow .25s,transform .22s;
    overflow:hidden;
  }
  .card:hover {
    box-shadow:0 4px 10px rgba(15,23,42,.07),0 18px 44px rgba(15,23,42,.11);
    transform:translateY(-3px);
  }
  .card-inner {padding:16px 18px 14px;display:flex;flex-direction:column;gap:8px;flex:1;}
  .top-row {display:flex;align-items:center;gap:8px;}
  .dot {
    width:7px;height:7px;border-radius:50%;
    background:var(--accent,#1D4ED8);
    flex-shrink:0;
    box-shadow:0 0 0 2px color-mix(in srgb,var(--accent,#1D4ED8) 20%,transparent);
  }
  .fuente {
    font-size:.61rem;font-weight:800;letter-spacing:.7px;
    text-transform:uppercase;color:var(--accent,#1D4ED8);
    background:var(--accent-bg,#DBEAFE);
    padding:2px 8px;border-radius:20px;
  }
  .fecha {font-size:.64rem;color:#94A3B8;margin-left:auto;white-space:nowrap;}
  .titulo {
    font-size:.87rem;color:#0F172A;line-height:1.55;flex:1;
    text-decoration:none;font-weight:500;
    display:-webkit-box;-webkit-line-clamp:3;-webkit-box-orient:vertical;overflow:hidden;
  }
  .titulo:hover {color:var(--accent,#1D4ED8);}
  .footer {
    display:flex;align-items:center;justify-content:flex-end;
    margin-top:auto;padding-top:6px;
  }
  .btn {
    display:inline-flex;align-items:center;gap:5px;
    color:var(--accent,#1D4ED8);font-size:.71rem;font-weight:600;
    border:1.5px solid var(--accent,#1D4ED8);border-radius:20px;
    padding:4px 13px;text-decoration:none;background:transparent;
    transition:background .15s,color .15s;
  }
  .btn:hover {background:var(--accent,#1D4ED8);color:#fff;}
  .empty {color:#9CA3AF;font-size:.85rem;padding:16px 4px;}
  @media(max-width:700px) {.grid{grid-template-columns:1fr;}}
</style>"""

def _tarjetas_html(filtro: str):
    items = [n for n in noticias if filtro in n["cats"]]
    if not items:
        return "<p class='empty'>Sin novedades disponibles en este momento.</p>"
    cards = ""
    for n in items[:9]:
        href       = _esc(n["link"]) if n["link"] else "#"
        fuente     = n["fuente"]
        accent     = _ACCENT.get(fuente, _ACCENT_DEFAULT)
        accent_bg  = _ACCENT_BG.get(fuente, "#EFF6FF")
        fecha      = f'<span class="fecha">{_esc(n["fecha"])}</span>' if n["fecha"] else ""
        cards += f"""
<div class="card" style="--accent:{accent};--accent-bg:{accent_bg}"
     onclick="window.open('{href}','_blank','noopener,noreferrer')">
  <div class="card-inner">
    <div class="top-row">
      <span class="dot"></span>
      <span class="fuente">{_esc(fuente)}</span>
      {fecha}
    </div>
    <a class="titulo" href="{href}" target="_blank"
       rel="noreferrer noopener" referrerpolicy="no-referrer">{_esc(n["titulo"])}</a>
    <div class="footer">
      <a class="btn" href="{href}" target="_blank"
         rel="noreferrer noopener" referrerpolicy="no-referrer">↗ Ver fuente</a>
    </div>
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
_ISR_URL = "https://www.diputados.gob.mx/LeyesBiblio/ref/lisr.htm"
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
_IVA_URL = "https://www.diputados.gob.mx/LeyesBiblio/ref/liva.htm"
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


def _render_ley_html(titulo: str, resumen: str, articulos: list, url_ley: str,
                     color: str, stats: list[tuple] | None = None) -> None:
    """Renderiza resumen de ley + artículos clave + botón ley completa."""
    # Pills de estadísticas clave
    stats_html = ""
    if stats:
        for label, val in stats:
            stats_html += (
                f'<div class="stat">'
                f'<span class="stat-val">{val}</span>'
                f'<span class="stat-lbl">{label}</span>'
                f'</div>'
            )

    arts_html = ""
    for codigo, nombre, desc in articulos:
        arts_html += f"""
<div class="art-card">
  <div class="art-top">
    <span class="art-num">{codigo}</span>
    <span class="art-name">{nombre}</span>
  </div>
  <p class="art-desc">{desc}</p>
</div>"""

    rows = (len(articulos) + 2) // 3
    h = (110 if stats else 0) + 130 + rows * 128 + 80

    components.html(f"""<!DOCTYPE html><html><head>
<style>
  :root {{--c:{color};}}
  *, *::before, *::after {{box-sizing:border-box;margin:0;padding:0;}}
  body {{
    background:transparent;
    font-family:'Segoe UI',system-ui,-apple-system,sans-serif;
    padding:4px 2px 6px;
    -webkit-font-smoothing:antialiased;
  }}
  /* ── Stats pills ──────────────────────── */
  .stats-row {{
    display:flex;gap:10px;flex-wrap:wrap;margin-bottom:16px;
  }}
  .stat {{
    background:color-mix(in srgb,var(--c) 10%,white);
    border:1.5px solid color-mix(in srgb,var(--c) 30%,white);
    border-radius:12px;padding:10px 18px;text-align:center;
    flex:1;min-width:110px;
  }}
  .stat-val {{
    display:block;font-size:1.25rem;font-weight:800;color:var(--c);line-height:1.2;
  }}
  .stat-lbl {{
    display:block;font-size:.67rem;color:#64748B;font-weight:600;
    letter-spacing:.3px;margin-top:3px;text-transform:uppercase;
  }}
  /* ── Resumen ──────────────────────────── */
  .resumen {{
    background:color-mix(in srgb,var(--c) 7%,white);
    border-left:5px solid var(--c);
    border-radius:0 12px 12px 0;
    padding:16px 20px;
    margin-bottom:18px;
    font-size:.88rem;color:#1E293B;line-height:1.7;
  }}
  .resumen strong {{color:var(--c);}}
  /* ── Sección header ───────────────────── */
  .sec-hdr {{
    display:flex;align-items:center;gap:8px;
    margin-bottom:12px;
  }}
  .sec-line {{flex:1;height:1.5px;background:color-mix(in srgb,var(--c) 20%,#E2E8F0);}}
  .sec-label {{
    font-size:.72rem;font-weight:800;color:var(--c);
    letter-spacing:.8px;text-transform:uppercase;white-space:nowrap;
  }}
  /* ── Grid artículos ───────────────────── */
  .grid {{display:grid;grid-template-columns:repeat(3,1fr);gap:11px;margin-bottom:18px;}}
  .art-card {{
    background:#FFFFFF;
    border-radius:12px;
    padding:14px 15px;
    min-height:112px;
    border-left:3px solid var(--c);
    border:1px solid rgba(226,232,240,.9);
    border-left:3px solid var(--c);
    box-shadow:0 1px 3px rgba(15,23,42,.05),0 4px 12px rgba(15,23,42,.04);
    display:flex;flex-direction:column;gap:6px;
    transition:transform .2s,box-shadow .2s;
  }}
  .art-card:hover {{
    transform:translateY(-2px);
    box-shadow:0 4px 8px rgba(15,23,42,.06),0 12px 28px rgba(15,23,42,.09);
  }}
  .art-top {{display:flex;align-items:center;gap:8px;flex-wrap:wrap;}}
  .art-num {{
    font-family:'Courier New',monospace;
    background:var(--c);color:#fff;
    font-size:.65rem;font-weight:700;
    border-radius:5px;padding:2px 8px;
    flex-shrink:0;white-space:nowrap;
    letter-spacing:.3px;
  }}
  .art-name {{font-size:.82rem;font-weight:700;color:#0F172A;line-height:1.3;}}
  .art-desc {{font-size:.75rem;color:#475569;line-height:1.55;margin-top:2px;}}
  /* ── Botón ley completa ───────────────── */
  .btn-wrap {{display:flex;justify-content:center;padding-top:2px;}}
  .btn-ley {{
    display:inline-flex;align-items:center;gap:10px;
    background:linear-gradient(135deg,var(--c) 0%,color-mix(in srgb,var(--c) 80%,black) 100%);
    color:#fff;font-size:.88rem;font-weight:700;
    border-radius:14px;padding:14px 30px;
    text-decoration:none;
    box-shadow:0 4px 18px color-mix(in srgb,var(--c) 40%,transparent);
    transition:transform .18s,box-shadow .18s,opacity .18s;
    letter-spacing:.2px;
  }}
  .btn-ley:hover {{
    transform:translateY(-2px);
    box-shadow:0 8px 26px color-mix(in srgb,var(--c) 55%,transparent);
    opacity:.92;
  }}
  .btn-icon {{font-size:1rem;}}
  @media(max-width:700px) {{.grid{{grid-template-columns:1fr;}}.stats-row{{flex-direction:column;}}}}
</style>
</head><body>
  {'<div class="stats-row">' + stats_html + '</div>' if stats_html else ''}
  <div class="resumen">{resumen}</div>
  <div class="sec-hdr">
    <span class="sec-line"></span>
    <span class="sec-label">📌 Artículos más importantes</span>
    <span class="sec-line"></span>
  </div>
  <div class="grid">{arts_html}</div>
  <div class="btn-wrap">
    <a class="btn-ley" href="{url_ley}" target="_blank"
       rel="noreferrer noopener" referrerpolicy="no-referrer">
      <span class="btn-icon">📄</span>
      Ver {titulo} completa — Cámara de Diputados
    </a>
  </div>
</body></html>""", height=h, scrolling=False)


tab_sat, tab_isr, tab_iva = st.tabs(["🏛️ SAT", "📊 ISR", "💵 IVA"])

def _render_tab(filtro: str):
    """Renderiza tarjetas de noticias con componentes nativos de Streamlit (sin iframe)."""
    items = [n for n in noticias if filtro in n["cats"]]
    if not items:
        st.markdown(
            "<p style='color:#9CA3AF;font-size:.86rem;padding:8px 0'>"
            "Sin novedades disponibles en este momento.</p>",
            unsafe_allow_html=True,
        )
        return

    for row_start in range(0, min(len(items), 9), 3):
        cols = st.columns(3)
        for j, col in enumerate(cols):
            idx = row_start + j
            if idx >= min(len(items), 9):
                break
            n = items[idx]
            fuente   = n["fuente"]
            accent   = _ACCENT.get(fuente, _ACCENT_DEFAULT)
            abg      = _ACCENT_BG.get(fuente, "#EFF6FF")
            titulo   = _esc(n["titulo"][:115])
            fecha_h  = (f'<span style="font-size:.63rem;color:#94A3B8;margin-left:auto">'
                        f'{_esc(n["fecha"])}</span>') if n.get("fecha") else ""
            with col:
                st.markdown(f"""
<div style="background:#fff;border-radius:14px;
            border:1px solid rgba(226,232,240,.9);border-left:4px solid {accent};
            box-shadow:0 1px 3px rgba(15,23,42,.05),0 5px 18px rgba(15,23,42,.04);
            padding:14px 16px 10px;margin-bottom:4px;min-height:90px;">
  <div style="display:flex;align-items:center;gap:7px;margin-bottom:7px;">
    <span style="width:7px;height:7px;border-radius:50%;background:{accent};
                 display:inline-block;flex-shrink:0;
                 box-shadow:0 0 0 2px {accent}30"></span>
    <span style="font-size:.59rem;font-weight:800;letter-spacing:.7px;
                 text-transform:uppercase;color:{accent};
                 background:{abg};padding:2px 9px;border-radius:20px">{_esc(fuente)}</span>
    {fecha_h}
  </div>
  <div style="font-size:.87rem;color:#0F172A;line-height:1.5;font-weight:500">{titulo}</div>
</div>""", unsafe_allow_html=True)
                st.link_button("↗ Ver fuente", n["link"], use_container_width=True)

with tab_sat:
    _render_tab("SAT")

    # ── Noticias directas del SAT ─────────────────────────────────────────────
    st.markdown("---")
    col_sh, col_sb = st.columns([5, 1])
    with col_sh:
        st.markdown("##### 🏛️ Noticias y trámites del SAT")
    with col_sb:
        if st.button("↻", key="refresh_sat", help="Actualizar noticias SAT"):
            _fetch_sat_noticias.clear()
            st.rerun()

    sat_items, sat_ts = _fetch_sat_noticias()

    # CSS inyectado en la página de Streamlit (NO en iframe)
    st.markdown("""
<style>
.sat-card {
    background:#FFFFFF;
    border-radius:14px;
    border:1px solid rgba(226,232,240,.85);
    border-left:4px solid #B91C1C;
    box-shadow:0 1px 3px rgba(15,23,42,.05),0 5px 18px rgba(15,23,42,.04);
    padding:14px 16px 10px;
    margin-bottom:4px;
    min-height:92px;
}
.sat-badge {
    font-size:.59rem;font-weight:800;letter-spacing:.7px;
    text-transform:uppercase;color:#B91C1C;
    background:#FEE2E2;padding:2px 9px;border-radius:20px;
    display:inline-block;margin-bottom:7px;
}
.sat-title {
    font-size:.87rem;color:#0F172A;line-height:1.5;font-weight:500;
    display:-webkit-box;-webkit-line-clamp:3;-webkit-box-orient:vertical;overflow:hidden;
}
/* Botones ↗ Ver en SAT */
div[data-testid="stLinkButton"] a {
    border-radius:20px !important;
    font-size:.76rem !important;
    font-weight:600 !important;
    padding:5px 14px !important;
}
</style>
""", unsafe_allow_html=True)

    n_show = min(len(sat_items), 12) if sat_items else 0
    if n_show:
        for row_start in range(0, n_show, 3):
            cols = st.columns(3)
            for j, col in enumerate(cols):
                idx = row_start + j
                if idx < n_show:
                    it = sat_items[idx]
                    with col:
                        st.markdown(
                            f'<div class="sat-card">'
                            f'<span class="sat-badge">● SAT</span>'
                            f'<div class="sat-title">{_esc(it["titulo"][:110])}</div>'
                            f'</div>',
                            unsafe_allow_html=True,
                        )
                        st.link_button(
                            "↗ Ver en SAT", it["link"],
                            use_container_width=True,
                        )
    else:
        st.info("No se pudieron obtener noticias del SAT en este momento.")

    st.caption(f"Fuente: sat.gob.mx · Última actualización: {sat_ts}")

    # ── Botones de acceso rápido (st.link_button — abre sin iframe, sin Referer) ──
    col_b1, col_b2, col_b3 = st.columns(3)
    with col_b1:
        st.link_button("🏛️ Portal SAT", "https://www.sat.gob.mx",
                       use_container_width=True)
    with col_b2:
        st.link_button("📄 Diario Oficial de la Federación", "https://www.dof.gob.mx",
                       use_container_width=True)
    with col_b3:
        st.link_button("📚 Leyes — Cámara de Diputados",
                       "https://www.diputados.gob.mx/LeyesBiblio/",
                       use_container_width=True)

with tab_isr:
    _render_tab("ISR")
    st.markdown("---")
    _render_ley_html(
        titulo    = "LISR",
        resumen   = _ISR_RESUMEN,
        articulos = _ISR_ARTS,
        url_ley   = _ISR_URL,
        color     = "#1D4ED8",
        stats     = [
            ("Tasa PM", "30 %"),
            ("Tarifa PF", "0–35 %"),
            ("Dividendos", "+10 %"),
            ("Arts. clave", str(len(_ISR_ARTS))),
        ],
    )

with tab_iva:
    _render_tab("IVA")
    st.markdown("---")
    _render_ley_html(
        titulo    = "LIVA",
        resumen   = _IVA_RESUMEN,
        articulos = _IVA_ARTS,
        url_ley   = _IVA_URL,
        color     = "#059669",
        stats     = [
            ("Tasa general", "16 %"),
            ("Frontera norte", "8 %"),
            ("Tasa cero", "0 %"),
            ("Arts. clave", str(len(_IVA_ARTS))),
        ],
    )

st.markdown("---")
st.caption("AUXILIAR DE REGISTROS · La Sanitaria · v2.0")
