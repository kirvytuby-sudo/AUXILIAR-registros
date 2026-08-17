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
_SAT_FALLBACK = [
    {
        "titulo": "Resolución Miscelánea Fiscal — última versión vigente",
        "link":   "https://www.sat.gob.mx/consultas/miscelanea-fiscal",
        "fecha":  "",
    },
    {
        "titulo": "CFDI 4.0 — Complementos y esquemas actualizados",
        "link":   "https://www.sat.gob.mx/consultas/63226/genera-tus-facturas-electronicas",
        "fecha":  "",
    },
    {
        "titulo": "Declaración Anual 2024 — Personas físicas y morales",
        "link":   "https://www.sat.gob.mx/declaracion/83241/declaracion-anual--personas-fisicas",
        "fecha":  "",
    },
    {
        "titulo": "Catálogos SAT para CFDI — claves c_ClaveProdServ y c_Impuesto",
        "link":   "https://www.sat.gob.mx/consultas/63315/catalogo-de-clave-de-producto-o-servicio",
        "fecha":  "",
    },
    {
        "titulo": "Buzón Tributario — notificaciones y trámites electrónicos",
        "link":   "https://www.sat.gob.mx/tramites/28525/abre-tu-buzon-tributario",
        "fecha":  "",
    },
    {
        "titulo": "Constancia de Situación Fiscal — descarga en línea con RFC",
        "link":   "https://www.sat.gob.mx/aplicacion/53027/constancia-de-situacion-fiscal",
        "fecha":  "",
    },
    {
        "titulo": "Opinión de Cumplimiento 32-D — verifica en el portal SAT",
        "link":   "https://www.sat.gob.mx/tramites/16703/obten-tu-opinion-de-cumplimiento-de-obligaciones-fiscales",
        "fecha":  "",
    },
    {
        "titulo": "Servicio de Declaraciones y Pagos (DyP) — presentación mensual",
        "link":   "https://www.sat.gob.mx/declaracion/82923/presenta-tu-declaracion-provisional-o-definitiva-de-impuestos-federales",
        "fecha":  "",
    },
    {
        "titulo": "e.firma (antes FIEL) — renovación y trámite en línea",
        "link":   "https://www.sat.gob.mx/tramites/16703/obten-tu-e.firma-(antes-fiel)",
        "fecha":  "",
    },
    {
        "titulo": "Lista de 69-B LISC — contribuyentes con operaciones inexistentes",
        "link":   "https://www.sat.gob.mx/consultas/76288/consulta-la-lista-del-articulo-69-b",
        "fecha":  "",
    },
    {
        "titulo": "Tarifas y tablas ISR vigentes — subsidio al empleo 2025",
        "link":   "https://www.sat.gob.mx/consultas/14027/tablas-y-tarifas",
        "fecha":  "",
    },
    {
        "titulo": "Pagos provisionales ISR — cálculo y presentación mensual",
        "link":   "https://www.sat.gob.mx/declaracion/82923/presenta-tu-declaracion-provisional",
        "fecha":  "",
    },
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

    # CSS compartido con el SAT (mismo que CARD_CSS pero acento rojo SAT)
    SAT_NEWS_CSS = """
<style>
  *{box-sizing:border-box;margin:0;padding:0;}
  body{background:transparent;font-family:'Segoe UI',Arial,sans-serif;}
  .grid{display:grid;grid-template-columns:repeat(3,1fr);gap:10px;padding:2px;}
  .card{
    background:#fff;border:1.5px solid #FCA5A533;border-radius:11px;
    padding:14px 15px 11px;min-height:110px;
    box-shadow:0 2px 7px rgba(220,38,38,.07);
    display:flex;flex-direction:column;gap:7px;
    transition:box-shadow .2s,border-color .2s,transform .15s;cursor:pointer;
  }
  .card:hover{
    box-shadow:0 5px 18px rgba(220,38,38,.17);
    border-color:#DC2626;transform:translateY(-2px);
  }
  .badge{
    display:inline-block;background:#FEE2E2;color:#B91C1C;
    font-size:.62rem;font-weight:800;border-radius:5px;
    padding:2px 8px;letter-spacing:.4px;width:fit-content;
  }
  .titulo{
    font-size:.84rem;color:#1E293B;line-height:1.5;flex:1;
    text-decoration:none;
  }
  .titulo:hover{color:#DC2626;text-decoration:underline;}
  .footer{display:flex;justify-content:flex-end;margin-top:4px;}
  .btn{
    display:inline-flex;align-items:center;gap:4px;
    background:#FEF2F2;color:#B91C1C;
    font-size:.71rem;font-weight:600;
    border:1.5px solid #FECACA;border-radius:6px;
    padding:3px 10px;text-decoration:none;
    transition:background .15s,border-color .15s;
  }
  .btn:hover{background:#FEE2E2;border-color:#DC2626;}
  .empty{color:#9CA3AF;font-size:.85rem;padding:14px;}
  @media(max-width:700px){.grid{grid-template-columns:1fr;}}
</style>"""

    if sat_items:
        cards_html = ""
        for it in sat_items[:12]:
            href = _esc(it["link"])
            cards_html += f"""
<div class="card" onclick="window.open('{href}','_blank')">
  <span class="badge">SAT</span>
  <a class="titulo" href="{href}" target="_blank">{_esc(it['titulo'])}</a>
  <div class="footer">
    <a class="btn" href="{href}" target="_blank">🔗 Ver en SAT</a>
  </div>
</div>"""
        rows = max(1, (min(len(sat_items), 12) + 2) // 3)
        h = rows * 145 + 24
        components.html(
            f"<!DOCTYPE html><html><head>{SAT_NEWS_CSS}</head>"
            f"<body><div class='grid'>{cards_html}</div></body></html>",
            height=h, scrolling=False,
        )
    else:
        st.info("No se pudieron obtener noticias del SAT en este momento.")

    st.caption(f"Fuente: sat.gob.mx · Última actualización: {sat_ts}")

    # Botón directo al portal SAT
    components.html("""
<style>
  *{box-sizing:border-box;margin:0;padding:0;}
  body{background:transparent;font-family:'Segoe UI',Arial,sans-serif;padding:8px 2px;}
  .wrap{display:flex;gap:12px;flex-wrap:wrap;}
  .btn{
    display:inline-flex;align-items:center;gap:7px;
    color:#fff;font-size:.83rem;font-weight:700;
    border:none;border-radius:9px;
    padding:11px 22px;text-decoration:none;
    transition:opacity .2s,transform .12s;cursor:pointer;
  }
  .btn:hover{opacity:.85;transform:translateY(-1px);}
  .b1{background:#DC2626;box-shadow:0 3px 12px rgba(220,38,38,.35);}
  .b2{background:#7C3AED;box-shadow:0 3px 12px rgba(124,58,237,.35);}
  .b3{background:#D97706;box-shadow:0 3px 12px rgba(217,119,6,.35);}
</style>
<div class="wrap">
  <a class="btn b1" href="https://www.sat.gob.mx/noticias" target="_blank">
    📰 Todas las noticias SAT
  </a>
  <a class="btn b2" href="https://www.sat.gob.mx/consultas/comunicados" target="_blank">
    📢 Comunicados de prensa
  </a>
  <a class="btn b3" href="https://www.dof.gob.mx" target="_blank">
    📄 Diario Oficial de la Federación
  </a>
</div>""", height=58, scrolling=False)

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
