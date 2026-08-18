"""
_theme.py — Tema compartido para todas las páginas del Auxiliar de Registros.
Uso:
    import _theme
    _theme.aplicar_header("💼 Pagos Bancarios", "PDF → Excel consolidado")
"""
import streamlit as st

_BASE_CSS = """
<style>
[data-testid="stAppViewContainer"] { background: #dbeafe; }

/* ══════════════════════════════════════════════════
   ANIMACIONES GLOBALES
   ══════════════════════════════════════════════════ */
@keyframes _fi {
  from { opacity:0; }
  to   { opacity:1; }
}
@keyframes _fiu {
  from { opacity:0; transform:translateY(14px); }
  to   { opacity:1; transform:translateY(0); }
}
@keyframes _hdr {
  from { opacity:0; transform:translateY(-10px) scale(.985); }
  to   { opacity:1; transform:translateY(0)     scale(1); }
}
@keyframes _sid {
  from { opacity:0; transform:translateX(-10px); }
  to   { opacity:1; transform:translateX(0); }
}

/* Fade-in general de la página */
section[data-testid="stMain"] > div:first-child {
  animation: _fi .4s ease both;
}

/* Slide-up al cambiar de tab */
div[data-baseweb="tab-panel"] {
  animation: _fiu .32s ease both;
}

/* Alertas (info, success, warning, error) */
div[data-testid="stAlert"] {
  animation: _sid .35s ease both;
}

/* ── Header de módulo ── */
.page-header {
    background: linear-gradient(135deg, #1E3A8A 0%, #2563EB 100%);
    padding: 18px 28px;
    border-radius: 10px;
    margin-bottom: 20px;
    animation: _hdr .45s cubic-bezier(.22,.61,.36,1) both;
    transition: box-shadow .25s;
}
.page-header:hover {
    box-shadow: 0 8px 32px rgba(30,58,138,.35);
}
.page-header h1 { margin: 0; font-size: 1.5rem; font-weight: 700; color: #FBCFE8; letter-spacing: .3px; }
.page-header p  { margin: 4px 0 0; font-size: .9rem; color: #93C5FD; }

/* ── Botones ── */
button[data-testid="baseButton-primary"],
button[data-testid="baseButton-secondary"],
button[data-testid="baseButton-secondaryFormSubmit"] {
    transition: transform .15s ease, box-shadow .15s ease !important;
}
button[data-testid="baseButton-primary"]:hover {
    transform: translateY(-2px) !important;
    box-shadow: 0 4px 14px rgba(30,58,138,.30) !important;
}
button[data-testid="baseButton-secondary"]:hover,
button[data-testid="baseButton-secondaryFormSubmit"]:hover {
    transform: translateY(-1px) !important;
    box-shadow: 0 3px 10px rgba(30,58,138,.18) !important;
}

/* ── Link buttons ── */
div[data-testid="stLinkButton"] a {
    transition: transform .18s ease, box-shadow .18s ease !important;
}
div[data-testid="stLinkButton"] a:hover {
    transform: translateY(-2px) !important;
    box-shadow: 0 4px 14px rgba(30,58,138,.22) !important;
}

/* ── File uploader ── */
section[data-testid="stFileUploadDropzone"] {
    transition: border-color .22s, background .22s, box-shadow .22s !important;
    border-radius: 10px !important;
}
section[data-testid="stFileUploadDropzone"]:hover {
    border-color: #1E3A8A !important;
    background: rgba(30,58,138,.04) !important;
    box-shadow: 0 0 0 3px rgba(30,58,138,.12) !important;
}

/* ── Métricas ── */
div[data-testid="metric-container"] {
    transition: transform .2s ease, box-shadow .2s ease !important;
    border-radius: 10px;
}
div[data-testid="metric-container"]:hover {
    transform: translateY(-2px) !important;
    box-shadow: 0 4px 16px rgba(30,58,138,.12) !important;
}

/* ── Expanders ── */
details[data-testid="stExpander"] > summary {
    transition: background .18s !important;
    border-radius: 8px;
}
details[data-testid="stExpander"] > summary:hover {
    background: rgba(30,58,138,.07) !important;
}

/* ── Tabs ── */
button[data-baseweb="tab"] {
    font-weight: 700 !important;
    font-size: 0.92rem !important;
    letter-spacing: .2px;
    padding: 10px 22px !important;
    transition: color .18s !important;
}
button[data-baseweb="tab"][aria-selected="true"] {
    color: #1E3A8A !important;
}
div[data-baseweb="tab-highlight"] {
    background-color: #1E3A8A !important;
    height: 3px !important;
    border-radius: 3px 3px 0 0 !important;
}

/* ── Scrollbar ── */
::-webkit-scrollbar { width: 6px; height: 6px; }
::-webkit-scrollbar-track { background: #F1F5F9; }
::-webkit-scrollbar-thumb { background: #CBD5E1; border-radius: 3px; }
::-webkit-scrollbar-thumb:hover { background: #94A3B8; }

#MainMenu { visibility: hidden; }
footer    { visibility: hidden; }

/* ── Divisores de sección en el sidebar ── */
[data-testid="stNavSectionHeader"] {
    border-top: 1.5px solid rgba(30, 58, 138, 0.30) !important;
    margin-top: 10px !important;
    padding-top: 8px !important;
}
[data-testid="stNavSectionHeader"] p,
[data-testid="stNavSectionHeader"] span {
    font-size: 0.65rem !important;
    font-weight: 700 !important;
    text-transform: uppercase !important;
    letter-spacing: 0.6px !important;
    color: #1E3A8A !important;
    opacity: 0.85 !important;
}

/* ── Sidebar nav items ── */
[data-testid="stNavLink"] {
    transition: padding-left .18s ease !important;
    border-radius: 8px !important;
}
[data-testid="stNavLink"]:hover {
    padding-left: 16px !important;
}
</style>
"""


def aplicar_header(titulo: str, subtitulo: str = "") -> None:
    """Renderiza el header estándar azul con fondo #dbeafe."""
    sub = f"<p>{subtitulo}</p>" if subtitulo else ""
    st.markdown(
        f"{_BASE_CSS}<div class='page-header'><h1>{titulo}</h1>{sub}</div>",
        unsafe_allow_html=True,
    )


def solo_css() -> None:
    """Aplica solo el fondo y oculta menú/footer, sin header (para páginas con diseño propio)."""
    st.markdown(_BASE_CSS, unsafe_allow_html=True)
