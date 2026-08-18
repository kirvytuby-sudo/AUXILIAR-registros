"""
_theme.py — Tema compartido para todas las páginas del Auxiliar de Registros.

Paleta unificada:
  Primary:       #0891B2  (blue-600)
  Primary dark:  #0E7490  (blue-700)  — headers, status bars
  Primary mid:   #06B6D4  (blue-500)  — elementos secundarios, hover
  Light border:  #A5F3FC  (blue-200)  — bordes, separadores
  Pale bg:       #ECFEFF  (blue-50)   — fondos de campos
  Page bg:       #F0F6FF              — fondo de página

Uso:
    import _theme
    _theme.aplicar_header("💼 Pagos Bancarios", "PDF → Excel consolidado")
"""
import streamlit as st

_BASE_CSS = """
<style>
/* ══════════════════════════════════════════════════
   FONDO Y SUPERFICIE
   ══════════════════════════════════════════════════ */
[data-testid="stAppViewContainer"] { background: #ECFEFF; }
[data-testid="stSidebar"]          { background: #06B6D4; }
[data-testid="stSidebar"] * { color: #CFFAFE !important; }
[data-testid="stSidebar"] [data-testid="stNavLink"]:hover {
    background: rgba(255,255,255,.12) !important;
    border-radius: 8px !important;
}
[data-testid="stSidebar"] [data-testid="stNavLink"][aria-current="page"] {
    background: rgba(255,255,255,.20) !important;
    border-radius: 8px !important;
    color: #FFFFFF !important;
    font-weight: 700 !important;
}

/* ══════════════════════════════════════════════════
   ANIMACIONES GLOBALES
   ══════════════════════════════════════════════════ */
@keyframes _fi  { from{opacity:0}                              to{opacity:1} }
@keyframes _fiu { from{opacity:0;transform:translateY(14px)}  to{opacity:1;transform:translateY(0)} }
@keyframes _hdr { from{opacity:0;transform:translateY(-8px) scale(.99)} to{opacity:1;transform:translateY(0) scale(1)} }
@keyframes _sid { from{opacity:0;transform:translateX(-8px)}  to{opacity:1;transform:translateX(0)} }

section[data-testid="stMain"] > div:first-child { animation: _fi .4s ease both; }
div[data-baseweb="tab-panel"]                    { animation: _fiu .30s ease both; }
div[data-testid="stAlert"]                       { animation: _sid .32s ease both; }

/* ══════════════════════════════════════════════════
   HEADER DE MÓDULO
   ══════════════════════════════════════════════════ */
.page-header {
    background: linear-gradient(135deg, #0E7490 0%, #06B6D4 100%);
    padding: 16px 26px;
    border-radius: 10px;
    margin-bottom: 18px;
    box-shadow: 0 4px 16px rgba(29,78,216,.18);
    animation: _hdr .40s cubic-bezier(.22,.61,.36,1) both;
    transition: box-shadow .22s;
}
.page-header:hover { box-shadow: 0 8px 28px rgba(29,78,216,.28); }
.page-header h1 {
    margin: 0;
    font-size: 1.45rem;
    font-weight: 700;
    color: #FFFFFF;
    letter-spacing: .2px;
}
.page-header p {
    margin: 4px 0 0;
    font-size: .88rem;
    color: #A5F3FC;
}

/* ══════════════════════════════════════════════════
   BOTONES
   ══════════════════════════════════════════════════ */
button[data-testid="baseButton-primary"] {
    background: #0891B2 !important;
    border: none !important;
    color: #fff !important;
    font-weight: 600 !important;
    border-radius: 8px !important;
    transition: background .15s, transform .15s, box-shadow .15s !important;
}
button[data-testid="baseButton-primary"]:hover {
    background: #0E7490 !important;
    transform: translateY(-2px) !important;
    box-shadow: 0 4px 14px rgba(29,78,216,.32) !important;
}
button[data-testid="baseButton-secondary"],
button[data-testid="baseButton-secondaryFormSubmit"] {
    border: 1.5px solid #A5F3FC !important;
    color: #0E7490 !important;
    background: #ECFEFF !important;
    border-radius: 8px !important;
    font-weight: 600 !important;
    transition: background .15s, border-color .15s, transform .15s, box-shadow .15s !important;
}
button[data-testid="baseButton-secondary"]:hover,
button[data-testid="baseButton-secondaryFormSubmit"]:hover {
    background: #CFFAFE !important;
    border-color: #06B6D4 !important;
    transform: translateY(-1px) !important;
    box-shadow: 0 3px 10px rgba(59,130,246,.20) !important;
}

/* ══════════════════════════════════════════════════
   FILE UPLOADER
   ══════════════════════════════════════════════════ */
section[data-testid="stFileUploadDropzone"] {
    border: 2px dashed #A5F3FC !important;
    border-radius: 10px !important;
    background: #F8FCFF !important;
    transition: border-color .20s, background .20s, box-shadow .20s !important;
}
section[data-testid="stFileUploadDropzone"]:hover {
    border-color: #06B6D4 !important;
    background: #ECFEFF !important;
    box-shadow: 0 0 0 3px rgba(59,130,246,.12) !important;
}

/* ══════════════════════════════════════════════════
   MÉTRICAS
   ══════════════════════════════════════════════════ */
div[data-testid="metric-container"] {
    background: #FFFFFF;
    border: 1px solid #A5F3FC;
    border-radius: 10px;
    padding: 10px 14px !important;
    transition: transform .18s, box-shadow .18s !important;
}
div[data-testid="metric-container"]:hover {
    transform: translateY(-2px) !important;
    box-shadow: 0 4px 16px rgba(59,130,246,.14) !important;
}
div[data-testid="metric-container"] [data-testid="stMetricLabel"] {
    color: #06B6D4 !important;
    font-weight: 700 !important;
    font-size: .80rem !important;
    text-transform: uppercase;
    letter-spacing: .4px;
}
div[data-testid="metric-container"] [data-testid="stMetricValue"] {
    color: #0E7490 !important;
    font-weight: 800 !important;
}

/* ══════════════════════════════════════════════════
   EXPANDERS
   ══════════════════════════════════════════════════ */
details[data-testid="stExpander"] {
    border: 1px solid #A5F3FC !important;
    border-radius: 10px !important;
    background: #FAFCFF !important;
}
details[data-testid="stExpander"] > summary {
    color: #0E7490 !important;
    font-weight: 600 !important;
    border-radius: 8px !important;
    transition: background .16s !important;
}
details[data-testid="stExpander"] > summary:hover {
    background: rgba(59,130,246,.07) !important;
}

/* ══════════════════════════════════════════════════
   TABS
   ══════════════════════════════════════════════════ */
button[data-baseweb="tab"] {
    font-weight: 700 !important;
    font-size: .90rem !important;
    letter-spacing: .15px;
    padding: 10px 20px !important;
    color: #64748B !important;
    transition: color .16s !important;
}
button[data-baseweb="tab"][aria-selected="true"] {
    color: #0E7490 !important;
}
div[data-baseweb="tab-highlight"] {
    background-color: #0891B2 !important;
    height: 3px !important;
    border-radius: 3px 3px 0 0 !important;
}
div[data-baseweb="tab-border"] {
    background-color: #A5F3FC !important;
}

/* ══════════════════════════════════════════════════
   INPUTS / SELECTBOX / SLIDER
   ══════════════════════════════════════════════════ */
div[data-testid="stTextInput"] input,
div[data-testid="stNumberInput"] input,
div[data-testid="stSelectbox"] > div > div {
    border-color: #A5F3FC !important;
    border-radius: 8px !important;
    transition: border-color .16s, box-shadow .16s !important;
}
div[data-testid="stTextInput"] input:focus,
div[data-testid="stNumberInput"] input:focus {
    border-color: #06B6D4 !important;
    box-shadow: 0 0 0 3px rgba(59,130,246,.15) !important;
}
div[data-testid="stSlider"] [data-testid="stThumbValue"],
div[data-testid="stSlider"] > div > div > div > div {
    background: #0891B2 !important;
}

/* ══════════════════════════════════════════════════
   DATAFRAMES Y TABLAS
   ══════════════════════════════════════════════════ */
[data-testid="stDataFrame"] { border: 1px solid #A5F3FC !important; border-radius: 8px; }
[data-testid="stDataEditor"] { border: 1.5px solid #A5F3FC !important; border-radius: 0 0 8px 8px !important; }
[data-testid="stDataEditor"] td:focus-within {
    outline: 2px solid #06B6D4 !important;
    background: #ECFEFF !important;
}

/* ══════════════════════════════════════════════════
   ALERTAS
   ══════════════════════════════════════════════════ */
div[data-testid="stAlert"][data-baseweb="notification"] {
    border-radius: 8px !important;
    border-left-width: 4px !important;
}

/* ══════════════════════════════════════════════════
   SIDEBAR — NAV ITEMS
   ══════════════════════════════════════════════════ */
[data-testid="stNavSectionHeader"] {
    border-top: 1px solid rgba(191,219,254,.30) !important;
    margin-top: 10px !important;
    padding-top: 8px !important;
}
[data-testid="stNavSectionHeader"] p,
[data-testid="stNavSectionHeader"] span {
    font-size: .64rem !important;
    font-weight: 700 !important;
    text-transform: uppercase !important;
    letter-spacing: .6px !important;
    color: #67E8F9 !important;
    opacity: .90 !important;
}

/* ══════════════════════════════════════════════════
   LINK BUTTONS
   ══════════════════════════════════════════════════ */
div[data-testid="stLinkButton"] a {
    transition: transform .16s, box-shadow .16s !important;
    border-radius: 8px !important;
}
div[data-testid="stLinkButton"] a:hover {
    transform: translateY(-2px) !important;
    box-shadow: 0 4px 14px rgba(37,99,235,.22) !important;
}

/* ══════════════════════════════════════════════════
   SCROLLBAR
   ══════════════════════════════════════════════════ */
::-webkit-scrollbar       { width: 6px; height: 6px; }
::-webkit-scrollbar-track { background: #ECFEFF; }
::-webkit-scrollbar-thumb { background: #A5F3FC; border-radius: 3px; }
::-webkit-scrollbar-thumb:hover { background: #67E8F9; }

/* ══════════════════════════════════════════════════
   OCULTAR MENÚ / FOOTER
   ══════════════════════════════════════════════════ */
#MainMenu { visibility: hidden; }
footer    { visibility: hidden; }
</style>
"""


def aplicar_header(titulo: str, subtitulo: str = "") -> None:
    """Renderiza el header estándar con la paleta unificada."""
    sub = f"<p>{subtitulo}</p>" if subtitulo else ""
    st.markdown(
        f"{_BASE_CSS}<div class='page-header'><h1>{titulo}</h1>{sub}</div>",
        unsafe_allow_html=True,
    )


def solo_css() -> None:
    """Aplica solo el CSS base sin header (para páginas con diseño propio)."""
    st.markdown(_BASE_CSS, unsafe_allow_html=True)
