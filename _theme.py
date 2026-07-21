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
.page-header {
    background: linear-gradient(135deg, #1E3A8A 0%, #2563EB 100%);
    padding: 18px 28px;
    border-radius: 10px;
    margin-bottom: 20px;
}
.page-header h1 { margin: 0; font-size: 1.5rem; font-weight: 700; color: #FBCFE8; letter-spacing: .3px; }
.page-header p  { margin: 4px 0 0; font-size: .9rem; color: #93C5FD; }
#MainMenu { visibility: hidden; }
footer    { visibility: hidden; }
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
