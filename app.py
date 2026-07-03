"""
AUXILIAR DE REGISTROS — Página de inicio
"""
import streamlit as st
import streamlit.components.v1 as components

st.set_page_config(
    page_title="Auxiliar de Registros",
    page_icon="🧾",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
    .header-bar {
        background: #1E3A8A; padding: 22px 28px;
        border-radius: 10px; margin-bottom: 24px;
    }
    .header-bar h1 { color: #FBCFE8; font-size: 1.8rem; margin: 0; }
    .header-bar p  { color: #93C5FD; margin: 6px 0 0; font-size: 1rem; }

    .mod-card {
        background: #F8FAFC; border: 2px solid #E2E8F0;
        border-radius: 10px; padding: 20px 16px 12px;
        text-align: center; min-height: 130px;
        display: flex; flex-direction: column; justify-content: center;
    }
    .mod-card.activo { border-color: #1E3A8A; background: #EFF6FF; }
    .mod-card.pronto { border-color: #CBD5E1; opacity: 0.7; }
    .mod-icon  { font-size: 2rem; margin-bottom: 6px; }
    .mod-title { font-weight: 700; color: #1E3A8A; font-size: 1rem; }
    .mod-desc  { color: #64748B; font-size: 0.82rem; margin-top: 4px; }
    .badge-soon {
        background: #F1F5F9; color: #94A3B8;
        border-radius: 12px; padding: 2px 10px; font-size: 0.75rem;
        display: inline-block; margin-top: 8px;
    }
    #MainMenu { visibility: hidden; }
    footer     { visibility: hidden; }
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="header-bar">
    <h1>🧾 AUXILIAR DE REGISTROS</h1>
    <p>Sistema contable de La Sanitaria — Selecciona un módulo</p>
</div>
""", unsafe_allow_html=True)

MODULOS = [
    {"icon": "💼", "title": "Pagos Bancarios",
     "desc": "Conciliación de nómina BBVA Net Cash — PDF → Excel",
     "estado": "activo", "url": "Pagos_Bancarios"},
    {"icon": "📋", "title": "Provisión de Nómina",
     "desc": "XML (CFDI) → Plantilla SINUBE con columnas dinámicas",
     "estado": "activo", "url": "Provision_Nomina"},
    {"icon": "💳", "title": "Préstamos",
     "desc": "PDFs de préstamos → Excel con catálogo de cuentas",
     "estado": "activo", "url": "Prestamos"},
    {"icon": "⛽", "title": "Ventas del Día",
     "desc": "Reporte de ventas diarias — póliza contable",
     "estado": "pronto", "url": None},
    {"icon": "🏦", "title": "Estado de Cuenta",
     "desc": "Análisis y conciliación de estados de cuenta bancarios",
     "estado": "pronto", "url": None},
    {"icon": "📑", "title": "Reconciliación",
     "desc": "Reconciliación contable con plantilla SINUBE",
     "estado": "pronto", "url": None},
    {"icon": "🔗", "title": "Conciliación SAT",
     "desc": "Conciliación de CFDIs contra registros contables",
     "estado": "pronto", "url": None},
]

# Track which button was clicked BEFORE rendering components
ir_a = None

cols = st.columns(4)
for i, mod in enumerate(MODULOS):
    with cols[i % 4]:
        st.markdown(f"""
<div class="mod-card {mod['estado']}">
    <div class="mod-icon">{mod['icon']}</div>
    <div class="mod-title">{mod['title']}</div>
    <div class="mod-desc">{mod['desc']}</div>
    {'<span class="badge-soon">🔜 Próximamente</span>' if mod['estado'] == 'pronto' else ''}
</div>
""", unsafe_allow_html=True)
        if mod["estado"] == "activo":
            if st.button("Abrir →", key=f"nav_{i}", use_container_width=True, type="primary"):
                ir_a = mod["url"]
        else:
            st.markdown("<div style='height:38px'></div>", unsafe_allow_html=True)

# Navigate via JS after all widgets rendered
if ir_a:
    components.html(
        f"<script>window.parent.location.href = '/{ir_a}';</script>",
        height=0,
    )

st.markdown("---")
st.caption("AUXILIAR DE REGISTROS · La Sanitaria · v2.0  —  Usa también la barra lateral (☰) para navegar.")
