"""
AUXILIAR DE REGISTROS — Página de inicio
"""
import streamlit as st

st.set_page_config(
    page_title="Auxiliar de Registros",
    page_icon="🧾",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
    .header-bar {
        background: #1E3A8A;
        padding: 22px 28px;
        border-radius: 10px;
        margin-bottom: 24px;
    }
    .header-bar h1 { color: #FBCFE8; font-size: 1.8rem; margin: 0; }
    .header-bar p  { color: #93C5FD; margin: 6px 0 0 0; font-size: 1rem; }

    .mod-card {
        background: #F8FAFC;
        border: 2px solid #E2E8F0;
        border-radius: 10px;
        padding: 20px 16px;
        text-align: center;
        height: 160px;
        display: flex;
        flex-direction: column;
        justify-content: center;
        transition: border-color 0.2s;
    }
    .mod-card.activo  { border-color: #1E3A8A; background: #EFF6FF; }
    .mod-card.pronto  { border-color: #CBD5E1; background: #F8FAFC; opacity: 0.7; }
    .mod-icon  { font-size: 2rem; margin-bottom: 8px; }
    .mod-title { font-weight: 700; color: #1E3A8A; font-size: 1rem; }
    .mod-desc  { color: #64748B; font-size: 0.82rem; margin-top: 4px; }
    .badge-ok   { background:#DCFCE7; color:#166534; border-radius:12px; padding:2px 10px; font-size:0.75rem; }
    .badge-soon { background:#F1F5F9; color:#94A3B8; border-radius:12px; padding:2px 10px; font-size:0.75rem; }

    #MainMenu { visibility: hidden; }
    footer     { visibility: hidden; }
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="header-bar">
    <h1>🧾 AUXILIAR DE REGISTROS</h1>
    <p>Sistema contable de La Sanitaria — Selecciona un módulo en la barra lateral o en las tarjetas de abajo</p>
</div>
""", unsafe_allow_html=True)

MODULOS = [
    {
        "icon": "💼",
        "title": "Pagos Bancarios",
        "desc": "Conciliación de nómina BBVA Net Cash — PDF → Excel",
        "estado": "activo",
        "pagina": "pages/1_Pagos_Bancarios.py",
    },
    {
        "icon": "📋",
        "title": "Provisión de Nómina",
        "desc": "XML (CFDI) → Plantilla SINUBE con columnas dinámicas",
        "estado": "activo",
        "pagina": "pages/2_Provision_Nomina.py",
    },
    {
        "icon": "💳",
        "title": "Préstamos",
        "desc": "PDFs de préstamos → Excel con catálogo de cuentas",
        "estado": "activo",
        "pagina": "pages/3_Prestamos.py",
    },
    {
        "icon": "⛽",
        "title": "Ventas del Día",
        "desc": "Reporte de ventas diarias — póliza contable",
        "estado": "activo",
        "pagina": "pages/4_Ventas_del_Dia.py",
    },
    {
        "icon": "🏦",
        "title": "Estado de Cuenta",
        "desc": "Análisis y conciliación de estados de cuenta bancarios",
        "estado": "pronto",
        "pagina": None,
    },
    {
        "icon": "📑",
        "title": "Reconciliación",
        "desc": "Reconciliación contable con plantilla SINUBE",
        "estado": "pronto",
        "pagina": None,
    },
    {
        "icon": "🔗",
        "title": "Conciliación SAT",
        "desc": "Conciliación de CFDIs contra registros contables",
        "estado": "activo",
        "pagina": "pages/5_Conciliacion_SAT.py",
    },
]

cols = st.columns(4)
for i, mod in enumerate(MODULOS):
    with cols[i % 4]:
        badge = (
            '<span class="badge-ok">✅ Disponible</span>'
            if mod["estado"] == "activo"
            else '<span class="badge-soon">🔜 Próximamente</span>'
        )
        st.markdown(f"""
<div class="mod-card {'activo' if mod['estado'] == 'activo' else 'pronto'}">
    <div class="mod-icon">{mod['icon']}</div>
    <div class="mod-title">{mod['title']}</div>
    <div class="mod-desc">{mod['desc']}</div>
    <div style="margin-top:10px">{badge}</div>
</div>
<br>
""", unsafe_allow_html=True)

st.markdown("---")
st.caption("AUXILIAR DE REGISTROS · La Sanitaria · v2.0  —  Usa la barra lateral izquierda (☰) para navegar entre módulos.")
