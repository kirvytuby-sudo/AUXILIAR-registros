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

    /* Tarjetas-botón */
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
        font-family: inherit;
        cursor: pointer;
        transition: border-color 0.2s, background 0.2s, box-shadow 0.2s;
        margin-bottom: 4px;
    }
    div[data-testid="column"] .stButton > button:hover {
        background: #DBEAFE;
        border-color: #2563EB;
        box-shadow: 0 4px 12px rgba(30,58,138,0.18);
    }
    div[data-testid="column"] .stButton > button:active {
        background: #BFDBFE;
    }

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
        "pagina": "pages/1_Pagos_Bancarios.py",
    },
    {
        "icon": "📋",
        "title": "Provisión de Nómina",
        "desc": "XML (CFDI) → Plantilla SINUBE con columnas dinámicas",
        "pagina": "pages/2_Provision_Nomina.py",
    },
    {
        "icon": "💳",
        "title": "Préstamos",
        "desc": "PDFs de préstamos → Excel con catálogo de cuentas",
        "pagina": "pages/3_Prestamos.py",
    },
    {
        "icon": "⛽",
        "title": "Ventas del Día",
        "desc": "Reporte de ventas diarias — póliza contable",
        "pagina": "pages/4_Ventas_del_Dia.py",
    },
    {
        "icon": "📊",
        "title": "Control Despacho vs Ventas",
        "desc": "Concilia despachos contra póliza de ventas — UUID, IVA, IEPS por producto",
        "pagina": "pages/10_Control_Despacho_vs_Ventas.py",
    },
    {
        "icon": "🏦",
        "title": "Depósitos Bancarios",
        "desc": "BBVA, Banorte e Inbursa → póliza de depósitos con cuentas de tránsito",
        "pagina": "pages/11_Depositos_Bancarios.py",
    },
    {
        "icon": "📈",
        "title": "Estado de Cuenta",
        "desc": "Análisis y conciliación de estados de cuenta bancarios",
        "pagina": "pages/9_Estado_de_Cuenta.py",
    },
    {
        "icon": "📑",
        "title": "Reconciliación",
        "desc": "Reconciliación contable con plantilla SINUBE",
        "pagina": "pages/6_Reconciliacion.py",
    },
    {
        "icon": "🔗",
        "title": "Conciliación SAT",
        "desc": "Conciliación de CFDIs contra registros contables",
        "pagina": "pages/5_Conciliacion_SAT.py",
    },
    {
        "icon": "🏛️",
        "title": "Constancia y Opinión SAT",
        "desc": "Genera Constancia Fiscal y Opinión 32-D con e.firma",
        "pagina": "pages/12_Constancia_y_Opinion_SAT.py",
    },
]

cols = st.columns(4)
for i, mod in enumerate(MODULOS):
    with cols[i % 4]:
        label = f"{mod['icon']}\n\n{mod['title']}\n\n{mod['desc']}\n\n✅ Disponible"
        if st.button(label, key=f"mod_{i}", use_container_width=True):
            st.switch_page(mod["pagina"])

st.markdown("---")
st.caption("AUXILIAR DE REGISTROS · La Sanitaria · v2.0  —  Haz clic en cualquier tarjeta o usa la barra lateral (☰) para navegar.")
