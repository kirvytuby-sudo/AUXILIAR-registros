"""
AUXILIAR DE REGISTROS — Punto de entrada con st.navigation()
Controla el sidebar: labels correctos, orden definido, sin páginas duplicadas.
"""
import streamlit as st

pg = st.navigation(
    {
        "": [
            st.Page("pages/0_Inicio.py",                      title="Inicio",                      icon="🏠"),
        ],
        "Nómina": [
            st.Page("pages/1_Pagos_Bancarios.py",             title="Pagos Bancarios",             icon="💼"),
            st.Page("pages/2_Provision_Nomina.py",            title="Provisión de Nómina",         icon="📋"),
            st.Page("pages/3_Prestamos.py",                   title="Préstamos",                   icon="💳"),
            st.Page("pages/14_Poliza_Nomina.py",              title="Póliza de Nómina",            icon="📄"),
        ],
        "Ventas": [
            st.Page("pages/4_Ventas_del_Dia.py",              title="Ventas del Día",              icon="⛽"),
            st.Page("pages/10_Control_Despacho_vs_Ventas.py", title="Control Despacho vs Ventas",  icon="📊"),
        ],
        "Banco": [
            st.Page("pages/11_Depositos_Bancarios.py",        title="Depósitos Bancarios",         icon="🏦"),
            st.Page("pages/9_Estado_de_Cuenta.py",            title="Estado de Cuenta",            icon="📈"),
            st.Page("pages/8_Conciliacion_Banco_Auxiliar.py", title="Conciliación Banco/Auxiliar", icon="🔀"),
        ],
        "SAT / IMSS / Contabilidad": [
            st.Page("pages/5_Conciliacion_SAT.py",            title="Conciliación SAT",            icon="🔗"),
            st.Page("pages/6_Reconciliacion.py",              title="Reconciliación",              icon="📑"),
            st.Page("pages/12_Constancia_y_Opinion_SAT.py",   title="Constancia y Opinión SAT",    icon="🏛️"),
            st.Page("pages/15_Opinion_IMSS.py",               title="Opinión IMSS",                icon="🏥"),
            st.Page("pages/16_Buzon_SAT.py",                 title="Buzón SAT",                   icon="🗂"),
        ],
    },
    position="sidebar",
)
pg.run()
