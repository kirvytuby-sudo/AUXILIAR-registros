"""
AUXILIAR DE REGISTROS — app.py
Usa st.navigation() para registrar páginas correctamente.
"""
import streamlit as st

st.set_page_config(
    page_title="Auxiliar de Registros",
    page_icon="🧾",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Definir páginas (ANTES de la función home) ─────────────────────────────────
p_pagos = st.Page("pages/1_Pagos_Bancarios.py",  title="Pagos Bancarios",    icon="💼")
p_prov  = st.Page("pages/2_Provision_Nomina.py", title="Provisión de Nómina",icon="📋")
p_prest = st.Page("pages/3_Prestamos.py",         title="Préstamos",           icon="💳")
p_vd    = st.Page("pages/4_Ventas_del_Dia.py",   title="Ventas del Día",      icon="⛽")
p_ec    = st.Page("pages/5_Estado_de_Cuenta.py", title="Estado de Cuenta",    icon="🏦")
p_rec   = st.Page("pages/6_Reconciliacion.py",   title="Reconciliación",      icon="📑")
p_sat   = st.Page("pages/7_Conciliacion_SAT.py", title="Conciliación SAT",    icon="🔗")

# ── Página de inicio ───────────────────────────────────────────────────────────
p_cba   = st.Page("pages/8_Conciliacion_Banco_Auxiliar.py", title="Banco vs Auxiliar", icon="🔀")
def pagina_inicio():
    st.markdown("""
<style>
    .header-bar { background:#1E3A8A; padding:22px 28px; border-radius:10px; margin-bottom:24px; }
    .header-bar h1 { color:#FBCFE8; font-size:1.8rem; margin:0; }
    .header-bar p  { color:#93C5FD; margin:6px 0 0; font-size:1rem; }
    .mod-card { background:#F8FAFC; border:2px solid #E2E8F0; border-radius:10px;
                padding:18px 16px 10px; text-align:center; }
    .mod-card.activo { border-color:#1E3A8A; background:#EFF6FF; }
    .mod-card.pronto { border-color:#CBD5E1; opacity:0.7; }
    .mod-icon  { font-size:2rem; margin-bottom:6px; }
    .mod-title { font-weight:700; color:#1E3A8A; font-size:1rem; }
    .mod-desc  { color:#64748B; font-size:0.82rem; margin-top:4px; }
    .badge-soon { background:#F1F5F9; color:#94A3B8; border-radius:12px;
                  padding:2px 10px; font-size:0.75rem; display:inline-block; margin-top:8px; }
    #MainMenu { visibility:hidden; } footer { visibility:hidden; }
</style>
""", unsafe_allow_html=True)

    st.markdown("""
<div class="header-bar">
    <h1>🧾 AUXILIAR DE REGISTROS</h1>
    <p>Sistema contable de La Sanitaria — Selecciona un módulo</p>
</div>
""", unsafe_allow_html=True)

    MODULOS = [
        {"icon":"💼","title":"Pagos Bancarios",    "desc":"Conciliación de nómina BBVA Net Cash — PDF → Excel",       "estado":"activo","page":p_pagos},
        {"icon":"📋","title":"Provisión de Nómina","desc":"XML (CFDI) → Plantilla SINUBE con columnas dinámicas",      "estado":"activo","page":p_prov},
        {"icon":"💳","title":"Préstamos",           "desc":"PDFs de préstamos → Excel con catálogo de cuentas",         "estado":"activo","page":p_prest},
                {"icon":"🔀","title":"Banco vs Auxiliar",  "desc":"Excel banco + auxiliar → conciliación bilateral",                                       "estado":"activo","page":p_cba},
        {"icon":"⛽","title":"Ventas del Día",      "desc":"Reporte de ventas diarias — póliza contable",               "estado":"pronto","page":None},
        {"icon":"🏦","title":"Estado de Cuenta",   "desc":"Análisis y conciliación de estados de cuenta bancarios",    "estado":"pronto","page":None},
        {"icon":"📑","title":"Reconciliación",      "desc":"Reconciliación contable con plantilla SINUBE",              "estado":"pronto","page":None},
        {"icon":"🔗","title":"Conciliación SAT",   "desc":"Conciliación de CFDIs contra registros contables",          "estado":"pronto","page":None},
    ]

    cols = st.columns(4)
    for i, mod in enumerate(MODULOS):
        with cols[i % 4]:
            badge = '<span class="badge-soon">🔜 Próximamente</span>' if mod["estado"] == "pronto" else ""
            st.markdown(f"""
<div class="mod-card {mod['estado']}">
    <div class="mod-icon">{mod['icon']}</div>
    <div class="mod-title">{mod['title']}</div>
    <div class="mod-desc">{mod['desc']}</div>
    {badge}
</div>
""", unsafe_allow_html=True)
            if mod["estado"] == "activo" and mod["page"]:
                st.page_link(mod["page"], label=f"  Abrir {mod['title']}  →", use_container_width=True)
            st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)

    st.markdown("---")
    st.caption("AUXILIAR DE REGISTROS · La Sanitaria · v2.0")

# ── Navegación ─────────────────────────────────────────────────────────────────
p_home = st.Page(pagina_inicio, title="Inicio", icon="🏠", default=True)

pg = st.navigation(
    {"": [p_home],
     "Módulos disponibles": [p_pagos, p_prov, p_prest, p_cba],
     "Próximamente": [p_vd, p_ec, p_rec, p_sat]},
    position="sidebar",
)
pg.run()
