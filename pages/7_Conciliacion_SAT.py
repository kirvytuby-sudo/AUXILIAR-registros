"""AUXILIAR DE REGISTROS — Conciliación SAT (próximamente)"""
import streamlit as st
st.set_page_config(page_title="Conciliación SAT · Auxiliar", page_icon="🔗", layout="wide")
st.markdown("""
<style>
[data-testid="stAppViewContainer"] { background: #dbeafe; }
.header-bar{background:#1E3A8A;padding:18px 28px;border-radius:10px;margin-bottom:20px;}
.header-bar h1{color:#FBCFE8;font-size:1.5rem;margin:0;}
.header-bar p{color:#93C5FD;margin:4px 0 0;font-size:.9rem;}
#MainMenu{visibility:hidden;}footer{visibility:hidden;}
</style>
<div class="header-bar">
  <h1>🔗 Conciliación SAT</h1>
  <p>Conciliación de CFDIs contra registros contables</p>
</div>
""", unsafe_allow_html=True)
st.info("🔜 **Módulo en desarrollo** — Este módulo estará disponible en una próxima versión.")
st.markdown("---")
st.caption("Módulo Conciliación SAT · Próximamente")
