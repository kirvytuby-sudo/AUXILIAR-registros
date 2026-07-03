"""
AUXILIAR DE REGISTROS — Módulo: Préstamos
PDFs de préstamos BBVA Net Cash → Excel con catálogo de cuentas.
"""
import os, tempfile
import streamlit as st

try:
    import conciliacion_nomina as cn
    _CN_OK = True
except ImportError:
    cn = None
    _CN_OK = False

st.set_page_config(
    page_title="Préstamos · Auxiliar",
    page_icon="💳",
    layout="wide",
)

st.markdown("""
<style>
    .header-bar { background:#1E3A8A; padding:18px 28px; border-radius:10px; margin-bottom:20px; }
    .header-bar h1 { color:#FBCFE8; font-size:1.5rem; margin:0; }
    .header-bar p  { color:#93C5FD; margin:4px 0 0; font-size:0.9rem; }
    #MainMenu { visibility:hidden; } footer { visibility:hidden; }
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="header-bar">
    <h1>💳 Préstamos</h1>
    <p>PDFs de préstamos BBVA Net Cash → Excel con catálogo de cuentas</p>
</div>
""", unsafe_allow_html=True)

if not _CN_OK:
    st.error("⚠️ No se encontró **conciliacion_nomina.py** — debe estar en la misma carpeta que app.py.")
    st.stop()

if "pr_resultado_bytes" not in st.session_state:
    st.session_state.pr_resultado_bytes = None
if "pr_tmp" not in st.session_state:
    st.session_state.pr_tmp = tempfile.mkdtemp(prefix="pr_")

TMP = st.session_state.pr_tmp

st.subheader("1️⃣  Cargar archivos")
col_pdf, col_cat = st.columns([2, 1])

with col_pdf:
    pdfs = st.file_uploader(
        "📂 PDFs de préstamos (puedes seleccionar varios)",
        type=["pdf"],
        accept_multiple_files=True,
    )

with col_cat:
    catalogo_file = st.file_uploader(
        "📊 Catálogo de cuentas (.xlsx)",
        type=["xlsx", "xls"],
    )

archivos_validos = []
archivos_rechazados = []

if pdfs:
    with st.spinner("Verificando PDFs..."):
        for uf in pdfs:
            tmp_path = os.path.join(TMP, uf.name)
            with open(tmp_path, "wb") as f: f.write(uf.read())
            try:
                tipo = cn.detect_template(tmp_path)
                meta, _ = (cn.parse_dispersion if tipo == "dispersion" else cn.parse_pagos_transferencias)(tmp_path)
                archivos_validos.append((uf.name, meta.get("descripcion", uf.name), tmp_path))
            except Exception as e: archivos_rechazados.append((uf.name, str(e)))

catalogo_path = None
if catalogo_file:
    catalogo_path = os.path.join(TMP, catalogo_file.name)
    with open(catalogo_path, "wb") as f: f.write(catalogo_file.read())

if archivos_validos:
    st.subheader("2️⃣  Seleccionar PDFs")
    opciones = [f"{n}—{d}" for n, d, _ in archivos_validos]
    sel = st.multiselect("PDFs a procesar:", options=opciones, default=opciones)
    pdfs_sel = [p for (n, d, p), et in zip(archivos_validos, opciones) if et in sel]
    st.subheader("3️⃣  Generar reporte")
    col_btn, col_dl = st.columns([1, 2])
    with col_btn:
        generar = st.button("📦 Generar Excel", disabled=(not pdfs_sel or not catalogo_file), type="primary", use_container_width=True)
    if generar:
        with st.spinner("Procesando..."):
            try:
                catalogo = {"empleados": {}, "prestamos": {}}
                if catalogo_path and os.path.isfile(catalogo_path):
                    mapa = cn.load_poliza(catalogo_path)
                    for k in ("empleados", "prestamos"):
                        df = mapa.get(k)
                        if df is not None and not df.empty:
                            catalogo[k] = {i: str(r["Cuenta"]) for i, r in df.iterrows()}
                out = os.path.join(TMP, "Prestamos_Consolidado.xlsx")
                cn.escribir_pagos_bancarios_todo(pdfs_sel, catalogo, out)
                with open(out, "rb") as f: st.session_state.pr_resultado_bytes = f.read()
                st.success("✅ Excel generado.")
            except Exception as e:
                import traceback; st.error(f"❌ {e}"); st.code(traceback.format_exc())
    if st.session_state.pr_resultado_bytes:
        with col_dl:
            st.download_button("⬇️ Descargar Excel", data=st.session_state.pr_resultado_bytes, file_name="Prestamos_Consolidado.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="secondary", use_container_width=True)
elif pdfs:
    st.warning("⚠️ Ningún PDF pudo leerse.")
else:
    st.info("👆 Cvarga los PDFs de préstamos y el catálogo para comenzar.")

st.markdown("---")
st.caption("Módulo Préstamos · v2.0")
