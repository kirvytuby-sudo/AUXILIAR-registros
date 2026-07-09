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
[data-testid="stAppViewContainer"] { background: #dbeafe; }
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

# ── Estado de sesión ───────────────────────────────────────────────────────────
if "pr_resultado_bytes" not in st.session_state:
    st.session_state.pr_resultado_bytes = None
if "pr_tmp" not in st.session_state:
    st.session_state.pr_tmp = tempfile.mkdtemp(prefix="pr_")

TMP = st.session_state.pr_tmp

# ── Sección 1: Carga ───────────────────────────────────────────────────────────
st.subheader("1️⃣  Cargar archivos")
col_pdf, col_cat = st.columns([2, 1])

with col_pdf:
    pdfs = st.file_uploader(
        "📂 PDFs de préstamos (puedes seleccionar varios)",
        type=["pdf"],
        accept_multiple_files=True,
        help="Selecciona los PDFs de dispersión de préstamos BBVA Net Cash.",
    )

with col_cat:
    catalogo_file = st.file_uploader(
        "📊 Catálogo de cuentas (.xlsx)",
        type=["xlsx", "xls"],
        help="Archivo Excel con las hojas EMPLEADOS y PRÉSTAMOS.",
    )

# ── Verificar y clasificar PDFs ────────────────────────────────────────────────
archivos_validos = []
archivos_rechazados = []

if pdfs:
    with st.spinner("Verificando PDFs..."):
        for uf in pdfs:
            tmp_path = os.path.join(TMP, uf.name)
            with open(tmp_path, "wb") as f:
                f.write(uf.read())
            try:
                tipo = cn.detect_template(tmp_path)
                if tipo == "dispersion":
                    meta, _ = cn.parse_dispersion(tmp_path)
                else:
                    meta, _ = cn.parse_pagos_transferencias(tmp_path)
                archivos_validos.append((uf.name, meta.get("descripcion", uf.name), tmp_path))
            except Exception as e:
                archivos_rechazados.append((uf.name, str(e)))

catalogo_path = None
if catalogo_file:
    catalogo_path = os.path.join(TMP, catalogo_file.name)
    with open(catalogo_path, "wb") as f:
        f.write(catalogo_file.read())

# ── Sección 2: Selección ───────────────────────────────────────────────────────
if archivos_validos:
    st.subheader("2️⃣  Seleccionar PDFs")
    opciones = [f"{n}  —  {d}" for n, d, _ in archivos_validos]
    sel = st.multiselect(
        "PDFs a procesar:",
        options=opciones,
        default=opciones,
    )
    pdfs_seleccionados = [
        p for (n, d, p), etq in zip(archivos_validos, opciones) if etq in sel
    ]

    if archivos_rechazados:
        with st.expander(f"⚠️ {len(archivos_rechazados)} PDF(s) no se pudieron leer"):
            for nombre, err in archivos_rechazados:
                st.warning(f"**{nombre}**: {err}")

    # ── Sección 3: Generar ─────────────────────────────────────────────────────
    st.subheader("3️⃣  Generar reporte")
    n_sel = len(pdfs_seleccionados)
    st.caption(f"Total seleccionado: **{n_sel}** PDF(s)")

    col_btn, col_dl = st.columns([1, 2])
    with col_btn:
        generar = st.button(
            "📦 Generar Excel de Préstamos",
            disabled=(n_sel == 0 or catalogo_file is None),
            use_container_width=True,
            type="primary",
        )
        if n_sel == 0:
            st.caption("Selecciona al menos un PDF.")
        if catalogo_file is None:
            st.caption("Carga el catálogo de cuentas.")

    if generar:
        with st.spinner(f"Procesando {n_sel} PDF(s)..."):
            try:
                # Catálogo — usa tabla PRESTAMOS para cruzar nombres
                catalogo = {"empleados": {}, "prestamos": {}}
                if catalogo_path and os.path.isfile(catalogo_path):
                    mapa = cn.load_poliza(catalogo_path)
                    for k in ("empleados", "prestamos"):
                        df = mapa.get(k)
                        if df is not None and not df.empty:
                            catalogo[k] = {idx: str(row["Cuenta"]) for idx, row in df.iterrows()}

                out_path = os.path.join(TMP, "Prestamos_Consolidado.xlsx")
                cn.escribir_pagos_bancarios_todo(pdfs_seleccionados, catalogo, out_path)

                with open(out_path, "rb") as f:
                    st.session_state.pr_resultado_bytes = f.read()
                st.success("✅ Excel generado correctamente.")
            except Exception as e:
                import traceback
                st.error(f"❌ Error: {e}")
                with st.expander("Ver detalle"):
                    st.code(traceback.format_exc())

    if st.session_state.pr_resultado_bytes:
        with col_dl:
            st.download_button(
                label="⬇️  Descargar Excel de Préstamos",
                data=st.session_state.pr_resultado_bytes,
                file_name="Prestamos_Consolidado.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                type="secondary",
            )

elif pdfs:
    st.warning("⚠️ Ningún PDF pudo leerse. Verifica que sean archivos válidos de BBVA Net Cash.")
else:
    st.info("👆 Carga los PDFs de préstamos y el catálogo de cuentas para comenzar.")

st.markdown("---")
st.caption("Módulo Préstamos · v2.0")
