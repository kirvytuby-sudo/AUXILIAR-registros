"""
AUXILIAR DE REGISTROS — Módulo: Pagos Bancarios
Conciliación de nómina BBVA Net Cash. PDFs → Excel.
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
    page_title="Pagos Bancarios · Auxiliar",
    page_icon="💼",
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
    <h1>💼 Pagos Bancarios</h1>
    <p>Conciliación de nómina BBVA Net Cash — PDF → Excel consolidado</p>
</div>
""", unsafe_allow_html=True)

if not _CN_OK:
    st.error("⚠️ No se encontró **conciliacion_nomina.py** — debe estar en la misma carpeta que app.py.")
    st.stop()

# ── Estado de sesión ───────────────────────────────────────────────────────────
if "pb_clasificados" not in st.session_state:
    st.session_state.pb_clasificados = {"nomina": [], "complementos": [], "vacaciones": [], "no_reconocidos": []}
if "pb_resultado_bytes" not in st.session_state:
    st.session_state.pb_resultado_bytes = None
if "pb_tmp" not in st.session_state:
    st.session_state.pb_tmp = tempfile.mkdtemp(prefix="pb_")

TMP = st.session_state.pb_tmp

def _clasificar(path):
    try:
        tipo = cn.detect_template(path)
    except Exception:
        return "no_reconocidos", os.path.basename(path)
    try:
        if tipo == "dispersion":
            meta, _ = cn.parse_dispersion(path)
        else:
            meta, _ = cn.parse_pagos_transferencias(path)
    except Exception:
        return "no_reconocidos", os.path.basename(path)
    desc = (meta.get("descripcion") or "").upper()
    if "VAC" in desc:
        cat = "vacaciones"
    elif "PRESTAMO" in desc or "PRÉSTAMO" in desc:
        cat = "no_reconocidos"   # préstamos van a módulo Préstamos
    elif tipo == "dispersion":
        cat = "nomina"
    else:
        cat = "complementos"
    return cat, meta.get("descripcion") or os.path.basename(path)

def clasificar_uploads(uploaded_files):
    res = {"nomina": [], "complementos": [], "vacaciones": [], "no_reconocidos": []}
    for uf in uploaded_files:
        tmp_path = os.path.join(TMP, uf.name)
        with open(tmp_path, "wb") as f:
            f.write(uf.read())
        cat, desc = _clasificar(tmp_path)
        res[cat].append((uf.name, desc, tmp_path))
    return res

# ── Sección 1: Carga ───────────────────────────────────────────────────────────
st.subheader("1️⃣  Cargar archivos")
col_pdf, col_cat = st.columns([2, 1])

with col_pdf:
    pdfs = st.file_uploader(
        "📂 PDFs de nómina (puedes seleccionar varios)",
        type=["pdf"], accept_multiple_files=True,
    )

with col_cat:
    catalogo_file = st.file_uploader(
        "📊 Catálogo de cuentas (.xlsx)",
        type=["xlsx", "xls"],
    )

if pdfs:
    with st.spinner("Clasificando PDFs..."):
        st.session_state.pb_clasificados = clasificar_uploads(pdfs)
        st.session_state.pb_resultado_bytes = None

catalogo_path = None
if catalogo_file:
    catalogo_path = os.path.join(TMP, catalogo_file.name)
    with open(catalogo_path, "wb") as f:
        f.write(catalogo_file.read())

# ── Sección 2: Selección ───────────────────────────────────────────────────────
clas = st.session_state.pb_clasificados
total_pdfs = sum(len(v) for v in clas.values())

if total_pdfs > 0:
    st.subheader("2️⃣  Seleccionar archivos a procesar")
    CATEGORIAS = [
        ("nomina",       "💼 Nómina principal"),
        ("complementos", "💗 Complementos"),
        ("vacaciones",   "🌴 Vacaciones"),
    ]
    col1, col2, col3 = st.columns(3)
    cols_ui = [col1, col2, col3]
    seleccionados = {}
    for (key, titulo), col in zip(CATEGORIAS, cols_ui):
        items = clas[key]
        with col:
            st.markdown(f"**{titulo}** — {len(items)} archivo(s)")
            if items:
                opciones = [f"{n}  —  {d}" for n, d, _ in items]
                sel = st.multiselect(
                    titulo, options=opciones, default=opciones,
                    label_visibility="collapsed", key=f"sel_{key}",
                )
                seleccionados[key] = [
                    p for (n, d, p), etq in zip(items, opciones) if etq in sel
                ]
            else:
                st.caption("Sin archivos en esta categoría")
                seleccionados[key] = []

    if clas["no_reconocidos"]:
        nombres_nr = ", ".join(n for n, _, _ in clas["no_reconocidos"])
        st.warning(f"⚠️ No reconocidos (posibles préstamos, úsalos en el módulo Préstamos): {nombres_nr}")

    # ── Sección 3: Generar ─────────────────────────────────────────────────────
    st.subheader("3️⃣  Generar reporte")
    todos_pdfs = (
        seleccionados.get("nomina", []) +
        seleccionados.get("complementos", []) +
        seleccionados.get("vacaciones", [])
    )
    n_sel = len(todos_pdfs)
    st.caption(f"Total seleccionado: **{n_sel}** PDF(s)")

    col_btn, col_dl = st.columns([1, 2])
    with col_btn:
        generar = st.button(
            "📦 Generar Excel consolidado",
            disabled=(n_sel == 0 or catalogo_file is None),
            use_container_width=True, type="primary",
        )
        if n_sel == 0:
            st.caption("Selecciona al menos un PDF.")
        if catalogo_file is None:
            st.caption("Carga el catálogo de cuentas.")

    if generar:
        with st.spinner(f"Procesando {n_sel} PDF(s)..."):
            try:
                catalogo = {"empleados": {}, "prestamos": {}}
                if catalogo_path and os.path.isfile(catalogo_path):
                    mapa = cn.load_poliza(catalogo_path)
                    for k in ("empleados", "prestamos"):
                        df = mapa.get(k)
                        if df is not None and not df.empty:
                            catalogo[k] = {idx: str(row["Cuenta"]) for idx, row in df.iterrows()}

                out_path = os.path.join(TMP, "PagosBancarios_Consolidado.xlsx")
                cn.escribir_pagos_bancarios_todo(todos_pdfs, catalogo, out_path)
                with open(out_path, "rb") as f:
                    st.session_state.pb_resultado_bytes = f.read()
                st.success("✅ Excel generado correctamente.")
            except Exception as e:
                import traceback
                st.error(f"❌ Error: {e}")
                with st.expander("Ver detalle"):
                    st.code(traceback.format_exc())

    if st.session_state.pb_resultado_bytes:
        with col_dl:
            st.download_button(
                label="⬇️  Descargar Excel",
                data=st.session_state.pb_resultado_bytes,
                file_name="PagosBancarios_Consolidado.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True, type="secondary",
            )
else:
    st.info("👆 Carga los PDFs de nómina para comenzar.")

st.markdown("---")
st.caption("Módulo Pagos Bancarios · v2.0")
