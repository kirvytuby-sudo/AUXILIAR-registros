"""
AUXILIAR DE REGISTROS — Pagos Bancarios
App web Streamlit. Requiere: conciliacion_nomina.py en la misma carpeta.
"""
import os, tempfile, shutil
import streamlit as st

# ── Importar motor contable ────────────────────────────────────────────────────
try:
    import conciliacion_nomina as cn
    _CN_OK = True
except ImportError:
    cn = None
    _CN_OK = False

# ── Configuración de página ────────────────────────────────────────────────────
st.set_page_config(
    page_title="Auxiliar de Registros",
    page_icon="🧾",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ── Estilos ────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
    /* Encabezado */
    .header-bar {
        background: #1E3A8A;
        padding: 18px 28px;
        border-radius: 10px;
        margin-bottom: 20px;
    }
    .header-bar h1 { color: #FBCFE8; font-size: 1.6rem; margin: 0; }
    .header-bar p  { color: #93C5FD; margin: 4px 0 0 0; font-size: 0.9rem; }

    /* Tarjetas de categoría */
    .cat-card {
        background: #F8FAFC;
        border: 1px solid #E2E8F0;
        border-radius: 8px;
        padding: 14px;
        height: 100%;
    }
    .cat-title { font-weight: 700; color: #1E3A8A; margin-bottom: 8px; }

    /* Alerta de error del motor */
    .engine-warning {
        background: #FEF2F2;
        border-left: 4px solid #EF4444;
        padding: 12px 16px;
        border-radius: 4px;
        color: #991B1B;
    }

    /* Ocultar menú hamburguesa */
    #MainMenu { visibility: hidden; }
    footer { visibility: hidden; }
</style>
""", unsafe_allow_html=True)

# ── Encabezado ─────────────────────────────────────────────────────────────────
st.markdown("""
<div class="header-bar">
    <h1>🧾 AUXILIAR DE REGISTROS</h1>
    <p>💼 Módulo: Pagos Bancarios — Conciliación de nómina BBVA Net Cash</p>
</div>
""", unsafe_allow_html=True)

if not _CN_OK:
    st.markdown("""
    <div class="engine-warning">
    ⚠️ <strong>No se encontró conciliacion_nomina.py</strong> —
    el archivo debe estar en la misma carpeta que app.py.
    </div>
    """, unsafe_allow_html=True)
    st.stop()

# ── Inicializar estado de sesión ───────────────────────────────────────────────
if "clasificados" not in st.session_state:
    st.session_state.clasificados = {
        "nomina": [],        # list of (nombre, path_temp)
        "complementos": [],
        "vacaciones": [],
        "prestamos": [],
        "no_reconocidos": [],
    }
if "resultado_bytes" not in st.session_state:
    st.session_state.resultado_bytes = None
if "resultado_nombre" not in st.session_state:
    st.session_state.resultado_nombre = "PagosBancarios.xlsx"
if "tmp_dir" not in st.session_state:
    st.session_state.tmp_dir = tempfile.mkdtemp(prefix="auxiliar_")

TMP = st.session_state.tmp_dir

# ── Función de clasificación ───────────────────────────────────────────────────
def clasificar_uploads(uploaded_files):
    """Guarda uploads en tmp y los clasifica con clasificar_pdf()."""
    resultado = {k: [] for k in ["nomina", "complementos", "vacaciones", "prestamos", "no_reconocidos"]}
    for uf in uploaded_files:
        tmp_path = os.path.join(TMP, uf.name)
        with open(tmp_path, "wb") as f:
            f.write(uf.read())
        try:
            categoria, descripcion = cn.detect_template and _clasificar(tmp_path)
        except Exception:
            categoria, descripcion = "no_reconocidos", uf.name
        if categoria not in resultado:
            categoria = "no_reconocidos"
        resultado[categoria].append((uf.name, descripcion, tmp_path))
    return resultado

def _clasificar(path):
    """Wrapper de clasificar_pdf del módulo gui."""
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
        cat = "prestamos"
    elif tipo == "dispersion":
        cat = "nomina"
    else:
        cat = "complementos"
    return cat, meta.get("descripcion") or os.path.basename(path)

# ══════════════════════════════════════════════════════════════════════════════
# SECCIÓN 1 — CARGA DE ARCHIVOS
# ══════════════════════════════════════════════════════════════════════════════
st.subheader("1️⃣  Cargar archivos")

col_pdf, col_cat = st.columns([2, 1])

with col_pdf:
    pdfs = st.file_uploader(
        "📂 PDFs de nómina (puedes seleccionar varios a la vez)",
        type=["pdf"],
        accept_multiple_files=True,
        help="Selecciona todos los PDFs de dispersión o comprobantes de pago.",
    )

with col_cat:
    catalogo_file = st.file_uploader(
        "📊 Catálogo de cuentas (Excel .xlsx)",
        type=["xlsx", "xls"],
        help="Archivo Excel con las hojas EMPLEADOS y PRÉSTAMOS.",
    )

# Clasificar PDFs cuando se cargan
if pdfs:
    with st.spinner("Clasificando PDFs..."):
        st.session_state.clasificados = clasificar_uploads(pdfs)
        st.session_state.resultado_bytes = None  # limpiar resultado anterior

# Guardar catálogo en tmp
catalogo_path = None
if catalogo_file:
    catalogo_path = os.path.join(TMP, catalogo_file.name)
    with open(catalogo_path, "wb") as f:
        f.write(catalogo_file.read())

# ══════════════════════════════════════════════════════════════════════════════
# SECCIÓN 2 — SELECCIÓN POR CATEGORÍA
# ══════════════════════════════════════════════════════════════════════════════
clas = st.session_state.clasificados
total_pdfs = sum(len(v) for v in clas.values())

if total_pdfs > 0:
    st.subheader("2️⃣  Seleccionar archivos a procesar")

    CATEGORIAS = [
        ("nomina",       "💼 Nómina principal"),
        ("complementos", "💗 Complementos"),
        ("vacaciones",   "🌴 Vacaciones"),
    ]

    col1, col2, col3 = st.columns(3)
    cols = [col1, col2, col3]
    seleccionados = {}

    for (key, titulo), col in zip(CATEGORIAS, cols):
        items = clas[key]
        with col:
            st.markdown(f"**{titulo}** — {len(items)} archivo(s)")
            if items:
                opciones = [f"{nombre}  —  {desc}" for nombre, desc, _ in items]
                sel = st.multiselect(
                    f"Seleccionar de {titulo}",
                    options=opciones,
                    default=opciones,  # todos seleccionados por defecto
                    label_visibility="collapsed",
                    key=f"sel_{key}",
                )
                # Mapear selección a rutas
                seleccionados[key] = [
                    path for (nombre, desc, path), etiqueta in zip(items, opciones)
                    if etiqueta in sel
                ]
            else:
                st.caption("Sin archivos en esta categoría")
                seleccionados[key] = []

    # Préstamos (no se procesan en Pagos Bancarios pero se muestran)
    if clas["prestamos"]:
        st.info(f"ℹ️  {len(clas['prestamos'])} PDF(s) clasificados como **Préstamos** — se procesan en el módulo Préstamos.")

    if clas["no_reconocidos"]:
        nombres_nr = ", ".join(n for n, _, _ in clas["no_reconocidos"])
        st.warning(f"⚠️  No reconocidos: {nombres_nr}")

    # ══════════════════════════════════════════════════════════════════════════
    # SECCIÓN 3 — GENERAR EXCEL
    # ══════════════════════════════════════════════════════════════════════════
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
                # Cargar catálogo
                catalogo = {"empleados": {}, "prestamos": {}}
                if catalogo_path and os.path.isfile(catalogo_path):
                    mapa = cn.load_poliza(catalogo_path)
                    import pandas as pd
                    for k in ("empleados", "prestamos"):
                        df = mapa.get(k)
                        if df is not None and not df.empty:
                            cols_df = [c for c in df.columns if c]
                            if len(cols_df) >= 2:
                                col_nombre = cols_df[0]
                                col_cuenta  = cols_df[1]
                                import unicodedata, re
                                def norm(s):
                                    s = unicodedata.normalize("NFKD", str(s).strip())
                                    s = s.encode("ascii", "ignore").decode("utf-8")
                                    return re.sub(r"\s+", " ", s).upper()
                                catalogo[k] = {
                                    norm(row[col_nombre]): str(row[col_cuenta]).strip()
                                    for _, row in df.iterrows()
                                    if pd.notna(row[col_nombre])
                                }

                # Generar Excel en tmp
                out_path = os.path.join(TMP, "PagosBancarios_Consolidado.xlsx")
                cn.escribir_pagos_bancarios_todo(todos_pdfs, catalogo, out_path)

                with open(out_path, "rb") as f:
                    st.session_state.resultado_bytes = f.read()
                st.session_state.resultado_nombre = "PagosBancarios_Consolidado.xlsx"
                st.success("✅ Excel generado correctamente.")

            except Exception as e:
                import traceback
                st.error(f"❌ Error al generar: {e}")
                with st.expander("Ver detalle del error"):
                    st.code(traceback.format_exc())

    # Botón de descarga (aparece si ya se generó)
    if st.session_state.resultado_bytes:
        with col_dl:
            st.download_button(
                label="⬇️  Descargar Excel",
                data=st.session_state.resultado_bytes,
                file_name=st.session_state.resultado_nombre,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                type="secondary",
            )

else:
    st.info("👆 Carga los PDFs de nómina para comenzar.")

# ── Footer ─────────────────────────────────────────────────────────────────────
st.markdown("---")
st.caption("AUXILIAR DE REGISTROS — Módulo Pagos Bancarios · v1.0")
