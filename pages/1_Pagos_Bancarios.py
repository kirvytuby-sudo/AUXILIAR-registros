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
    nombre = os.path.basename(path).upper()
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
    # Revisa descripción interna Y nombre del archivo para no perder casos
    # donde el PDF viene con descripción genérica pero el nombre indica vacaciones/préstamo
    buscar = desc + " " + nombre
    if "VAC" in buscar:
        cat = "vacaciones"
    elif "PRESTAMO" in buscar or "PRÉSTAMO" in buscar or "PRÉSTAMO" in nombre:
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

# ── Estilos adicionales ────────────────────────────────────────────────────────
st.markdown("""
<style>
.listbox-header {
    background:#1E3A8A; color:#FBCFE8; font-weight:700;
    padding:5px 10px; border-radius:6px 6px 0 0; font-size:.9rem;
}
.listbox-body {
    border:1.5px solid #1E3A8A; border-top:none;
    border-radius:0 0 6px 6px; background:#fff;
    min-height:140px; padding:6px 8px; margin-bottom:6px;
}
.listbox-item { padding:3px 4px; font-size:.85rem; color:#1e293b; }
.reporte-bar {
    background:#1E3A8A; color:#fff; padding:6px 14px;
    border-radius:6px 6px 0 0; font-weight:700; font-size:.9rem;
    display:flex; align-items:center; gap:8px;
}
.log-bar {
    background:#0D1117; color:#93C5FD; padding:5px 12px;
    border-radius:6px 6px 0 0; font-weight:700; font-size:.85rem;
}
div[data-testid="stButton"] > button[kind="secondary"] {
    background:#FFF0F5; border:1px solid #1E3A8A;
    color:#1E3A8A; font-size:.82rem;
}
</style>
""", unsafe_allow_html=True)

# ── Área de carga ──────────────────────────────────────────────────────────────
_c_pdf, _c_cat = st.columns([3, 2])
with _c_pdf:
    pdfs = st.file_uploader(
        "📂 PDFs...", type=["pdf"], accept_multiple_files=True, label_visibility="collapsed",
        help="Selecciona uno o varios PDFs de nómina BBVA Net Cash",
    )
with _c_cat:
    catalogo_file = st.file_uploader(
        "📊 Catálogo de cuentas (.xlsx)", type=["xlsx","xls"], label_visibility="visible",
    )

# Clasificar al subir
if pdfs:
    with st.spinner("Clasificando PDFs..."):
        st.session_state.pb_clasificados = clasificar_uploads(pdfs)
        st.session_state.pb_resultado_bytes = None
        st.session_state.pb_log = ["PDFs cargados y clasificados."]

catalogo_path = None
if catalogo_file:
    catalogo_path = os.path.join(TMP, catalogo_file.name)
    with open(catalogo_path, "wb") as f:
        f.write(catalogo_file.read())

# ── Inicializar log y selecciones ─────────────────────────────────────────────
if "pb_log" not in st.session_state:
    st.session_state.pb_log = []
if "pb_todos_nom" not in st.session_state:
    st.session_state.pb_todos_nom = True
if "pb_todos_comp" not in st.session_state:
    st.session_state.pb_todos_comp = True
if "pb_todos_vac" not in st.session_state:
    st.session_state.pb_todos_vac = True

clas = st.session_state.pb_clasificados

# ── Tres listboxes ─────────────────────────────────────────────────────────────
_col1, _col2, _col3 = st.columns(3)

def _render_lista(col, key, titulo, emoji):
    items = clas.get(key, [])
    todos_key = f"pb_todos_{key[:3]}"
    with col:
        st.markdown(f'<div class="listbox-header">{emoji} {titulo}</div>', unsafe_allow_html=True)
        # Multiselect como listbox
        opciones = [f"{n}  —  {d}" for n, d, _ in items] if items else []
        default  = opciones if st.session_state.get(todos_key, True) else []
        sel = st.multiselect(
            titulo, options=opciones, default=default,
            label_visibility="collapsed", key=f"ms_{key}",
        )
        # Botones Todos / Ninguno
        _bt, _bn = st.columns(2)
        with _bt:
            if st.button("✔ Todos",  key=f"btn_todos_{key}", use_container_width=True):
                st.session_state[todos_key] = True
                st.rerun()
        with _bn:
            if st.button("✖ Ninguno", key=f"btn_ning_{key}", use_container_width=True):
                st.session_state[todos_key] = False
                st.rerun()
        # Rutas seleccionadas
        paths_sel = [p for (n, d, p), etq in zip(items, opciones) if etq in sel]
        return paths_sel

_sel_nom  = _render_lista(_col1, "nomina",       "Nómina principal", "💼")
_sel_comp = _render_lista(_col2, "complementos", "Complementos",     "💗")
_sel_vac  = _render_lista(_col3, "vacaciones",   "Vacaciones",       "🌴")

# ── Botones de proceso individuales ───────────────────────────────────────────
_pb1, _pb2, _pb3 = st.columns(3)

def _procesar(pdfs_lista, etiqueta, out_nombre):
    if not pdfs_lista:
        st.session_state.pb_log.append(f"⚠ Sin PDFs seleccionados en {etiqueta}.")
        return
    if not catalogo_path:
        st.session_state.pb_log.append("⚠ Carga el catálogo de cuentas primero.")
        return
    try:
        catalogo = {"empleados": {}, "prestamos": {}}
        mapa = cn.load_poliza(catalogo_path)
        for k in ("empleados", "prestamos"):
            df = mapa.get(k)
            if df is not None and not df.empty:
                catalogo[k] = {idx: str(row["Cuenta"]) for idx, row in df.iterrows()}
        out_path = os.path.join(TMP, out_nombre)
        cn.escribir_pagos_bancarios_todo(pdfs_lista, catalogo, out_path)
        with open(out_path, "rb") as f:
            st.session_state.pb_resultado_bytes = f.read()
        st.session_state.pb_resultado_nombre = out_nombre
        st.session_state.pb_log.append(f"✅ {etiqueta}: {len(pdfs_lista)} PDF(s) → {out_nombre}")
    except Exception as e:
        import traceback
        st.session_state.pb_log.append(f"❌ Error en {etiqueta}: {e}")
        st.session_state.pb_log.append(traceback.format_exc())

with _pb1:
    if st.button("▶ Conciliar seleccionados", use_container_width=True, key="btn_conc"):
        with st.spinner("Procesando nómina..."):
            _procesar(_sel_nom, "Nómina", "Nomina_Conciliada.xlsx")
        st.rerun()

with _pb2:
    if st.button("▶ Procesar complementos", use_container_width=True, key="btn_comp"):
        with st.spinner("Procesando complementos..."):
            _procesar(_sel_comp, "Complementos", "Complementos.xlsx")
        st.rerun()

with _pb3:
    if st.button("▶ Procesar vacaciones", use_container_width=True, key="btn_vac"):
        with st.spinner("Procesando vacaciones..."):
            _procesar(_sel_vac, "Vacaciones", "Vacaciones.xlsx")
        st.rerun()

# ── Generar TODO ───────────────────────────────────────────────────────────────
_todos_pdfs = _sel_nom + _sel_comp + _sel_vac
_hay_cat    = catalogo_path is not None
if st.button(
    f"⚙ Generar TODO en un solo Excel  (usa lo seleccionado en las 3 listas)  — {len(_todos_pdfs)} PDF(s)",
    disabled=(len(_todos_pdfs) == 0 or not _hay_cat),
    use_container_width=True, type="primary", key="btn_todo",
):
    with st.spinner(f"Procesando {len(_todos_pdfs)} PDF(s)..."):
        _procesar(_todos_pdfs, "TODO", "PagosBancarios_Consolidado.xlsx")
    st.rerun()

if not _hay_cat:
    st.caption("⚠️ Carga el catálogo de cuentas para habilitar los botones.")

st.divider()

# ── Reporte ────────────────────────────────────────────────────────────────────
_nombre_rep = st.session_state.get("pb_resultado_nombre", "— (genera el reporte primero)")
st.markdown(f'<div class="reporte-bar">📋 Reporte: {_nombre_rep}</div>', unsafe_allow_html=True)

_r_bytes = st.session_state.get("pb_resultado_bytes")
if _r_bytes:
    _rc1, _rc2 = st.columns([1, 1])
    with _rc1:
        st.download_button(
            "📂 Abrir / Descargar Excel",
            data=_r_bytes,
            file_name=_nombre_rep,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True, key="dl_rep",
        )
    # Vista previa
    try:
        import pandas as pd, io
        _df_prev = pd.read_excel(io.BytesIO(_r_bytes), nrows=50)
        st.dataframe(_df_prev, use_container_width=True, height=200)
    except Exception:
        st.info("Vista previa no disponible.")
else:
    st.markdown(
        '<div style="text-align:center;padding:30px;color:#94a3b8;">'
        '📋 Genera el reporte para ver la vista previa aquí</div>',
        unsafe_allow_html=True,
    )

st.divider()

# ── Registro de actividad ──────────────────────────────────────────────────────
st.markdown('<div class="log-bar">📜 Registro de actividad:</div>', unsafe_allow_html=True)
_log_txt = "\n".join(st.session_state.pb_log[-30:]) if st.session_state.pb_log else "(sin actividad)"
st.code(_log_txt, language=None)

if st.session_state.pb_log:
    if st.button("🗑 Limpiar log", key="btn_limpiar_log"):
        st.session_state.pb_log = []
        st.rerun()

if clas.get("no_reconocidos"):
    nombres_nr = ", ".join(n for n, _, _ in clas["no_reconocidos"])
    st.warning(f"⚠️ No reconocidos (posibles préstamos — úsalos en el módulo Préstamos): {nombres_nr}")
