"""
pages/13_Impuestos_PF.py — Módulo Impuestos Persona Física
Calcula ISR (Art.106 LISR) e IVA (Art.5-D LIVA) para Régimen 612.
Genera libro Excel completo: PAGOS AL SAT, CÁLCULO LEY, ISR, IVA,
GASTOS DEDUCIBLES, hojas mensuales, GASTOS EXCLUIDOS, AMARRE FACTURAS,
TARIFA ISR.
"""
import os, sys, tempfile, shutil
import streamlit as st

st.set_page_config(
    page_title="Impuestos PF | Auxiliar de Registros",
    page_icon="📊",
    layout="wide",
)

import _theme
_theme.aplicar_header(
    "📊 Impuestos PF — Régimen 612",
    "ISR Art.106 LISR (pago provisional acumulado) · IVA Art.5-D LIVA (mensual)"
)

# ── Importar motor ────────────────────────────────────────────────────────────
_APP_DIR = os.path.join(os.path.dirname(__file__), '..', '_app')
if _APP_DIR not in sys.path:
    sys.path.insert(0, os.path.abspath(_APP_DIR))

try:
    from calc_impuestos_pf import calcular_impuestos, generar_excel
    _ENGINE_OK = True
except ImportError as _err:
    _ENGINE_OK = False
    st.error(f"❌ No se encontró el motor `calc_impuestos_pf.py` en `_app/`.\n\n`{_err}`")
    st.stop()

# ── Estado de sesión ──────────────────────────────────────────────────────────
if 'ipf_resultado' not in st.session_state:
    st.session_state.ipf_resultado      = None
if 'ipf_resultado_nombre' not in st.session_state:
    st.session_state.ipf_resultado_nombre = None
if 'ipf_log' not in st.session_state:
    st.session_state.ipf_log = []
if 'ipf_tmp' not in st.session_state:
    st.session_state.ipf_tmp = tempfile.mkdtemp(prefix='ipf_')

TMP = st.session_state.ipf_tmp

# ── Sección 1: Configuración ──────────────────────────────────────────────────
st.subheader("1️⃣  Configuración del contribuyente")
c1, c2, c3 = st.columns([2, 3, 1])
with c1:
    rfc = st.text_input("RFC", value="GOFC7909293X6",
                        max_chars=13, help="RFC de la persona física").strip().upper()
with c2:
    nombre = st.text_input("Nombre / Razón social",
                           value="", placeholder="Carlos González Flores").strip()
with c3:
    import datetime
    year = st.number_input("Año fiscal", min_value=2020, max_value=2030,
                           value=datetime.date.today().year - 1, step=1)

if not nombre:
    nombre = rfc

# ── Sección 2: XMLs ───────────────────────────────────────────────────────────
st.subheader("2️⃣  Archivos XML (CFDI)")
col_e, col_r = st.columns(2)
with col_e:
    st.caption("**Emitidos** — facturas propias (ingresos)")
    emit_files = st.file_uploader(
        "XMLs emitidos", type=["xml"], accept_multiple_files=True,
        key="ipf_emit", label_visibility="collapsed")
with col_r:
    st.caption("**Recibidos** — gastos / proveedores")
    rec_files = st.file_uploader(
        "XMLs recibidos", type=["xml"], accept_multiple_files=True,
        key="ipf_rec", label_visibility="collapsed")

# ── Sección 3: Estados de cuenta ─────────────────────────────────────────────
st.subheader("3️⃣  Estados de cuenta bancarios")
col_bbva, col_ban, col_inb, col_ext = st.columns(4)
with col_bbva:
    st.caption("**BBVA** (PDF)")
    bbva_files = st.file_uploader(
        "PDFs BBVA", type=["pdf"], accept_multiple_files=True,
        key="ipf_bbva", label_visibility="collapsed")
with col_ban:
    st.caption("**Banorte** (Excel)")
    ban_files = st.file_uploader(
        "Excel Banorte", type=["xlsx","xls"], accept_multiple_files=True,
        key="ipf_ban", label_visibility="collapsed")
with col_inb:
    st.caption("**Inbursa** (Excel)")
    inb_files = st.file_uploader(
        "Excel Inbursa", type=["xlsx","xls"], accept_multiple_files=True,
        key="ipf_inb", label_visibility="collapsed")
with col_ext:
    st.caption("**Otro banco** (Excel)")
    ext_nombre = st.text_input("Nombre del banco", value="Otro", key="ipf_ext_nombre")
    ext_files = st.file_uploader(
        "Excel otro banco", type=["xlsx","xls"], accept_multiple_files=True,
        key="ipf_ext", label_visibility="collapsed")

# ── Helper: guardar uploads a tmp ─────────────────────────────────────────────
def _save_uploads(files, subdir):
    """Guarda archivos en TMP/subdir y retorna lista de rutas."""
    if not files:
        return []
    d = os.path.join(TMP, subdir)
    os.makedirs(d, exist_ok=True)
    paths = []
    for f in files:
        dest = os.path.join(d, f.name)
        with open(dest, 'wb') as fp:
            fp.write(f.getbuffer())
        paths.append(dest)
    return paths

# ── Resumen de archivos cargados ──────────────────────────────────────────────
n_emit = len(emit_files or []); n_rec = len(rec_files or [])
n_bbva = len(bbva_files or []); n_ban = len(ban_files or [])
n_inb  = len(inb_files or []);  n_ext = len(ext_files or [])

st.info(
    f"📥 **Cargados:** "
    f"{n_emit} XML emit. · {n_rec} XML recib. · "
    f"{n_bbva} PDF BBVA · {n_ban} Banorte · {n_inb} Inbursa · {n_ext} Otro"
)

verificar_vigencia = st.toggle(
    "🔍 Verificar vigencia de CFDIs en el SAT (excluye cancelados del cálculo)",
    value=True,
    help="Consulta el servicio del SAT para cada CFDI emitido. Los cancelados se marcan en rojo y NO se suman al ISR/IVA."
)

if n_emit == 0 and n_rec == 0:
    st.warning("⚠️ Sube al menos XMLs emitidos **o** recibidos para continuar.")
    st.stop()

# ── Botón generar ─────────────────────────────────────────────────────────────
st.markdown("---")
generar = st.button(
    "⚡  GENERAR CÁLCULO DE IMPUESTOS",
    type="primary", use_container_width=True)

if generar:
    st.session_state.ipf_log = []
    log_msgs = st.session_state.ipf_log

    def _log(msg):
        log_msgs.append(msg)

    with st.spinner("Calculando ISR e IVA..."):
        try:
            # Guardar en tmp
            emit_paths = _save_uploads(emit_files, 'emit')
            rec_paths  = _save_uploads(rec_files,  'rec')
            bbva_paths = _save_uploads(bbva_files, 'bbva')
            ban_paths  = _save_uploads(ban_files,  'banorte')
            inb_paths  = _save_uploads(inb_files,  'inbursa')
            ext_paths  = _save_uploads(ext_files,  'ext')

            results, emit_log, ded_log, no_ded_log = calcular_impuestos(
                rfc=rfc,
                year=year,
                carpetas_xml_emit=emit_paths,
                carpetas_xml_rec=rec_paths,
                pdfs_bbva=bbva_paths or None,
                excels_banorte=ban_paths or None,
                excels_inbursa=inb_paths or None,
                excels_extra={ext_nombre: ext_paths} if ext_paths else None,
                verificar_vigencia=verificar_vigencia,
                log=_log,
            )

            out_path = os.path.join(TMP, f'Impuestos_{rfc}_{year}.xlsx')
            generar_excel(results, emit_log, ded_log, no_ded_log,
                          rfc, nombre, year, out_path, log=_log)

            with open(out_path, 'rb') as fp:
                st.session_state.ipf_resultado       = fp.read()
            st.session_state.ipf_resultado_nombre = f'Impuestos_{rfc}_{year}.xlsx'
            st.success("✅ Cálculo completado")

        except Exception as ex:
            import traceback
            st.error(f"❌ Error: {ex}")
            st.code(traceback.format_exc())

# ── Log de proceso ────────────────────────────────────────────────────────────
if st.session_state.ipf_log:
    with st.expander("📋 Log de proceso", expanded=False):
        for msg in st.session_state.ipf_log:
            st.text(msg)

# ── Descarga ──────────────────────────────────────────────────────────────────
if st.session_state.ipf_resultado:
    st.markdown("---")
    st.subheader("4️⃣  Resultado")

    st.download_button(
        label="⬇️  Descargar Excel de Impuestos",
        data=st.session_state.ipf_resultado,
        file_name=st.session_state.ipf_resultado_nombre,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        type="primary",
    )

    # Vista previa de PAGOS AL SAT
    try:
        import io, openpyxl, pandas as pd
        wb = openpyxl.load_workbook(
            io.BytesIO(st.session_state.ipf_resultado), data_only=True, read_only=True)
        if 'PAGOS AL SAT' in wb.sheetnames:
            ws = wb['PAGOS AL SAT']
            rows = []
            for ri, row in enumerate(ws.iter_rows(values_only=True)):
                if ri < 4: continue                      # saltar encabezados
                if all(v is None for v in row): continue
                rows.append([str(v or '') for v in row])
                if len(rows) >= 14: break
            wb.close()
            if rows:
                st.caption("📊 Vista previa — PAGOS AL SAT")
                hdrs = ['MES','ING.MES','DED.MES','BASE ACUM','ISR TARIFA',
                        'RET.ISR','ISR PAGAR','IVA CAUS','IVA ACRED','IVA RET',
                        'IVA PAGAR','ISR','IVA','TOTAL SAT']
                df = pd.DataFrame(rows, columns=hdrs[:len(rows[0])])
                st.dataframe(df, use_container_width=True, hide_index=True)
    except Exception:
        pass
