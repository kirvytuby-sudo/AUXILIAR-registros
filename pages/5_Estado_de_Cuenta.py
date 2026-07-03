"""
pages/5_Estado_de_Cuenta.py — Módulo Estado de Cuenta
"""
import streamlit as st
import tempfile
import os
import pandas as pd
from datetime import date

st.set_page_config(
    page_title="Estado de Cuenta | Auxiliar de Registros",
    page_icon="🏦",
    layout="wide",
)

st.markdown("""
<style>
[data-testid="stAppViewContainer"] { background: #F8FAFC; }
.block-container { padding-top: 1.5rem; padding-bottom: 2rem; }
.ec-header {
    background: linear-gradient(135deg, #1E3A8A 0%, #2563EB 100%);
    color: white; border-radius: 10px;
    padding: 1.2rem 1.8rem; margin-bottom: 1.5rem;
}
.ec-header h1 { margin: 0; font-size: 1.6rem; font-weight: 700; letter-spacing: .5px; }
.ec-header p  { margin: .3rem 0 0; opacity: .8; font-size: .9rem; }
.ec-card {
    background: white; border-radius: 8px; padding: 1.2rem 1.5rem;
    box-shadow: 0 1px 6px rgba(0,0,0,.07); margin-bottom: 1rem;
}
.ec-stat {
    background: #EEF2FF; border-radius: 8px; padding: .9rem 1.2rem;
    text-align: center; margin-bottom: .5rem;
}
.ec-stat .val { font-size: 1.4rem; font-weight: 700; color: #1E3A8A; }
.ec-stat .lbl { font-size: .75rem; color: #64748B; margin-top: .1rem; }
.dep-val { color: #1E6FBF; font-weight: 600; }
.ret-val { color: #E14B3D; font-weight: 600; }
.sal-val { font-weight: 600; }
.diff-ok  { background: #DCFCE7; border-radius: 6px; padding: .4rem .9rem; color: #166534; font-weight: 600; }
.diff-err { background: #FEE2E2; border-radius: 6px; padding: .4rem .9rem; color: #991B1B; font-weight: 600; }
.diff-na  { background: #F1F5F9; border-radius: 6px; padding: .4rem .9rem; color: #475569; }
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="ec-header">
  <h1>🏦 Estado de Cuenta</h1>
  <p>Extrae movimientos de estados de cuenta bancarios (PDF o Excel) y genera reporte Excel con conciliación.</p>
</div>
""", unsafe_allow_html=True)

# ── Imports opcionales ────────────────────────────────────────────────────────

@st.cache_resource
def _get_pdfplumber():
    try:
        import pdfplumber
        return pdfplumber
    except ImportError:
        return None

@st.cache_resource
def _get_openpyxl():
    try:
        import openpyxl
        return openpyxl
    except ImportError:
        return None

def _importar_parser():
    """Importa ec_parser. Busca en el directorio del repositorio."""
    import sys, os
    dirs = [
        os.path.dirname(os.path.abspath(__file__)),    # pages/
        os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."),  # raíz repo
    ]
    for d in dirs:
        if d not in sys.path:
            sys.path.insert(0, d)
    import ec_parser
    return ec_parser

# ── BANCOS ────────────────────────────────────────────────────────────────────
BANCOS = [
    "Auto-detectar",
    "Banorte Débito", "Banorte Empresarial",
    "BBVA Débito", "BBVA Pyme", "BBVA Cash Management",
    "Banamex Débito", "Banamex Empresarial",
    "Santander", "HSBC", "Scotiabank",
    "Banregio", "Inbursa", "American Express", "Afirme",
]

# ── UI — configuración ────────────────────────────────────────────────────────
col_cfg, col_res = st.columns([1, 2], gap="large")

with col_cfg:
    st.markdown('<div class="ec-card">', unsafe_allow_html=True)
    st.markdown("##### ⚙️ Configuración")

    banco_sel = st.selectbox(
        "Banco / tipo de estado de cuenta",
        BANCOS,
        index=0,
        help="Si el banco no está en la lista o no estás seguro, elige 'Auto-detectar'.",
    )

    archivo = st.file_uploader(
        "Archivo de estado de cuenta",
        type=["pdf", "xlsx", "xls"],
        help="PDF o Excel del banco.",
    )

    saldo_ini = st.number_input(
        "Saldo inicial ($)",
        value=0.0,
        step=0.01,
        format="%.2f",
        help="Saldo al inicio del período (si el archivo lo trae, no es necesario).",
    )

    usar_saldo_esp = st.checkbox("Verificar saldo final esperado")
    saldo_esp = None
    if usar_saldo_esp:
        saldo_esp = st.number_input(
            "Saldo final esperado ($)",
            value=0.0,
            step=0.01,
            format="%.2f",
            help="Saldo según el banco al cierre del período.",
        )

    generar = st.button("🔍 Procesar estado de cuenta", type="primary", use_container_width=True)
    st.markdown('</div>', unsafe_allow_html=True)

    # Info bancos
    with st.expander("ℹ️ Bancos soportados"):
        st.markdown("""
**Parseo nativo completo:**
- Banorte (Débito / Empresarial)
- BBVA (Débito / Maestra PYME / Cash Management)
- Banamex / Citibanamex
- Santander · HSBC · Banregio
- Scotiabank · Inbursa · Afirme
- American Express

**Excel:** cualquier formato tabular con columnas de fecha, descripción, cargo y abono.
        """)

# ── Procesamiento ─────────────────────────────────────────────────────────────
with col_res:
    if generar:
        if archivo is None:
            st.warning("⚠️ Por favor sube un archivo antes de procesar.")
            st.stop()

        pdfplumber_mod = _get_pdfplumber()
        openpyxl_mod   = _get_openpyxl()

        if pdfplumber_mod is None or openpyxl_mod is None:
            st.error("Faltan dependencias: `pdfplumber` y `openpyxl` deben estar instalados.")
            st.stop()

        try:
            ec = _importar_parser()
        except Exception as e:
            st.error(f"No se pudo importar ec_parser: {e}")
            st.stop()

        ext  = os.path.splitext(archivo.name)[1].lower()
        nombre_base = os.path.splitext(archivo.name)[0]

        with st.spinner("Extrayendo movimientos…"):
            movs_raw = []
            try:
                with tempfile.NamedTemporaryFile(suffix=ext, delete=False) as tmp:
                    tmp.write(archivo.read()); tmp_path = tmp.name
                if ext == ".pdf":
                    if pdfplumber_mod is None:
                        st.error("pdfplumber no disponible para leer PDF."); st.stop()
                    movs_raw = ec.leer_pdf(tmp_path, pdfplumber_mod, banco_sel)
                else:
                    if openpyxl_mod is None:
                        st.error("openpyxl no disponible para leer Excel."); st.stop()
                    movs_raw = ec.leer_excel(tmp_path, openpyxl_mod)
            except Exception as e:
                st.error(f"Error al procesar el archivo: {e}")
                import traceback; st.text(traceback.format_exc())
                st.stop()
            finally:
                try: os.unlink(tmp_path)
                except Exception: pass

        if not movs_raw:
            st.warning("No se encontraron movimientos. Verifica que el banco seleccionado sea correcto o intenta con 'Auto-detectar'.")
            st.stop()

        # Calcular saldos
        filas, total_dep, total_ret = ec.calcular_saldos(movs_raw, saldo_ini)
        saldo_fin = filas[-1][4] if filas else saldo_ini
        n = len(filas)

        # ── Métricas ─────────────────────────────────────────────────────────
        mc1, mc2, mc3, mc4 = st.columns(4)
        with mc1:
            st.markdown(f'<div class="ec-stat"><div class="val">{n}</div><div class="lbl">Movimientos</div></div>', unsafe_allow_html=True)
        with mc2:
            st.markdown(f'<div class="ec-stat"><div class="val dep-val">${total_dep:,.2f}</div><div class="lbl">Total depósitos</div></div>', unsafe_allow_html=True)
        with mc3:
            st.markdown(f'<div class="ec-stat"><div class="val ret-val">${total_ret:,.2f}</div><div class="lbl">Total retiros</div></div>', unsafe_allow_html=True)
        with mc4:
            st.markdown(f'<div class="ec-stat"><div class="val sal-val">${saldo_fin:,.2f}</div><div class="lbl">Saldo final</div></div>', unsafe_allow_html=True)

        # ── Conciliación rápida ───────────────────────────────────────────────
        if saldo_esp is not None:
            diff = saldo_fin - saldo_esp
            if abs(diff) < 0.01:
                st.markdown(f'<div class="diff-ok">✅ Conciliación OK — diferencia ${diff:,.2f}</div>', unsafe_allow_html=True)
            else:
                st.markdown(f'<div class="diff-err">⚠️ Diferencia ${diff:,.2f} (calculado ${saldo_fin:,.2f} vs esperado ${saldo_esp:,.2f})</div>', unsafe_allow_html=True)
        else:
            st.markdown('<div class="diff-na">ℹ️ Ingresa "Saldo final esperado" para verificar conciliación.</div>', unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)

        # ── Vista previa ─────────────────────────────────────────────────────
        with st.expander("📋 Vista previa de movimientos", expanded=True):
            def _fmt_date(v):
                if isinstance(v, date): return v.strftime("%d/%m/%Y")
                return str(v)
            preview_data = []
            for f in filas:
                fecha, desc, dep, ret, sal = f
                preview_data.append({
                    "Fecha":       _fmt_date(fecha),
                    "Descripción": desc,
                    "Depósito":    f"${dep:,.2f}"  if dep  else "—",
                    "Retiro":      f"${ret:,.2f}"  if ret  else "—",
                    "Saldo":       f"${sal:,.2f}",
                })
            df_prev = pd.DataFrame(preview_data)
            st.dataframe(
                df_prev,
                use_container_width=True,
                hide_index=True,
                column_config={
                    "Fecha":       st.column_config.TextColumn("Fecha", width=90),
                    "Descripción": st.column_config.TextColumn("Descripción", width=350),
                    "Depósito":    st.column_config.TextColumn("Depósito", width=110),
                    "Retiro":      st.column_config.TextColumn("Retiro", width=110),
                    "Saldo":       st.column_config.TextColumn("Saldo", width=110),
                },
                height=min(40 + n * 35, 520),
            )

        # ── Generar Excel ─────────────────────────────────────────────────────
        st.markdown("---")
        with st.spinner("Generando Excel…"):
            try:
                excel_bytes = ec.generar_excel_bytes(
                    filas,
                    nombre_base,
                    saldo_ini=saldo_ini,
                    saldo_esp=saldo_esp if usar_saldo_esp else None,
                )
            except Exception as e:
                st.error(f"Error generando Excel: {e}")
                import traceback; st.text(traceback.format_exc())
                st.stop()

        nombre_out = f"{nombre_base}_estado_de_cuenta.xlsx"
        st.download_button(
            label="⬇️ Descargar Excel (Movimientos + Conciliación)",
            data=excel_bytes,
            file_name=nombre_out,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
        )
        st.success(f"✅ Listo — {n} movimientos exportados a **{nombre_out}**")

    else:
        # Estado inicial
        st.markdown("""
<div class="ec-card" style="text-align:center; padding: 3rem 2rem;">
  <div style="font-size:3rem; margin-bottom:1rem;">🏦</div>
  <h3 style="color:#1E3A8A; margin:0 0 .5rem;">Estado de Cuenta</h3>
  <p style="color:#64748B; max-width:360px; margin:0 auto;">
    Selecciona el banco, sube el PDF o Excel y presiona <strong>Procesar</strong>
    para extraer movimientos y generar el reporte de conciliación.
  </p>
</div>
""", unsafe_allow_html=True)
