"""
pages/9_Estado_de_Cuenta.py — Módulo Estado de Cuenta
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

import _theme
_theme.aplicar_header("🏦 Estado de Cuenta", "Análisis y conciliación de estados de cuenta bancarios")
# ── Imports opcionales ────────────────────────────────────────────────────────

@st.cache_resource
@st.cache_data
def _get_pdfplumber():
    try:
        import pdfplumber
        return pdfplumber
    except ImportError:
        return None

@st.cache_resource
@st.cache_data
def _get_openpyxl():
    try:
        import openpyxl
        return openpyxl
    except ImportError:
        return None

@st.cache_data
def _importar_parser():
    """Importa ec_parser. Busca en el directorio del repositorio."""
    import sys, os, importlib
    dirs = [
        os.path.dirname(os.path.abspath(__file__)),
        os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."),
    ]
    for d in dirs:
        if d not in sys.path:
            sys.path.insert(0, d)
    import ec_parser
    importlib.reload(ec_parser)  # fuerza recarga para evitar caché de sys.modules
    return ec_parser


def _obtener_api_key():
    """Obtiene la API key de Anthropic: session_state → secrets → env."""
    if st.session_state.get("_ec_ia_key"):
        return st.session_state["_ec_ia_key"]
    try:
        k = st.secrets.get("ANTHROPIC_API_KEY", "")
        if k: return k
    except Exception:
        pass
    import os
    return os.environ.get("ANTHROPIC_API_KEY", "")


def _parsear_con_ia(texto_pdf: str, banco_hint: str, api_key: str) -> list:
    """Llama a Claude Haiku para extraer movimientos cuando el parser falla."""
    import re, json
    from datetime import datetime
    try:
        import anthropic
    except ImportError:
        st.error("Paquete `anthropic` no instalado en el servidor. Agrégalo a requirements.txt.")
        return []

    texto_recortado = texto_pdf[:8000] if len(texto_pdf) > 8000 else texto_pdf
    prompt = f"""Eres un experto en estados de cuenta bancarios mexicanos.
Banco: {banco_hint or 'desconocido'}

Extrae TODOS los movimientos del siguiente texto. Para cada uno devuelve:
- fecha: DD/MM/YYYY
- descripcion: descripción del movimiento (sin montos)
- deposito: monto depositado/abonado (0 si no aplica)
- retiro: monto retirado/cargado (0 si no aplica)

Responde ÚNICAMENTE con JSON válido:
{{"movimientos": [{{"fecha":"DD/MM/YYYY","descripcion":"...","deposito":0.0,"retiro":0.0}}]}}

Texto del estado de cuenta:
{texto_recortado}"""

    try:
        client = anthropic.Anthropic(api_key=api_key)
        resp = client.messages.create(
            model="claude-haiku-4-5-20251001",
            max_tokens=4096,
            messages=[{"role": "user", "content": prompt}],
        )
        raw = resp.content[0].text.strip()
        m = re.search(r'\{.*\}', raw, re.DOTALL)
        if not m:
            return []
        data = json.loads(m.group())
        result = []
        for mov in data.get("movimientos", []):
            try:
                fecha_str = mov.get("fecha", "")
                fecha = None
                for fmt in ("%d/%m/%Y", "%d/%m/%y", "%Y-%m-%d", "%d-%m-%Y"):
                    try:
                        fecha = datetime.strptime(fecha_str, fmt).date()
                        break
                    except ValueError:
                        continue
                if fecha is None:
                    continue
                desc = str(mov.get("descripcion", "—")).strip() or "—"
                dep  = float(mov.get("deposito", 0) or 0)
                ret  = float(mov.get("retiro",   0) or 0)
                result.append((fecha, desc, dep, ret))
            except Exception:
                continue
        return result
    except Exception as e:
        st.error(f"Error llamando a la API de Claude: {e}")
        return []

# ── BANCOS ────────────────────────────────────────────────────────────────────
BANCOS = [
    "Auto-detectar",
    "Banorte Débito", "Banorte Empresarial",
    "BBVA Débito", "BBVA Pyme", "BBVA Cash Management", "BBVA TDC", "BBVA Libretón",
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

    # ── Opción C: IA fallback ────────────────────────────────────────
    with st.expander("🤖 Fallback con IA (cuando el parser falla)"):
        st.markdown("""
**¿Qué hace?** Si el parser no encuentra movimientos, la app llama automáticamente
a la API de Claude (modelo Haiku, rápido y económico) para extraerlos.

**Para Streamlit Cloud:** agrega `ANTHROPIC_API_KEY = "sk-ant-..."` en
*Settings → Secrets* de tu app.

**Para uso temporal:** escribe la key aquí (solo dura esta sesión).
        """)
        _key_placeholder = ""
        try:
            _key_placeholder = "••••" + st.secrets["ANTHROPIC_API_KEY"][-4:] \
                               if st.secrets.get("ANTHROPIC_API_KEY") else ""
        except Exception:
            pass
        _ia_key_input = st.text_input(
            "API Key de Anthropic (opcional)",
            value=_key_placeholder,
            type="password",
            placeholder="sk-ant-...",
            help="Si ya configuraste ANTHROPIC_API_KEY en Secrets, déjalo vacío.",
        )
        if _ia_key_input and not _ia_key_input.startswith("••••"):
            st.session_state["_ec_ia_key"] = _ia_key_input

    with st.expander("ℹ️ Bancos soportados"):
        st.markdown("""
**Parseo nativo completo:**
- Banorte (Débito / Empresarial)
- BBVA (Débito / Maestra PYME / Cash Management)
- **BBVA TDC** (T Negoc / LCDigital — Tarjeta de Crédito)
- Banamex / Citibanamex
- Santander · HSBC · Banregio
- Scotiabank · Inbursa · Afirme
- American Express

**Excel:** cualquier formato tabular con columnas de fecha, descripción, cargo y abono.

> **Nota BBVA TDC:** Para PDFs descargados directamente del portal BBVA (texto seleccionable) el parseo es instantáneo. Para PDFs generados con "Imprimir → Guardar como PDF" (imagen) se usa OCR automático — requiere `tesseract-ocr` instalado en el servidor. En la columna *Depósito* aparecen los **abonos (pagos)** y en *Retiro* los **cargos (compras)**.
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

        _texto_pdf_para_ia = ""   # guardamos para fallback IA
        with st.spinner("Extrayendo movimientos…"):
            movs_raw = []
            try:
                with tempfile.NamedTemporaryFile(suffix=ext, delete=False) as tmp:
                    tmp.write(archivo.read()); tmp_path = tmp.name
                if ext == ".pdf":
                    # Guardar texto para posible fallback IA
                    try:
                        with pdfplumber_mod.open(tmp_path) as _pdf:
                            _texto_pdf_para_ia = "\n".join(
                                (p.extract_text() or "") for p in _pdf.pages
                            )
                    except Exception:
                        pass
                    movs_raw = ec.leer_pdf(tmp_path, pdfplumber_mod, banco_sel)
                else:
                    movs_raw = ec.leer_excel(tmp_path, openpyxl_mod)
            except Exception as e:
                st.error(f"Error al procesar el archivo: {e}")
                import traceback; st.text(traceback.format_exc())
                st.stop()
            finally:
                try: os.unlink(tmp_path)
                except Exception: pass

        # ── Opción C: fallback con IA si parser no encontró nada ──────
        if not movs_raw and ext == ".pdf":
            _api_key = _obtener_api_key()
            if _api_key:
                with st.spinner("🤖 Parser convencional sin resultados — intentando con IA (Claude Haiku)…"):
                    movs_raw = _parsear_con_ia(_texto_pdf_para_ia, banco_sel, _api_key)
                if movs_raw:
                    st.info(f"🤖 IA extrajo {len(movs_raw)} movimientos (el parser convencional no encontró ninguno).")

        if not movs_raw:
            st.warning("No se encontraron movimientos. Verifica que el banco seleccionado sea correcto o intenta con 'Auto-detectar'.")
            if ext == ".pdf" and not _obtener_api_key():
                st.info("💡 Tip: configura una API Key de Anthropic en el panel izquierdo para que la IA lo intente automáticamente.")
            st.stop()

        filas, total_dep, total_ret = ec.calcular_saldos(movs_raw, saldo_ini)
        saldo_fin = filas[-1][4] if filas else saldo_ini
        n = len(filas)

        mc1, mc2, mc3, mc4 = st.columns(4)
        with mc1:
            st.markdown(f'<div class="ec-stat"><div class="val">{n}</div><div class="lbl">Movimientos</div></div>', unsafe_allow_html=True)
        with mc2:
            st.markdown(f'<div class="ec-stat"><div class="val dep-val">${total_dep:,.2f}</div><div class="lbl">Total depósitos</div></div>', unsafe_allow_html=True)
        with mc3:
            st.markdown(f'<div class="ec-stat"><div class="val ret-val">${total_ret:,.2f}</div><div class="lbl">Total retiros</div></div>', unsafe_allow_html=True)
        with mc4:
            st.markdown(f'<div class="ec-stat"><div class="val sal-val">${saldo_fin:,.2f}</div><div class="lbl">Saldo final</div></div>', unsafe_allow_html=True)

        if saldo_esp is not None:
            diff = saldo_fin - saldo_esp
            if abs(diff) < 0.01:
                st.markdown(f'<div class="diff-ok">✅ Conciliación OK — diferencia ${diff:,.2f}</div>', unsafe_allow_html=True)
            else:
                st.markdown(f'<div class="diff-err">⚠️ Diferencia ${diff:,.2f} (calculado ${saldo_fin:,.2f} vs esperado ${saldo_esp:,.2f})</div>', unsafe_allow_html=True)
        else:
            st.markdown('<div class="diff-na">ℹ️ Ingresa "Saldo final esperado" para verificar conciliación.</div>', unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)

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
