"""
Página 14 — Póliza de Nómina
Genera la póliza de nómina a partir del Excel de Pagos Bancarios consolidado.
"""

import io
import streamlit as st
import pandas as pd

# ── page config ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Póliza de Nómina",
    page_icon="📋",
    layout="wide",
)

st.title("📋 Póliza de Nómina")
st.caption("Genera la póliza de nómina a partir del Excel de Pagos Bancarios consolidado.")

# ── imports con error amigable ────────────────────────────────────────────────
try:
    import conciliacion_nomina as cn
except ImportError as _e:
    st.error(f"No se pudo importar conciliacion_nomina: {_e}")
    st.stop()

# ── sidebar: instrucciones ────────────────────────────────────────────────────
with st.sidebar:
    st.header("Instrucciones")
    st.markdown(
        """
        1. **Pagos Bancarios** — Sube el Excel consolidado generado por el módulo de Pagos Bancarios.
        2. **Plantilla** *(opcional)* — Sube una plantilla `.xlsx` existente. Si no la subes, se genera una en blanco.
        3. Presiona **Generar Póliza**.
        4. Descarga el resultado con el botón **Descargar Excel**.
        """
    )

# ── controles ─────────────────────────────────────────────────────────────────
col1, col2 = st.columns(2)

with col1:
    pagos_file = st.file_uploader(
        "📊 Pagos Bancarios (Excel)",
        type=["xlsx", "xls"],
        key="pnom_pagos",
    )

with col2:
    plantilla_file = st.file_uploader(
        "📋 Plantilla (opcional)",
        type=["xlsx", "xls"],
        key="pnom_plantilla",
    )

generar = st.button("🔄 Generar Póliza de Nómina", type="primary",
                    disabled=(pagos_file is None))

# ── generación ────────────────────────────────────────────────────────────────
if generar and pagos_file is not None:
    import tempfile, os, shutil

    with tempfile.TemporaryDirectory() as tmp:
        # Guardar pagos
        pagos_path = os.path.join(tmp, "pagos.xlsx")
        with open(pagos_path, "wb") as f:
            f.write(pagos_file.getvalue())

        # Guardar plantilla si se subió
        plantilla_path = None
        if plantilla_file is not None:
            plantilla_path = os.path.join(tmp, "plantilla.xlsx")
            with open(plantilla_path, "wb") as f:
                f.write(plantilla_file.getvalue())

        out_path = os.path.join(tmp, "PolizaNomina.xlsx")

        with st.spinner("Generando póliza..."):
            try:
                n_hojas, n_emp = cn.generar_poliza_nomina(
                    pagos_path, out_path, plantilla_path
                )
                st.success(f"✅ Póliza generada — {n_hojas} hoja(s), {n_emp} empleado(s)")

                # Leer resultado para descarga
                with open(out_path, "rb") as f:
                    xlsx_bytes = f.read()

                st.download_button(
                    label="⬇️ Descargar Excel",
                    data=xlsx_bytes,
                    file_name="PolizaNomina.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

                # Vista previa
                st.subheader("Vista previa — Póliza IA")
                try:
                    df_prev = pd.read_excel(
                        io.BytesIO(xlsx_bytes),
                        sheet_name=0,
                        header=2,       # fila 3 = nombres
                        nrows=30,
                        dtype=str,
                    ).fillna("")
                    st.dataframe(df_prev, use_container_width=True, height=350)
                except Exception as _pe:
                    st.warning(f"Vista previa no disponible: {_pe}")

                st.subheader("Vista previa — Conciliación")
                try:
                    df_conc = pd.read_excel(
                        io.BytesIO(xlsx_bytes),
                        sheet_name="Conciliación",
                        header=0,
                        nrows=40,
                        dtype=str,
                    ).fillna("")
                    st.dataframe(df_conc, use_container_width=True, height=250)
                except Exception as _ce:
                    st.info(f"Hoja Conciliación: {_ce}")

            except Exception as exc:
                import traceback
                st.error(f"Error al generar: {exc}")
                st.code(traceback.format_exc())
