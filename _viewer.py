"""
_viewer.py — Visor Excel editable compartido para todas las páginas del Auxiliar.

Uso básico:
    import _viewer
    edited = _viewer.show(file_bytes, filename="reporte.xlsx", key="pagos")

    # Si quieres ocultar el botón de descarga propio del viewer (porque la página ya tiene uno):
    edited = _viewer.show(..., show_download=False)

La función devuelve los bytes del archivo con los cambios que el usuario haya hecho.
Si el usuario no hizo ningún cambio, devuelve los mismos bytes originales.
"""

from __future__ import annotations
import io
import streamlit as st
import pandas as pd


# ── CSS para la barra superior del visor ───────────────────────────────────────
_VIEWER_CSS = """
<style>
/* Barra de encabezado del visor */
.vwr-header {
    background: linear-gradient(135deg,#1E3A8A 0%,#2563EB 100%);
    border-radius: 8px 8px 0 0;
    padding: 8px 16px;
    display: flex; align-items: center; gap: 10px;
    margin-bottom: 0;
}
.vwr-header span { color:#FBCFE8; font-weight:700; font-size:.92rem; }
.vwr-badge {
    background:rgba(255,255,255,.18); color:#fff;
    border-radius:4px; padding:2px 8px; font-size:.78rem;
}
/* Resaltar celda activa del data_editor */
[data-testid="stDataEditor"] td:focus-within {
    outline: 2px solid #2563EB !important;
    background: #EFF6FF !important;
}
</style>
"""


def show(
    file_bytes: bytes,
    filename: str = "reporte.xlsx",
    key: str = "vwr",
    height: int = 440,
    show_download: bool = True,
) -> bytes:
    """
    Muestra un visor de Excel editable tipo spreadsheet.

    Parámetros
    ----------
    file_bytes    : bytes del archivo .xlsx a mostrar
    filename      : nombre sugerido al descargar
    key           : prefijo único de Streamlit (evita colisiones entre páginas)
    height        : altura en píxeles del editor
    show_download : si True muestra botón "Descargar con cambios"

    Devuelve
    --------
    bytes del Excel reconstruido con los cambios del usuario.
    """
    if not file_bytes:
        return file_bytes

    st.markdown(_VIEWER_CSS, unsafe_allow_html=True)

    # ── Leer hojas disponibles ─────────────────────────────────────────────────
    try:
        xf = pd.ExcelFile(io.BytesIO(file_bytes))
        sheet_names = xf.sheet_names
    except Exception as e:
        st.error(f"⚠️ No se pudo abrir el archivo para edición: {e}")
        return file_bytes

    # ── Selector de hoja ───────────────────────────────────────────────────────
    if len(sheet_names) > 1:
        cols_top = st.columns([3, 1])
        with cols_top[0]:
            active_sheet = st.selectbox(
                "📋 Hoja:",
                sheet_names,
                key=f"{key}_sheet",
                label_visibility="collapsed",
            )
        with cols_top[1]:
            st.markdown(
                f'<div style="padding-top:6px;font-size:.8rem;color:#64748B;">'
                f'{len(sheet_names)} hojas</div>',
                unsafe_allow_html=True,
            )
    else:
        active_sheet = sheet_names[0]

    # ── Leer hoja activa ───────────────────────────────────────────────────────
    try:
        df_orig = pd.read_excel(io.BytesIO(file_bytes), sheet_name=active_sheet)
    except Exception as e:
        st.error(f"Error al leer hoja '{active_sheet}': {e}")
        return file_bytes

    n_rows, n_cols = df_orig.shape

    # ── Cabecera del visor ─────────────────────────────────────────────────────
    st.markdown(
        f'<div class="vwr-header">'
        f'<span>📊 {active_sheet}</span>'
        f'<span class="vwr-badge">{n_rows} filas</span>'
        f'<span class="vwr-badge">{n_cols} columnas</span>'
        f'<span style="color:#93C5FD;font-size:.78rem;margin-left:auto;">'
        f'✏️ Haz clic en una celda para editar · puedes añadir o eliminar filas</span>'
        f'</div>',
        unsafe_allow_html=True,
    )

    # ── Editor ────────────────────────────────────────────────────────────────
    edited_df = st.data_editor(
        df_orig,
        use_container_width=True,
        num_rows="dynamic",
        height=height,
        key=f"{key}_editor",
    )

    # ── Reconstruir Excel con todas las hojas ──────────────────────────────────
    buf = io.BytesIO()
    try:
        all_data: dict[str, pd.DataFrame] = {}
        for sh in sheet_names:
            if sh == active_sheet:
                all_data[sh] = edited_df
            else:
                try:
                    all_data[sh] = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sh)
                except Exception:
                    all_data[sh] = pd.DataFrame()

        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            for sh, sdf in all_data.items():
                sdf.to_excel(writer, sheet_name=sh, index=False)

        edited_bytes = buf.getvalue()
    except Exception as e:
        st.warning(f"No se pudo reconstruir el Excel: {e}")
        edited_bytes = file_bytes

    # ── Botón de descarga ──────────────────────────────────────────────────────
    if show_download:
        st.download_button(
            "💾 Descargar con cambios",
            data=edited_bytes,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            key=f"{key}_dl",
            type="secondary",
        )

    return edited_bytes
