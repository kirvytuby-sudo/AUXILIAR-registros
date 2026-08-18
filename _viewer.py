"""
_viewer.py — Visor Excel editable con cinta de herramientas al estilo Excel.

Uso:
    import _viewer
    edited_bytes = _viewer.show(file_bytes, filename="reporte.xlsx", key="pagos")
"""

from __future__ import annotations
import io
import streamlit as st
import pandas as pd

# ── CSS de la cinta ────────────────────────────────────────────────────────────
_CSS = """
<style>
/* ── Tabs de la cinta ─────────────── */
.vwr-tabs {
    display: flex;
    background: #4A6FA5;
    border-radius: 8px 8px 0 0;
    padding: 0 8px;
    gap: 2px;
}
.vwr-tab {
    color: #DBEAFE;
    font-size: .75rem;
    font-weight: 700;
    padding: 6px 16px;
    cursor: default;
    letter-spacing: .3px;
    border-radius: 6px 6px 0 0;
    transition: background .15s;
}
.vwr-tab.active {
    background: #FFFFFF;
    color: #2563EB;
}
/* ── Barra de grupos de herramientas ─ */
.vwr-bar {
    background: #F8FAFF;
    border: 1.5px solid #BFDBFE;
    border-top: none;
    padding: 6px 8px 3px 8px;
    display: flex;
    align-items: flex-end;
    gap: 0;
}
.vwr-grp-lbl {
    font-size: .60rem;
    color: #4A6FA5;
    font-weight: 800;
    text-transform: uppercase;
    letter-spacing: .5px;
    text-align: center;
    padding: 3px 4px 0;
    border-top: 2px solid #BFDBFE;
    margin-top: 5px;
    width: 100%;
}
.vwr-sep {
    width: 1px;
    background: #BFDBFE;
    align-self: stretch;
    margin: 3px 8px;
}
/* ── Botones Streamlit en la cinta ── */
.vwr-bar [data-testid="stBaseButton-secondary"] > button,
.vwr-bar [data-testid="stButton"] > button {
    background: #EFF6FF !important;
    border: 1.5px solid #93C5FD !important;
    color: #1D4ED8 !important;
    font-weight: 700 !important;
    border-radius: 6px !important;
    box-shadow: 0 1px 3px rgba(37,99,235,.15) !important;
    transition: all .15s !important;
}
.vwr-bar [data-testid="stBaseButton-secondary"] > button:hover,
.vwr-bar [data-testid="stButton"] > button:hover {
    background: #DBEAFE !important;
    border-color: #3B82F6 !important;
    box-shadow: 0 2px 6px rgba(37,99,235,.25) !important;
}
/* ── Encabezado de hoja ─────────────── */
.vwr-hdr {
    background: linear-gradient(135deg, #4A6FA5 0%, #6B93C9 100%);
    padding: 6px 14px;
    display: flex;
    align-items: center;
    gap: 10px;
}
.vwr-hdr-title { color: #FFF; font-weight: 700; font-size: .90rem; letter-spacing: .2px; }
.vwr-badge {
    background: rgba(255,255,255,.25);
    color: #fff;
    border-radius: 4px;
    padding: 2px 9px;
    font-size: .73rem;
    font-weight: 600;
}
.vwr-hint { color: #DBEAFE; font-size: .72rem; margin-left: auto; }
/* ── Barra de estado ────────────────── */
.vwr-status {
    background: #4A6FA5;
    border-radius: 0 0 6px 6px;
    padding: 4px 14px;
    display: flex;
    gap: 18px;
    font-size: .70rem;
    color: #DBEAFE;
    font-weight: 500;
}
/* ── Celda activa ───────────────────── */
[data-testid="stDataEditor"] td:focus-within {
    outline: 2px solid #3B82F6 !important;
    background: #EFF6FF !important;
}
[data-testid="stDataEditor"] { border: 1.5px solid #BFDBFE !important; border-top: none !important; }
</style>
"""


def show(
    file_bytes: bytes,
    filename: str = "reporte.xlsx",
    key: str = "vwr",
    height: int = 430,
    show_download: bool = True,
) -> bytes:
    """
    Muestra un visor Excel editable con cinta de herramientas.
    Devuelve bytes del archivo con los cambios aplicados.
    """
    if not file_bytes:
        return file_bytes

    st.markdown(_CSS, unsafe_allow_html=True)

    # ── Leer hojas disponibles ─────────────────────────────────────────────
    try:
        xf = pd.ExcelFile(io.BytesIO(file_bytes))
        sheet_names = xf.sheet_names
    except Exception as e:
        st.error(f"No se pudo abrir el archivo: {e}")
        return file_bytes

    # ── Estado de sesión para la cinta ──────────────────────────────────────
    def _ss_init(k, default):
        if k not in st.session_state:
            st.session_state[k] = default

    K_sheet   = f"{key}_sheet"
    K_sort    = f"{key}_sort_col"
    K_asc     = f"{key}_sort_asc"
    K_flt     = f"{key}_filter"
    K_fmt     = f"{key}_num_fmt"
    K_editor  = f"{key}_editor"

    _ss_init(K_sort, "")
    _ss_init(K_asc,  True)
    _ss_init(K_flt,  "")
    _ss_init(K_fmt,  "auto")

    # ── Selector de hoja ──────────────────────────────────────────────────
    if len(sheet_names) > 1:
        active_sheet = st.selectbox(
            "Hoja:", sheet_names, key=K_sheet, label_visibility="collapsed",
        )
    else:
        active_sheet = sheet_names[0]

    # ── Leer hoja activa ──────────────────────────────────────────────────
    try:
        df_orig = pd.read_excel(io.BytesIO(file_bytes), sheet_name=active_sheet)
    except Exception as e:
        st.error(f"Error al leer '{active_sheet}': {e}")
        return file_bytes

    cols_list = list(df_orig.columns)
    n_rows_orig, n_cols = df_orig.shape

    # ════════════════════════════════════════════════════════════════════════
    #  CINTA DE HERRAMIENTAS
    # ════════════════════════════════════════════════════════════════════════

    # ── Tabs (visual) ──────────────────────────────────────────────────────
    st.markdown(
        '<div class="vwr-tabs">'
        '<span class="vwr-tab active">Inicio</span>'
        '<span class="vwr-tab">Insertar</span>'
        '<span class="vwr-tab">Datos</span>'
        '<span class="vwr-tab">Vista</span>'
        '</div>',
        unsafe_allow_html=True,
    )

    # ── Controles funcionales ─────────────────────────────────────────────
    c_fmt, c_sep1, c_sort, c_sep2, c_flt, c_sep3, c_rows, c_sep4, c_exp = st.columns(
        [2.2, 0.05, 2.8, 0.05, 3.2, 0.05, 1.8, 0.05, 2.2]
    )

    # ── Grupo: Número ─────────────────────────────────────────────────────
    with c_fmt:
        f1, f2, f3, f4 = st.columns([1, 1, 1, 1])
        with f1:
            if st.button("**$**", key=f"{key}_cur", help="Formato moneda  $#,##0.00",
                         use_container_width=True):
                st.session_state[K_fmt] = "currency"
                st.rerun()
        with f2:
            if st.button("**%**", key=f"{key}_pct", help="Formato porcentaje",
                         use_container_width=True):
                st.session_state[K_fmt] = "percent"
                st.rerun()
        with f3:
            if st.button("**,**", key=f"{key}_num", help="Número con separadores",
                         use_container_width=True):
                st.session_state[K_fmt] = "number"
                st.rerun()
        with f4:
            if st.button("**A**", key=f"{key}_auto", help="Formato automático",
                         use_container_width=True):
                st.session_state[K_fmt] = "auto"
                st.rerun()
        st.markdown('<div class="vwr-grp-lbl">Número</div>', unsafe_allow_html=True)

    st.markdown('<div class="vwr-sep"></div>', unsafe_allow_html=True)

    # ── Grupo: Ordenar ────────────────────────────────────────────────────
    with c_sort:
        s1, s2, s3 = st.columns([3, 1, 1])
        with s1:
            st.selectbox(
                "Col", ["(sin orden)"] + cols_list,
                key=K_sort, label_visibility="collapsed",
            )
        with s2:
            if st.button("↑ A→Z", key=f"{key}_asc", help="Ordenar ascendente",
                         use_container_width=True):
                st.session_state[K_asc] = True
                st.rerun()
        with s3:
            if st.button("↓ Z→A", key=f"{key}_desc", help="Ordenar descendente",
                         use_container_width=True):
                st.session_state[K_asc] = False
                st.rerun()
        st.markdown('<div class="vwr-grp-lbl">Ordenar</div>', unsafe_allow_html=True)

    # ── Grupo: Buscar/Filtrar ─────────────────────────────────────────────
    with c_flt:
        b1, b2 = st.columns([4, 1])
        with b1:
            st.text_input(
                "Buscar", placeholder="🔍  Buscar en tabla…",
                key=K_flt, label_visibility="collapsed",
            )
        with b2:
            if st.button("✖", key=f"{key}_cls", help="Limpiar filtro",
                         use_container_width=True):
                st.session_state[K_flt] = ""
                st.rerun()
        st.markdown('<div class="vwr-grp-lbl">Buscar & Filtrar</div>', unsafe_allow_html=True)

    # ── Grupo: Info filas ─────────────────────────────────────────────────
    with c_rows:
        fmt_label = {
            "currency": "$ Moneda", "percent": "% Porcent.",
            "number": ", Número", "auto": "Auto",
        }.get(st.session_state[K_fmt], "Auto")
        st.markdown(
            f"<div style='font-size:.78rem;color:#374151;padding-top:4px;'>"
            f"<b>{n_rows_orig}</b> filas · <b>{n_cols}</b> cols<br>"
            f"<span style='color:#6B7280'>Núm: {fmt_label}</span></div>",
            unsafe_allow_html=True,
        )
        st.markdown('<div class="vwr-grp-lbl">Datos</div>', unsafe_allow_html=True)

    # ── Grupo: Exportar ───────────────────────────────────────────────────
    with c_exp:
        # (Los botones de descarga van abajo del editor para acceso fácil)
        if st.button("🔄 Recargar original", key=f"{key}_reload", use_container_width=True,
                     help="Descarta cambios y vuelve al archivo original"):
            # Borrar estado del editor
            if K_editor in st.session_state:
                del st.session_state[K_editor]
            st.session_state[K_sort] = ""
            st.session_state[K_flt]  = ""
            st.session_state[K_fmt]  = "auto"
            st.rerun()
        st.markdown('<div class="vwr-grp-lbl">Edición</div>', unsafe_allow_html=True)

    # ════════════════════════════════════════════════════════════════════════
    #  Aplicar sort y filtro al dataframe
    # ════════════════════════════════════════════════════════════════════════
    df = df_orig.copy()

    sort_col_val = st.session_state.get(K_sort, "")
    if sort_col_val and sort_col_val in df.columns:
        try:
            df = df.sort_values(sort_col_val, ascending=st.session_state.get(K_asc, True))
        except Exception:
            pass

    filter_val = st.session_state.get(K_flt, "")
    if filter_val:
        mask = df.astype(str).apply(
            lambda col: col.str.contains(filter_val, case=False, na=False)
        ).any(axis=1)
        df = df[mask]

    n_visible = len(df)

    # ── Column config: formato de números ─────────────────────────────────
    num_fmt = st.session_state.get(K_fmt, "auto")
    col_config = {}
    for col in df.columns:
        if pd.api.types.is_numeric_dtype(df[col]):
            if num_fmt == "currency":
                col_config[col] = st.column_config.NumberColumn(col, format="$%,.2f")
            elif num_fmt == "percent":
                col_config[col] = st.column_config.NumberColumn(col, format="%.2f%%")
            elif num_fmt == "number":
                col_config[col] = st.column_config.NumberColumn(col, format="%,.2f")

    # ════════════════════════════════════════════════════════════════════════
    #  Encabezado de hoja + Editor
    # ════════════════════════════════════════════════════════════════════════
    flt_badge = f'<span class="vwr-badge">Filtro: {n_visible} de {n_rows_orig}</span>' \
                if filter_val else ""

    st.markdown(
        f'<div class="vwr-hdr">'
        f'<span class="vwr-hdr-title">📊 {active_sheet}</span>'
        f'{flt_badge}'
        f'<span class="vwr-badge">{n_visible} filas · {n_cols} cols</span>'
        f'<span class="vwr-hint">✏️ Clic en celda para editar — arrastra columnas para reordenar</span>'
        f'</div>',
        unsafe_allow_html=True,
    )

    edited_df = st.data_editor(
        df.reset_index(drop=True),
        use_container_width=True,
        num_rows="dynamic",
        height=height,
        key=K_editor,
        column_config=col_config if col_config else None,
    )

    # ── Barra de estado ────────────────────────────────────────────────────
    numeric_cols = [c for c in edited_df.columns if pd.api.types.is_numeric_dtype(edited_df[c])]
    status_parts = [f"Filas: {len(edited_df)}  ·  Cols: {n_cols}"]
    for nc in numeric_cols[:4]:
        s = edited_df[nc].sum()
        status_parts.append(f"Σ {nc}: {s:,.2f}")
    st.markdown(
        f'<div class="vwr-status">{"  ·  ".join(status_parts)}</div>',
        unsafe_allow_html=True,
    )

    # ════════════════════════════════════════════════════════════════════════
    #  Reconstruir Excel (todas las hojas)
    # ════════════════════════════════════════════════════════════════════════
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

    # ── Botones de descarga ────────────────────────────────────────────────
    if show_download:
        dl1, dl2 = st.columns(2)
        with dl1:
            st.download_button(
                "💾 Descargar Excel (con cambios)",
                data=edited_bytes,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                key=f"{key}_dl_xl",
                type="secondary",
            )
        with dl2:
            csv_buf = io.StringIO()
            edited_df.to_csv(csv_buf, index=False, encoding="utf-8-sig")
            st.download_button(
                "📄 Descargar CSV",
                data=csv_buf.getvalue().encode("utf-8-sig"),
                file_name=filename.replace(".xlsx", ".csv").replace(".xls", ".csv"),
                mime="text/csv",
                use_container_width=True,
                key=f"{key}_dl_csv",
                type="secondary",
            )

    return edited_bytes
