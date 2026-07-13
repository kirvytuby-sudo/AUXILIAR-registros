"""
pages/8_Conciliacion_Banco_Auxiliar.py — Módulo Conciliación Banco vs Auxiliar
Carga un Excel del banco y un auxiliar contable, concilia depósitos↔cargos
y retiros↔abonos con tolerancia ±$0.05 y ±2 días.
"""
import io
import os
import tempfile
import traceback
from datetime import date, datetime, timedelta

import streamlit as st

st.set_page_config(
    page_title="Conciliación Banco vs Auxiliar | Auxiliar de Registros",
    page_icon="🔀",
    layout="wide",
)

st.markdown("""
<style>
[data-testid="stAppViewContainer"] { background: #F8FAFC; }
.block-container { padding-top: 1.5rem; padding-bottom: 2rem; }
.cba-header {
    background: linear-gradient(135deg, #1E3A8A 0%, #2563EB 100%);
    color: white; border-radius: 10px;
    padding: 1.2rem 1.8rem; margin-bottom: 1.5rem;
}
.cba-header h1 { margin: 0; font-size: 1.6rem; font-weight: 700; }
.cba-header p  { margin: .3rem 0 0; opacity: .8; font-size: .9rem; }
.cba-card {
    background: white; border-radius: 8px; padding: 1.2rem 1.5rem;
    box-shadow: 0 1px 6px rgba(0,0,0,.07); margin-bottom: 1rem;
}
.cba-stat {
    background: #EEF2FF; border-radius: 8px; padding: .9rem 1.2rem;
    text-align: center; margin-bottom: .5rem;
}
.cba-stat .val { font-size: 1.4rem; font-weight: 700; color: #1E3A8A; }
.cba-stat .lbl { font-size: .75rem; color: #64748B; margin-top: .1rem; }
.leg { display:inline-block; width:14px; height:14px; border-radius:3px;
       margin-right:5px; vertical-align:middle; }
#MainMenu { visibility:hidden; } footer { visibility:hidden; }
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="cba-header">
  <h1>🔀 Conciliación Banco vs Auxiliar</h1>
  <p>Excel del banco + Auxiliar contable → conciliación bilateral con reporte Excel</p>
</div>
""", unsafe_allow_html=True)

# ── Utilidades ────────────────────────────────────────────────────────────────

def _to_date(v):
    """Convierte un valor a date. Acepta date, datetime, string, número Excel."""
    if v is None:
        return None
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    if isinstance(v, (int, float)):
        try:
            from openpyxl.utils.datetime import from_excel
            return from_excel(v).date()
        except Exception:
            return None
    s = str(v).strip()
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%m/%d/%Y"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None


def _to_float(v):
    """Convierte valor a float, devuelve 0.0 si no se puede."""
    if v is None:
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(str(v).replace(",", "").replace("$", "").strip())
    except Exception:
        return 0.0


def _detect_header(rows, keywords):
    """Devuelve (header_row_index, {col_name: index}) buscando keywords en todas las filas."""
    for i, row in enumerate(rows):
        vals = [str(v).strip().lower() if v is not None else "" for v in row]
        if all(kw in vals for kw in keywords):
            # construir mapa nombre→índice
            mapping = {v: j for j, v in enumerate(vals) if v}
            return i, mapping
    return None, {}


def _col(mapping, *candidates):
    """Devuelve el primer índice que coincida con alguna de las palabras clave."""
    for cand in candidates:
        for k, v in mapping.items():
            if cand in k:
                return v
    return None


def _read_banco(wb):
    """
    Lee el archivo del banco (Excel).
    Detecta encabezado dinámicamente buscando 'fecha'.
    Devuelve lista de dicts: fecha, descripcion, deposito, retiro, saldo.
    """
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))

    hdr_idx, mapping = _detect_header(rows, ["fecha"])
    if hdr_idx is None:
        raise ValueError("No se encontró encabezado con columna 'Fecha' en el archivo del banco.")

    col_f   = _col(mapping, "fecha")
    col_d   = _col(mapping, "desc", "concepto", "referencia", "movimiento")
    col_dep = _col(mapping, "dep", "abono", "crédito", "credito", "entrada")
    col_ret = _col(mapping, "ret", "cargo", "débito", "debito", "salida")
    col_sal = _col(mapping, "saldo", "sal")

    if col_f is None:
        raise ValueError("No se detectó la columna 'Fecha' en el banco.")

    movs = []
    for row in rows[hdr_idx + 1:]:
        if not row or all(v is None for v in row):
            continue
        fecha = _to_date(row[col_f] if col_f < len(row) else None)
        if fecha is None:
            continue
        desc  = str(row[col_d]   if col_d   is not None and col_d   < len(row) else "").strip()
        dep   = _to_float(row[col_dep] if col_dep is not None and col_dep < len(row) else 0)
        ret   = _to_float(row[col_ret] if col_ret is not None and col_ret < len(row) else 0)
        sal   = _to_float(row[col_sal] if col_sal is not None and col_sal < len(row) else 0)
        movs.append({"fecha": fecha, "desc": desc, "dep": dep, "ret": ret, "sal": sal})

    return movs


def _read_auxiliar(wb):
    """
    Lee el auxiliar contable (Excel).
    Detecta encabezado dinámicamente buscando 'cargo' Y 'fecha'.
    Devuelve (aux_cargo, aux_abono) — listas de dicts.
    """
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))

    hdr_idx, mapping = _detect_header(rows, ["cargo", "fecha"])
    if hdr_idx is None:
        raise ValueError("No se encontró encabezado con columnas 'Fecha' y 'Cargo' en el auxiliar.")

    col_f     = _col(mapping, "fecha")

    # Póliza: buscar número/folio primero para no confundir con "Tipo de Póliza"
    col_pol   = _col(mapping, "núm. póliza", "num. póliza", "número póliza",
                     "numero poliza", "num poliza", "folio póliza", "folio")
    if col_pol is None:
        # Fallback genérico — evitar columnas que empiecen con "tipo"
        for cand in ("póliza", "poliza", "pol"):
            for k, v in mapping.items():
                if cand in k and not k.startswith("tipo"):
                    col_pol = v
                    break
            if col_pol is not None:
                break

    # Desc. Póliza (fuente principal para Concepto Aux)
    col_dp    = _col(mapping, "desc. póliza", "desc poliza", "descripción póliza",
                     "descripcion poliza", "descripción", "descripcion", "desc")

    # Detalle / Concepto (secundario)
    col_det   = _col(mapping, "detalle", "concepto", "cliente", "referencia")
    if col_det is not None and col_det == col_dp:
        col_det = None

    col_cargo = _col(mapping, "cargo")
    col_abono = _col(mapping, "abono")

    def _s(row, col):
        """Extrae string de celda; devuelve '' para None/nan/None-literal."""
        if col is None or col >= len(row):
            return ""
        v = row[col]
        s = "" if v is None else str(v).strip()
        return "" if s.lower() in ("none", "nan", "#n/a", "") else s

    aux_cargo = []
    aux_abono = []

    for row in rows[hdr_idx + 1:]:
        if not row or all(v is None for v in row):
            continue
        fecha = _to_date(row[col_f] if col_f is not None and col_f < len(row) else None)
        if fecha is None:
            continue
        pol      = _s(row, col_pol)
        dp_val   = _s(row, col_dp)
        det_val  = _s(row, col_det)
        # Concepto Aux: priorizar Desc. Póliza; si vacía, usar Detalle/Concepto
        concepto = dp_val or det_val
        cargo    = _to_float(row[col_cargo] if col_cargo is not None and col_cargo < len(row) else 0)
        abono    = _to_float(row[col_abono] if col_abono is not None and col_abono < len(row) else 0)

        base = {"fecha": fecha, "poliza": pol, "concepto": concepto, "matched": False}
        if cargo > 0:
            aux_cargo.append({**base, "monto": cargo})
        if abono > 0:
            aux_abono.append({**base, "monto": abono})

    return aux_cargo, aux_abono


def _match(banco_movs, aux_list, tipo_banco):
    """
    banco_movs: lista de movs del banco con campo 'dep' o 'ret' según tipo_banco.
    aux_list:   lista de entradas del auxiliar (cargo o abono).
    tipo_banco: 'dep' o 'ret'.
    Devuelve banco_movs con campo 'match_aux' relleno donde corresponda.
    """
    TOL_AMT  = 0.05

    aux_copy = [dict(a) for a in aux_list]  # copia para marcar matched

    for mov in banco_movs:
        monto_b = mov[tipo_banco]
        if monto_b <= 0:
            mov["match_aux"] = None
            continue
        fecha_b = mov["fecha"]
        best_i = None
        best_diff = None
        for i, a in enumerate(aux_copy):
            if a["matched"]:
                continue
            if abs(a["monto"] - monto_b) > TOL_AMT:
                continue
            diff_d = abs((a["fecha"] - fecha_b).days)
            if best_diff is None or diff_d < best_diff:
                best_diff = diff_d
                best_i = i
        if best_i is not None:
            aux_copy[best_i]["matched"] = True
            mov["match_aux"] = aux_copy[best_i]
        else:
            mov["match_aux"] = None

    return banco_movs, aux_copy


def _fmt_date(v):
    if v is None:
        return ""
    if isinstance(v, (date, datetime)):
        return v.strftime("%d/%m/%Y")
    return str(v)


def _fmt_money(v):
    if not v:
        return ""
    return f"${v:,.2f}"


# ── Generación de Excel ───────────────────────────────────────────────────────

def _generar_excel(depositos, retiros, aux_cargo_list, aux_abono_list):
    """Genera el Excel de conciliación en memoria y devuelve bytes."""
    try:
        from openpyxl import Workbook
        from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
        from openpyxl.utils import get_column_letter
    except ImportError:
        raise ImportError("Se requiere openpyxl para generar el Excel.")

    wb = Workbook()

    # Colores
    fill_conc_dep = PatternFill("solid", fgColor="C6EFCE")   # verde — dep conciliado
    fill_conc_ret = PatternFill("solid", fgColor="DDEEFF")   # azul claro — ret conciliado
    fill_no_conc  = PatternFill("solid", fgColor="FFC7CE")   # rojo — no conciliado
    fill_solo     = PatternFill("solid", fgColor="FFEB9C")   # amarillo — solo en banco
    fill_hdr      = PatternFill("solid", fgColor="1E3A8A")   # azul encabezado
    fill_tot      = PatternFill("solid", fgColor="FFC000")   # ámbar totales

    font_hdr  = Font(color="FFFFFF", bold=True, size=10)
    font_bold = Font(bold=True)
    font_tot  = Font(bold=True, size=10)
    al_c      = Alignment(horizontal="center", vertical="center")
    al_r      = Alignment(horizontal="right",  vertical="center")
    al_l      = Alignment(horizontal="left",   vertical="center")

    thin = Side(style="thin", color="CCCCCC")
    bord = Border(left=thin, right=thin, top=thin, bottom=thin)

    def _hdr_cell(ws, row, col, text, width=None):
        c = ws.cell(row=row, column=col, value=text)
        c.fill = fill_hdr; c.font = font_hdr
        c.alignment = al_c; c.border = bord
        if width:
            ws.column_dimensions[get_column_letter(col)].width = width
        return c

    def _data_cell(ws, row, col, val, fill=None, align=None, bold=False):
        c = ws.cell(row=row, column=col, value=val)
        if fill:  c.fill = fill
        if align: c.alignment = align
        if bold:  c.font = Font(bold=True)
        c.border = bord
        return c

    # ── Hoja Depósitos ────────────────────────────────────────────────────────
    ws_dep = wb.active
    ws_dep.title = "Depositos"
    ws_dep.row_dimensions[1].height = 22

    headers_dep = [
        ("Fecha Banco",   12),
        ("Descripción",   35),
        ("Depósito",      13),
        ("Saldo",         13),
        ("Estatus",       16),
        ("Fecha Aux",     12),
        ("Póliza",        12),
        ("Concepto Aux",  30),
        ("Monto Aux",     13),
    ]
    for c_idx, (h, w) in enumerate(headers_dep, 1):
        _hdr_cell(ws_dep, 1, c_idx, h, w)

    conc = 0; no_conc = 0
    mto_conc = 0.0; mto_no = 0.0

    for r_idx, mov in enumerate(depositos, 2):
        dep = mov["dep"]
        if dep <= 0:
            continue
        ma = mov.get("match_aux")
        if ma:
            fill = fill_conc_dep; est = "✅ CONCILIADO"; conc += 1; mto_conc += dep
        else:
            fill = fill_no_conc;  est = "❌ NO CONCILIADO"; no_conc += 1; mto_no += dep

        ws_dep.row_dimensions[r_idx].height = 18
        _data_cell(ws_dep, r_idx, 1, _fmt_date(mov["fecha"]), fill, al_c)
        _data_cell(ws_dep, r_idx, 2, mov["desc"],             fill, al_l)
        _data_cell(ws_dep, r_idx, 3, dep,                     fill, al_r)
        ws_dep.cell(r_idx, 3).number_format = '#,##0.00'
        _data_cell(ws_dep, r_idx, 4, mov["sal"] or None,      fill, al_r)
        ws_dep.cell(r_idx, 4).number_format = '#,##0.00'
        _data_cell(ws_dep, r_idx, 5, est,                     fill, al_c)
        if ma:
            _data_cell(ws_dep, r_idx, 6, _fmt_date(ma["fecha"]),  fill, al_c)
            _data_cell(ws_dep, r_idx, 7, ma["poliza"],            fill, al_c)
            _data_cell(ws_dep, r_idx, 8, ma["concepto"],          fill, al_l)
            _data_cell(ws_dep, r_idx, 9, ma["monto"],             fill, al_r)
            ws_dep.cell(r_idx, 9).number_format = '#,##0.00'
        else:
            for c in range(6, 10):
                _data_cell(ws_dep, r_idx, c, "", fill)

    # Solo en auxiliar (no conciliados del aux)
    aux_no_match = [a for a in aux_cargo_list if not a["matched"]]
    if aux_no_match:
        r_solo = ws_dep.max_row + 2
        ws_dep.cell(r_solo, 1).value = "⬇ Solo en Auxiliar (sin match en banco)"
        ws_dep.cell(r_solo, 1).font = Font(bold=True, italic=True, color="FF0000")
        for a in aux_no_match:
            r_solo += 1
            ws_dep.row_dimensions[r_solo].height = 18
            _data_cell(ws_dep, r_solo, 1, "",                    fill_solo, al_c)
            _data_cell(ws_dep, r_solo, 2, "(sin mov. en banco)", fill_solo, al_l)
            _data_cell(ws_dep, r_solo, 3, "",                    fill_solo)
            _data_cell(ws_dep, r_solo, 4, "",                    fill_solo)
            _data_cell(ws_dep, r_solo, 5, "⚠ SOLO AUXILIAR",    fill_solo, al_c)
            _data_cell(ws_dep, r_solo, 6, _fmt_date(a["fecha"]), fill_solo, al_c)
            _data_cell(ws_dep, r_solo, 7, a["poliza"],           fill_solo, al_c)
            _data_cell(ws_dep, r_solo, 8, a["concepto"],         fill_solo, al_l)
            _data_cell(ws_dep, r_solo, 9, a["monto"],            fill_solo, al_r)
            ws_dep.cell(r_solo, 9).number_format = '#,##0.00'

    # ── Hoja Retiros ──────────────────────────────────────────────────────────
    ws_ret = wb.create_sheet("Retiros")
    ws_ret.row_dimensions[1].height = 22

    headers_ret = [
        ("Fecha Banco",   12),
        ("Descripción",   35),
        ("Retiro",        13),
        ("Saldo",         13),
        ("Estatus",       16),
        ("Fecha Aux",     12),
        ("Póliza",        12),
        ("Concepto Aux",  30),
        ("Monto Aux",     13),
    ]
    for c_idx, (h, w) in enumerate(headers_ret, 1):
        _hdr_cell(ws_ret, 1, c_idx, h, w)

    conc_r = 0; no_conc_r = 0
    mto_conc_r = 0.0; mto_no_r = 0.0

    for r_idx, mov in enumerate(retiros, 2):
        ret_v = mov["ret"]
        if ret_v <= 0:
            continue
        ma = mov.get("match_aux")
        if ma:
            fill = fill_conc_ret; est = "✅ CONCILIADO"; conc_r += 1; mto_conc_r += ret_v
        else:
            fill = fill_no_conc;  est = "❌ NO CONCILIADO"; no_conc_r += 1; mto_no_r += ret_v

        ws_ret.row_dimensions[r_idx].height = 18
        _data_cell(ws_ret, r_idx, 1, _fmt_date(mov["fecha"]), fill, al_c)
        _data_cell(ws_ret, r_idx, 2, mov["desc"],             fill, al_l)
        _data_cell(ws_ret, r_idx, 3, ret_v,                   fill, al_r)
        ws_ret.cell(r_idx, 3).number_format = '#,##0.00'
        _data_cell(ws_ret, r_idx, 4, mov["sal"] or None,      fill, al_r)
        ws_ret.cell(r_idx, 4).number_format = '#,##0.00'
        _data_cell(ws_ret, r_idx, 5, est,                     fill, al_c)
        if ma:
            _data_cell(ws_ret, r_idx, 6, _fmt_date(ma["fecha"]),  fill, al_c)
            _data_cell(ws_ret, r_idx, 7, ma["poliza"],            fill, al_c)
            _data_cell(ws_ret, r_idx, 8, ma["concepto"],          fill, al_l)
            _data_cell(ws_ret, r_idx, 9, ma["monto"],             fill, al_r)
            ws_ret.cell(r_idx, 9).number_format = '#,##0.00'
        else:
            for c in range(6, 10):
                _data_cell(ws_ret, r_idx, c, "", fill)

    # Solo en auxiliar (abonos sin match)
    aux_abono_no_match = [a for a in aux_abono_list if not a["matched"]]
    if aux_abono_no_match:
        r_solo = ws_ret.max_row + 2
        ws_ret.cell(r_solo, 1).value = "⬇ Solo en Auxiliar — Abonos (sin match en banco)"
        ws_ret.cell(r_solo, 1).font = Font(bold=True, italic=True, color="FF0000")
        for a in aux_abono_no_match:
            r_solo += 1
            ws_ret.row_dimensions[r_solo].height = 18
            _data_cell(ws_ret, r_solo, 1, "",                    fill_solo, al_c)
            _data_cell(ws_ret, r_solo, 2, "(sin mov. en banco)", fill_solo, al_l)
            _data_cell(ws_ret, r_solo, 3, "",                    fill_solo)
            _data_cell(ws_ret, r_solo, 4, "",                    fill_solo)
            _data_cell(ws_ret, r_solo, 5, "⚠ SOLO AUXILIAR",    fill_solo, al_c)
            _data_cell(ws_ret, r_solo, 6, _fmt_date(a["fecha"]), fill_solo, al_c)
            _data_cell(ws_ret, r_solo, 7, a["poliza"],           fill_solo, al_c)
            _data_cell(ws_ret, r_solo, 8, a["concepto"],         fill_solo, al_l)
            _data_cell(ws_ret, r_solo, 9, a["monto"],            fill_solo, al_r)
            ws_ret.cell(r_solo, 9).number_format = '#,##0.00'

    # ── Hoja Resumen ──────────────────────────────────────────────────────────
    ws_res = wb.create_sheet("Resumen")
    ws_res.column_dimensions["A"].width = 32
    ws_res.column_dimensions["B"].width = 18
    ws_res.column_dimensions["C"].width = 18

    def _res_hdr(row, text):
        c = ws_res.cell(row=row, column=1, value=text)
        c.fill = fill_hdr; c.font = font_hdr
        ws_res.merge_cells(start_row=row, start_column=1, end_row=row, end_column=3)
        c.alignment = al_c; c.border = bord
        ws_res.row_dimensions[row].height = 20

    def _res_row(row, label, cant, monto, fill=None):
        for col, val, al in [(1, label, al_l), (2, cant, al_c), (3, monto, al_r)]:
            c = ws_res.cell(row=row, column=col, value=val)
            if fill: c.fill = fill
            c.alignment = al; c.border = bord
        ws_res.cell(row, 3).number_format = '#,##0.00'
        ws_res.row_dimensions[row].height = 18

    r = 1
    _res_hdr(r, "DEPÓSITOS"); r += 1
    _res_row(r, "Conciliados",     conc,   mto_conc,  fill_conc_dep); r += 1
    _res_row(r, "No conciliados",  no_conc, mto_no,   fill_no_conc);  r += 1
    total_dep_cnt = conc + no_conc
    total_dep_mto = mto_conc + mto_no
    c_tot = ws_res.cell(r, 1, "TOTAL DEPÓSITOS"); c_tot.fill = fill_tot; c_tot.font = font_tot; c_tot.border = bord; c_tot.alignment = al_l
    c_tot2 = ws_res.cell(r, 2, total_dep_cnt);    c_tot2.fill = fill_tot; c_tot2.font = font_tot; c_tot2.border = bord; c_tot2.alignment = al_c
    c_tot3 = ws_res.cell(r, 3, total_dep_mto);    c_tot3.fill = fill_tot; c_tot3.font = font_tot; c_tot3.border = bord; c_tot3.alignment = al_r; c_tot3.number_format = '#,##0.00'
    ws_res.row_dimensions[r].height = 20; r += 2

    _res_hdr(r, "RETIROS"); r += 1
    _res_row(r, "Conciliados",     conc_r,    mto_conc_r, fill_conc_ret); r += 1
    _res_row(r, "No conciliados",  no_conc_r, mto_no_r,   fill_no_conc);  r += 1
    total_ret_cnt = conc_r + no_conc_r
    total_ret_mto = mto_conc_r + mto_no_r
    c_tot = ws_res.cell(r, 1, "TOTAL RETIROS");   c_tot.fill = fill_tot; c_tot.font = font_tot; c_tot.border = bord; c_tot.alignment = al_l
    c_tot2 = ws_res.cell(r, 2, total_ret_cnt);    c_tot2.fill = fill_tot; c_tot2.font = font_tot; c_tot2.border = bord; c_tot2.alignment = al_c
    c_tot3 = ws_res.cell(r, 3, total_ret_mto);    c_tot3.fill = fill_tot; c_tot3.font = font_tot; c_tot3.border = bord; c_tot3.alignment = al_r; c_tot3.number_format = '#,##0.00'
    ws_res.row_dimensions[r].height = 20; r += 2

    # Leyenda
    _res_hdr(r, "LEYENDA DE COLORES"); r += 1
    for label, fill in [
        ("Verde  — Depósito conciliado",       fill_conc_dep),
        ("Azul   — Retiro conciliado",          fill_conc_ret),
        ("Rojo   — No conciliado",              fill_no_conc),
        ("Amarillo — Solo en auxiliar / banco", fill_solo),
    ]:
        for col in range(1, 4):
            c = ws_res.cell(r, col, label if col == 1 else "")
            c.fill = fill; c.border = bord; c.alignment = al_l
        ws_res.row_dimensions[r].height = 18; r += 1

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


# ── Estado de sesión ──────────────────────────────────────────────────────────
for key in ("cba_resultado_bytes", "cba_resumen"):
    if key not in st.session_state:
        st.session_state[key] = None

# ── UI — Carga de archivos ────────────────────────────────────────────────────
st.subheader("1️⃣  Cargar archivos")
col_banco, col_aux = st.columns(2)

with col_banco:
    banco_file = st.file_uploader(
        "🏦 Archivo del banco (.xlsx)",
        type=["xlsx", "xls"],
        help="Estado de cuenta o movimientos bancarios en Excel.",
        key="cba_banco",
    )

with col_aux:
    aux_file = st.file_uploader(
        "📋 Auxiliar contable (.xlsx)",
        type=["xlsx", "xls"],
        help="Auxiliar con columnas Fecha, Póliza, Concepto, Cargo, Abono.",
        key="cba_aux",
    )

# ── Leyenda ───────────────────────────────────────────────────────────────────
st.markdown("""
<div style="margin:0.5rem 0 1rem; font-size:0.85rem;">
  <span class="leg" style="background:#C6EFCE;"></span>Depósito conciliado &nbsp;
  <span class="leg" style="background:#DDEEFF;"></span>Retiro conciliado &nbsp;
  <span class="leg" style="background:#FFC7CE;"></span>No conciliado &nbsp;
  <span class="leg" style="background:#FFEB9C;"></span>Solo en auxiliar
</div>
""", unsafe_allow_html=True)

# ── Botón generar ─────────────────────────────────────────────────────────────
st.subheader("2️⃣  Generar conciliación")

col_btn, col_dl = st.columns([1, 2])
with col_btn:
    can_gen = banco_file is not None and aux_file is not None
    generar = st.button(
        "🔀 Generar conciliación",
        disabled=not can_gen,
        type="primary",
        use_container_width=True,
    )
    if not can_gen:
        st.caption("Carga ambos archivos para continuar.")

if generar:
    st.session_state.cba_resultado_bytes = None
    st.session_state.cba_resumen = None

    with st.spinner("Leyendo archivos..."):
        try:
            from openpyxl import load_workbook

            # Banco
            wb_banco = load_workbook(
                filename=io.BytesIO(banco_file.read()),
                read_only=True, data_only=True,
            )
            movs_banco = _read_banco(wb_banco)
            wb_banco.close()

            # Auxiliar
            wb_aux = load_workbook(
                filename=io.BytesIO(aux_file.read()),
                read_only=True, data_only=True,
            )
            aux_cargo_raw, aux_abono_raw = _read_auxiliar(wb_aux)
            wb_aux.close()

        except Exception as e:
            st.error(f"❌ Error al leer archivos: {e}")
            with st.expander("Ver detalle"):
                st.code(traceback.format_exc())
            st.stop()

    with st.spinner("Conciliando..."):
        try:
            depositos = [m for m in movs_banco if m["dep"] > 0]
            retiros   = [m for m in movs_banco if m["ret"] > 0]

            depositos, aux_cargo_list = _match(depositos, aux_cargo_raw, "dep")
            retiros,   aux_abono_list = _match(retiros,   aux_abono_raw, "ret")

        except Exception as e:
            st.error(f"❌ Error en conciliación: {e}")
            with st.expander("Ver detalle"):
                st.code(traceback.format_exc())
            st.stop()

    with st.spinner("Generando Excel..."):
        try:
            excel_bytes = _generar_excel(depositos, retiros, aux_cargo_list, aux_abono_list)
            st.session_state.cba_resultado_bytes = excel_bytes

            # Resumen para mostrar en UI
            conc_d  = sum(1 for m in depositos if m.get("match_aux"))
            no_d    = sum(1 for m in depositos if not m.get("match_aux"))
            conc_r  = sum(1 for m in retiros   if m.get("match_aux"))
            no_r    = sum(1 for m in retiros   if not m.get("match_aux"))
            st.session_state.cba_resumen = {
                "conc_d": conc_d, "no_d": no_d,
                "conc_r": conc_r, "no_r": no_r,
                "total_d": len(depositos), "total_r": len(retiros),
                "n_banco": len(movs_banco),
                "n_aux_cargo": len(aux_cargo_raw),
                "n_aux_abono": len(aux_abono_raw),
            }
        except Exception as e:
            st.error(f"❌ Error generando Excel: {e}")
            with st.expander("Ver detalle"):
                st.code(traceback.format_exc())
            st.stop()

# ── Resultado ─────────────────────────────────────────────────────────────────
if st.session_state.cba_resultado_bytes and st.session_state.cba_resumen:
    res = st.session_state.cba_resumen

    st.success("✅ Conciliación completada")

    # Métricas
    m1, m2, m3, m4 = st.columns(4)
    with m1:
        st.markdown(f'<div class="cba-stat"><div class="val">{res["n_banco"]}</div><div class="lbl">Movimientos banco</div></div>', unsafe_allow_html=True)
    with m2:
        st.markdown(f'<div class="cba-stat"><div class="val" style="color:#166534">{res["conc_d"]}/{res["total_d"]}</div><div class="lbl">Depósitos conciliados</div></div>', unsafe_allow_html=True)
    with m3:
        st.markdown(f'<div class="cba-stat"><div class="val" style="color:#1E40AF">{res["conc_r"]}/{res["total_r"]}</div><div class="lbl">Retiros conciliados</div></div>', unsafe_allow_html=True)
    with m4:
        no_total = res["no_d"] + res["no_r"]
        color_no = "#991B1B" if no_total > 0 else "#166534"
        st.markdown(f'<div class="cba-stat"><div class="val" style="color:{color_no}">{no_total}</div><div class="lbl">Sin conciliar</div></div>', unsafe_allow_html=True)

    # Descarga
    st.download_button(
        label="⬇️  Descargar Excel de conciliación",
        data=st.session_state.cba_resultado_bytes,
        file_name="Conciliacion_Banco_Auxiliar.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True,
    )

else:
    st.markdown("""
<div class="cba-card" style="text-align:center; padding: 3rem 2rem;">
  <div style="font-size:3rem; margin-bottom:1rem;">🔀</div>
  <h3 style="color:#1E3A8A; margin:0 0 .5rem;">Conciliación Banco vs Auxiliar</h3>
  <p style="color:#64748B; max-width:400px; margin:0 auto;">
    Carga el Excel del banco y el auxiliar contable, luego presiona
    <strong>Generar conciliación</strong>. El resultado se descarga como Excel
    con hojas de Depósitos, Retiros y Resumen.
  </p>
</div>
""", unsafe_allow_html=True)

st.markdown("---")
st.caption("Módulo Conciliación Banco vs Auxiliar · v1.1  ·  Tolerancia ±$0.05 · Concepto Aux = Desc. Póliza")
