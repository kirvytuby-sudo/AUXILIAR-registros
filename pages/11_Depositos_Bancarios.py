"""
Módulo: Depósitos Bancarios
Genera automáticamente la póliza de depósitos bancarios a partir de los
estados de cuenta de BBVA, Banorte e Inbursa.
"""

import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from io import BytesIO
import re
from datetime import datetime

st.set_page_config(page_title="Depósitos Bancarios", page_icon="🏦", layout="wide")

# ─── Estilos ──────────────────────────────────────────────────────────────────
import _theme
_theme.aplicar_header("🏦 Depósitos Bancarios", "BBVA, Banorte e Inbursa → póliza de depósitos con cuentas de tránsito")
st.caption("Genera la póliza contable desde los estados de cuenta de BBVA, Banorte e Inbursa.")

# ─── Constantes contables ──────────────────────────────────────────────────────
CARGOS = {
    "BANORTE": ("102-01-0001-0001", "Banorte"),
    "BBVA":    ("102-01-0001-0002", "BBVA"),
    "INBURSA": ("102-01-0001-0005", "INBURSA"),
}

# Orden de columnas de abonos (cols 12-19 en la póliza)
ABONOS = [
    {"col": 12, "cuenta": "106-01-0001-0010", "nombre": "DEPOSITO EN TRANSITO T AMERICAN EXPRESS"},
    {"col": 13, "cuenta": "106-01-0001-0005", "nombre": "DEPOSITO EN TRANSITO  EFECTIVALE"},
    {"col": 14, "cuenta": "106-01-0001-0009", "nombre": "DEPOSITO EN TRANSITO  TICKET CARD EDENRED"},
    {"col": 15, "cuenta": "101-01-0001",       "nombre": "Fondo Fijo de Caja"},
    {"col": 16, "cuenta": "106-01-0001-0011",  "nombre": "DEPOSITO EN TRANSITO T BANORTE"},
    {"col": 17, "cuenta": "106-01-0001-0003",  "nombre": "DEPOSITO EN TRANSITO  SMARTBT - SHELL FLEET"},
    {"col": 18, "cuenta": "106-01-0001-0013",  "nombre": "BBVA"},
    {"col": 19, "cuenta": "106-01-0001-0012",  "nombre": "DEPOSITO EN TRANSITO T INBURSA"},
]

COL_ABONO_IDX = {a["col"]: a for a in ABONOS}

# ─── Clasificación de depósitos ───────────────────────────────────────────────
def _hora_citi(desc: str) -> int:
    """Extrae la hora del SPEI de CITI para distinguir AMEX vs Shell."""
    m = re.search(r"HR LIQ:\s*(\d{2}):", desc)
    return int(m.group(1)) if m else 0


def clasificar_bbva(desc: str, monto: float):
    """
    BBVA:
      *AMEX* en cualquier parte           → col 12 (AMEX transit)
      VENTAS PUNTOS TDC / VENTAS CREDITO
        / TERMINALES PV / TDC INTER       → col 18 (Dep. Tránsito Tarjetas Bancomer)
      DEPOSITO EN EFECTIVO                → col 15 (Caja 101-01-0001)
      Cualquier otra descripción          → None  (no clasificado)
    """
    d = desc.upper()
    if "AMEX" in d:
        return 12
    if ("VENTAS PUNTOS TDC" in d or "VENTAS CREDITO" in d
            or "TERMINALES PUNTO DE VENTA" in d
            or "VENTAS TDC INTER" in d or "TDC INTER" in d):
        return 18
    if "DEPOSITO EN EFECTIVO" in d or "DEP.EFECTIVO" in d or "DEP EN EFECTIVO" in d:
        return 15
    return None


def clasificar_inbursa(desc: str, monto: float):
    """INBURSA: solo 'INBURED' → INBURSA transit (col 19). Resto → None."""
    d = desc.upper()
    if "INBURED" in d:
        return 19
    return None


def clasificar_banorte(desc: str, monto: float):
    """
    BANORTE: busca la palabra clave en la descripción y devuelve la columna
    de abono. Devuelve None si no hay coincidencia (no se registra).
    """
    d = desc.upper()

    # Cheques → no clasificar
    if "DEP. CH." in d or "CHEQUE SBC" in d:
        return None

    # Compensación desfase → ignorar
    if "COMPENSACION DESFASE" in d:
        return None

    # FELUSA SPEI (transferencia interna desde cualquier banco) → no clasificar
    if "SPEI RECIBIDO" in d and "SERVICIOS FELUSA" in d:
        return None

    # SHELL FLEET / SMARTBT (excluye BNET genérico que no es Shell)
    if "SHELL" in d or "SMARTBT" in d:
        return 17

    # AMERICAN EXPRESS / CITI MEXICO
    if "AMERICAN EXPRESS" in d or "BCO:0124" in d or "CITI MEXICO" in d:
        m = re.search(r"HR LIQ:\s*(\d{2}):", desc)
        hora = int(m.group(1)) if m else 0
        return 17 if hora >= 12 else 12

    # EFECTIVALE / SANTANDER
    if "EFECTIVALE" in d or "EFE8908015L3" in d or "BCO:0014" in d or (
            "SANTANDER" in d and "SPEI RECIBIDO" in d):
        return 13

    # EDENRED / HSBC
    if "EDENRED" in d or "HSBCPGMD" in d or "BCO:0021" in d:
        return 14

    # Efectivo en caja
    if "DEP.EFECTIVO" in d or "DEPOSITO EN EFECTIVO" in d:
        return 15

    # SERV <NOMBRE> <DÍGITOS>C/D
    if "07277262C" in d or "07277262D" in d or (
            "SERV" in d and re.search(r'\d{5,}[CD]', d)):
        if "AMERICAN" in d: return 12
        if "EFECTIVALE" in d: return 13
        if "EDENRED" in d or "TICKET" in d: return 14
        if "SHELL" in d or "SMARTBT" in d: return 17
        if "INBURSA" in d: return 19
        return 16

    return None


CLASIFICADORES = {
    "BBVA":    clasificar_bbva,
    "BANORTE": clasificar_banorte,
    "INBURSA": clasificar_inbursa,
}

CARGO_COL = {"BBVA": 9, "BANORTE": 8, "INBURSA": 10}
# ─── Lectura de estados de cuenta ─────────────────────────────────────────────
def leer_banco(file_obj, banco: str) -> tuple[list, list]:
    wb = openpyxl.load_workbook(file_obj, read_only=True, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    wb.close()

    inicio = 2
    for i, r in enumerate(rows):
        if r and str(r[0]).strip().lower() == "fecha":
            inicio = i + 1
            break

    ok, sin_clasif = [], []
    fn = CLASIFICADORES[banco]
    col_cargo = CARGO_COL[banco]

    for idx, r in enumerate(rows[inicio:], start=inicio):
        if not r or r[0] is None:
            continue
        fecha = r[0]
        if not hasattr(fecha, "day"):
            continue
        desc  = str(r[1] or "").strip()
        monto = r[2]
        if not monto or monto <= 0:
            continue

        col_abono = fn(desc, monto)
        if col_abono is None:
            sin_clasif.append({"fecha": fecha, "banco": banco,
                               "descripcion": desc[:80], "monto": monto})
        else:
            ok.append({
                "fecha":      fecha,
                "banco":      banco,
                "ref":        f"DEPOSITOS {banco}",
                "desc":       desc,
                "monto":      monto,
                "col_cargo":  col_cargo,
                "col_abono":  col_abono,
                "fila_excel": idx + 1,
            })

    return ok, sin_clasif


def marcar_estado_cuenta(file_bytes: bytes, filas_usadas: set, banco: str = "") -> bytes:
    BANK_STYLE = {
        "BBVA":    ("A9C9EF", "004481"),
        "BANORTE": ("FBBDBD", "C8102E"),
        "INBURSA": ("C8E6C9", "1B5E20"),
    }
    fondo, color_txt = BANK_STYLE.get(banco.upper(), ("A8D5B5", "1A5276"))
    row_fill = PatternFill("solid", fgColor=fondo)
    _s = Side(style="thin", color=color_txt)
    row_border = Border(left=_s, right=_s, top=_s, bottom=_s)

    wb = openpyxl.load_workbook(BytesIO(file_bytes))
    ws = wb.active
    for fila in sorted(filas_usadas):
        for col in range(1, ws.max_column + 1):
            cell = ws.cell(row=fila, column=col)
            cell.fill = row_fill
            cell.border = row_border
            old = cell.font
            cell.font = Font(
                name=old.name or "Calibri",
                size=old.size or 11,
                bold=True,
                color=color_txt,
                italic=old.italic,
            )
    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()
# ─── Generación del Excel ────────────────────────────────────────────────────
def generar_excel(registros: list, plantilla=None) -> bytes:
    FONT_NAME = "Aptos Narrow"
    FONT_SIZE = 11

    F_ADMIN   = PatternFill("solid", fgColor="4B5563")
    F_H_BBVA  = PatternFill("solid", fgColor="2563EB")
    F_H_BNT   = PatternFill("solid", fgColor="DC2626")
    F_H_INB   = PatternFill("solid", fgColor="059669")
    F_H_TCARG = PatternFill("solid", fgColor="D97706")
    F_H_AMEX  = PatternFill("solid", fgColor="4F46E5")
    F_H_EFEC  = PatternFill("solid", fgColor="E11D48")
    F_H_EDEN  = PatternFill("solid", fgColor="EA580C")
    F_H_CAJA  = PatternFill("solid", fgColor="F59E0B")
    F_H_BNTTR = PatternFill("solid", fgColor="B91C1C")
    F_H_SHLL  = PatternFill("solid", fgColor="CA8A04")
    F_H_BBVAT = PatternFill("solid", fgColor="1D4ED8")
    F_H_INBTR = PatternFill("solid", fgColor="047857")
    F_H_TABON = PatternFill("solid", fgColor="15803D")
    F_H_DIFF  = PatternFill("solid", fgColor="7C3AED")
    F_GRAY2   = PatternFill("solid", fgColor="374151")
    F_BBVA_1  = PatternFill("solid", fgColor="EFF6FF")
    F_BBVA_2  = PatternFill("solid", fgColor="DBEAFE")
    F_BNT_1   = PatternFill("solid", fgColor="FFF1F2")
    F_BNT_2   = PatternFill("solid", fgColor="FFE4E6")
    F_INB_1   = PatternFill("solid", fgColor="F0FDF4")
    F_INB_2   = PatternFill("solid", fgColor="DCFCE7")
    F_NONE    = PatternFill(fill_type=None)

    _S = Side(style="thin", color="999999")
    BORDER = Border(left=_S, right=_S, top=_S, bottom=_S)
    _SH = Side(style="medium", color="FFFFFF")
    BORDER_H = Border(left=_SH, right=_SH, top=_SH, bottom=_SH)

    def fnt(bold=False, color="000000", size=None, italic=False):
        return Font(name=FONT_NAME, size=size or FONT_SIZE,
                    bold=bold, color=color, italic=italic)

    A_CTR   = Alignment(horizontal="center", vertical="center", wrap_text=False)
    A_CTR_W = Alignment(horizontal="center", vertical="center", wrap_text=True)
    A_LEFT  = Alignment(horizontal="left",   vertical="center", wrap_text=False)
    A_RIGHT = Alignment(horizontal="right",  vertical="center", wrap_text=False)

    FMT_NUM  = '#,##0.00'
    FMT_DATE = 'DD/MM/YYYY'

    def set_cell(ws, row, col, value=None, font=None, fill=F_NONE,
                 align=A_CTR, num_format=None, border=None):
        c = ws.cell(row=row, column=col, value=value)
        if font:       c.font = font
        if fill:       c.fill = fill
        if align:      c.alignment = align
        if num_format: c.number_format = num_format
        if border:     c.border = border
        return c
    if plantilla is not None:
        wb = openpyxl.load_workbook(plantilla)
        ws = wb["POLIZA"] if "POLIZA" in wb.sheetnames else wb.active
        for row_idx in range(1, ws.max_row + 1):
            for col_idx in range(1, ws.max_column + 1):
                ws.cell(row=row_idx, column=col_idx).value = None
    else:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "POLIZA"

    for pol_idx in range(21):
        set_cell(ws, 1, pol_idx + 1, value=pol_idx,
                 font=fnt(color="000000"), fill=F_NONE, align=A_CTR)

    fnt_cta = fnt(bold=False, color="FFFFFF", size=9, italic=True)
    for pol_idx, banco in [(8, "BANORTE"), (9, "BBVA"), (10, "INBURSA")]:
        set_cell(ws, 2, pol_idx + 1, value=CARGOS[banco][0],
                 font=fnt_cta, fill=F_GRAY2, align=A_CTR, border=BORDER)
    for a in ABONOS:
        set_cell(ws, 2, a["col"] + 1, value=a["cuenta"],
                 font=fnt_cta, fill=F_GRAY2, align=A_CTR, border=BORDER)

    fnt_h = lambda: fnt(bold=True, color="FFFFFF")
    for col, lbl in [(1,"TIPO"),(2,"FECHA"),(3,"REFERENCIA"),(4,"CONCEPTO"),
                     (5,"ERROR"),(6,"UIDD"),(7,"NÚM PÓLIZA"),(8,"PROCESADO")]:
        set_cell(ws, 3, col, value=lbl, font=fnt_h(), fill=F_ADMIN,
                 align=A_CTR, border=BORDER_H)

    for pol_idx, banco, fill_h in [
        (8, "BANORTE", F_H_BNT), (9, "BBVA", F_H_BBVA), (10, "INBURSA", F_H_INB)]:
        set_cell(ws, 3, pol_idx + 1, value=CARGOS[banco][1].strip(),
                 font=fnt_h(), fill=fill_h, align=A_CTR, border=BORDER_H)

    set_cell(ws, 3, 12, value="TOTAL CARGOS",
             font=fnt_h(), fill=F_H_TCARG, align=A_CTR, border=BORDER_H)

    ABONO_FILLS = {12:F_H_AMEX,13:F_H_EFEC,14:F_H_EDEN,15:F_H_CAJA,
                   16:F_H_BNTTR,17:F_H_SHLL,18:F_H_BBVAT,19:F_H_INBTR}
    for a in ABONOS:
        dark = a["col"] in {15, 17}
        set_cell(ws, 3, a["col"]+1, value=a["nombre"],
                 font=fnt(bold=True, color="1A1A1A" if dark else "FFFFFF"),
                 fill=ABONO_FILLS.get(a["col"], F_ADMIN), align=A_CTR_W, border=BORDER_H)

    set_cell(ws, 3, 21, value="TOTAL ABONOS",
             font=fnt_h(), fill=F_H_TABON, align=A_CTR, border=BORDER_H)
    set_cell(ws, 3, 22, value="DIFERENCIA",
             font=fnt_h(), fill=F_H_DIFF,  align=A_CTR, border=BORDER_H)
    orden_banco = {"BBVA": 0, "INBURSA": 1, "BANORTE": 2}
    registros_sorted = sorted(registros, key=lambda x: (orden_banco[x["banco"]], x["fecha"]))

    FILLS_BANCO = {
        "BBVA":    (F_BBVA_1, F_BBVA_2),
        "INBURSA": (F_INB_1,  F_INB_2),
        "BANORTE": (F_BNT_1,  F_BNT_2),
    }
    BANCO_COLOR = {"BBVA": "2563EB", "BANORTE": "DC2626", "INBURSA": "059669"}
    BANCO_ABREV = {"BBVA": "BBV", "BANORTE": "BNT", "INBURSA": "INB"}

    for fila_num, r in enumerate(registros_sorted, start=4):
        f1, f2 = FILLS_BANCO[r["banco"]]
        fill_row = f1 if fila_num % 2 == 0 else f2
        fn_dat   = fnt(color="000000")

        def dat(col, val, num_fmt=None, align=A_CTR):
            c = set_cell(ws, fila_num, col, value=val,
                         font=fn_dat, fill=fill_row, align=align, border=BORDER)
            if num_fmt: c.number_format = num_fmt
            return c

        monto = r["monto"]
        # Col TIPO: negrita con color del banco
        set_cell(ws, fila_num, 1, value="I",
                 font=fnt(bold=True, color=BANCO_COLOR[r["banco"]]),
                 fill=fill_row, align=A_CTR, border=BORDER)
        dat(2, r["fecha"], num_fmt=FMT_DATE)
        dat(3, r["ref"],   align=A_LEFT)
        dat(4, r["ref"],   align=A_LEFT)
        for col in [5, 6, 7, 8]: dat(col, None)
        dat(r["col_cargo"] + 1, monto, num_fmt=FMT_NUM, align=A_RIGHT)
        dat(12, monto, num_fmt=FMT_NUM, align=A_RIGHT)
        dat(r["col_abono"] + 1, monto, num_fmt=FMT_NUM, align=A_RIGHT)
        dat(21, monto, num_fmt=FMT_NUM, align=A_RIGHT)

        col_L = get_column_letter(12)
        col_U = get_column_letter(21)
        c_diff = ws.cell(row=fila_num, column=22,
                         value=f"={col_L}{fila_num}-{col_U}{fila_num}")
        c_diff.font   = fnt(bold=True, color="7C3AED")
        c_diff.fill   = fill_row
        c_diff.border = BORDER
        c_diff.alignment  = A_RIGHT
        c_diff.number_format = FMT_NUM
        ws.row_dimensions[fila_num].height = 18
    anchos = {
        1:14.4, 2:13.0, 3:36.9, 4:26.3, 5:6.7, 6:5.3, 7:11.3, 8:11.6,
        9:16.0, 10:16.0, 11:16.0, 12:16.3, 13:32.0, 14:26.0, 15:30.0,
        16:18.0, 17:28.0, 18:32.0, 19:12.0, 20:28.0, 21:14.1, 22:12.6,
    }
    for col_num, w in anchos.items():
        ws.column_dimensions[get_column_letter(col_num)].width = w
    ws.row_dimensions[1].height = 14
    ws.row_dimensions[2].height = 20
    ws.row_dimensions[3].height = 36
    ws.freeze_panes = "B4"

    if plantilla is None:
        wc = wb.create_sheet("CUENTAS")
        fn_th  = fnt(bold=True, color="FFFFFF")
        fn_row = fnt(color="000000")
        F_TH_C = PatternFill("solid", fgColor="4B5563")
        F_TH_A = PatternFill("solid", fgColor="15803D")
        F_ROW1 = PatternFill("solid", fgColor="EFF6FF")
        F_ROW2 = PatternFill("solid", fgColor="F0FDF4")
        A_L  = Alignment(horizontal="left",   vertical="center")
        A_C2 = Alignment(horizontal="center", vertical="center")

        set_cell(wc,1,1,"CARGOS",font=fn_th,fill=F_TH_C,align=A_C2,border=BORDER_H)
        set_cell(wc,1,2,"",      font=fn_th,fill=F_TH_C,align=A_C2,border=BORDER_H)
        set_cell(wc,1,4,"ABONOS",font=fn_th,fill=F_TH_A,align=A_C2,border=BORDER_H)
        set_cell(wc,1,5,"",      font=fn_th,fill=F_TH_A,align=A_C2,border=BORDER_H)
        for col,lbl,fl in [(1,"N° Cuenta",F_TH_C),(2,"Banco",F_TH_C),
                            (4,"N° Cuenta",F_TH_A),(5,"Nombre",F_TH_A)]:
            set_cell(wc,2,col,lbl,font=fn_th,fill=fl,align=A_C2,border=BORDER_H)
        for i,(banco,(cuenta,nombre)) in enumerate(CARGOS.items(),start=3):
            fl = F_ROW1 if i%2==0 else F_ROW2
            set_cell(wc,i,1,cuenta,        font=fn_row,fill=fl,align=A_L,border=BORDER)
            set_cell(wc,i,2,nombre.strip(),font=fn_row,fill=fl,align=A_L,border=BORDER)
        for i,a in enumerate(ABONOS,start=3):
            fl = F_ROW1 if i%2==0 else F_ROW2
            set_cell(wc,i,4,a["cuenta"],font=fn_row,fill=fl,align=A_L,border=BORDER)
            set_cell(wc,i,5,a["nombre"],font=fn_row,fill=fl,align=A_L,border=BORDER)
        for col_num,w in {1:20.0,2:12.0,4:20.0,5:48.0}.items():
            wc.column_dimensions[get_column_letter(col_num)].width = w
        wc.row_dimensions[1].height = 22
        wc.row_dimensions[2].height = 22

    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()
# ─── UI ─────────────────────────────────────────────────────────────────────
col1, col2, col3 = st.columns(3)

with col1:
    st.markdown('<div class="upload-box">', unsafe_allow_html=True)
    st.markdown("**🟦 BBVA Bancomer**")
    file_bbva = st.file_uploader("Estado de cuenta BBVA (.xlsx)",
                                  type=["xlsx"], key="bbva",
                                  label_visibility="collapsed")
    st.markdown('</div>', unsafe_allow_html=True)

with col2:
    st.markdown('<div class="upload-box">', unsafe_allow_html=True)
    st.markdown("**🟥 Banorte**")
    file_banorte = st.file_uploader("Estado de cuenta Banorte (.xlsx)",
                                     type=["xlsx"], key="banorte",
                                     label_visibility="collapsed")
    st.markdown('</div>', unsafe_allow_html=True)

with col3:
    st.markdown('<div class="upload-box">', unsafe_allow_html=True)
    st.markdown("**🟧 Inbursa**")
    file_inbursa = st.file_uploader("Estado de cuenta Inbursa (.xlsx)",
                                     type=["xlsx"], key="inbursa",
                                     label_visibility="collapsed")
    st.markdown('</div>', unsafe_allow_html=True)

st.markdown('<div class="upload-box">', unsafe_allow_html=True)
st.markdown("**📋 PLANTILLA DE DEPOSITOS**  *(opcional — si se sube, los datos se escriben sobre ella)*")
file_plantilla = st.file_uploader(
    "Plantilla de Depósitos (.xlsx)",
    type=["xlsx", "xlsm"], key="plantilla",
    label_visibility="collapsed",
)
st.markdown('</div>', unsafe_allow_html=True)

st.divider()

if "dep_result" not in st.session_state:
    st.session_state.dep_result = None

if st.button("⚙️ Generar Póliza", type="primary",
             disabled=not any([file_bbva, file_banorte, file_inbursa])):

    todos_ok, todos_nc = [], []
    archivos_bytes = {}
    progress = st.progress(0)

    archivos = [
        (file_bbva,    "BBVA",    "🟦"),
        (file_banorte, "BANORTE", "🟥"),
        (file_inbursa, "INBURSA", "🟧"),
    ]

    for i, (f, banco, ico) in enumerate(archivos):
        if f:
            raw = f.read()
            archivos_bytes[banco] = raw
            with st.spinner(f"{ico} Leyendo {banco}..."):
                try:
                    ok, nc = leer_banco(BytesIO(raw), banco)
                    todos_ok.extend(ok)
                    todos_nc.extend(nc)
                    st.success(f"{ico} {banco}: **{len(ok)}** depósitos clasificados"
                               + (f", {len(nc)} sin clasificar" if nc else ""))
                except Exception as e:
                    st.error(f"Error leyendo {banco}: {e}")
        progress.progress((i + 1) / 3)
    if todos_ok:
        excel_bytes = generar_excel(todos_ok, plantilla=file_plantilla)
        nombre_archivo = f"DEPOSITOS BANCARIOS {datetime.now().strftime('%Y-%m')}.xlsx"

        marked = {}
        for _, banco, ico in archivos:
            if banco in archivos_bytes:
                filas = {r["fila_excel"] for r in todos_ok if r["banco"] == banco}
                if filas:
                    marked[banco] = {
                        "ico":   ico,
                        "bytes": marcar_estado_cuenta(archivos_bytes[banco], filas, banco),
                        "n":     len(filas),
                    }

        st.session_state.dep_result = {
            "todos_ok":       todos_ok,
            "todos_nc":       todos_nc,
            "excel_bytes":    excel_bytes,
            "nombre_archivo": nombre_archivo,
            "marked":         marked,
        }
    else:
        st.session_state.dep_result = None

if st.session_state.dep_result:
    import pandas as pd
    res      = st.session_state.dep_result
    todos_ok = res["todos_ok"]
    todos_nc = res["todos_nc"]

    st.subheader("📊 Resumen")
    df_ok = pd.DataFrame(todos_ok)
    total_general = df_ok["monto"].sum()

    sc = st.columns(4)
    for ci, (banco, ico) in enumerate([("BBVA","🟦"),("BANORTE","🟥"),("INBURSA","🟧")]):
        sub = df_ok[df_ok["banco"] == banco]
        sc[ci].metric(f"{ico} {banco}",
                      f"${sub['monto'].sum():,.2f}",
                      f"{len(sub)} mov.")
    sc[3].metric("💰 TOTAL", f"${total_general:,.2f}", f"{len(df_ok)} mov.")

    st.subheader("📋 Vista previa")
    abono_nombre = {a["col"]: a["nombre"][:30] for a in ABONOS}
    cargo_nombre = {9:"BANORTE",10:"BBVA",11:"INBURSA"}

    preview_rows = []
    orden = {"BBVA":0,"BANORTE":1,"INBURSA":2}
    for r in sorted(todos_ok, key=lambda x:(x["fecha"],orden[x["banco"]])):
        preview_rows.append({
            "Fecha":       r["fecha"].strftime("%d/%m/%Y") if hasattr(r["fecha"],"strftime") else str(r["fecha"]),
            "Banco":       r["banco"],
            "Cargo":       cargo_nombre.get(r["col_cargo"]+1, str(r["col_cargo"])),
            "Abono":       abono_nombre.get(r["col_abono"], str(r["col_abono"])),
            "Monto":       r["monto"],
            "Descripción": r["desc"][:55],
        })

    st.dataframe(pd.DataFrame(preview_rows).style.format({"Monto":"${:,.2f}"}),
                 use_container_width=True, height=400)

    if todos_nc:
        st.subheader(f"⚠️ Sin clasificar ({len(todos_nc)} movimientos — no incluidos)")
        df_nc = pd.DataFrame(todos_nc)
        df_nc["fecha"] = df_nc["fecha"].apply(
            lambda x: x.strftime("%d/%m/%Y") if hasattr(x,"strftime") else str(x))
        st.dataframe(df_nc.style.format({"monto":"${:,.2f}"}), use_container_width=True)

    st.subheader("💾 Descargar")
    st.download_button(
        label="📥 Descargar Póliza Excel",
        data=res["excel_bytes"],
        file_name=res["nombre_archivo"],
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )

    if res["marked"]:
        st.subheader("🗙️ Estados de cuenta marcados")
        st.caption("Las filas incluidas en la póliza aparecen resaltadas en el color del banco.")
        for banco, info in res["marked"].items():
            st.download_button(
                label=f"{info['ico']} Descargar {banco} marcado  ({info['n']} movimientos incluidos)",
                data=info["bytes"],
                file_name=f"MARCADO_{banco}_{datetime.now().strftime('%Y-%m')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"dl_marcado_{banco}",
            )

else:
    st.info("⬆️ Sube al menos un archivo de depósito para continuar.")
