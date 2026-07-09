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
st.markdown("""
<style>
[data-testid="stAppViewContainer"] { background: #0f172a; }
[data-testid="stHeader"] { background: transparent; }
h1, h2, h3, .stMarkdown p { color: #e2e8f0; }
.upload-box { background: #1e293b; border: 1px solid #334155;
              border-radius: 10px; padding: 16px; margin-bottom: 12px; }
.stat-card  { background: #1e3a5f; border-radius: 8px; padding: 12px 16px;
              text-align: center; }
.no-clasif  { background: #7f1d1d; border-radius: 8px; padding: 12px 16px; }
</style>
""", unsafe_allow_html=True)

st.title("🏦 Depósitos Bancarios")
st.caption("Genera la póliza contable desde los estados de cuenta de BBVA, Banorte e Inbursa.")

# ─── Constantes contables ──────────────────────────────────────────────────────
CARGOS = {
    "BANORTE": ("102-01-0001-0001", " Banorte"),
    "BBVA":    ("102-01-0001-0003", "BBVA"),
    "INBURSA": ("102-01-0001-0002", "INBURSA"),
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
    """BBVA: todos los depósitos → BBVA transit (col 18)."""
    return 18


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
    # Mañana (antes de 12:00) → col 12 (AMEX)  |  Tarde (≥12:00) → col 17 (Shell/ajuste)
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

    # TPV Banorte (SERVICIOS FELUSA terminales C/D)
    if "07277262C" in d or "07277262D" in d:
        return 16

    # No clasificado
    return None


CLASIFICADORES = {
    "BBVA":    clasificar_bbva,
    "BANORTE": clasificar_banorte,
    "INBURSA": clasificar_inbursa,
}

# Columna de cargo según banco
CARGO_COL = {"BBVA": 9, "BANORTE": 8, "INBURSA": 10}

# ─── Lectura de estados de cuenta ─────────────────────────────────────────────
def leer_banco(file_obj, banco: str) -> tuple[list, list]:
    """
    Lee el estado de cuenta y devuelve (registros_ok, registros_sin_clasif).
    Cada registro_ok = dict con claves: fecha, banco, desc, monto, col_cargo, col_abono
    """
    wb = openpyxl.load_workbook(file_obj, read_only=True, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    wb.close()

    # Detectar fila de encabezado (busca 'Fecha' en col 0)
    inicio = 2
    for i, r in enumerate(rows):
        if r and str(r[0]).strip().lower() == "fecha":
            inicio = i + 1
            break

    ok, sin_clasif = [], []
    fn = CLASIFICADORES[banco]
    col_cargo = CARGO_COL[banco]

    for r in rows[inicio:]:
        if not r or r[0] is None:
            continue
        fecha = r[0]
        if not hasattr(fecha, "day"):
            continue
        desc  = str(r[1] or "").strip()
        monto = r[2]  # columna Depósito
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
            })

    return ok, sin_clasif


# ─── Generación del Excel (respetando formato del template) ───────────────────
def generar_excel(registros: list, plantilla=None) -> bytes:
    """
    Genera el Excel de póliza respetando el formato del template
    DEPOSITOS BANCARIOS FELUSA.xlsx:
    - Fuente: Aptos Narrow 11
    - Colores de encabezados exactos del template
    - freeze_panes en B4
    - Columna DIFERENCIA = TOTAL CARGOS - TOTAL ABONOS
    - Hoja CUENTAS igual al template
    """
    # ── Fuente base del template ──
    FONT_NAME = "Aptos Narrow"
    FONT_SIZE = 11

    # ── Fills exactos del template (RGB hex) ──
    # Theme colors aproximados con los colores reales del Office default:
    F_TIPO    = PatternFill("solid", fgColor="E2EFDA")   # theme:6 tint:0.599 ≈ verde claro
    F_FECHA   = PatternFill("solid", fgColor="FFC000")   # amber exacto del template
    F_REF     = PatternFill("solid", fgColor="7030A0")   # purple exacto
    F_CONC    = PatternFill("solid", fgColor="00B050")   # green exacto
    F_ERROR   = PatternFill("solid", fgColor="843C0C")   # theme:5 tint:-0.249 ≈ naranja oscuro
    F_ADMIN   = PatternFill("solid", fgColor="D0D6DC")   # theme:3 tint:0.749 ≈ azul gris claro
    F_BANCO   = PatternFill("solid", fgColor="BDD7EE")   # celeste para columnas de cargo
    F_ABONO   = PatternFill("solid", fgColor="FFE699")   # amarillo claro para columnas de abono
    F_GRAY2   = PatternFill("solid", fgColor="BFBFBF")   # theme:0 tint:-0.249 ≈ gris (fila 2)
    # Sin fill
    F_NONE    = PatternFill(fill_type=None)

    # ── Fuentes ──
    def fnt(bold=False, color="000000", size=None, italic=False):
        return Font(name=FONT_NAME, size=size or FONT_SIZE,
                    bold=bold, color=color, italic=italic)

    # ── Alineación ──
    A_CTR  = Alignment(horizontal="center", vertical="center", wrap_text=False)
    A_CTR_W = Alignment(horizontal="center", vertical="center", wrap_text=True)
    A_LEFT = Alignment(horizontal="left",   vertical="center", wrap_text=False)
    A_RIGHT= Alignment(horizontal="right",  vertical="center", wrap_text=False)

    # ── Formato numérico ──
    FMT_NUM  = '#,##0.00'
    FMT_DATE = 'DD/MM/YYYY'

    def set_cell(ws, row, col, value=None, font=None, fill=F_NONE,
                 align=A_CTR, num_format=None):
        c = ws.cell(row=row, column=col, value=value)
        if font:    c.font = font
        if fill:    c.fill = fill
        if align:   c.alignment = align
        if num_format: c.number_format = num_format
        return c

    # ══════════════════════════════════════════════════════════════════════════
    # Hoja POLIZA
    # ══════════════════════════════════════════════════════════════════════════
    if plantilla is not None:
        # Cargar la plantilla y limpiar filas de datos (fila 4 en adelante)
        wb = openpyxl.load_workbook(plantilla)
        ws = wb["POLIZA"] if "POLIZA" in wb.sheetnames else wb.active
        max_row = ws.max_row
        for row_idx in range(4, max_row + 1):
            for col_idx in range(1, ws.max_column + 1):
                ws.cell(row=row_idx, column=col_idx).value = None
    else:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "POLIZA"

        # ── FILA 1: índices de columna (0-20) ──
        for pol_idx in range(21):
            set_cell(ws, 1, pol_idx + 1,
                     value=pol_idx,
                     font=fnt(color="000000"),
                     fill=F_NONE, align=A_CTR)

        # ── FILA 2: números de cuenta (solo en cols de banco y tránsito) ──
        for pol_idx, banco in [(8, "BANORTE"), (9, "BBVA"), (10, "INBURSA")]:
            set_cell(ws, 2, pol_idx + 1,
                     value=CARGOS[banco][0],
                     font=fnt(color="000000"),
                     fill=F_GRAY2, align=A_CTR)
        for a in ABONOS:
            set_cell(ws, 2, a["col"] + 1,
                     value=a["cuenta"],
                     font=fnt(color="000000"),
                     fill=F_GRAY2, align=A_CTR)

        # ── FILA 3: etiquetas de columna con colores del template ──
        headers_admin = [
            (1,  "TIPO DE POLIZA", F_TIPO,  fnt(color="000000")),
            (2,  "Fecha",          F_FECHA, fnt(bold=True, color="FFFFFF")),
            (3,  "REFERENCIA",     F_REF,   fnt(bold=True, color="FFFFFF")),
            (4,  "CONCEPTO",       F_CONC,  fnt(bold=True, color="FFFFFF")),
            (5,  "ERROR",          F_ERROR, fnt(bold=True, color="FFFFFF")),
            (6,  "UIDD",           F_ADMIN, fnt(bold=True, color="FFFFFF")),
            (7,  "NUM POLIZA",     F_ADMIN, fnt(bold=True, color="FFFFFF")),
            (8,  "PROCESADO",      F_ADMIN, fnt(bold=True, color="FFFFFF")),
        ]
        for col, lbl, fill, font in headers_admin:
            set_cell(ws, 3, col, value=lbl, font=font, fill=fill, align=A_CTR)

        for pol_idx, banco in [(8, "BANORTE"), (9, "BBVA"), (10, "INBURSA")]:
            set_cell(ws, 3, pol_idx + 1,
                     value=CARGOS[banco][1].strip(),
                     font=fnt(bold=True, color="FFFFFF"),
                     fill=F_BANCO, align=A_CTR)

        set_cell(ws, 3, 12, value="TOTAL CARGOS",
                 font=fnt(color="FF0000"), fill=F_NONE, align=A_CTR)

        for a in ABONOS:
            set_cell(ws, 3, a["col"] + 1,
                     value=a["nombre"],
                     font=fnt(bold=True, color="000000"),
                     fill=F_ABONO, align=A_CTR_W)

        set_cell(ws, 3, 21, value="TOTAL ABONOS",
                 font=fnt(color="FF0000"), fill=F_NONE, align=A_CTR)

        set_cell(ws, 3, 22, value="DIFERENCIA",
                 font=fnt(color="FF0000"), fill=F_NONE, align=A_CTR)

    # ── FILAS DE DATOS (desde fila 4) ──
    # Orden: BBVA → INBURSA → BANORTE, y dentro de cada banco por fecha
    orden_banco = {"BBVA": 0, "INBURSA": 1, "BANORTE": 2}
    registros_sorted = sorted(
        registros,
        key=lambda x: (orden_banco[x["banco"]], x["fecha"])
    )

    # Color base por banco (tono claro) + alternado más oscuro
    F_BBVA_1  = PatternFill("solid", fgColor="DDEEFF")   # azul muy claro
    F_BBVA_2  = PatternFill("solid", fgColor="C5DCF5")   # azul claro alt
    F_INBU_1  = PatternFill("solid", fgColor="FFF0DC")   # naranja muy claro
    F_INBU_2  = PatternFill("solid", fgColor="FFE0B5")   # naranja claro alt
    F_BNT_1   = PatternFill("solid", fgColor="FFE4E4")   # rojo muy claro
    F_BNT_2   = PatternFill("solid", fgColor="FFCCCC")   # rojo claro alt

    FILLS_BANCO = {
        "BBVA":    (F_BBVA_1, F_BBVA_2),
        "INBURSA": (F_INBU_1, F_INBU_2),
        "BANORTE": (F_BNT_1,  F_BNT_2),
    }

    for fila_num, r in enumerate(registros_sorted, start=4):
        f1, f2 = FILLS_BANCO[r["banco"]]
        fill_row = f1 if fila_num % 2 == 0 else f2
        fn_dat   = fnt(color="000000")

        def dat(col, val, num_fmt=None, align=A_CTR):
            c = set_cell(ws, fila_num, col, value=val,
                         font=fn_dat, fill=fill_row, align=align)
            if num_fmt: c.number_format = num_fmt
            return c

        monto = r["monto"]

        # A: "I"
        dat(1,  "I", align=A_CTR)
        # B: Fecha
        dat(2, r["fecha"], num_fmt=FMT_DATE)
        # C: Referencia
        dat(3,  r["ref"], align=A_LEFT)
        # D: Concepto
        dat(4,  r["ref"], align=A_LEFT)
        # E, F, G, H: vacías
        for col in [5, 6, 7, 8]:
            dat(col, None)

        # Cargo en columna del banco correspondiente
        dat(r["col_cargo"] + 1, monto, num_fmt=FMT_NUM, align=A_RIGHT)

        # TOTAL CARGOS (col 12)
        dat(12, monto, num_fmt=FMT_NUM, align=A_RIGHT)

        # Abono en columna de tránsito correspondiente
        dat(r["col_abono"] + 1, monto, num_fmt=FMT_NUM, align=A_RIGHT)

        # TOTAL ABONOS (col 21)
        dat(21, monto, num_fmt=FMT_NUM, align=A_RIGHT)

        # DIFERENCIA (col 22): fórmula =L{n}-U{n}
        col_L = get_column_letter(12)
        col_U = get_column_letter(21)
        c_diff = ws.cell(row=fila_num, column=22,
                         value=f"={col_L}{fila_num}-{col_U}{fila_num}")
        c_diff.font = fnt(color="000000")
        c_diff.fill = fill_row
        c_diff.alignment = A_RIGHT
        c_diff.number_format = FMT_NUM

    # ── Anchos de columna (siguiendo template) ──
    anchos = {
        1: 14.4,   # A TIPO
        2: 13.0,   # B Fecha
        3: 36.9,   # C REFERENCIA
        4: 26.3,   # D CONCEPTO
        5: 6.7,    # E ERROR
        6: 5.3,    # F UIDD
        7: 11.3,   # G NUM POLIZA
        8: 11.6,   # H PROCESADO
        9: 16.0,   # I BANORTE
        10: 16.0,  # J BBVA
        11: 16.0,  # K INBURSA
        12: 16.3,  # L TOTAL CARGOS (= template I)
        13: 32.0,  # M AMEX
        14: 26.0,  # N EFECTIVALE
        15: 30.0,  # O EDENRED
        16: 18.0,  # P FONDO
        17: 28.0,  # Q BANORTE TRANSIT
        18: 32.0,  # R SHELL
        19: 12.0,  # S BBVA TRANSIT
        20: 28.0,  # T INBURSA TRANSIT
        21: 14.1,  # U TOTAL ABONOS (= template J)
        22: 12.6,  # V DIFERENCIA (= template K)
    }
    for col_num, w in anchos.items():
        ws.column_dimensions[get_column_letter(col_num)].width = w

    # Altura fila 3
    ws.row_dimensions[3].height = 15.0

    # freeze_panes: igual al template (B4 → congela fila 3 y col A)
    ws.freeze_panes = "B4"

    # ══════════════════════════════════════════════════════════════════════════
    # Hoja CUENTAS (solo si no se usa plantilla; la plantilla ya la tiene)
    # ══════════════════════════════════════════════════════════════════════════
    if plantilla is None:
        wc = wb.create_sheet("CUENTAS")

        fn_plain = fnt(color="000000")
        A_CTR_C  = Alignment(horizontal="center")

        set_cell(wc, 1, 1, "CARGOS", font=fn_plain, fill=F_NONE, align=A_CTR_C)
        set_cell(wc, 1, 4, "ABONOS", font=fn_plain, fill=F_NONE, align=A_CTR_C)
        set_cell(wc, 2, 1, "N° Cuenta",           font=fn_plain, fill=F_NONE)
        set_cell(wc, 2, 2, "Banco",               font=fn_plain, fill=F_NONE)
        set_cell(wc, 2, 4, "N° Cuenta",           font=fn_plain, fill=F_NONE)
        set_cell(wc, 2, 5, "Nombre del Deposito", font=fn_plain, fill=F_NONE)
        for i, (banco, (cuenta, nombre)) in enumerate(CARGOS.items(), start=3):
            set_cell(wc, i, 1, cuenta,         font=fn_plain, fill=F_NONE)
            set_cell(wc, i, 2, nombre.strip(), font=fn_plain, fill=F_NONE)
        for i, a in enumerate(ABONOS, start=3):
            set_cell(wc, i, 4, a["cuenta"], font=fn_plain, fill=F_NONE)
            set_cell(wc, i, 5, a["nombre"], font=fn_plain, fill=F_NONE)
        for col_num, w in {1: 16.3, 2: 10.0, 4: 16.3, 5: 42.9}.items():
            wc.column_dimensions[get_column_letter(col_num)].width = w

    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ─── UI ───────────────────────────────────────────────────────────────────────
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

if st.button("⚙️ Generar Póliza", type="primary",
             disabled=not any([file_bbva, file_banorte, file_inbursa])):

    todos_ok, todos_nc = [], []
    progress = st.progress(0)

    archivos = [
        (file_bbva,    "BBVA",    "🟦"),
        (file_banorte, "BANORTE", "🟥"),
        (file_inbursa, "INBURSA", "🟧"),
    ]

    for i, (f, banco, ico) in enumerate(archivos):
        if f:
            with st.spinner(f"{ico} Leyendo {banco}..."):
                try:
                    ok, nc = leer_banco(f, banco)
                    todos_ok.extend(ok)
                    todos_nc.extend(nc)
                    st.success(f"{ico} {banco}: **{len(ok)}** depósitos clasificados"
                               + (f", {len(nc)} sin clasificar" if nc else ""))
                except Exception as e:
                    st.error(f"Error leyendo {banco}: {e}")
        progress.progress((i + 1) / 3)

    if todos_ok:
        import pandas as pd

        # ── Resumen ──
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

        # ── Vista previa ──
        st.subheader("📋 Vista previa")
        abono_nombre = {a["col"]: a["nombre"][:30] for a in ABONOS}
        cargo_nombre = {9: "BANORTE (102-01)", 10: "BBVA (102-03)", 11: "INBURSA (102-02)"}

        preview_rows = []
        orden = {"BBVA": 0, "BANORTE": 1, "INBURSA": 2}
        for r in sorted(todos_ok, key=lambda x: (x["fecha"], orden[x["banco"]])):
            preview_rows.append({
                "Fecha":       r["fecha"].strftime("%d/%m/%Y") if hasattr(r["fecha"], "strftime") else str(r["fecha"]),
                "Banco":       r["banco"],
                "Cargo":       cargo_nombre.get(r["col_cargo"] + 1, str(r["col_cargo"])),
                "Abono":       abono_nombre.get(r["col_abono"], str(r["col_abono"])),
                "Monto":       r["monto"],
                "Descripción": r["desc"][:55],
            })

        st.dataframe(
            pd.DataFrame(preview_rows).style.format({"Monto": "${:,.2f}"}),
            use_container_width=True,
            height=400,
        )

        # ── No clasificados ──
        if todos_nc:
            st.subheader(f"⚠️ Sin clasificar ({len(todos_nc)} movimientos — no incluidos)")
            df_nc = pd.DataFrame(todos_nc)
            df_nc["fecha"] = df_nc["fecha"].apply(
                lambda x: x.strftime("%d/%m/%Y") if hasattr(x, "strftime") else str(x))
            st.dataframe(df_nc.style.format({"monto": "${:,.2f}"}),
                         use_container_width=True)

        # ── Descargar ──
        st.subheader("💾 Descargar")
        excel_bytes = generar_excel(todos_ok, plantilla=file_plantilla)
        nombre_archivo = f"DEPOSITOS BANCARIOS {datetime.now().strftime('%Y-%m')}.xlsx"
        st.download_button(
            label="📥 Descargar Póliza Excel",
            data=excel_bytes,
            file_name=nombre_archivo,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )
    else:
        st.warning("No se encontraron depósitos clasificables.")
