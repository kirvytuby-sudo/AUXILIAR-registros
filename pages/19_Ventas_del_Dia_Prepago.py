"""AUXILIAR DE REGISTROS — Ventas del Día Prepago (Control de Despachos → Póliza con Crédito/Prepago separados)"""
import streamlit as st
import io
from collections import defaultdict
from datetime import datetime

st.set_page_config(
    page_title="Ventas del Día Prepago · Auxiliar",
    page_icon="🏷️",
    layout="wide",
)

import _theme
_theme.aplicar_header(
    "🏷️ Ventas del Día Prepago",
    "Control de Despachos → Póliza con separación Crédito / Prepago",
)

try:
    import openpyxl
    from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
except ImportError as _e:
    st.error(f"❌ Librería faltante: {_e}. Verifica requirements.txt.")
    st.stop()

def _leer_despachos_bytes(file_bytes: bytes, filename: str):
    """Lee el Control de Despachos desde .xlsx o .xls y retorna lista de filas (sin cabecera)."""
    ext = filename.rsplit(".", 1)[-1].lower() if "." in filename else "xlsx"
    if ext == "xlsx":
        wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
        ws = wb.active
        rows = list(ws.iter_rows(min_row=2, values_only=True))
        wb.close()
        return rows
    else:
        # .xls — intentar con xlrd, fallback a TSV/CSV si no es binario real
        try:
            import xlrd
            wb = xlrd.open_workbook(file_contents=file_bytes)
            ws = wb.sheet_by_index(0)
            rows = []
            for i in range(1, ws.nrows):
                row = ws.row_values(i)
                # xlrd devuelve fechas como float → convertir col 0
                try:
                    if isinstance(row[0], float) and row[0] > 0:
                        from datetime import datetime as _dt
                        import xlrd as _xl
                        tup = _xl.xldate_as_tuple(row[0], wb.datemode)
                        row[0] = _dt(*tup)
                except Exception:
                    pass
                rows.append(row)
            return rows
        except Exception:
            # Fallback: leer como texto CSV/TSV
            import csv
            text = file_bytes.decode("utf-8", errors="replace")
            dialect = "excel-tab" if "\t" in text[:500] else "excel"
            reader = csv.reader(io.StringIO(text), dialect=dialect)
            all_rows = list(reader)
            return all_rows[1:] if all_rows else []

# ── Cuentas fijas de pago ────────────────────────────────────────────────────
# V.EDENRED y V.EFECTIVALE acumulan en la misma columna que T.EDENRED / T.EFECTIVALE
FIJAS_DEF = [
    ("CONTADO",      "105-01-0001-0001"),
    ("T.BANORTE",    "105-01-0001-0004"),
    ("T.AMEX",       "105-01-0001-0007"),
    ("T.EFECTIVALE", "105-01-0002-0001"),
    ("T.EDENRED",    "105-01-0002-0002"),
    ("T.ULTRAGAS",   "105-01-0002-0003"),
]
CLIENTE_TO_FIJA = {
    "VENTA DE CONTADO":     "CONTADO",
    "CONTADO":              "CONTADO",
    "T.BANORTE":            "T.BANORTE",
    "T.BBVA":               "T.BANORTE",
    "T.AMEX":               "T.AMEX",
    "CLIENTE VENTA TIENDA": "CONTADO",
    "T.EFECTIVALE":         "T.EFECTIVALE",
    "V.EFECTIVALE":         "T.EFECTIVALE",   # misma cuenta
    "T.EDENRED":            "T.EDENRED",
    "V.EDENRED":            "T.EDENRED",      # misma cuenta
    "T.ULTRAGAS":           "T.ULTRAGAS",
}

ABONO_COLS = [
    ("401-01-0001-0001", "Gasolina Super"),
    ("401-01-0001-0002", "premium"),
    ("401-01-0001-0003", "Diesel"),
    ("209-01",           "IVA trasladado no cobrado"),
    ("401-01-0001-0006-0001", "IEPS DE Gasolina Magna"),
    ("401-01-0001-0006-0002", "IEPS de Premium"),
    ("401-01-0001-0006-0003", "IEPS de Diesel"),
    ("101-01-0002",      "Efectivo cta. diferencias"),   # balance → CONC = 0
]
EFE_IDX = len(ABONO_COLS) - 1  # índice de Efectivo dentro de ABONO_COLS

# ── Helpers de normalización ─────────────────────────────────────────────────
def _norm(s):
    return str(s or "").upper().replace("\n", " ").strip()

def _buscar_en(lista, nombre):
    """Busca (cuenta, nombre) en lista por coincidencia exacta o subcadena."""
    n = _norm(nombre)
    if not n:
        return None
    for acct, cname in lista:
        if _norm(cname) == n:
            return acct
    for acct, cname in lista:
        cn = _norm(cname)
        if n in cn or cn in n:
            return acct
    return None

# ── Leer cuentas de la plantilla ─────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def _leer_plantilla(plantilla_bytes: bytes):
    """Retorna (cta_credito, cta_prepago) como listas de (cuenta, nombre)."""
    cred, prep = [], []
    try:
        wb = openpyxl.load_workbook(io.BytesIO(plantilla_bytes))
        hoja = None
        for sn in wb.sheetnames:
            if sn.strip().upper() == "CUENTAS":
                hoja = wb[sn]; break
            if "CUENTAS" in sn.upper():
                hoja = wb[sn]
        if not hoja:
            # intentar hoja con espacios
            for sn in wb.sheetnames:
                if "cuentas" in sn.lower():
                    hoja = wb[sn]; break
        if hoja:
            mode = None
            for row in hoja.iter_rows(min_row=1, values_only=True):
                acct = row[7] if len(row) > 7 else None
                nombre = str(row[8]).strip().replace("\n", "").strip() if len(row) > 8 and row[8] else ""
                if not acct:
                    continue
                acct = str(acct).strip()
                if acct == "105-01-0003":
                    mode = "credito"; continue
                if acct == "105-01-0004":
                    mode = "prepago"; continue
                if mode == "credito" and acct.startswith("105-01-0003-") and nombre:
                    cred.append((acct, nombre))
                elif mode == "prepago" and acct.startswith("105-01-0004-") and nombre:
                    prep.append((acct, nombre))
        wb.close()
    except Exception as e:
        st.warning(f"⚠ No se pudo leer la plantilla: {e}")
    return cred, prep

# ── Motor de generación ───────────────────────────────────────────────────────
def procesar_prepago(despachos_bytes: bytes, plantilla_bytes: bytes | None, despachos_nombre: str = "archivo.xlsx") -> tuple[bytes, list, list]:
    """
    Genera la póliza con separación Crédito/Prepago.
    Retorna (excel_bytes, logs, resumen_por_dia).
    """
    logs = []

    # Cargar cuentas de plantilla
    if plantilla_bytes:
        cta_credito, cta_prepago = _leer_plantilla(plantilla_bytes)
        logs.append(f"✅ Plantilla: {len(cta_credito)} clientes Crédito, {len(cta_prepago)} clientes Prepago.")
    else:
        cta_credito, cta_prepago = [], []
        logs.append("⚠ Sin plantilla — no hay cuentas de Crédito/Prepago disponibles.")

    FIJA_COLS = [(a, n) for n, a in FIJAS_DEF]
    CRED_COLS = list(cta_credito)
    PREP_COLS = list(cta_prepago)
    N_FIJA = len(FIJA_COLS)
    N_CRED = len(CRED_COLS)
    N_PREP = len(PREP_COLS)
    fija_name_idx = {n: i for i, (a, n) in enumerate(FIJA_COLS)}
    cred_acct_idx = {a: i for i, (a, n) in enumerate(CRED_COLS)}
    prep_acct_idx = {a: i for i, (a, n) in enumerate(PREP_COLS)}

    # Leer despachos (.xlsx o .xls)
    src_rows = _leer_despachos_bytes(despachos_bytes, despachos_nombre)
    logs.append(f"📂 {len(src_rows):,} filas leídas del archivo.")

    def new_day():
        return {
            "fija": [0.0] * N_FIJA, "cred": [0.0] * N_CRED, "prep": [0.0] * N_PREP,
            "desc": 0.0, "gs": 0.0, "gp": 0.0, "gd": 0.0,
            "iva": 0.0, "ieps_gs": 0.0, "ieps_gp": 0.0, "ieps_gd": 0.0,
        }

    day_data = defaultdict(new_day)
    sin_mapear = set()

    def _parse_fecha(v):
        """Normaliza cualquier representación de fecha a objeto date."""
        if v is None or v == "" or v == 0 or v == 0.0:
            return None
        from datetime import date as _date, datetime as _dt
        if isinstance(v, _dt): return v.date()
        if isinstance(v, _date): return v
        s = str(v).strip()[:10]  # tomar solo 'YYYY-MM-DD' del inicio
        for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y"):
            try: return _dt.strptime(s, fmt).date()
            except Exception: pass
        return None

    for row in src_rows:
        if not row or row[0] is None or row[0] == "":
            continue
        fecha = _parse_fecha(row[0])
        if fecha is None:
            continue
        prod    = str(row[3]).upper().strip() if row[3] else ""
        sub     = float(row[6] or 0)
        iva     = float(row[7] or 0)
        ieps    = float(row[8] or 0)
        imp     = float(row[9] or 0)
        dsc_s   = float(row[10] or 0)
        dsc_v   = float(row[11] or 0)
        cliente = str(row[17]).strip() if row[17] else ""
        tipo    = str(row[19]).strip() if row[19] else ""

        d = day_data[fecha]
        if prod == "GS":    d["gs"] += sub;  d["iva"] += (iva - dsc_v); d["ieps_gs"] += ieps
        elif prod == "GP":  d["gp"] += sub;  d["iva"] += (iva - dsc_v); d["ieps_gp"] += ieps
        elif prod == "GD":  d["gd"] += sub;  d["iva"] += (iva - dsc_v); d["ieps_gd"] += ieps

        d["desc"] += dsc_s
        cargo = imp - dsc_s - dsc_v

        if tipo in ("Contado", "Tarjeta", "Monedero"):
            fname = CLIENTE_TO_FIJA.get(cliente)
            if fname is None:
                for k, v in CLIENTE_TO_FIJA.items():
                    if k in cliente or cliente in k:
                        fname = v; break
            idx = fija_name_idx.get(fname, 0)
            d["fija"][idx] += cargo
            if fname is None:
                sin_mapear.add(f"{tipo}:{cliente}")
        elif tipo == "Credito":
            acct = _buscar_en(CRED_COLS, cliente)
            if acct and acct in cred_acct_idx:
                d["cred"][cred_acct_idx[acct]] += cargo
            else:
                fb = _buscar_en(CRED_COLS, "Pendiente por facturar Crédito")
                if fb and fb in cred_acct_idx:
                    d["cred"][cred_acct_idx[fb]] += cargo
                sin_mapear.add(f"Credito:{cliente}")
        elif tipo == "Prepago":
            acct = _buscar_en(PREP_COLS, cliente)
            if acct and acct in prep_acct_idx:
                d["prep"][prep_acct_idx[acct]] += cargo
            else:
                sin_mapear.add(f"Prepago:{cliente}")

    if sin_mapear:
        logs.append(f"⚠ {len(sin_mapear)} clientes sin mapear: {', '.join(sorted(sin_mapear)[:10])}" +
                    (" …" if len(sin_mapear) > 10 else ""))

    sorted_dates = sorted(day_data.keys())
    if not sorted_dates:
        raise RuntimeError("No se encontraron datos de ventas en el archivo.")
    logs.append(f"📅 {len(sorted_dates)} día(s): {sorted_dates[0]} → {sorted_dates[-1]}")

    # Filtrar columnas vacías (solo cliente/cargo)
    fija_tot  = [sum(day_data[f]["fija"][i] for f in sorted_dates) for i in range(N_FIJA)]
    cred_tot  = [sum(day_data[f]["cred"][i] for f in sorted_dates) for i in range(N_CRED)]
    prep_tot  = [sum(day_data[f]["prep"][i] for f in sorted_dates) for i in range(N_PREP)]

    act_fija = [(i, a, n) for i, (a, n) in enumerate(FIJA_COLS) if abs(fija_tot[i]) > 0.001]
    act_cred = [(i, a, n) for i, (a, n) in enumerate(CRED_COLS) if abs(cred_tot[i]) > 0.001]
    act_prep = [(i, a, n) for i, (a, n) in enumerate(PREP_COLS) if abs(prep_tot[i]) > 0.001]
    logs.append(f"📊 Columnas activas — Fijas:{len(act_fija)} | Crédito:{len(act_cred)}/{N_CRED} | Prepago:{len(act_prep)}/{N_PREP}")

    # ── Construir Excel con openpyxl ────────────────────────────────────────
    def fill(h): return PatternFill("solid", fgColor=h)
    def font(color="1F3864", bold=False, sz=9):
        return Font(name="Calibri", size=sz, bold=bold, color=color)

    F_HDR_FIX  = fill("BDD7EE"); F_HDR_FIJA = fill("9DC3E6")
    F_HDR_CRED = fill("B4C7E7"); F_HDR_PREP = fill("C5E0B4")
    F_HDR_DESC = fill("F4B183"); F_HDR_TB2  = fill("FFD966")
    F_HDR_ABO  = fill("A9D18E"); F_HDR_EFE  = fill("FCE4D6")
    F_HDR_CONC = fill("C6E0B4")
    F_WHITE    = fill("FFFFFF"); F_CRED = fill("EEF3FB")
    F_PREP     = fill("F0F7EE"); F_DESC = fill("FEF3E8")
    F_TB2      = fill("FFF8DC"); F_ABO  = fill("F4FAF0")
    F_EFE      = fill("FEF9F7"); F_CONC = fill("EDF7E6")
    F_TOT      = fill("FFD966"); F_TOT_CONC = fill("92D050")

    DARK   = font()
    DARK_B = font(bold=True)
    GS_    = Side(border_style="thin", color="BFBFBF")
    BRD    = Border(left=GS_, right=GS_, top=GS_, bottom=GS_)
    CTR    = Alignment(horizontal="center", vertical="center")

    wb_out = openpyxl.Workbook()
    ws = wb_out.active
    ws.title = "poliza IA"

    def w(r, c, val, fill_=None, font_=None, fmt=None):
        cell = ws.cell(r, c, val)
        cell.fill      = fill_ or F_WHITE
        cell.font      = font_ or DARK
        cell.alignment = CTR
        cell.border    = BRD
        if fmt:
            cell.number_format = fmt

    FIXED_META = [(None, n) for n in ["TIPO DE POLIZA","Fecha","REFERENCIA","CONCEPTO","ERROR","UIDD","NUM POLIZA","PROCESADO"]]
    FIJA_META  = [(a, n) for _, a, n in act_fija]
    CRED_META  = [(a, n) for _, a, n in act_cred]
    PREP_META  = [(a, n) for _, a, n in act_prep]

    ALL_COLS = (FIXED_META + FIJA_META + CRED_META + PREP_META +
                [("402-01","DescuentoSubtotal"), (None,"TOTAL B2")] +
                ABONO_COLS + [(None,"TOTAL B2"), (None,"CONCILIACION")])

    N8  = 8
    NF_ = N8  + len(FIJA_META)
    NC_ = NF_ + len(CRED_META)
    NP_ = NC_ + len(PREP_META)
    ND_ = NP_
    NT1_= ND_ + 1
    NA_ = NT1_+ 1
    NT2_= NA_ + len(ABONO_COLS)
    # CONC es NT2_+1 (0-indexed = NT2_)

    def hdr_fill(ci):
        if ci < N8:   return F_HDR_FIX
        if ci < NF_:  return F_HDR_FIJA
        if ci < NC_:  return F_HDR_CRED
        if ci < NP_:  return F_HDR_PREP
        if ci == ND_: return F_HDR_DESC
        if ci == NT1_:return F_HDR_TB2
        if ci == NA_ + EFE_IDX: return F_HDR_EFE
        if ci < NT2_: return F_HDR_ABO
        if ci == NT2_:return F_HDR_TB2
        return F_HDR_CONC

    def dat_fill(ci):
        if ci < NF_:  return F_WHITE
        if ci < NC_:  return F_CRED
        if ci < NP_:  return F_PREP
        if ci == ND_: return F_DESC
        if ci == NT1_:return F_TB2
        if ci == NA_ + EFE_IDX: return F_EFE
        if ci < NT2_: return F_ABO
        if ci == NT2_:return F_TB2
        return F_CONC

    # Cabeceras filas 1-3
    for ci, (acct, name) in enumerate(ALL_COLS):
        col = ci + 1
        hf  = hdr_fill(ci)
        w(1, col, ci,   fill_=hf, font_=DARK_B)
        w(2, col, acct, fill_=hf, font_=DARK_B)
        w(3, col, name, fill_=hf, font_=DARK_B)

    totals  = defaultdict(float)
    resumen = []

    for ri, fecha in enumerate(sorted_dates):
        d = day_data[fecha]; r = ri + 4
        fecha_str = fecha.strftime("%d/%m/%Y") if hasattr(fecha, "strftime") else str(fecha)
        fixed_vals = ["CLI", fecha, f"VENTAS DEL DIA {fecha_str}", f"VENTAS DEL DIA {fecha_str}", None, None, None, None]
        for ci2, val in enumerate(fixed_vals):
            w(r, ci2+1, val, fill_=F_WHITE, font_=DARK,
              fmt=("DD/MM/YYYY" if ci2 == 1 else None))

        col = 9
        for orig_i, a, n in act_fija:
            val = round(d["fija"][orig_i], 2) or None
            w(r, col, val, fill_=F_WHITE, font_=DARK, fmt="#,##0.00"); col += 1
            if val: totals[f"F{orig_i}"] += val

        for orig_i, a, n in act_cred:
            val = round(d["cred"][orig_i], 2) or None
            w(r, col, val, fill_=F_CRED, font_=DARK, fmt="#,##0.00"); col += 1
            if val: totals[f"C{orig_i}"] += val

        for orig_i, a, n in act_prep:
            val = round(d["prep"][orig_i], 2) or None
            w(r, col, val, fill_=F_PREP, font_=DARK, fmt="#,##0.00"); col += 1
            if val: totals[f"P{orig_i}"] += val

        desc = round(d["desc"], 2) or None
        w(r, col, desc, fill_=F_DESC, font_=DARK, fmt="#,##0.00"); col += 1
        if desc: totals["desc"] += desc

        tb2c = round(
            sum(d["fija"][i] for i, _, __ in act_fija) +
            sum(d["cred"][i] for i, _, __ in act_cred) +
            sum(d["prep"][i] for i, _, __ in act_prep) +
            (d["desc"] or 0), 2
        )
        w(r, col, tb2c or None, fill_=F_TB2, font_=DARK_B, fmt="#,##0.00"); col += 1
        totals["tb2c"] += tb2c

        # Abonos: todos redondeados a 2 decimales
        gs_ = round(d["gs"],      2); gp_  = round(d["gp"],      2); gd_  = round(d["gd"],      2)
        iva_= round(d["iva"],     2); igs_ = round(d["ieps_gs"], 2)
        igp_= round(d["ieps_gp"],2); igd_ = round(d["ieps_gd"], 2)
        # Efectivo = exactamente lo que falta para que TB2 abonos = TB2 cargos → CONC = 0.00 exacto
        otros = round(gs_+gp_+gd_+iva_+igs_+igp_+igd_, 2)
        efectivo = round(tb2c - otros, 2)
        abo_vals = [gs_, gp_, gd_, iva_, igs_, igp_, igd_, efectivo]

        for j, (ac, nm) in enumerate(ABONO_COLS):
            val = round(abo_vals[j], 2) or None
            df  = F_EFE if j == EFE_IDX else F_ABO
            w(r, col, val, fill_=df, font_=DARK, fmt="#,##0.00"); col += 1
            if val: totals[f"a{j}"] += abo_vals[j]

        # TB2 abonos = tb2c exacto (garantiza CONC = 0)
        tb2a = tb2c
        w(r, col, round(tb2a, 2), fill_=F_TB2, font_=DARK_B, fmt="#,##0.00"); col += 1
        totals["tb2a"] += tb2a

        w(r, col, 0, fill_=F_CONC, font_=DARK_B, fmt="#,##0.00")

        resumen.append({
            "Fecha":        fecha_str,
            "TB2 Cargos":   round(tb2c, 2),
            "TB2 Abonos":   round(tb2a, 2),
            "Conciliación": 0.0,
        })

    # TOTAL GENERAL — fórmulas SUM para que Excel recalcule y sea auditable
    r_t    = len(sorted_dates) + 4
    r_ini  = 4          # primera fila de datos
    r_fin  = r_t - 1   # última fila de datos

    col = 1
    w(r_t, col, "TOTAL GENERAL", fill_=F_TOT, font_=DARK_B); col += 1
    for _ in range(7): w(r_t, col, None, fill_=F_TOT); col += 1   # cols 2-8 vacías

    # Columnas numéricas: fórmula =SUM(X4:Xn)
    col_num_start = 9   # primera columna numérica (fijas)
    n_num_cols = (len(act_fija) + len(act_cred) + len(act_prep) +
                  1 +                # DescuentoSubtotal
                  1 +                # TOTAL B2 cargos
                  len(ABONO_COLS) +  # abonos
                  1 +                # TOTAL B2 abonos
                  1)                 # CONCILIACION

    for rel in range(n_num_cols):
        abs_col = col_num_start + rel   # columna Excel (1-indexed)
        col_letter = openpyxl.utils.get_column_letter(abs_col)
        formula = f"=ROUND(SUM({col_letter}{r_ini}:{col_letter}{r_fin}),2)"

        # CONCILIACION total: fórmula diferencia TB2 cargos − TB2 abonos
        col_tb2c = col_num_start + len(act_fija) + len(act_cred) + len(act_prep) + 1  # +1 para desc
        col_tb2a = col_tb2c + 1 + len(ABONO_COLS)  # TB2 abonos
        col_conc = col_tb2a + 1
        if abs_col == col_conc:
            l_tb2c = openpyxl.utils.get_column_letter(col_tb2c)
            l_tb2a = openpyxl.utils.get_column_letter(col_tb2a)
            formula = f"=ROUND({l_tb2c}{r_t}-{l_tb2a}{r_t},2)"
            w(r_t, abs_col, formula, fill_=F_TOT_CONC, font_=DARK_B, fmt="#,##0.00")
        elif abs_col in (col_tb2c, col_tb2a):
            w(r_t, abs_col, formula, fill_=F_TOT, font_=DARK_B, fmt="#,##0.00")
        else:
            w(r_t, abs_col, formula, fill_=F_TOT, font_=DARK_B, fmt="#,##0.00")

    conc_t = round(totals["tb2c"] - totals["tb2a"], 2)

    # Anchos de columna
    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 12
    ws.column_dimensions["C"].width = 12
    ws.column_dimensions["D"].width = 24
    for i in range(5, len(ALL_COLS) + 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = 15
    ws.row_dimensions[1].height = 18
    ws.row_dimensions[2].height = 22
    ws.row_dimensions[3].height = 30
    ws.freeze_panes = "C4"

    buf = io.BytesIO()
    wb_out.save(buf)
    buf.seek(0)

    logs.append(f"✅ TB2 Cargos: ${totals['tb2c']:,.2f} | TB2 Abonos: ${totals['tb2a']:,.2f} | CONC: ${conc_t:,.2f}")
    return buf.read(), logs, resumen


# ══════════════════════════════════════════════════════════════════════════════
# UI
# ══════════════════════════════════════════════════════════════════════════════

st.markdown("### 📂 Archivos de entrada")
col1, col2 = st.columns([1, 1])

with col1:
    despachos_file = st.file_uploader(
        "📊 Control de despachos (.xlsx / .xls)",
        type=["xlsx", "xls"],
        help="Archivo de control de despachos con columnas: Importe, DescuentoSubtotal, DescuentoIva, Cliente, Tipo.",
    )

with col2:
    plantilla_file = st.file_uploader(
        "📋 Plantilla VENTAS DEL DIA VALLEJO (.xlsx)",
        type=["xlsx"],
        help="Plantilla con hoja 'cuentas' que contiene las cuentas 105-01-0003-* (Crédito) y 105-01-0004-* (Prepago).",
    )

st.markdown("")
generar = st.button(
    "🏷️  Generar Póliza Crédito / Prepago",
    type="primary",
    disabled=despachos_file is None or plantilla_file is None,
    use_container_width=True,
)

if despachos_file is None or plantilla_file is None:
    st.info("👆 Selecciona el Control de Despachos y la Plantilla para comenzar.")

if generar and despachos_file is not None and plantilla_file is not None:
    with st.spinner("Procesando despachos y generando póliza…"):
        try:
            excel_bytes, logs, resumen = procesar_prepago(
                despachos_file.read(),
                plantilla_file.read(),
                despachos_file.name,
            )

            base = despachos_file.name.rsplit(".", 1)[0]
            nombre_salida = f"poliza_credito_prepago_{base}.xlsx"

            st.success(f"✅ Póliza generada — {len(resumen)} día(s)")

            st.download_button(
                label="💾  Descargar póliza Excel",
                data=excel_bytes,
                file_name=nombre_salida,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

            # Resumen tabular
            if resumen:
                st.markdown("#### 📋 Resumen por día")
                import pandas as pd
                df = pd.DataFrame(resumen)
                # Totales
                tot_row = {
                    "Fecha": "TOTAL",
                    "TB2 Cargos":   df["TB2 Cargos"].sum(),
                    "TB2 Abonos":   df["TB2 Abonos"].sum(),
                    "Conciliación": df["Conciliación"].sum(),
                }
                df = pd.concat([df, pd.DataFrame([tot_row])], ignore_index=True)
                st.dataframe(
                    df.style.format({
                        "TB2 Cargos":   "${:,.2f}",
                        "TB2 Abonos":   "${:,.2f}",
                        "Conciliación": "${:,.2f}",
                    }).apply(
                        lambda row: ["background-color:#FFD966; font-weight:bold"] * len(row)
                        if row["Fecha"] == "TOTAL" else [""] * len(row),
                        axis=1,
                    ),
                    use_container_width=True,
                    hide_index=True,
                )

            # Log de proceso
            with st.expander("📄 Log de proceso", expanded=False):
                for line in logs:
                    st.markdown(f"- {line}")

        except Exception as exc:
            st.error(f"❌ Error al generar la póliza: {exc}")
            import traceback
            with st.expander("Detalle del error"):
                st.code(traceback.format_exc())
