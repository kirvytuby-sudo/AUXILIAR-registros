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
    """
    Lee el Control de Despachos y retorna (cabecera: list[str], filas: list[tuple]).
    Soporta .xlsx, .xls binario (xlrd), .xls-en-realidad-xlsx (openpyxl) y TSV/CSV.
    """
    ext = filename.rsplit(".", 1)[-1].lower() if "." in filename else "xlsx"

    # ── xlsx ──────────────────────────────────────────────────────────────────
    if ext == "xlsx":
        wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
        ws = wb.active
        all_rows = list(ws.iter_rows(min_row=1, values_only=True))
        wb.close()
        hdr  = [str(c).strip() if c is not None else "" for c in (all_rows[0] if all_rows else [])]
        data = all_rows[1:] if len(all_rows) > 1 else []
        return hdr, data

    # ── .xls — 4 intentos en cascada ─────────────────────────────────────────
    # 1) xlrd (BIFF binario real)
    try:
        import xlrd
        wb = xlrd.open_workbook(file_contents=file_bytes)
        ws = wb.sheet_by_index(0)
        hdr  = [str(ws.cell_value(0, c)).strip() for c in range(ws.ncols)]
        data = []
        for i in range(1, ws.nrows):
            row = list(ws.row_values(i))
            try:
                if isinstance(row[0], float) and row[0] > 0:
                    from datetime import datetime as _dt
                    import xlrd as _xl
                    tup = _xl.xldate_as_tuple(row[0], wb.datemode)
                    row[0] = _dt(*tup)
            except Exception:
                pass
            data.append(tuple(row))
        return hdr, data
    except Exception:
        pass

    # 2) openpyxl (algunos .xls son xlsx con extensión incorrecta)
    try:
        wb2 = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
        ws2 = wb2.active
        all_rows = list(ws2.iter_rows(min_row=1, values_only=True))
        wb2.close()
        hdr  = [str(c).strip() if c is not None else "" for c in (all_rows[0] if all_rows else [])]
        return hdr, all_rows[1:] if len(all_rows) > 1 else []
    except Exception:
        pass

    # 3) pandas
    try:
        import pandas as pd
        df = pd.read_excel(io.BytesIO(file_bytes), header=0, engine="xlrd")
        hdr  = [str(c).strip() for c in df.columns]
        data = [tuple(r) for r in df.itertuples(index=False)]
        return hdr, data
    except Exception:
        pass

    # 4) TSV/CSV (archivos de texto renombrados como .xls)
    import csv
    text = file_bytes.decode("utf-8", errors="replace")
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    dialect = "excel-tab" if "\t" in text[:500] else "excel"
    reader   = csv.reader(io.StringIO(text), dialect=dialect)
    all_rows = list(reader)
    if not all_rows:
        return [], []
    hdr  = [str(c).strip() for c in all_rows[0]]
    return hdr, all_rows[1:]

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
    ("101-01-0001",      "Efectivo cta. diferencias"),   # balance → CONC = 0
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
    """Retorna (cta_fija, cta_credito, cta_prepago, cta_desc) como listas de (cuenta, nombre).
    Detecta cuentas por prefijo de número — no depende de filas sentinela."""
    fija, cred, prep, desc = [], [], [], []
    try:
        wb = openpyxl.load_workbook(io.BytesIO(plantilla_bytes))
        hoja = None
        for sn in wb.sheetnames:
            if sn.strip().upper() == "CUENTAS":
                hoja = wb[sn]; break
            if "CUENTAS" in sn.upper():
                hoja = wb[sn]
        if not hoja:
            for sn in wb.sheetnames:
                if "cuentas" in sn.lower():
                    hoja = wb[sn]; break
        if hoja:
            for row in hoja.iter_rows(min_row=1, values_only=True):
                acct = row[7] if len(row) > 7 else None
                nombre = str(row[8]).strip().replace("\n", "").strip() if len(row) > 8 and row[8] else ""
                if not acct:
                    continue
                acct = str(acct).strip()
                # Detectar por prefijo numérico (no requiere filas sentinela)
                if (acct.startswith("105-01-0001-") or acct.startswith("105-01-0002-")) and nombre:
                    fija.append((acct, nombre))
                elif acct.startswith("105-01-0003-") and nombre:
                    cred.append((acct, nombre))
                elif acct.startswith("105-01-0004-") and nombre:
                    prep.append((acct, nombre))
                elif acct.startswith("402-") and nombre:
                    desc.append((acct, nombre))
        wb.close()
    except Exception as e:
        st.warning(f"⚠ No se pudo leer la plantilla: {e}")
    return fija, cred, prep, desc

# ── Motor de generación ───────────────────────────────────────────────────────
def procesar_prepago(despachos_bytes: bytes, plantilla_bytes: bytes | None, despachos_nombre: str = "archivo.xlsx") -> tuple[bytes, list, list]:
    """
    Genera la póliza con separación Crédito/Prepago.
    Retorna (excel_bytes, logs, resumen_por_dia).
    """
    logs = []

    # Cargar cuentas de plantilla
    if plantilla_bytes:
        cta_fija, cta_credito, cta_prepago, cta_desc = _leer_plantilla(plantilla_bytes)
        logs.append(f"✅ Plantilla: {len(cta_fija)} FIJA, {len(cta_credito)} Crédito, {len(cta_prepago)} Prepago, {len(cta_desc)} Descuento.")
    else:
        cta_fija, cta_credito, cta_prepago, cta_desc = [], [], [], []
        logs.append("⚠ Sin plantilla — no hay cuentas disponibles.")

    # FIJA_COLS: prioridad a plantilla (evita hardcoded), fallback a FIJAS_DEF
    if cta_fija:
        FIJA_COLS = list(cta_fija)   # (acct, nombre) de la hoja CUENTAS
        logs.append(f"📋 {len(FIJA_COLS)} cuentas FIJA de plantilla: "
                    + ", ".join(f"{a}" for a, n in FIJA_COLS[:4]) + ("…" if len(FIJA_COLS) > 4 else ""))
    else:
        FIJA_COLS = [(a, n) for n, a in FIJAS_DEF]   # fallback hardcoded
        logs.append("ℹ️ Usando FIJA_COLS hardcoded (FIJAS_DEF).")

    CRED_COLS = list(cta_credito)
    PREP_COLS = list(cta_prepago)
    N_FIJA = len(FIJA_COLS)
    N_CRED = len(CRED_COLS)
    N_PREP = len(PREP_COLS)
    # Índices por número de cuenta
    fija_acct_idx = {a: i for i, (a, n) in enumerate(FIJA_COLS)}
    cred_acct_idx = {a: i for i, (a, n) in enumerate(CRED_COLS)}
    prep_acct_idx = {a: i for i, (a, n) in enumerate(PREP_COLS)}
    # fija canonical-name → account (para lookup via CLIENTE_TO_FIJA)
    fija_fname_to_acct = {fn: fa for fn, fa in FIJAS_DEF}
    # fija canonical-name → índice (compatible con código que use fija_name_idx)
    fija_name_idx = {}
    for fn, fa in FIJAS_DEF:
        if fa in fija_acct_idx:
            fija_name_idx[fn] = fija_acct_idx[fa]

    # Leer despachos (.xlsx o .xls) — retorna (cabecera, filas)
    hdr, src_rows = _leer_despachos_bytes(despachos_bytes, despachos_nombre)
    logs.append(f"📂 {len(src_rows):,} filas leídas del archivo.")

    # ── Mapeo de columnas por nombre (robusto a cambios de orden) ────────────
    # Nombres canónicos que buscamos en la cabecera (case-insensitive, sin espacios)
    _ALIAS = {
        "fecha":    ["fecha_hora", "fecha", "date"],
        "producto": ["producto", "product"],
        "subtotal": ["subtotal", "sub"],
        "iva":      ["iva"],
        "ieps":     ["ieps"],
        "importe":  ["importe", "total", "monto"],
        "dsc_s":    ["descuento", "descuentosubtotal", "desc_subtotal", "discount"],
        "dsc_v":    ["descuentoiva", "desc_iva", "descuento_iva"],
        "dsc_i":    ["descuentoieps", "desc_ieps", "descuento_ieps"],
        "cliente":  ["cliente", "client", "customer"],
        "tipo":     ["tipo", "type", "tipocliente", "tipo_pago"],
    }
    def _ci(hdr_list, aliases):
        """Devuelve el índice de la primera columna que coincide con algún alias."""
        norm = [h.lower().replace(" ", "").replace("_", "") for h in hdr_list]
        for alias in aliases:
            a = alias.lower().replace("_", "").replace(" ", "")
            for i, n in enumerate(norm):
                if n == a:
                    return i
        return None

    CI = {k: _ci(hdr, v) for k, v in _ALIAS.items()}
    logs.append(f"📋 Columnas detectadas: fecha={CI['fecha']} prod={CI['producto']} "
                f"imp={CI['importe']} cliente={CI['cliente']} tipo={CI['tipo']}")

    def _g(row, key, default=None):
        """Obtiene el valor de una columna por su índice detectado."""
        idx = CI.get(key)
        if idx is None or idx >= len(row):
            return default
        return row[idx]

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
        s = str(v).strip()[:10]
        for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%m/%d/%Y"):
            try: return _dt.strptime(s, fmt).date()
            except Exception: pass
        return None

    # ── Helper de matching FIJA (normalizado) ─────────────────────────────────
    import re as _re_fija
    def _nk_fija(s):
        """Normaliza para matching: strip 'Clientes/Cientes', quita no-alfanuméricos."""
        s = _re_fija.sub(r'^(CLIENTES?|CIENTES?)\s*', '', str(s or '').upper().strip())
        return _re_fija.sub(r'[^A-Z0-9]', '', s)

    def _buscar_fija_acct(cliente_raw):
        """Retorna el número de cuenta FIJA para este cliente, o None.
        Estrategia: 1) CLIENTE_TO_FIJA exacto+normalizado → 2) búsqueda normalizada en FIJA_COLS."""
        # 1a. Exacto
        fname = CLIENTE_TO_FIJA.get(cliente_raw)
        if fname is None:
            # 1b. Normalizado contra CLIENTE_TO_FIJA
            cn = _nk_fija(cliente_raw)
            for k, v in CLIENTE_TO_FIJA.items():
                if _nk_fija(k) == cn:
                    fname = v; break
        if fname is not None:
            return fija_fname_to_acct.get(fname)
        # 2. Búsqueda normalizada directa en FIJA_COLS (nombres de plantilla)
        cn = _nk_fija(cliente_raw)
        if not cn:
            return None
        # Exacto primero
        for fa, fn in FIJA_COLS:
            if _nk_fija(fn) == cn:
                return fa
        # Substring
        for fa, fn in FIJA_COLS:
            kn = _nk_fija(fn)
            if kn and (kn in cn or cn in kn):
                return fa
        return None

    # ── Detectar modo: transaccional (con Fecha_Hora) vs resumen mensual ──────
    is_resumen = CI.get("fecha") is None and CI.get("cliente") is not None
    if is_resumen:
        logs.append("📋 Modo RESUMEN mensual detectado (sin columna Fecha_Hora). "
                    "Se genera póliza de totales del período.")

        from datetime import date as _date
        FECHA_RESUMEN = _date(1900, 1, 1)   # sentinel → fila única en la póliza

        for row in src_rows:
            if not row:
                continue
            cliente_raw = str(_g(row, "cliente") or "").strip()
            if not cliente_raw:
                continue
            if cliente_raw.upper().startswith("TOTAL") or cliente_raw.upper() == "CLIENTE":
                continue   # filas de subtotal / encabezado repetido
            prod   = str(_g(row, "producto") or "").upper().strip()
            sub    = float(_g(row, "subtotal") or 0)
            iva    = float(_g(row, "iva")     or 0)
            ieps   = float(_g(row, "ieps")    or 0)
            imp    = float(_g(row, "importe") or 0)

            d = day_data[FECHA_RESUMEN]
            if prod == "GS":   d["gs"] += sub; d["iva"] += iva; d["ieps_gs"] += ieps
            elif prod == "GP": d["gp"] += sub; d["iva"] += iva; d["ieps_gp"] += ieps
            elif prod == "GD": d["gd"] += sub; d["iva"] += iva; d["ieps_gd"] += ieps

            cargo = imp

            # Mapear cliente a fija o crédito usando matching normalizado
            acct_f = _buscar_fija_acct(cliente_raw)
            if acct_f is not None:
                idx = fija_acct_idx.get(acct_f, 0)
                d["fija"][idx] += cargo
            else:
                acct = _buscar_en(CRED_COLS, cliente_raw)
                if acct and acct in cred_acct_idx:
                    d["cred"][cred_acct_idx[acct]] += cargo
                else:
                    sin_mapear.add(f"Credito:{cliente_raw}")
    else:
        # ── Modo transaccional: una fila por fecha ─────────────────────────────
        import re as _re
        _DATE_RE = _re.compile(r'^\d{4}-\d{2}-\d{2}$')
        for row in src_rows:
            if not row:
                continue
            fecha_raw = _g(row, "fecha")
            fecha = _parse_fecha(fecha_raw)
            if fecha is None:
                continue
            prod    = str(_g(row, "producto") or "").upper().strip()
            sub     = float(_g(row, "subtotal") or 0)
            iva     = float(_g(row, "iva")      or 0)
            ieps    = float(_g(row, "ieps")     or 0)
            imp     = float(_g(row, "importe")  or 0)
            dsc_s   = float(_g(row, "dsc_s")    or 0)
            dsc_v   = float(_g(row, "dsc_v")    or 0)
            dsc_i   = float(_g(row, "dsc_i")    or 0)
            cliente = str(_g(row, "cliente")    or "").strip()
            tipo    = str(_g(row, "tipo")       or "").strip()

            d = day_data[fecha]
            if prod == "GS":    d["gs"] += sub;  d["iva"] += (iva - dsc_v); d["ieps_gs"] += (ieps - dsc_i)
            elif prod == "GP":  d["gp"] += sub;  d["iva"] += (iva - dsc_v); d["ieps_gp"] += (ieps - dsc_i)
            elif prod == "GD":  d["gd"] += sub;  d["iva"] += (iva - dsc_v); d["ieps_gd"] += ieps

            d["desc"] += dsc_s - dsc_v
            cargo = imp - dsc_s - dsc_v

            if tipo in ("Contado", "Tarjeta", "Monedero"):
                acct_f = _buscar_fija_acct(cliente)
                idx = fija_acct_idx.get(acct_f, 0) if acct_f else 0
                d["fija"][idx] += cargo
                if acct_f is None:
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

    # DescuentoSubtotal — cuenta de plantilla (402-XX) o fallback hardcoded
    _desc_tot  = sum(day_data[f]["desc"] for f in sorted_dates)
    show_desc  = abs(_desc_tot) > 0.001
    if cta_desc:
        _desc_acct, _desc_name = cta_desc[0]
    else:
        _desc_acct, _desc_name = "402-01", "DescuentoSubtotal"
    DESC_META  = [(_desc_acct, _desc_name)] if show_desc else []

    ALL_COLS = (FIXED_META + FIJA_META + CRED_META + PREP_META +
                DESC_META + [(None,"TOTAL B2")] +
                ABONO_COLS + [(None,"TOTAL B2"), (None,"CONCILIACION")])

    N8  = 8
    NF_ = N8  + len(FIJA_META)
    NC_ = NF_ + len(CRED_META)
    NP_ = NC_ + len(PREP_META)
    ND_ = NP_ if show_desc else -1          # -1 = columna no presente
    NT1_= NP_ + (1 if show_desc else 0)     # TOTAL B2 Cargos
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

    # En modo resumen la "fecha" sentinel se muestra como etiqueta del período
    _SENTINEL_RESUMEN = __import__('datetime').date(1900, 1, 1)

    for ri, fecha in enumerate(sorted_dates):
        d = day_data[fecha]; r = ri + 4
        if fecha == _SENTINEL_RESUMEN:
            # Modo resumen: extraer período del nombre del archivo
            import re as _re2
            _m = _re2.search(r'(\d{4}-\d{2}-\d{2})\s+al\s+(\d{4}-\d{2}-\d{2})', despachos_nombre)
            if _m:
                from datetime import datetime as _dt2
                _d1 = _dt2.strptime(_m.group(1), '%Y-%m-%d').strftime('%d/%m/%Y')
                _d2 = _dt2.strptime(_m.group(2), '%Y-%m-%d').strftime('%d/%m/%Y')
                periodo_label = f"{_d1} al {_d2}"
            else:
                periodo_label = "PERÍODO"
            fecha_str  = periodo_label
            fecha_cell = None   # sin valor de fecha en celda (se escribe la etiqueta)
        else:
            fecha_str  = fecha.strftime("%d/%m/%Y") if hasattr(fecha, "strftime") else str(fecha)
            fecha_cell = fecha
        fixed_vals = ["CLI", fecha_cell, f"VENTAS DEL DIA {fecha_str}", f"VENTAS DEL DIA {fecha_str}", None, None, None, None]
        for ci2, val in enumerate(fixed_vals):
            # En modo resumen col 2 (fecha_cell=None) → escribir la etiqueta como texto
            if ci2 == 1 and val is None and is_resumen:
                w(r, ci2+1, fecha_str, fill_=F_WHITE, font_=DARK)
            else:
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

        if show_desc:
            desc = round(d["desc"], 2) or None
            w(r, col, desc, fill_=F_DESC, font_=DARK, fmt="#,##0.00"); col += 1
            if desc: totals["desc"] += desc

        # Posiciones 1-indexed (para fórmulas Excel) — derivadas de los índices 0-indexed
        L      = openpyxl.utils.get_column_letter
        c_tb2c = NT1_ + 1          # columna TOTAL B2 Cargos  (1-indexed)
        c_abo0 = NA_  + 1          # primera columna de abonos (1-indexed)
        c_efe  = NA_  + EFE_IDX + 1  # columna Efectivo (última abono, 1-indexed)
        c_tb2a = NT2_ + 1          # columna TOTAL B2 Abonos  (1-indexed)
        c_conc = NT2_ + 2          # columna CONCILIACION      (1-indexed)

        # TB2 Cargos — fórmula autosuma de todos los cargos (fija+cred+prep+desc)
        tb2c_py = round(
            sum(d["fija"][i] for i, _, __ in act_fija) +
            sum(d["cred"][i] for i, _, __ in act_cred) +
            sum(d["prep"][i] for i, _, __ in act_prep) +
            (d["desc"] if show_desc else 0), 2
        )
        formula_tb2c = f"=ROUND(SUM({L(9)}{r}:{L(c_tb2c-1)}{r}),2)"
        w(r, col, formula_tb2c, fill_=F_TB2, font_=DARK_B, fmt="#,##0.00"); col += 1
        totals["tb2c"] += tb2c_py

        # Abonos individuales (Python values, excepto Efectivo que es fórmula)
        gs_ = round(d["gs"],      2); gp_  = round(d["gp"],      2); gd_  = round(d["gd"],      2)
        iva_= round(d["iva"],     2); igs_ = round(d["ieps_gs"], 2)
        igp_= round(d["ieps_gp"],2); igd_ = round(d["ieps_gd"], 2)
        otros    = round(gs_+gp_+gd_+iva_+igs_+igp_+igd_, 2)
        efectivo = round(tb2c_py - otros, 2)
        abo_vals = [gs_, gp_, gd_, iva_, igs_, igp_, igd_, efectivo]

        for j, (ac, nm) in enumerate(ABONO_COLS):
            df = F_EFE if j == EFE_IDX else F_ABO
            if j == EFE_IDX:
                # Efectivo = TB2 Cargos − suma de los demás abonos → garantiza CONC = 0
                formula_efe = f"=ROUND({L(c_tb2c)}{r}-SUM({L(c_abo0)}{r}:{L(c_efe-1)}{r}),2)"
                w(r, col, formula_efe, fill_=df, font_=DARK, fmt="#,##0.00")
            else:
                val = round(abo_vals[j], 2) or None
                w(r, col, val, fill_=df, font_=DARK, fmt="#,##0.00")
                if val: totals[f"a{j}"] += abo_vals[j]
            col += 1
        totals[f"a{EFE_IDX}"] += efectivo

        # TB2 Abonos — autosuma de todos los abonos (incluye Efectivo)
        formula_tb2a = f"=ROUND(SUM({L(c_abo0)}{r}:{L(c_efe)}{r}),2)"
        w(r, col, formula_tb2a, fill_=F_TB2, font_=DARK_B, fmt="#,##0.00"); col += 1
        totals["tb2a"] += tb2c_py  # por diseño tb2a = tb2c

        # CONCILIACION = TB2 Cargos − TB2 Abonos (Excel calcula; debe ser 0)
        formula_conc = f"=ROUND({L(c_tb2c)}{r}-{L(c_tb2a)}{r},2)"
        w(r, col, formula_conc, fill_=F_CONC, font_=DARK_B, fmt="#,##0.00")

        resumen.append({
            "Fecha":        fecha_str,
            "TB2 Cargos":   round(tb2c_py, 2),
            "TB2 Abonos":   round(tb2c_py, 2),
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
