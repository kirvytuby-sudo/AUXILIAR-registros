"""AUXILIAR DE REGISTROS â Ventas del DÃ­a (Control de Despachos â PÃ³liza Excel)"""
import streamlit as st
import io
from collections import defaultdict
from datetime import datetime

st.set_page_config(page_title="Ventas del DÃ­a Â· Auxiliar", page_icon="â½", layout="wide")

st.markdown("""<style>
[data-testid="stAppViewContainer"] { background: #dbeafe; }
.header-bar{background:#1E3A8A;padding:18px 28px;border-radius:10px;margin-bottom:20px;}
.header-bar h1{color:#FBCFE8;font-size:1.5rem;margin:0;}
.header-bar p{color:#93C5FD;margin:4px 0 0;font-size:.9rem;}
#MainMenu{visibility:hidden;}footer{visibility:hidden;}
</style>
<div class="header-bar">
  <h1>â½ Ventas del DÃ­a</h1>
  <p>Genera la pÃ³liza contable de ventas diarias desde el control de despachos</p>
</div>
""", unsafe_allow_html=True)

try:
    import xlsxwriter
    import openpyxl
except ImportError as _e:
    st.error(f"â LibrerÃ­a faltante: {_e}. Verifica requirements.txt.")
    st.stop()

# ââ Cuentas IEPS (fijas) ââââââââââââââââââââââââââââââââââââââââââââââââââââ
IEPS_GS = "401-01-0001-0006-0001"
IEPS_GP = "401-01-0001-0006-0002"
IEPS_GD = "401-01-0001-0006-0003"

CLIENTES_TPL = [
    "T. BANORTE","Contado","T. EDENRED (TICKET CAR)","T. EFECTICARD",
    "T. SODEXO (GASOPASS)","T. AMERICAN EXPRESS","V. EFECTIVALE",
    "ACEROS Y CEMENTOS TEPEPAN SA DE CV","CONCESIONARIA KIOTO SA DE CV",
    "EXCELENCIA COREANA S.A. DE C.V.","GOTESA SA DE CV",
    "INGENIERIA Y SERVICIO EN TRATAMIENTO DE AGUA SA DE CV",
    "MF CONSTRUCCIONES Y ASOCIADOS SA DE CV","PHIEMES SA DE CV",
    "SUOMI PUBLICITY AND LEAFLETING",
]
CUENTAS_TPL = [
    "105-01-0001-0004","105-01-0001-0001","105-01-0002-0002","105-01-0002-0009",
    "105-01-0002-0008","105-01-0001-0007","105-01-0002-0005",
    "105-01-0003-0005","105-01-0003-0602",
    "105-01-0003-1200","105-01-0003-1801",
    "105-01-0003-2400",
    "105-01-0003-3601","105-01-0003-4700",
    "105-01-0003-5701",
]
META_HDRS = ["TIPO DE POLIZA", "Fecha", "REFERENCIA", "CONCEPTO", "ERROR", "UIDD", "NUM POLIZA", "PROCESADO"]
PRODS     = ["GS", "GP", "GD", "IVA", "IEPS GS", "IEPS GP", "IEPS GD"]
CTAS_PROD = [
    "401-01-0001-0001", "401-01-0001-0002", "401-01-0001-0003", "209-01",
    IEPS_GS, IEPS_GP, IEPS_GD,
]


def _leer_cuentas_plantilla(plantilla_bytes):
    """Lee la hoja CUENTAS de la plantilla; retorna dict {cuenta: nombre}."""
    cuentas_map = {
        IEPS_GS: "IEPS De Gasolina Magna",
        IEPS_GP: "IEPS de Premium",
        IEPS_GD: "IEPS de Diesel",
    }
    try:
        wb = openpyxl.load_workbook(io.BytesIO(plantilla_bytes), data_only=True)
        hoja = None
        for sn in wb.sheetnames:
            if sn.strip().upper() == "CUENTAS":
                hoja = wb[sn]
                break
        if hoja:
            for r in hoja.iter_rows(values_only=True):
                for idx_num, idx_nom in [(7, 8), (11, 12)]:
                    num = r[idx_num] if len(r) > idx_num else None
                    nom = r[idx_nom] if len(r) > idx_nom else None
                    if num and nom and str(num)[0] in ('1', '2', '4'):
                        cuentas_map[str(num).strip()] = str(nom).strip().replace('\n', '')
        wb.close()
    except Exception:
        pass
    return cuentas_map


def _leer_despachos(file_bytes, filename, logs):
    """Lee el archivo de despachos (.xlsx o .xls). Retorna lista de filas (sin encabezado)."""
    ext = filename.rsplit('.', 1)[-1].lower() if '.' in filename else 'xlsx'
    if ext == 'xls':
        try:
            import xlrd
            wb_xls = xlrd.open_workbook(file_contents=file_bytes)
            ws_xls = wb_xls.sheet_by_index(0)
            rows = [tuple(ws_xls.row_values(r)) for r in range(ws_xls.nrows)]
            data = rows[1:]
            logs.append(f"  {len(data):,} registros leÃ­dos (.xls vÃ­a xlrd).")
            return data
        except ImportError:
            raise RuntimeError(
                "El archivo estÃ¡ en formato antiguo .xls y la librerÃ­a 'xlrd' no estÃ¡ instalada.\n"
                "SoluciÃ³n: Abre el archivo en Excel y guÃ¡rdalo como .xlsx."
            )
    else:
        wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True, read_only=True)
        data = list(wb.active.iter_rows(values_only=True))[1:]
        wb.close()
        logs.append(f"  {len(data):,} registros leÃ­dos (.xlsx).")
        return data


def procesar_ventas(despachos_bytes, despachos_nombre, plantilla_bytes=None):
    """
    Procesa el control de despachos y genera la pÃ³liza Excel.
    Retorna (excel_bytes: bytes, logs: list[str]).
    """
    logs = []

    logs.append("â½ Leyendo control de despachos...")
    data = _leer_despachos(despachos_bytes, despachos_nombre, logs)

    # ââ Leer plantilla de cuentas ââââââââââââââââââââââââââââââââââââââââââ
    if plantilla_bytes:
        logs.append("â½ Leyendo plantilla de cuentas...")
        cuentas_map = _leer_cuentas_plantilla(plantilla_bytes)
        logs.append(f"  {len(cuentas_map)} cuentas cargadas.")
    else:
        cuentas_map = {
            IEPS_GS: "IEPS De Gasolina Magna",
            IEPS_GP: "IEPS de Premium",
            IEPS_GD: "IEPS de Diesel",
        }
        logs.append("  â¹ Sin plantilla â usando nombres predeterminados.")

    NOMBRES_TPL = [cuentas_map.get(c, cli) for c, cli in zip(CUENTAS_TPL, CLIENTES_TPL)]
    NOMS_PROD   = [cuentas_map.get(c, p)   for c, p   in zip(CTAS_PROD, PRODS)]

    # ââ Acumular por fecha / cliente / producto ââââââââââââââââââââââââââââ
    cli_day   = defaultdict(float)
    prod_day  = defaultdict(float)
    iva_day   = defaultdict(float)
    ieps_prod = defaultdict(float)

    for r in data:
        try:
            fecha   = str(r[0])[:10] if r[0] is not None else ""
            cliente = str(r[16] or "").strip()
            prod    = str(r[3]  or "")
            cli_day[(fecha, cliente)]  += float(r[9]  or 0)
            prod_day[(fecha, prod)]    += float(r[6]  or 0)
            iva_day[fecha]             += float(r[7]  or 0)
            ieps_prod[(fecha, prod)]   += float(r[8]  or 0)
        except Exception:
            continue

    fechas = sorted(set(k[0] for k in cli_day if k[0]))
    if not fechas:
        raise RuntimeError("No se encontraron datos de ventas en el archivo.")
    logs.append(f"  {len(fechas)} fecha(s) detectada(s): {fechas[0]} â¦ {fechas[-1]}")

    # ââ Ãndices de columnas ââââââââââââââââââââââââââââââââââââââââââââââââ
    N_META    = len(META_HDRS)    # 8
    N_CLI     = len(CLIENTES_TPL) # 20
    N_PROD    = len(PRODS)        # 7
    OFF       = N_META            # 8  â inicio clientes
    COL_TOT1  = OFF + N_CLI       # 28 â TOTAL B2 clientes
    COL_PROD0 = OFF + N_CLI + 1   # 29 â inicio productos
    COL_TOT2  = OFF + N_CLI + 1 + N_PROD  # 36 â TOTAL B2 productos
    COL_CONC  = OFF + N_CLI + 1 + N_PROD + 1  # 37 â CONCILIACION
    TOTAL_COLS = COL_CONC + 1     # 38

    # ââ Generar Excel con xlsxwriter âââââââââââââââââââââââââââââââââââââââ
    logs