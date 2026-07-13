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
    "T BANORTE", "Contado", "T EDENRED (ACCOR)", "T EFECTIVALE",
    "M Y T INTEGRALES PARA LA SALUD", "T AMERICAN EXPRESS",
    "CENTRO DE DISTRIBUCION ORIENTE", "T ULTRAGAS", "T PLUXEE MEXICO",
    "JUAN ANTONIO CRUZ MONDRAGON", "T INBURSA", "V EFECTIVALE",
    "ETIQUETAS CCL", "MARIA DEL CARMEN PEÃALOZA PEIMBERT",
    "ROTULACIONES E IMPRESIONES MEXICANAS SA DE CV",
    "ALMACENAJE Y DISTRIBUCION TRANSGALLA", "PETRO ASFALTOS DEL SURESTE",
    "GRAFIARTE DELA", "MARICELA GONZALEZ RODRIGUEZ", "ADEPT SERVICES MEXICO",
]
CUENTAS_TPL = [
    "105-01-0001-0004", "105-01-0001-0001", "105-01-0002-0002", "105-01-0002-0001",
    "105-01-0003-3600", "105-01-0001-0007", "105-01-0003-0601", "105-01-0002-0003",
    "105-01-0002-0007", "105-01-0003-2700", "105-01-0002-0006", "105-01-0002-0005",
    "105-01-0003-1200", "105-01-0003-3601", "105-01-0003-5400", "105-01-0003-0002",
    "105-01-0003-4701", "105-01-0003-1800", "105-01-0003-3602", "105-01-0003-0001",
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
    logs.append("â½ Generando pÃ³liza Excel...")
    buf = io.BytesIO()
    wb  = xlsxwriter.Workbook(buf, {'in_memory': True})
    ws  = wb.add_worksheet("poliza IA")

    def fmt(d):
        return wb.add_format(d)

    CURR = 'General'
    BASE = {'font_name': 'Arial', 'font_size': 9, 'border': 1,
            'border_color': '#CCCCCC', 'valign': 'vcenter'}

    f_title   = fmt({'bold': True, 'font_color': '#FFFFFF', 'bg_color': '#C2185B',
                     'font_size': 12, 'align': 'center', 'valign': 'vcenter', 'font_name': 'Arial'})
    f_acct    = fmt({**BASE, 'bold': True, 'bg_color': '#3B0764', 'font_color': '#FFFFFF',
                     'align': 'center', 'font_size': 8})
    f_hdr_m   = fmt({**BASE, 'bold': True, 'bg_color': '#6A1B9A', 'font_color': '#FFFFFF',
                     'align': 'center', 'text_wrap': True, 'font_size': 8})
    f_hdr_c   = fmt({**BASE, 'bold': True, 'bg_color': '#1565C0', 'font_color': '#FFFFFF',
                     'align': 'center', 'text_wrap': True, 'font_size': 8})
    f_hdr_p   = fmt({**BASE, 'bold': True, 'bg_color': '#0D47A1', 'font_color': '#FFFFFF',
                     'align': 'center', 'text_wrap': True, 'font_size': 8})
    f_hdr_tot = fmt({**BASE, 'bold': True, 'bg_color': '#1B5E20', 'font_color': '#FFFFFF',
                     'align': 'center', 'font_size': 8})
    f_hdr_con = fmt({**BASE, 'bold': True, 'bg_color': '#E65100', 'font_color': '#FFFFFF',
                     'align': 'center', 'font_size': 8})
    f_fecha   = fmt({**BASE, 'bold': True, 'bg_color': '#F3E5F5', 'align': 'center'})
    f_meta    = fmt({**BASE, 'bg_color': '#EDE7F6', 'align': 'left', 'font_size': 8})
    f_num0    = fmt({**BASE, 'bg_color': '#FFFFFF', 'align': 'right', 'num_format': CURR})
    f_num1    = fmt({**BASE, 'bg_color': '#B7D9EF', 'align': 'right', 'num_format': CURR})
    f_tot0    = fmt({**BASE, 'bold': True, 'bg_color': '#E6F2FB', 'align': 'right', 'num_format': CURR})
    f_tot1    = fmt({**BASE, 'bold': True, 'bg_color': '#DCEDC8', 'align': 'right', 'num_format': CURR})
    f_conc0   = fmt({**BASE, 'bold': True, 'bg_color': '#FFF9C4', 'font_color': '#E65100',
                     'align': 'right', 'num_format': CURR})
    f_conc1   = fmt({**BASE, 'bold': True, 'bg_color': '#FFF176', 'font_color': '#E65100',
                     'align': 'right', 'num_format': CURR})
    f_grand   = fmt({**BASE, 'bold': True, 'bg_color': '#1B5E20', 'font_color': '#FFFFFF',
                     'align': 'right', 'num_format': CURR, 'border': 2, 'border_color': '#000000'})
    f_grand_l = fmt({**BASE, 'bold': True, 'bg_color': '#1B5E20', 'font_color': '#FFFFFF',
                     'align': 'center', 'border': 2, 'border_color': '#000000'})
    f_grand_c = fmt({**BASE, 'bold': True, 'bg_color': '#E65100', 'font_color': '#FFFFFF',
                     'align': 'right', 'num_format': CURR, 'border': 2, 'border_color': '#000000'})

    # ââ Fila 0: tÃ­tulo âââââââââââââââââââââââââââââââââââââââââââââââââââââ
    ws.set_row(0, 24)
    ws.set_row(1, 18)
    ws.set_row(2, 50)
    titulo = "VENTAS DEL DIA â SUPER SERVICIO PERIFERICO"
    ws.merge_range(0, 0, 0, TOTAL_COLS - 1, titulo, f_title)

    # ââ Fila 1: nÃºmeros de cuenta ââââââââââââââââââââââââââââââââââââââââââ
    for c in range(N_META):
        ws.write(1, c, "", f_acct)
    for i, cta in enumerate(CUENTAS_TPL):
        ws.write(1, OFF + i, cta, f_acct)
    ws.write(1, COL_TOT1, "", f_acct)
    for i, cta in enumerate(CTAS_PROD):
        ws.write(1, COL_PROD0 + i, cta, f_acct)
    ws.write(1, COL_TOT2, "", f_acct)
    ws.write(1, COL_CONC, "", f_acct)

    # ââ Fila 2: encabezados ââââââââââââââââââââââââââââââââââââââââââââââââ
    for i, h in enumerate(META_HDRS):
        ws.write(2, i, h, f_hdr_m)
    for i, nom in enumerate(NOMBRES_TPL):
        ws.write(2, OFF + i, nom, f_hdr_c)
    ws.write(2, COL_TOT1, "TOTAL B2", f_hdr_tot)
    for i, nom in enumerate(NOMS_PROD):
        ws.write(2, COL_PROD0 + i, nom, f_hdr_p)
    ws.write(2, COL_TOT2, "TOTAL B2", f_hdr_tot)
    ws.write(2, COL_CONC, "CONCILIACION", f_hdr_con)

    # ââ Anchos de columna ââââââââââââââââââââââââââââââââââââââââââââââââââ
    ws.set_column(0, 0, 14)
    ws.set_column(1, 1, 12)
    for c in range(2, N_META):
        ws.set_column(c, c, 20)
    for i in range(N_CLI):
        ws.set_column(OFF + i, OFF + i, 14)
    ws.set_column(COL_TOT1, COL_TOT1, 13)
    for i in range(N_PROD):
        ws.set_column(COL_PROD0 + i, COL_PROD0 + i, 13)
    ws.set_column(COL_TOT2, COL_TOT2, 13)
    ws.set_column(COL_CONC, COL_CONC, 14)
    ws.freeze_panes(3, 2)

    # ââ Filas de datos âââââââââââââââââââââââââââââââââââââââââââââââââââââ
    gran_cli  = [0.0] * N_CLI
    gran_tot1 = 0.0
    gran_prod = [0.0] * N_PROD
    gran_tot2 = 0.0
    gran_conc = 0.0

    for ri, fecha in enumerate(fechas):
        row = ri + 3
        fn  = f_num0 if ri % 2 == 0 else f_num1
        ft  = f_tot0 if ri % 2 == 0 else f_tot1
        fc  = f_conc0 if ri % 2 == 0 else f_conc1

        # Convertir fecha a DD/MM/YYYY
        try:
            _fd = datetime.strptime(fecha[:10], '%Y-%m-%d')
            fecha_display = _fd.strftime('%d/%m/%Y')
        except Exception:
            fecha_display = fecha

        ws.write(row, 0, "D", f_meta)
        ws.write(row, 1, fecha_display, f_fecha)
        ws.write(row, 2, "VENTA DEL DIA " + fecha_display, f_meta)
        ws.write(row, 3, "VENTA DEL DIA " + fecha_display, f_meta)
        for c in range(4, N_META):
            ws.write(row, c, "", f_meta)

        # Clientes
        total_b2 = 0.0
        for i, cli in enumerate(CLIENTES_TPL):
            v = round(cli_day.get((fecha, cli), 0.0), 2)
            ws.write(row, OFF + i, v if v else None, fn)
            total_b2 += v
            gran_cli[i] += v
        ws.write(row, COL_TOT1, round(total_b2, 2), ft)
        gran_tot1 += total_b2

        # Productos
        prod_vals = [
            prod_day.get((fecha, "GS"), 0.0),
            prod_day.get((fecha, "GP"), 0.0),
            prod_day.get((fecha, "GD"), 0.0),
            iva_day.get(fecha, 0.0),
            ieps_prod.get((fecha, "GS"), 0.0),
            ieps_prod.get((fecha, "GP"), 0.0),
            ieps_prod.get((fecha, "GD"), 0.0),
        ]
        total_prod = sum(prod_vals)
        for i, v in enumerate(prod_vals):
            ws.write(row, COL_PROD0 + i, round(v, 2) if v else None, fn)
            gran_prod[i] += v
        ws.write(row, COL_TOT2, round(total_prod, 2), ft)
        gran_tot2 += total_prod

        diferencia = round(total_b2 - total_prod, 2)
        ws.write(row, COL_CONC, diferencia, fc)
        gran_conc += diferencia

    # ââ Fila totales generales âââââââââââââââââââââââââââââââââââââââââââââ
    tr = len(fechas) + 3
    ws.merge_range(tr, 0, tr, N_META - 1, "TOTAL GENERAL", f_grand_l)
    for i in range(N_CLI):
        ws.write(tr, OFF + i, round(gran_cli[i], 2), f_grand)
    ws.write(tr, COL_TOT1, round(gran_tot1, 2), f_grand)
    for i in range(N_PROD):
        ws.write(tr, COL_PROD0 + i, round(gran_prod[i], 2), f_grand)
    ws.write(tr, COL_TOT2, round(gran_tot2, 2), f_grand)
    ws.write(tr, COL_CONC, round(gran_conc, 2), f_grand_c)

    wb.close()
    buf.seek(0)

    logs.append(f"â PÃ³liza generada â {len(fechas)} fecha(s), {TOTAL_COLS} columnas.")
    if abs(gran_conc) < 0.02:
        logs.append(f"â CONCILIACION cuadra: {gran_conc:,.2f}")
    else:
        logs.append(f"â ï¸  CONCILIACION con diferencia: {gran_conc:,.2f}")

    # ââ Resumen por fecha (para mostrar en UI) âââââââââââââââââââââââââââââ
    resumen = []
    for fecha in fechas:
        try:
            fd = datetime.strptime(fecha[:10], '%Y-%m-%d')
            fd_str = fd.strftime('%d/%m/%Y')
        except Exception:
            fd_str = fecha
        tot_cli  = sum(cli_day.get((fecha, cli), 0.0) for cli in CLIENTES_TPL)
        tot_prod = (
            prod_day.get((fecha, "GS"), 0.0) +
            prod_day.get((fecha, "GP"), 0.0) +
            prod_day.get((fecha, "GD"), 0.0) +
            iva_day.get(fecha, 0.0) +
            ieps_prod.get((fecha, "GS"), 0.0) +
            ieps_prod.get((fecha, "GP"), 0.0) +
            ieps_prod.get((fecha, "GD"), 0.0)
        )
        resumen.append({
            "Fecha": fd_str,
            "Total Clientes": round(tot_cli, 2),
            "Total Productos": round(tot_prod, 2),
            "Diferencia": round(tot_cli - tot_prod, 2),
        })

    return buf.read(), logs, resumen


# ââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââ
# UI
# ââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââââ

st.markdown("### ð Archivos de entrada")
col1, col2 = st.columns([1, 1])
with col1:
    despachos_file = st.file_uploader(
        "ð Control de despachos (.xlsx / .xls)",
        type=["xlsx", "xls"],
        help="Archivo de control de despachos generado por el sistema de ventas.",
    )
with col2:
    plantilla_file = st.file_uploader(
        "ð Plantilla de cuentas (.xlsx) â opcional",
        type=["xlsx"],
        help="Plantilla SINUBE con hoja 'CUENTAS'. Si no se proporciona, se usan nombres predeterminados.",
    )

st.markdown("")
generar = st.button(
    "â½  Generar PÃ³liza Ventas del DÃ­a",
    type="primary",
    disabled=despachos_file is None,
    use_container_width=True,
)

if despachos_file is None:
    st.info("ð Selecciona al menos el archivo de control de despachos para comenzar.")

if generar and despachos_file is not None:
    with st.spinner("Procesando..."):
        try:
            plantilla_bytes = plantilla_file.read() if plantilla_file else None
            excel_bytes, logs, resumen = procesar_ventas(
                despachos_file.read(),
                despachos_file.name,
                plantilla_bytes,
            )

            base = despachos_file.name.rsplit('.', 1)[0]
            nombre_salida = f"poliza_ventas_dia_{base}.xlsx"

            st.success(f"â PÃ³liza generada â {len(resumen)} fecha(s)")

            st.download_button(
                label="ð¾  Descargar pÃ³liza Excel",
                data=excel_bytes,
                file_name=nombre_salida,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

            # ââ Tabla resumen ââââââââââââââââââââââââââââââââââââââââââââââ
            if resumen:
                st.markdown("### ð Resumen por fecha")
                import pandas as pd
                df = pd.DataFrame(resumen)
                # Color diferencia
                def _color_diff(v):
                    if abs(v) < 0.02:
                        return "background-color:#D1FAE5; color:#065F46"
                    return "background-color:#FEF3C7; color:#92400E"

                styled = (
                    df.style
                    .format({
                        "Total Clientes":  "{:,.2f}",
                        "Total Productos": "{:,.2f}",
                        "Diferencia":      "{:,.2f}",
                    })
                    .map(_color_diff, subset=["Diferencia"])
                )
                st.dataframe(styled, use_container_width=True, hide_index=True)

            # ââ Log ââââââââââââââââââââââââââââââââââââââââââââââââââââââââ
            with st.expander("ð Log de procesamiento"):
                for line in logs:
                    st.text(line)

        except Exception as exc:
            import traceback
            st.error(f"â Error al generar pÃ³liza: {exc}")
            with st.expander("Detalle del error"):
                st.code(traceback.format_exc())

st.markdown("---")
st.caption("MÃ³dulo Ventas del DÃ­a Â· AUXILIAR DE REGISTROS")
