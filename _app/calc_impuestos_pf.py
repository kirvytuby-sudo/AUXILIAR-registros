# ─────────────────────────────────────────────────────────────────────────────
#  calc_impuestos_pf.py  —  Motor de cálculo ISR/IVA Régimen 612
#  Persona Física | Art.106 LISR (ISR provisional) | Art.5-D LIVA (IVA mensual)
#  Versión 1.0  —  Oct 2026
# ─────────────────────────────────────────────────────────────────────────────

import glob, os, re, datetime
from xml.etree import ElementTree as ET
from openpyxl import Workbook
from openpyxl.styles import (PatternFill, Font, Alignment, Border, Side)
from openpyxl.utils import get_column_letter

try:
    import pdfplumber
    _PDFPLUMBER_OK = True
except ImportError:
    _PDFPLUMBER_OK = False

# ─── Constantes ───────────────────────────────────────────────────────────────
MESES = {1:'ENERO',2:'FEBRERO',3:'MARZO',4:'ABRIL',5:'MAYO',6:'JUNIO',
         7:'JULIO',8:'AGOSTO',9:'SEPTIEMBRE',10:'OCTUBRE',11:'NOVIEMBRE',12:'DICIEMBRE'}
MESES_C = {1:'ENE',2:'FEB',3:'MAR',4:'ABR',5:'MAY',6:'JUN',
            7:'JUL',8:'AGO',9:'SEP',10:'OCT',11:'NOV',12:'DIC'}

# Tarifa ISR 2024 Art.96 LISR (límite inf, límite sup, cuota fija, % excedente)
TARIFA = [
    (0.01,        5_952.84,      0.00,      1.92),
    (5_952.85,   50_524.92,    114.24,      6.40),
    (50_524.93,  88_793.04,  2_966.91,     10.88),
    (88_793.05, 103_218.00,  7_130.48,     16.00),
    (103_218.01,123_580.20, 9_438.47,      17.92),
    (123_580.21,249_243.48,13_087.37,      21.36),
    (249_243.49,391_841.97,39_929.05,      23.52),
    (391_841.98,750_000.00,73_524.64,      30.00),
    (750_000.01,1_000_000.00,181_121.21,   32.00),
    (1_000_000.01,3_000_000.00,261_121.21, 34.00),
    (3_000_000.01,999_999_999,941_121.22,  35.00),
]

# INPC mensual 2024-2026 (índice general)
INPC_TABLE = {
    (2024, 1):128.648,(2024, 2):129.528,(2024, 3):130.185,(2024, 4):131.028,
    (2024, 5):131.453,(2024, 6):131.882,(2024, 7):132.323,(2024, 8):132.609,
    (2024, 9):133.148,(2024,10):133.672,(2024,11):134.247,(2024,12):134.889,
    (2025, 1):136.172,(2025, 2):137.130,(2025, 3):137.714,(2025, 4):138.203,
    (2025, 5):138.508,(2025, 6):138.870,(2025, 7):139.294,(2025, 8):139.766,
    (2025, 9):140.394,(2025,10):140.845,(2025,11):141.478,(2025,12):142.210,
    (2026, 1):143.287,(2026, 2):144.043,(2026, 3):144.624,(2026, 4):145.008,
    (2026, 5):145.381,(2026, 6):145.633,(2026, 7):145.750,(2026, 8):145.750,
    (2026, 9):145.750,(2026,10):145.750,(2026,11):145.750,(2026,12):145.750,
}
INPC_REF = 145.75  # Septiembre 2026 (referencia actualización)

# Tasa de recargos Art.21 CFF (mensual)
RECARGO_RATES = {2024: 0.0147, 2025: 0.0147, 2026: 0.0207}

# ─── Clasificación de gastos ──────────────────────────────────────────────────
NO_DED_PROV   = ['GASOLINA','GAS','UBER','DIDI','TAXI','COMBUSTIBLE','FERRETERIA','FERRETERÍA']
NO_DED_CONC   = ['MULTA','RECARGO','SANCION','SANCIÓN','DONATIVO','DONACION','DONACIÓN',
                  'CLAUSULA PENAL','CLÁUSULA PENAL']
NO_DED_MED_PROV = ['FARMACIA','DROGUE','MEDICAMENTO','SALUD']
NO_DED_MED_CONC = ['MEDICAMENTO','MEDICINA','CONSULTA MEDICA','ANÁLISIS CLINICO','LABORATORIO']

def clasificar_gasto(conc: str, nom_e: str = '') -> tuple:
    """Regresa ('DEDUCIBLE','') o ('NO_DED', motivo)."""
    c = (conc or '').upper(); n = (nom_e or '').upper()
    for kw in NO_DED_CONC:
        if kw in c: return 'NO_DED', f'Concepto no deducible: {kw}'
    for kw in NO_DED_PROV:
        if kw in n or kw in c: return 'NO_DED', f'Proveedor/concepto no autorizado: {kw}'
    for kw in NO_DED_MED_PROV:
        if kw in n: return 'NO_DED', f'Farmacia/salud (uso personal): {kw}'
    for kw in NO_DED_MED_CONC:
        if kw in c: return 'NO_DED', f'Gastos médicos personales: {kw}'
    return 'DEDUCIBLE', ''

# ─── Cálculo ISR ─────────────────────────────────────────────────────────────
def calc_isr(base: float) -> float:
    """Aplica tarifa Art.96 LISR sobre base acumulada anual."""
    if base <= 0: return 0.0
    for li, ls, cf, pct in TARIFA:
        if li <= base <= ls:
            return round(cf + (base - li) * pct / 100, 2)
    li, ls, cf, pct = TARIFA[-1]
    return round(cf + (base - li) * pct / 100, 2)

# ─── Recargos y actualización Art.21 / Art.17-A CFF ─────────────────────────
def calc_rec_act(monto: float, mes_imp: int, year_imp: int = 2024):
    """Calcula recargos y actualización para un pago en INPC_REF."""
    if monto <= 0: return 0.0, 0.0, 0.0
    rate  = RECARGO_RATES.get(year_imp, 0.0147)
    meses = max(1, 12 - mes_imp + 1)          # meses de retraso hasta dic
    rec   = round(monto * rate * meses, 2)
    inpc_orig = INPC_TABLE.get((year_imp, mes_imp), INPC_REF)
    fact_act  = INPC_REF / inpc_orig
    act       = round(monto * (fact_act - 1), 2) if fact_act > 1 else 0.0
    total_cr  = round(monto + rec + act, 2)
    return rec, act, total_cr

# ─── Lector de PDFs BBVA ─────────────────────────────────────────────────────
def _leer_pdfs_bbva(pdfs: list, log=print) -> list:
    """
    Extrae movimientos de estados de cuenta BBVA (PDF).
    Retorna lista de dicts: fecha, desc, dep, ret, banco.
    """
    if not _PDFPLUMBER_OK:
        log("⚠ pdfplumber no instalado — PDFs BBVA omitidos"); return []
    movs = []
    for path in pdfs:
        try:
            with pdfplumber.open(path) as pdf:
                in_det = False
                for page in pdf.pages:
                    txt = page.extract_text() or ''
                    for line in txt.splitlines():
                        if 'DETALLE DE MOVIMIENTOS' in line.upper(): in_det = True; continue
                        if not in_det: continue
                        m = re.match(
                            r'(\d{2}/\d{2}/\d{4})\s+(.+?)\s+([\d,]+\.\d{2})?\s*([\d,]+\.\d{2})?$',
                            line.strip())
                        if not m: continue
                        fecha_s, desc, dep_s, ret_s = m.groups()
                        try:
                            fecha = datetime.datetime.strptime(fecha_s, '%d/%m/%Y').date()
                        except Exception:
                            continue
                        dep = float((dep_s or '0').replace(',',''))
                        ret = float((ret_s or '0').replace(',',''))
                        movs.append({'fecha':fecha,'desc':desc.strip(),'dep':dep,'ret':ret,'banco':'BBVA'})
        except Exception as e:
            log(f"  ✗ PDF BBVA {os.path.basename(path)}: {e}")
    return movs

# ─── Lector de Excel de bancos (Banorte / Inbursa / extra) ───────────────────
def _leer_excels_banco(excels: list, banco: str, log=print) -> list:
    """
    Extrae movimientos de estados de cuenta en formato Excel.
    Detecta automáticamente columnas FECHA y ABONO/DEP/CARGO.
    """
    try:
        import openpyxl as ox
    except ImportError:
        log("⚠ openpyxl no instalado"); return []
    movs = []
    KEY_FECHA = ['FECHA','DATE']
    KEY_DEP   = ['ABONO','DEPOSITO','DEPÓSITO','CREDITO','CRÉDITO','CARGO_ABONO',
                 'DEP','CREDIT','IMPORTE_ABONO']
    KEY_RET   = ['CARGO','DEBITO','DÉBITO','RETIRO','DEBIT','IMPORTE_CARGO']
    KEY_DESC  = ['DESCRIPCION','DESCRIPCIÓN','CONCEPTO','MOVIMIENTO','REFERENCIA',
                 'DESC','DETAIL','NARRATION']

    for path in excels:
        try:
            wb = ox.load_workbook(path, data_only=True)
            ws = wb.active
            hdr_row = None; col_map = {}
            for ri, row in enumerate(ws.iter_rows(max_row=20), 1):
                row_vals = [(c.column, str(c.value or '').upper().strip()) for c in row if c.value]
                if not row_vals: continue
                tmp = {}
                for ci, v in row_vals:
                    if any(k in v for k in KEY_FECHA) and 'fecha' not in tmp: tmp['fecha']=ci
                    if any(k in v for k in KEY_DEP)   and 'dep'   not in tmp: tmp['dep']  =ci
                    if any(k in v for k in KEY_RET)   and 'ret'   not in tmp: tmp['ret']  =ci
                    if any(k in v for k in KEY_DESC)  and 'desc'  not in tmp: tmp['desc'] =ci
                if 'fecha' in tmp and ('dep' in tmp or 'ret' in tmp):
                    hdr_row = ri; col_map = tmp; break
            if hdr_row is None:
                log(f"  ✗ Excel {banco} {os.path.basename(path)}: no se detectó encabezado"); continue
            for row in ws.iter_rows(min_row=hdr_row+1):
                def gv(ci): return row[ci-1].value if ci and ci <= len(row) else None
                fval = gv(col_map['fecha'])
                if fval is None: continue
                if isinstance(fval, datetime.datetime): fecha = fval.date()
                elif isinstance(fval, datetime.date):   fecha = fval
                else:
                    try: fecha = datetime.datetime.strptime(str(fval).strip(), '%d/%m/%Y').date()
                    except Exception:
                        try: fecha = datetime.datetime.strptime(str(fval).strip(), '%Y-%m-%d').date()
                        except Exception: continue
                dep_v = gv(col_map.get('dep')); ret_v = gv(col_map.get('ret'))
                dep = float(str(dep_v).replace(',','')) if dep_v else 0.0
                ret = float(str(ret_v).replace(',','')) if ret_v else 0.0
                desc_v = gv(col_map.get('desc',''))
                desc = str(desc_v).strip() if desc_v else ''
                if dep == 0 and ret == 0: continue
                movs.append({'fecha':fecha,'desc':desc,'dep':dep,'ret':ret,'banco':banco})
        except Exception as e:
            log(f"  ✗ Excel {banco} {os.path.basename(path)}: {e}")
    return movs

# ─── Lector de XMLs CFDI ─────────────────────────────────────────────────────
_NS = {'cfdi': 'http://www.sat.gob.mx/cfd/4',
       'cfdi3': 'http://www.sat.gob.mx/cfd/3',
       'tfd':  'http://www.sat.gob.mx/TimbreFiscalDigital'}

def _parse_xml(path: str, rfc_contribuyente: str, year: int):
    """Parsea un CFDI XML. Retorna dict o None si no aplica."""
    try:
        tree = ET.parse(path)
        root = tree.getroot()
        ns_cfdi = 'cfdi' if root.tag.startswith('{http://www.sat.gob.mx/cfd/4}') else 'cfdi3'
        ns = {ns_cfdi: _NS[ns_cfdi], 'tfd': _NS['tfd']}
        tag = lambda t: f'{{{_NS[ns_cfdi]}}}{t}'

        tipo     = root.get('TipoDeComprobante','')   # I=Ingreso E=Egreso P=Pago N=Nómina
        fecha_s  = root.get('Fecha','')[:10]
        total    = float(root.get('Total','0'))
        subtotal = float(root.get('SubTotal','0'))
        serie    = root.get('Serie','')
        folio    = root.get('Folio','')

        try:
            fecha = datetime.datetime.strptime(fecha_s, '%Y-%m-%d').date()
        except Exception:
            return None
        if fecha.year != year: return None

        rfc_e   = root.find(f'{tag("Emisor")}').get('Rfc','')  if root.find(f'{tag("Emisor")}') is not None else ''
        rfc_r   = root.find(f'{tag("Receptor")}').get('Rfc','') if root.find(f'{tag("Receptor")}') is not None else ''
        nom_r   = root.find(f'{tag("Receptor")}').get('Nombre','') if root.find(f'{tag("Receptor")}') is not None else ''
        nom_e   = root.find(f'{tag("Emisor")}').get('Nombre','')  if root.find(f'{tag("Emisor")}') is not None else ''

        # Timbre
        tfd = root.find(f'.//{{{_NS["tfd"]}}}TimbreFiscalDigital')
        uuid = tfd.get('UUID','') if tfd is not None else ''

        # IVA y retenciones
        iva = ret_isr = ret_iva = 0.0
        for imp in root.iter(tag('Traslado')):
            if imp.get('Impuesto','')=='002': iva += float(imp.get('Importe','0'))
        for ret in root.iter(tag('Retencion')):
            if ret.get('Impuesto','')=='001':  ret_isr += float(ret.get('Importe','0'))
            if ret.get('Impuesto','')=='002':  ret_iva += float(ret.get('Importe','0'))

        neto = round(total - ret_isr - ret_iva, 2)
        mes  = fecha.month

        # Concepto principal
        conc_node = root.find(f'.//{tag("Concepto")}')
        conc = conc_node.get('Descripcion','') if conc_node is not None else ''

        return dict(tipo=tipo, fecha=fecha_s, mes=mes,
                    serie=serie, folio=folio, folio_full=f'{serie}{folio}',
                    rfc_e=rfc_e, nom_e=nom_e, rfc_r=rfc_r, nom_r=nom_r,
                    sub=round(subtotal,2), iva=round(iva,2),
                    ret_isr=round(ret_isr,2), ret_iva=round(ret_iva,2),
                    total=round(total,2), neto=round(neto,2),
                    uuid=uuid, conc=conc,
                    estado='PENDIENTE',
                    banco_fecha='', banco_desc='', banco_monto=0.0, banco_nombre='')
    except Exception:
        return None

# ─── Amarre facturas vs banco ─────────────────────────────────────────────────
def _amarre(emitidas: list, movs_banco: list):
    """Une facturas con movimientos bancarios. Tolerancia ±$1, ventana -5..+90 días."""
    usados = set()
    for fact in emitidas:
        if fact['tipo'] not in ('I','E'): continue
        try: fdt = datetime.datetime.strptime(fact['fecha'], '%Y-%m-%d').date()
        except Exception: continue
        meta = fact['neto']
        best = None; best_score = -1
        for idx, mov in enumerate(movs_banco):
            if idx in usados: continue
            if mov['dep'] <= 0: continue
            diff_d = (mov['fecha'] - fdt).days
            if diff_d < -5 or diff_d > 90: continue
            if abs(mov['dep'] - meta) > 1.0: continue
            score = 10 - abs(diff_d)*0.1
            if fact['rfc_r'] and fact['rfc_r'] in mov['desc'].upper(): score += 5
            if score > best_score: best = idx; best_score = score
        if best is not None:
            m = movs_banco[best]; usados.add(best)
            fact['estado']      = 'LOCALIZADA'
            fact['banco_fecha'] = str(m['fecha'])
            fact['banco_desc']  = m['desc']
            fact['banco_monto'] = m['dep']
            fact['banco_nombre']= m['banco']

# ─── Motor principal ──────────────────────────────────────────────────────────
def calcular_impuestos(rfc: str, year: int,
                       carpetas_xml_emit: list = None,
                       carpetas_xml_rec:  list = None,
                       pdfs_bbva:         list = None,
                       excels_banorte:    list = None,
                       excels_inbursa:    list = None,
                       excels_extra:      dict = None,
                       log=print):
    """
    Calcula ISR/IVA mensuales para persona física régimen 612.

    Parámetros
    ----------
    rfc             : RFC del contribuyente
    year            : Año de cálculo
    carpetas_xml_emit: Rutas a carpetas con XMLs emitidos (o lista de archivos .xml)
    carpetas_xml_rec : Rutas a carpetas con XMLs recibidos
    pdfs_bbva       : Lista de archivos PDF BBVA
    excels_banorte  : Lista de archivos Excel Banorte
    excels_inbursa  : Lista de archivos Excel Inbursa
    excels_extra    : Dict {nombre_banco: [lista archivos]}

    Retorna
    -------
    (results, emitidas_log, ded_log, no_ded_log)
    """
    log(f"▶ Iniciando cálculo {rfc} — {year}")

    # 1. Recopilar XMLs
    emit_xmls = []; rec_xmls = []
    for src in (carpetas_xml_emit or []):
        if os.path.isdir(src):   emit_xmls += glob.glob(os.path.join(src,'**','*.xml'), recursive=True)
        elif src.endswith('.xml'): emit_xmls.append(src)
    for src in (carpetas_xml_rec or []):
        if os.path.isdir(src):   rec_xmls += glob.glob(os.path.join(src,'**','*.xml'), recursive=True)
        elif src.endswith('.xml'): rec_xmls.append(src)
    log(f"  XMLs emitidos: {len(emit_xmls)}  |  recibidos: {len(rec_xmls)}")

    # 2. Parsear emitidas
    emitidas = []
    for path in emit_xmls:
        r = _parse_xml(path, rfc, year)
        if r: emitidas.append(r)

    # 3. Parsear recibidas (gastos)
    recibidas = []
    for path in rec_xmls:
        r = _parse_xml(path, rfc, year)
        if r: recibidas.append(r)

    # 4. Leer bancos
    movs_banco = []
    if pdfs_bbva:       movs_banco += _leer_pdfs_bbva(pdfs_bbva, log)
    if excels_banorte:  movs_banco += _leer_excels_banco(excels_banorte, 'Banorte', log)
    if excels_inbursa:  movs_banco += _leer_excels_banco(excels_inbursa, 'Inbursa', log)
    for bnk, files in (excels_extra or {}).items():
        movs_banco += _leer_excels_banco(files, bnk, log)
    log(f"  Movimientos bancarios: {len(movs_banco)}")

    # 5. Amarre facturas vs banco
    _amarre(emitidas, movs_banco)

    # 6. Clasificar gastos recibidos
    ded_log = []; no_ded_log = []
    for g in recibidas:
        if g['tipo'] not in ('I','E'): continue
        # Solo facturas donde el RFC emisor ≠ contribuyente (son gastos)
        if g['rfc_e'].upper() == rfc.upper(): continue
        cls, motivo = clasificar_gasto(g['conc'], g['nom_e'])
        row = dict(mes=g['mes'], fecha=g['fecha'], folio=g['folio_full'],
                   nom_e=g['nom_e'], sub=g['sub'], iva=g['iva'],
                   total=g['total'], conc=g['conc'])
        if cls == 'DEDUCIBLE':
            row['iva_acred'] = g['iva']; row['nota'] = ''
            ded_log.append(row)
        else:
            row['motivo'] = motivo
            no_ded_log.append(row)

    # 7. Acumular por mes
    ing  = [0.0]*13; ded  = [0.0]*13
    ret_isr_m = [0.0]*13
    iva_caus = [0.0]*13; iva_acred = [0.0]*13; iva_ret_c = [0.0]*13

    for e in emitidas:
        if e['tipo'] not in ('I','E'): continue
        m = e['mes']
        ing[m]      += e['sub']
        iva_caus[m] += e['iva']
        ret_isr_m[m]+= e['ret_isr']
        iva_ret_c[m]+= e['ret_iva']

    for g in ded_log:
        m = g['mes']
        ded[m]      += g['sub']
        iva_acred[m]+= g['iva_acred']

    # 8. Cálculos acumulados
    results = []
    ing_acum = ded_acum = isr_pag_acum = ret_isr_acum = 0.0
    iva_fav_prev = 0.0

    for mes in range(1, 13):
        ing_acum  = round(ing_acum  + ing[mes],  2)
        ded_acum  = round(ded_acum  + ded[mes],  2)
        ret_isr_acum = round(ret_isr_acum + ret_isr_m[mes], 2)
        base_acum = max(0.0, round(ing_acum - ded_acum, 2))

        # ISR
        isr_tar   = calc_isr(base_acum)
        isr_cargo = round(isr_tar - ret_isr_acum - isr_pag_acum, 2)
        isr_pay   = max(0.0, isr_cargo)
        isr_pag_acum = round(isr_pag_acum + isr_pay, 2)
        tasa = round(isr_tar / base_acum * 100, 4) if base_acum > 0 else 0.0

        # IVA (mensual)
        iva_neto  = round(iva_caus[mes] - iva_acred[mes] - iva_ret_c[mes] - iva_fav_prev, 2)
        iva_pay   = max(0.0, iva_neto)
        iva_fav   = abs(iva_neto) if iva_neto < 0 else 0.0
        iva_fav_prev = iva_fav

        # Recargos (cero si se paga a tiempo)
        rec_isr = act_isr = tot_rec_isr = 0.0
        rec_iva = act_iva = tot_rec_iva = 0.0
        total_con_rec = round(isr_pay + iva_pay, 2)
        total_sat     = total_con_rec

        n_emit = sum(1 for e in emitidas if e['mes']==mes and e['tipo'] in ('I','E'))
        n_ded  = sum(1 for g in ded_log  if g['mes']==mes)
        n_no_ded = sum(1 for g in no_ded_log if g['mes']==mes)

        results.append(dict(
            mes=mes, nom=MESES[mes],
            ing_m=round(ing[mes],2), ded_m=round(ded[mes],2),
            ing_acum=ing_acum, ded_acum=ded_acum, base_acum=base_acum,
            isr_tar=isr_tar, tasa=tasa,
            ret_isr_m=round(ret_isr_m[mes],2), ret_isr_acum=ret_isr_acum,
            isr_pag_ant=round(isr_pag_acum - isr_pay, 2),
            isr_cargo=isr_cargo, isr_pay=isr_pay, isr_pag_acum=isr_pag_acum,
            rec_isr=rec_isr, act_isr=act_isr, total_con_rec_isr=tot_rec_isr,
            iva_caus=round(iva_caus[mes],2), iva_acred=round(iva_acred[mes],2),
            iva_ret=round(iva_ret_c[mes],2), iva_fav_ant=iva_fav_prev,
            iva_neto=iva_neto, iva_pay=iva_pay, iva_fav=iva_fav,
            iva_pag_acum=0.0,
            rec_iva=rec_iva, act_iva=act_iva, total_con_rec_iva=tot_rec_iva,
            rec_total=0.0, act_total=0.0, total_con_rec=total_con_rec,
            total_sat=total_sat,
            n_emit=n_emit, n_ded=n_ded, n_no_ded=n_no_ded,
        ))

    emitidas_log = emitidas
    log(f"✔ Cálculo completado — {sum(r['total_sat'] for r in results):,.2f} total SAT")
    return results, emitidas_log, ded_log, no_ded_log


# ─────────────────────────────────────────────────────────────────────────────
#  GENERACIÓN EXCEL
# ─────────────────────────────────────────────────────────────────────────────
def generar_excel(results, emitidas_log, ded_log, no_ded_log,
                  rfc, nombre, year, out_path, log=print):
    """Genera el libro Excel completo y lo guarda en out_path."""

    # ── Estilos ────────────────────────────────────────────────────────────────
    Az  = PatternFill('solid', fgColor='1E3A8A')
    Az2 = PatternFill('solid', fgColor='2563EB')
    Vd  = PatternFill('solid', fgColor='14532D')
    Gs  = PatternFill('solid', fgColor='374151')
    Am  = PatternFill('solid', fgColor='78350F')
    La  = PatternFill('solid', fgColor='EFF6FF')
    Lg  = PatternFill('solid', fgColor='F9FAFB')
    Lv  = PatternFill('solid', fgColor='F0FDF4')
    Lr  = PatternFill('solid', fgColor='FEF2F2')
    Blk = PatternFill('solid', fgColor='0F172A')
    ISR_PAG = PatternFill('solid', fgColor='111827')
    IVA_PAG = PatternFill('solid', fgColor='052E16')
    TOT_F   = PatternFill('solid', fgColor='431407')

    thin = Side(style='thin', color='CBD5E1')
    BD   = Border(left=thin, right=thin, top=thin, bottom=thin)
    CC   = Alignment(horizontal='center', vertical='center', wrap_text=True)
    RR   = Alignment(horizontal='right',  vertical='center')
    LL   = Alignment(horizontal='left',   vertical='center')
    MXN  = '#,##0.00'
    PCT  = '0.00%'

    def f(bold=False, color='000000', size=9, italic=False):
        return Font(name='Arial', bold=bold, color=color, size=size, italic=italic)

    def put(ws, row, col, val, bold=False, color='000000', fill=None,
            align=None, fmt=None, size=9):
        cell = ws.cell(row, col, val)
        cell.font   = f(bold, color, size)
        if fill:  cell.fill  = fill
        if align: cell.alignment = align
        if fmt:   cell.number_format = fmt
        cell.border = BD
        return cell

    def title(ws, span, row, txt, fill=None, sz=11):
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=span)
        c = ws.cell(row, 1, txt)
        c.font = f(True, 'FFFFFF', sz); c.fill = fill or Az
        c.alignment = CC; c.border = BD
        ws.row_dimensions[row].height = 22

    wb = Workbook()

    # ══ HOJA 1: PAGOS AL SAT ══════════════════════════════════════════════════
    ws = wb.active; ws.title = 'PAGOS AL SAT'
    title(ws, 14, 1, f'PAGOS PROVISIONALES AL SAT {year} — {rfc} — {nombre}', Az)
    title(ws, 14, 2, 'Régimen 612 | ISR Art.106 LISR (base devengada, deducciones autorizadas) | IVA Art.5-D LIVA', Az2, 8)
    ws.row_dimensions[2].height = 14

    for s, e, lbl, fl in [('A3','C3','INGRESOS / DED. AUTORIZADAS',Az),
                           ('D3','G3','CÁLCULO ISR — PAGO PROVISIONAL',Gs),
                           ('H3','K3','CÁLCULO IVA — PAGO MENSUAL',Vd),
                           ('L3','N3','TOTAL A ENTREGAR AL SAT',Am)]:
        ws.merge_cells(f'{s}:{e}')
        c = ws[s]; c.value=lbl; c.fill=fl; c.font=f(True,'FFFFFF',9); c.alignment=CC; c.border=BD
    ws.row_dimensions[3].height = 28

    hdrs = [('MES',Az),('INGRESOS\nMES',Az),('DED.AUTOR.\nMES',Az),
            ('BASE ACUM.\n(ing-ded)',Gs),('ISR TARIFA\nACUM.',Gs),('RET.ISR\nACUM.',Gs),('ISR\nA PAGAR',Gs),
            ('IVA\nCAUSADO',Vd),('IVA\nACREDIT.',Vd),('IVA RET.\nCLIENTE',Vd),('IVA\nA PAGAR',Vd),
            ('ISR',Am),('IVA',Am),('TOTAL SAT',Am)]
    for ci, (h, fl) in enumerate(hdrs, 1):
        c = ws.cell(4, ci, h); c.fill=fl; c.font=f(True,'FFFFFF',8); c.alignment=CC; c.border=BD
    ws.row_dimensions[4].height = 36

    for ri, r in enumerate(results, 5):
        bg = La if ri%2==0 else Lg
        vals = [(r['nom'],None),(r['ing_m'],MXN),(r['ded_m'],MXN),
                (r['base_acum'],MXN),(r['isr_tar'],MXN),(r['ret_isr_acum'],MXN),(r['isr_pay'],MXN),
                (r['iva_caus'],MXN),(r['iva_acred'],MXN),(r['iva_ret'],MXN),(r['iva_pay'],MXN),
                (r['isr_pay'],MXN),(r['iva_pay'],MXN),(r['total_sat'],MXN)]
        for ci, (v, fmt) in enumerate(vals, 1):
            bold = ci in (7,11,14)
            col  = '14532D' if (ci in (7,11) and isinstance(v,float) and v>0) else \
                   '78350F' if ci==14 and isinstance(v,float) and v>0 else '000000'
            put(ws, ri, ci, v, bold, col, bg, LL if ci==1 else RR, fmt)
        ws.row_dimensions[ri].height = 15

    tr = len(results)+5
    ws.merge_cells(f'A{tr}:C{tr}')
    put(ws, tr, 1, f'TOTAL {year}', True, 'FFFFFF', Blk, LL)
    isr_t = sum(r['isr_pay'] for r in results); iva_t = sum(r['iva_pay'] for r in results)
    tv = [None,None,None,None,None,None,isr_t,None,None,None,iva_t,isr_t,iva_t,isr_t+iva_t]
    for ci, v in enumerate(tv, 1):
        c = ws.cell(tr, ci); c.fill=Blk; c.border=BD
        if v is not None: c.value=v; c.font=f(True,'FFFFFF'); c.number_format=MXN; c.alignment=RR
    ws.row_dimensions[tr].height = 18
    for ci, w in enumerate([13,13,13,15,14,13,13,13,13,12,13,12,12,13], 1):
        ws.column_dimensions[get_column_letter(ci)].width = w
    ws.freeze_panes = 'B5'

    # ══ HOJA 2: CÁLCULO ACUMULADO — LEY ══════════════════════════════════════
    ws_ley = wb.create_sheet('CÁLCULO ACUMULADO — LEY')
    ws_ley.merge_cells('A1:R1')
    c = ws_ley['A1']
    c.value = f'CÁLCULO DE IMPUESTOS {year} — {rfc} — {nombre} — RÉG. 612'
    c.font=f(True,'FFFFFF',12); c.fill=Az; c.alignment=CC; c.border=BD
    ws_ley.row_dimensions[1].height = 26
    ws_ley.merge_cells('A2:R2')
    c = ws_ley['A2']
    c.value = 'Art. 106 LISR — base ACUMULADA ene→mes | Tarifa Art. 96 LISR | Art. 5-D LIVA — IVA mensual con saldo a favor arrastrado'
    c.font=f(False,'FFFFFF',8,True); c.fill=Az2; c.alignment=CC; c.border=BD
    ws_ley.row_dimensions[2].height = 14

    secs = [(1,1,'MES',Az),(2,6,'ISR — PAGO PROVISIONAL (Art.106 LISR)',Gs),
            (7,7,'ISR\nA PAGAR',ISR_PAG),(8,12,'IVA — PAGO MENSUAL (Art.5-D LIVA)',Vd),
            (13,13,'IVA\nA PAGAR',IVA_PAG),(14,17,'RESUMEN ANUAL',Am),(18,18,'TOTAL\nSAT',TOT_F)]
    for cs, ce, lbl, fl in secs:
        if cs == ce:
            c = ws_ley.cell(3,cs,lbl); c.fill=fl; c.font=f(True,'FFFFFF',8); c.alignment=CC; c.border=BD
        else:
            ws_ley.merge_cells(start_row=3,start_column=cs,end_row=3,end_column=ce)
            c = ws_ley.cell(3,cs,lbl); c.fill=fl; c.font=f(True,'FFFFFF',9); c.alignment=CC; c.border=BD
    ws_ley.row_dimensions[3].height = 30

    col_hdrs2 = [('ING.\nACUM.',Gs),('DED.\nACUM.',Gs),('BASE\nACUM.',Gs),('ISR s/TAR.',Gs),
                 ('RET.ISR\nACUM.',Gs),('ISR\nA PAGAR',ISR_PAG),
                 ('IVA\nCAUS.',Vd),('IVA\nACRED.',Vd),('IVA RET.',Vd),('SAL.FAV.\nANT.',Vd),('IVA\nNETO',Vd),
                 ('IVA\nA PAGAR',IVA_PAG),
                 ('ISR\nACUM.',Am),('IVA\nACUM.',Am),('SAL.FAV.\nIVA',Am),('PP ANT.\nISR',Am),('TOTAL SAT',TOT_F)]
    c2 = ws_ley.cell(4,1,'MES'); c2.fill=Az; c2.font=f(True,'FFFFFF',8); c2.alignment=CC; c2.border=BD
    for ci, (h, fl) in enumerate(col_hdrs2, 2):
        c2 = ws_ley.cell(4,ci,h); c2.fill=fl; c2.font=f(True,'FFFFFF',8); c2.alignment=CC; c2.border=BD
    ws_ley.row_dimensions[4].height = 48

    Gf = PatternFill('solid',fgColor='F8FAFC')
    ipp_d = ivp_d = fav_d = 0.0
    for ri, r in enumerate(results, 5):
        bg = La if ri%2==1 else Gf
        c2 = ws_ley.cell(ri,1,r['nom']); c2.font=f(True,'000000',9); c2.fill=bg; c2.alignment=LL; c2.border=BD
        for ci, (val, col_c) in enumerate([
            (r['ing_acum'],'1E3A8A'),(r['ded_acum'],'14532D'),(r['base_acum'],'000000'),
            (r['isr_tar'],'374151'),(r['ret_isr_acum'],'991B1B')], 2):
            c2=ws_ley.cell(ri,ci,val); c2.number_format=MXN; c2.fill=bg; c2.font=f(False,col_c,9); c2.alignment=RR; c2.border=BD
        v = r['isr_pay']
        c2 = ws_ley.cell(ri,7,v if v>0 else '—')
        if v>0: c2.number_format=MXN; c2.font=f(True,'FFFFFF',10); c2.fill=ISR_PAG
        else: c2.font=f(False,'9CA3AF',9); c2.fill=Gf
        c2.alignment=CC; c2.border=BD
        for ci, (val, col_c) in enumerate([
            (r['iva_caus'],'1E3A8A'),(r['iva_acred'],'14532D'),(r['iva_ret'],'991B1B'),
            (fav_d if fav_d>0 else 0,'78350F'),(r['iva_neto'],'000000')], 8):
            c2=ws_ley.cell(ri,ci,val); c2.number_format=MXN; c2.fill=bg; c2.font=f(False,col_c,9); c2.alignment=RR; c2.border=BD
        v = r['iva_pay']
        c2 = ws_ley.cell(ri,13,v if v>0 else '—')
        if v>0: c2.number_format=MXN; c2.font=f(True,'FFFFFF',10); c2.fill=IVA_PAG
        else: c2.font=f(False,'9CA3AF',9); c2.fill=Gf
        c2.alignment=CC; c2.border=BD
        ipp_d=round(ipp_d+r['isr_pay'],2); ivp_d=round(ivp_d+r['iva_pay'],2)
        for ci, (val,col_c) in enumerate([
            (ipp_d,'374151'),(ivp_d,'14532D'),
            (r['iva_fav'] if r['iva_fav']>0 else 0,'78350F'),(r['isr_pag_ant'],'6B7280')], 14):
            c2=ws_ley.cell(ri,ci,val); c2.number_format=MXN; c2.fill=bg; c2.font=f(False,col_c,9); c2.alignment=RR; c2.border=BD
        v = r['total_sat']
        c2 = ws_ley.cell(ri,18,v if v>0 else '—')
        if v>0: c2.number_format=MXN; c2.font=f(True,'FFFFFF',11); c2.fill=TOT_F
        else: c2.font=f(False,'9CA3AF',9); c2.fill=Gf
        c2.alignment=CC; c2.border=BD
        fav_d = r['iva_fav']
        ws_ley.row_dimensions[ri].height = 18

    ws_ley.column_dimensions['A'].width = 13
    for ci, w in enumerate([14,14,15,14,14,13,13,13,12,12,13,13,13,13,12,13,13], 2):
        ws_ley.column_dimensions[get_column_letter(ci)].width = w
    ws_ley.freeze_panes = 'B5'

    # ══ HOJA 3: CÁLCULO ISR ═══════════════════════════════════════════════════
    ws2 = wb.create_sheet('CÁLCULO ISR')
    title(ws2, 13, 1, f'CÁLCULO ISR PAGOS PROVISIONALES {year} — Art. 106 LISR', Az)
    title(ws2, 13, 2, 'Base = Ingresos acumulados − Deducciones acumuladas | ISR = Tarifa − Ret.ISR − PP anteriores', Az2, 8)
    ws2.row_dimensions[2].height = 14
    h2 = ['MES','ING.\nMES','DED.\nMES','ING.\nACUM','DED.\nACUM','BASE\nACUM',
          'ISR\nTARIFA','RET.ISR\nMES','RET.ISR\nACUM','PP\nANT.','ISR\nCARGO','ISR\nA PAGAR','TASA\nEFECT.']
    for ci, h in enumerate(h2, 1):
        c=ws2.cell(3,ci,h); c.fill=Az; c.font=f(True,'FFFFFF',8); c.alignment=CC; c.border=BD
    ws2.row_dimensions[3].height = 40
    for ri, r in enumerate(results, 4):
        bg = La if ri%2==1 else Lg
        rv = [(r['nom'],None),(r['ing_m'],MXN),(r['ded_m'],MXN),(r['ing_acum'],MXN),(r['ded_acum'],MXN),
              (r['base_acum'],MXN),(r['isr_tar'],MXN),(r['ret_isr_m'],MXN),(r['ret_isr_acum'],MXN),
              (r['isr_pag_ant'],MXN),(r['isr_cargo'],MXN),(r['isr_pay'],MXN),(r['tasa']/100,PCT)]
        for ci, (v, fmt) in enumerate(rv, 1):
            bold = ci in (11,12)
            col  = '991B1B' if ci==11 and isinstance(v,float) and v<0 else \
                   '14532D' if ci==12 and isinstance(v,float) and v>0 else '000000'
            put(ws2, ri, ci, v, bold, col, bg, LL if ci==1 else RR, fmt)
        ws2.row_dimensions[ri].height = 15
    for ci, w in enumerate([13,13,13,14,14,15,14,12,12,13,13,13,10], 1):
        ws2.column_dimensions[get_column_letter(ci)].width = w
    ws2.freeze_panes = 'B4'

    # ══ HOJA 4: CÁLCULO IVA ═══════════════════════════════════════════════════
    ws3 = wb.create_sheet('CÁLCULO IVA')
    title(ws3, 9, 1, f'CÁLCULO IVA MENSUAL {year} — Art. 5-D LIVA', Vd)
    title(ws3, 9, 2, 'IVA a pagar = IVA Causado − IVA Acreditable − IVA Retenido + Saldo a favor mes anterior', Gs, 8)
    ws3.row_dimensions[2].height = 14
    h3 = ['MES','IVA CAUSADO\n(Emit.)','SALDO FAV.\nANT.','IVA ACREDIT.\n(Ded.)','IVA RET.\nCLIENTES',
          'IVA NETO','IVA A PAGAR','SALDO A\nFAVOR','CFDIs\nE/R']
    for ci, h in enumerate(h3, 1):
        c=ws3.cell(3,ci,h); c.fill=Vd; c.font=f(True,'FFFFFF',8); c.alignment=CC; c.border=BD
    ws3.row_dimensions[3].height = 44
    iva_fav_prev2 = 0.0
    for ri, r in enumerate(results, 4):
        bg = Lv if ri%2==1 else Lg
        rv = [(r['nom'],None),(r['iva_caus'],MXN),(iva_fav_prev2,MXN),(r['iva_acred'],MXN),(r['iva_ret'],MXN),
              (r['iva_neto'],MXN),(r['iva_pay'],MXN),(r['iva_fav'],MXN),(f"{r['n_emit']}/{r['n_ded']}",None)]
        iva_fav_prev2 = r['iva_fav']
        for ci, (v, fmt) in enumerate(rv, 1):
            bold = ci in (6,7,8)
            col  = '991B1B' if ci==6 and isinstance(v,float) and v<0 else \
                   '14532D' if ci==7 and isinstance(v,float) and v>0 else \
                   '78350F' if ci==8 and isinstance(v,float) and v>0 else '000000'
            put(ws3, ri, ci, v, bold, col, bg, LL if ci in(1,9) else RR, fmt)
        ws3.row_dimensions[ri].height = 15
    for ci, w in enumerate([13,15,13,15,13,13,13,13,9], 1):
        ws3.column_dimensions[get_column_letter(ci)].width = w
    ws3.freeze_panes = 'B4'

    # ══ GASTOS DEDUCIBLES ══════════════════════════════════════════════════════
    ws4 = wb.create_sheet('GASTOS DEDUCIBLES')
    title(ws4, 7, 1, f'GASTOS DEDUCIBLES {year} — {len(ded_log)} facturas', Vd)
    hdrs4 = ['MES','FECHA','FOLIO','PROVEEDOR / EMISOR','SUBTOTAL DED.','IVA ACRED.','CONCEPTO']
    for ci, h in enumerate(hdrs4, 1):
        c=ws4.cell(3,ci,h); c.fill=Vd; c.font=f(True,'FFFFFF',8); c.alignment=CC; c.border=BD
    ws4.row_dimensions[3].height = 20
    rows4 = sorted(ded_log, key=lambda x:(x['mes'],x['fecha']))
    for ri, g in enumerate(rows4, 4):
        bg = Lv if ri%2==0 else Lg
        rv = [(MESES[g['mes']][:3],None),(g['fecha'],None),(g['folio'],None),(g['nom_e'],None),
              (g['sub'],MXN),(g['iva_acred'],MXN),(g['conc'],None)]
        for ci, (v, fmt) in enumerate(rv, 1):
            put(ws4, ri, ci, v, fill=bg, align=LL if ci in(1,2,3,4,7) else RR, fmt=fmt)
        ws4.row_dimensions[ri].height = 13
    if ded_log:
        tr4 = len(ded_log)+4
        put(ws4, tr4, 4, 'TOTAL GENERAL', True, 'FFFFFF', Blk, LL)
        put(ws4, tr4, 5, sum(g['sub'] for g in ded_log), True, 'FFFFFF', Blk, RR, MXN)
        put(ws4, tr4, 6, sum(g['iva_acred'] for g in ded_log), True, 'FFFFFF', Blk, RR, MXN)
        for ci in [1,2,3,7]:
            ws4.cell(tr4,ci).fill=Blk; ws4.cell(tr4,ci).border=BD
        ws4.row_dimensions[tr4].height = 18
    for ci, w in enumerate([9,11,12,40,14,13,50], 1):
        ws4.column_dimensions[get_column_letter(ci)].width = w

    # ══ HOJAS MENSUALES ═══════════════════════════════════════════════════════
    FILL_ING = PatternFill('solid',fgColor='1E3A8A')
    FILL_GAS = PatternFill('solid',fgColor='14532D')
    for mes in range(1,13):
        r_res = results[mes-1]
        ws_m = wb.create_sheet(MESES_C[mes])
        ws_m.merge_cells('A1:J1')
        c = ws_m['A1']
        c.value = f'{MESES[mes]} {year} — {rfc} — {nombre}'
        c.font=f(True,'FFFFFF',11); c.fill=FILL_ING; c.alignment=CC; c.border=BD
        ws_m.row_dimensions[1].height = 22
        res_data = [('Ingresos del mes',r_res['ing_m']),('Deducciones del mes',r_res['ded_m']),
                    ('IVA causado',r_res['iva_caus']),('IVA acreditable',r_res['iva_acred']),
                    ('ISR a pagar',r_res['isr_pay']),('IVA a pagar',r_res['iva_pay'])]
        for ri_s, (lbl, val) in enumerate(res_data, 2):
            c1=ws_m.cell(ri_s,1,lbl); c1.fill=La; c1.font=f(False,'000000',9); c1.alignment=LL; c1.border=BD
            c2=ws_m.cell(ri_s,2,val); c2.number_format=MXN; c2.fill=La; c2.font=f(True,'000000',9); c2.alignment=RR; c2.border=BD
            ws_m.row_dimensions[ri_s].height = 14
        mes_ing = [e for e in emitidas_log if e['mes']==mes and e['tipo'] in ('I','E')]
        mes_ded = [g for g in ded_log if g['mes']==mes]
        row_s = 9
        ws_m.merge_cells(f'A{row_s}:J{row_s}')
        c = ws_m.cell(row_s,1,f'INGRESOS ({len(mes_ing)} facturas)')
        c.font=f(True,'FFFFFF',9); c.fill=FILL_ING; c.alignment=CC; c.border=BD
        ws_m.row_dimensions[row_s].height = 16; row_s += 1
        for ci, h in enumerate(['FECHA','SERIE','FOLIO','CLIENTE','RFC','SUBTOTAL','IVA','RET.ISR','RET.IVA','TOTAL'], 1):
            c=ws_m.cell(row_s,ci,h); c.fill=FILL_ING; c.font=f(True,'FFFFFF',8); c.alignment=CC; c.border=BD
        ws_m.row_dimensions[row_s].height = 18; row_s += 1
        for ri_i, e in enumerate(mes_ing):
            bg = La if ri_i%2==0 else Lg
            rv = [(e['fecha'],None),(e['serie'],None),(e['folio'],None),(e['nom_r'],None),(e['rfc_r'],None),
                  (e['sub'],MXN),(e['iva'],MXN),(e['ret_isr'],MXN),(e['ret_iva'],MXN),(e['total'],MXN)]
            for ci, (v, fmt) in enumerate(rv, 1):
                put(ws_m, row_s, ci, v, fill=bg, align=LL if ci<=5 else RR, fmt=fmt)
            ws_m.row_dimensions[row_s].height = 13; row_s += 1
        row_s += 1
        ws_m.merge_cells(f'A{row_s}:J{row_s}')
        c = ws_m.cell(row_s,1,f'GASTOS DEDUCIBLES ({len(mes_ded)} facturas)')
        c.font=f(True,'FFFFFF',9); c.fill=FILL_GAS; c.alignment=CC; c.border=BD
        ws_m.row_dimensions[row_s].height = 16; row_s += 1
        for ci, h in enumerate(['FECHA','FOLIO','PROVEEDOR','CONCEPTO','','','SUBTOTAL','IVA ACRED.','',''], 1):
            c=ws_m.cell(row_s,ci,h); c.fill=FILL_GAS; c.font=f(True,'FFFFFF',8); c.alignment=CC; c.border=BD
        ws_m.row_dimensions[row_s].height = 18; row_s += 1
        for ri_d, g in enumerate(mes_ded):
            bg = Lv if ri_d%2==0 else Lg
            for ci, (v, fmt) in enumerate([(g['fecha'],None),(g['folio'],None),(g['nom_e'],None),(g['conc'],None),
                                            (None,None),(None,None),(g['sub'],MXN),(g['iva_acred'],MXN),(None,None),(None,None)], 1):
                if v is not None:
                    put(ws_m, row_s, ci, v, fill=bg, align=LL if ci<=4 else RR, fmt=fmt)
                else:
                    ws_m.cell(row_s,ci).fill=bg; ws_m.cell(row_s,ci).border=BD
            ws_m.row_dimensions[row_s].height = 13; row_s += 1
        for ci, w in enumerate([11,8,10,32,12,12,14,14,8,8], 1):
            ws_m.column_dimensions[get_column_letter(ci)].width = w

    # ══ GASTOS EXCLUIDOS ══════════════════════════════════════════════════════
    ws_ex = wb.create_sheet('GASTOS EXCLUIDOS')
    title(ws_ex, 8, 1, f'GASTOS NO DEDUCIBLES {year} — {len(no_ded_log)} facturas', Gs)
    title(ws_ex, 8, 2, 'Art. 28 LISR — no reducen base ISR ni generan IVA acreditable', Gs, 8)
    ws_ex.row_dimensions[2].height = 14
    hdrs_ex = ['MES','FECHA','FOLIO','PROVEEDOR','SUBTOTAL','IVA','CONCEPTO','MOTIVO']
    for ci, h in enumerate(hdrs_ex, 1):
        c=ws_ex.cell(3,ci,h); c.fill=Gs; c.font=f(True,'FFFFFF',8); c.alignment=CC; c.border=BD
    ws_ex.row_dimensions[3].height = 20
    for ri, g in enumerate(sorted(no_ded_log, key=lambda x:(x['mes'],x['fecha'])), 4):
        bg = Lr if ri%2==0 else Lg
        rv = [(MESES[g['mes']][:3],None),(g['fecha'],None),(g['folio'],None),(g['nom_e'],None),
              (g['sub'],MXN),(g['iva'],MXN),(g['conc'],None),(g['motivo'],None)]
        for ci, (v, fmt) in enumerate(rv, 1):
            put(ws_ex, ri, ci, v, fill=bg, align=LL if ci in(1,2,3,4,7,8) else RR, fmt=fmt)
        ws_ex.row_dimensions[ri].height = 13
    for ci, w in enumerate([9,11,12,38,13,12,50,20], 1):
        ws_ex.column_dimensions[get_column_letter(ci)].width = w

    # ══ AMARRE FACTURAS ═══════════════════════════════════════════════════════
    ws_am = wb.create_sheet('AMARRE FACTURAS')
    title(ws_am, 15, 1, f'AMARRE FACTURAS vs ESTADOS DE CUENTA {year} — {rfc}', Az)
    ws_am.row_dimensions[1].height = 20
    hdrs_am = ['MES','FECHA','SERIE','FOLIO','CLIENTE','SUBTOTAL','IVA TRASL.',
               'ISR RET.','IVA RET.','TOTAL CFDI','NETO COBRADO',
               'ESTADO','FECHA BANCO','MONTO BANCO','DESC BANCO']
    for ci, h in enumerate(hdrs_am, 1):
        c=ws_am.cell(2,ci,h); c.fill=Az; c.font=f(True,'FFFFFF',8); c.alignment=CC; c.border=BD
    ws_am.row_dimensions[2].height = 28
    FILL_LOC = PatternFill('solid',fgColor='D1FAE5')
    FILL_PEN = PatternFill('solid',fgColor='FEF2F2')
    for ri, e in enumerate([x for x in emitidas_log if x['tipo'] != 'P'], 3):
        bg = FILL_LOC if e['estado']=='LOCALIZADA' else FILL_PEN
        rv = [(MESES[e['mes']][:3],None),(e['fecha'],None),(e['serie'],None),(e['folio'],None),(e['nom_r'],None),
              (e['sub'],MXN),(e['iva'],MXN),(e['ret_isr'],MXN),(e['ret_iva'],MXN),(e['total'],MXN),(e['neto'],MXN),
              (e['estado'],None),(e.get('banco_fecha',''),None),(e.get('banco_monto') or '',MXN if e.get('banco_monto') else None),(e.get('banco_desc',''),None)]
        for ci, (v, fmt) in enumerate(rv, 1):
            bold = ci == 12
            col  = '14532D' if e['estado']=='LOCALIZADA' and ci==12 else \
                   '991B1B' if e['estado']=='PENDIENTE'  and ci==12 else '000000'
            put(ws_am, ri, ci, v or '', bold, col, bg, LL if ci in(1,2,3,4,5,12,13,15) else RR, fmt)
        ws_am.row_dimensions[ri].height = 13
    for ci, w in enumerate([8,11,7,10,32,13,12,11,11,13,13,11,11,14,40], 1):
        ws_am.column_dimensions[get_column_letter(ci)].width = w
    ws_am.freeze_panes = 'A3'

    # ══ TARIFA ISR ════════════════════════════════════════════════════════════
    ws5 = wb.create_sheet(f'TARIFA ISR {year}')
    title(ws5, 5, 1, f'TARIFA ISR {year} — Art. 96 LISR (Base anual)', Az)
    for ci, h in enumerate(['LÍM. INFERIOR','LÍM. SUPERIOR','CUOTA FIJA','% S/EXCEDENTE'], 1):
        c=ws5.cell(2,ci,h); c.fill=Az; c.font=f(True,'FFFFFF',8); c.alignment=CC; c.border=BD
    for ri, (li,ls,cf,pct) in enumerate(TARIFA, 3):
        bg = La if ri%2==1 else Lg
        for ci, v in enumerate([li, min(ls,9_999_999), cf, pct/100], 1):
            put(ws5, ri, ci, v, fill=bg, align=RR, fmt=MXN if ci<4 else PCT)
        ws5.row_dimensions[ri].height = 13
    for ci, w in enumerate([17,17,14,16], 1):
        ws5.column_dimensions[get_column_letter(ci)].width = w

    wb.save(out_path)
    log(f"✔ Excel guardado → {out_path}")
    return out_path
