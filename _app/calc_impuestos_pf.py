# ─────────────────────────────────────────────────────────────────────────────
#  calc_impuestos_pf.py  —  Motor de cálculo ISR/IVA Régimen 612
#  Persona Física | Art.106 LISR (ISR provisional) | Art.5-D LIVA (IVA mensual)
#  Versión 1.0  —  Oct 2026
# ─────────────────────────────────────────────────────────────────────────────

import glob, os, re, datetime, time
from xml.etree import ElementTree as ET
from openpyxl import Workbook
from openpyxl.styles import (PatternFill, Font, Alignment, Border, Side)
from openpyxl.utils import get_column_letter

try:
    import pdfplumber
    _PDFPLUMBER_OK = True
except ImportError:
    _PDFPLUMBER_OK = False

# ─── Verificación de vigencia CFDI en el SAT ─────────────────────────────────
_SAT_CACHE: dict = {}   # uuid → 'Vigente'|'Cancelado'|'No encontrado'

def verificar_cfdi_sat(uuid: str, rfc_e: str, rfc_r: str, total: float,
                       log=print) -> str:
    """
    Consulta el servicio del SAT para verificar si un CFDI está vigente o cancelado.
    Retorna: 'Vigente', 'Cancelado', 'No encontrado' o 'Error'.
    Usa caché para no repetir consultas.
    """
    if not uuid:
        return 'Sin UUID'
    key = uuid.upper()
    if key in _SAT_CACHE:
        return _SAT_CACHE[key]

    try:
        import requests
        # Servicio público del SAT — ConsultaCFDI
        url = 'https://consultaqr.facturaelectronica.sat.gob.mx/ConsultaCFDIService.svc'
        headers = {
            'Content-Type': 'text/xml; charset=utf-8',
            'SOAPAction': 'http://tempuri.org/IConsultaCFDIService/Consulta',
        }
        expr = (f'?re={rfc_e.strip()}&rr={rfc_r.strip()}'
                f'&tt={total:.6f}&id={uuid.strip()}')
        body = f"""<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"
               xmlns:tem="http://tempuri.org/">
  <soap:Body>
    <tem:Consulta>
      <tem:expresionImpresa><![CDATA[{expr}]]></tem:expresionImpresa>
    </tem:Consulta>
  </soap:Body>
</soap:Envelope>"""
        resp = requests.post(url, data=body.encode('utf-8'), headers=headers,
                             timeout=10)
        resp.raise_for_status()
        # Parsear respuesta
        root_r = ET.fromstring(resp.content)
        # Buscar CodigoEstatus y Estado
        ns_r = {'s': 'http://schemas.xmlsoap.org/soap/envelope/',
                 't': 'http://tempuri.org/'}
        estado_node = root_r.find('.//{http://tempuri.org/}Estado')
        if estado_node is not None:
            estado = estado_node.text or ''
            if 'Vigente' in estado:
                result = 'Vigente'
            elif 'Cancelado' in estado or 'Cancel' in estado:
                result = 'Cancelado'
            else:
                result = estado or 'No encontrado'
        else:
            result = 'No encontrado'
    except Exception as exc:
        result = 'Error'
        log(f"  ⚠ SAT consulta UUID {uuid[:8]}...: {exc}")

    _SAT_CACHE[key] = result
    time.sleep(0.15)   # respetar rate-limit del SAT
    return result


def verificar_lote(cfdi_list: list, log=print) -> dict:
    """
    Verifica un lote de CFDIs.
    cfdi_list: lista de dicts con uuid, rfc_e, rfc_r, total.
    Retorna dict {uuid → estado_sat}.
    """
    resultados = {}
    total_n = len(cfdi_list)
    for i, cfdi in enumerate(cfdi_list, 1):
        uuid = cfdi.get('uuid', '')
        if not uuid:
            continue
        if i % 10 == 0 or i == 1:
            log(f"  Verificando vigencia SAT: {i}/{total_n}...")
        est = verificar_cfdi_sat(
            uuid, cfdi.get('rfc_e', ''), cfdi.get('rfc_r', ''),
            cfdi.get('total', 0.0), log)
        resultados[uuid.upper()] = est
    return resultados

# ─── Constantes ───────────────────────────────────────────────────────────────
MESES = {1:'ENERO',2:'FEBRERO',3:'MARZO',4:'ABRIL',5:'MAYO',6:'JUNIO',
         7:'JULIO',8:'AGOSTO',9:'SEPTIEMBRE',10:'OCTUBRE',11:'NOVIEMBRE',12:'DICIEMBRE'}
MESES_C = {1:'ENE',2:'FEB',3:'MAR',4:'ABR',5:'MAY',6:'JUN',
            7:'JUL',8:'AGO',9:'SEP',10:'OCT',11:'NOV',12:'DIC'}

# Tarifa ISR 2024 Art.96 LISR (límite inf, límite sup, cuota fija, % excedente)
# Fuente: SAT — Art.152 LISR tabla anual 2024
TARIFA = [
    (0.01,          8_952.49,          0.00,       1.92),
    (8_952.50,     75_984.55,        171.88,       6.40),
    (75_984.56,   133_536.07,      4_461.94,      10.88),
    (133_536.08,  155_229.80,     10_723.55,      16.00),
    (155_229.81,  185_852.57,     14_194.54,      17.92),
    (185_852.58,  374_837.88,     19_682.13,      21.36),
    (374_837.89,  590_795.99,     60_049.40,      23.52),
    (590_796.00, 1_127_926.84,   110_842.74,      30.00),
    (1_127_926.85,1_503_902.46,  271_981.99,      32.00),
    (1_503_902.47,4_511_707.37,  392_294.17,      34.00),
    (4_511_707.38,999_999_999, 1_414_947.85,      35.00),
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
NO_DED_PROV   = ['UBER','DIDI',
                  'WAL MART','WALMART','CHEDRAUI',
                  'MASCOTAS','PET ',
                  'RESTAURANTE','RESTAURANT','TACO ','MARISCOS','ASADOR','GASTRONOMIC',
                  'CEVICHI','MANTALA','BEER FACTORY','LIKOREAK','LA CALLE DEL VINO',
                  'SILENT BEATS',
                  'STAR MEDICA','ULTRASONIDO','MASTOGRAFIA',
                  'CLUB RECREATIVO','CLUB DEPORTIVO',
                  'OPERADORA HOTELERA','SERVICIOS AEROPORTUARIOS',
                  'OXXO','AEROCOMIDA']
NO_DED_CONC   = ['MULTA','RECARGO','SANCION','SANCIÓN','DONATIVO','DONACION','DONACIÓN',
                  'CLAUSULA PENAL','CLÁUSULA PENAL',
                  'HOSPEDAJE']
# Gastos con deducción parcial 8.5% — Art. 28 fr. XX LISR
PARCIAL_DED_CONC   = ['CONSUMO','RESTAURANTE','RESTAURANT','ALIMENTO','COMIDA','VIATICO']
PARCIAL_DED_FACTOR = 0.085
NO_DED_MED_PROV = ['FARMACIA','DROGUE','MEDICAMENTO','SALUD','PHARMA']
NO_DED_MED_CONC = ['MEDICAMENTO','MEDICINA','CONSULTA MEDICA','ANÁLISIS CLINICO','LABORATORIO']

def clasificar_gasto(conc: str, nom_e: str = '') -> tuple:
    """Regresa ('DEDUCIBLE','',1.0) | ('PARCIAL_DED', motivo, 0.085) | ('NO_DED', motivo, 0.0)."""
    c = (conc or '').upper(); n = (nom_e or '').upper()
    for kw in PARCIAL_DED_CONC:
        if kw in c or kw in n:
            return 'PARCIAL_DED', f'Consumo/restaurante 8.5% Art.28 fr.XX LISR', PARCIAL_DED_FACTOR
    for kw in NO_DED_CONC:
        if kw in c: return 'NO_DED', f'Concepto no deducible: {kw}', 0.0
    for kw in NO_DED_PROV:
        if kw in n or kw in c: return 'NO_DED', f'Proveedor/concepto no autorizado: {kw}', 0.0
    for kw in NO_DED_MED_PROV:
        if kw in n: return 'NO_DED', f'Farmacia/salud (uso personal): {kw}', 0.0
    for kw in NO_DED_MED_CONC:
        if kw in c: return 'NO_DED', f'Gastos médicos personales: {kw}', 0.0
    return 'DEDUCIBLE', '', 1.0

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
# Soporta formato "DD-MMM-YY" (ej. 02-ENE-24) del estado de cuenta BBVA actual.
_MESES_ES = {'ENE':1,'FEB':2,'MAR':3,'ABR':4,'MAY':5,'JUN':6,
             'JUL':7,'AGO':8,'SEP':9,'OCT':10,'NOV':11,'DIC':12}
_DATE_BBVA = re.compile(
    r'^(\d{2})-(ENE|FEB|MAR|ABR|MAY|JUN|JUL|AGO|SEP|OCT|NOV|DIC)-(\d{2,4})',
    re.IGNORECASE)
_AMT_PAT   = re.compile(r'([\d,]+\.\d{2})')

def _leer_pdfs_bbva(pdfs: list, log=print) -> list:
    """
    Extrae movimientos de estados de cuenta BBVA (PDF).
    Soporta formato DD-MMM-YY (ej. 02-ENE-24) y DD/MM/YYYY.
    Retorna lista de dicts: fecha, desc, dep, ret, banco.
    """
    if not _PDFPLUMBER_OK:
        log("⚠ pdfplumber no instalado — PDFs BBVA omitidos"); return []
    movs = []
    for path in pdfs:
        try:
            year_pdf   = None
            in_det     = False
            prev_saldo = None
            with pdfplumber.open(path) as pdf:
                for page in pdf.pages:
                    txt = page.extract_text() or ''
                    # Detectar año del periodo (ej. "Periodo Del 01/Enero/2024")
                    m_yr = re.search(r'[Pp]eriodo\s+[Dd]el\s+\d+/\w+/(\d{4})', txt)
                    if m_yr:
                        year_pdf = int(m_yr.group(1))
                    for line in txt.splitlines():
                        if 'DETALLE DE MOVIMIENTOS' in line.upper():
                            in_det = True
                            continue
                        if not in_det:
                            continue
                        # Línea de saldo anterior
                        if 'SALDO ANTERIOR' in line.upper():
                            amts = _AMT_PAT.findall(line)
                            if amts:
                                prev_saldo = float(amts[-1].replace(',', ''))
                            continue
                        # --- Intento 1: formato DD-MMM-YY (BBVA actual) ---
                        dm = _DATE_BBVA.match(line)
                        if dm:
                            day_s, mon_s, yr_s = dm.groups()
                            mon = _MESES_ES.get(mon_s.upper(), 0)
                            if not mon:
                                continue
                            yr = year_pdf if year_pdf else (
                                int(yr_s) if len(yr_s) == 4 else 2000 + int(yr_s))
                            try:
                                fecha = datetime.date(yr, mon, int(day_s))
                            except ValueError:
                                continue
                            # Extraer todos los montos de la línea
                            amts = _AMT_PAT.findall(line)
                            if len(amts) < 2:
                                continue
                            saldo = float(amts[-1].replace(',', ''))
                            monto = float(amts[-2].replace(',', ''))
                            # Determinar dep vs ret por cambio en saldo
                            dep = ret = 0.0
                            if prev_saldo is not None:
                                diff = round(saldo - prev_saldo, 2)
                                if abs(diff - monto) < 2.0:
                                    dep = monto
                                elif abs(diff + monto) < 2.0:
                                    ret = monto
                                else:
                                    # fallback: signo del cambio
                                    if diff > 0:
                                        dep = monto
                                    else:
                                        ret = monto
                            else:
                                prev_saldo = saldo
                                continue
                            prev_saldo = saldo
                            # Descripción: texto entre la fecha y los montos
                            after_date = line[dm.end():]
                            desc_part = re.sub(r'\s*[\d,]+\.\d{2}.*$', '', after_date).strip()
                            if not desc_part:
                                continue
                            movs.append({'fecha': fecha, 'desc': desc_part,
                                         'dep': dep, 'ret': ret, 'banco': 'BBVA'})
                            continue
                        # --- Intento 2: formato DD/MM/YYYY (legacy) ---
                        m2 = re.match(
                            r'(\d{2}/\d{2}/\d{4})\s+(.+?)\s+([\d,]+\.\d{2})?\s*([\d,]+\.\d{2})?$',
                            line.strip())
                        if m2:
                            fecha_s, desc, dep_s, ret_s = m2.groups()
                            try:
                                fecha = datetime.datetime.strptime(fecha_s, '%d/%m/%Y').date()
                            except Exception:
                                continue
                            dep2 = float((dep_s or '0').replace(',', ''))
                            ret2 = float((ret_s or '0').replace(',', ''))
                            movs.append({'fecha': fecha, 'desc': desc.strip(),
                                         'dep': dep2, 'ret': ret2, 'banco': 'BBVA'})
        except Exception as e:
            log(f"  ✗ PDF BBVA {os.path.basename(path)}: {e}")
    log(f"  BBVA: {len(movs)} movimientos extraídos de {len(pdfs)} PDFs")
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

        tipo         = root.get('TipoDeComprobante','')   # I=Ingreso E=Egreso P=Pago N=Nómina
        fecha_s      = root.get('Fecha','')[:10]
        total        = float(root.get('Total','0'))
        subtotal     = float(root.get('SubTotal','0'))
        serie        = root.get('Serie','')
        folio        = root.get('Folio','')
        metodo_pago  = root.get('MetodoPago','')   # PUE=una exhibición / PPD=parcialidades
        forma_pago   = root.get('FormaPago','')    # 01=Efectivo 03=TRF 04=Tarjeta etc.

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

        # IVA y retenciones — solo desde el nodo raíz cfdi:Impuestos
        # (evita doble conteo: CFDI 4.0 repite los nodos en Concepto Y en raíz)
        iva = ret_isr = ret_iva = 0.0
        imp_root = root.find(tag('Impuestos'))
        if imp_root is not None:
            for tr in imp_root.findall(f'.//{tag("Traslado")}'):
                if tr.get('Impuesto','') == '002':
                    iva += float(tr.get('Importe','0'))
            for rt in imp_root.findall(f'.//{tag("Retencion")}'):
                if rt.get('Impuesto','') == '001': ret_isr += float(rt.get('Importe','0'))
                if rt.get('Impuesto','') == '002': ret_iva += float(rt.get('Importe','0'))

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
                    metodo_pago=metodo_pago, forma_pago=forma_pago,
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
        meta = fact['total']
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
                       verificar_vigencia: bool = True,
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

    # 5. Verificar vigencia SAT (opcional)
    if verificar_vigencia and emitidas:
        log(f"  Verificando vigencia de {len(emitidas)} CFDIs emitidos en el SAT...")
        estados_sat = verificar_lote(emitidas, log)
        cancelados = 0
        for e in emitidas:
            sat_est = estados_sat.get((e.get('uuid') or '').upper(), 'Sin verificar')
            e['sat_estado'] = sat_est
            if sat_est == 'Cancelado':
                e['cancelado'] = True
                cancelados += 1
            else:
                e['cancelado'] = False
        log(f"  ✔ Vigentes: {len(emitidas)-cancelados} | Cancelados: {cancelados}")
    else:
        for e in emitidas:
            e['sat_estado'] = 'Sin verificar'
            e['cancelado'] = False

    # Excluir cancelados de los cálculos (pero mantenerlos en el log para el Excel)
    emitidas_vigentes = [e for e in emitidas if not e.get('cancelado', False)]

    # 5b. Amarre facturas vs banco (solo vigentes)
    _amarre(emitidas_vigentes, movs_banco)
    # Marcar los cancelados también en el log completo
    for e in emitidas:
        if e.get('cancelado'):
            e['estado'] = 'CANCELADO'

    # 6. Clasificar gastos recibidos
    ded_log = []; no_ded_log = []
    for g in recibidas:
        if g['tipo'] not in ('I','E'): continue
        # Solo facturas donde el RFC emisor ≠ contribuyente (son gastos)
        if g['rfc_e'].upper() == rfc.upper(): continue
        cls, motivo, factor = clasificar_gasto(g['conc'], g['nom_e'])
        row = dict(mes=g['mes'], fecha=g['fecha'], folio=g['folio_full'],
                   nom_e=g['nom_e'], sub=g['sub'], iva=g['iva'],
                   total=g['total'], conc=g['conc'],
                   metodo_pago=g.get('metodo_pago',''), forma_pago=g.get('forma_pago',''))

        # ── Reglas de deducibilidad por forma de pago ──────────────────────────
        if cls in ('DEDUCIBLE', 'PARCIAL_DED'):
            _fp   = g.get('forma_pago', '')
            _conc = (g.get('conc','') + ' ' + g.get('nom_e','')).upper()
            _es_combustible = any(k in _conc for k in
                                  ('GASOLINA','DIESEL','COMBUSTIBLE','GAS NATURAL','GAS L.P.'))
            if _fp == '01':  # Efectivo
                if _es_combustible:
                    # Art. 28 fr. V LISR — combustible en efectivo → no deducible
                    cls    = 'NO_DED'
                    motivo = 'Combustible pagado en EFECTIVO (Art.28 fr.V LISR — no deducible)'
                elif g['sub'] > 2000:
                    # Art. 27 fr. III LISR — pago en efectivo > $2,000 → no deducible
                    cls    = 'NO_DED'
                    motivo = f'Efectivo >${g["sub"]:,.2f} excede $2,000 (Art.27 fr.III LISR — no deducible)'
        # ───────────────────────────────────────────────────────────────────────

        if cls in ('DEDUCIBLE', 'PARCIAL_DED'):
            sub_ded = round(g['sub'] * factor, 2)
            iva_ded = round(g['iva'] * factor, 2)
            row['iva_acred']    = iva_ded
            row['sub_original'] = g['sub']
            row['factor']       = factor
            row['sub_ded']      = sub_ded
            row['nota']         = motivo  # '' si DEDUCIBLE, descripción si PARCIAL_DED
            ded_log.append(row)
        else:  # NO_DED
            row['motivo'] = motivo
            no_ded_log.append(row)

    # 7. Acumular por mes
    ing  = [0.0]*13; ded  = [0.0]*13
    ret_isr_m = [0.0]*13
    iva_caus = [0.0]*13; iva_acred = [0.0]*13; iva_ret_c = [0.0]*13

    for e in emitidas_vigentes:
        if e['tipo'] not in ('I','E'): continue
        m = e['mes']
        ing[m]      += e['sub']
        iva_caus[m] += e['iva']
        ret_isr_m[m]+= e['ret_isr']
        iva_ret_c[m]+= e['ret_iva']

    for g in ded_log:
        m = g['mes']
        ded[m]      += g.get('sub_ded', g['sub'])  # sub_ded usa factor 8.5% si PARCIAL_DED
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

        # IVA (mensual) — Art. 5-D LIVA
        # IVA_neto = IVA_causado − IVA_acreditable − IVA_retenido − Saldo_a_favor_anterior
        # Si > 0 → se paga; si < 0 → nuevo saldo a favor que se arrastra al mes siguiente
        iva_fav_ant_save = iva_fav_prev          # saldo a favor del mes anterior (para display)
        iva_neto  = round(iva_caus[mes] - iva_acred[mes] - iva_ret_c[mes] - iva_fav_prev, 2)
        iva_pay   = max(0.0, iva_neto)
        iva_fav   = abs(iva_neto) if iva_neto < 0 else 0.0
        iva_fav_prev = iva_fav

        # Recargos y actualización (Art. 21 / Art. 17-A CFF) — pagos extemporáneos
        rec_isr, act_isr, tot_rec_isr = calc_rec_act(isr_pay, mes, year)
        rec_iva, act_iva, tot_rec_iva = calc_rec_act(iva_pay, mes, year)
        total_con_rec = round(tot_rec_isr + tot_rec_iva, 2)
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
            iva_ret=round(iva_ret_c[mes],2), iva_fav_ant=iva_fav_ant_save,
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
    title(ws3, 9, 2, 'IVA a pagar = IVA Causado − IVA Acreditable − IVA Retenido − Saldo a favor mes anterior  (Art. 5-D LIVA)', Gs, 8)
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



    # ══ CÉDULA DE CÁLCULO (vertical: conceptos x meses — con fórmulas Excel) ═
    # Mapa de hojas referenciadas:
    #   'CÁLCULO ISR':  row4=ENE..row15=DIC
    #     B=ing_m  C=ded_m  D=ing_acum  E=ded_acum  F=base_acum
    #     G=isr_tar  H=ret_isr_m  I=ret_isr_acum  J=pp_ant  K=isr_cargo
    #     L=isr_pay  M=tasa
    #   'CÁLCULO IVA':  row4=ENE..row15=DIC
    #     B=iva_caus  C=saldo_fav_ant  D=iva_acred  E=iva_ret
    #     F=iva_neto  G=iva_pay  H=saldo_a_favor
    # Columnas CÉDULA: B=ENE(ci=2)..M=DIC(ci=13)  → isr/iva_row = ci+2

    ws_ced = wb.create_sheet('CÉDULA DE CÁLCULO')
    MCOLS_CED  = 13   # última col de mes (M=DIC)
    MCOLS_TOT  = 14   # col N = TOTAL
    ws_ced.merge_cells(f'A1:{get_column_letter(MCOLS_TOT)}1')
    c = ws_ced['A1']
    c.value = f'CÉDULA DE CÁLCULO DE IMPUESTOS {year} — {rfc} — {nombre} — RÉG. 612'
    c.font=f(True,'FFFFFF',12); c.fill=Az; c.alignment=CC; c.border=BD
    ws_ced.row_dimensions[1].height = 26
    ws_ced.merge_cells(f'A2:{get_column_letter(MCOLS_TOT)}2')
    c = ws_ced['A2']
    c.value = ('Art. 106 LISR — Pagos provisionales ISR acumulado '
               '| Art. 5-D LIVA — IVA mensual '
               '| Art. 21 CFF — Recargos y actualización')
    c.font=f(False,'FFFFFF',8,True); c.fill=Az2; c.alignment=CC; c.border=BD
    ws_ced.row_dimensions[2].height = 14

    _MESES_CORTOS = ['ENE','FEB','MAR','ABR','MAY','JUN','JUL','AGO','SEP','OCT','NOV','DIC']
    ch0 = ws_ced.cell(3,1,'CONCEPTO')
    ch0.fill=PatternFill('solid',fgColor='374151'); ch0.font=f(True,'FFFFFF',9)
    ch0.alignment=CC; ch0.border=BD
    for _ci, _mn in enumerate(_MESES_CORTOS, 2):
        _ch = ws_ced.cell(3,_ci,_mn)
        _ch.fill=Az; _ch.font=f(True,'FFFFFF',9); _ch.alignment=CC; _ch.border=BD
    # Encabezado columna TOTAL (col N)
    _ct = ws_ced.cell(3, MCOLS_TOT, 'TOTAL')
    _ct.fill=PatternFill('solid',fgColor='1E3A8A'); _ct.font=f(True,'FFFFFF',9)
    _ct.alignment=CC; _ct.border=BD
    ws_ced.row_dimensions[3].height = 22

    _LBL_ISR = PatternFill('solid', fgColor='EFF6FF')
    _LBL_IVA = PatternFill('solid', fgColor='F0FDF4')
    _LBL_REC = PatternFill('solid', fgColor='F5F3FF')
    _TOT_ISR  = PatternFill('solid', fgColor='374151')

    # ── Factores INPC/recargos por mes (para incrustar en fórmulas) ──────────
    _REC_RATE_CED  = RECARGO_RATES.get(year, 0.0147)
    _INPC_REF_CED  = INPC_REF                            # 145.75

    def _inpc_m(mes):
        return INPC_TABLE.get((year, mes), _INPC_REF_CED)

    # Filas clave en la CÉDULA (row_i = index_en_CEDULA_ROWS + 4)
    # Índice 0=sec_ISR(row4) 15=ISR_A_PAGAR(row19) 24=IVA_A_PAGAR(row28)
    # 28=ACT_ISR(row32) 29=REC_ISR(row33) 30=ISR_CR(row34)
    # 31=ACT_IVA(row35) 32=REC_IVA(row36) 33=IVA_CR(row37) 34=TOTAL(row38)
    _R_ISR_PAY = 19   # ISR A PAGAR AL SAT
    _R_IVA_PAY = 28   # IVA A PAGAR AL SAT
    _R_ACT_ISR = 32   # Actualización ISR
    _R_REC_ISR = 33   # Recargos ISR
    _R_ISR_CR  = 34   # ISR CON ACTUALIZ. Y RECARGOS
    _R_ACT_IVA = 35   # Actualización IVA
    _R_REC_IVA = 36   # Recargos IVA
    _R_IVA_CR  = 37   # IVA CON ACTUALIZ. Y RECARGOS

    # (label, fill_lbl, is_sec_hdr, is_total, fmt_key, formula_fn, total_mode)
    # total_mode: 'SUM'=suma B:M | 'DIC'=valor DIC (acumulado) | 'AVG'=promedio | None=vacío
    _CEDULA_ROWS = [
        # ── Sección ISR ──────────────────────────────────────────────────────
        ('ISR — PAGO PROVISIONAL MENSUAL  (Art. 106 LISR)',
            Az, True, False, None, None, None),                                 # row 4
        ('(+)  Ingresos del mes',
            _LBL_ISR, False, False, 'MXN',
            lambda ci: f"='CÁLCULO ISR'!B{ci+2}", 'SUM'),                      # row 5
        ('(+)  Acum. ingresos meses anteriores',
            _LBL_ISR, False, False, 'MXN',
            lambda ci: f"='CÁLCULO ISR'!D{ci+2}-'CÁLCULO ISR'!B{ci+2}", None), # row 6
        ('(=)  Ingresos acumulados ene→mes',
            _TOT_ISR, False, True, 'MXN',
            lambda ci: f"={get_column_letter(ci)}5+{get_column_letter(ci)}6", 'DIC'), # row 7
        ('(--)  Deducciones autorizadas del mes',
            _LBL_ISR, False, False, 'MXN',
            lambda ci: f"='CÁLCULO ISR'!C{ci+2}", 'SUM'),                      # row 8
        ('(--)  Acum. deducciones meses anteriores',
            _LBL_ISR, False, False, 'MXN',
            lambda ci: f"='CÁLCULO ISR'!E{ci+2}-'CÁLCULO ISR'!C{ci+2}", None), # row 9
        ('(=)  Deducciones acumuladas ene→mes',
            Vd, False, True, 'MXN',
            lambda ci: f"={get_column_letter(ci)}8+{get_column_letter(ci)}9", 'DIC'), # row 10
        ('(=)  BASE GRAVABLE ACUMULADA',
            Az, False, True, 'MXN',
            lambda ci: f"=MAX(0,{get_column_letter(ci)}7-{get_column_letter(ci)}10)", 'DIC'), # row 11
        ('(\xd7)  ISR segun tarifa Art. 96 LISR',
            _LBL_ISR, False, False, 'MXN',
            lambda ci: f"='CÁLCULO ISR'!G{ci+2}", 'DIC'),                      # row 12
        ('       Tasa efectiva ISR',
            _LBL_ISR, False, False, 'PCT',
            lambda ci: f"='CÁLCULO ISR'!M{ci+2}", 'DIC'),                      # row 13
        ('(--)  Retenciones ISR del mes',
            _LBL_ISR, False, False, 'MXN',
            lambda ci: f"='CÁLCULO ISR'!H{ci+2}", 'SUM'),                      # row 14
        ('(--)  Acum. retenciones meses anteriores',
            _LBL_ISR, False, False, 'MXN',
            lambda ci: f"='CÁLCULO ISR'!I{ci+2}-'CÁLCULO ISR'!H{ci+2}", None), # row 15
        ('(=)  Retenciones ISR acumuladas ene→mes',
            _TOT_ISR, False, True, 'MXN',
            lambda ci: f"={get_column_letter(ci)}14+{get_column_letter(ci)}15", 'DIC'), # row 16
        ('(--)  Pagos provisionales anteriores',
            _LBL_ISR, False, False, 'MXN',
            lambda ci: f"='CÁLCULO ISR'!J{ci+2}", None),                       # row 17
        ('(=)  ISR CARGO / (FAVOR)',
            _TOT_ISR, False, True, 'MXN',
            lambda ci: (f"={get_column_letter(ci)}12"
                        f"-{get_column_letter(ci)}16"
                        f"-{get_column_letter(ci)}17"), 'SUM'),                 # row 18
        ('(=)  ISR A PAGAR AL SAT',
            ISR_PAG, False, True, 'MXN',
            lambda ci: f"=MAX(0,{get_column_letter(ci)}18)", 'SUM'),            # row 19
        ('(+)  ISR pagado acumulado año',
            _LBL_ISR, False, False, 'MXN',
            lambda ci: f"=SUM('CÁLCULO ISR'!$L$4:L{ci+2})", 'DIC'),           # row 20
        # blank                                                                  # row 21
        (None, None, False, False, None, None, None),
        # ── Sección IVA ──────────────────────────────────────────────────────
        ('IVA — PAGO MENSUAL  (Art. 5-D LIVA)',
            Vd, True, False, None, None, None),                                 # row 22
        ('(+)  IVA causado del mes',
            _LBL_IVA, False, False, 'MXN',
            lambda ci: f"='CÁLCULO IVA'!B{ci+2}", 'SUM'),                      # row 23
        ('(--)  IVA acreditable del mes',
            _LBL_IVA, False, False, 'MXN',
            lambda ci: f"='CÁLCULO IVA'!D{ci+2}", 'SUM'),                      # row 24
        ('(--)  IVA retenido por clientes',
            _LBL_IVA, False, False, 'MXN',
            lambda ci: f"='CÁLCULO IVA'!E{ci+2}", 'SUM'),                      # row 25
        ('(+)  Saldo a favor IVA mes anterior',
            _LBL_IVA, False, False, 'MXN',
            lambda ci: (
                f"='CÁLCULO IVA'!C4"
                if ci == 2 else
                f"=IF({get_column_letter(ci-1)}27<0,ABS({get_column_letter(ci-1)}27),0)"
            ), None),                                                            # row 26
        ('(=)  IVA NETO DEL MES',
            Vd, False, True, 'MXN',
            lambda ci: (f"={get_column_letter(ci)}23"
                        f"-{get_column_letter(ci)}24"
                        f"-{get_column_letter(ci)}25"
                        f"-{get_column_letter(ci)}26"), 'SUM'),                 # row 27
        ('(=)  IVA A PAGAR AL SAT',
            IVA_PAG, False, True, 'MXN',
            lambda ci: f"=MAX(0,{get_column_letter(ci)}27)", 'SUM'),            # row 28
        ('       Saldo a favor → sig. mes',
            _LBL_IVA, False, False, 'MXN',
            lambda ci: f"=IF({get_column_letter(ci)}27<0,ABS({get_column_letter(ci)}27),0)", 'DIC'), # row 29
        # blank                                                                  # row 30
        (None, None, False, False, None, None, None),
        # ── Sección Recargos ─────────────────────────────────────────────────
        ('RECARGOS Y ACTUALIZACION  (Art. 21 CFF)',
            PatternFill('solid',fgColor='4C1D95'), True, False, None, None, None), # row 31
        (f'(+)  Actualizacion ISR  [INPC ref {_INPC_REF_CED} ÷ INPC {year}/mes]',
            _LBL_REC, False, False, 'MXN',
            lambda ci: (
                f"={get_column_letter(ci)}{_R_ISR_PAY}"
                f"*MAX(0,{_INPC_REF_CED}/{_inpc_m(ci-1):.3f}-1)"
            ), 'SUM'),                                                           # row 32
        (f'(+)  Recargos ISR  [{year}: {_REC_RATE_CED*100:.2f}%/mes × meses de retraso]',
            _LBL_REC, False, False, 'MXN',
            lambda ci: (
                f"={get_column_letter(ci)}{_R_ISR_PAY}"
                f"*{_REC_RATE_CED}*MAX(1,{14-ci})"
            ), 'SUM'),                                                           # row 33
        ('(=)  ISR CON ACTUALIZ. Y RECARGOS',
            ISR_PAG, False, True, 'MXN',
            lambda ci: (
                f"={get_column_letter(ci)}{_R_ISR_PAY}"
                f"+{get_column_letter(ci)}{_R_ACT_ISR}"
                f"+{get_column_letter(ci)}{_R_REC_ISR}"
            ), 'SUM'),                                                           # row 34
        (f'(+)  Actualizacion IVA  [INPC ref {_INPC_REF_CED} ÷ INPC {year}/mes]',
            _LBL_REC, False, False, 'MXN',
            lambda ci: (
                f"={get_column_letter(ci)}{_R_IVA_PAY}"
                f"*MAX(0,{_INPC_REF_CED}/{_inpc_m(ci-1):.3f}-1)"
            ), 'SUM'),                                                           # row 35
        (f'(+)  Recargos IVA  [{year}: {_REC_RATE_CED*100:.2f}%/mes × meses de retraso]',
            _LBL_REC, False, False, 'MXN',
            lambda ci: (
                f"={get_column_letter(ci)}{_R_IVA_PAY}"
                f"*{_REC_RATE_CED}*MAX(1,{14-ci})"
            ), 'SUM'),                                                           # row 36
        ('(=)  IVA CON ACTUALIZ. Y RECARGOS',
            IVA_PAG, False, True, 'MXN',
            lambda ci: (
                f"={get_column_letter(ci)}{_R_IVA_PAY}"
                f"+{get_column_letter(ci)}{_R_ACT_IVA}"
                f"+{get_column_letter(ci)}{_R_REC_IVA}"
            ), 'SUM'),                                                           # row 37
        ('(=)  TOTAL CON RECARGOS Y ACTUALIZ.',
            TOT_F, False, True, 'MXN',
            lambda ci: (
                f"={get_column_letter(ci)}{_R_ISR_CR}"
                f"+{get_column_letter(ci)}{_R_IVA_CR}"
            ), 'SUM'),                                                           # row 38
    ]

    _TOT_COL = get_column_letter(MCOLS_TOT)   # 'N'

    for _ri_off, (_lbl, _lfill, _is_sec, _is_tot, _fmt_key, _ffunc, _tot_mode) in enumerate(_CEDULA_ROWS):
        _row_i = _ri_off + 4
        _fmt = MXN if _fmt_key == 'MXN' else (PCT if _fmt_key == 'PCT' else None)

        if _lbl is None:                              # separador en blanco
            ws_ced.row_dimensions[_row_i].height = 5
            for _ci_b in range(1, MCOLS_TOT+1):
                ws_ced.cell(_row_i, _ci_b).border = BD
            continue

        if _is_sec:                                   # encabezado de sección
            ws_ced.merge_cells(start_row=_row_i, start_column=1,
                               end_row=_row_i, end_column=MCOLS_TOT)
            _c = ws_ced.cell(_row_i, 1, _lbl)
            _c.fill=_lfill; _c.font=f(True,'FFFFFF',9); _c.alignment=CC; _c.border=BD
            ws_ced.row_dimensions[_row_i].height = 18
            continue

        # Celda etiqueta
        _lc = ws_ced.cell(_row_i, 1, _lbl)
        _lc.fill=_lfill; _lc.border=BD
        if _is_tot:
            _lc.font=f(True,'FFFFFF',9); _lc.alignment=LL
        else:
            _lc.font=f(False,'1E293B',9); _lc.alignment=LL

        # Celdas de fórmula: una por mes (B→M)
        for _ci2 in range(2, MCOLS_CED+1):
            _formula = _ffunc(_ci2) if _ffunc else None
            _dc = ws_ced.cell(_row_i, _ci2, _formula)
            _dc.fill=_lfill; _dc.border=BD
            if _fmt:
                _dc.number_format = _fmt
            _dc.font = f(True,'FFFFFF',10) if _is_tot else f(False,'1E293B',9)
            _dc.alignment = RR

        # Columna TOTAL (col N)
        _tc = ws_ced.cell(_row_i, MCOLS_TOT)
        if _tot_mode == 'SUM':
            _tc.value = f"=SUM(B{_row_i}:M{_row_i})"
        elif _tot_mode == 'DIC':
            _tc.value = f"=M{_row_i}"          # DIC ya es el acumulado anual
        elif _tot_mode == 'AVG':
            _tc.value = f"=IFERROR(AVERAGEIF(B{_row_i}:M{_row_i},\"<>0\"),0)"
        # else None → celda vacía
        _tc.fill = PatternFill('solid', fgColor='1E3A8A') if _is_tot else \
                   PatternFill('solid', fgColor='DBEAFE')
        _tc.border = BD
        if _fmt:
            _tc.number_format = _fmt
        _tc.font  = f(True,'FFFFFF',10) if _is_tot else f(False,'1E3A8A',9)
        _tc.alignment = RR

        ws_ced.row_dimensions[_row_i].height = 16 if _is_tot else 14

    ws_ced.column_dimensions['A'].width = 55
    for _ci3 in range(2, MCOLS_CED+1):
        ws_ced.column_dimensions[get_column_letter(_ci3)].width = 13
    ws_ced.column_dimensions[_TOT_COL].width = 15   # columna TOTAL un poco más ancha
    ws_ced.freeze_panes = 'B4'


    # ══ GASTOS DEDUCIBLES (con subtotales por mes) ════════════════════════════
    ws4 = wb.create_sheet('GASTOS DEDUCIBLES')
    _GD_COLS = 8
    title(ws4, _GD_COLS, 1, f'GASTOS DEDUCIBLES {year} — {len(ded_log)} facturas', Vd)
    title(ws4, _GD_COLS, 2,
          'Art. 103 LISR | Deducciones autorizadas Regimen 612 | Consumo/restaurante al 8.5% Art.28 fr.XX',
          Vd, 8)
    ws4.row_dimensions[2].height = 14
    _hdrs4 = ['MES','FECHA','FOLIO','PROVEEDOR / EMISOR','SUBTOTAL DED.','IVA ACRED.','CONCEPTO / NOTA','MÉTODO/FORMA PAGO']
    for _ci4h, _h4 in enumerate(_hdrs4, 1):
        _c4=ws4.cell(3,_ci4h,_h4)
        _c4.fill=Vd; _c4.font=f(True,'FFFFFF',8); _c4.alignment=CC; _c4.border=BD
    ws4.row_dimensions[3].height = 20

    def _gd_subtotal(ws, row_n, mes_num, mes_buf):
        _sub_t = sum(g.get('sub_ded', g['sub']) for g in mes_buf)
        _iva_t = sum(g['iva_acred'] for g in mes_buf)
        _lbl_m = MESES[mes_num][:3]
        _sub_lbl = f'Subtotal {_lbl_m} ({len(mes_buf)} facturas)'
        _c = ws.cell(row_n, 1, _sub_lbl)
        _c.fill=Vd; _c.font=f(True,'FFFFFF',8); _c.alignment=LL; _c.border=BD
        for _ci_st in [2,3,4]:
            ws.cell(row_n,_ci_st).fill=Vd; ws.cell(row_n,_ci_st).border=BD
        _c5=ws.cell(row_n,5,_sub_t)
        _c5.fill=Vd; _c5.font=f(True,'FFFFFF',9); _c5.number_format=MXN; _c5.alignment=RR; _c5.border=BD
        _c6=ws.cell(row_n,6,_iva_t)
        _c6.fill=Vd; _c6.font=f(True,'FFFFFF',9); _c6.number_format=MXN; _c6.alignment=RR; _c6.border=BD
        ws.cell(row_n,7).fill=Vd; ws.cell(row_n,7).border=BD
        ws.cell(row_n,8).fill=Vd; ws.cell(row_n,8).border=BD
        ws.row_dimensions[row_n].height = 16

    _rows4 = sorted(ded_log, key=lambda x:(x['mes'],x['fecha']))
    _row4 = 4
    _mes_actual4 = None
    _mes_buf4 = []

    for _g4 in _rows4:
        if _mes_actual4 != _g4['mes']:
            if _mes_actual4 is not None and _mes_buf4:
                _gd_subtotal(ws4, _row4, _mes_actual4, _mes_buf4)
                _row4 += 1
            _mes_actual4 = _g4['mes']
            _mes_buf4 = []

        _parcial4 = _g4.get('factor', 1.0) < 1.0
        _bg4 = PatternFill('solid', fgColor='FEF9C3') if _parcial4 else (Lv if _row4%2==0 else Lg)
        _nota4 = _g4.get('nota','')
        _sub_ded4 = _g4.get('sub_ded', _g4['sub'])
        # Método de pago: "PUE – Pago en una sola exhibición" ó "PPD – Pago en parcialidades"
        _mp = _g4.get('metodo_pago','')
        _fp = _g4.get('forma_pago','')
        _MP_DESC = {'PUE':'PUE – Una exhibición','PPD':'PPD – Parcialidades o diferido'}
        _FP_DESC = {
            '01':'01-Efectivo','02':'02-Cheque','03':'03-Transferencia','04':'04-Tarjeta crédito',
            '28':'28-Tarjeta débito','99':'99-Por definir',
        }
        _mp_str = _MP_DESC.get(_mp, _mp)
        _fp_str = _FP_DESC.get(_fp, _fp)
        _metodo_col = f"{_mp_str} / {_fp_str}" if _mp_str or _fp_str else ''
        _rv4 = [
            (MESES[_g4['mes']][:3], None),
            (_g4['fecha'],           None),
            (_g4['folio'],           None),
            (_g4['nom_e'],           None),
            (_sub_ded4,              MXN),
            (_g4['iva_acred'],       MXN),
            (_g4['conc'] + (f' [{_nota4}]' if _nota4 else ''), None),
            (_metodo_col,            None),
        ]
        for _ci4d, (_v4, _fmt4d) in enumerate(_rv4, 1):
            put(ws4, _row4, _ci4d, _v4, fill=_bg4,
                align=LL if _ci4d in(1,2,3,4,7,8) else RR, fmt=_fmt4d)
        ws4.row_dimensions[_row4].height = 13
        _mes_buf4.append(_g4)
        _row4 += 1

    if _mes_actual4 is not None and _mes_buf4:
        _gd_subtotal(ws4, _row4, _mes_actual4, _mes_buf4)
        _row4 += 1

    if ded_log:
        _tr4 = _row4
        put(ws4, _tr4, 1, 'TOTAL GENERAL', True, 'FFFFFF', Blk, LL)
        put(ws4, _tr4, 5, sum(g.get('sub_ded',g['sub']) for g in ded_log), True, 'FFFFFF', Blk, RR, MXN)
        put(ws4, _tr4, 6, sum(g['iva_acred'] for g in ded_log), True, 'FFFFFF', Blk, RR, MXN)
        for _ci4t in [2,3,4,7,8]:
            ws4.cell(_tr4,_ci4t).fill=Blk; ws4.cell(_tr4,_ci4t).border=BD
        ws4.row_dimensions[_tr4].height = 18

    for _ci4w, _w4 in enumerate([9,11,12,40,14,13,52,32], 1):
        ws4.column_dimensions[get_column_letter(_ci4w)].width = _w4


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
            parcial = g.get('factor', 1.0) < 1.0
            bg = PatternFill('solid', fgColor='FEF9C3') if parcial else (Lv if ri_d%2==0 else Lg)
            nota = g.get('nota', '')
            sub_d = g.get('sub_ded', g['sub'])
            for ci, (v, fmt) in enumerate([(g['fecha'],None),(g['folio'],None),(g['nom_e'],None),
                                            (g['conc']+(f' [{nota}]' if nota else ''),None),
                                            (None,None),(None,None),(sub_d,MXN),(g['iva_acred'],MXN),(None,None),(None,None)], 1):
                if v is not None:
                    put(ws_m, row_s, ci, v, fill=bg, align=LL if ci<=4 else RR, fmt=fmt)
                else:
                    ws_m.cell(row_s,ci).fill=bg; ws_m.cell(row_s,ci).border=BD
            ws_m.row_dimensions[row_s].height = 13; row_s += 1
        # Totals row for gastos
        if mes_ded:
            put(ws_m, row_s, 6, 'TOTAL DED.', True, 'FFFFFF', Vd, LL)
            put(ws_m, row_s, 7, sum(g.get('sub_ded',g['sub']) for g in mes_ded), True, 'FFFFFF', Vd, RR, MXN)
            put(ws_m, row_s, 8, sum(g['iva_acred'] for g in mes_ded), True, 'FFFFFF', Vd, RR, MXN)
            for ci in [1,2,3,4,5,9,10]:
                ws_m.cell(row_s,ci).fill=Vd; ws_m.cell(row_s,ci).border=BD
            ws_m.row_dimensions[row_s].height = 15; row_s += 1
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
    FILL_LOC = PatternFill('solid',fgColor='D1FAE5')
    FILL_PEN = PatternFill('solid',fgColor='FEF2F2')
    FILL_CAN = PatternFill('solid',fgColor='FEE2E2')
    FILL_SEC = PatternFill('solid',fgColor='1E3A8A')   # separador de sección

    def _am_hdr_row(ws, row):
        for ci, h in enumerate(hdrs_am, 1):
            c=ws.cell(row,ci,h); c.fill=Az; c.font=f(True,'FFFFFF',8); c.alignment=CC; c.border=BD
        ws.row_dimensions[row].height = 24

    def _am_sec(ws, row, label, n):
        ws.merge_cells(start_row=row,start_column=1,end_row=row,end_column=15)
        c=ws.cell(row,1,label); c.fill=FILL_SEC; c.font=f(True,'FFFFFF',9)
        c.alignment=CC; c.border=BD; ws.row_dimensions[row].height = 18

    def _am_write_rows(ws, rows, start_row):
        ri = start_row
        for e in rows:
            cancelado = e.get('cancelado', False)
            estado_txt = e['estado']
            if cancelado:    bg = FILL_CAN
            elif estado_txt=='LOCALIZADA': bg = FILL_LOC
            else:            bg = FILL_PEN
            rv = [(MESES[e['mes']][:3],None),(e['fecha'],None),(e['serie'],None),(e['folio'],None),(e['nom_r'],None),
                  (e['sub'],MXN),(e['iva'],MXN),(e['ret_isr'],MXN),(e['ret_iva'],MXN),(e['total'],MXN),(e['neto'],MXN),
                  (estado_txt,None),(e.get('banco_fecha',''),None),(e.get('banco_monto') or '',MXN if e.get('banco_monto') else None),(e.get('banco_desc',''),None)]
            for ci, (v, fmt) in enumerate(rv, 1):
                bold = ci == 12
                col  = '991B1B' if cancelado and ci==12 else \
                       '14532D' if estado_txt=='LOCALIZADA' and ci==12 else \
                       '6B7280' if estado_txt=='PENDIENTE'  and ci==12 else '000000'
                put(ws_am, ri, ci, v or '', bold, col, bg, LL if ci in(1,2,3,4,5,12,13,15) else RR, fmt)
            ws_am.row_dimensions[ri].height = 13; ri += 1
        return ri

    facturas = [x for x in emitidas_log if x['tipo'] != 'P']
    localizadas = sorted([e for e in facturas if e['estado']=='LOCALIZADA'], key=lambda x:(x['mes'],x['fecha']))
    pendientes  = sorted([e for e in facturas if e['estado']=='PENDIENTE'],  key=lambda x:(x['mes'],x['fecha']))
    canceladas  = sorted([e for e in facturas if e.get('cancelado',False)],  key=lambda x:(x['mes'],x['fecha']))

    cur_row = 2
    _am_sec(ws_am, cur_row, f'✅ LOCALIZADAS — {len(localizadas)} factura(s) con depósito identificado', len(localizadas)); cur_row += 1
    _am_hdr_row(ws_am, cur_row); cur_row += 1
    cur_row = _am_write_rows(ws_am, localizadas, cur_row)
    cur_row += 1
    _am_sec(ws_am, cur_row, f'⏳ PENDIENTES — {len(pendientes)} factura(s) sin depósito localizado', len(pendientes)); cur_row += 1
    _am_hdr_row(ws_am, cur_row); cur_row += 1
    cur_row = _am_write_rows(ws_am, pendientes, cur_row)
    if canceladas:
        cur_row += 1
        _am_sec(ws_am, cur_row, f'❌ CANCELADAS — {len(canceladas)} CFDI(s) cancelado(s) ante el SAT', len(canceladas)); cur_row += 1
        _am_hdr_row(ws_am, cur_row); cur_row += 1
        _am_write_rows(ws_am, canceladas, cur_row)
    for ci, w in enumerate([8,11,7,10,32,13,12,11,11,13,13,11,11,14,40], 1):
        ws_am.column_dimensions[get_column_letter(ci)].width = w
    ws_am.freeze_panes = 'A2'

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
