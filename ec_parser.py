"""
ec_parser.py — Parsers de estados de cuenta bancarios (sin GUI).
Portado desde gui_conciliacion.py para uso en Streamlit.
"""
import re
from datetime import date, datetime
from collections import defaultdict


# ── Utilidad ──────────────────────────────────────────────────────────────────

def parse_fecha(s):
    """Convierte una cadena de fecha a datetime.date, o devuelve la cadena original."""
    for fmt in ("%d/%m/%Y", "%d/%m/%y", "%Y/%m/%d", "%d-%m-%Y",
                "%d-%m-%y", "%Y-%m-%d", "%m/%d/%Y", "%m/%d/%y"):
        try:
            return datetime.strptime(str(s).strip(), fmt).date()
        except ValueError:
            pass
    return str(s).strip()


# ── Parsers por banco ──────────────────────────────────────────────────────────

def _parsear_bbva(texto, tablas, variante="BBVA Débito"):
    """Parser BBVA: tablas pdfplumber → fallback texto DD/MM/YYYY."""
    movimientos = []
    for tbl in tablas:
        if not tbl or len(tbl) < 2:
            continue
        encabezado = [str(c or "").upper() for c in tbl[0]]
        def col_idx(keywords):
            for kw in keywords:
                for i, h in enumerate(encabezado):
                    if kw in h:
                        return i
            return -1
        i_fecha = col_idx(["FECHA"])
        i_desc  = col_idx(["DESCRIPCI", "CONCEPTO", "MOVIMIENTO"])
        i_cargo = col_idx(["CARGO", "RETIRO", "DÉBITO", "DEBITO"])
        i_abono = col_idx(["ABONO", "DEPÓSITO", "DEPOSITO", "CRÉDITO", "CREDITO"])
        i_saldo = col_idx(["SALDO"])
        if i_fecha < 0 or i_desc < 0:
            continue
        for fila in tbl[1:]:
            if not fila or len(fila) <= max(i_fecha, i_desc):
                continue
            f_str = str(fila[i_fecha] or "").strip()
            desc  = str(fila[i_desc]  or "").strip()
            if not f_str or not desc:
                continue
            fecha = parse_fecha(f_str)
            if not fecha or (isinstance(fecha, str) and len(fecha) > 12):
                continue
            def _num(idx):
                if idx < 0 or idx >= len(fila):
                    return 0.0
                try:
                    return float(str(fila[idx] or "0").replace(",","").replace("$","") or "0")
                except Exception:
                    return 0.0
            cargo = _num(i_cargo); abono = _num(i_abono)
            saldo = _num(i_saldo) if i_saldo >= 0 else None
            dep = abono; ret = cargo
            if saldo is not None:
                movimientos.append((fecha, desc, dep, ret, saldo))
            else:
                movimientos.append((fecha, desc, dep, ret))
        if movimientos:
            return movimientos
    # fallback texto
    pat_f = re.compile(r"(\d{2}[/\-]\d{2}[/\-]\d{4})")
    pat_n = re.compile(r"[\d,]+\.\d{2}")
    for linea in texto.splitlines():
        m = pat_f.search(linea)
        if not m:
            continue
        fecha = parse_fecha(m.group(1))
        montos = [float(x.replace(",","")) for x in pat_n.findall(linea)]
        if not montos:
            continue
        desc = linea[:m.start()].strip() + " " + linea[m.end():].strip()
        desc = re.sub(r"[\d,]+\.\d{2}", "", desc).strip() or "—"
        if len(montos) >= 3:
            cargo, abono, saldo = montos[-3], montos[-2], montos[-1]
            movimientos.append((fecha, desc, abono, cargo, saldo))
        elif len(montos) == 2:
            monto, saldo = montos
            movimientos.append((fecha, desc, monto, 0.0, saldo))
        elif len(montos) == 1:
            movimientos.append((fecha, desc, montos[0], 0.0))
    return movimientos


def _parsear_bbva_cashmanagement(ruta, pdfplumber_mod):
    """Parser BBVA Cash Management posicional."""
    MESES = {"ENE":1,"FEB":2,"MAR":3,"ABR":4,"MAY":5,"JUN":6,
             "JUL":7,"AGO":8,"SEP":9,"OCT":10,"NOV":11,"DIC":12}
    pat_fecha = re.compile(
        r"^(\d{2})/(ENE|FEB|MAR|ABR|MAY|JUN|JUL|AGO|SEP|OCT|NOV|DIC)$", re.IGNORECASE)
    pat_monto = re.compile(r"^\d{1,3}(?:,\d{3})*\.\d{2}$")
    movimientos = []

    with pdfplumber_mod.open(ruta) as pdf:
        for page in pdf.pages:
            words = page.extract_words(keep_blank_chars=False, x_tolerance=2, y_tolerance=3)
            if not words:
                continue
            rows_tmp = defaultdict(list)
            for w in words:
                rows_tmp[round(w["top"])].append(w)

            x_cargo_hdr = x_abono_hdr = x_saldo_hdr = None
            for y_h in sorted(rows_tmp.keys()):
                row_w = rows_tmp[y_h]
                row_texts = [w["text"].upper() for w in row_w]
                if "CARGOS" in row_texts and "ABONOS" in row_texts:
                    for w in sorted(row_w, key=lambda w: w["x0"]):
                        t = w["text"].upper()
                        if t == "CARGOS" and x_cargo_hdr is None:
                            x_cargo_hdr = w["x0"]
                        elif t == "ABONOS" and x_abono_hdr is None:
                            x_abono_hdr = w["x0"]
                        elif t in ("OPERACIÓN","OPERACION","LIQUIDACIÓN","LIQUIDACION") \
                             and x_saldo_hdr is None:
                            x_saldo_hdr = w["x0"]
                    break
            if x_cargo_hdr is None or x_abono_hdr is None:
                continue

            x_sep   = (x_cargo_hdr + x_abono_hdr) / 2
            x_saldo = x_saldo_hdr if x_saldo_hdr else x_abono_hdr + 50
            rows = rows_tmp

            anio = date.today().year
            texto_pag = page.extract_text() or ""
            m_anio = re.search(r"/(\d{4})", texto_pag)
            if m_anio:
                anio = int(m_anio.group(1))

            pat_solo_num = re.compile(r"^[\d\s]+$")
            pat_hex      = re.compile(r"^[0-9A-Fa-f]{8,}$")
            pat_clabe    = re.compile(r"^\d{10,}$")
            _SKIP_CONT = ("BBVA MEXICO","AV. PASEO DE","ESTIMADO CLIENTE",
                          "SU ESTADO DE CUENTA","TOTAL DE MOVIMIENTOS","TOTAL IMPORTE",
                          "INSTITUCION DE BANCA")

            def _util_cont(tok):
                t = tok.strip()
                if not t or len(t) < 3: return False
                if pat_clabe.match(t): return False
                if pat_hex.match(t): return False
                if pat_solo_num.match(t): return False
                if len(t) > 100: return False
                letters = sum(1 for c in t if c.isalpha())
                if letters < 4: return False
                tu = t.upper()
                if any(s in tu for s in _SKIP_CONT): return False
                return True

            cur_fecha = cur_desc = None
            cur_dep = cur_ret = 0.0
            cur_saldo = None
            cur_conts = []

            def _flush():
                if cur_fecha is None or cur_desc is None:
                    return
                utiles = [c for c in cur_conts if _util_cont(c)]
                desc_final = cur_desc
                if utiles:
                    desc_final = cur_desc + " | " + " | ".join(utiles[:3])
                if cur_saldo is not None:
                    movimientos.append((cur_fecha, desc_final, cur_dep, cur_ret, cur_saldo))
                else:
                    movimientos.append((cur_fecha, desc_final, cur_dep, cur_ret))

            for y in sorted(rows.keys()):
                ws2 = sorted(rows[y], key=lambda w: w["x0"])
                tokens = [w["text"] for w in ws2]
                xs     = [w["x0"]  for w in ws2]
                if len(tokens) < 2:
                    continue
                m1 = pat_fecha.match(tokens[0])
                m2 = pat_fecha.match(tokens[1]) if len(tokens) > 1 else None
                if m1 and m2:
                    _flush(); cur_conts = []
                    try:
                        dia = int(m1.group(1))
                        mes = MESES[m1.group(2).upper()]
                        cur_fecha = date(anio, mes, dia)
                    except Exception:
                        cur_fecha = None; continue
                    cur_dep = cur_ret = 0.0; cur_saldo = None
                    for tok, x in zip(tokens, xs):
                        if not pat_monto.match(tok): continue
                        val = float(tok.replace(",", ""))
                        if x >= x_saldo - 5:   cur_saldo = val
                        elif x >= x_sep:        cur_dep   = val
                        else:                   cur_ret   = val
                    if cur_dep == 0.0 and cur_ret == 0.0:
                        cur_fecha = None; continue
                    desc_tokens = []
                    for tok in tokens[3:]:
                        if pat_monto.match(tok): break
                        desc_tokens.append(tok)
                    cur_desc = " ".join(desc_tokens).strip() or (tokens[2] if len(tokens) > 2 else "")
                elif cur_fecha is not None:
                    line = " ".join(tokens).strip()
                    if line: cur_conts.append(line)
            _flush()
    return movimientos


def _parsear_bbva_pyme(texto):
    """Parser BBVA Maestra PYME."""
    MESES = {"ENE":1,"FEB":2,"MAR":3,"ABR":4,"MAY":5,"JUN":6,
             "JUL":7,"AGO":8,"SEP":9,"OCT":10,"NOV":11,"DIC":12}
    y = re.search(r"\b(20\d{2})\b", texto)
    anio = int(y.group(1)) if y else date.today().year
    pat = re.compile(
        r"^(\d{2})/(ENE|FEB|MAR|ABR|MAY|JUN|JUL|AGO|SEP|OCT|NOV|DIC)"
        r"\s+\d{2}/[A-Z]{3}\s+(\w+)\s+(.*?)$", re.IGNORECASE)
    pat_monto = re.compile(r"([\d,]+\.\d{2})")
    _BBVA_SKIP = ("BBVA MEXICO","MAESTRA PYME BBVA","ESTADO DE CUENTA",
                  "PAGINA ","PÁGINA ","NO. CUENTA","NO. CLIENTE",
                  "FECHA SALDO","OPER LIQ COD","AV. PASEO DE LA REFORMA",
                  "R.F.C. BBA","ALCALDIA","CIUDAD DE MEXICO")
    movimientos = []
    cur_fecha = cur_cod = cur_desc = None
    cur_dep = cur_ret = 0.0; cur_saldo = None; cur_conts = []

    def _flush_bbva():
        if cur_fecha is None or cur_desc is None: return
        desc = cur_desc
        if cur_conts:
            utiles = []
            for cl in cur_conts[:6]:
                if len(cl) > 120: continue
                if re.match(r"^\d{6,}", cl): continue
                if re.match(r"^BNET", cl, re.I): continue
                if re.match(r"^\d+TRANSFERENCIA", cl, re.I): continue
                utiles.append(cl.strip())
            if utiles: desc = desc + " | " + " ".join(utiles)
        desc = re.sub(r"\s+", " ", desc).strip() or "—"
        if cur_saldo is not None:
            movimientos.append((cur_fecha, desc, cur_dep, cur_ret, cur_saldo))
        else:
            movimientos.append((cur_fecha, desc, cur_dep, cur_ret))

    for line in texto.splitlines():
        lu = line.strip().upper()
        if any(s in lu for s in _BBVA_SKIP): continue
        m = pat.match(line.strip())
        if m:
            _flush_bbva()
            dia = int(m.group(1)); mes_str = m.group(2).upper()
            cur_cod = m.group(3); resto = m.group(4).strip()
            try: cur_fecha = date(anio, MESES[mes_str], dia)
            except Exception: cur_fecha = None; continue
            numeros = [float(x.replace(",","")) for x in pat_monto.findall(resto)]
            if not numeros: cur_fecha = None; continue
            monto = numeros[0]
            is_abono = (cur_cod.upper() == "V45")
            cur_dep = monto if is_abono else 0.0
            cur_ret = 0.0 if is_abono else monto
            cur_desc = re.sub(r"\s+", " ", pat_monto.sub("", resto)).strip() or cur_cod
            cur_saldo = numeros[-1] if len(numeros) >= 2 else None
            cur_conts = []
        else:
            if cur_fecha is not None and line.strip():
                cur_conts.append(line.strip())
    _flush_bbva()

    return movimientos


def _parsear_banamex(texto, variante="Banamex Débito"):
    """Parser Citibanamex MiCuenta."""
    MESES = {"ENE":1,"FEB":2,"MAR":3,"ABR":4,"MAY":5,"JUN":6,
             "JUL":7,"AGO":8,"SEP":9,"OCT":10,"NOV":11,"DIC":12}
    anio_m = re.search(r"AL\s+\d+\s+DE\s+\w+\s+DE\s+(20\d{2})", texto, re.I)
    anio = int(anio_m.group(1)) if anio_m else date.today().year
    si_m = re.search(r"Saldo Anterior\s+\$?([\d,]+\.\d{2})", texto)
    saldo_ant = float(si_m.group(1).replace(",","")) if si_m else 0.0
    pat_fecha = re.compile(
        r"^(\d{1,2})\s+(ENE|FEB|MAR|ABR|MAY|JUN|JUL|AGO|SEP|OCT|NOV|DIC)\s+(.*?)$", re.IGNORECASE)
    pat_fin   = re.compile(r"([\d,]+\.\d{2})\s+([\d,]+\.\d{2})\s*$")
    pat_num   = re.compile(r"([\d,]+\.\d{2})")
    pat_skip  = re.compile(
        r"^ESTADO DE CUENTA|^CLIENTE:|^P[aá]gina\s*\d|^IRVING|"
        r"Centro de Atenci|^Ciudad de M[eé]xico|^Resto del pa[ií]s|"
        r"^DETALLE DE OPERACIONES|^FECHA\s+CONCEPTO|^\d{6}\.", re.IGNORECASE)
    pat_fin_det = re.compile(
        r"GRAFICO TRANSACCIONAL|Otros cargos|TOTAL DE MOVIMIENTOS|"
        r"CONCEPTOS\s*$|ClavProd Serv|SUBTOTALES", re.IGNORECASE)
    lineas = texto.splitlines(); i = 0; movimientos = []; saldo_act = saldo_ant
    while i < len(lineas):
        linea = lineas[i].strip()
        if pat_fin_det.search(linea): break
        m = pat_fecha.match(linea)
        if not m: i += 1; continue
        dia, mes_str, desc_ini = int(m.group(1)), m.group(2).upper(), m.group(3).strip()
        if re.search(r"SALDO ANTERIOR", desc_ini, re.I):
            nums = pat_num.findall(desc_ini)
            if nums: saldo_act = float(nums[-1].replace(",",""))
            i += 1; continue
        if re.search(r"^(FECHA|CONCEPTO|RETIROS|DEPOSITOS)", desc_ini, re.I):
            i += 1; continue
        try: fecha = date(anio, MESES[mes_str], dia)
        except Exception: i += 1; continue
        j = i + 1; bloque_lines = [linea]
        while j < len(lineas):
            nl = lineas[j].strip()
            if pat_fin_det.search(nl): break
            if pat_fecha.match(nl): break
            if pat_skip.search(nl): j += 1; continue
            if nl: bloque_lines.append(nl)
            j += 1
        monto = None; saldo_nuevo = None
        for bl in reversed(bloque_lines):
            pm = pat_fin.search(bl)
            if pm:
                monto = float(pm.group(1).replace(",",""))
                saldo_nuevo = float(pm.group(2).replace(",",""))
                break
        if saldo_nuevo is not None:
            diff = round(saldo_nuevo - saldo_act, 2)
            if monto is None: monto = abs(diff)
            dep = monto if diff >= 0 else 0.0
            ret = 0.0 if diff >= 0 else monto
            saldo_act = saldo_nuevo
            desc = re.sub(r"\s*([\d,]+\.\d{2})\s*", " ", desc_ini).strip()
            desc = re.sub(r"\s+", " ", desc).strip() or "—"
            movimientos.append((fecha, desc, dep, ret, saldo_nuevo))
        i = j
    return movimientos


def _parsear_santander(texto):
    """Parser Santander/HSBC/Banregio."""
    MESES = {"ENE":1,"FEB":2,"MAR":3,"ABR":4,"MAY":5,"JUN":6,
             "JUL":7,"AGO":8,"SEP":9,"OCT":10,"NOV":11,"DIC":12}
    pat_fecha = re.compile(r"^(\d{2})-([A-Z]{3})-(\d{4})\s+(\d+)\s+(.*?)$")
    pat_monto = re.compile(r"([\d,]+\.\d{2})")
    movimientos = []; saldo_ant = None
    si_m = re.search(
        r"(?:SALDO\s+FINAL\s+DEL\s+PERIODO\s+ANTERIOR|SALDO\s+ANTERIOR|Saldo\s+inicial)"
        r"[:\s\$]*([\d,]+\.\d{2})", texto, re.I)
    if si_m:
        try: saldo_ant = float(si_m.group(1).replace(",",""))
        except Exception: pass
    lineas = texto.splitlines(); i = 0
    while i < len(lineas):
        m = pat_fecha.match(lineas[i].strip())
        if not m: i += 1; continue
        dia = int(m.group(1)); mes_str = m.group(2); year = int(m.group(3))
        desc_ini = m.group(5).strip()
        if mes_str not in MESES: i += 1; continue
        try: fecha = date(year, MESES[mes_str], dia)
        except Exception: i += 1; continue
        j = i + 1; bloque = desc_ini
        while j < len(lineas) and not pat_fecha.match(lineas[j].strip()):
            nl = lineas[j].strip()
            if nl: bloque += " " + nl
            j += 1
        montos = [float(x.replace(",","")) for x in pat_monto.findall(bloque)]
        if montos:
            saldo = montos[-1]; monto = montos[-2] if len(montos) >= 2 else 0.0
            if saldo_ant is not None:
                diff = round(saldo - saldo_ant, 2)
                dep, ret = (monto, 0.0) if diff >= 0 else (0.0, monto)
            else:
                dep, ret = monto, 0.0
            desc = pat_monto.sub("", desc_ini).strip()
            desc = re.sub(r"\s+", " ", desc).strip() or "—"
            saldo_ant = saldo
            movimientos.append((fecha, desc, dep, ret, saldo))
        i = j
    if not movimientos:
        pat_f2 = re.compile(r"(\d{2}/\d{2}/\d{4})")
        for linea in texto.splitlines():
            m2 = pat_f2.search(linea)
            if not m2: continue
            fecha = parse_fecha(m2.group(1))
            montos = [float(x.replace(",","")) for x in pat_monto.findall(linea)]
            if not montos: continue
            saldo = montos[-1]; monto = montos[-2] if len(montos) >= 2 else 0.0
            if saldo_ant is not None:
                diff = round(saldo - saldo_ant, 2)
                dep, ret = (monto, 0.0) if diff >= 0 else (0.0, monto)
            else:
                dep, ret = monto, 0.0
            desc = pat_f2.sub("", linea); desc = pat_monto.sub("", desc).strip() or "—"
            saldo_ant = saldo
            movimientos.append((fecha, desc, dep, ret, saldo))
    return movimientos


def _parsear_scotiabank(texto, tablas=None):
    """Parser Scotiabank."""
    MESES = {"ENE":1,"FEB":2,"MAR":3,"ABR":4,"MAY":5,"JUN":6,
             "JUL":7,"AGO":8,"SEP":9,"OCT":10,"NOV":11,"DIC":12,
             "JAN":1,"APR":4,"AUG":8,"OCT":10,"NOV":11,"DEC":12}
    anio = 2024
    y4 = re.search(r"\b(20\d{2})\b", texto)
    if y4: anio = int(y4.group(1))
    else:
        y2 = re.search(r"\b(\d{2})[/\-](\d{2})[/\-](\d{2})\b", texto)
        if y2: anio = 2000 + int(y2.group(3))
    pat_fecha = re.compile(r"^(\d{2})\s+([A-Z]{3})\s+(.*?)$")
    pat_monto_d = re.compile(r"\$([\d,]+\.\d{2})")
    pat_monto   = re.compile(r"([\d,]+\.\d{2})")
    movimientos = []; saldo_ant = None
    si_m = re.search(r"(?:Saldo\s+inicial|SALDO\s+ANTERIOR|Saldo\s+anterior)[^\d]+([\d,]+\.\d{2})", texto, re.I)
    if si_m:
        try: saldo_ant = float(si_m.group(1).replace(",",""))
        except Exception: pass
    lineas = texto.splitlines(); i = 0
    while i < len(lineas):
        m = pat_fecha.match(lineas[i].strip())
        if not m: i += 1; continue
        dia = int(m.group(1)); mes_str = m.group(2); desc_ini = m.group(3).strip()
        if mes_str not in MESES or dia < 1 or dia > 31: i += 1; continue
        try: fecha = date(anio, MESES[mes_str], dia)
        except Exception: i += 1; continue
        j = i + 1; bloque_lines = [desc_ini]
        while j < len(lineas) and not pat_fecha.match(lineas[j].strip()):
            nl = lineas[j].strip()
            if nl: bloque_lines.append(nl)
            j += 1
        bloque = " ".join(bloque_lines)
        montos = [float(x.replace(",","")) for x in pat_monto_d.findall(bloque)]
        if not montos: montos = [float(x.replace(",","")) for x in pat_monto.findall(bloque)]
        if montos:
            saldo = montos[-1]; monto = montos[-2] if len(montos) >= 2 else 0.0
            if saldo_ant is not None:
                diff = round(saldo - saldo_ant, 2)
                dep, ret = (monto, 0.0) if diff >= 0 else (0.0, monto)
            else:
                dep, ret = monto, 0.0
            desc = pat_monto_d.sub("", desc_ini); desc = pat_monto.sub("", desc).strip() or "—"
            desc = re.sub(r"\s+", " ", desc).strip() or "—"
            saldo_ant = saldo
            movimientos.append((fecha, desc, dep, ret, saldo))
        i = j
    return movimientos


def _parsear_inbursa(texto):
    """Parser Inbursa: MMM. DD CONCEPTO monto saldo."""
    MESES = {"ENE":1,"FEB":2,"MAR":3,"ABR":4,"MAY":5,"JUN":6,
             "JUL":7,"AGO":8,"SEP":9,"OCT":10,"NOV":11,"DIC":12}
    year_m = re.search(r"\b(20\d{2})\b", texto)
    anio = int(year_m.group(1)) if year_m else date.today().year
    pat_linea = re.compile(r"^([A-Z]{3})\.\s{1,3}(\d{1,2})\s+(.*)", re.MULTILINE)
    pat_monto = re.compile(r"([\d,]+\.\d{2})")
    movimientos = []; saldo_ant = None
    for m in pat_linea.finditer(texto):
        mes_str = m.group(1); dia_str = m.group(2); resto = m.group(3).strip()
        if mes_str not in MESES: continue
        if re.search(r"\b(REFERENCIA|CONCEPTO|FECHA)\b", resto): continue
        try: fecha = date(anio, MESES[mes_str], int(dia_str))
        except Exception: continue
        montos = []
        for s in pat_monto.findall(resto):
            try: montos.append(float(s.replace(",","")))
            except Exception: pass
        if not montos: continue
        saldo = montos[-1]
        if "BALANCE INICIAL" in resto.upper(): saldo_ant = saldo; continue
        monto = montos[-2] if len(montos) >= 2 else abs(saldo - (saldo_ant or saldo))
        if saldo_ant is not None:
            diff = round(saldo - saldo_ant, 2)
            dep, ret = (monto, 0.0) if diff >= 0 else (0.0, monto)
        else:
            dep, ret = monto, 0.0
        desc = pat_monto.sub("", resto).strip()
        desc = re.sub(r"^\S+\s+", "", desc).strip()
        desc = re.sub(r"\s+", " ", desc).strip() or "—"
        saldo_ant = saldo
        movimientos.append((fecha, desc, dep, ret, saldo))
    return movimientos


def _parsear_amex(texto):
    """Parser American Express."""
    MESES = {"enero":1,"febrero":2,"marzo":3,"abril":4,"mayo":5,"junio":6,
             "julio":7,"agosto":8,"septiembre":9,"octubre":10,"noviembre":11,"diciembre":12}
    year_m = re.search(r"\b(20\d{2})\b", texto)
    anio = int(year_m.group(1)) if year_m else date.today().year
    pat_fecha = re.compile(r"^(\d{1,2})\s+de\s*([A-Za-záéíóúÁÉÍÓÚ]+)\s+(.*?)$", re.IGNORECASE)
    pat_monto = re.compile(r"([\d,]+\.\d{2})")
    pat_corte = re.compile(
        r"Total de las|Estado de Cuenta P[aá]g|Este no es un|Resumen de|Abreviaci[oó]n|N[uú]mero de Cuenta", re.I)
    movimientos = []; lineas = texto.splitlines(); i = 0
    while i < len(lineas):
        m = pat_fecha.match(lineas[i].strip())
        if not m: i += 1; continue
        dia = int(m.group(1)); mes_str = m.group(2).lower().strip(); desc_ini = m.group(3).strip()
        if mes_str not in MESES: i += 1; continue
        try: fecha = date(anio, MESES[mes_str], dia)
        except Exception: i += 1; continue
        j = i + 1; cont = 0; bloque_lines = [desc_ini]
        while j < len(lineas) and cont < 4:
            nl = lineas[j].strip()
            if pat_fecha.match(nl) or pat_corte.search(nl): break
            if nl: bloque_lines.append(nl); cont += 1
            j += 1
        bloque = " ".join(bloque_lines)
        montos = [float(x.replace(",","")) for x in pat_monto.findall(bloque)]
        if not montos: i += 1; continue
        monto = montos[0]
        is_cr = bool(re.search(r"\bCR\b", bloque)) or "PAGO RECIBIDO" in bloque.upper()
        dep = monto if is_cr else 0.0; ret = 0.0 if is_cr else monto
        desc = pat_monto.sub("", desc_ini).strip()
        desc = re.sub(r"\bCR\b", "", desc).strip()
        desc = re.sub(r"\s+", " ", desc).strip() or "—"
        movimientos.append((fecha, desc, dep, ret))
        i += 1
    return movimientos


def _parsear_afirme(texto, ruta=None, pdfplumber_mod=None):
    """Parser Banca Afirme."""
    import unicodedata as _ud
    MESES = {"ENE":1,"FEB":2,"MAR":3,"ABR":4,"MAY":5,"JUN":6,
             "JUL":7,"AGO":8,"SEP":9,"OCT":10,"NOV":11,"DIC":12}
    per_m = re.search(r'Per[ií]odo\s+de\s+(\d{2})([A-Z]{3})(\d{4})AL', texto, re.IGNORECASE)
    if not per_m: per_m = re.search(r'(\d{2})([A-Z]{3})(\d{4})AL', texto)
    if per_m:
        mes = MESES.get(per_m.group(2).upper(), 1); anio = int(per_m.group(3))
    else:
        anio, mes = date.today().year, date.today().month
    ini_m = re.search(r'Saldo\s+inicial\s+\$\s*([\d,]+\.\d{2})', texto)
    saldo_ant = float(ini_m.group(1).replace(",","")) if ini_m else None
    pat_monto = re.compile(r'\$\s*([\d,]+\.\d{2})')
    _SKIP_TX = ("descripci","detalle","informaci","resumen","comision","deposito","retiro",
                "saldo","ganancia","rendimient","numero de","tasa de","pagina","sucursal",
                "direcci","cliente","pyme","periodo","fecha de")
    _SKIP_CONT = ("pagina","sello digital","cadena original","este documento",
                  "representaci","detalle de operaciones","regimen fiscal",
                  "av. ju","banca afirme","instituci","r.f.c. ba","ipab","condusef",
                  "sus ahorros","saldo inicial","numero de cuenta","clave bancaria",
                  "lider pyme","gas 122 sa de cv","cll poniente","estado de cuenta al",
                  "numero de cliente","metodo de pago","tipo o factor","uso cfdi",
                  "claveprodserv","forma de pago","||1.0|","dqgfup","ty5ilfr","20-104",
                  "dia descripcion","referencia depositos")
    def _norm(s):
        return _ud.normalize('NFKD', s).encode('ascii','ignore').decode('ascii').lower()
    def _es_cont_valida(line):
        ln = _norm(line)
        if any(s in ln for s in _SKIP_CONT): return False
        if len(line) > 200: return False
        if line.count('+') + line.count('=') + line.count('/') > 10: return False
        if '||' in line: return False
        if re.search(r'[A-Za-z0-9+/]{40,}', line): return False
        return True
    movimientos = []
    if ruta is not None and pdfplumber_mod is not None:
        X_DIA_MAX = 50; all_rows = []
        with pdfplumber_mod.open(ruta) as pdf:
            for page in pdf.pages:
                words = page.extract_words(keep_blank_chars=False, x_tolerance=2, y_tolerance=3)
                if not words: continue
                rows_tmp = defaultdict(list)
                for w in words: rows_tmp[round(w["top"])].append(w)
                for y in sorted(rows_tmp.keys()):
                    ws2 = sorted(rows_tmp[y], key=lambda w: w["x0"])
                    line = " ".join(w["text"] for w in ws2)
                    all_rows.append((y, ws2[0]["x0"], line))
        pat_dia = re.compile(r"^(\d{1,2})\s+")
        blocks = []; cur_dia = None; cur_rows = []
        for y, x0, line in all_rows:
            m = pat_dia.match(line)
            if m and x0 < X_DIA_MAX:
                dia = int(m.group(1))
                if 1 <= dia <= 31:
                    if not any(kw in line.lower() for kw in _SKIP_TX):
                        if cur_dia is not None: blocks.append((cur_dia, cur_rows))
                        cur_dia = dia; cur_rows = [line]; continue
            if cur_dia is not None and _es_cont_valida(line): cur_rows.append(line)
        if cur_dia is not None: blocks.append((cur_dia, cur_rows))
        for dia, rows in blocks:
            bt = " ".join(rows)
            montos = pat_monto.findall(bt)
            if len(montos) < 2: continue
            try:
                saldo = float(montos[-1].replace(",",""))
                monto = float(montos[-2].replace(",",""))
                fecha = date(anio, mes, dia)
            except Exception: continue
            fl = rows[0]; desc = pat_monto.sub("", fl)
            desc = re.sub(r"\$\s*", "", desc); desc = re.sub(r"^\d{1,2}\s+", "", desc)
            desc = re.sub(r"\s+", " ", desc).strip() or "—"
            dest_m = re.search(r"DESTINATARIO:(.+?)(?=\s*\(?DATO\s+NO|\s+RFC\s+DEST|\s+ND\s+CVE|$)", bt, re.IGNORECASE)
            dest = ""
            if dest_m:
                dest = dest_m.group(1).strip()
                dest = re.sub(r"\s*\(?DA\s+TO\s+NO.+$", "", dest, flags=re.I).strip()
                dest = re.sub(r"\s*\(?DATO\s+NO.+$", "", dest, flags=re.I).strip()
                dest = dest.rstrip("(").strip()
            conc_m = re.search(r"CONCEPTO:(.+?)(?=\s+HORA:|\s+CVE\s+RASTREO:|\s+M[eé]todo|$)", bt, re.IGNORECASE)
            concepto = conc_m.group(1).strip()[:80] if conc_m else ""
            if dest: desc += " | " + dest
            if concepto and concepto.upper() not in desc.upper(): desc += " | " + concepto
            if saldo_ant is not None:
                diff = round(saldo - saldo_ant, 2)
                dep = round(diff, 2) if diff > 0 else 0.0
                ret = round(-diff, 2) if diff < 0 else 0.0
            else:
                dep, ret = 0.0, monto
            saldo_ant = saldo
            movimientos.append((fecha, desc, dep, ret, saldo))
        if movimientos: return movimientos
    # fallback texto
    pat_tx = re.compile(r"^(\d{1,2})\s+([^\n]+)", re.MULTILINE)
    saldo_ant2 = None
    ini_m2 = re.search(r"Saldo\s+inicial\s+\$\s*([\d,]+\.\d{2})", texto)
    if ini_m2: saldo_ant2 = float(ini_m2.group(1).replace(",",""))
    for m in pat_tx.finditer(texto):
        dia = int(m.group(1)); resto = m.group(2).strip()
        if dia < 1 or dia > 31: continue
        if any(kw in resto.lower() for kw in _SKIP_TX): continue
        montos_raw = pat_monto.findall(resto)
        if len(montos_raw) < 2: continue
        try:
            saldo = float(montos_raw[-1].replace(",","")); monto = float(montos_raw[-2].replace(",",""))
            fecha = date(anio, mes, dia)
        except Exception: continue
        desc = pat_monto.sub("", resto); desc = re.sub(r"\$\s*", "", desc)
        desc = re.sub(r"\s+\d{4,}\s*$", "", desc); desc = re.sub(r"\s+", " ", desc).strip() or "—"
        if saldo_ant2 is not None:
            diff = round(saldo - saldo_ant2, 2)
            dep = round(diff, 2) if diff > 0 else 0.0; ret = round(-diff, 2) if diff < 0 else 0.0
        else:
            dep, ret = 0.0, monto
        saldo_ant2 = saldo
        movimientos.append((fecha, desc, dep, ret, saldo))
    return movimientos


def _parsear_banorte(texto, ruta=None, pdfplumber_mod=None):
    """Parser Banorte: DD-MMM-YY."""
    MESES = {"ENE":1,"FEB":2,"MAR":3,"ABR":4,"MAY":5,"JUN":6,
             "JUL":7,"AGO":8,"SEP":9,"OCT":10,"NOV":11,"DIC":12}
    pat_fecha = re.compile(r"^(\d{2}-(?:ENE|FEB|MAR|ABR|MAY|JUN|JUL|AGO|SEP|OCT|NOV|DIC)-\d{2})")
    pat_monto = re.compile(r"(?<!\S)([\d,]+\.\d{2})(?!\S|\d)")
    if ruta is not None and pdfplumber_mod is not None:
        lineas = []
        with pdfplumber_mod.open(ruta) as _pdf:
            for _page in _pdf.pages:
                _words = _page.extract_words(keep_blank_chars=False, x_tolerance=1, y_tolerance=3)
                _rows = defaultdict(list)
                for _w in _words: _rows[round(_w["top"])].append(_w)
                for _y in sorted(_rows.keys()):
                    _ws = sorted(_rows[_y], key=lambda w: w["x0"])
                    lineas.append(" ".join(w["text"] for w in _ws))
    else:
        lineas = texto.splitlines()
    _SKIP_HDR = ("FECHA DESCRIPCI","MONTO DEL","DETALLE DE MOVIMIENTOS",
                 "ESTADO DE CUENTA","ENLACE GLOBAL","LINEA DIRECTA","LÍNEA DIRECTA",
                 "BANCO MERCANTIL","CIUDAD DE MEXICO","NUEVO LEON",
                 "PAGINA ","PÁGINA ","TIPO DE ENVIO","INFORMACION DEL",
                 "NO. DE CLIENTE","DATOS DE SUCURSAL","PLAZA:","TELEFONO:",
                 "RESUMEN INTEGRAL","RESUMEN DEL PERIODO","BANCO MERCANTIL DEL NORTE",
                 "/63","5140 5640","3669 9040")
    movimientos = []; saldo_anterior = None; cur_fecha_str = None; cur_lineas = []

    def _flush(fecha_str, lineas_desc):
        nonlocal saldo_anterior
        if not fecha_str or not lineas_desc: return
        try:
            d, mes_str, a = fecha_str.split("-")
            mes = MESES[mes_str]; anio = 2000 + int(a)
            fecha = date(anio, mes, int(d))
        except Exception: return
        linea1 = lineas_desc[0] if lineas_desc else ""
        continuas = lineas_desc[1:] if len(lineas_desc) > 1 else []
        tu = linea1.upper()
        if "SALDO ANTERIOR" in tu:
            m2 = pat_monto.findall(" " + linea1)
            if m2:
                try: saldo_anterior = float(m2[-1].replace(",",""))
                except Exception: pass
            return
        if any(s in tu for s in ("DETALLE DE MOVIMIENTOS","MONTO DEL","DESCRIPCIÓN",
                                   "FECHA DESCRIPCI","SALDO FINAL","TOTAL DE")): return
        montos_raw = pat_monto.findall(" " + linea1)
        montos = []
        for s in montos_raw:
            try: montos.append(float(s.replace(",","")))
            except ValueError: pass
        if not montos: return
        saldo = montos[-1]
        if len(montos) >= 2: monto = montos[-2]
        else: monto = abs(saldo - saldo_anterior) if saldo_anterior is not None else 0.0
        if saldo_anterior is not None:
            diff = round(saldo - saldo_anterior, 2)
            dep, ret = (monto, 0.0) if diff >= 0 else (0.0, monto)
        else:
            if any(k in tu for k in ("DEP","DEPOSITO","ABONO","SPEI RECIBIDO","SUPER SERV")):
                dep, ret = monto, 0.0
            else:
                dep, ret = 0.0, monto
        desc1 = linea1.strip()
        for s in montos_raw[-2:]:
            idx = desc1.rfind(s)
            if idx >= 0: desc1 = desc1[:idx].rstrip()
        desc1 = re.sub(r"\s+", " ", desc1).strip()
        extras = []
        for cl in continuas:
            cl = cl.strip()
            if re.match(r"^[\d\s]{1,20}$", cl): continue
            if "CVE RAST" in cl.upper() or "HORA LIQ" in cl.upper(): continue
            extras.append(cl)
        desc_extra = " ".join(extras).strip()
        texto_cont = " ".join(continuas)
        if "BENEF:" in texto_cont.upper():
            _bm = re.search(r"BENEF:([^,\(\[]+)", texto_cont, re.IGNORECASE)
            _rm = re.search(r"RFC:\s*([A-Z&\xc0-\xff0-9]{10,13})", texto_cont, re.IGNORECASE)
            _benef = _bm.group(1).strip() if _bm else ""
            _rfc   = _rm.group(1).strip() if _rm else ""
            if _benef: desc = desc1 + " | " + _benef + (" RFC:"+_rfc if _rfc else "")
            else: desc = (desc1 + " " + desc_extra).strip() or "—"
        else:
            desc = (desc1 + (" " + desc_extra if desc_extra else "")).strip() or "—"
        desc = re.sub(r"\s+", " ", desc).strip() or "—"
        saldo_anterior = saldo
        movimientos.append((fecha, desc, dep, ret, saldo))



    for linea in lineas:
        lu = linea.upper()
        if any(s in lu for s in _SKIP_HDR): continue
        m = pat_fecha.match(linea)
        if m:
            _flush(cur_fecha_str, cur_lineas)
            cur_fecha_str = m.group(1)
            _resto = linea[m.end():]
            _resto = re.sub(r"^([A-ZÁÉÍÓÚÑ]) ([A-ZÁÉÍÓÚÑ])", r"\1\2", _resto)
            cur_lineas = [_resto.strip()]
        else:
            if cur_fecha_str is not None: cur_lineas.append(linea.strip())
    _flush(cur_fecha_str, cur_lineas)

    return movimientos


def _parsear_fila(fila):
    """Parsea fila de tabla. Retorna (fecha,desc,dep,ret) o None."""
    patron_monto = re.compile(r"^-?[\d,]+\.\d{2}$")
    patron_fecha = re.compile(r"\b(\d{1,2}[/\-]\d{1,2}[/\-]\d{2,4})\b")
    celdas = [str(c).strip() if c else "" for c in fila]
    fecha_str = ""; desc_parts = []; montos = []
    for celda in celdas:
        m = patron_fecha.search(celda)
        if m and not fecha_str: fecha_str = m.group(1)
        elif patron_monto.match(celda.replace(" ","").replace("(","-").replace(")","")):
            val = celda.replace(",","").replace(" ","").replace("(","-").replace(")","")
            try: montos.append(float(val))
            except ValueError: pass
        elif celda and not patron_fecha.search(celda): desc_parts.append(celda)
    if not fecha_str: return None
    fecha = parse_fecha(fecha_str); desc = " ".join(desc_parts).strip() or "—"
    dep = ret = 0.0
    if len(montos) == 1:
        v = montos[0]; dep, ret = (v, 0.0) if v > 0 else (0.0, abs(v))
    elif len(montos) >= 2:
        dep = montos[0] if montos[0] > 0 else 0.0
        ret = montos[1] if montos[1] > 0 else 0.0
        if montos[0] < 0 and montos[1] >= 0: ret, dep = abs(montos[0]), montos[1]
    else: return None
    return (fecha, desc, dep, ret)


def _parsear_linea(linea, patron_fecha, patron_monto):
    """Parsea línea de texto libre."""
    m_fecha = patron_fecha.search(linea)
    if not m_fecha: return None
    fecha = parse_fecha(m_fecha.group(1))
    montos_raw = patron_monto.findall(linea)
    montos = []
    for s in montos_raw:
        try: montos.append(float(s.replace(",","")))
        except ValueError: pass
    pos_fecha_fin = m_fecha.end()
    primer_monto_pos = linea.find(montos_raw[0]) if montos_raw else len(linea)
    desc = linea[pos_fecha_fin:primer_monto_pos].strip(" |-\t") or "—"
    dep = ret = 0.0
    if len(montos) == 1: dep = montos[0]
    elif len(montos) >= 2: dep = montos[0]; ret = montos[1]
    else: return None
    return (fecha, desc, dep, ret)


# ── API pública ────────────────────────────────────────────────────────────────

def leer_pdf(ruta, pdfplumber_mod, banco_key=""):
    """Extrae movimientos de un PDF de estado de cuenta.
    Retorna lista de tuplas (fecha, desc, dep, ret) o (fecha, desc, dep, ret, saldo).
    """
    paginas_texto = []; paginas_tablas = []
    with pdfplumber_mod.open(ruta) as pdf:
        for page in pdf.pages:
            txt = page.extract_text() or ""
            if txt.strip(): paginas_texto.append(txt)
            tbls = page.extract_tables() or []
            paginas_tablas.extend(tbls)
    texto_total = "\n".join(paginas_texto)
    bk = banco_key.lstrip("🔍 ─").strip()

    if bk.startswith("Banorte"):
        return _parsear_banorte(texto_total, ruta=ruta, pdfplumber_mod=pdfplumber_mod)
    if bk == "BBVA Pyme" or bk.startswith("BBVA"):
        if any(k in texto_total.upper() for k in ("CASH MANAGEMENT","MAESTRA PYME","OPER LIQ COD")):
            movs = _parsear_bbva_cashmanagement(ruta, pdfplumber_mod)
            if movs: return movs
        if bk == "BBVA Pyme":
            movs = _parsear_bbva_pyme(texto_total)
            if movs: return movs
        if bk.startswith("BBVA"):
            movs = _parsear_bbva(texto_total, paginas_tablas, bk)
            if movs: return movs
    if bk.startswith("Banamex"):
        movs = _parsear_banamex(texto_total, bk)
        if movs: return movs
    if bk.startswith("Santander") or bk.startswith("HSBC") or bk.startswith("Banregio"):
        movs = _parsear_santander(texto_total)
        if movs: return movs
    if bk.startswith("Scotiabank"):
        movs = _parsear_scotiabank(texto_total, paginas_tablas)
        if movs: return movs
    if bk.startswith("Inbursa"):
        movs = _parsear_inbursa(texto_total)
        if movs: return movs
    if bk.startswith("American Express"):
        movs = _parsear_amex(texto_total)
        if movs: return movs
    if bk.startswith("Afirme"):
        movs = _parsear_afirme(texto_total, ruta=ruta, pdfplumber_mod=pdfplumber_mod)
        if movs: return movs

    # Auto-detección
    BANORTE_PAT = re.compile(r"\d{2}-(?:ENE|FEB|MAR|ABR|MAY|JUN|JUL|AGO|SEP|OCT|NOV|DIC)-\d{2}")
    if BANORTE_PAT.search(texto_total):
        movs = _parsear_banorte(texto_total)
        if movs: return movs
    if any(k in texto_total.upper() for k in ("CASH MANAGEMENT","MAESTRA PYME","OPER LIQ COD")):
        movs = _parsear_bbva_cashmanagement(ruta, pdfplumber_mod)
        if movs: return movs
    if any("CARGO" in str(t) or "ABONO" in str(t) for t in paginas_tablas):
        movs = _parsear_bbva(texto_total, paginas_tablas, "BBVA Débito")
        if movs: return movs
    patron_fecha = re.compile(r"\b(\d{1,2}[/\-]\d{1,2}[/\-]\d{2,4})\b")
    patron_monto = re.compile(r"[\d,]+\.\d{2}")
    movimientos = []
    for linea in texto_total.splitlines():
        if not patron_fecha.search(linea): continue
        mov = _parsear_linea(linea, patron_fecha, patron_monto)
        if mov: movimientos.append(mov)
    return movimientos


def leer_excel(ruta, openpyxl_mod):
    """Extrae movimientos de un Excel de estado de cuenta."""
    patron_fecha = re.compile(r"\d{1,2}[/\-]\d{1,2}[/\-]\d{2,4}")
    wb = openpyxl_mod.load_workbook(ruta, data_only=True); ws = wb.active
    movimientos = []; col_fecha = col_desc = col_dep = col_ret = None
    for row in ws.iter_rows():
        vals = [str(c.value or "").lower() for c in row]
        for i, v in enumerate(vals):
            if "fecha" in v and col_fecha is None: col_fecha = i
            if any(x in v for x in ("concepto","descripci","referencia")) and col_desc is None: col_desc = i
            if any(x in v for x in ("dep","abono","cargo_cr")) and col_dep is None: col_dep = i
            if any(x in v for x in ("retiro","cargo","egreso")) and col_ret is None: col_ret = i
        if col_fecha is not None: break
    if col_fecha is None: col_fecha, col_desc, col_dep, col_ret = 0, 1, 2, 3
    for row in ws.iter_rows(min_row=2, values_only=True):
        try:
            fecha_raw = row[col_fecha] if col_fecha < len(row) else None
            desc_raw  = row[col_desc]  if col_desc  < len(row) else None
            dep_raw   = row[col_dep]   if col_dep   < len(row) else None
            ret_raw   = row[col_ret]   if col_ret   < len(row) else None
            if fecha_raw is None: continue
            fecha = parse_fecha(str(fecha_raw)) if not hasattr(fecha_raw, "year") else fecha_raw
            dep = float(str(dep_raw).replace(",","")) if dep_raw not in (None,"","None") else 0.0
            ret = float(str(ret_raw).replace(",","")) if ret_raw not in (None,"","None") else 0.0
            desc = str(desc_raw or "—").strip()
            if dep == 0.0 and ret == 0.0: continue
            movimientos.append((fecha, desc, abs(dep), abs(ret)))
        except Exception: continue
    wb.close()
    return movimientos


def calcular_saldos(movimientos, saldo_ini=0.0):
    """Calcula saldo acumulado para movimientos que no lo traen."""
    filas = []; saldo = saldo_ini; total_dep = total_ret = 0.0
    for mov in movimientos:
        if len(mov) == 5:
            fecha, desc, dep, ret, saldo = mov
        else:
            fecha, desc, dep, ret = mov
            saldo += dep - ret
        total_dep += dep; total_ret += ret
        filas.append((fecha, desc, dep, ret, saldo))
    return filas, total_dep, total_ret


def generar_excel_bytes(filas, nombre_base, saldo_ini=0.0, saldo_esp=None):
    """Genera el Excel de reporte y lo retorna como bytes."""
    import io
    import xlsxwriter

    total_dep = sum(f[2] for f in filas)
    total_ret = sum(f[3] for f in filas)
    saldo_fin_real = filas[-1][4] if filas else saldo_ini

    buf = io.BytesIO()
    wb = xlsxwriter.Workbook(buf, {"in_memory": True})
    ws  = wb.add_worksheet("Movimientos")
    ws_c = wb.add_worksheet("Conciliación")

    fmt_h   = wb.add_format({"bold":True,"bg_color":"#1E6FBF","font_color":"#FFFFFF","border":1,"align":"center"})
    fmt_d   = wb.add_format({"num_format":"DD/MM/YYYY","border":1})
    fmt_t   = wb.add_format({"border":1})
    fmt_dep = wb.add_format({"num_format":"#,##0.00","border":1,"bg_color":"#E6F2FB"})
    fmt_ret = wb.add_format({"num_format":"#,##0.00","border":1,"bg_color":"#FDE8E4"})
    fmt_sal = wb.add_format({"num_format":"#,##0.00","border":1,"bold":True})
    fmt_ttl = wb.add_format({"bold":True,"bg_color":"#CFE7F8","num_format":"#,##0.00","border":1})
    fmt_lbl = wb.add_format({"bold":True,"border":1})
    fmt_tit = wb.add_format({"bold":True,"font_size":14,"font_color":"#1E6FBF","align":"center","valign":"vcenter"})
    fmt_sub = wb.add_format({"bold":True,"font_size":11,"bg_color":"#AED6F2","border":1,"align":"center"})
    fmt_cl  = wb.add_format({"bold":True,"border":1,"bg_color":"#B7D9EF","align":"left","indent":1})
    fmt_cv  = wb.add_format({"num_format":"#,##0.00","border":1,"align":"right"})
    fmt_ct  = wb.add_format({"bold":True,"num_format":"#,##0.00","border":2,"bg_color":"#CFE7F8","align":"right"})
    fmt_ok  = wb.add_format({"bold":True,"num_format":"#,##0.00","border":2,"bg_color":"#E6F2FB","font_color":"#3E6E96","align":"right"})
    fmt_err = wb.add_format({"bold":True,"num_format":"#,##0.00","border":2,"bg_color":"#FDE8E4","font_color":"#E14B3D","align":"right"})

    n_filas = len(filas)
    ws.merge_range(0,0,0,4,f"Estado de Cuenta — {nombre_base}  ({n_filas} movimientos)", fmt_tit)
    ws.set_row(0, 24)
    ws.set_column(0,0,14); ws.set_column(1,1,55); ws.set_column(2,4,17)
    for c, h in enumerate(["Fecha","Descripción","Depósito","Retiro","Saldo"]):
        ws.write(1, c, h, fmt_h)
    ws.freeze_panes(2, 0)
    ws.autofilter(1, 0, 1+n_filas, 4)
    for r, (fecha, desc, dep, ret, sal) in enumerate(filas, start=2):
        if isinstance(fecha, str): ws.write(r, 0, fecha, fmt_t)
        else: ws.write_datetime(r, 0, datetime.combine(fecha, datetime.min.time()), fmt_d)
        ws.write(r, 1, desc, fmt_t)
        ws.write(r, 2, dep if dep else "", fmt_dep if dep else fmt_t)
        ws.write(r, 3, ret if ret else "", fmt_ret if ret else fmt_t)
        ws.write(r, 4, sal, fmt_sal)
    tot_row = n_filas + 2
    ws.write(tot_row, 1, "TOTALES", fmt_lbl)
    ws.write_formula(tot_row, 2, f"=SUM(C3:C{tot_row})", fmt_ttl)
    ws.write_formula(tot_row, 3, f"=SUM(D3:D{tot_row})", fmt_ttl)
    ws.write(tot_row, 4, saldo_fin_real, fmt_ttl)
    ws.set_landscape(); ws.fit_to_pages(1, 0); ws.repeat_rows(1)

    diferencia = (saldo_fin_real - saldo_esp) if saldo_esp is not None else None
    ws_c.set_column(0,0,36); ws_c.set_column(1,1,20)
    ws_c.merge_range(0,0,0,1,"Conciliación Bancaria", fmt_tit)
    ws_c.set_row(0, 24)
    ws_c.merge_range(1,0,1,1, nombre_base, fmt_sub)
    conc_data = [
        ("Saldo inicial", saldo_ini, fmt_cv),
        ("(+) Total depósitos", total_dep, fmt_cv),
        ("(-) Total retiros", total_ret, fmt_cv),
        ("= Saldo final calculado", saldo_fin_real, fmt_ct),
    ]
    if saldo_esp is not None:
        conc_data.append(("Saldo final esperado (banco)", saldo_esp, fmt_cv))
        f_dif = fmt_ok if diferencia is not None and abs(diferencia) < 0.01 else fmt_err
        conc_data.append(("Diferencia (calculado − esperado)", diferencia, f_dif))
    for i, (lbl, val, fmt_v) in enumerate(conc_data, start=2):
        ws_c.write(i, 0, lbl, fmt_cl); ws_c.write(i, 1, val, fmt_v)

    wb.close()
    buf.seek(0)
    return buf.read()


BANCOS = [
    "Auto-detectar",
    "Banorte Débito", "Banorte Empresarial",
    "BBVA Débito", "BBVA Pyme", "BBVA Cash Management",
    "Banamex Débito", "Banamex Empresarial",
    "Santander", "HSBC", "Scotiabank",
    "Banregio", "Inbursa", "American Express", "Afirme",
]
