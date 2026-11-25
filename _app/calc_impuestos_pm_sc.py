"""
_app/calc_impuestos_pm_sc.py — Motor de Impuestos Persona Moral · Sociedad Civil (v1, 2026-11-25)

Pagos provisionales mensuales de una S.C. (Título II LISR, ingresos efectivamente cobrados):
  · ISR PM (art. 14 LISR): facturado → ajuste por cobranza → cobrado, coeficiente de utilidad,
    PTU, anticipos de remanente distribuible a socios (art. 14 fr. II), pérdidas, pagos anteriores.
  · IVA (arts. 1, 1-A, 5 y 5-D LIVA): trasladado cobrado, acreditable (DIOT), saldo a favor, IVA retenido.
  · Retenciones de IVA a terceros por concepto y tasa (art. 1-A LIVA, art. 3 RLIVA).
  · ISR retenido por sueldos y por remanente distribuible (asimilados).
  · Actualización (art. 17-A CFF) y recargos (art. 21 CFF) por pago extemporáneo.

Cada mes se calcula con los acumulados de los meses ya guardados (enero → diciembre).
Funciones puras: no dependen de Streamlit.
"""
from __future__ import annotations

import datetime as _dt
import io
import json
import math
import re
from decimal import Decimal, ROUND_HALF_UP

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

MESES = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio",
         "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]

HOJA_DATOS = "_datos_modulo"
TASA_IVA = 0.16
TASA_ISR_PM = 0.30

# Conceptos de retención de IVA (art. 1-A LIVA / art. 3 RLIVA)
CONCEPTOS_RET_IVA = [
    {"clave": "honorarios",   "nombre": "Servicios personales independientes", "tasa": 0.16 * 2 / 3, "texto": "2/3 del IVA (10.6667%)"},
    {"clave": "arrendamiento", "nombre": "Arrendamiento a personas físicas",   "tasa": 0.16 * 2 / 3, "texto": "2/3 del IVA (10.6667%)"},
    {"clave": "comisionistas", "nombre": "Comisionistas personas físicas",     "tasa": 0.16 * 2 / 3, "texto": "2/3 del IVA (10.6667%)"},
    {"clave": "fletes",       "nombre": "Autotransporte terrestre de carga",   "tasa": 0.04,         "texto": "4%"},
    {"clave": "personal",     "nombre": "Servicios con personal a disposición", "tasa": 0.06,        "texto": "6%"},
    {"clave": "desperdicios", "nombre": "Desperdicios para uso industrial",    "tasa": 0.16,         "texto": "100% del IVA (16%)"},
    {"clave": "extranjeros",  "nombre": "Residentes en el extranjero sin EP",  "tasa": 0.16,         "texto": "100% del IVA (16%)"},
]
_RET_POR_CLAVE = {c["clave"]: c for c in CONCEPTOS_RET_IVA}

# INPC (base 2Q julio 2018 = 100, INEGI). El usuario agrega meses nuevos en Configuración.
INPC_DEFAULT = {
    "2025-11": 142.645, "2025-12": 143.042,
    "2026-01": 143.588, "2026-02": 144.307, "2026-03": 145.544, "2026-04": 145.831,
    "2026-05": 145.527, "2026-06": 145.131, "2026-07": 145.169, "2026-08": 145.462,
    "2026-09": 146.075, "2026-10": 147.026,
}


def r2(x) -> float:
    """Redondeo a 2 decimales, mitad hacia arriba (como el SAT)."""
    return float(Decimal(str(x or 0)).quantize(Decimal("0.01"), ROUND_HALF_UP))


def config_default(ejercicio: int) -> dict:
    return {
        "ejercicio": ejercicio,
        "coef_1": 0.0,             # coeficiente del ejercicio anterior al anterior (enero hasta mes_cambio-1)
        "coef_2": 0.0,             # coeficiente de la última anual presentada
        "mes_cambio_coef": 1,      # primer mes en que aplica coef_2
        "tasa_recargos": 0.0207,   # mensual por mora (LIF 2026)
        "inpc": dict(INPC_DEFAULT),
    }


def coeficiente_mes(cfg: dict, mes: int) -> float:
    if mes >= int(cfg.get("mes_cambio_coef", 1) or 1):
        return float(cfg.get("coef_2") or 0)
    return float(cfg.get("coef_1") or 0)


# ══════════════════════════════════════════════════════════════════════════════
# LECTURA DE BALANZA Y DIOT
# ══════════════════════════════════════════════════════════════════════════════
def _norm(s) -> str:
    return re.sub(r"\s+", " ", str(s or "")).strip().lower()


def leer_balanza(data: bytes) -> dict:
    """
    Lee una balanza de comprobación en Excel.
    Busca la fila de encabezados (Cuenta / Saldo inicial / Cargos / Abonos / Saldo final).
    Retorna {"periodo": int|None, "empresa": str, "rfc": str, "cuentas": {codigo: {...}}}.
    """
    wb = load_workbook(io.BytesIO(data), read_only=True, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    meta = {"periodo": None, "empresa": "", "rfc": ""}
    hdr_i, cols = None, {}
    for i, r in enumerate(rows):
        if not r:
            continue
        k = _norm(r[0])
        if k.startswith("mes inicial") and len(r) > 1:
            meta["periodo"] = r[1]
        elif k.startswith("empresa") and len(r) > 1:
            meta["empresa"] = str(r[1] or "").strip()
        elif k.startswith("r. f. c") or k.startswith("rfc") or k.startswith("r.f.c"):
            meta["rfc"] = str(r[1] or "").strip()
        names = [_norm(v) for v in r]
        if "cuenta" in names and any("cargo" in n for n in names) and any("abono" in n for n in names):
            hdr_i = i
            for j, n in enumerate(names):
                if n == "cuenta":
                    cols["cuenta"] = j
                elif n.startswith("descrip"):
                    cols["desc"] = j
                elif "saldo inicial" in n:
                    cols["si"] = j
                elif n.startswith("cargo"):
                    cols["cargos"] = j
                elif n.startswith("abono"):
                    cols["abonos"] = j
                elif "saldo final" in n:
                    cols["sf"] = j
            break
    if hdr_i is None:
        raise ValueError("No encontré la fila de encabezados (Cuenta, Cargos, Abonos) en la balanza.")
    cuentas = {}
    for r in rows[hdr_i + 1:]:
        if not r:
            continue
        cod = r[cols["cuenta"]]
        if not cod or not re.match(r"^\d{3}", str(cod).strip()):
            continue
        g = lambda key: float(r[cols[key]] or 0) if key in cols and isinstance(r[cols[key]], (int, float)) else 0.0
        cuentas[str(cod).strip()] = {
            "desc": str(r[cols["desc"]] or "").strip() if "desc" in cols else "",
            "si": g("si"), "cargos": g("cargos"), "abonos": g("abonos"), "sf": g("sf"),
        }
    meta["cuentas"] = cuentas
    return meta


def leer_diot(data: bytes) -> dict:
    """
    Lee el reporte de movimientos para DIOT (Excel).
    Retorna totales y la lista de proveedores con IVA retenido.
    """
    wb = load_workbook(io.BytesIO(data), read_only=True, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    hdr_i, c = None, {}
    for i, r in enumerate(rows):
        names = [_norm(v) for v in (r or [])]
        if any(n.startswith("rfc prov") for n in names) and any(n == "actos 16%" for n in names):
            hdr_i = i
            for j, n in enumerate(names):
                c.setdefault(n, j)
            break
    if hdr_i is None:
        raise ValueError("No encontré los encabezados de la DIOT (RFC Prov., Actos 16%, IVA retenido…).")

    def col(*nombres):
        for n in nombres:
            if n in c:
                return c[n]
        return None

    j_rfc = col("rfc prov.", "rfc prov")
    j_nom = col("razón social", "razon social")
    j_per = col("mes póliza", "mes poliza")
    j_a16, j_i16 = col("actos 16%"), col("iva acred. 16%")
    j_a8, j_i8 = col("actos 8%"), col("iva acred. 8%")
    j_a0, j_ex = col("actos 0%"), col("actos iva exento")
    j_ret = col("iva retenido")
    j_imp, j_iimp = col("actos 16% importación", "actos 16% importacion"), None

    def num(r, j):
        v = r[j] if j is not None and j < len(r) else 0
        return float(v) if isinstance(v, (int, float)) else 0.0

    tot = {"actos_16": 0.0, "iva_16": 0.0, "actos_8": 0.0, "iva_8": 0.0,
           "actos_0": 0.0, "exentos": 0.0, "iva_retenido": 0.0, "importacion": 0.0}
    proveedores, periodos = [], set()
    for r in rows[hdr_i + 1:]:
        if not r or j_rfc is None or not r[j_rfc]:
            continue
        a16, i16, ret = num(r, j_a16), num(r, j_i16), num(r, j_ret)
        tot["actos_16"] += a16
        tot["iva_16"] += i16
        tot["actos_8"] += num(r, j_a8)
        tot["iva_8"] += num(r, j_i8)
        tot["actos_0"] += num(r, j_a0)
        tot["exentos"] += num(r, j_ex)
        tot["iva_retenido"] += ret
        tot["importacion"] += num(r, j_imp)
        if j_per is not None and r[j_per]:
            periodos.add(r[j_per])
        if ret:
            proveedores.append({"rfc": str(r[j_rfc]).strip(),
                                "nombre": str(r[j_nom] or "").strip() if j_nom is not None else "",
                                "base": a16, "iva": i16, "retenido": ret,
                                "concepto": clasificar_retencion(a16, ret, str(r[j_rfc]))})
    return {"totales": {k: r2(v) for k, v in tot.items()},
            "proveedores": proveedores, "periodos": sorted(periodos)}


def clasificar_retencion(base: float, retenido: float, rfc: str = "") -> str:
    """Clasifica una retención por su tasa efectiva."""
    if not base:
        return "honorarios"
    t = retenido / base
    if abs(t - 0.04) < 0.005:
        return "fletes"
    if abs(t - 0.06) < 0.005:
        return "personal"
    if abs(t - 0.16) < 0.005:
        return "desperdicios"
    return "honorarios"   # 2/3 del IVA (honorarios / arrendamiento / comisionistas)


def _cta(bal: dict, prefijo: str, campo: str) -> float:
    """Valor de la cuenta exacta; si no existe, suma de subcuentas inmediatas que empiezan con el prefijo."""
    cu = bal["cuentas"]
    if prefijo in cu:
        return cu[prefijo][campo]
    return 0.0


def extraer_mes(bal: dict, diot: dict | None) -> tuple[dict, list[str]]:
    """
    Arma los datos de captura del mes a partir de balanza (+ DIOT).
    Códigos agrupadores SAT:
      401 ingresos · 402 devoluciones · 208-01 IVA trasladado cobrado · 601-01 sueldos ·
      216-01 ISR sueldos · 601-32 / 601-31 anticipos a socios · 216-02 / 216-13 ISR remanente ·
      215-01 PTU · 113-08 ISR retenido bancos · 113-01 IVA a favor · 118-01 IVA acreditable.
    Retorna (datos, avisos).
    """
    avisos: list[str] = []
    facturado = _cta(bal, "401", "abonos") - _cta(bal, "401", "cargos")
    devol = _cta(bal, "402", "cargos")
    iva_cobrado = _cta(bal, "208-01", "abonos")
    cobrado = r2(iva_cobrado / TASA_IVA) if iva_cobrado else 0.0

    sueldos = _cta(bal, "601-01", "cargos")
    isr_sueldos = _cta(bal, "216-01", "abonos")
    rem_32 = _cta(bal, "601-32", "cargos")
    rem_31 = _cta(bal, "601-31", "cargos")
    isr_rem = _cta(bal, "216-02", "abonos") + max(_cta(bal, "216-13", "abonos"), 0)
    remanente = rem_32 + rem_31
    if sueldos == 0 and rem_31 > 0:
        avisos.append(
            f"La cuenta 601-01 Sueldos no tiene cargos y la 601-31 Asimilados trae ${rem_31:,.2f}. "
            "Si ahí quedó la nómina, separa sueldos y remanente (y su ISR) en la captura.")

    ptu = _cta(bal, "215-01", "cargos")
    ret_bancos = _cta(bal, "113-08", "cargos") or _cta(bal, "113-08-0001", "cargos")
    saldo_favor_iva = _cta(bal, "113-01", "sf")

    d = {
        "facturado": r2(facturado), "devoluciones": r2(devol), "cobrado": cobrado,
        "iva_trasladado": r2(iva_cobrado),
        "actos_0": 0.0, "actos_exentos": 0.0,
        "gastos_16": 0.0, "iva_acreditable": r2(_cta(bal, "118-01", "cargos")),
        "saldo_favor_aplicado": 0.0,
        "retenciones": {c["clave"]: 0.0 for c in CONCEPTOS_RET_IVA},
        "iva_retenido_diot": 0.0,
        "sueldos": r2(sueldos), "isr_sueldos": r2(isr_sueldos),
        "remanente": r2(remanente), "isr_remanente": r2(isr_rem),
        "ptu_pagada": r2(ptu), "isr_ret_bancos": r2(ret_bancos),
        "saldo_favor_iva_disponible": r2(saldo_favor_iva),
        "isr_honorarios_216_04": r2(_cta(bal, "216-04", "abonos")),
    }
    if diot:
        t = diot["totales"]
        d["gastos_16"] = r2(t["actos_16"] + t["actos_8"])
        d["iva_acreditable"] = r2(t["iva_16"] + t["iva_8"])
        d["iva_retenido_diot"] = r2(t["iva_retenido"])
        for p in diot["proveedores"]:
            d["retenciones"][p["concepto"]] = r2(d["retenciones"][p["concepto"]] + p["base"])
        bal_ret = _cta(bal, "216-10", "abonos")
        if bal_ret and abs(bal_ret - t["iva_retenido"]) > 1:
            avisos.append(f"IVA retenido: DIOT ${t['iva_retenido']:,.2f} vs balanza 216-10 ${bal_ret:,.2f}.")
        if abs(d["iva_acreditable"] - _cta(bal, "118-01", "cargos")) > 1 and _cta(bal, "118-01", "cargos"):
            avisos.append(f"IVA acreditable: DIOT ${d['iva_acreditable']:,.2f} vs balanza 118-01 "
                          f"${_cta(bal, '118-01', 'cargos'):,.2f}.")
    else:
        avisos.append("Sin DIOT: el IVA acreditable se toma de la cuenta 118-01 y no hay desglose de retenciones.")
    return d, avisos


# ══════════════════════════════════════════════════════════════════════════════
# ACTUALIZACIÓN Y RECARGOS
# ══════════════════════════════════════════════════════════════════════════════
def _inpc_mes_anterior(cfg: dict, fecha: _dt.date) -> float:
    """INPC del mes anterior a la fecha; si aún no se publica, el último disponible."""
    tabla = cfg.get("inpc") or {}
    if not tabla:
        return 0.0
    y, m = fecha.year, fecha.month - 1
    if m == 0:
        y, m = y - 1, 12
    clave = f"{y:04d}-{m:02d}"
    if clave in tabla:
        return float(tabla[clave])
    previas = sorted(k for k in tabla if k <= clave)
    return float(tabla[previas[-1]]) if previas else 0.0


def vencimiento(ejercicio: int, mes: int) -> _dt.date:
    y, m = (ejercicio, mes + 1) if mes < 12 else (ejercicio + 1, 1)
    return _dt.date(y, m, 17)


def meses_atraso(fecha_pago: _dt.date, venc: _dt.date) -> int:
    if fecha_pago <= venc:
        return 0
    n = (fecha_pago.year - venc.year) * 12 + fecha_pago.month - venc.month
    if fecha_pago.day > venc.day:
        n += 1
    return max(n, 0)


def accesorios(cfg: dict, ejercicio: int, mes: int, recargos: dict | None) -> dict:
    """Factor de actualización y meses de atraso del mes."""
    out = {"aplica": False, "factor": 1.0, "meses": 0, "inpc_pago": 0.0, "inpc_venc": 0.0,
           "vencimiento": vencimiento(ejercicio, mes).isoformat(), "fecha_pago": None}
    if not recargos or not recargos.get("aplica") or not recargos.get("fecha_pago"):
        return out
    fp = recargos["fecha_pago"]
    fp = _dt.date.fromisoformat(fp) if isinstance(fp, str) else fp
    venc = recargos.get("vencimiento") or vencimiento(ejercicio, mes)
    venc = _dt.date.fromisoformat(venc) if isinstance(venc, str) else venc
    ip, iv = _inpc_mes_anterior(cfg, fp), _inpc_mes_anterior(cfg, venc)
    factor = max(math.trunc(ip / iv * 10000) / 10000, 1.0) if ip and iv else 1.0
    out.update(aplica=True, factor=factor, meses=meses_atraso(fp, venc), inpc_pago=ip, inpc_venc=iv,
               vencimiento=venc.isoformat(), fecha_pago=fp.isoformat())
    return out


def _act_rec(monto: float, acc: dict, tasa: float) -> tuple[float, float]:
    if not acc["aplica"] or monto <= 0:
        return 0.0, 0.0
    act = r2(monto * (acc["factor"] - 1))
    rec = r2((monto + act) * tasa * acc["meses"])
    return act, rec


# ══════════════════════════════════════════════════════════════════════════════
# CÁLCULO DEL MES
# ══════════════════════════════════════════════════════════════════════════════
def acumulados_previos(doc: dict, mes: int) -> dict:
    """Suma lo guardado de enero a mes-1. Lista los meses que faltan."""
    acc = {"ingresos": 0.0, "anticipos": 0.0, "ptu": 0.0, "pagos_prov": 0.0, "ret_bancos": 0.0,
           "faltan": [], "usados": []}
    meses = doc.get("meses", {})
    for m in range(1, mes):
        g = meses.get(str(m))
        if not g:
            acc["faltan"].append(m)
            continue
        e, res = g["entrada"], g["resultado"]
        acc["ingresos"] += res["isr"]["ingresos_mes"]
        acc["anticipos"] += e.get("remanente", 0)
        acc["ptu"] += e.get("ptu_pagada", 0)
        acc["ret_bancos"] += e.get("isr_ret_bancos", 0)
        pagado = g.get("isr_pagado")
        acc["pagos_prov"] += float(pagado if pagado is not None else res["isr"]["a_pagar"])
        acc["usados"].append(m)
    ini = doc.get("acumulados_iniciales") or {}
    if acc["faltan"] and ini:
        acc["ingresos"] += float(ini.get("ingresos", 0))
        acc["anticipos"] += float(ini.get("anticipos", 0))
        acc["ptu"] += float(ini.get("ptu", 0))
        acc["pagos_prov"] += float(ini.get("pagos_prov", 0))
        acc["ret_bancos"] += float(ini.get("ret_bancos", 0))
    return {k: (r2(v) if isinstance(v, float) else v) for k, v in acc.items()}


def calcular_mes(doc: dict, mes: int, e: dict) -> dict:
    """Calcula ISR, IVA, retenciones y accesorios del mes con los acumulados guardados."""
    cfg = doc["config"]
    ejercicio = int(cfg["ejercicio"])
    prev = acumulados_previos(doc, mes)
    tasa_rec = float(cfg.get("tasa_recargos", 0.0207))
    acc = accesorios(cfg, ejercicio, mes, e.get("recargos"))

    # ── ISR PM ────────────────────────────────────────────────────────────────
    ingresos_mes = r2(e["cobrado"] - e.get("devoluciones", 0))
    ingresos_acum = r2(prev["ingresos"] + ingresos_mes)
    coef = coeficiente_mes(cfg, mes)
    utilidad = r2(ingresos_acum * coef)
    ptu_acum = r2(prev["ptu"] + e.get("ptu_pagada", 0))
    anticipos_acum = r2(prev["anticipos"] + e.get("remanente", 0))
    perdidas = float(e.get("perdidas", 0) or 0)
    base = r2(max(utilidad - ptu_acum - anticipos_acum - perdidas, 0))
    causado = r2(base * TASA_ISR_PM)
    ret_bancos_acum = r2(prev["ret_bancos"] + e.get("isr_ret_bancos", 0))
    a_pagar = r2(max(causado - prev["pagos_prov"] - ret_bancos_acum, 0))
    dif_cobranza = r2(e["cobrado"] - e["facturado"])
    isr = {
        "facturado": e["facturado"],
        "aumento_cobranza": max(dif_cobranza, 0.0),
        "disminucion_no_cobrado": max(-dif_cobranza, 0.0),
        "cobrado": e["cobrado"], "devoluciones": e.get("devoluciones", 0),
        "ingresos_mes": ingresos_mes, "ingresos_previos": prev["ingresos"], "ingresos_acum": ingresos_acum,
        "coeficiente": coef, "utilidad": utilidad, "ptu_acum": ptu_acum,
        "remanente_mes": e.get("remanente", 0), "anticipos_acum": anticipos_acum,
        "perdidas": perdidas, "base": base, "causado": causado,
        "pagos_anteriores": prev["pagos_prov"], "ret_bancos_acum": ret_bancos_acum, "a_pagar": a_pagar,
    }

    # ── Retenciones de IVA ────────────────────────────────────────────────────
    ret_det, ret_total = [], 0.0
    for c in CONCEPTOS_RET_IVA:
        base_c = float(e["retenciones"].get(c["clave"], 0) or 0)
        if not base_c:
            continue
        ret = r2(base_c * c["tasa"])
        ret_det.append({"concepto": c["nombre"], "tasa": c["texto"], "valor": r2(base_c),
                        "iva_trasladado": r2(base_c * TASA_IVA), "iva_retenido": ret})
        ret_total += ret
    ret_total = r2(ret_total)

    # ── IVA ───────────────────────────────────────────────────────────────────
    gravados = r2(e["iva_trasladado"] / TASA_IVA) if e["iva_trasladado"] else 0.0
    tot_actos = gravados + e.get("actos_0", 0) + e.get("actos_exentos", 0)
    proporcion = 1.0 if not tot_actos else (gravados + e.get("actos_0", 0)) / tot_actos
    acreditable = r2(e["iva_acreditable"] * proporcion)
    a_cargo = r2(e["iva_trasladado"] - acreditable - e.get("saldo_favor_aplicado", 0))
    total_iva = r2(max(a_cargo, 0) + ret_total)
    iva = {"actos_16": gravados, "trasladado": e["iva_trasladado"], "gastos_16": e.get("gastos_16", 0),
           "proporcion": round(proporcion, 4), "acreditable": acreditable,
           "saldo_favor_aplicado": e.get("saldo_favor_aplicado", 0), "a_cargo": a_cargo,
           "saldo_a_favor": r2(max(-a_cargo, 0)), "retenido": ret_total, "total": total_iva}

    # ── Accesorios ────────────────────────────────────────────────────────────
    lineas = [
        ("ISR personas morales", a_pagar),
        ("IVA (incluye retenido)", total_iva),
        ("ISR retenciones por salarios", e.get("isr_sueldos", 0)),
        ("ISR retenciones por asimilados (remanente)", e.get("isr_remanente", 0)),
    ]
    resumen = []
    for nombre, monto in lineas:
        act, rec = _act_rec(monto, acc, tasa_rec)
        resumen.append({"impuesto": nombre, "a_cargo": r2(monto), "actualizacion": act,
                        "recargos": rec, "total": r2(monto + act + rec)})
    total = r2(sum(x["total"] for x in resumen))

    return {"mes": mes, "isr": isr, "iva": iva, "retenciones": ret_det, "accesorios": acc,
            "resumen": resumen, "total": total, "previos": prev}


# ══════════════════════════════════════════════════════════════════════════════
# PAPEL DE TRABAJO EN EXCEL (un mes por columna)
# ══════════════════════════════════════════════════════════════════════════════
_F = Font(name="Arial", size=10)
_FB = Font(name="Arial", size=10, bold=True)
_FT = Font(name="Arial", size=12, bold=True)
_GR = PatternFill("solid", fgColor="F3F3F3")
_TOP = Border(top=Side(style="thin"))
_NUM = "#,##0.00"


def generar_excel(doc: dict) -> bytes:
    cfg = doc["config"]
    meses = {int(k): v for k, v in (doc.get("meses") or {}).items()}
    wb = Workbook()
    titulo = f"{doc.get('razon_social') or doc.get('rfc')} — Impuestos {cfg['ejercicio']}"

    def hoja(nombre, filas, primero=False):
        ws = wb.active if primero else wb.create_sheet(nombre)
        ws.title = nombre
        ws["A1"] = titulo
        ws["A1"].font = _FT
        ws["A2"] = nombre
        ws["A2"].font = _FB
        ws.cell(row=4, column=1, value="Concepto").font = _FB
        for i, m in enumerate(MESES):
            c = ws.cell(row=4, column=2 + i, value=m)
            c.font, c.fill = _FB, _GR
        ws.cell(row=4, column=1).fill = _GR
        for r, (etq, fn, fmt, negrita) in enumerate(filas, start=5):
            a = ws.cell(row=r, column=1, value=etq)
            a.font = _FB if negrita else _F
            for m in range(1, 13):
                g = meses.get(m)
                v = fn(g) if g else None
                c = ws.cell(row=r, column=1 + m, value=v)
                c.number_format = fmt or _NUM
                c.font = _FB if negrita else _F
                if negrita:
                    c.border = _TOP
            if negrita:
                a.border = _TOP
        ws.column_dimensions["A"].width = 62
        for col in "BCDEFGHIJKLM":
            ws.column_dimensions[col].width = 14
        ws.freeze_panes = "B5"
        return ws

    R = lambda k: (lambda g: g["resultado"]["isr"][k])
    V = lambda k: (lambda g: g["resultado"]["iva"][k])
    E = lambda k: (lambda g: g["entrada"].get(k, 0))

    def res_linea(i, k):
        return lambda g: g["resultado"]["resumen"][i][k]

    hoja("Resumen", [
        ("ISR personas morales", res_linea(0, "a_cargo"), None, False),
        ("IVA (incluye retenido)", res_linea(1, "a_cargo"), None, False),
        ("ISR retenciones por salarios", res_linea(2, "a_cargo"), None, False),
        ("ISR retenciones por asimilados (remanente)", res_linea(3, "a_cargo"), None, False),
        ("Actualización", lambda g: r2(sum(x["actualizacion"] for x in g["resultado"]["resumen"])), None, False),
        ("Recargos", lambda g: r2(sum(x["recargos"] for x in g["resultado"]["resumen"])), None, False),
        ("Total a pagar", lambda g: g["resultado"]["total"], None, True),
        ("ISR provisional efectivamente pagado (declarado)", lambda g: g.get("isr_pagado"), None, False),
    ], primero=True)
    hoja("ISR PM", [
        ("Ingresos facturados del mes (sin IVA)", R("facturado"), None, False),
        ("(+) Aumento: cobranza de facturas de meses anteriores", R("aumento_cobranza"), None, False),
        ("(−) Disminución: facturado en el mes pendiente de cobro", R("disminucion_no_cobrado"), None, False),
        ("Ingresos efectivamente cobrados del mes", R("cobrado"), None, True),
        ("(−) Devoluciones, descuentos y bonificaciones", R("devoluciones"), None, False),
        ("Ingresos cobrados acumulados", R("ingresos_acum"), None, False),
        ("Coeficiente de utilidad", R("coeficiente"), "0.0000", False),
        ("Utilidad fiscal estimada", R("utilidad"), None, False),
        ("(−) PTU pagada (acumulada)", R("ptu_acum"), None, False),
        ("Remanente distribuible pagado a socios en el mes", R("remanente_mes"), None, False),
        ("(−) Anticipos de remanente a socios, acumulado (art. 14 fr. II LISR)", R("anticipos_acum"), None, False),
        ("(−) Pérdidas fiscales por amortizar", R("perdidas"), None, False),
        ("Base del pago provisional", R("base"), None, True),
        ("ISR causado 30%", R("causado"), None, False),
        ("(−) Pagos provisionales anteriores", R("pagos_anteriores"), None, False),
        ("(−) ISR retenido por bancos (acumulado)", R("ret_bancos_acum"), None, False),
        ("ISR a pagar del mes", R("a_pagar"), None, True),
    ])
    hoja("IVA", [
        ("Actos gravados 16% cobrados", V("actos_16"), None, False),
        ("IVA trasladado cobrado", V("trasladado"), None, False),
        ("Gastos con IVA pagados (DIOT)", V("gastos_16"), None, False),
        ("Proporción de acreditamiento", V("proporcion"), "0.0000", False),
        ("(−) IVA acreditable", V("acreditable"), None, False),
        ("(−) Saldo a favor aplicado", V("saldo_favor_aplicado"), None, False),
        ("IVA a cargo / (a favor) del periodo", V("a_cargo"), None, True),
        ("(+) IVA retenido a terceros", V("retenido"), None, False),
        ("Total IVA a enterar", V("total"), None, True),
    ])
    filas_ret = []
    for c in CONCEPTOS_RET_IVA:
        k = c["clave"]
        filas_ret.append((f"{c['nombre']} — valor de la contraprestación", lambda g, k=k: g["entrada"]["retenciones"].get(k, 0), None, False))
        filas_ret.append((f"   IVA retenido ({c['texto']})", lambda g, k=k, t=c["tasa"]: r2(g["entrada"]["retenciones"].get(k, 0) * t), None, False))
    filas_ret.append(("Total IVA retenido", V("retenido"), None, True))
    filas_ret.append(("IVA retenido según DIOT (referencia)", E("iva_retenido_diot"), None, False))
    hoja("Retenciones IVA", filas_ret)
    hoja("Sueldos y remanente", [
        ("Sueldos y salarios pagados", E("sueldos"), None, False),
        ("ISR retenido por sueldos", E("isr_sueldos"), None, True),
        ("Remanente distribuible a socios (asimilados)", E("remanente"), None, False),
        ("ISR retenido por remanente", E("isr_remanente"), None, True),
    ])
    acc = lambda k: (lambda g: g["resultado"]["accesorios"].get(k))
    hoja("Recargos", [
        ("¿Con recargos?", lambda g: "Sí" if g["resultado"]["accesorios"]["aplica"] else "No", "@", False),
        ("Fecha de vencimiento", acc("vencimiento"), "@", False),
        ("Fecha de pago", acc("fecha_pago"), "@", False),
        ("INPC mes anterior al pago", acc("inpc_pago"), "0.000", False),
        ("INPC mes anterior al vencimiento", acc("inpc_venc"), "0.000", False),
        ("Factor de actualización", acc("factor"), "0.0000", False),
        ("Meses de atraso", acc("meses"), "0", False),
    ])
    p = wb.create_sheet("Parámetros")
    p["A1"], p["A1"].font = "Parámetros", _FT
    filas = [("Ejercicio", cfg["ejercicio"]), ("Coeficiente de utilidad (enero en adelante)", cfg.get("coef_1")),
             (f"Coeficiente desde {MESES[int(cfg.get('mes_cambio_coef', 1)) - 1]}", cfg.get("coef_2")),
             ("Tasa ISR personas morales", TASA_ISR_PM), ("Tasa IVA", TASA_IVA),
             ("Tasa mensual de recargos", cfg.get("tasa_recargos"))]
    for i, (k, v) in enumerate(filas, start=3):
        p.cell(row=i, column=1, value=k).font = _F
        p.cell(row=i, column=2, value=v).font = _F
    p.cell(row=11, column=1, value="INPC (base 2Q jul 2018 = 100)").font = _FB
    for i, (k, v) in enumerate(sorted((cfg.get("inpc") or {}).items()), start=12):
        p.cell(row=i, column=1, value=k).font = _F
        p.cell(row=i, column=2, value=v).font = _F
    p.column_dimensions["A"].width = 48

    # Hoja oculta con los datos completos: es lo que el módulo lee el mes siguiente.
    h = wb.create_sheet(HOJA_DATOS)
    h["A1"] = "No modificar: datos del módulo Impuestos PM S.C."
    texto = json.dumps(doc, ensure_ascii=False, default=str)
    for i in range(0, len(texto), 30000):
        h.cell(row=2 + i // 30000, column=1, value=texto[i:i + 30000])
    h.sheet_state = "hidden"
    bio = io.BytesIO()
    wb.save(bio)
    return bio.getvalue()


def leer_papel(data: bytes) -> dict:
    """Lee un papel de trabajo generado por este módulo y devuelve el documento con todos los meses."""
    wb = load_workbook(io.BytesIO(data), read_only=True, data_only=True)
    if HOJA_DATOS not in wb.sheetnames:
        raise ValueError("Este Excel no es un papel de trabajo del módulo (no trae la hoja de datos). "
                         "Sube el archivo que descargaste del mes anterior.")
    partes = [r[0] for r in wb[HOJA_DATOS].iter_rows(min_row=2, max_col=1, values_only=True) if r and r[0]]
    doc = json.loads("".join(partes))
    doc.setdefault("meses", {})
    return doc


def recalcular(doc: dict, desde: int = 1) -> list[int]:
    """Recalcula en orden los meses del documento a partir de `desde` (usa el ISR pagado capturado)."""
    hechos = []
    for m in range(desde, 13):
        g = doc["meses"].get(str(m))
        if not g:
            continue
        g["resultado"] = calcular_mes(doc, m, g["entrada"])
        if g.get("isr_pagado_auto", True):
            g["isr_pagado"] = g["resultado"]["isr"]["a_pagar"]
        hechos.append(m)
    return hechos
