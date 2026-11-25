"""
pages/24_Impuestos_PM_SC.py — Impuestos Persona Moral · Sociedad Civil (v2, 2026-11-25)

Sin guardar nada en la app: el papel de trabajo en Excel es el que lleva la información.
  · Se sube el papel de trabajo del mes anterior (el Excel que generó este módulo),
    la balanza y la DIOT del mes nuevo.
  · El ISR usa los acumulados que trae el Excel anterior (ingresos cobrados, remanente a socios,
    PTU, pagos provisionales efectivamente pagados).
  · Se descarga un Excel nuevo con todos los meses + el mes calculado; ese es el que se sube el mes siguiente.
Motor: _app/calc_impuestos_pm_sc.py
"""
import copy
import datetime as dt
import json

import pandas as pd
import streamlit as st

st.set_page_config(page_title="Impuestos PM S.C. | Auxiliar de Registros", page_icon="🏛️", layout="wide")

import _theme
from _app import calc_impuestos_pm_sc as calc

_theme.aplicar_header(
    "🏛️ Impuestos PM — Sociedad Civil",
    "Sube el papel de trabajo del mes anterior + balanza y DIOT del mes → descarga el papel de trabajo actualizado",
)

MESES = calc.MESES


def _money(v) -> str:
    return "—" if v is None else f"${v:,.2f}"


def _tabla(filas, col_valor="Importe"):
    """Tabla de dos columnas con importes ya formateados como texto."""
    return pd.DataFrame(
        [(c, "" if v is None else (v if isinstance(v, str) else f"{v:,.2f}")) for c, v in filas],
        columns=["Concepto", col_valor])


# ══════════════════════════════════════════════════════════════════════════════
# 1. ARCHIVOS
# ══════════════════════════════════════════════════════════════════════════════
st.markdown("##### 1️⃣ Archivos")
a1, a2, a3 = st.columns(3)
with a1:
    f_prev = st.file_uploader("📒 Papel de trabajo del mes anterior (Excel de este módulo)",
                              type=["xlsx"], key="pm2_prev",
                              help="Déjalo vacío si es el primer mes que calculas (enero) o si empiezas a usar el módulo.")
with a2:
    f_bal = st.file_uploader("📗 Balanza de comprobación del mes", type=["xlsx", "xlsm"], key="pm2_bal")
with a3:
    f_diot = st.file_uploader("📘 Movimientos para DIOT del mes", type=["xlsx", "xlsm"], key="pm2_diot")

if st.button("📖 Leer archivos", type="primary", disabled=not f_bal):
    try:
        bal = calc.leer_balanza(f_bal.getvalue())
        diot = calc.leer_diot(f_diot.getvalue()) if f_diot else None
        per = str(bal.get("periodo") or "")
        if not (len(per) == 6 and per.isdigit()):
            raise ValueError("No encontré el periodo (Mes inicial: AAAAMM) en la balanza.")
        ejercicio, mes = int(per[:4]), int(per[4:])
        if f_prev:
            doc = calc.leer_papel(f_prev.getvalue())
            if int(doc["config"]["ejercicio"]) != ejercicio:
                raise ValueError(f"El papel de trabajo es del ejercicio {doc['config']['ejercicio']} y la balanza de {ejercicio}.")
            if bal.get("rfc") and doc.get("rfc") and bal["rfc"].upper() != doc["rfc"].upper():
                raise ValueError(f"El papel de trabajo es de {doc['rfc']} y la balanza de {bal['rfc']}.")
        else:
            doc = {"rfc": (bal.get("rfc") or "").upper(), "razon_social": bal.get("empresa", ""),
                   "config": calc.config_default(ejercicio), "meses": {}}
        datos, avisos = calc.extraer_mes(bal, diot)
        if diot and diot["periodos"] and int(diot["periodos"][0]) != int(per):
            avisos.insert(0, f"La DIOT trae el periodo {diot['periodos']} y la balanza {per}.")
        if str(mes) in doc["meses"]:
            avisos.insert(0, f"El papel de trabajo ya traía {MESES[mes - 1]}: se reemplaza con esta balanza.")
        datos["recargos"] = {"aplica": False, "fecha_pago": None,
                             "vencimiento": calc.vencimiento(ejercicio, mes).isoformat()}
        datos["perdidas"] = 0.0
        st.session_state.pm2 = {"doc": doc, "mes": mes, "entrada": datos, "avisos": avisos,
                                "provs": diot["proveedores"] if diot else [], "con_prev": bool(f_prev)}
        st.session_state.pm2_ver = st.session_state.get("pm2_ver", 0) + 1
    except Exception as e:
        st.error(f"❌ {e}")


def _panel():
    S = st.session_state.get("pm2")
    if not S:
        st.info("Sube la **balanza** y la **DIOT** del mes y, a partir de febrero, el **papel de trabajo** que "
                "descargaste el mes anterior. Luego presiona **Leer archivos**.")
        return
    DOC, mes, E = S["doc"], S["mes"], S["entrada"]
    CFG = DOC["config"]
    ejercicio = int(CFG["ejercicio"])
    v = st.session_state.get("pm2_ver", 0)
    K = f"pm2_{v}"

    st.markdown(f"#### {DOC.get('razon_social') or ''} · {DOC.get('rfc') or ''} · **{MESES[mes - 1]} {ejercicio}**")
    previos = sorted(int(k) for k in DOC["meses"] if int(k) < mes)
    faltan = [m for m in range(1, mes) if m not in previos]
    for a in S["avisos"]:
        st.warning(a)
    if faltan:
        txt = ", ".join(MESES[m - 1] for m in faltan)
        if S["con_prev"]:
            st.warning(f"⚠️ El papel de trabajo no trae: **{txt}**. Captura esos acumulados en Configuración o el ISR saldrá incompleto.")
        else:
            st.warning(f"⚠️ No subiste papel de trabajo anterior. Para que el ISR salga completo captura en "
                       f"**Configuración → Acumulados iniciales** los totales de {txt}.")

    def num(label, campo, help=None, cont=None, dct=None, fmt="%.2f", step=100.0):
        dct = E if dct is None else dct
        tgt = cont or st
        val = tgt.number_input(label, value=float(dct.get(campo, 0) or 0), step=step,
                               format=fmt, key=f"{K}_{campo}", help=help)
        dct[campo] = val if fmt != "%.2f" else calc.r2(val)

    # ── Configuración (viaja dentro del Excel) ────────────────────────────────
    with st.expander("⚙️ Configuración — coeficiente de utilidad, recargos, INPC y acumulados iniciales",
                     expanded=not S["con_prev"]):
        k1, k2, k3, k4 = st.columns(4)
        CFG["coef_1"] = k1.number_input("Coeficiente (enero en adelante)", value=float(CFG.get("coef_1", 0)),
                                        format="%.4f", step=0.0001, key=f"{K}_c1")
        CFG["coef_2"] = k2.number_input("Coeficiente de la última anual", value=float(CFG.get("coef_2", 0)),
                                        format="%.4f", step=0.0001, key=f"{K}_c2")
        CFG["mes_cambio_coef"] = k3.selectbox("Aplica desde", list(range(1, 13)),
                                              index=int(CFG.get("mes_cambio_coef", 1)) - 1,
                                              format_func=lambda m: MESES[m - 1], key=f"{K}_mc")
        CFG["tasa_recargos"] = k4.number_input("Tasa mensual de recargos", value=float(CFG.get("tasa_recargos", 0.0207)),
                                               format="%.4f", step=0.0001, key=f"{K}_tr")
        st.caption("Ejemplo 2026: 0.2663 (anual 2024) y 0.3807 (anual 2025) desde junio. Si solo hay un coeficiente, pon el mismo en los dos.")
        st.markdown("**INPC** (base 2Q julio 2018 = 100) — agrega cada mes cuando lo publique el INEGI")
        inpc_df = pd.DataFrame(sorted((CFG.get("inpc") or {}).items()), columns=["Mes (AAAA-MM)", "INPC"])
        inpc_ed = st.data_editor(inpc_df, num_rows="dynamic", hide_index=True, key=f"{K}_inpc")
        tabla = {}
        for _, row in inpc_ed.dropna().iterrows():
            kk = str(row["Mes (AAAA-MM)"]).strip()
            if len(kk) == 7 and kk[4] == "-":
                tabla[kk] = float(row["INPC"])
        CFG["inpc"] = tabla
        st.markdown("**Acumulados iniciales** (solo para los meses que no trae el papel de trabajo)")
        ini = DOC.setdefault("acumulados_iniciales", {})
        j = st.columns(5)
        for i, (campo, label) in enumerate([("ingresos", "Ingresos cobrados"), ("anticipos", "Remanente a socios"),
                                            ("ptu", "PTU pagada"), ("pagos_prov", "Pagos provisionales"),
                                            ("ret_bancos", "ISR retenido bancos")]):
            num(label, campo, cont=j[i], dct=ini)
    if (CFG.get("coef_1") or 0) == 0 and (CFG.get("coef_2") or 0) == 0:
        st.error("Falta el **coeficiente de utilidad** (en Configuración).")

    # ── Meses anteriores: ISR efectivamente pagado ────────────────────────────
    if previos:
        with st.expander(f"📒 Meses que trae el papel de trabajo ({len(previos)}) — ISR provisional pagado", expanded=False):
            st.caption("Los meses siguientes restan como *pagos provisionales anteriores* lo que realmente pagaste. "
                       "Si declaraste algo distinto al calculado, corrígelo aquí.")
            filas = [{"Mes": MESES[m - 1],
                      "ISR calculado": DOC["meses"][str(m)]["resultado"]["isr"]["a_pagar"],
                      "ISR pagado": float(DOC["meses"][str(m)]["isr_pagado"] if DOC["meses"][str(m)].get("isr_pagado") is not None else DOC["meses"][str(m)]["resultado"]["isr"]["a_pagar"])} for m in previos]
            ed = st.data_editor(pd.DataFrame(filas), hide_index=True, disabled=["Mes", "ISR calculado"],
                                key=f"{K}_pag", column_config={
                                    "ISR calculado": st.column_config.NumberColumn(format="%.2f"),
                                    "ISR pagado": st.column_config.NumberColumn(format="%.2f")})
            for m, (_, row) in zip(previos, ed.iterrows()):
                g = DOC["meses"][str(m)]
                g["isr_pagado"] = calc.r2(row["ISR pagado"])
                g["isr_pagado_auto"] = abs(g["isr_pagado"] - g["resultado"]["isr"]["a_pagar"]) < 0.005

    # ── Captura del mes ───────────────────────────────────────────────────────
    st.markdown("##### 2️⃣ Datos del mes (puedes corregirlos)")
    t1, t2, t3, t4 = st.tabs(["Ingresos / ISR", "IVA", "Retenciones de IVA", "Sueldos, remanente y recargos"])
    with t1:
        x = st.columns(3)
        num("Ingresos facturados (401)", "facturado", cont=x[0])
        num("Ingresos efectivamente cobrados (208-01 ÷ 16%)", "cobrado", cont=x[1])
        num("Devoluciones, descuentos y bonificaciones", "devoluciones", cont=x[2])
        x = st.columns(3)
        num("PTU pagada en el mes", "ptu_pagada", cont=x[0])
        num("ISR retenido por bancos en el mes", "isr_ret_bancos", cont=x[1])
        num("Pérdidas fiscales por amortizar", "perdidas", cont=x[2])
    with t2:
        x = st.columns(3)
        num("IVA trasladado cobrado (208-01)", "iva_trasladado", cont=x[0])
        num("Gastos con IVA pagados (DIOT)", "gastos_16", cont=x[1])
        num("IVA acreditable (DIOT)", "iva_acreditable", cont=x[2])
        x = st.columns(3)
        num("Actos tasa 0% cobrados", "actos_0", cont=x[0])
        num("Actos exentos cobrados", "actos_exentos", cont=x[1])
        num("Saldo a favor de IVA aplicado", "saldo_favor_aplicado", cont=x[2],
            help=f"Saldo a favor en balanza (113-01): {_money(E.get('saldo_favor_iva_disponible'))}")
        x = st.columns(3)
        hay_ant = str(mes - 1) in DOC["meses"]
        E.setdefault("iva_ret_mes_anterior", 0.0)
        E["iva_ret_ant_auto"] = x[0].checkbox("I.V.A. retenido del mes anterior: tomarlo del papel de trabajo",
                                              value=hay_ant and E.get("iva_ret_ant_auto", True),
                                              disabled=not hay_ant, key=f"{K}_retant_auto")
        if E["iva_ret_ant_auto"] and hay_ant:
            x[1].metric("I.V.A. retenido del mes anterior", _money(DOC["meses"][str(mes - 1)]["resultado"]["iva"]["retenido"]))
        else:
            num("I.V.A. retenido del mes anterior", "iva_ret_mes_anterior", cont=x[1],
                help=f"IVA retenido a terceros del mes anterior. Referencia: cargos de la 216-10 en este mes "
                     f"{_money(E.get('iva_ret_enterado_216_10'))}.")
    with t3:
        st.caption("Valor de la contraprestación pagada **sin IVA** por concepto; el IVA retenido se calcula con la tasa de ley.")
        cols = st.columns(2)
        for i, c in enumerate(calc.CONCEPTOS_RET_IVA):
            num(f"{c['nombre']} — {c['texto']}", c["clave"], cont=cols[i % 2], dct=E["retenciones"])
        if S["provs"]:
            with st.expander("Proveedores con IVA retenido en la DIOT"):
                st.dataframe(pd.DataFrame(S["provs"]).rename(columns={
                    "rfc": "RFC", "nombre": "Proveedor", "base": "Valor", "iva": "IVA",
                    "retenido": "IVA retenido", "concepto": "Concepto"}).style.format(
                    {"Valor": "{:,.2f}", "IVA": "{:,.2f}", "IVA retenido": "{:,.2f}"}),
                    hide_index=True, use_container_width=True)
        if E.get("iva_retenido_diot"):
            st.caption(f"IVA retenido según DIOT: {_money(E['iva_retenido_diot'])}")
    with t4:
        x = st.columns(2)
        num("Sueldos y salarios pagados", "sueldos", cont=x[0])
        num("ISR retenido por sueldos (según balanza)", "isr_sueldos", cont=x[1])
        x = st.columns(2)
        num("Remanente distribuible pagado a socios (bruto)", "remanente", cont=x[0])
        num("ISR retenido por remanente (según balanza)", "isr_remanente", cont=x[1])
        st.divider()
        R = E["recargos"]
        y = st.columns(3)
        R["aplica"] = y[0].toggle("¿La declaración va con recargos?", value=bool(R.get("aplica")), key=f"{K}_rec")
        venc = dt.date.fromisoformat(R.get("vencimiento") or calc.vencimiento(ejercicio, mes).isoformat())
        R["vencimiento"] = y[1].date_input("Fecha de vencimiento", value=venc, format="DD/MM/YYYY", key=f"{K}_venc").isoformat()
        if R["aplica"]:
            fp = dt.date.fromisoformat(R["fecha_pago"]) if R.get("fecha_pago") else dt.date.today()
            R["fecha_pago"] = y[2].date_input("Fecha de pago", value=fp, format="DD/MM/YYYY", key=f"{K}_fp").isoformat()

    # ── ISR pagado del mes ────────────────────────────────────────────────────
    s1, s2 = st.columns([1.4, 1])
    manual = s1.checkbox("El ISR provisional que declaré este mes es distinto al calculado", key=f"{K}_man")
    pagado = s2.number_input("ISR provisional declarado", value=0.0, step=1.0, format="%.2f",
                             disabled=not manual, key=f"{K}_pagado")

    # ── Cálculo (documento nuevo = anterior + este mes; se recalculan meses posteriores) ──
    nuevo = copy.deepcopy(DOC)
    nuevo["meses"][str(mes)] = {"entrada": json.loads(json.dumps(E, default=str)),
                                "isr_pagado": calc.r2(pagado) if manual else None,
                                "isr_pagado_auto": not manual,
                                "generado": dt.datetime.now().strftime("%d/%m/%Y %H:%M")}
    calc.recalcular(nuevo, 1)
    RES = nuevo["meses"][str(mes)]["resultado"]
    I, V = RES["isr"], RES["iva"]

    st.markdown(f"##### 3️⃣ Impuestos de {MESES[mes - 1]} {ejercicio}")
    m = st.columns(5)
    m[0].metric("Total a pagar", _money(RES["total"]))
    m[1].metric("ISR PM", _money(I["a_pagar"]))
    m[2].metric("IVA (incl. retenido)", _money(V["total"]))
    m[3].metric("ISR sueldos", _money(E.get("isr_sueldos", 0)))
    m[4].metric("ISR remanente", _money(E.get("isr_remanente", 0)))

    r1, r2c = st.columns(2)
    with r1:
        st.markdown("**ISR personas morales**")
        st.dataframe(_tabla([
            ("Ingresos facturados", I["facturado"]),
            ("(+) Cobranza de facturas de meses anteriores", I["aumento_cobranza"]),
            ("(−) Facturado en el mes pendiente de cobro", I["disminucion_no_cobrado"]),
            ("= Ingresos efectivamente cobrados", I["cobrado"]),
            ("(−) Devoluciones", I["devoluciones"]),
            ("(+) Ingresos de meses anteriores", I["ingresos_previos"]),
            ("Ingresos acumulados", I["ingresos_acum"]),
            ("× Coeficiente de utilidad", f"{I['coeficiente']:.4f}"),
            ("Utilidad fiscal estimada", I["utilidad"]),
            ("(−) PTU pagada", I["ptu_acum"]),
            (f"(−) Remanente a socios acumulado (mes: {_money(I['remanente_mes'])})", I["anticipos_acum"]),
            ("(−) Pérdidas fiscales", I["perdidas"]),
            ("Base", I["base"]),
            ("ISR causado 30%", I["causado"]),
            ("(−) Pagos provisionales anteriores", I["pagos_anteriores"]),
            ("(−) ISR retenido por bancos", I["ret_bancos_acum"]),
            ("ISR a pagar", I["a_pagar"]),
        ]), hide_index=True, use_container_width=True, height=630)
    with r2c:
        st.markdown("**IVA**")
        st.dataframe(_tabla([
            ("Actos gravados 16% cobrados", V["actos_16"]),
            ("IVA trasladado cobrado", V["trasladado"]),
            ("(−) IVA acreditable", V["acreditable"]),
            ("(−) Saldo a favor aplicado", V["saldo_favor_aplicado"]),
            ("(−) I.V.A. retenido del mes anterior", V["retenido_mes_anterior"]),
            ("IVA a cargo / (a favor)", V["a_cargo"]),
            ("(+) IVA retenido a terceros", V["retenido"]),
            ("Total IVA a enterar", V["total"]),
        ]), hide_index=True, use_container_width=True)
        if RES["retenciones"]:
            st.markdown("**Retenciones de IVA**")
            st.dataframe(pd.DataFrame(RES["retenciones"]).rename(columns={
                "concepto": "Concepto", "tasa": "Tasa", "valor": "Valor pagado",
                "iva_trasladado": "IVA trasladado", "iva_retenido": "IVA retenido"}).style.format(
                {"Valor pagado": "{:,.2f}", "IVA trasladado": "{:,.2f}", "IVA retenido": "{:,.2f}"}),
                hide_index=True, use_container_width=True)
        st.markdown("**Resumen de pago**")
        acc = RES["accesorios"]
        if acc["aplica"]:
            st.caption(f"Factor {acc['factor']:.4f} (INPC {acc['inpc_pago']:.3f} / {acc['inpc_venc']:.3f}) · "
                       f"{acc['meses']} mes(es) de atraso · tasa {CFG.get('tasa_recargos', 0):.2%} mensual")
        dfr = pd.DataFrame(RES["resumen"]).rename(columns={
            "impuesto": "Impuesto", "a_cargo": "A cargo", "actualizacion": "Actualización",
            "recargos": "Recargos", "total": "Total"})
        dfr.loc[len(dfr)] = ["TOTAL", dfr["A cargo"].sum(), dfr["Actualización"].sum(), dfr["Recargos"].sum(), dfr["Total"].sum()]
        st.dataframe(dfr.style.format({c: "{:,.2f}" for c in ["A cargo", "Actualización", "Recargos", "Total"]}),
                     hide_index=True, use_container_width=True)

    # ── Resumen del año y descarga ────────────────────────────────────────────
    st.markdown("##### 4️⃣ Papel de trabajo actualizado")
    filas = []
    for mm in sorted(int(k) for k in nuevo["meses"]):
        g = nuevo["meses"][str(mm)]
        rs = g["resultado"]
        filas.append({"Mes": MESES[mm - 1], "Ingresos cobrados": rs["isr"]["cobrado"],
                      "Remanente socios": rs["isr"]["remanente_mes"], "ISR PM": rs["isr"]["a_pagar"],
                      "ISR pagado": g.get("isr_pagado"), "IVA total": rs["iva"]["total"],
                      "ISR sueldos": rs["resumen"][2]["a_cargo"], "ISR remanente": rs["resumen"][3]["a_cargo"],
                      "Act. y recargos": calc.r2(sum(x["actualizacion"] + x["recargos"] for x in rs["resumen"])),
                      "Total": rs["total"]})
    df = pd.DataFrame(filas)
    nc = [c for c in df.columns if c != "Mes"]
    df[nc] = df[nc].apply(pd.to_numeric, errors="coerce")
    df = pd.concat([df, pd.DataFrame([{"Mes": "TOTAL", **{c: df[c].sum() for c in nc}}])], ignore_index=True)
    st.dataframe(df.style.format({c: "{:,.2f}" for c in nc}, na_rep="—"), hide_index=True, use_container_width=True)
    rfc = (nuevo.get("rfc") or "SC").upper()
    st.download_button(f"📥 Descargar papel de trabajo con {MESES[mes - 1]}", calc.generar_excel(nuevo),
                       file_name=f"Impuestos_PM_{rfc}_{ejercicio}_{mes:02d}_{MESES[mes - 1]}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")
    st.caption("Guarda este Excel: es el que subes como *papel de trabajo del mes anterior* el próximo mes. "
               "No borres ni edites la hoja oculta de datos.")


_panel()
