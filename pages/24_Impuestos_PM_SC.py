"""
pages/24_Impuestos_PM_SC.py — Impuestos Persona Moral · Sociedad Civil (v1, 2026-11-25)

Pagos provisionales mensuales de una S.C.:
  · Se carga la balanza y la DIOT del mes → se llenan los datos (editables).
  · El ISR usa los acumulados de los meses ya guardados (ingresos cobrados, remanente a socios,
    PTU, pagos provisionales efectivamente pagados).
  · Al guardar un mes, los meses posteriores ya guardados se recalculan solos.
Motor: _app/calc_impuestos_pm_sc.py · Guardado: _app/store_impuestos_pm.py
"""
import datetime as dt
import json

import pandas as pd
import streamlit as st

st.set_page_config(page_title="Impuestos PM S.C. | Auxiliar de Registros", page_icon="🏛️", layout="wide")

import _auth
import _theme
from _app import calc_impuestos_pm_sc as calc
from _app import store_impuestos_pm as store

_theme.aplicar_header(
    "🏛️ Impuestos PM — Sociedad Civil",
    "ISR art. 14 LISR (ingresos cobrados y remanente a socios) · IVA · Retenciones · Sueldos y asimilados",
)
_auth.require_login("Impuestos PM")

MESES = calc.MESES


def _money(v) -> str:
    return "—" if v is None else f"${v:,.2f}"


def _tabla(filas, col_valor="Importe"):
    """Tabla de dos columnas con importes ya formateados como texto."""
    return pd.DataFrame(
        [(c, "" if v is None else (v if isinstance(v, str) else f"{v:,.2f}")) for c, v in filas],
        columns=["Concepto", col_valor])


@st.cache_resource(ttl=300)
def _almacen():
    try:
        return store.Almacen(st.secrets)
    except Exception:
        return store.Almacen(None)


ALM = _almacen()

# ══════════════════════════════════════════════════════════════════════════════
# 1. CONTRIBUYENTE
# ══════════════════════════════════════════════════════════════════════════════
c1, c2, c3, c4 = st.columns([1.3, 2.6, 0.9, 1.4])
with c1:
    rfc = st.text_input("RFC", value=st.session_state.get("pm_rfc", "QUA031205AR2"), max_chars=13).strip().upper()
with c2:
    razon = st.text_input("Razón social", value=st.session_state.get("pm_razon", "QUALOM SC")).strip()
with c3:
    ejercicio = int(st.number_input("Ejercicio", 2020, 2035, value=dt.date.today().year, step=1))
with c4:
    st.caption("Guardado")
    if ALM.modo == "supabase":
        st.success("☁️ Supabase (en línea)", icon="✅")
    else:
        st.info("💻 Archivo local", icon="💾")
st.session_state.pm_rfc, st.session_state.pm_razon = rfc, razon

if not rfc or len(rfc) < 12:
    st.warning("Captura un RFC válido para continuar.")
    st.stop()

_KEY_DOC = f"pm_doc_{rfc}_{ejercicio}"
if _KEY_DOC not in st.session_state:
    try:
        d = ALM.cargar(rfc, ejercicio)
    except Exception as e:
        st.error(f"No pude leer lo guardado: {e}")
        d = None
    st.session_state[_KEY_DOC] = d or {"rfc": rfc, "razon_social": razon,
                                       "config": calc.config_default(ejercicio), "meses": {}}
DOC = st.session_state[_KEY_DOC]
DOC["razon_social"] = razon or DOC.get("razon_social", "")
CFG = DOC["config"]


def _guardar_doc():
    try:
        ALM.guardar(DOC)
        return True
    except Exception as e:
        st.error(f"❌ No se pudo guardar: {e}")
        return False


def _recalcular_desde(mes_ini: int) -> list[int]:
    """Recalcula los meses guardados posteriores a mes_ini con sus entradas guardadas."""
    hechos = []
    for m in range(mes_ini + 1, 13):
        g = DOC["meses"].get(str(m))
        if not g:
            continue
        nuevo = calc.calcular_mes(DOC, m, g["entrada"])
        if g.get("isr_pagado_auto", True):
            g["isr_pagado"] = nuevo["isr"]["a_pagar"]
        g["resultado"] = nuevo
        hechos.append(m)
    return hechos


guardados = sorted(int(k) for k in DOC["meses"])
chips = " ".join(
    f"<span style='padding:3px 10px;border-radius:12px;margin-right:4px;font-size:.8rem;"
    f"background:{'#0891B2' if (i + 1) in guardados else '#E2E8F0'};"
    f"color:{'#fff' if (i + 1) in guardados else '#64748B'}'>{m[:3]}</span>"
    for i, m in enumerate(MESES))
st.markdown(f"**Meses guardados {ejercicio}:** {chips}", unsafe_allow_html=True)
if CFG.get("coef_2", 0) == 0 and CFG.get("coef_1", 0) == 0:
    st.warning("⚙️ Falta capturar el **coeficiente de utilidad** en la pestaña Configuración.")

tab_mes, tab_anual, tab_cfg = st.tabs(["📅 Calcular mes", "📊 Resumen anual", "⚙️ Configuración"])

# ══════════════════════════════════════════════════════════════════════════════
# 2. CALCULAR MES
# ══════════════════════════════════════════════════════════════════════════════
def _panel_mes():
    sig = next((m for m in range(1, 13) if m not in guardados), 12)
    mes = st.selectbox("Mes a calcular", list(range(1, 13)), index=sig - 1,
                       format_func=lambda m: f"{MESES[m - 1]}{'  ✔ guardado' if m in guardados else ''}")
    K = f"pm_in_{rfc}_{ejercicio}_{mes}"
    VER = f"{K}_ver"
    st.session_state.setdefault(VER, 0)

    prev = calc.acumulados_previos(DOC, mes)
    if prev["faltan"]:
        faltan = ", ".join(MESES[m - 1] for m in prev["faltan"])
        if DOC.get("acumulados_iniciales"):
            st.info(f"Meses sin guardar ({faltan}): se usan los **acumulados iniciales** de Configuración.")
        else:
            st.warning(f"⚠️ Faltan meses anteriores: **{faltan}**. Guárdalos primero, o captura los acumulados "
                       "iniciales en Configuración si empiezas a mitad de año. Sin eso, el ISR sale incompleto.")
    with st.expander(f"📥 Lo que se lee de los meses guardados (enero a {MESES[mes - 2] if mes > 1 else '—'})",
                     expanded=False):
        st.dataframe(_tabla([
            ("Ingresos cobrados acumulados", prev["ingresos"]),
            ("Remanente a socios acumulado", prev["anticipos"]),
            ("PTU pagada acumulada", prev["ptu"]),
            ("Pagos provisionales efectuados", prev["pagos_prov"]),
            ("ISR retenido por bancos acumulado", prev["ret_bancos"]),
        ]), hide_index=True, use_container_width=True)

    # ── Archivos ──────────────────────────────────────────────────────────────
    st.markdown("##### 1️⃣ Balanza y DIOT del mes")
    a1, a2, a3 = st.columns([2, 2, 1])
    with a1:
        f_bal = st.file_uploader("Balanza de comprobación (Excel)", type=["xlsx", "xlsm"], key=f"{K}_bal")
    with a2:
        f_diot = st.file_uploader("Movimientos para DIOT (Excel)", type=["xlsx", "xlsm"], key=f"{K}_diot")
    with a3:
        st.write("")
        st.write("")
        leer = st.button("📖 Leer archivos", type="primary", use_container_width=True, disabled=not f_bal)

    if leer and f_bal:
        try:
            bal = calc.leer_balanza(f_bal.getvalue())
            diot = calc.leer_diot(f_diot.getvalue()) if f_diot else None
            datos, avisos = calc.extraer_mes(bal, diot)
            per = str(bal.get("periodo") or "")
            if per and per != f"{ejercicio}{mes:02d}":
                avisos.insert(0, f"La balanza es del periodo {per} y estás calculando {MESES[mes - 1]} {ejercicio}.")
            if bal.get("rfc") and bal["rfc"].upper() != rfc:
                avisos.insert(0, f"La balanza es del RFC {bal['rfc']}, no de {rfc}.")
            if diot and diot["periodos"] and int(diot["periodos"][0]) != int(f"{ejercicio}{mes:02d}"):
                avisos.insert(0, f"La DIOT trae el periodo {diot['periodos']}.")
            datos["recargos"] = {"aplica": False, "fecha_pago": None, "vencimiento": calc.vencimiento(ejercicio, mes).isoformat()}
            datos["perdidas"] = 0.0
            st.session_state[K] = datos
            st.session_state[f"{K}_avisos"] = avisos
            st.session_state[f"{K}_diot"] = diot["proveedores"] if diot else []
            st.session_state[VER] += 1
        except Exception as e:
            st.error(f"❌ No pude leer los archivos: {e}")

    if K not in st.session_state and str(mes) in DOC["meses"]:
        st.session_state[K] = json.loads(json.dumps(DOC["meses"][str(mes)]["entrada"]))
    if K not in st.session_state:
        st.info("Sube la balanza (y la DIOT) del mes y presiona **Leer archivos**.")
        return

    E = st.session_state[K]
    for a in st.session_state.get(f"{K}_avisos", []):
        st.warning(a)
    v = st.session_state[VER]

    def num(label, campo, help=None, cont=None, dct=None):
        dct = E if dct is None else dct
        tgt = cont or st
        dct[campo] = calc.r2(tgt.number_input(label, value=float(dct.get(campo, 0) or 0), step=100.0,
                                              format="%.2f", key=f"{K}_{campo}_{v}", help=help))

    # ── Captura editable ──────────────────────────────────────────────────────
    st.markdown("##### 2️⃣ Datos del mes (puedes corregirlos)")
    t1, t2, t3, t4 = st.tabs(["Ingresos / ISR", "IVA", "Retenciones de IVA", "Sueldos, remanente y recargos"])
    with t1:
        x1, x2, x3 = st.columns(3)
        num("Ingresos facturados (401)", "facturado", cont=x1)
        num("Ingresos efectivamente cobrados (208-01 ÷ 16%)", "cobrado", cont=x2)
        num("Devoluciones, descuentos y bonificaciones", "devoluciones", cont=x3)
        x1, x2, x3 = st.columns(3)
        num("PTU pagada en el mes", "ptu_pagada", cont=x1)
        num("ISR retenido por bancos en el mes", "isr_ret_bancos", cont=x2)
        num("Pérdidas fiscales por amortizar (acumulado)", "perdidas", cont=x3)
    with t2:
        x1, x2, x3 = st.columns(3)
        num("IVA trasladado cobrado (208-01)", "iva_trasladado", cont=x1)
        num("Gastos con IVA pagados (DIOT)", "gastos_16", cont=x2)
        num("IVA acreditable (DIOT)", "iva_acreditable", cont=x3)
        x1, x2, x3 = st.columns(3)
        num("Actos tasa 0% cobrados", "actos_0", cont=x1)
        num("Actos exentos cobrados", "actos_exentos", cont=x2)
        num("Saldo a favor de IVA aplicado", "saldo_favor_aplicado", cont=x3,
            help=f"Saldo a favor disponible en balanza (113-01): {_money(E.get('saldo_favor_iva_disponible'))}")
    with t3:
        st.caption("Captura el **valor de la contraprestación pagada (sin IVA)** por concepto; el IVA retenido se calcula con la tasa de ley.")
        cols = st.columns(2)
        for i, c in enumerate(calc.CONCEPTOS_RET_IVA):
            num(f"{c['nombre']} — {c['texto']}", c["clave"], cont=cols[i % 2], dct=E["retenciones"])
        provs = st.session_state.get(f"{K}_diot") or []
        if provs:
            with st.expander("Proveedores con IVA retenido en la DIOT"):
                st.dataframe(pd.DataFrame(provs).rename(columns={
                    "rfc": "RFC", "nombre": "Proveedor", "base": "Valor", "iva": "IVA",
                    "retenido": "IVA retenido", "concepto": "Concepto"}).style.format(
                    {"Valor": "{:,.2f}", "IVA": "{:,.2f}", "IVA retenido": "{:,.2f}"}),
                    hide_index=True, use_container_width=True)
        if E.get("iva_retenido_diot"):
            st.caption(f"IVA retenido según DIOT: {_money(E['iva_retenido_diot'])}")
    with t4:
        x1, x2 = st.columns(2)
        num("Sueldos y salarios pagados", "sueldos", cont=x1)
        num("ISR retenido por sueldos (según balanza)", "isr_sueldos", cont=x2)
        x1, x2 = st.columns(2)
        num("Remanente distribuible pagado a socios (bruto)", "remanente", cont=x1)
        num("ISR retenido por remanente (según balanza)", "isr_remanente", cont=x2)
        st.divider()
        R = E.setdefault("recargos", {"aplica": False, "fecha_pago": None,
                                      "vencimiento": calc.vencimiento(ejercicio, mes).isoformat()})
        y1, y2, y3 = st.columns([1, 1, 1])
        R["aplica"] = y1.toggle("¿La declaración va con recargos?", value=bool(R.get("aplica")), key=f"{K}_rec_{v}")
        venc_def = dt.date.fromisoformat(R.get("vencimiento") or calc.vencimiento(ejercicio, mes).isoformat())
        R["vencimiento"] = y2.date_input("Fecha de vencimiento", value=venc_def, format="DD/MM/YYYY",
                                         key=f"{K}_venc_{v}").isoformat()
        if R["aplica"]:
            fp_def = dt.date.fromisoformat(R["fecha_pago"]) if R.get("fecha_pago") else dt.date.today()
            R["fecha_pago"] = y3.date_input("Fecha de pago", value=fp_def, format="DD/MM/YYYY",
                                            key=f"{K}_fp_{v}").isoformat()

    # ── Cálculo ───────────────────────────────────────────────────────────────
    RES = calc.calcular_mes(DOC, mes, E)
    I, V = RES["isr"], RES["iva"]
    st.markdown(f"##### 3️⃣ Impuestos de {MESES[mes - 1]} {ejercicio}")
    m1, m2, m3, m4, m5 = st.columns(5)
    m1.metric("Total a pagar", _money(RES["total"]))
    m2.metric("ISR PM", _money(I["a_pagar"]))
    m3.metric("IVA (incl. retenido)", _money(V["total"]))
    m4.metric("ISR sueldos", _money(E.get("isr_sueldos", 0)))
    m5.metric("ISR remanente", _money(E.get("isr_remanente", 0)))

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

    # ── Guardar ───────────────────────────────────────────────────────────────
    st.markdown("##### 4️⃣ Guardar el mes")
    g_prev = DOC["meses"].get(str(mes), {})
    auto = g_prev.get("isr_pagado_auto", True)
    s1, s2, s3 = st.columns([1.4, 1.2, 1])
    with s1:
        manual = st.checkbox("El ISR provisional declarado/pagado fue distinto al calculado", value=not auto,
                             key=f"{K}_manual_{v}",
                             help="Los meses siguientes restan lo efectivamente pagado como pagos provisionales anteriores.")
    with s2:
        pagado = st.number_input("ISR provisional pagado", value=float(g_prev.get("isr_pagado", I["a_pagar"]) if manual else I["a_pagar"]),
                                 step=1.0, format="%.2f", disabled=not manual, key=f"{K}_pagado_{v}_{manual}")
    with s3:
        st.write("")
        if st.button(f"💾 Guardar {MESES[mes - 1]}", type="primary", use_container_width=True):
            DOC["meses"][str(mes)] = {"entrada": json.loads(json.dumps(E, default=str)), "resultado": RES,
                                      "isr_pagado": calc.r2(pagado if manual else I["a_pagar"]),
                                      "isr_pagado_auto": not manual,
                                      "guardado": dt.datetime.now().strftime("%d/%m/%Y %H:%M")}
            rehechos = _recalcular_desde(mes)
            if _guardar_doc():
                msg = f"✅ {MESES[mes - 1]} guardado."
                if rehechos:
                    msg += " Se recalcularon: " + ", ".join(MESES[m - 1] for m in rehechos) + "."
                st.success(msg)
                st.rerun()

with tab_mes:
    _panel_mes()

# ══════════════════════════════════════════════════════════════════════════════
# 3. RESUMEN ANUAL
# ══════════════════════════════════════════════════════════════════════════════
with tab_anual:
    if not DOC["meses"]:
        st.info("Aún no hay meses guardados.")
    else:
        filas = []
        for m in sorted(int(k) for k in DOC["meses"]):
            g = DOC["meses"][str(m)]
            rs = g["resultado"]
            filas.append({
                "Mes": MESES[m - 1],
                "Ingresos cobrados": rs["isr"]["cobrado"],
                "Remanente socios": rs["isr"]["remanente_mes"],
                "ISR PM": rs["isr"]["a_pagar"],
                "ISR pagado": g.get("isr_pagado"),
                "IVA a cargo": max(rs["iva"]["a_cargo"], 0),
                "IVA retenido": rs["iva"]["retenido"],
                "ISR sueldos": rs["resumen"][2]["a_cargo"],
                "ISR remanente": rs["resumen"][3]["a_cargo"],
                "Act. y recargos": calc.r2(sum(x["actualizacion"] + x["recargos"] for x in rs["resumen"])),
                "Total": rs["total"],
            })
        df = pd.DataFrame(filas)
        num_cols = [c for c in df.columns if c != "Mes"]
        df[num_cols] = df[num_cols].apply(pd.to_numeric, errors="coerce")
        tot = {"Mes": "TOTAL", **{c: df[c].sum() for c in num_cols}}
        df = pd.concat([df, pd.DataFrame([tot])], ignore_index=True)
        st.dataframe(df.style.format({c: "{:,.2f}" for c in num_cols}, na_rep="—"), hide_index=True, use_container_width=True)
        b1, b2, b3 = st.columns([1.3, 1.3, 2])
        with b1:
            st.download_button("📥 Papel de trabajo (Excel)", calc.generar_excel(DOC),
                               file_name=f"Impuestos_PM_{rfc}_{ejercicio}.xlsx", use_container_width=True,
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        with b2:
            st.download_button("🗂️ Respaldo (JSON)", json.dumps(store.limpiar(DOC), ensure_ascii=False, indent=1),
                               file_name=f"Impuestos_PM_{rfc}_{ejercicio}.json", use_container_width=True)
        with b3:
            borrar = st.selectbox("Borrar un mes guardado", ["—"] + [MESES[int(k) - 1] for k in sorted(DOC["meses"], key=int)])
            if borrar != "—" and st.button(f"🗑️ Borrar {borrar}"):
                mb = MESES.index(borrar) + 1
                DOC["meses"].pop(str(mb), None)
                _recalcular_desde(mb)
                if _guardar_doc():
                    st.success(f"{borrar} borrado; los meses posteriores se recalcularon.")
                    st.rerun()

# ══════════════════════════════════════════════════════════════════════════════
# 4. CONFIGURACIÓN
# ══════════════════════════════════════════════════════════════════════════════
with tab_cfg:
    st.markdown("**Coeficiente de utilidad**")
    k1, k2, k3, k4 = st.columns(4)
    coef_1 = k1.number_input("Coeficiente (enero en adelante)", value=float(CFG.get("coef_1", 0)), format="%.4f", step=0.0001)
    coef_2 = k2.number_input("Coeficiente de la última anual", value=float(CFG.get("coef_2", 0)), format="%.4f", step=0.0001)
    mes_cambio = k3.selectbox("Aplica desde el mes", list(range(1, 13)), index=int(CFG.get("mes_cambio_coef", 1)) - 1,
                              format_func=lambda m: MESES[m - 1])
    tasa_rec = k4.number_input("Tasa mensual de recargos", value=float(CFG.get("tasa_recargos", 0.0207)),
                               format="%.4f", step=0.0001)
    st.caption("Ejemplo 2026: 0.2663 (anual 2024) de enero a mayo y 0.3807 (anual 2025, presentada en julio) desde junio.")

    st.markdown("**INPC** (base 2Q julio 2018 = 100) — agrega cada mes cuando lo publique el INEGI")
    inpc_df = pd.DataFrame(sorted((CFG.get("inpc") or {}).items()), columns=["Mes (AAAA-MM)", "INPC"])
    inpc_ed = st.data_editor(inpc_df, num_rows="dynamic", hide_index=True, use_container_width=False, key=f"pm_inpc_{rfc}_{ejercicio}")

    with st.expander("Acumulados iniciales (solo si empiezas a capturar a mitad de año)"):
        ini = DOC.get("acumulados_iniciales") or {}
        st.caption("Totales de los meses que **no** vas a guardar en el módulo. Se usan solo cuando falta algún mes anterior.")
        j1, j2, j3, j4, j5 = st.columns(5)
        ini_v = {
            "ingresos": j1.number_input("Ingresos cobrados", value=float(ini.get("ingresos", 0)), format="%.2f"),
            "anticipos": j2.number_input("Remanente a socios", value=float(ini.get("anticipos", 0)), format="%.2f"),
            "ptu": j3.number_input("PTU pagada", value=float(ini.get("ptu", 0)), format="%.2f"),
            "pagos_prov": j4.number_input("Pagos provisionales", value=float(ini.get("pagos_prov", 0)), format="%.2f"),
            "ret_bancos": j5.number_input("ISR retenido bancos", value=float(ini.get("ret_bancos", 0)), format="%.2f"),
        }

    if st.button("💾 Guardar configuración", type="primary"):
        CFG.update(coef_1=calc.r2(coef_1 * 10000) / 10000, coef_2=calc.r2(coef_2 * 10000) / 10000,
                   mes_cambio_coef=int(mes_cambio), tasa_recargos=float(tasa_rec))
        tabla = {}
        for _, row in inpc_ed.dropna().iterrows():
            k = str(row["Mes (AAAA-MM)"]).strip()
            if len(k) == 7 and k[4] == "-":
                tabla[k] = float(row["INPC"])
        CFG["inpc"] = tabla
        DOC["acumulados_iniciales"] = ini_v if any(ini_v.values()) else {}
        rehechos = _recalcular_desde(0)
        if _guardar_doc():
            st.success("✅ Configuración guardada." + (" Se recalcularon los meses guardados." if rehechos else ""))
            st.rerun()

    st.divider()
    st.markdown("**Respaldo**")
    up = st.file_uploader("Restaurar desde un respaldo JSON", type=["json"], key=f"pm_restore_{rfc}_{ejercicio}")
    if up and st.button("♻️ Restaurar respaldo"):
        try:
            nuevo = json.loads(up.getvalue())
            if nuevo.get("rfc", "").upper() != rfc or int(nuevo["config"]["ejercicio"]) != ejercicio:
                st.error("El respaldo es de otro RFC o ejercicio.")
            else:
                st.session_state[_KEY_DOC] = nuevo
                DOC = nuevo
                if _guardar_doc():
                    st.success("✅ Respaldo restaurado.")
                    st.rerun()
        except Exception as e:
            st.error(f"Respaldo inválido: {e}")

    if ALM.modo == "local":
        st.info("💾 Se guarda en un archivo local de la app. En la versión en línea (Streamlit Cloud) ese archivo se "
                "pierde al reiniciar: para guardar en línea crea la tabla en Supabase con este SQL y descarga "
                "respaldos JSON de vez en cuando.")
        st.code(store.SQL_TABLA, language="sql")
        if ALM.error:
            st.caption(f"Supabase: {ALM.error}")
