"""
pages/22_Depuracion_de_Cuentas.py — DEPURACIÓN DE CUENTAS (v1, 2026-11-18)

Auditoría del auxiliar de proveedores exportado del ERP:
  · Integridad aritmética (saldo corrido y saldo final)
  · Relación cargo-abono: pagos y notas de crédito ligados a su factura
  · Hallazgos (saldos deudores, NC sin factura, pagos duplicados, egresos manuales, catálogo, antigüedad…)
  · Descargas: auxiliar original marcado (amarillo / azul) y papel de trabajo en Excel
  · Comparación opcional contra un auxiliar anterior
Motor: _app/depuracion_cuentas.py
"""
import hashlib
import io

import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Depuración de Cuentas | Auxiliar de Registros",
    page_icon="🧹",
    layout="wide",
)

import _theme
from _app import depuracion_cuentas as dc

_theme.aplicar_header("🧹 DEPURACIÓN DE CUENTAS",
                      "Auditoría del auxiliar de proveedores: relación cargo-abono, hallazgos y papel de trabajo")

AMARILLO, AZUL = "#FFF59D", "#BBDEFB"


@st.cache_data(show_spinner=False, max_entries=6)
def _analizar(data: bytes, _key: str):
    return dc.analizar(data)


@st.cache_data(show_spinner=False, max_entries=6)
def _excel_marcado(data: bytes, _key: str) -> bytes:
    return dc.auxiliar_marcado(data, _analizar(data, _key))


@st.cache_data(show_spinner=False, max_entries=6)
def _conciliacion(data: bytes, _key: str):
    res = _analizar(data, _key)
    det = dc.detalle_conciliacion(res)
    det["folio"] = det.folio.fillna("")
    return dc.resumen_por_proveedor(res).reset_index(), det


@st.cache_data(show_spinner=False, max_entries=6)
def _excel_proveedor(data: bytes, _key: str) -> bytes:
    return dc.reporte_por_proveedor(_analizar(data, _key))


@st.cache_data(show_spinner=False, max_entries=6)
def _excel_papel(data: bytes, _key: str) -> bytes:
    return dc.papel_de_trabajo(_analizar(data, _key))


def _money(v) -> str:
    return "—" if v is None or pd.isna(v) else f"${v:,.2f}"


def _md(t: str) -> str:
    """Escapa $ para que Markdown no lo interprete como fórmula."""
    return str(t).replace("$", "\\$")


def _mill(v) -> str:
    return f"${v / 1e6:,.2f} M" if abs(v) >= 1e6 else f"${v:,.0f}"


def _d(v) -> str:
    return v.strftime("%d-%m-%Y") if hasattr(v, "strftime") else str(v or "")


def _sty(df: pd.DataFrame, money=(), color_col=None, color_val="Y"):
    """Styler con formato $ y miles (compatible con cualquier versión de Streamlit) y fondo amarillo/azul."""
    money = [c for c in money if c in df.columns]
    sty = df.style.format({c: (lambda v: "" if pd.isna(v) else f"${v:,.2f}") for c in money})
    if color_col:
        sty = sty.apply(lambda r: [f"background-color:{AMARILLO if r[color_col] == color_val else AZUL}"] * len(r), axis=1)
    return sty


def _money_cols(df: pd.DataFrame) -> list:
    return [c for c in df.columns if pd.api.types.is_float_dtype(df[c])]


# ── Carga ─────────────────────────────────────────────────────────────────────
with st.container(border=True):
    c1, c2 = st.columns([3, 2])
    with c1:
        up = st.file_uploader("Auxiliar de cuentas (Excel exportado del ERP, hoja **Datos**)", type=["xlsx"], key="dep_aux")
    with c2:
        up_prev = st.file_uploader("Opcional: auxiliar anterior para comparar", type=["xlsx"], key="dep_prev")
    st.caption("Se revisa el 100% de los movimientos. El archivo no se guarda: se procesa en memoria.")

if not up:
    st.info("Sube el auxiliar de proveedores (cuenta 201) para iniciar la depuración.")
    with st.expander("¿Qué revisa este módulo?"):
        st.markdown("""
- **Integridad:** recalcula el saldo corrido de cada movimiento y el saldo final de cada cuenta.
- **Relación cargo-abono:** cada pago (EA) se liga a su factura (PRO) por folio o UUID; cada nota de crédito
  se liga cuando viene aplicada dentro del pago o cubre exacto el saldo de una factura.
  Los pagos de facturas del ejercicio anterior se aceptan hasta el saldo inicial del proveedor.
- **Amarillo** = relacionado · **Azul** = sin relación (factura sin pago o parcial, NC sin factura,
  pago en exceso, egreso manual, importe cero).
- **Hallazgos:** saldos deudores, NC sin factura, pagos duplicados, egresos manuales, NC altas vs facturado,
  catálogo vs CFDI, pagos antes de la factura, saldos iniciales sin liquidar, antigüedad y concentración.
""")
    st.stop()

data = up.getvalue()
key = hashlib.md5(data).hexdigest()
try:
    with st.spinner("Analizando el auxiliar…"):
        res = _analizar(data, key)
except Exception as e:  # formato incorrecto
    st.error(f"No se pudo leer el archivo: {e}")
    st.stop()

meta, f, m, rel = res.meta, res.cuentas, res.mov, res.relacion
st.markdown(f"**{meta.get('Empresa', '')}** · RFC {meta.get('R. F. C.', '')} · "
            f"{_d(meta.get('Fecha inicial'))} a {_d(meta.get('Fecha final'))} · "
            f"cuentas {meta.get('Cuenta inicial', '')} a {next((v for k, v in meta.items() if k.startswith('Cuenta final')), '')}")

amar = (rel.color == "Y").sum()
k = st.columns(5)
k[0].metric("Cuentas", f"{len(f):,}")
k[1].metric("Movimientos", f"{len(m):,}")
k[2].metric("Saldo final", _mill(f.fin.sum()))
k[3].metric("Saldos deudores", _mill(f[f.fin < -0.005].fin.sum()), f"{(f.fin < -0.005).sum()} cuentas", delta_color="off")
k[4].metric("Relacionados (amarillo)", f"{amar / len(rel):.0%}", f"{len(rel) - amar:,} en azul", delta_color="off")

if res.errores_aritmeticos:
    st.error(f"⚠️ {res.errores_aritmeticos} diferencias aritméticas en saldos corridos o finales.")
else:
    st.success("✅ Integridad aritmética: todos los saldos corridos y finales coinciden al centavo.")

tabs = st.tabs(["📋 Hallazgos", "🔗 Relación cargo-abono", "🏢 Proveedores", "⚖️ Por proveedor", "⏳ Antigüedad",
                "🔄 Comparar", "⬇️ Descargas"])

# ── Hallazgos ─────────────────────────────────────────────────────────────────
with tabs[0]:
    ini, fac, pag, nc = f.ini.sum(), f.facturado.sum(), f.pagos.sum(), f.notas_credito.sum()
    cuadre = pd.DataFrame({
        "Concepto": ["Saldo inicial", "(+) Facturas", "(−) Pagos", "(−) Notas de crédito", "Saldo final",
                     "Saldos acreedores", "Saldos deudores"],
        "Importe": [ini, fac, -pag, -nc, f.fin.sum(), f[f.fin > 0.005].fin.sum(), f[f.fin < -0.005].fin.sum()],
    })
    cA, cB = st.columns([1, 2])
    with cA:
        st.markdown("##### Cuadre")
        st.dataframe(_sty(cuadre, ["Importe"]), hide_index=True, use_container_width=True)
    with cB:
        st.markdown("##### Hallazgos")
        ico = {"Alto": "🔴", "Medio": "🟠", "Bajo": "🟢"}
        for i, h in enumerate(res.hallazgos, 1):
            with st.expander(_md(f"{ico[h['riesgo']]} {i}. {h['titulo']}" + (f" — {_money(h['importe'])}" if h["importe"] else ""))):
                if h["detalle"]:
                    st.markdown(_md(h["detalle"]))
                st.markdown(_md(f"**Acción:** {h['accion']}"))
                tkey = {"Saldos deudores": "saldos_deudores", "NC sin factura": "nc_sin_factura",
                        "Pagos en exceso": "pagos_exceso", "Egresos manuales": "egresos_manuales",
                        "Pagos sin factura": "pagos_sin_factura", "NC por proveedor": "nc_por_proveedor",
                        "Catálogo": "catalogo", "Saldo deudor temporal": "saldo_deudor_temporal",
                        "Saldos iniciales": "saldos_iniciales"}.get(h["hoja"])
                if tkey and tkey in res.tablas and len(res.tablas[tkey]):
                    st.dataframe(_sty(res.tablas[tkey], _money_cols(res.tablas[tkey])), hide_index=True, use_container_width=True, height=min(400, 38 + 35 * len(res.tablas[tkey])))

# ── Relación cargo-abono ──────────────────────────────────────────────────────
with tabs[1]:
    rr = dc.resumen_relacion(res)
    st.dataframe(_sty(rr, ["cargos", "abonos"], "color", "Amarillo"), hide_index=True, use_container_width=True,
                 column_config={"color": "Color", "resultado": "Resultado", "renglones": "Renglones",
                                "cargos": "Cargos", "abonos": "Abonos"})
    c1, c2, c3 = st.columns([1, 2, 2])
    col = c1.radio("Color", ["Azul", "Amarillo", "Todos"], horizontal=True)
    cats = c2.multiselect("Resultado", sorted(rel.categoria.unique()))
    prov = c3.text_input("Buscar proveedor o cuenta")
    v = rel.copy()
    if col != "Todos":
        v = v[v.color == ("B" if col == "Azul" else "Y")]
    if cats:
        v = v[v.categoria.isin(cats)]
    if prov:
        q = prov.upper()
        v = v[v["Cta. Desc."].str.upper().str.contains(q, regex=False) | v["Cta."].str.contains(q, regex=False)]
    st.caption(f"{len(v):,} renglones · Renglón = fila del Excel original")
    show = v[["row", "Cta.", "Cta. Desc.", "Tipo póliza", "Num. Póliza", "Fecha", "Cargo", "Abono", "motivo", "color"]]
    st.dataframe(_sty(show, ["Cargo", "Abono"], "color"), hide_index=True, use_container_width=True, height=520,
                 column_config={"row": "Renglón", "Num. Póliza": st.column_config.NumberColumn("Póliza", format="%d"), "Fecha": st.column_config.DateColumn(format="DD/MM/YYYY"),
                                "motivo": "Relación", "color": None})

# ── Proveedores ───────────────────────────────────────────────────────────────
with tabs[2]:
    azul_cta = rel[rel.color == "B"].groupby("Cta.").size()
    tb = f.reset_index()[["cta", "nombre", "ini", "facturado", "pagos", "notas_credito", "fin", "n"]].copy()
    tb["en_azul"] = tb.cta.map(azul_cta).fillna(0).astype(int)
    solo = st.toggle("Solo saldos deudores", value=False)
    if solo:
        tb = tb[tb.fin < -0.005]
    st.dataframe(_sty(tb.sort_values("fin"), ["ini", "facturado", "pagos", "notas_credito", "fin"]),
                 hide_index=True, use_container_width=True, height=380,
                 column_config={"cta": "Cuenta", "nombre": "Proveedor", "n": "Movs.", "en_azul": "En azul",
                                "ini": "Saldo inicial", "facturado": "Facturado", "pagos": "Pagos",
                                "notas_credito": "Notas de crédito", "fin": "Saldo final"})
    opciones = [f"{c} · {n}" for c, n in zip(f.index, f.nombre)]
    sel = st.selectbox("Ver movimientos de la cuenta", opciones, index=None, placeholder="Elige una cuenta…")
    if sel:
        cta = sel.split(" · ")[0]
        mv = m[m["Cta."] == cta][["row", "Tipo póliza", "Num. Póliza", "Fecha", "Desc. Detalle", "Cargo", "Abono", "Saldo"]] \
            .merge(rel[["row", "color", "motivo"]], on="row")
        st.dataframe(_sty(mv, ["Cargo", "Abono", "Saldo"], "color"), hide_index=True, use_container_width=True,
                     column_config={"row": "Renglón", "Num. Póliza": st.column_config.NumberColumn("Póliza", format="%d"), "Fecha": st.column_config.DateColumn(format="DD/MM/YYYY"),
                                    "motivo": "Relación", "color": None})

# ── Por proveedor: lo que se cancela vs pendiente ────────────────────────────
with tabs[3]:
    with st.spinner("Agrupando movimientos que se cancelan…"):
        rp, det = _conciliacion(data, key)
    c = st.columns(4)
    c[0].metric("Cuentas conciliadas", f"{(rp.estatus == 'Conciliada').sum():,}")
    c[1].metric("Pendientes de conciliar", f"{(rp.estatus != 'Conciliada').sum():,}")
    c[2].metric("Movimientos pendientes", f"{int(rp.movs_pendientes.sum()):,}")
    c[3].metric("Saldo pendiente", _mill(rp.saldo_pendiente.sum()))
    st.caption("Saldo pendiente = saldo inicial pendiente + abonos pendientes − cargos pendientes. "
               "Abonos conciliados = cargos conciliados: esos movimientos se cancelan entre sí.")
    solo_p = st.toggle("Solo cuentas pendientes de conciliar", value=True)
    vista = rp[rp.estatus != "Conciliada"] if solo_p else rp
    cols_m = ["saldo_inicial", "cargos_saldo_inicial", "saldo_inicial_pendiente", "abonos_conciliados", "cargos_conciliados",
              "abonos_pendientes", "cargos_pendientes", "saldo_pendiente", "saldo_final", "diferencia"]
    st.dataframe(_sty(vista[["cta", "nombre"] + cols_m + ["movs_pendientes", "estatus"]], cols_m, "estatus", "Conciliada"),
                 hide_index=True, use_container_width=True, height=380,
                 column_config={"cta": "Cuenta", "nombre": "Proveedor", "saldo_inicial": "Saldo inicial",
                                "cargos_saldo_inicial": "Cargos a saldo inicial", "saldo_inicial_pendiente": "Saldo inicial pendiente",
                                "abonos_conciliados": "Abonos conciliados", "cargos_conciliados": "Cargos conciliados",
                                "abonos_pendientes": "Abonos pendientes", "cargos_pendientes": "Cargos pendientes",
                                "saldo_pendiente": "Saldo pendiente", "saldo_final": "Saldo final", "diferencia": "Diferencia",
                                "movs_pendientes": "Movs. pendientes", "estatus": "Estatus"})
    opc = [f"{a} · {b}" for a, b in zip(vista.cta, vista.nombre)]
    selp = st.selectbox("Ver pendientes y conciliados de la cuenta", opc, index=None, placeholder="Elige una cuenta…", key="dep_prov")
    if selp:
        cta = selp.split(" · ")[0]
        a = rp.set_index("cta").loc[cta]
        dd = det[det["Cta."] == cta]
        pend = dd[dd.grupo_estado == "Pendiente"]
        conc = dd[dd.grupo_estado.isin(["Se cancela", "Aplicado a saldo inicial"])]
        st.caption("Cada grupo junta la factura con los pagos y notas de crédito que la cancelan, en renglones consecutivos; "
                   "grupos del más reciente al más antiguo.")
        st.markdown(_md(f"**Saldo inicial pendiente:** {_money(a.saldo_inicial_pendiente)} · "
                        f"**Abonos pendientes:** {_money(a.abonos_pendientes)} · **Cargos pendientes:** {_money(a.cargos_pendientes)} · "
                        f"**Saldo pendiente:** {_money(a.saldo_pendiente)} (auxiliar {_money(a.saldo_final)})"))
        st.markdown("###### Pendientes de conciliar")
        st.dataframe(_sty(pend[["grupo", "grupo_tipo", "row", "Tipo póliza", "Num. Póliza", "Fecha", "folio", "Cargo", "Abono",
                                "cargo_pendiente", "abono_pendiente", "estado", "motivo"]],
                          ["Cargo", "Abono", "cargo_pendiente", "abono_pendiente"], "estado", "Conciliado"),
                     hide_index=True, use_container_width=True,
                     column_config={"grupo": "Grupo", "grupo_tipo": "Tipo de grupo", "row": "Renglón", "Num. Póliza": st.column_config.NumberColumn("Póliza", format="%d"), "Fecha": st.column_config.DateColumn(format="DD/MM/YYYY"), "folio": "Folio",
                                    "cargo_pendiente": "Cargo pendiente", "abono_pendiente": "Abono pendiente",
                                    "estado": "Estado", "motivo": "Motivo"})
        st.markdown(_md(f"###### Conciliados (se cancelan): abonos {_money(conc.abono_conciliado.sum())} · "
                        f"cargos {_money(conc.cargo_conciliado.sum())} · aplicados a saldo inicial {_money(conc.cargo_saldo_inicial.sum())}"))
        st.dataframe(_sty(conc[["grupo", "grupo_tipo", "row", "Tipo póliza", "Num. Póliza", "Fecha", "folio", "cargo_conciliado",
                                "cargo_saldo_inicial", "abono_conciliado", "motivo"]],
                          ["cargo_conciliado", "cargo_saldo_inicial", "abono_conciliado"]),
                     hide_index=True, use_container_width=True,
                     column_config={"grupo": "Grupo", "grupo_tipo": "Tipo de grupo", "row": "Renglón", "Num. Póliza": st.column_config.NumberColumn("Póliza", format="%d"), "Fecha": st.column_config.DateColumn(format="DD/MM/YYYY"), "folio": "Folio",
                                    "cargo_conciliado": "Cargo conciliado", "cargo_saldo_inicial": "Cargo a saldo inicial",
                                    "abono_conciliado": "Abono conciliado", "motivo": "Relación"})

# ── Antigüedad ────────────────────────────────────────────────────────────────
with tabs[4]:
    ag = res.tablas.get("antiguedad")
    if ag is None or not len(ag):
        st.info("No hay saldos acreedores.")
    else:
        rangos = ["0-30", "31-60", "61-90", "91-180", ">180 / saldo inicial"]
        tot = ag[rangos].sum()
        st.markdown(_md(f"##### Saldo acreedor {_money(ag.saldo.sum())} · más de 90 días: {_money(ag.mas_90.sum())} "
                        f"({ag.mas_90.sum() / ag.saldo.sum():.0%})"))
        st.bar_chart(pd.DataFrame({"Rango": rangos, "Importe": tot.values}).set_index("Rango"), horizontal=True, color="#0891B2")
        st.dataframe(_sty(ag.sort_values("mas_90", ascending=False), ["saldo", "mas_90"] + rangos),
                     hide_index=True, use_container_width=True, height=400,
                     column_config={"cta": "Cuenta", "nombre": "Proveedor", "saldo": "Saldo", "mas_90": "Más de 90 días"})
        st.caption("Método PEPS: el saldo de cada proveedor se asigna a sus facturas más recientes.")
    cc = res.tablas["concentracion"].copy()
    cc["pct"] = cc.facturado / f.facturado.sum() * 100
    cc["acum"] = cc.pct.cumsum()
    st.markdown(f"##### Concentración: los 10 proveedores principales son el {cc.pct.head(10).sum():.0f}% de lo facturado")
    st.dataframe(_sty(cc, ["facturado", "fin"]), hide_index=True, use_container_width=True,
                 column_config={"cta": "Cuenta", "nombre": "Proveedor", "facturado": "Facturado", "fin": "Saldo final",
                                "pct": st.column_config.NumberColumn("% del total", format="%.1f%%"),
                                "acum": st.column_config.NumberColumn("% acumulado", format="%.1f%%")})
    st.caption("% sobre lo facturado en el periodo.")

# ── Comparar contra auxiliar anterior ─────────────────────────────────────────
with tabs[5]:
    if not up_prev:
        st.info("Sube un auxiliar anterior (arriba, a la derecha) para comprobar que los movimientos ya registrados no cambiaron.")
    else:
        dp = up_prev.getvalue()
        try:
            prev = _analizar(dp, hashlib.md5(dp).hexdigest())
        except Exception as e:
            st.error(f"No se pudo leer el auxiliar anterior: {e}")
            prev = None
        if prev is not None:
            pm = prev.mov
            ult = pm.Periodo.max()
            claves = ["Cta.", "Tipo póliza", "Num. Póliza", "Periodo", "Fecha", "Cargo", "Abono", "Desc. Detalle", "Ren."]
            a = pm[claves]
            b = m[m.Periodo <= ult][claves]
            mg = a.merge(b, how="outer", indicator=True)
            quit_ = mg[mg._merge == "left_only"]
            nuevos = mg[mg._merge == "right_only"]
            c = st.columns(4)
            c[0].metric("Periodo comparado", f"hasta {int(ult)}")
            c[1].metric("Movimientos iguales", f"{(mg._merge == 'both').sum():,}")
            c[2].metric("Eliminados o modificados", f"{len(quit_):,}")
            c[3].metric("Agregados después", f"{len(nuevos):,}")
            if len(quit_) == 0 and len(nuevos) == 0:
                st.success("✅ Los movimientos del periodo anterior no cambiaron.")
            else:
                st.warning("Hay movimientos de periodos ya revisados que cambiaron.")
                st.dataframe(pd.concat([quit_.assign(cambio="Ya no está / modificado"), nuevos.assign(cambio="Nuevo / modificado")])
                             .drop(columns="_merge"), hide_index=True, use_container_width=True)
            nuevas = sorted(set(f.index) - set(prev.cuentas.index))
            dif_ini = (f.ini - prev.cuentas.ini.reindex(f.index)).dropna()
            dif_ini = dif_ini[dif_ini.abs() > 0.005]
            st.markdown(f"**Cuentas nuevas:** {len(nuevas)}")
            if nuevas:
                st.dataframe(_sty(f.loc[nuevas, ["nombre", "ini", "fin", "n"]].reset_index(), ["ini", "fin"]), hide_index=True, use_container_width=True)
            if len(dif_ini):
                st.warning(f"{len(dif_ini)} cuentas cambiaron de saldo inicial.")
            prev_deu = set(prev.cuentas[prev.cuentas.fin < -0.005].index)
            act_deu = set(f[f.fin < -0.005].index)
            if prev_deu - act_deu:
                st.markdown("**Dejaron de tener saldo deudor:** " + ", ".join(f.loc[list(prev_deu - act_deu), "nombre"]))
            if act_deu - prev_deu:
                st.markdown("**Nuevos saldos deudores:** " + ", ".join(f.loc[list(act_deu - prev_deu), "nombre"]))

# ── Descargas ─────────────────────────────────────────────────────────────────
with tabs[6]:
    base = up.name.rsplit(".", 1)[0]
    c1, c2, c3 = st.columns(3)
    with c3:
        st.markdown("##### Resumen por proveedor")
        st.caption("Por cada proveedor: saldo inicial, cargos y abonos que se cancelan y los que quedan pendientes, "
                   "con el detalle de cada movimiento pendiente de conciliar.")
        with st.spinner("Generando…"):
            b3 = _excel_proveedor(data, key)
        st.download_button("⬇️ Descargar resumen por proveedor", b3, file_name=f"Resumen_por_proveedor_{base}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
    with c1:
        st.markdown("##### Auxiliar marcado")
        st.caption("Tu archivo original con cada movimiento en amarillo (relacionado) o azul (sin relación), "
                   "columna con el motivo y hoja 'Resumen relación' con fórmulas.")
        with st.spinner("Generando…"):
            b1 = _excel_marcado(data, key)
        st.download_button("⬇️ Descargar auxiliar marcado", b1, file_name=f"{base}_revisado.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
    with c2:
        st.markdown("##### Papel de trabajo de auditoría")
        st.caption("Resumen con cuadre, hallazgos con seguimiento (responsable, fecha, estatus) y una hoja de detalle por hallazgo.")
        with st.spinner("Generando…"):
            b2 = _excel_papel(data, key)
        st.download_button("⬇️ Descargar papel de trabajo", b2, file_name=f"Depuracion_{base}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
