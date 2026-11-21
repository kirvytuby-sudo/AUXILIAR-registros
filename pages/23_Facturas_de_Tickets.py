"""
pages/23_Facturas_de_Tickets.py — FACTURAS DE TICKETS (v1, 2026-11-20)

Sube fotos de tickets (gasolina, casetas, restaurantes, tiendas…) y el módulo:
  · Lee cada ticket con la API de Claude (comercio, fecha, total, forma de pago y los
    datos que pide el portal: ticket, folio, Web ID, estación, tienda…)
  · Liga al portal de facturación y calcula la fecha límite
  · Avisa lo que no es deducible (gasolina en efectivo, propinas, recargas TeleVía…)
  · Arma el correo de solicitud para los comercios sin portal
  · Lleva el control de estatus y lo descarga en Excel
Motor: _app/tickets_factura.py
"""
import hashlib
import os
from datetime import date

import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Facturas de Tickets | Auxiliar de Registros",
    page_icon="🧾",
    layout="wide",
)

import _theme
from _app import tickets_factura as tf

_theme.aplicar_header("🧾 FACTURAS DE TICKETS",
                      "Fotos de tickets → datos para facturar, portal, fecha límite, avisos y control")

ss = st.session_state
ss.setdefault("tk_receptor", tf.Receptor())
ss.setdefault("tk_tickets", [])          # list[dict]
ss.setdefault("tk_leidos", set())        # huellas de fotos ya procesadas
ss.setdefault("tk_fotos", {})            # nombre → bytes


def _api_key() -> str:
    if ss.get("tk_key"):
        return ss["tk_key"]
    try:
        k = st.secrets.get("ANTHROPIC_API_KEY", "")
        if k:
            return k
    except Exception:
        pass
    return os.environ.get("ANTHROPIC_API_KEY", "")


def _md(t) -> str:
    return str(t).replace("$", "\\$")


tab_rec, tab_sub, tab_tk, tab_ctl = st.tabs(["1 · Receptor", "2 · Subir tickets", "3 · Facturar", "4 · Control"])

# ── 1. Receptor ───────────────────────────────────────────────────────────────
with tab_rec:
    r: tf.Receptor = ss.tk_receptor
    st.markdown("Datos fiscales de quien recibe la factura. Puedes leerlos de la **Constancia de Situación Fiscal**.")
    csf = st.file_uploader("Constancia de Situación Fiscal (PDF)", type=["pdf"], key="tk_csf")
    if csf is not None and ss.get("tk_csf_hash") != hashlib.md5(csf.getvalue()).hexdigest():
        try:
            leido = tf.receptor_de_constancia(csf.getvalue())
            leido.uso_cfdi, leido.correo = r.uso_cfdi, r.correo
            ss.tk_receptor = r = leido
            ss.tk_csf_hash = hashlib.md5(csf.getvalue()).hexdigest()
            st.success(f"Leí la constancia de **{r.nombre}** ({r.rfc}).")
        except Exception as e:
            st.error(f"No pude leer la constancia: {e}")

    c1, c2 = st.columns(2)
    r.rfc = c1.text_input("RFC", r.rfc).strip().upper()
    r.nombre = c2.text_input("Nombre o razón social (como en la constancia, sin régimen societario)", r.nombre).strip().upper()
    c3, c4, c5 = st.columns([2, 1, 2])
    regs = list(tf.REGIMENES)
    r.regimen = c3.selectbox("Régimen fiscal", regs, index=regs.index(r.regimen) if r.regimen in regs else 0,
                             format_func=lambda k: f"{k} – {tf.REGIMENES[k]}")
    r.cp = c4.text_input("Código postal fiscal", r.cp).strip()
    usos = list(tf.USOS_CFDI)
    r.uso_cfdi = c5.selectbox("Uso de CFDI", usos, index=usos.index(r.uso_cfdi) if r.uso_cfdi in usos else 0,
                              format_func=lambda k: f"{k} – {tf.USOS_CFDI[k]}")
    r.correo = st.text_input("Correo donde llegan el XML y el PDF", r.correo).strip()
    if r.completo():
        st.success("Receptor completo.")
    else:
        st.info("Faltan datos del receptor.")

# ── 2. Subir tickets ──────────────────────────────────────────────────────────
with tab_sub:
    if not _api_key():
        st.warning("Falta la API key de Anthropic. En Streamlit Cloud agrégala en Secrets como "
                   "`ANTHROPIC_API_KEY`, o pégala aquí para esta sesión.")
        k = st.text_input("API key de Anthropic", type="password")
        if k:
            ss.tk_key = k.strip()
    fotos = st.file_uploader("Fotos de los tickets (puedes subir varias; si viene el voucher junto, mejor)",
                             type=["jpg", "jpeg", "png", "webp"], accept_multiple_files=True, key="tk_up")
    nuevas = [f for f in (fotos or []) if hashlib.md5(f.getvalue()).hexdigest() not in ss.tk_leidos]
    if fotos:
        st.caption(f"{len(fotos)} foto(s) · {len(nuevas)} sin leer")
    if st.button(f"📷 Leer {len(nuevas)} ticket(s)", type="primary", disabled=not nuevas or not _api_key()):
        barra = st.progress(0.0)
        for i, f in enumerate(nuevas, 1):
            data = f.getvalue()
            try:
                leidos = tf.extraer_tickets(f.name, data, _api_key())
                ss.tk_tickets.extend(tf.a_dict(t) for t in leidos)
                ss.tk_fotos[f.name] = data
                ss.tk_leidos.add(hashlib.md5(data).hexdigest())
                st.write(f"✅ {f.name}: " + ", ".join(_md(f"{t.comercio} ${t.total:,.2f}") for t in leidos))
            except Exception as e:
                st.error(f"❌ {f.name}: {e}")
            barra.progress(i / len(nuevas))
        st.rerun()

# ── 3. Facturar ───────────────────────────────────────────────────────────────
with tab_tk:
    r = ss.tk_receptor
    tks = [tf.completar(tf.de_dict(d)) for d in ss.tk_tickets]
    if not tks:
        st.info("Sube fotos en la pestaña **2 · Subir tickets**.")
    if tks and not r.completo():
        st.warning("Completa los datos del receptor en la pestaña 1 para armar los correos.")
    hoy = date.today()
    orden = sorted(range(len(tks)), key=lambda i: (tks[i].estatus not in ("Pendiente", "Vencida"), tks[i].fecha_limite or "9"))
    for i in orden:
        t = tks[i]
        lim = t.fecha_limite
        dias = (date.fromisoformat(lim) - hoy).days if lim else None
        semaforo = "🔴" if t.estatus == "Vencida" else "🟢" if t.estatus == "Recibida" else \
                   "🔵" if t.estatus == "Solicitada" else "⚪" if t.estatus == "No facturable" else \
                   "🟠" if dias is not None and dias <= 3 else "🟡"
        titulo = f"{semaforo} {t.comercio or 'Ticket'} · {t.fecha} · ${t.total:,.2f} · {t.estatus}"
        with st.expander(_md(titulo), expanded=t.estatus in ("Pendiente", "Vencida")):
            a, b = st.columns([3, 2])
            with a:
                st.markdown(f"**{_md(t.razon_social or t.comercio)}**" + (f" · RFC {t.rfc_emisor}" if t.rfc_emisor else ""))
                st.markdown(_md(f"Categoría: {t.categoria} · Forma de pago: {t.forma_pago or '—'}"
                                + (f" · Propina: ${t.propina:,.2f}" if t.propina else "")))
                if lim:
                    st.markdown(f"**Fecha límite:** {lim}" + (f" ({dias} días)" if dias is not None and dias >= 0 else " — vencida"))
                st.markdown("**Datos para el portal**")
                datos = "\n".join(f"{k}: {v}" for k, v in t.datos_facturar.items())
                datos += f"\nFecha: {t.fecha}\nImporte: {t.total:.2f}"
                st.code(datos + "\n\n" + "\n".join(r.lineas()) + f"\nCorreo: {r.correo}", language=None)
                for av in t.avisos:
                    st.warning(_md(av))
            with b:
                if t.archivo in ss.tk_fotos:
                    st.image(ss.tk_fotos[t.archivo], use_container_width=True)
            c1, c2, c3, c4 = st.columns(4)
            if t.portal:
                c1.link_button("🌐 Abrir portal", t.portal, use_container_width=True)
                if t.nota_portal:
                    st.caption(_md(t.nota_portal))
            if t.correo_facturacion and r.completo():
                c2.link_button("✉️ Correo de solicitud", tf.mailto(t, r), use_container_width=True)
            nuevo = c3.selectbox("Estatus", tf.ESTATUS, index=tf.ESTATUS.index(t.estatus),
                                 key=f"tk_est_{i}", label_visibility="collapsed")
            if nuevo != ss.tk_tickets[i].get("estatus"):
                ss.tk_tickets[i]["estatus"] = nuevo
            if c4.button("🗑️ Quitar", key=f"tk_del_{i}", use_container_width=True):
                ss.tk_tickets.pop(i)
                st.rerun()
            with st.popover("✏️ Corregir datos"):
                d = ss.tk_tickets[i]
                d["comercio"] = st.text_input("Comercio", d["comercio"], key=f"tk_c_{i}")
                d["fecha"] = st.text_input("Fecha (AAAA-MM-DD)", d["fecha"], key=f"tk_f_{i}")
                d["total"] = st.number_input("Total", value=float(d["total"]), step=0.01, key=f"tk_t_{i}")
                cats = tf.CATEGORIAS
                d["categoria"] = st.selectbox("Categoría", cats, index=cats.index(d["categoria"]) if d["categoria"] in cats else len(cats) - 1, key=f"tk_k_{i}")
                d["forma_pago"] = st.text_input("Forma de pago", d["forma_pago"], key=f"tk_p_{i}")
                d["portal"] = st.text_input("Portal", d["portal"], key=f"tk_u_{i}")
                d["correo_facturacion"] = st.text_input("Correo de facturación", d["correo_facturacion"], key=f"tk_m_{i}")
                txt = st.text_area("Datos para facturar (uno por renglón, Etiqueta: valor)",
                                   "\n".join(f"{k}: {v}" for k, v in d["datos_facturar"].items()), key=f"tk_d_{i}")
                d["datos_facturar"] = {k.strip(): v.strip() for k, _, v in
                                       (x.partition(":") for x in txt.splitlines()) if k.strip() and v.strip()}

# ── 4. Control ────────────────────────────────────────────────────────────────
with tab_ctl:
    r = ss.tk_receptor
    tks = [tf.completar(tf.de_dict(d)) for d in ss.tk_tickets]
    if tks:
        df = pd.DataFrame([{
            "Fecha": t.fecha, "Comercio": t.comercio, "Categoría": t.categoria, "Total": t.total,
            "Forma de pago": t.forma_pago, "Fecha límite": t.fecha_limite, "Estatus": t.estatus,
            "Avisos": len(t.avisos),
        } for t in tks]).sort_values("Fecha")
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Tickets", len(df))
        m2.metric("Importe", f"${df.Total.sum():,.2f}")
        m3.metric("Pendientes", int(df.Estatus.isin(["Pendiente"]).sum()))
        m4.metric("Vencidos", int((df.Estatus == "Vencida").sum()))
        st.dataframe(df, hide_index=True, use_container_width=True,
                     column_config={"Total": st.column_config.NumberColumn(format="$%.2f")})
        st.download_button("⬇️ Control en Excel", tf.control_excel(tks, r),
                           file_name=f"control_facturas_tickets_{date.today():%Y%m%d}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    else:
        st.info("Todavía no hay tickets leídos.")
    st.caption("La factura la timbra el comercio. Este módulo prepara la solicitud; los portales con "
               "captcha o con cuenta de usuario se terminan a mano.")
