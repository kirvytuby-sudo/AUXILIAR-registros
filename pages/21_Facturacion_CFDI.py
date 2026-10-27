"""
pages/21_Facturacion_CFDI.py — Facturación CFDI 4.0 (emitir y timbrar)

Emite facturas de ingreso y las TIMBRA ante el SAT a través del PAC Facturapi.
  · Nueva factura: cliente (buscar o capturar), conceptos, impuestos → timbrar
  · Facturas emitidas: buscar, descargar PDF/XML, reenviar por correo, cancelar
  · Configuración: estado de la llave (pruebas / producción) y cómo darla de alta

Requisitos (una sola vez):
  1. Cuenta en https://dashboard.facturapi.io → crear la organización (empresa)
  2. Subir ahí los CSD (.cer, .key y contraseña) y datos fiscales del emisor
  3. Copiar la llave secreta a los Secrets de Streamlit (ver _facturapi.py)
Con una llave sk_test_ las facturas NO tienen validez fiscal (modo pruebas).
"""
import re
from datetime import date, timedelta

import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Facturación CFDI | Auxiliar de Registros",
    page_icon="🧾",
    layout="wide",
)

import _theme
import _auth as _auth_mod
import _facturapi as fx

_theme.aplicar_header("🧾 Facturación CFDI 4.0",
                      "Emite y timbra facturas ante el SAT (PAC: Facturapi)")

_auth_user, _auth_name, _es_admin = _auth_mod.require_login("Facturación")

RFC_RE = re.compile(r"^[A-ZÑ&]{3,4}\d{6}[A-Z0-9]{3}$")
RFC_PUBLICO = "XAXX010101000"


# ── Empresas / llaves ─────────────────────────────────────────────────────────
def _empresas() -> dict[str, dict]:
    """{alias: {"nombre":..., "secret_key":...}} desde Secrets."""
    try:
        cfg = st.secrets.get("facturapi")
    except Exception:
        cfg = None
    if not cfg:
        return {}
    out: dict[str, dict] = {}
    emps = cfg.get("empresas") if hasattr(cfg, "get") else None
    if emps:
        for alias, d in dict(emps).items():
            d = dict(d)
            if d.get("secret_key"):
                out[alias] = {"nombre": d.get("nombre", alias.upper()),
                              "secret_key": d["secret_key"]}
    elif cfg.get("secret_key"):
        out["principal"] = {"nombre": cfg.get("nombre", "Empresa"),
                            "secret_key": cfg["secret_key"]}
    return out


EMPRESAS = _empresas()

with st.sidebar:
    st.markdown("---")
    st.markdown("**🧾 Empresa emisora**")
    if EMPRESAS:
        _alias = st.selectbox("Emisor", list(EMPRESAS),
                              format_func=lambda a: EMPRESAS[a]["nombre"],
                              label_visibility="collapsed", key="fx_emisor")
    else:
        _alias = None

if not EMPRESAS:
    st.warning("Falta configurar la llave de Facturapi en los **Secrets** de la app.")
    st.markdown("""
**Pasos (una sola vez):**
1. Crea tu cuenta en **[dashboard.facturapi.io](https://dashboard.facturapi.io)** y da de alta la empresa.
2. En *Configuración → Certificados* sube tu **CSD** (.cer, .key y contraseña) y llena los datos fiscales del emisor.
3. En *Desarrolladores → Llaves* copia la **llave secreta** (`sk_test_…` para pruebas, `sk_live_…` para producción).
4. En Streamlit Cloud → *Settings → Secrets* pega:
""")
    st.code('[facturapi]\nsecret_key = "sk_test_xxxxxxxxxxxxxxxx"\n\n'
            '# o, para varias empresas:\n'
            '# [facturapi.empresas.sanitaria]\n'
            '# nombre = "La Sanitaria"\n'
            '# secret_key = "sk_live_xxxxxxxx"', language="toml")
    st.stop()

api = fx.Facturapi(EMPRESAS[_alias]["secret_key"])

if api.modo_pruebas:
    st.info("🧪 **Modo pruebas** — las facturas se timbran en el ambiente de pruebas "
            "y **no tienen validez fiscal**. Cambia a una llave `sk_live_` para facturar en real.")
else:
    st.success(f"🟢 **Producción** — emisor: **{EMPRESAS[_alias]['nombre']}**. "
               "Las facturas se timbran ante el SAT con validez fiscal.")


def _money(x) -> str:
    try:
        return f"${float(x):,.2f}"
    except (TypeError, ValueError):
        return "—"


tab_nueva, tab_emitidas = st.tabs(["➕ Nueva factura", "📂 Facturas emitidas"])

# ══════════════════════════════════════════════════════════════════════════════
#  NUEVA FACTURA
# ══════════════════════════════════════════════════════════════════════════════
with tab_nueva:
    # ── 1. Cliente ────────────────────────────────────────────────────────
    st.markdown("#### 1 · Cliente (receptor)")
    modo_cli = st.radio("Cliente", ["Buscar cliente guardado", "Capturar cliente nuevo",
                                    "Público en general"],
                        horizontal=True, label_visibility="collapsed", key="fx_modo_cli")

    cliente_payload = None
    cliente_email = ""

    if modo_cli == "Buscar cliente guardado":
        q = st.text_input("Buscar por nombre o RFC", key="fx_q_cli")
        try:
            clientes = api.buscar_clientes(q)
        except fx.FacturapiError as e:
            clientes = []
            st.error(f"No se pudieron cargar clientes: {e}")
        if clientes:
            sel = st.selectbox(
                "Cliente", range(len(clientes)),
                format_func=lambda i: f"{clientes[i].get('legal_name')} — {clientes[i].get('tax_id')}",
                key="fx_sel_cli")
            c = clientes[sel]
            cliente_payload = c["id"]
            cliente_email = c.get("email") or ""
            st.caption(f"Régimen {c.get('tax_system', '—')} · C.P. "
                       f"{(c.get('address') or {}).get('zip', '—')} · {cliente_email or 'sin correo'}")
        else:
            st.caption("Sin resultados. Usa *Capturar cliente nuevo*.")

    elif modo_cli == "Capturar cliente nuevo":
        c1, c2 = st.columns([2, 1])
        nombre = c1.text_input("Razón social (como en su Constancia, sin régimen societario: "
                               "sin “S.A. de C.V.”)", key="fx_cli_nombre").strip().upper()
        rfc = c2.text_input("RFC", key="fx_cli_rfc").strip().upper()
        c3, c4, c5 = st.columns([2, 1, 2])
        regimen = c3.selectbox("Régimen fiscal", list(fx.REGIMENES),
                               format_func=lambda k: f"{k} · {fx.REGIMENES[k]}", key="fx_cli_reg")
        cp = c4.text_input("C.P. fiscal", max_chars=5, key="fx_cli_cp").strip()
        cliente_email = c5.text_input("Correo (opcional)", key="fx_cli_mail").strip()
        guardar_cli = st.checkbox("Guardar cliente para la próxima vez", value=True, key="fx_cli_save")

        errores_cli = []
        if rfc and not RFC_RE.match(rfc):
            errores_cli.append("RFC con formato inválido.")
        if cp and not re.fullmatch(r"\d{5}", cp):
            errores_cli.append("El C.P. debe tener 5 dígitos.")
        for e in errores_cli:
            st.error(e)
        if nombre and rfc and cp and not errores_cli:
            cliente_payload = {"legal_name": nombre, "tax_id": rfc, "tax_system": regimen,
                               "address": {"zip": cp}}
            if cliente_email:
                cliente_payload["email"] = cliente_email

    else:  # Público en general
        st.caption("RFC XAXX010101000 · Régimen 616 · Uso S01. "
                   "El C.P. debe ser el del lugar de expedición (tu C.P. fiscal).")
        cp_pg = st.text_input("C.P. de expedición", max_chars=5, key="fx_pg_cp").strip()
        if re.fullmatch(r"\d{5}", cp_pg or ""):
            cliente_payload = {"legal_name": "PUBLICO EN GENERAL", "tax_id": RFC_PUBLICO,
                               "tax_system": "616", "address": {"zip": cp_pg}}

    # ── 2. Datos del comprobante ──────────────────────────────────────────
    st.markdown("#### 2 · Datos del comprobante")
    d1, d2, d3, d4 = st.columns(4)
    es_pg = modo_cli == "Público en general"
    uso = d1.selectbox("Uso CFDI", list(fx.USOS_CFDI),
                       index=list(fx.USOS_CFDI).index("S01" if es_pg else "G03"),
                       format_func=lambda k: f"{k} · {fx.USOS_CFDI[k]}", key="fx_uso")
    metodo = d2.selectbox("Método de pago", list(fx.METODOS_PAGO),
                          format_func=lambda k: f"{k} · {fx.METODOS_PAGO[k]}", key="fx_metodo")
    formas = list(fx.FORMAS_PAGO)
    forma = d3.selectbox("Forma de pago", formas,
                         index=formas.index("99" if metodo == "PPD" else "03"),
                         format_func=lambda k: f"{k} · {fx.FORMAS_PAGO[k]}", key="fx_forma",
                         disabled=(metodo == "PPD"))
    if metodo == "PPD":
        forma = "99"
    serie = d4.text_input("Serie (opcional)", max_chars=25, key="fx_serie").strip().upper()

    # ── 3. Conceptos ──────────────────────────────────────────────────────
    st.markdown("#### 3 · Conceptos")
    with st.expander("🔎 Buscar clave de producto/servicio o unidad del SAT"):
        b1, b2 = st.columns(2)
        qp = b1.text_input("Producto o servicio (ej. «gasolina», «consultoría»)", key="fx_qprod")
        if qp:
            try:
                res = api.buscar_producto_sat(qp)
                b1.dataframe(pd.DataFrame([{"Clave": r.get("key"), "Descripción": r.get("description")}
                                           for r in res]), hide_index=True, use_container_width=True)
            except fx.FacturapiError as e:
                b1.error(str(e))
        qu = b2.text_input("Unidad (ej. «litro», «pieza»)", key="fx_qunit")
        if qu:
            try:
                res = api.buscar_unidad_sat(qu)
                b2.dataframe(pd.DataFrame([{"Clave": r.get("key"), "Descripción": r.get("description")}
                                           for r in res]), hide_index=True, use_container_width=True)
            except fx.FacturapiError as e:
                b2.error(str(e))

    if "fx_conceptos" not in st.session_state:
        st.session_state["fx_conceptos"] = pd.DataFrame([{
            "Cantidad": 1.0, "Descripción": "", "Clave SAT": "", "Unidad": "E48",
            "Precio unitario": 0.0, "Impuesto": "IVA 16%",
        }])

    conceptos = st.data_editor(
        st.session_state["fx_conceptos"],
        num_rows="dynamic", use_container_width=True, key="fx_editor",
        column_config={
            "Cantidad": st.column_config.NumberColumn(min_value=0.000001, format="%.2f", required=True),
            "Descripción": st.column_config.TextColumn(width="large", required=True),
            "Clave SAT": st.column_config.TextColumn(help="Clave ProdServ de 8 dígitos", max_chars=8),
            "Unidad": st.column_config.SelectboxColumn(
                options=list(fx.UNIDADES_COMUNES), required=True,
                help=", ".join(f"{k}={v}" for k, v in fx.UNIDADES_COMUNES.items())),
            "Precio unitario": st.column_config.NumberColumn(min_value=0.0, format="$%.2f", required=True),
            "Impuesto": st.column_config.SelectboxColumn(options=list(fx.IVA_OPCIONES), required=True),
        },
    )

    o1, o2, o3 = st.columns(3)
    iva_incluido = o1.checkbox("El precio YA incluye IVA", value=False, key="fx_incl")
    ret_iva = o2.checkbox("Retener IVA (2/3)", key="fx_retiva",
                          help="Servicios profesionales / arrendamiento de PF a PM")
    ret_isr = o3.checkbox("Retener ISR 10%", key="fx_retisr")

    # Validar conceptos y calcular totales estimados
    filas = conceptos.dropna(how="all")
    errores, items = [], []
    subtotal = iva_tot = ret_tot = 0.0
    for i, r in filas.iterrows():
        desc = str(r.get("Descripción") or "").strip()
        clave = str(r.get("Clave SAT") or "").strip()
        cant = float(r.get("Cantidad") or 0)
        precio = float(r.get("Precio unitario") or 0)
        imp = r.get("Impuesto") or "IVA 16%"
        if not desc and not clave and precio == 0:
            continue
        n = len(items) + 1
        if not desc:
            errores.append(f"Concepto {n}: falta descripción.")
        if not re.fullmatch(r"\d{8}", clave):
            errores.append(f"Concepto {n}: la clave SAT debe tener 8 dígitos.")
        if cant <= 0 or precio <= 0:
            errores.append(f"Concepto {n}: cantidad y precio deben ser mayores a 0.")
        tasa = fx.IVA_OPCIONES.get(imp) or 0.0
        base = cant * precio / (1 + tasa) if iva_incluido else cant * precio
        subtotal += base
        iva_tot += base * tasa
        ret_tot += base * ((0.106667 if ret_iva else 0) + (0.10 if ret_isr else 0))
        items.append({
            "quantity": cant,
            "product": {
                "description": desc,
                "product_key": clave,
                "unit_key": r.get("Unidad") or "E48",
                "unit_name": fx.UNIDADES_COMUNES.get(r.get("Unidad"), ""),
                "price": round(precio, 6),
                "tax_included": bool(iva_incluido),
                "taxes": fx.impuestos_concepto(imp, ret_iva, ret_isr),
            },
        })

    t1, t2, t3, t4 = st.columns(4)
    t1.metric("Subtotal", _money(subtotal))
    t2.metric("IVA", _money(iva_tot))
    t3.metric("Retenciones", _money(-ret_tot) if ret_tot else "$0.00")
    t4.metric("Total (estimado)", _money(subtotal + iva_tot - ret_tot))
    st.caption("El total definitivo lo calcula el PAC al timbrar.")

    # ── 4. Timbrar ────────────────────────────────────────────────────────
    st.markdown("#### 4 · Timbrar")
    faltantes = []
    if not cliente_payload:
        faltantes.append("completa los datos del cliente")
    if not items:
        faltantes.append("agrega al menos un concepto")
    for e in errores:
        st.error(e)
    if faltantes:
        st.caption("Para timbrar: " + " y ".join(faltantes) + ".")

    enviar_mail = st.checkbox("Enviar PDF y XML al correo del cliente al timbrar",
                              value=bool(cliente_email), disabled=not cliente_email,
                              key="fx_mail_auto")
    confirmar = st.checkbox("Revisé los datos: timbrar esta factura ante el SAT", key="fx_confirm")
    puede = bool(cliente_payload and items and not errores and confirmar)

    if st.button("🧾 Timbrar factura", type="primary", disabled=not puede, key="fx_timbrar"):
        try:
            with st.spinner("Timbrando ante el SAT…"):
                if (modo_cli == "Capturar cliente nuevo" and st.session_state.get("fx_cli_save")
                        and isinstance(cliente_payload, dict)):
                    try:
                        cliente_payload = api.crear_cliente(cliente_payload)["id"]
                    except fx.FacturapiError as e:
                        st.warning(f"No se guardó el cliente ({e}); se factura con los datos capturados.")
                datos = {
                    "type": "I",
                    "customer": cliente_payload,
                    "items": items,
                    "use": uso,
                    "payment_form": forma,
                    "payment_method": metodo,
                }
                if serie:
                    datos["series"] = serie
                if es_pg:
                    hoy = date.today()
                    datos["global"] = {"periodicity": "day", "months": f"{hoy.month:02d}",
                                       "year": hoy.year}
                fac = api.crear_factura(datos)
            st.session_state["fx_ultima"] = fac
            if enviar_mail and cliente_email:
                try:
                    api.enviar_correo(fac["id"], cliente_email)
                    st.session_state["fx_ultima_mail"] = cliente_email
                except fx.FacturapiError as e:
                    st.warning(f"Timbrada, pero no se pudo enviar el correo: {e}")
            for _k in ("fx_conceptos", "fx_editor", "fx_confirm"):
                st.session_state.pop(_k, None)
            st.rerun()
        except fx.FacturapiError as e:
            st.error(f"❌ El PAC rechazó la factura: {e}")

    fac = st.session_state.get("fx_ultima")
    if fac:
        st.success(f"✅ Factura timbrada · Folio **{fac.get('series') or ''}{fac.get('folio_number', '')}** · "
                   f"UUID `{fac.get('uuid')}` · Total **{_money(fac.get('total'))}**")
        if st.session_state.get("fx_ultima_mail"):
            st.caption(f"📧 Enviada a {st.session_state['fx_ultima_mail']}")
        k1, k2, k3 = st.columns(3)
        try:
            k1.download_button("⬇️ PDF", api.descargar(fac["id"], "pdf"),
                               file_name=f"{fac.get('uuid')}.pdf", mime="application/pdf",
                               use_container_width=True)
            k2.download_button("⬇️ XML", api.descargar(fac["id"], "xml"),
                               file_name=f"{fac.get('uuid')}.xml", mime="application/xml",
                               use_container_width=True)
        except fx.FacturapiError as e:
            st.error(f"No se pudieron descargar los archivos: {e}")
        if k3.button("Nueva factura", use_container_width=True, key="fx_nueva"):
            st.session_state.pop("fx_ultima", None)
            st.session_state.pop("fx_ultima_mail", None)
            st.rerun()

# ══════════════════════════════════════════════════════════════════════════════
#  FACTURAS EMITIDAS
# ══════════════════════════════════════════════════════════════════════════════
with tab_emitidas:
    f1, f2, f3 = st.columns([2, 1, 1])
    q_fac = f1.text_input("Buscar (cliente, RFC, folio o UUID)", key="fx_q_fac")
    desde = f2.date_input("Desde", date.today() - timedelta(days=30), key="fx_desde")
    hasta = f3.date_input("Hasta", date.today(), key="fx_hasta")

    try:
        res = api.listar_facturas(q_fac, limit=100,
                                  desde=f"{desde.isoformat()}T00:00:00",
                                  hasta=f"{hasta.isoformat()}T23:59:59")
        facturas = res.get("data", [])
    except fx.FacturapiError as e:
        facturas = []
        st.error(f"No se pudieron consultar las facturas: {e}")

    if not facturas:
        st.caption("No hay facturas en ese periodo.")
    else:
        df = pd.DataFrame([{
            "Fecha": (f.get("date") or f.get("created_at") or "")[:10],
            "Folio": f"{f.get('series') or ''}{f.get('folio_number') or ''}",
            "Cliente": (f.get("customer") or {}).get("legal_name", ""),
            "RFC": (f.get("customer") or {}).get("tax_id", ""),
            "Total": f.get("total", 0),
            "Método": f.get("payment_method", ""),
            "Estatus": {"valid": "Vigente", "canceled": "Cancelada"}.get(f.get("status"), f.get("status")),
            "UUID": f.get("uuid", ""),
        } for f in facturas])
        st.dataframe(df, hide_index=True, use_container_width=True,
                     column_config={"Total": st.column_config.NumberColumn(format="$%.2f")})
        vig = df[df["Estatus"] == "Vigente"]["Total"].sum()
        st.caption(f"{len(df)} factura(s) · Vigentes: {_money(vig)}")

        st.markdown("##### Acciones")
        idx = st.selectbox(
            "Factura", range(len(facturas)),
            format_func=lambda i: f"{df.iloc[i]['Folio'] or '(s/folio)'} · {df.iloc[i]['Cliente']} · "
                                  f"{_money(df.iloc[i]['Total'])} · {df.iloc[i]['Estatus']}",
            key="fx_sel_fac")
        f = facturas[idx]
        a1, a2, a3 = st.columns(3)
        with a1:
            if st.button("Preparar descarga PDF + XML", use_container_width=True, key="fx_prep_dl"):
                try:
                    st.session_state["fx_zip"] = (f["id"], api.descargar(f["id"], "zip"))
                except fx.FacturapiError as e:
                    st.error(str(e))
            z = st.session_state.get("fx_zip")
            if z and z[0] == f["id"]:
                st.download_button("⬇️ Descargar ZIP", z[1], file_name=f"{f.get('uuid')}.zip",
                                   mime="application/zip", use_container_width=True)
        with a2:
            mail = st.text_input("Correo", value=(f.get("customer") or {}).get("email") or "",
                                 key=f"fx_mail_{f['id']}", label_visibility="collapsed",
                                 placeholder="correo@cliente.com")
            if st.button("📧 Enviar por correo", use_container_width=True, key="fx_send",
                         disabled=not mail):
                try:
                    api.enviar_correo(f["id"], mail)
                    st.success("Enviada.")
                except fx.FacturapiError as e:
                    st.error(str(e))
        with a3:
            if f.get("status") == "valid":
                with st.popover("🚫 Cancelar ante el SAT", use_container_width=True):
                    motivo = st.selectbox("Motivo", list(fx.MOTIVOS_CANCELACION),
                                          format_func=lambda k: f"{k} · {fx.MOTIVOS_CANCELACION[k]}",
                                          key="fx_motivo")
                    sust = ""
                    if motivo == "01":
                        sust = st.text_input("ID de Facturapi de la factura que la sustituye",
                                             key="fx_sust").strip()
                    ok = st.checkbox("Confirmo la cancelación", key="fx_conf_cancel")
                    if st.button("Cancelar factura", type="primary", key="fx_cancel",
                                 disabled=not ok or (motivo == "01" and not sust)):
                        try:
                            r = api.cancelar(f["id"], motivo, sust or None)
                            est = r.get("cancellation_status") or r.get("status")
                            if est == "pending":
                                st.info("Solicitud enviada: el receptor debe aceptar la cancelación "
                                        "en el buzón del SAT.")
                            else:
                                st.success("Factura cancelada.")
                        except fx.FacturapiError as e:
                            st.error(str(e))
            else:
                st.caption(f"Estatus: {df.iloc[idx]['Estatus']}"
                           + (f" · cancelación {f.get('cancellation_status')}"
                              if f.get("cancellation_status") not in (None, "none") else ""))
