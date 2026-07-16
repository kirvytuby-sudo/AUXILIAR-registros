"""
pages/13_Descarga_CFDI_SAT.py
Descarga masiva de CFDIs emitidos y recibidos desde el SAT.
Autenticación: e.firma (.cer + .key + contraseña).
"""

import streamlit as st
import base64, zipfile, io, os, time
from datetime import date, datetime, timedelta
from lxml import etree
import pandas as pd

# ── Page config ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Descarga CFDI SAT",
    page_icon="📥",
    layout="wide",
)

# ── Tema (azul royal + rosa, igual que el resto de la app) ───────────────────
st.markdown("""
<style>
:root {
    --az: #1E3A8A; --az2: #2563EB; --az3: #DBEAFE;
    --rosa: #FFF0F5; --txt: #1e293b;
}
[data-testid="stAppViewContainer"] { background: var(--rosa); }
[data-testid="stHeader"] { background: var(--az); }
h1,h2,h3 { color: var(--az); }
.stButton>button { background: var(--az); color: #fff; border-radius: 6px; border: none; }
.stButton>button:hover { background: var(--az2); }
.sec-header {
    background: var(--az); color: #fff; padding: 8px 14px;
    border-radius: 8px 8px 0 0; font-weight: 700; font-size: 1rem;
    margin-bottom: 0;
}
.sec-body {
    background: #fff; border: 1.5px solid var(--az3);
    border-top: none; border-radius: 0 0 8px 8px;
    padding: 16px; margin-bottom: 18px;
}
.badge-vigente  { background:#d1fae5; color:#065f46; padding:2px 10px; border-radius:12px; font-size:.82rem; font-weight:700; }
.badge-cancelado{ background:#fee2e2; color:#991b1b; padding:2px 10px; border-radius:12px; font-size:.82rem; font-weight:700; }
.badge-nd       { background:#f3f4f6; color:#6b7280; padding:2px 10px; border-radius:12px; font-size:.82rem; font-weight:700; }
</style>
""", unsafe_allow_html=True)

# ── Constantes ───────────────────────────────────────────────────────────────
_TIPOS_COMP = {
    "Todos": None,
    "I — Ingreso": "I",
    "E — Egreso": "E",
    "P — Pago": "P",
    "N — Nómina": "N",
    "T — Traslado": "T",
}
_ESTADOS_COMP = {
    "Todos": "Todos",
    "Vigente": "Vigente",
    "Cancelado": "Cancelado",
}

NS4  = "http://www.sat.gob.mx/cfd/4"
NS3  = "http://www.sat.gob.mx/cfd/3"
NSTFD = "http://www.sat.gob.mx/TimbreFiscalDigital"

# ── Helpers de CFDI ──────────────────────────────────────────────────────────
def _parse_meta(xml_bytes: bytes) -> dict:
    """Extrae metadatos de un XML CFDI (v3.3 o v4.0)."""
    try:
        root = etree.fromstring(xml_bytes)
        tag  = root.tag
        ns   = NS4 if NS4 in tag else NS3

        emisor     = root.find(f"{{{ns}}}Emisor")
        receptor   = root.find(f"{{{ns}}}Receptor")
        complemento = root.find(f"{{{ns}}}Complemento")
        tfd = None
        if complemento is not None:
            tfd = complemento.find(f"{{{NSTFD}}}TimbreFiscalDigital")

        _tipo_map = {
            "I": "Ingreso","E": "Egreso","P": "Pago",
            "N": "Nómina","T": "Traslado",
        }
        tipo_raw = root.get("TipoDeComprobante", "")

        return {
            "uuid":            (tfd.get("UUID","") if tfd is not None else "").upper(),
            "fecha":           root.get("Fecha","")[:10],
            "tipo":            _tipo_map.get(tipo_raw, tipo_raw),
            "tipo_raw":        tipo_raw,
            "serie":           root.get("Serie",""),
            "folio":           root.get("Folio",""),
            "subtotal":        float(root.get("SubTotal","0") or "0"),
            "total":           float(root.get("Total","0")    or "0"),
            "moneda":          root.get("Moneda","MXN"),
            "emisor_rfc":      (emisor.get("Rfc","")    if emisor   is not None else ""),
            "emisor_nombre":   (emisor.get("Nombre","") if emisor   is not None else ""),
            "receptor_rfc":    (receptor.get("Rfc","")    if receptor is not None else ""),
            "receptor_nombre": (receptor.get("Nombre","") if receptor is not None else ""),
            "estado":          "—",
            "xml_bytes":       xml_bytes,
        }
    except Exception as e:
        return {"uuid":"ERROR","fecha":"","tipo":"","tipo_raw":"",
                "serie":"","folio":"","subtotal":0,"total":0,
                "moneda":"","emisor_rfc":"","emisor_nombre":"",
                "receptor_rfc":"","receptor_nombre":"",
                "estado":"—","xml_bytes":xml_bytes}


def _pdf_bytes_from_xml(xml_bytes: bytes) -> bytes | None:
    """Genera PDF desde bytes de XML CFDI. Retorna None si falla."""
    try:
        from satcfdi.cfdi import cfdi_objectify
        from satcfdi.render import pdf_bytes
        cfdi = cfdi_objectify(xml_bytes)
        return pdf_bytes(cfdi)
    except Exception as e:
        st.warning(f"No se pudo generar PDF: {e}")
        return None


def _html_from_xml(xml_bytes: bytes) -> str | None:
    """Genera HTML desde bytes de XML CFDI para preview."""
    try:
        from satcfdi.cfdi import cfdi_objectify
        from satcfdi.render import html_str
        cfdi = cfdi_objectify(xml_bytes)
        return html_str(cfdi)
    except Exception as e:
        return None


def _make_signer(cer_bytes: bytes, key_bytes: bytes, pwd: str):
    from satcfdi.models import Signer
    return Signer.load(cer_bytes, key_bytes, pwd)


def _make_sat(signer):
    from satcfdi.pacs.sat import SAT
    return SAT(signer)


def _unzip_xmls(zip_b64: str) -> list[tuple[str, bytes]]:
    """Desempaca paquete ZIP base64 del SAT. Retorna [(nombre, xml_bytes)]."""
    raw = base64.b64decode(zip_b64)
    result = []
    with zipfile.ZipFile(io.BytesIO(raw)) as zf:
        for name in zf.namelist():
            if name.lower().endswith(".xml"):
                result.append((name, zf.read(name)))
    return result


def _zip_files(files: list[tuple[str, bytes]]) -> bytes:
    """Empaca lista [(nombre, bytes)] en ZIP."""
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, data in files:
            zf.writestr(name, data)
    return buf.getvalue()


def _estado_badge(estado: str) -> str:
    e = (estado or "—").strip()
    if e == "Vigente":
        return '<span class="badge-vigente">✅ Vigente</span>'
    elif e == "Cancelado":
        return '<span class="badge-cancelado">❌ Cancelado</span>'
    return '<span class="badge-nd">— N/D</span>'


# ── SESSION STATE init ────────────────────────────────────────────────────────
for _k, _v in {
    "cfdi_signer":      None,
    "cfdi_sat":         None,
    "cfdi_rfc":         "",
    "cfdi_solicitudes": [],   # [{"tipo":"emitidos"|"recibidos","id":str,"paquetes":[],"estado":str}]
    "cfdi_resultados":  [],   # [meta_dict, ...]
    "cfdi_preview_idx": None,
    "cfdi_verificando": False,
    "cfdi_start_time":  None,
}.items():
    if _k not in st.session_state:
        st.session_state[_k] = _v

# ════════════════════════════════════════════════════════════════════════════
# SECCIÓN 1 — AUTENTICACIÓN
# ════════════════════════════════════════════════════════════════════════════
st.markdown('<div class="sec-header">🔐 1. Autenticación con e.firma</div>', unsafe_allow_html=True)
with st.container():
    st.markdown('<div class="sec-body">', unsafe_allow_html=True)

    _ya_auth = st.session_state["cfdi_signer"] is not None
    if _ya_auth:
        st.success(f"✅ Autenticado como **{st.session_state['cfdi_rfc']}**")
        if st.button("🔓 Cerrar sesión", key="btn_logout"):
            st.session_state["cfdi_signer"]      = None
            st.session_state["cfdi_sat"]         = None
            st.session_state["cfdi_rfc"]         = ""
            st.session_state["cfdi_solicitudes"] = []
            st.session_state["cfdi_resultados"]  = []
            st.rerun()
    else:
        _c1, _c2, _c3 = st.columns([1, 1, 1])
        with _c1:
            _cer_file = st.file_uploader("📄 Archivo .cer (certificado)", type=["cer"], key="fu_cer")
        with _c2:
            _key_file = st.file_uploader("🔑 Archivo .key (llave privada)", type=["key"], key="fu_key")
        with _c3:
            _pwd = st.text_input("🔒 Contraseña de la llave", type="password", key="ti_pwd")

        if st.button("🔐 Conectar", key="btn_connect", type="primary", use_container_width=True):
            if not _cer_file or not _key_file or not _pwd:
                st.error("Debes proporcionar el .cer, el .key y la contraseña.")
            else:
                try:
                    with st.spinner("Verificando e.firma…"):
                        _signer = _make_signer(
                            _cer_file.read(),
                            _key_file.read(),
                            _pwd,
                        )
                        st.session_state["cfdi_signer"] = _signer
                        st.session_state["cfdi_rfc"]    = _signer.rfc
                        st.session_state["cfdi_sat"]    = _make_sat(_signer)
                    st.success(f"✅ Conectado como {_signer.rfc}")
                    st.rerun()
                except Exception as _e:
                    st.error(f"Error al cargar e.firma: {_e}")

    st.markdown('</div>', unsafe_allow_html=True)

# Si no hay auth, no mostrar más
if not st.session_state["cfdi_signer"]:
    st.stop()

_sat = st.session_state["cfdi_sat"]
_rfc = st.session_state["cfdi_rfc"]

# ════════════════════════════════════════════════════════════════════════════
# SECCIÓN 2 — FILTROS DE BÚSQUEDA
# ════════════════════════════════════════════════════════════════════════════
st.markdown('<div class="sec-header">🔎 2. Filtros de búsqueda</div>', unsafe_allow_html=True)
st.markdown('<div class="sec-body">', unsafe_allow_html=True)

_modo = st.radio(
    "Modo de búsqueda",
    ["📅 Rango de fechas", "🆔 Por UUID"],
    horizontal=True,
    key="rad_modo",
)

if _modo == "🆔 Por UUID":
    _uuid_input = st.text_input(
        "UUID del CFDI",
        placeholder="xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx",
        key="ti_uuid",
    ).strip().upper()
    _tipo_dir = st.radio(
        "¿Es emitido o recibido?",
        ["Emitido", "Recibido"],
        horizontal=True,
        key="rad_dir_uuid",
    )
    _fecha_ini = _fecha_fin = None
    _tipo_comp_k = "Todos"
    _rfc_filtro  = ""
    _estado_k    = "Todos"
    _tipo_dl     = _tipo_dir
else:
    _uuid_input = ""
    _ca, _cb = st.columns(2)
    with _ca:
        _fecha_ini = st.date_input("📅 Fecha inicial", value=date.today().replace(day=1), key="di_ini")
    with _cb:
        _fecha_fin = st.date_input("📅 Fecha final",   value=date.today(),                key="di_fin")

    _c1, _c2, _c3, _c4 = st.columns(4)
    with _c1:
        _tipo_dl = st.selectbox(
            "Dirección",
            ["Ambos", "Emitidos", "Recibidos"],
            key="sel_dir",
        )
    with _c2:
        _tipo_comp_k = st.selectbox(
            "Tipo de comprobante",
            list(_TIPOS_COMP.keys()),
            key="sel_tipo",
        )
    with _c3:
        _rfc_filtro = st.text_input(
            "RFC emisor (opcional)",
            placeholder="RFC a filtrar",
            key="ti_rfc_filtro",
        ).strip().upper()
    with _c4:
        _estado_k = st.selectbox(
            "Estatus del comprobante",
            list(_ESTADOS_COMP.keys()),
            key="sel_estado",
        )

st.markdown('</div>', unsafe_allow_html=True)

# ════════════════════════════════════════════════════════════════════════════
# SECCIÓN 3 — SOLICITAR / VERIFICAR / DESCARGAR
# ════════════════════════════════════════════════════════════════════════════
st.markdown('<div class="sec-header">📡 3. Solicitar descarga al SAT</div>', unsafe_allow_html=True)
st.markdown('<div class="sec-body">', unsafe_allow_html=True)

from satcfdi.pacs.sat import EstadoSolicitud, EstadoComprobante, TipoDescargaMasivaTerceros

_btn_solicitar, _btn_verificar, _btn_descargar, _btn_limpiar = st.columns([2, 2, 2, 1])

with _btn_solicitar:
    if st.button("📤 Solicitar descarga", key="btn_solicitar", type="primary", use_container_width=True):
        _solicitudes_nuevas = []
        try:
            with st.spinner("Enviando solicitud al SAT…"):
                _tipo_comp_v  = _TIPOS_COMP.get(_tipo_comp_k)
                _estado_v     = _ESTADOS_COMP.get(_estado_k, "Todos")
                _ec = (
                    EstadoComprobante.VIGENTE   if _estado_v == "Vigente" else
                    EstadoComprobante.CANCELADO if _estado_v == "Cancelado" else
                    None
                )

                if _modo == "🆔 Por UUID":
                    _res = _sat.recover_comprobante_uuid_request(folio=_uuid_input)
                    if _res:
                        _solicitudes_nuevas.append({
                            "tipo":     _tipo_dir.lower(),
                            "id":       _res.get("IdSolicitud",""),
                            "paquetes": [],
                            "estado":   "ENVIADA",
                            "cod":      _res.get("CodEstatus",""),
                        })
                else:
                    if _tipo_dl in ("Ambos","Emitidos"):
                        _re = _sat.recover_comprobante_emitted_request(
                            fecha_inicial       = _fecha_ini,
                            fecha_final         = _fecha_fin,
                            tipo_comprobante    = _tipo_comp_v,
                            estado_comprobante  = _ec,
                            rfc_receptor        = _rfc_filtro or None,
                        )
                        _solicitudes_nuevas.append({
                            "tipo":     "emitidos",
                            "id":       _re.get("IdSolicitud",""),
                            "paquetes": [],
                            "estado":   "ENVIADA",
                            "cod":      _re.get("CodEstatus",""),
                        })
                    if _tipo_dl in ("Ambos","Recibidos"):
                        _rr = _sat.recover_comprobante_received_request(
                            fecha_inicial       = _fecha_ini,
                            fecha_final         = _fecha_fin,
                            tipo_comprobante    = _tipo_comp_v,
                            estado_comprobante  = _ec,
                            rfc_emisor          = _rfc_filtro or None,
                        )
                        _solicitudes_nuevas.append({
                            "tipo":     "recibidos",
                            "id":       _rr.get("IdSolicitud",""),
                            "paquetes": [],
                            "estado":   "ENVIADA",
                            "cod":      _rr.get("CodEstatus",""),
                        })

            st.session_state["cfdi_solicitudes"] = _solicitudes_nuevas
            st.session_state["cfdi_resultados"]  = []
            st.session_state["cfdi_start_time"]  = time.time()
            for _sol in _solicitudes_nuevas:
                _cod = _sol.get("cod","")
                if _cod in ("5000","5004"):
                    st.success(f"✅ Solicitud {_sol['tipo']} enviada — ID: `{_sol['id']}`")
                else:
                    st.warning(f"⚠️ Solicitud {_sol['tipo']} — Cód: {_cod} ID: `{_sol['id']}`")
        except Exception as _e:
            st.error(f"Error al solicitar: {_e}")

# ── BOTÓN 2: VERIFICAR ESTADO (sin descargar) ──────────────────────────────────────────
with _btn_verificar:
    if st.button("🔄 Verificar estado", key="btn_verificar", use_container_width=True):
        _sols = st.session_state.get("cfdi_solicitudes", [])
        if not _sols:
            st.warning("Primero haz clic en **Solicitar descarga**.")
        else:
            _t0       = time.time()  # timer inicia AQUI, no desde solicitar
            _timer_ph = st.empty()
            _stat_ph  = st.empty()
            _prog_ph  = st.empty()
            _debug_ph = st.empty()
            _MAX_ESP  = 300
            _POLL     = 8
            _terminados = set()
            _errores    = set()

            while True:
                _elapsed = time.time() - _t0
                _mm, _ss = divmod(int(_elapsed), 60)
                _rm, _rs = divmod(int(max(0, _MAX_ESP - _elapsed)), 60)
                _timer_ph.markdown(
                    "⏱️ **Verificando…** &nbsp;"
                    f"Transcurrido: `{_mm:02d}:{_ss:02d}` &nbsp;| "
                    f"Máx. restante: `{_rm:02d}:{_rs:02d}`"
                )

                _todos_listos = True
                _debug_lines  = []
                for _sol in _sols:
                    _id_sol = _sol["id"]
                    if not _id_sol:
                        _debug_lines.append(
                            f"⚠️ **{_sol['tipo']}**: sin ID (rechazada al enviar)"
                        )
                        _sol["estado"] = "SIN_ID"
                        _errores.add("")
                        continue
                    if _id_sol in _terminados or _id_sol in _errores:
                        continue
                    try:
                        _st   = _sat.recover_comprobante_status(_id_sol)
                        _enum = _st.get("EstadoSolicitud")  # int 1-6
                        _cod  = _st.get("CodigoEstadoSolicitud", "")
                        _num  = _st.get("NumeroCFDIs", 0)
                        _paq  = _st.get("IdsPaquetes") or []
                        _sol["paquetes"] = _paq
                        _sol["estado"]   = str(int(_enum)) if _enum is not None else "ERR"
                        _debug_lines.append(
                            f"📋 **{_sol['tipo']}** → "
                            f"Estado=`{_enum}` | Cod=`{_cod}` | CFDIs=`{_num}` | Paq=`{len(_paq)}`"
                        )
                        _en = int(_enum) if _enum is not None else -1
                        if _en == 3:
                            _terminados.add(_id_sol)
                            _stat_ph.success(
                                f"✅ {_sol['tipo'].capitalize()}: "
                                f"{_num} CFDI(s), {len(_paq)} paquete(s). "
                                "Presiona 📥 Descargar paquetes."
                            )
                        elif _en in (4, 5, 6):
                            _errores.add(_id_sol)
                            _lbl_e = {4:"Error en SAT",5:"Rechazada",6:"Vencida"}.get(_en, str(_en))
                            _stat_ph.error(
                                f"❌ {_sol['tipo'].capitalize()}: {_lbl_e} "
                                f"(Cód: {_cod}) — haz una nueva solicitud."
                            )
                        else:
                            _todos_listos = False
                            _lbl_esp = {1:"Aceptada, en cola…",2:"SAT procesando…"}.get(_en, f"Estado {_en}")
                            _stat_ph.info(f"⏳ {_sol['tipo'].capitalize()}: {_lbl_esp}")
                    except Exception as _exc_v:
                        _todos_listos = False
                        _debug_lines.append(f"❌ **{_sol['tipo']}** → Excepción: `{_exc_v}`")
                        _stat_ph.error(f"Error consultando {_sol['tipo']}: {_exc_v}")

                _diag_txt = "  \n".join(_debug_lines) if _debug_lines else "Sin datos aún."
                _debug_ph.info("🔍 **Diagnóstico SAT:**  \n" + _diag_txt)

                _n_l = len(_terminados) + len(_errores)
                _n_t = max(len(_sols), 1)
                _prog_ph.progress(_n_l / _n_t, f"Verificadas: {_n_l}/{_n_t}")

                if _todos_listos or _elapsed >= _MAX_ESP:
                    break
                time.sleep(_POLL)

            _ef = time.time() - _t0
            _mf2, _sf2 = divmod(int(_ef), 60)
            if _terminados:
                _timer_ph.success(
                    f"⏱️ Listo en **{_mf2:02d}:{_sf2:02d}** — presiona 📥 Descargar paquetes"
                )
            else:
                _timer_ph.warning(
                    f"⏱️ Verificación en **{_mf2:02d}:{_sf2:02d}** — revisa el diagnóstico"
                )
            _prog_ph.empty()
            st.session_state["cfdi_solicitudes"] = _sols
            st.rerun()
# ── BOTÓN 3: DESCARGAR PAQUETES (independiente) ──────────────────────────────
with _btn_descargar:
    if st.button("📥 Descargar paquetes", key="btn_descargar", use_container_width=True):
        _sols = st.session_state.get("cfdi_solicitudes", [])
        _listos = [s for s in _sols if s.get("paquetes")]
        if not _listos:
            st.warning("No hay paquetes listos. Primero **Verifica el estado**.")
        else:
            _t0_dl = time.time()
            _nuevos = list(st.session_state.get("cfdi_resultados", []))
            _timer_dl = st.empty()
            _stat_dl  = st.empty()
            _total_paq = sum(len(s["paquetes"]) for s in _listos)
            _prog_dl   = st.progress(0, f"0 / {_total_paq} paquetes descargados")
            _descargados = 0

            for _sol in _listos:
                for _id_paq in _sol.get("paquetes", []):
                    try:
                        _el = time.time() - _t0_dl
                        _mm, _ss = divmod(int(_el), 60)
                        _timer_dl.markdown(
                            f"⏱️ **Descargando…** `{_mm:02d}:{_ss:02d}` — "
                            f"paquete `{_id_paq[:8]}…`"
                        )
                        _meta_paq, _zip_b64 = _sat.recover_comprobante_download(_id_paq)
                        _xmls = _unzip_xmls(_zip_b64)
                        for _fname, _xbytes in _xmls:
                            _m = _parse_meta(_xbytes)
                            _m["direccion"] = _sol["tipo"]
                            _nuevos.append(_m)
                        _descargados += 1
                        _prog_dl.progress(
                            _descargados / _total_paq,
                            f"{_descargados} / {_total_paq} paquetes descargados"
                        )
                    except Exception as _ep:
                        _stat_dl.warning(f"Error en paquete {_id_paq[:8]}: {_ep}")

            _ef2 = time.time() - _t0_dl
            _mf2, _sf2 = divmod(int(_ef2), 60)
            _timer_dl.success(
                f"⏱️ Descarga completa en **{_mf2:02d}:{_sf2:02d}** — "
                f"{len(_nuevos)} CFDI(s) obtenidos."
            )
            _prog_dl.empty()
            st.session_state["cfdi_resultados"] = _nuevos
            if _nuevos:
                st.rerun()

with _btn_limpiar:
    if st.button("🗑️ Limpiar", key="btn_limpiar", use_container_width=True):
        st.session_state["cfdi_solicitudes"] = []
        st.session_state["cfdi_resultados"]  = []
        st.session_state["cfdi_preview_idx"] = None
        st.rerun()

# Mostrar estado actual de solicitudes
_sols_act = st.session_state.get("cfdi_solicitudes", [])
if _sols_act:
    _est_labels = {
        "1": "Aceptada","2": "En proceso","3": "Terminada",
        "4": "Error","5": "Rechazada","6": "Vencida",
        "ENVIADA": "Enviada","SIN_ID": "Sin ID (SAT rechazó solicitud)","ERR": "Error al verificar",
    }
    for _s in _sols_act:
        _id_disp = _s['id'] if _s['id'] else "_(vacío)_"
        _lbl = _est_labels.get(str(_s.get("estado","")),"—")
        st.caption(f"**{_s['tipo'].capitalize()}** — ID: `{_id_disp}` — Estado: {_lbl} — Paquetes: {len(_s.get('paquetes',[]))}")

st.markdown('</div>', unsafe_allow_html=True)

# ════════════════════════════════════════════════════════════════════════════
# SECCIÓN 4 — RESULTADOS + FILTROS + DESCARGA
# ════════════════════════════════════════════════════════════════════════════
_resultados = st.session_state.get("cfdi_resultados", [])
if not _resultados:
    st.stop()

st.markdown(f'<div class="sec-header">📋 4. Resultados — {len(_resultados)} CFDI(s)</div>', unsafe_allow_html=True)
st.markdown('<div class="sec-body">', unsafe_allow_html=True)

# ── Filtros sobre resultados ─────────────────────────────────────────────────
_fca, _fcb, _fcc, _fcd, _fce = st.columns(5)
with _fca:
    _f_uuid = st.text_input("Buscar UUID", placeholder="parcial o completo", key="f_uuid").strip().upper()
with _fcb:
    _f_rfc_e = st.text_input("RFC Emisor", placeholder="filtrar…", key="f_rfc_e").strip().upper()
with _fcc:
    _f_tipo = st.selectbox("Tipo", ["Todos","Ingreso","Egreso","Pago","Nómina","Traslado"], key="f_tipo")
with _fcd:
    _f_estado = st.selectbox("Estatus", ["Todos","Vigente","Cancelado","—"], key="f_estado")
with _fce:
    _f_dir = st.selectbox("Dirección", ["Todos","emitidos","recibidos"], key="f_dir")

# Aplicar filtros
_filtrados = _resultados
if _f_uuid:
    _filtrados = [r for r in _filtrados if _f_uuid in r.get("uuid","")]
if _f_rfc_e:
    _filtrados = [r for r in _filtrados if _f_rfc_e in r.get("emisor_rfc","")]
if _f_tipo != "Todos":
    _filtrados = [r for r in _filtrados if r.get("tipo","") == _f_tipo]
if _f_estado != "Todos":
    _filtrados = [r for r in _filtrados if r.get("estado","") == _f_estado]
if _f_dir != "Todos":
    _filtrados = [r for r in _filtrados if r.get("direccion","") == _f_dir]

# ── Botón verificar estatus masivo ───────────────────────────────────────────
_vc1, _vc2 = st.columns([3,1])
with _vc2:
    if st.button("🔍 Verificar estatus de todos", key="btn_ver_est", use_container_width=True):
        _actualizados = 0
        _bar = st.progress(0, "Verificando estatus…")
        _n   = len(_resultados)
        for _i, _rec in enumerate(_resultados):
            try:
                _uuid_v = _rec.get("uuid","")
                if not _uuid_v or _uuid_v == "ERROR":
                    continue
                _status_r = _sat.recover_comprobante_uuid_request(folio=_uuid_v)
                if _status_r:
                    _sid = _status_r.get("IdSolicitud")
                    # Use status from metadata if available
                _bar.progress((_i+1)/_n, f"Verificando {_i+1}/{_n}…")
                _actualizados += 1
            except:
                pass
        _bar.empty()
        st.info(f"Verificación completada ({_actualizados} registros).")
        st.rerun()

# ── Tabla de resultados ──────────────────────────────────────────────────────
_df = pd.DataFrame([{
    "#":            _i + 1,
    "Dirección":    r.get("direccion","—").capitalize(),
    "UUID":         r.get("uuid",""),
    "Fecha":        r.get("fecha",""),
    "Tipo":         r.get("tipo",""),
    "Emisor RFC":   r.get("emisor_rfc",""),
    "Emisor":       r.get("emisor_nombre",""),
    "Receptor RFC": r.get("receptor_rfc",""),
    "Total":        f"${r.get('total',0):,.2f} {r.get('moneda','MXN')}",
    "Estatus":      r.get("estado","—"),
} for _i, r in enumerate(_filtrados)])

if len(_filtrados) == 0:
    st.warning("Sin resultados con los filtros actuales.")
else:
    st.dataframe(_df.drop(columns=["#"]), use_container_width=True, height=320)

    # ── Descargas masivas (ZIP) ──────────────────────────────────────────────
    st.markdown("#### 📦 Descargar todo")
    _da1, _da2, _da3 = st.columns(3)
    with _da1:
        _all_xml = _zip_files([
            (f"{r.get('uuid','SIN_UUID')}.xml", r["xml_bytes"])
            for r in _filtrados if r.get("xml_bytes")
        ])
        st.download_button(
            "📄 Descargar todos los XML",
            data=_all_xml,
            file_name=f"CFDIs_XML_{date.today()}.zip",
            mime="application/zip",
            use_container_width=True,
            key="dl_all_xml",
        )
    with _da2:
        if st.button("🖨️ Generar y descargar todos los PDF", key="btn_all_pdf", use_container_width=True):
            _pdf_files = []
            _pbar = st.progress(0, "Generando PDFs…")
            for _pi, _r in enumerate(_filtrados):
                _pdfb = _pdf_bytes_from_xml(_r["xml_bytes"])
                if _pdfb:
                    _pdf_files.append((f"{_r.get('uuid','SIN_UUID')}.pdf", _pdfb))
                _pbar.progress((_pi+1)/len(_filtrados))
            _pbar.empty()
            if _pdf_files:
                _zip_pdf = _zip_files(_pdf_files)
                st.download_button(
                    "💾 Descargar ZIP de PDFs",
                    data=_zip_pdf,
                    file_name=f"CFDIs_PDF_{date.today()}.zip",
                    mime="application/zip",
                    key="dl_all_pdf_zip",
                )
    with _da3:
        if st.button("📦 Generar y descargar todo (XML + PDF)", key="btn_all_ambos", use_container_width=True):
            _ambos_files = []
            _pbar2 = st.progress(0, "Generando archivos…")
            for _pi, _r in enumerate(_filtrados):
                _uid = _r.get("uuid","SIN_UUID")
                _ambos_files.append((f"xml/{_uid}.xml", _r["xml_bytes"]))
                _pdfb = _pdf_bytes_from_xml(_r["xml_bytes"])
                if _pdfb:
                    _ambos_files.append((f"pdf/{_uid}.pdf", _pdfb))
                _pbar2.progress((_pi+1)/len(_filtrados))
            _pbar2.empty()
            if _ambos_files:
                _zip_ambos = _zip_files(_ambos_files)
                st.download_button(
                    "💾 Descargar ZIP completo",
                    data=_zip_ambos,
                    file_name=f"CFDIs_Completo_{date.today()}.zip",
                    mime="application/zip",
                    key="dl_all_ambos_zip",
                )

    st.divider()

    # ── Tabla detallada con descarga individual ──────────────────────────────
    st.markdown("#### 📄 Descarga individual")
    for _i_r, _rec in enumerate(_filtrados):
        _uuid_r  = _rec.get("uuid","SIN_UUID")
        _tipo_r  = _rec.get("tipo","")
        _fecha_r = _rec.get("fecha","")
        _emit_r  = _rec.get("emisor_rfc","")
        _total_r = _rec.get("total",0)
        _mon_r   = _rec.get("moneda","MXN")
        _est_r   = _rec.get("estado","—")
        _dir_r   = _rec.get("direccion","").capitalize()

        with st.expander(
            f"{'📤' if _rec.get('direccion')=='emitidos' else '📥'} "
            f"{_uuid_r[:8]}… | {_fecha_r} | {_tipo_r} | {_emit_r} | "
            f"${_total_r:,.2f} {_mon_r} | {_est_r}",
            expanded=False,
        ):
            _ec1, _ec2, _ec3 = st.columns(3)
            _ec4, _ec5       = st.columns(2)

            with _ec1:
                st.caption(f"**UUID:** {_uuid_r}")
                st.caption(f"**Fecha:** {_fecha_r}")
                st.caption(f"**Tipo:** {_tipo_r} ({_dir_r})")
            with _ec2:
                st.caption(f"**Emisor:** {_emit_r}")
                st.caption(f"**Nombre:** {_rec.get('emisor_nombre','')}")
            with _ec3:
                st.caption(f"**Receptor:** {_rec.get('receptor_rfc','')}")
                st.caption(f"**Total:** ${_total_r:,.2f} {_mon_r}")
                _badge = _estado_badge(_est_r)
                st.markdown(f"**Estatus:** {_badge}", unsafe_allow_html=True)

            # Botones de descarga individuales
            _db1, _db2, _db3, _db4 = st.columns(4)
            with _db1:
                st.download_button(
                    "📄 XML",
                    data=_rec["xml_bytes"],
                    file_name=f"{_uuid_r}.xml",
                    mime="application/xml",
                    key=f"dl_xml_{_uuid_r}_{_i_r}",
                    use_container_width=True,
                )
            with _db2:
                if st.button("🖨️ Generar PDF", key=f"btn_pdf_{_uuid_r}_{_i_r}", use_container_width=True):
                    with st.spinner("Generando PDF…"):
                        _pdfb = _pdf_bytes_from_xml(_rec["xml_bytes"])
                    if _pdfb:
                        st.download_button(
                            "💾 Descargar PDF",
                            data=_pdfb,
                            file_name=f"{_uuid_r}.pdf",
                            mime="application/pdf",
                            key=f"dl_pdf_{_uuid_r}_{_i_r}",
                            use_container_width=True,
                        )
            with _db3:
                if st.button("📦 XML + PDF", key=f"btn_ambos_{_uuid_r}_{_i_r}", use_container_width=True):
                    with st.spinner("Generando…"):
                        _pdfb2 = _pdf_bytes_from_xml(_rec["xml_bytes"])
                    _archivos_zip = [(_uuid_r + ".xml", _rec["xml_bytes"])]
                    if _pdfb2:
                        _archivos_zip.append((_uuid_r + ".pdf", _pdfb2))
                    _zip_ind = _zip_files(_archivos_zip)
                    st.download_button(
                        "💾 Descargar ZIP",
                        data=_zip_ind,
                        file_name=f"{_uuid_r}.zip",
                        mime="application/zip",
                        key=f"dl_ambos_{_uuid_r}_{_i_r}",
                        use_container_width=True,
                    )
            with _db4:
                if st.button("👁️ Ver CFDI", key=f"btn_prev_{_uuid_r}_{_i_r}", use_container_width=True):
                    st.session_state["cfdi_preview_idx"] = _i_r

            # Preview inline (HTML renderizado del CFDI)
            if st.session_state.get("cfdi_preview_idx") == _i_r:
                with st.spinner("Renderizando…"):
                    _html_prev = _html_from_xml(_rec["xml_bytes"])
                if _html_prev:
                    st.components.v1.html(_html_prev, height=700, scrolling=True)
                else:
                    st.warning("No se pudo generar la vista previa.")

st.markdown('</div>', unsafe_allow_html=True)
