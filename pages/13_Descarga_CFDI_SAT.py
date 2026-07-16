"""
pages/13_Descarga_CFDI_SAT.py
Descarga masiva de CFDIs emitidos y recibidos desde el SAT.
Autenticación: e.firma (.cer + .key + contraseña).
"""

import streamlit as st
import base64, zipfile, io, os, time, json, uuid as _uuid_mod
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


# ── Historial de solicitudes ──────────────────────────────────────────────────
_HIST_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "cfdi_historial.json")

def _hist_cargar() -> list:
    """Carga el historial de solicitudes desde JSON."""
    try:
        if os.path.exists(_HIST_FILE):
            with open(_HIST_FILE, "r", encoding="utf-8") as _f:
                return json.load(_f)
    except Exception:
        pass
    return []

def _hist_guardar(hist: list):
    """Guarda historial (máx. 30 entradas)."""
    try:
        with open(_HIST_FILE, "w", encoding="utf-8") as _f:
            json.dump(hist[:30], _f, ensure_ascii=False, indent=2, default=str)
    except Exception:
        pass

def _hist_agregar(entry: dict):
    """Agrega (o reemplaza) una entrada al inicio del historial."""
    _h = _hist_cargar()
    _h = [x for x in _h if x.get("id_entrada") != entry.get("id_entrada")]
    _h.insert(0, entry)
    _hist_guardar(_h)

def _hist_actualizar(id_entrada: str, **kwargs):
    """Actualiza campos de una entrada existente por id_entrada."""
    _h = _hist_cargar()
    for _e in _h:
        if _e.get("id_entrada") == id_entrada:
            _e.update(kwargs)
            break
    _hist_guardar(_h)


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
    "cfdi_polling":     False,   # True mientras se verifica en background
    "cfdi_poll_start":  None,    # tiempo inicio del polling
    "cfdi_terminados":  [],      # IDs ya terminados
    "cfdi_errores_v":   [],      # IDs con error/rechazada/vencida
    "cfdi_zero_counts": {},      # {pid: n} contador de respuestas EstadoSolicitud=0
    "cfdi_hist_id":     None,    # ID de la entrada activa en el historial
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
# HISTORIAL de solicitudes
# ════════════════════════════════════════════════════════════════════════════
_hist = _hist_cargar()
_hist_lbl = f"📋 Historial — {len(_hist)} solicitud(es) anteriores" if _hist else "📋 Historial — sin solicitudes anteriores"
with st.expander(_hist_lbl, expanded=False):
        if not _hist:
            st.info("Aún no hay solicitudes registradas. Aparecerán aquí después de tu primera solicitud.")
        for _h in _hist:
            _ef  = _h.get("estado_final", "?")
            _ico = (
                "✅" if _ef in ("Terminada", "Descargado") else
                "❌" if _ef in ("Error", "Rechazada", "Vencida", "Sin paquetes") else
                "⏳"
            )
            _ha, _hb, _hc, _hd = st.columns([2.5, 2.5, 1.5, 0.8])
            with _ha:
                st.markdown(f"**{_h.get('fecha_solicitud','?')}**")
                st.caption(f"RFC: {_h.get('rfc','?')}")
                st.caption(f"Período: {_h.get('periodo','?')} · {_h.get('tipo_dl','?')}")
            with _hb:
                for _hs in _h.get("solicitudes", []):
                    _hid  = _hs.get("id", "")
                    _hest = _hs.get("estado", "")
                    if _hid:
                        st.caption(f"**{_hs.get('tipo','').capitalize()}**: `{_hid[:12]}…` [{_hest}]")
                    else:
                        st.caption(f"**{_hs.get('tipo','').capitalize()}**: sin ID")
            with _hc:
                st.markdown(f"{_ico} **{_ef}**")
                st.caption(f"CFDIs: {_h.get('total_cfdis', 0)}")
            with _hd:
                if st.button(
                    "↩ Cargar", key=f"h_load_{_h['id_entrada']}",
                    use_container_width=True,
                    help="Restaurar estas solicitudes en la sesión actual",
                ):
                    _sols_r = [
                        {
                            "tipo":     s["tipo"],
                            "id":       s.get("id", ""),
                            "paquetes": s.get("paquetes", []),
                            "estado":   s.get("estado", "ENVIADA"),
                            "cod":      s.get("cod", ""),
                        }
                        for s in _h.get("solicitudes", [])
                    ]
                    st.session_state["cfdi_solicitudes"] = _sols_r
                    st.session_state["cfdi_resultados"]  = []
                    st.session_state["cfdi_hist_id"]     = _h["id_entrada"]
                    st.rerun()
            st.divider()

        if st.button("🗑️ Borrar historial completo", key="btn_borrar_hist"):
            _hist_guardar([])
            st.rerun()

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

# ── AUTO-POLLING (estado de máquina — un check por rerun) ─────────────────────
# NOTA: El SAT puede tardar 30-60 min en registrar y procesar una solicitud.
# EstadoSolicitud=0 (CodEstatus 5002) = SAT aún no registró la solicitud → seguir esperando.
# No tratar como error permanente hasta que se confirme falla real (4/5/6).
_POLL_INTERVAL = 20   # segundos entre checks (no saturar al SAT)
_POLL_MAX      = 3600 # 1 hora máxima de polling

if st.session_state.get("cfdi_polling") and st.session_state.get("cfdi_signer"):
    _p_start  = st.session_state["cfdi_poll_start"] or time.time()
    _elapsed  = time.time() - _p_start
    _mm, _ss  = divmod(int(_elapsed), 60)
    _rm, _rs  = divmod(int(max(0, _POLL_MAX - _elapsed)), 60)

    _p_sols       = st.session_state.get("cfdi_solicitudes", [])
    _p_terminados = set(st.session_state.get("cfdi_terminados", []))
    _p_errores    = set(st.session_state.get("cfdi_errores_v", []))
    # Contador de respuestas "0" consecutivas por ID de solicitud
    _p_zero_cnt   = dict(st.session_state.get("cfdi_zero_counts", {}))
    _p_sat        = st.session_state["cfdi_sat"]

    with st.container():
        _ph_timer = st.markdown(
            f"⏱️ **Verificando…** Transcurrido: `{_mm:02d}:{_ss:02d}` | "
            f"Máx. restante: `{_rm:02d}:{_rs:02d}`"
        )
        _ph_prog  = st.progress(0, "Verificando solicitudes…")
        _diag     = []

        for _p_sol in _p_sols:
            _pid = _p_sol.get("id","")
            if not _pid:
                _diag.append(f"⚠️ **{_p_sol['tipo']}**: sin ID de solicitud")
                _p_sol["estado"] = "SIN_ID"
                _p_errores.add("__sin_id__" + _p_sol["tipo"])
                continue
            if _pid in _p_terminados or _pid in _p_errores:
                _diag.append(f"✅ **{_p_sol['tipo']}**: ya procesado ({_p_sol.get('estado','')})")
                continue
            try:
                _p_r   = _p_sat.recover_comprobante_status(_pid)
                _p_en  = int(_p_r.get("EstadoSolicitud", -1))
                # CodEstatus = resultado de la llamada (5000=ok, 5002=aún no registrada)
                _p_cod = _p_r.get("CodEstatus", _p_r.get("CodigoEstadoSolicitud", ""))
                _p_num = _p_r.get("NumeroCFDIs", 0)
                _p_paq = _p_r.get("IdsPaquetes") or []
                _p_sol["paquetes"] = _p_paq
                _p_sol["estado"]   = str(_p_en)

                if _p_en == 0:
                    # SAT aún no registró la solicitud (eventual consistency).
                    # Incrementar contador; seguir esperando hasta _POLL_MAX.
                    _p_zero_cnt[_pid] = _p_zero_cnt.get(_pid, 0) + 1
                    _zc = _p_zero_cnt[_pid]
                    if _zc < 6:          # < 2 min: silencioso
                        _diag.append(f"⏳ **{_p_sol['tipo']}**: Esperando que el SAT registre la solicitud… `(Cód:{_p_cod})`")
                    else:                # ≥ 2 min: aviso
                        _diag.append(
                            f"⏳ **{_p_sol['tipo']}**: SAT procesando… `(Cód:{_p_cod}, intento {_zc})`  \n"
                            f"  ℹ️ El SAT puede tardar 30-60 min — esta ventana seguirá intentando."
                        )
                else:
                    _p_zero_cnt.pop(_pid, None)   # resetear contador si ya no es 0
                    _lbl_e = {
                        1: "Aceptada ✅", 2: "En proceso ⏳", 3: "Terminada ✅",
                        4: "Error ❌", 5: "Rechazada ❌", 6: "Vencida ❌",
                    }.get(_p_en, f"Estado desconocido ({_p_en})")
                    _icono = "✅" if _p_en == 3 else "⏳" if _p_en in (1, 2) else "❌"
                    _diag.append(
                        f"{_icono} **{_p_sol['tipo']}**: {_lbl_e} | "
                        f"Cód:`{_p_cod}` | CFDIs:`{_p_num}` | Paq:`{len(_p_paq)}`"
                    )
                    if _p_en == 3:
                        _p_terminados.add(_pid)
                    elif _p_en in (4, 5, 6):
                        _p_errores.add(_pid)

                # Respuesta completa del SAT para debug
                with st.expander(f"🔍 Respuesta SAT cruda ({_p_sol['tipo']})", expanded=False):
                    st.json(_p_r)

            except Exception as _p_exc:
                _diag.append(f"❌ **{_p_sol['tipo']}** → Error en verificación: `{_p_exc}`")
                _p_errores.add(_pid)

        st.info("🔍 **Diagnóstico:**  \n" + "  \n".join(_diag) if _diag else "Sin datos aún.")

        _n_done = len(_p_terminados) + len(_p_errores)
        _n_tot  = max(len(_p_sols), 1)
        _ph_prog.progress(min(_n_done / _n_tot, 1.0), f"{_n_done}/{_n_tot} solicitudes procesadas")

    # Guardar estado actualizado
    st.session_state["cfdi_solicitudes"]  = _p_sols
    st.session_state["cfdi_terminados"]   = list(_p_terminados)
    st.session_state["cfdi_errores_v"]    = list(_p_errores)
    st.session_state["cfdi_zero_counts"]  = _p_zero_cnt

    _todos_ok = all(
        (s.get("id","") in _p_terminados or s.get("id","") in _p_errores or not s.get("id",""))
        for s in _p_sols
    )

    if _todos_ok or _elapsed >= _POLL_MAX:
        st.session_state["cfdi_polling"] = False
        # Determinar estado final: pendiente si aún estaba en 0, sin paquetes si realmente falló
        _hay_pendientes = any(
            s.get("id","") not in _p_terminados and s.get("id","") not in _p_errores
            for s in _p_sols if s.get("id","")
        )
        # ── Actualizar historial con estado final ────────────────────────
        _hist_id_poll = st.session_state.get("cfdi_hist_id")
        if _hist_id_poll:
            _ef_poll = (
                "Terminada" if _p_terminados else
                "Pendiente (reintenta)" if _hay_pendientes else
                "Sin paquetes"
            )
            _hist_actualizar(
                _hist_id_poll,
                estado_final = _ef_poll,
                solicitudes  = [
                    {
                        "tipo":     s["tipo"],
                        "id":       s.get("id", ""),
                        "estado":   s.get("estado", ""),
                        "paquetes": s.get("paquetes", []),
                        "cod":      s.get("cod", ""),
                    }
                    for s in _p_sols
                ],
            )
        if _p_terminados:
            st.success("✅ Verificación completa — presiona **📥 Descargar paquetes**")
        elif _hay_pendientes:
            st.warning(
                "⏳ **El SAT no respondió en 1 hora.** Los CFDIs pueden estar listos más tarde.  \n"
                "Usa el botón **↩ Cargar** del Historial y vuelve a presionar **🔄 Verificar estado**."
            )
        else:
            st.warning("⚠️ Verificación completada: el SAT reportó error/rechazo en la solicitud.")
    else:
        time.sleep(_POLL_INTERVAL)
        st.rerun()

_btn_solicitar, _btn_descargar, _btn_verificar, _btn_limpiar = st.columns([2, 2, 2, 1])

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

            # ── Separar aceptadas vs rechazadas ─────────────────────────
            _aceptadas  = [s for s in _solicitudes_nuevas
                           if s.get("cod","") in ("5000","5004") and s.get("id","")]
            _rechazadas = [s for s in _solicitudes_nuevas
                           if s not in _aceptadas]

            # Solo enviar al polling las solicitudes ACEPTADAS con ID válido
            st.session_state["cfdi_solicitudes"] = _aceptadas
            st.session_state["cfdi_resultados"]  = []
            st.session_state["cfdi_start_time"]  = time.time()

            # ── Mostrar resultado por solicitud ──────────────────────────
            for _sol in _aceptadas:
                st.success(f"✅ Solicitud **{_sol['tipo']}** enviada — ID: `{_sol['id']}`")
            for _sol in _rechazadas:
                _cod_r = _sol.get("cod","")
                _msg_r = {
                    "301": "El SAT no permite descargar **Recibidos** para este RFC. Usa solo **Emitidos**.",
                    "5002": "Límite de solicitudes alcanzado. Intenta mañana.",
                    "5005": "Solicitud duplicada en proceso.",
                }.get(_cod_r, f"SAT rechazó la solicitud (Cód: {_cod_r}).")
                st.warning(f"⚠️ **{_sol['tipo'].capitalize()}** rechazada — {_msg_r}")

            # ── Guardar en historial ─────────────────────────────────────
            _hist_id_new  = str(_uuid_mod.uuid4())
            st.session_state["cfdi_hist_id"] = _hist_id_new
            _periodo_h = (
                f"UUID: {_uuid_input}" if _modo == "🆔 Por UUID"
                else f"{_fecha_ini} al {_fecha_fin}"
            )
            _hist_agregar({
                "id_entrada":      _hist_id_new,
                "fecha_solicitud": datetime.now().strftime("%Y-%m-%d %H:%M"),
                "rfc":             _rfc,
                "periodo":         _periodo_h,
                "tipo_dl":         _tipo_dl if _modo != "🆔 Por UUID" else _tipo_dir,
                "solicitudes": [
                    {
                        "tipo":     s["tipo"],
                        "id":       s["id"],
                        "estado":   s["estado"],
                        "paquetes": s.get("paquetes", []),
                        "cod":      s.get("cod", ""),
                    }
                    for s in _solicitudes_nuevas
                ],
                "total_cfdis":  0,
                "estado_final": "Enviada",
            })
            if _aceptadas:
                st.info(
                    "ℹ️ **El SAT tarda entre 30 minutos y varias horas en procesar.**  \n"
                    "Presiona **🔄 Verificar estado** para monitorear — "
                    "revisará automáticamente cada 20 segundos hasta 1 hora."
                )
        except Exception as _e:
            st.error(f"Error al solicitar: {_e}")

# ── BOTÓN 2: DESCARGAR PAQUETES ──────────────────────────────────────────────
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
            # ── Actualizar historial con total de CFDIs ──────────────────
            _hist_id_dl = st.session_state.get("cfdi_hist_id")
            if _hist_id_dl:
                _hist_actualizar(
                    _hist_id_dl,
                    total_cfdis  = len(_nuevos),
                    estado_final = "Descargado",
                )
            if _nuevos:
                st.rerun()

# ── BOTÓN 3: VERIFICAR ESTADO ────────────────────────────────────────────────
with _btn_verificar:
    _is_polling = st.session_state.get("cfdi_polling", False)
    _btn_lbl    = "⏹ Detener verificación" if _is_polling else "🔄 Verificar estado"
    if st.button(_btn_lbl, key="btn_verificar", use_container_width=True):
        _sols = st.session_state.get("cfdi_solicitudes", [])
        if _is_polling:
            # Detener polling
            st.session_state["cfdi_polling"] = False
            st.rerun()
        elif not _sols:
            st.warning("Primero haz clic en **Solicitar descarga**.")
        else:
            # Iniciar polling (resetear contadores)
            st.session_state["cfdi_polling"]     = True
            st.session_state["cfdi_poll_start"]  = time.time()
            st.session_state["cfdi_terminados"]  = []
            st.session_state["cfdi_errores_v"]   = []
            st.session_state["cfdi_zero_counts"] = {}
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
