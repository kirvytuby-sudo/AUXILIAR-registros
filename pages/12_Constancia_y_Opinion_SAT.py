"""
pages/12_Constancia_y_Opinion_SAT.py — Módulo Constancia y Opinión SAT
Genera con un solo botón, para una o varias empresas:
  · Constancia de Situación Fiscal (CSF)
  · Opinión de Cumplimiento de Obligaciones Fiscales (32-D)
Se autentica en el portal del SAT con la e.firma (FIEL): .cer + .key + contraseña.
Las credenciales NUNCA se guardan: viven solo en la memoria de la sesión.
"""
import streamlit as st
import re
import zipfile
from io import BytesIO
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, TimeoutError as FuturoTimeout

st.set_page_config(
    page_title="Constancia y Opinión SAT | Auxiliar de Registros",
    page_icon="🏛️",
    layout="wide",
)

# ─── Estilos (tema azul corporativo) ──────────────────────────────────────────
st.markdown("""
<style>
[data-testid="stAppViewContainer"] { background: #dbeafe; }
.block-container { padding-top: 1.5rem; padding-bottom: 2rem; }
.sat-header {
    background: linear-gradient(135deg, #1E3A8A 0%, #2563EB 100%);
    color: white; border-radius: 10px;
    padding: 1.2rem 1.8rem; margin-bottom: 1.5rem;
}
.sat-header h1 { margin: 0; font-size: 1.6rem; font-weight: 700; letter-spacing: .5px; }
.sat-header p  { margin: .3rem 0 0; opacity: .85; font-size: .9rem; }
.emp-card {
    background: #eff6ff; border: 1.5px solid #93c5fd; border-radius: 8px;
    padding: .9rem 1.2rem; margin-bottom: .6rem;
}
.emp-card .rfc  { font-weight: 700; color: #1E3A8A; font-size: 1.05rem; }
.emp-card .name { color: #334155; font-size: .88rem; }
.emp-card .vig  { color: #64748B; font-size: .78rem; margin-top: .15rem; }
.ok-box   { background: #DCFCE7; border-radius: 6px; padding: .4rem .9rem; color: #166534; font-weight: 600; }
.err-box  { background: #FEE2E2; border-radius: 6px; padding: .4rem .9rem; color: #991B1B; font-weight: 600; }
.warn-box { background: #FEF3C7; border-radius: 6px; padding: .4rem .9rem; color: #92400E; font-weight: 600; }
.priv-note { background:#F1F5F9; border:1px dashed #94A3B8; border-radius:8px;
             padding:.6rem 1rem; color:#475569; font-size:.82rem; margin-bottom:1rem; }
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="sat-header">
    <h1>🏛️ Constancia y Opinión SAT</h1>
    <p>Genera la Constancia de Situación Fiscal y la Opinión de Cumplimiento (32-D)
       de todas tus empresas con un solo botón, usando la e.firma.</p>
</div>
""", unsafe_allow_html=True)

st.markdown("""
<div class="priv-note">🔒 <b>Privacidad:</b> los archivos de la e.firma y las contraseñas
solo se usan en memoria durante esta sesión para autenticarse con el SAT.
No se guardan en ningún servidor ni en el repositorio.</div>
""", unsafe_allow_html=True)

# ─── Dependencia satcfdi ──────────────────────────────────────────────────────
try:
    from satcfdi.models import Signer
    from satcfdi.portal import SATPortalConstancia, SATPortalOpinionCumplimiento
    from satcfdi.exceptions import CFDIError
except ImportError:
    st.error("Falta la librería **satcfdi**. Agrega `satcfdi` a `requirements.txt` "
             "y reinicia la app para usar este módulo.")
    st.stop()

RFC_RE = re.compile(r"[A-ZÑ&]{3,4}\d{6}[A-Z0-9]{3}", re.IGNORECASE)
TIMEOUT_SEG = 150  # tiempo máximo por documento


# ─── Utilidades ───────────────────────────────────────────────────────────────
def _stem(nombre: str) -> str:
    base = nombre.rsplit("/", 1)[-1]
    return base.rsplit(".", 1)[0].strip().lower()


def _info_cer(cer_bytes: bytes):
    """Lee el .cer y regresa (rfc, razon_social, vigencia_fin, tipo)."""
    from satcfdi.models import Certificate
    cert = Certificate.load_certificate(cer_bytes) if hasattr(Certificate, "load_certificate") else None
    if cert is None:
        from OpenSSL import crypto
        x509 = crypto.load_certificate(crypto.FILETYPE_ASN1, cer_bytes)
        cert = Certificate(x509)
    rfc = (cert.rfc or "").upper()
    nombre = cert.legal_name or ""
    not_after = cert.certificate.get_notAfter()  # b'AAAAMMDDHHMMSSZ'
    vig = datetime.strptime(not_after.decode()[:14], "%Y%m%d%H%M%S")
    tipo = str(cert.type.name) if cert.type else "?"
    return rfc, nombre, vig, tipo


def _emparejar(cers: dict, keys: dict):
    """Empareja archivos .key con .cer: por nombre idéntico, por RFC en el
    nombre del .key, o por orden. Regresa lista de dicts."""
    pares, usados = [], set()

    def _tomar(cer_nombre, key_nombre):
        usados.add(cer_nombre)
        pares.append({"cer": cer_nombre, "key": key_nombre})

    for key_nombre in keys:
        ks = _stem(key_nombre)
        # 1) mismo nombre base
        match = next((c for c in cers if c not in usados and _stem(c) == ks), None)
        # 2) RFC del cer aparece en el nombre del key (o viceversa)
        if not match:
            m = RFC_RE.search(ks.upper())
            rfc_key = m.group(0).upper() if m else None
            if rfc_key:
                match = next((c for c in cers if c not in usados
                              and cers[c]["rfc"] == rfc_key), None)
        # 3) primer cer libre
        if not match:
            match = next((c for c in cers if c not in usados), None)
        if match:
            _tomar(match, key_nombre)
    return pares


def _con_timeout(fn, segundos):
    """Ejecuta fn() con límite de tiempo; el portal del SAT a veces se cuelga."""
    with ThreadPoolExecutor(max_workers=1) as ex:
        fut = ex.submit(fn)
        return fut.result(timeout=segundos)


def _descargar_documentos(cer_bytes, key_bytes, password):
    """Autentica con la e.firma y descarga (constancia, opinion).
    Regresa dict con resultados y errores por documento."""
    res = {"rfc": None, "nombre": None,
           "constancia": None, "err_constancia": None,
           "opinion": None, "err_opinion": None, "err_general": None}
    try:
        signer = Signer.load(certificate=cer_bytes, key=key_bytes, password=password)
    except CFDIError:
        res["err_general"] = "El archivo .key NO corresponde a este .cer (llave y certificado no coinciden)."
        return res
    except ValueError:
        res["err_general"] = "Contraseña de la e.firma incorrecta (no se pudo abrir el .key)."
        return res
    except Exception as e:
        res["err_general"] = f"No se pudo leer la e.firma: {e}"
        return res

    res["rfc"] = signer.rfc
    res["nombre"] = signer.legal_name or ""

    # Vigencia
    try:
        vence = datetime.strptime(signer.certificate.get_notAfter().decode()[:14], "%Y%m%d%H%M%S")
        if vence < datetime.utcnow():
            res["err_general"] = f"La e.firma está VENCIDA (venció el {vence:%d/%m/%Y}). Renuévala en el SAT."
            return res
    except Exception:
        pass

    # 1) Constancia de Situación Fiscal
    try:
        pdf = _con_timeout(lambda: SATPortalConstancia(signer).generar_constancia(), TIMEOUT_SEG)
        if pdf and pdf[:4] == b"%PDF":
            res["constancia"] = pdf
        else:
            res["err_constancia"] = "El SAT no regresó un PDF válido (intenta de nuevo en unos minutos)."
    except FuturoTimeout:
        res["err_constancia"] = "El portal del SAT tardó demasiado (timeout). Intenta más tarde."
    except AssertionError:
        res["err_constancia"] = "El portal del SAT rechazó la sesión (¿e.firma revocada o portal en mantenimiento?)."
    except Exception as e:
        res["err_constancia"] = f"Error del portal SAT: {e}"

    # 2) Opinión de Cumplimiento 32-D
    try:
        pdf = _con_timeout(lambda: SATPortalOpinionCumplimiento(signer).generar_opinion_cumplimiento(), TIMEOUT_SEG)
        if pdf and pdf[:4] == b"%PDF":
            res["opinion"] = pdf
        else:
            res["err_opinion"] = "El SAT no regresó un PDF válido de la opinión (intenta de nuevo)."
    except FuturoTimeout:
        res["err_opinion"] = "El portal del SAT tardó demasiado (timeout). Intenta más tarde."
    except AssertionError:
        res["err_opinion"] = "El portal del SAT rechazó la sesión para la opinión 32-D."
    except Exception as e:
        res["err_opinion"] = f"Error del portal SAT: {e}"

    return res


# ─── 1. Carga de e.firmas ─────────────────────────────────────────────────────
st.subheader("1️⃣ Sube las e.firmas (.cer y .key)")
st.caption("Puedes subir los archivos de **varias empresas a la vez** — se emparejan solos.")

archivos = st.file_uploader(
    "Arrastra aquí los .cer y .key de todas las empresas",
    type=["cer", "key"], accept_multiple_files=True, key="efirmas",
)

cers, keys = {}, {}
for f in archivos or []:
    data = f.getvalue()
    if f.name.lower().endswith(".cer"):
        try:
            rfc, nombre, vig, tipo = _info_cer(data)
            cers[f.name] = {"bytes": data, "rfc": rfc, "nombre": nombre, "vig": vig, "tipo": tipo}
        except Exception:
            st.markdown(f'<div class="err-box">❌ No se pudo leer <b>{f.name}</b> — ¿es un .cer válido?</div>',
                        unsafe_allow_html=True)
    else:
        keys[f.name] = data

# Avisos de certificados CSD (sello) en lugar de FIEL
for nombre_c, info in cers.items():
    if info["tipo"].upper() == "CSD":
        st.markdown(
            f'<div class="warn-box">⚠️ <b>{nombre_c}</b> es un certificado de <b>sello (CSD)</b>, '
            f'no la e.firma (FIEL). El portal del SAT solo acepta la <b>e.firma</b>.</div>',
            unsafe_allow_html=True)

pares = _emparejar(cers, keys) if (cers and keys) else []

# ─── 2. Empresas detectadas y contraseñas ────────────────────────────────────
empresas = []
if pares:
    st.subheader("2️⃣ Empresas detectadas")
    misma_pwd = False
    if len(pares) > 1:
        misma_pwd = st.checkbox("Usar la misma contraseña para todas las empresas")
    pwd_comun = st.text_input("Contraseña de la e.firma (todas)", type="password",
                              key="pwd_comun") if misma_pwd else None

    for i, par in enumerate(pares):
        info = cers[par["cer"]]
        c1, c2 = st.columns([3, 2])
        with c1:
            st.markdown(f"""
            <div class="emp-card">
                <div class="rfc">🏢 {info['rfc'] or '(RFC no leído)'}</div>
                <div class="name">{info['nombre']}</div>
                <div class="vig">📜 {par['cer']} &nbsp;·&nbsp; 🔑 {par['key']}
                     &nbsp;·&nbsp; vigente hasta {info['vig']:%d/%m/%Y}</div>
            </div>""", unsafe_allow_html=True)
        with c2:
            pwd = pwd_comun if misma_pwd else st.text_input(
                f"Contraseña e.firma — {info['rfc']}", type="password", key=f"pwd_{i}")
        empresas.append({"cer": info["bytes"], "key": keys[par["key"]],
                         "pwd": pwd or "", "rfc": info["rfc"], "nombre": info["nombre"]})

    faltan_key = [c for c in cers if c not in {p["cer"] for p in pares}]
    for c in faltan_key:
        st.markdown(f'<div class="warn-box">⚠️ <b>{c}</b> no tiene archivo .key — súbelo para incluir esta empresa.</div>',
                    unsafe_allow_html=True)
elif cers and not keys:
    st.info("Ya subiste los .cer — ahora falta el archivo **.key** de cada empresa.")
elif keys and not cers:
    st.info("Ya subiste los .key — ahora falta el archivo **.cer** de cada empresa.")

# ─── 3. Botón único ───────────────────────────────────────────────────────────
if empresas:
    st.subheader("3️⃣ Generar documentos")
    sin_pwd = [e["rfc"] for e in empresas if not e["pwd"]]
    if sin_pwd:
        st.caption(f"✏️ Escribe la contraseña de: {', '.join(sin_pwd)}")

    if st.button("🏛️ GENERAR CONSTANCIAS Y OPINIONES", type="primary",
                 use_container_width=True, disabled=bool(sin_pwd)):
        resultados = []
        barra = st.progress(0.0, text="Conectando con el SAT…")
        for i, emp in enumerate(empresas):
            barra.progress(i / len(empresas),
                           text=f"🏢 {emp['rfc']} — autenticando con e.firma y generando documentos…")
            r = _descargar_documentos(emp["cer"], emp["key"], emp["pwd"])
            r["rfc"] = r["rfc"] or emp["rfc"]
            r["nombre"] = r["nombre"] or emp["nombre"]
            resultados.append(r)
        barra.progress(1.0, text="✅ Proceso terminado")
        st.session_state["sat_resultados"] = resultados
        st.session_state["sat_timestamp"] = datetime.now().strftime("%Y%m%d_%H%M")

# ─── 4. Resultados ────────────────────────────────────────────────────────────
resultados = st.session_state.get("sat_resultados")
if resultados:
    st.divider()
    st.subheader("📄 Resultados")
    ts = st.session_state.get("sat_timestamp", datetime.now().strftime("%Y%m%d_%H%M"))
    zip_buf = BytesIO()
    hay_algo = False

    with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for r in resultados:
            rfc = r["rfc"] or "SIN_RFC"
            st.markdown(f"""
            <div class="emp-card"><div class="rfc">🏢 {rfc}</div>
            <div class="name">{r['nombre'] or ''}</div></div>""", unsafe_allow_html=True)

            if r["err_general"]:
                st.markdown(f'<div class="err-box">❌ {r["err_general"]}</div>', unsafe_allow_html=True)
                continue

            c1, c2 = st.columns(2)
            with c1:
                if r["constancia"]:
                    hay_algo = True
                    nombre_pdf = f"{rfc}_Constancia_Situacion_Fiscal_{ts}.pdf"
                    zf.writestr(nombre_pdf, r["constancia"])
                    st.download_button(f"📥 Constancia de Situación Fiscal — {rfc}",
                                       data=r["constancia"], file_name=nombre_pdf,
                                       mime="application/pdf", key=f"dl_csf_{rfc}",
                                       use_container_width=True)
                else:
                    st.markdown(f'<div class="err-box">❌ Constancia: {r["err_constancia"]}</div>',
                                unsafe_allow_html=True)
            with c2:
                if r["opinion"]:
                    hay_algo = True
                    nombre_pdf = f"{rfc}_Opinion_Cumplimiento_32D_{ts}.pdf"
                    zf.writestr(nombre_pdf, r["opinion"])
                    st.download_button(f"📥 Opinión de Cumplimiento 32-D — {rfc}",
                                       data=r["opinion"], file_name=nombre_pdf,
                                       mime="application/pdf", key=f"dl_op_{rfc}",
                                       use_container_width=True)
                else:
                    st.markdown(f'<div class="err-box">❌ Opinión 32-D: {r["err_opinion"]}</div>',
                                unsafe_allow_html=True)

    if hay_algo:
        st.divider()
        st.download_button("🗜️ DESCARGAR TODO (ZIP)", data=zip_buf.getvalue(),
                           file_name=f"SAT_Constancias_y_Opiniones_{ts}.zip",
                           mime="application/zip", type="primary",
                           use_container_width=True, key="dl_zip")

# ─── Ayuda ────────────────────────────────────────────────────────────────────
with st.expander("❓ ¿Qué necesito? / Preguntas frecuentes"):
    st.markdown("""
- **e.firma (FIEL) vigente** de cada empresa: archivo **.cer**, archivo **.key** y su **contraseña**.
  *No sirve el CSD (certificado de sello digital) ni la Contraseña/CIEC.*
- Los documentos se generan directamente en los portales oficiales del SAT
  (Constancia: `rfcampc.siat.sat.gob.mx` · Opinión 32-D: `ptsc32d.clouda.sat.gob.mx`).
- Si el SAT está en mantenimiento (frecuente en la noche o en cierres de mes),
  vuelve a intentar más tarde.
- La **Opinión 32-D** indica si la empresa está al corriente (sentido *positivo* o *negativo*).
- Nada se guarda: al cerrar o recargar la página, las e.firmas y contraseñas se borran de la memoria.
""")
