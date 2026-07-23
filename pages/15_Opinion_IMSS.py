"""
pages/15_Opinion_IMSS.py — Módulo Opinión de Cumplimiento IMSS
Genera la Opinión de Cumplimiento de Obligaciones Fiscales en materia
de Seguridad Social (Art. 32-D CFF) directamente del Buzón IMSS.
Autentica con la e.firma (FIEL): .cer + .key + contraseña.
Portal oficial: https://buzon.imss.gob.mx/buzonimss/login
Sin captcha manual — automatización headless con Playwright.
"""
import streamlit as st
import base64
import hashlib
import hmac
import re
import secrets as _secrets_mod
import zipfile
from io import BytesIO
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, TimeoutError as _FutureTimeout

st.set_page_config(
    page_title="Opinión IMSS | Auxiliar de Registros",
    page_icon="🏥",
    layout="wide",
)

import _theme
_theme.aplicar_header("🏥 Opinión de Cumplimiento IMSS",
                      "Genera la Opinión 32-D del IMSS con e.firma • Buzón IMSS")

RFC_RE   = re.compile(r"[A-ZÑ&]{3,4}\d{6}[A-Z0-9]{3}", re.IGNORECASE)
TIMEOUT  = 180   # segundos máx por empresa

# ── Verificar / instalar Playwright una vez ───────────────────────────────────
@st.cache_resource(show_spinner=False)
def _preparar_playwright():
    """Instala chromium la primera vez que corre en Streamlit Cloud."""
    import subprocess, sys
    try:
        from playwright.sync_api import sync_playwright  # noqa: F401
    except ImportError:
        subprocess.run([sys.executable, "-m", "pip", "install", "playwright"],
                       capture_output=True)
    try:
        subprocess.run(["playwright", "install", "chromium", "--with-deps"],
                       capture_output=True, timeout=120)
        return True
    except Exception:
        return False

_playwright_ok = _preparar_playwright()

# ── Secrets / Supabase helpers ────────────────────────────────────────────────
def _get_secret(clave, default=None):
    try:
        return st.secrets.get(clave, default)
    except Exception:
        return default

@st.cache_resource(ttl=60)
def _sb_client():
    try:
        from supabase import create_client
        return create_client(st.secrets["supabase"]["url"],
                             st.secrets["supabase"]["key"])
    except Exception:
        return None

def _sb_fernet():
    try:
        from cryptography.fernet import Fernet
        return Fernet(st.secrets["supabase"]["enc_key"].encode())
    except Exception:
        return None

@st.cache_data(ttl=30)
def _sb_load_empresas():
    sb = _sb_client(); f = _sb_fernet()
    if not sb or not f: return {}
    try:
        rows = sb.table("imss_empresas").select("*").execute().data or []
    except Exception:
        return {}
    result = {}
    for row in rows:
        rfc = row["rfc"]
        try:
            import base64 as _b64
            cer_b = f.decrypt(row["cer_enc"].encode())
            key_b = f.decrypt(row["key_enc"].encode())
            pwd   = f.decrypt(row["pwd_enc"].encode()).decode()
            result[rfc] = {
                "nombre":  row.get("nombre", rfc),
                "cer_b64": _b64.b64encode(cer_b).decode(),
                "key_b64": _b64.b64encode(key_b).decode(),
                "password": pwd,
                "nrp": row.get("nrp", ""),
            }
        except Exception:
            pass
    return result

def _sb_save_empresa(rfc, nombre, cer_bytes, key_bytes, pwd, nrp, vigencia):
    sb = _sb_client(); f = _sb_fernet()
    if not sb or not f:
        st.error("Supabase no configurado en Secrets."); return False
    try:
        _pwd_b = pwd.encode() if isinstance(pwd, str) else (pwd or b"")
        sb.table("imss_empresas").upsert({
            "rfc":     rfc,
            "nombre":  nombre,
            "cer_enc": f.encrypt(cer_bytes).decode(),
            "key_enc": f.encrypt(key_bytes).decode(),
            "pwd_enc": f.encrypt(_pwd_b).decode(),
            "nrp":     nrp or "",
            "vigencia": vigencia.strftime("%d/%m/%Y") if hasattr(vigencia, "strftime") else str(vigencia or ""),
        }).execute()
        return True
    except Exception as e:
        st.error(f"Error guardando en Supabase: {e}"); return False

def _sb_delete_empresa(rfc):
    sb = _sb_client()
    if not sb: st.error("Supabase no configurado."); return False
    try:
        sb.table("imss_empresas").delete().eq("rfc", rfc).execute()
        _sb_load_empresas.clear(); return True
    except Exception as e:
        st.error(f"Error borrando: {e}"); return False

_USE_SUPABASE = bool(_sb_client() and _sb_fernet())

# ── Autenticación con contraseña (igual que SAT) ──────────────────────────────
def _pw_hash(password, salt):
    dk = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"),
                              salt.encode("utf-8"), 260_000)
    return dk.hex()

def _pw_verify(password, stored):
    try:
        salt, expected = stored.split(":", 1)
        return hmac.compare_digest(_pw_hash(password, salt), expected)
    except Exception:
        return False

@st.cache_resource
def _get_pendientes():
    return {"lista": []}

_imss_users = None
try:
    _raw = st.secrets.get("imss_users")
    if _raw: _imss_users = dict(_raw)
except Exception:
    pass

if _imss_users:
    if not st.session_state.get("imss_auth_user"):
        col_l, col_c, col_r = st.columns([1, 1.2, 1])
        with col_c:
            st.markdown("#### 🔐 Acceso al módulo IMSS")
            _usr = st.text_input("Usuario", key="imss_login_user", placeholder="tu usuario")
            _pwd = st.text_input("Contraseña", type="password", key="imss_login_pwd")
            if st.button("Entrar →", type="primary", use_container_width=True,
                         key="imss_login_btn"):
                _datos = _imss_users.get(_usr.strip().lower())
                if _datos and _pw_verify(_pwd, _datos.get("password_hash", "")):
                    st.session_state["imss_auth_user"] = _usr.strip().lower()
                    st.session_state["imss_auth_name"] = _datos.get("name", _usr.upper())
                    st.rerun()
                else:
                    st.error("❌ Usuario o contraseña incorrectos.")
            st.markdown("---")
            with st.expander("📝 Solicitar acceso"):
                with st.form("form_solicitud_imss"):
                    s_nombre  = st.text_input("Nombre completo")
                    s_usuario = st.text_input("Usuario deseado")
                    s_pwd1    = st.text_input("Contraseña", type="password")
                    s_pwd2    = st.text_input("Confirmar contraseña", type="password")
                    if st.form_submit_button("📨 Enviar", use_container_width=True):
                        _u = s_usuario.strip().lower().replace(" ", "_")
                        if not all([s_nombre.strip(), _u, s_pwd1]):
                            st.error("Completa todos los campos.")
                        elif s_pwd1 != s_pwd2:
                            st.error("Las contraseñas no coinciden.")
                        else:
                            _s2 = _secrets_mod.token_hex(16)
                            _get_pendientes()["lista"].append({
                                "nombre": s_nombre.strip(), "usuario": _u,
                                "password_hash": f"{_s2}:{_pw_hash(s_pwd1, _s2)}",
                                "fecha": datetime.now().strftime("%d/%m/%Y %H:%M"),
                            })
                            st.success("✅ Solicitud enviada.")
        st.stop()

    _auth_name = st.session_state.get("imss_auth_name", "")
    _auth_user = st.session_state.get("imss_auth_user", "")
    _es_admin  = (_auth_user == "kirvy")
    with st.sidebar:
        st.markdown(f"👤 **{_auth_name}**")
        st.caption("Módulo IMSS")
        if st.button("🚪 Cerrar sesión", key="imss_logout"):
            st.session_state.pop("imss_auth_user", None)
            st.session_state.pop("imss_auth_name", None)
            st.rerun()
        if _es_admin:
            st.markdown("---")
            st.markdown("**⚙️ Administración**")
            with st.expander("👥 Usuarios con acceso"):
                for _un, _ud in (_imss_users or {}).items():
                    _ud2 = _ud.get("name", _un.upper()) if isinstance(_ud, dict) else _un.upper()
                    st.markdown(f"• **{_ud2}** — `{_un}`")
            with st.expander("➕ Agregar usuario"):
                _na = st.text_input("Nombre", key="adm_i_nombre")
                _ua = st.text_input("Usuario", key="adm_i_user")
                _pa = st.text_input("Contraseña", type="password", key="adm_i_pwd")
                if st.button("Generar Secrets", key="adm_i_gen"):
                    if _na and _ua and _pa:
                        _sa = _secrets_mod.token_hex(16)
                        _ha = f"{_sa}:{_pw_hash(_pa, _sa)}"
                        st.session_state["adm_i_toml"] = (
                            f"[imss_users.{_ua.lower()}]\n"
                            f'name = "{_na}"\n'
                            f'password_hash = "{_ha}"'
                        )
                    else:
                        st.error("Completa todos los campos.")
                if st.session_state.get("adm_i_toml"):
                    st.code(st.session_state["adm_i_toml"], language="toml")
                    if st.button("✔ Listo", key="adm_i_done"):
                        del st.session_state["adm_i_toml"]
                        st.rerun()
            _pend = _get_pendientes()["lista"]
            if _pend:
                st.warning(f"📬 {len(_pend)} solicitud(es) pendiente(s)")
                for _i, _req in enumerate(_pend):
                    with st.expander(f"👤 {_req['nombre']} — @{_req['usuario']}"):
                        st.code(
                            f"[imss_users.{_req['usuario']}]\n"
                            f"name = \"{_req['nombre']}\"\n"
                            f"password_hash = \"{_req['password_hash']}\"",
                            language="toml")
                        _ca, _cb = st.columns(2)
                        with _ca:
                            if st.button("✅ Aprobar", key=f"apr_i_{_i}",
                                         type="primary", use_container_width=True):
                                _get_pendientes()["lista"].remove(_req)
                                st.rerun()
                        with _cb:
                            if st.button("❌ Rechazar", key=f"rec_i_{_i}",
                                         use_container_width=True):
                                _get_pendientes()["lista"].remove(_req)
                                st.rerun()

# ── Utilidades e.firma ────────────────────────────────────────────────────────
def _stem(nombre):
    return nombre.rsplit("/", 1)[-1].rsplit(".", 1)[0].strip().lower()

def _info_cer(cer_bytes):
    """Lee el .cer con satcfdi (ya está como dependencia) y regresa
    (rfc, razon_social, vigencia_fin). Fallback: cryptography."""
    try:
        from satcfdi.models import Certificate
        cert = Certificate.load_certificate(cer_bytes)
        rfc  = (cert.rfc or "").upper()
        nombre = cert.legal_name or ""
        not_after = cert.certificate.get_notAfter()
        vig = datetime.strptime(not_after.decode()[:14], "%Y%m%d%H%M%S")
        return rfc, nombre, vig
    except Exception:
        pass
    try:
        from cryptography import x509
        from cryptography.hazmat.backends import default_backend
        cert = x509.load_der_x509_certificate(cer_bytes, default_backend())
        rfc = ""
        for attr in cert.subject:
            if attr.oid.dotted_string in ("2.5.4.45", "1.2.840.113549.1.9.1"):
                val = attr.value
                m = RFC_RE.search(val.upper())
                if m: rfc = m.group(0)
        nombre = ""
        try: nombre = cert.subject.get_attributes_for_oid(x509.NameOID.COMMON_NAME)[0].value
        except Exception: pass
        vig = cert.not_valid_after_utc.replace(tzinfo=None)
        return rfc, nombre, vig
    except Exception as e:
        raise ValueError(f"No se pudo leer el .cer: {e}")

def _emparejar(cers, keys):
    pares, usados = [], set()
    def _tomar(cn, kn): usados.add(cn); pares.append({"cer": cn, "key": kn})
    for kn in keys:
        ks = _stem(kn)
        match = next((c for c in cers if c not in usados and _stem(c) == ks), None)
        if not match:
            m = RFC_RE.search(ks.upper())
            rfc_k = m.group(0).upper() if m else None
            if rfc_k:
                match = next((c for c in cers if c not in usados
                              and cers[c]["rfc"] == rfc_k), None)
        if not match:
            match = next((c for c in cers if c not in usados), None)
        if match: _tomar(match, kn)
    return pares

# ── Automatización Buzón IMSS ────────────────────────────────────────────────
def _descargar_opinion_imss(cer_bytes: bytes, key_bytes: bytes,
                             password: str, nrp: str = "") -> dict:
    """
    Abre el Buzón IMSS en Chromium headless, autentica con e.firma
    y descarga la Opinión de Cumplimiento como PDF.
    Regresa dict con claves: rfc, nombre, pdf, error.
    """
    res = {"rfc": "", "nombre": "", "pdf": None, "error": None}

    # Extraer RFC del .cer antes de abrir el navegador
    try:
        rfc, nombre, vig = _info_cer(cer_bytes)
        res["rfc"] = rfc; res["nombre"] = nombre
        if vig < datetime.utcnow():
            res["error"] = f"La e.firma está VENCIDA (venció el {vig:%d/%m/%Y})."
            return res
    except Exception as e:
        res["error"] = f"No se pudo leer el .cer: {e}"
        return res

    import tempfile, os, pathlib

    # Guardar .cer y .key en archivos temporales para que Playwright los suba
    tmp_cer = tempfile.NamedTemporaryFile(suffix=".cer", delete=False)
    tmp_key = tempfile.NamedTemporaryFile(suffix=".key", delete=False)
    tmp_cer.write(cer_bytes); tmp_cer.close()
    tmp_key.write(key_bytes); tmp_key.close()
    tmp_dl  = tempfile.mkdtemp()   # carpeta de descarga

    try:
        from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

        with sync_playwright() as pw:
            browser = pw.chromium.launch(
                headless=True,
                args=["--no-sandbox", "--disable-setuid-sandbox",
                      "--disable-dev-shm-usage"],
            )
            ctx = browser.new_context(
                accept_downloads=True,
                locale="es-MX",
            )
            page = ctx.new_page()
            page.set_default_timeout(60_000)

            # ── 1. Ir al Buzón IMSS ──────────────────────────────────────
            page.goto("https://buzon.imss.gob.mx/buzonimss/login",
                      wait_until="networkidle")

            # ── 2. Seleccionar autenticación con e.firma ─────────────────
            # El portal puede mostrar pestañas: "Persona Física" / "e.firma"
            try:
                efirma_tab = page.locator(
                    "text=e.firma, text=E.FIRMA, text=Firma Electrónica"
                ).first
                if efirma_tab.is_visible(timeout=5000):
                    efirma_tab.click()
            except PWTimeout:
                pass   # ya está en modo e.firma

            # ── 3. Subir .cer ────────────────────────────────────────────
            try:
                cer_input = page.locator(
                    "input[type='file'][accept*='cer'], "
                    "input[type='file']:near(:text('certificado')), "
                    "input[type='file']:near(:text('.cer'))"
                ).first
                cer_input.set_input_files(tmp_cer.name)
            except Exception as e:
                res["error"] = f"No se encontró campo para subir .cer: {e}"
                browser.close(); return res

            # ── 4. Subir .key ────────────────────────────────────────────
            try:
                key_input = page.locator(
                    "input[type='file'][accept*='key'], "
                    "input[type='file']:near(:text('llave')), "
                    "input[type='file']:near(:text('.key'))"
                ).nth(1)
                key_input.set_input_files(tmp_key.name)
            except Exception as e:
                res["error"] = f"No se encontró campo para subir .key: {e}"
                browser.close(); return res

            # ── 5. Contraseña de la llave privada ────────────────────────
            try:
                pwd_input = page.locator(
                    "input[type='password']:near(:text('contraseña')), "
                    "input[type='password']:near(:text('clave')), "
                    "input[type='password']"
                ).first
                pwd_input.fill(password)
            except Exception as e:
                res["error"] = f"No se encontró campo de contraseña: {e}"
                browser.close(); return res

            # ── 6. NRP opcional ──────────────────────────────────────────
            if nrp:
                try:
                    nrp_input = page.locator(
                        "input:near(:text('NRP')), "
                        "input:near(:text('registro patronal')), "
                        "input[placeholder*='NRP']"
                    ).first
                    if nrp_input.is_visible(timeout=3000):
                        nrp_input.fill(nrp)
                except PWTimeout:
                    pass   # NRP no requerido en este paso

            # ── 7. Ingresar ──────────────────────────────────────────────
            try:
                btn_ingresar = page.locator(
                    "button:has-text('Ingresar'), "
                    "button:has-text('INGRESAR'), "
                    "input[type='submit']:near(:text('Ingresar'))"
                ).first
                btn_ingresar.click()
                page.wait_for_load_state("networkidle", timeout=30_000)
            except Exception as e:
                res["error"] = f"Error al hacer clic en Ingresar: {e}"
                browser.close(); return res

            # Verificar error de login
            for err_text in ["contraseña incorrecta", "datos incorrectos",
                             "no coinciden", "credenciales inválidas"]:
                if page.locator(f"text={err_text}").is_visible(timeout=2000):
                    res["error"] = "Contraseña de e.firma incorrecta o certificado no válido."
                    browser.close(); return res

            # ── 8. Navegar a Opinión de Cumplimiento ─────────────────────
            # Intentar URL directa primero
            try:
                page.goto("https://buzon.imss.gob.mx/opinioncumplimiento/",
                          wait_until="networkidle", timeout=30_000)
            except PWTimeout:
                pass

            # Buscar enlace / botón "Opinión" / "32D" en la página
            try:
                opinion_link = page.locator(
                    "text=Opinión, text=OPINIÓN, text=32D, text=Cumplimiento"
                ).first
                if opinion_link.is_visible(timeout=5000):
                    opinion_link.click()
                    page.wait_for_load_state("networkidle", timeout=20_000)
            except PWTimeout:
                pass

            # ── 9. Descargar / generar el PDF ────────────────────────────
            pdf_bytes = None

            # A) Esperar un botón de descarga / "Consultar"
            for selector in [
                "button:has-text('Consultar')",
                "button:has-text('Generar')",
                "button:has-text('Descargar')",
                "a:has-text('Opinión')",
                "button:has-text('Obtener')",
            ]:
                try:
                    btn = page.locator(selector).first
                    if btn.is_visible(timeout=4000):
                        with ctx.expect_download(timeout=60_000) as dl_info:
                            btn.click()
                        dl = dl_info.value
                        import shutil
                        dest = os.path.join(tmp_dl, dl.suggested_filename or "opinion.pdf")
                        dl.save_as(dest)
                        with open(dest, "rb") as fh:
                            pdf_bytes = fh.read()
                        break
                except PWTimeout:
                    continue

            # B) Si no hubo descarga, capturar el PDF de la página actual
            if not pdf_bytes:
                try:
                    pdf_bytes = page.pdf(
                        format="Letter",
                        print_background=True,
                        margin={"top": "10mm", "bottom": "10mm",
                                "left": "10mm", "right": "10mm"},
                    )
                except Exception:
                    pass

            browser.close()

            if pdf_bytes and pdf_bytes[:4] == b"%PDF":
                res["pdf"] = pdf_bytes
            else:
                res["error"] = ("El Buzón IMSS no devolvió un PDF válido. "
                                "El portal puede estar en mantenimiento o la sesión expiró. "
                                "Intenta de nuevo en unos minutos.")
    except ImportError:
        res["error"] = ("La librería Playwright no está instalada. "
                        "Agrega `playwright` a `requirements.txt` y reinicia la app.")
    except Exception as e:
        import traceback
        res["error"] = f"Error inesperado: {e}\n{traceback.format_exc()[:800]}"
    finally:
        os.unlink(tmp_cer.name)
        os.unlink(tmp_key.name)
        import shutil
        shutil.rmtree(tmp_dl, ignore_errors=True)

    return res


def _con_timeout(fn, segundos):
    with ThreadPoolExecutor(max_workers=1) as ex:
        fut = ex.submit(fn)
        return fut.result(timeout=segundos)


# ─── Estado global ────────────────────────────────────────────────────────────
empresas = []
toml_empresas = {}

# ─── 0. Empresas guardadas ────────────────────────────────────────────────────
_guardadas = _sb_load_empresas() if _USE_SUPABASE else (_get_secret("imss_empresas", {}) or {})

if _guardadas:
    st.subheader("🔐 Empresas guardadas")
    st.caption("E.firmas guardadas en la base de datos — solo marca y genera.")
    for clave_emp in sorted(_guardadas.keys()):
        datos = _guardadas[clave_emp]
        try:
            cer_b = base64.b64decode(datos["cer_b64"])
            rfc, nombre, vig = _info_cer(cer_b)
        except Exception as e:
            st.error(f"❌ {clave_emp}: datos mal guardados ({e})"); continue
        c1, c2 = st.columns([1, 8])
        with c1:
            usar = st.checkbox(" ", value=True, key=f"isec_{clave_emp}",
                               label_visibility="collapsed")
        with c2:
            nrp_g = datos.get("nrp", "")
            st.markdown(f"""
            <div style="background:#EFF6FF;border-left:4px solid #1E3A8A;
                        padding:8px 14px;border-radius:6px;margin:4px 0;">
              <div style="font-weight:700;color:#1E3A8A;font-size:1.05rem;">
                🏥 {rfc}</div>
              <div style="color:#374151;">{nombre or datos.get('nombre', clave_emp)}</div>
              <div style="color:#6B7280;font-size:.85rem;">
                vigente hasta {vig:%d/%m/%Y}
                {"&nbsp;·&nbsp; NRP: " + nrp_g if nrp_g else ""}
              </div>
            </div>""", unsafe_allow_html=True)
        if usar:
            empresas.append({"cer": cer_b,
                             "key": base64.b64decode(datos["key_b64"]),
                             "pwd": datos.get("password", ""),
                             "nrp": datos.get("nrp", ""),
                             "rfc": rfc, "nombre": nombre})

    _marcadas  = [c for c in sorted(_guardadas.keys())
                  if st.session_state.get(f"isec_{c}", True)]
    _sin_pwd_g = [c for c in _marcadas if not _guardadas[c].get("password")]
    if _sin_pwd_g:
        st.caption(f"⚠️ Falta contraseña guardada para: {', '.join(_sin_pwd_g)}")
    _n_sel = len(_marcadas)
    if st.button(
        f"🤖  GENERAR  —  {_n_sel} empresa{'s' if _n_sel != 1 else ''}"
            if _n_sel else "🤖  GENERAR  —  (ninguna seleccionada)",
        type="primary", use_container_width=True,
        key="btn_auto_imss", disabled=(bool(_sin_pwd_g) or _n_sel == 0),
    ):
        _auto_list = []
        for _k in _marcadas:
            _d = _guardadas[_k]
            try:
                _cb = base64.b64decode(_d["cer_b64"])
                _kb = base64.b64decode(_d["key_b64"])
                _r2, _n2, _ = _info_cer(_cb)
                _auto_list.append({"cer": _cb, "key": _kb,
                                   "pwd": _d.get("password", ""),
                                   "nrp": _d.get("nrp", ""),
                                   "rfc": _r2, "nombre": _n2})
            except Exception as _ex:
                st.error(f"Error leyendo {_k}: {_ex}")
        if _auto_list:
            _res_auto = []
            _bar = st.progress(0.0, text="Conectando con el Buzón IMSS…")
            for _ia, _ea in enumerate(_auto_list):
                _bar.progress(_ia / len(_auto_list),
                              text=f"🏥 {_ea['rfc']} — generando Opinión IMSS…")
                try:
                    _ra = _con_timeout(
                        lambda ea=_ea: _descargar_opinion_imss(
                            ea["cer"], ea["key"], ea["pwd"], ea.get("nrp", "")),
                        TIMEOUT)
                except _FutureTimeout:
                    _ra = {"rfc": _ea["rfc"], "nombre": _ea["nombre"],
                           "pdf": None,
                           "error": "El portal IMSS tardó demasiado (timeout). Intenta más tarde."}
                _ra["rfc"]    = _ra.get("rfc") or _ea["rfc"]
                _ra["nombre"] = _ra.get("nombre") or _ea["nombre"]
                _res_auto.append(_ra)
            _bar.progress(1.0, text="✅ Terminado")
            st.session_state["imss_resultados"] = _res_auto
            st.session_state["imss_timestamp"]  = datetime.now().strftime("%Y%m%d_%H%M")
            st.rerun()

    # ── Quitar empresa guardada ──────────────────────────────────────────
    st.markdown("#### 🗑️ Quitar empresa guardada")
    _claves_del = sorted(_guardadas.keys())
    _cd1, _cd2 = st.columns([3, 1])
    with _cd1:
        _esel = st.selectbox(
            "Empresa", _claves_del, key="del_imss_sel",
            format_func=lambda k: f"{k}  —  {str(_guardadas[k].get('nombre', k))[:50]}",
            label_visibility="collapsed",
        )
    with _cd2:
        if st.button("🗑️ Quitar", key="btn_del_imss",
                     use_container_width=True, type="primary"):
            if _USE_SUPABASE:
                if _sb_delete_empresa(_esel):
                    st.success(f"✅ Empresa {_esel} eliminada."); st.rerun()
            else:
                _restantes = {k: v for k, v in _guardadas.items() if k != _esel}
                if _restantes:
                    _lineas = []
                    for _ek, _ed in _restantes.items():
                        _lineas.append(f"[imss_empresas.{_ek}]")
                        if isinstance(_ed, dict):
                            for _ek2, _ev2 in _ed.items():
                                _lineas.append(f'  {_ek2} = "{_ev2}"')
                        _lineas.append("")
                    st.session_state["del_imss_toml"] = "\n".join(_lineas).strip()
                else:
                    st.session_state["del_imss_toml"] = ""
                st.session_state["del_imss_nombre"] = _esel
    if not _USE_SUPABASE and "del_imss_toml" in st.session_state:
        _toml_d = st.session_state["del_imss_toml"]
        _qn     = st.session_state.get("del_imss_nombre", "")
        if _toml_d:
            st.warning(f"Reemplaza la sección `[imss_empresas]` en Secrets para borrar **{_qn}**:")
            st.code(_toml_d, language="toml")
        else:
            st.warning(f"Al borrar **{_qn}** ya no quedan empresas. Elimina toda la sección.")
        if st.button("✔ Listo, ya lo actualicé", key="btn_del_imss_done"):
            del st.session_state["del_imss_toml"]
            st.session_state.pop("del_imss_nombre", None); st.rerun()

    st.divider()

# ─── 1. Carga manual de e.firmas ──────────────────────────────────────────────
titulo_m = "➕ Agregar otras empresas (.cer y .key)" if _guardadas \
           else "1️⃣ Sube las e.firmas (.cer y .key)"
st.subheader(titulo_m)
st.caption("Sube los archivos de **varias empresas a la vez** — se emparejan solos.")

archivos = st.file_uploader(
    "Arrastra aquí los .cer y .key",
    type=["cer", "key"], accept_multiple_files=True, key="imss_efirmas",
)

cers, keys = {}, {}
for f in archivos or []:
    data = f.getvalue()
    if f.name.lower().endswith(".cer"):
        try:
            rfc, nombre, vig = _info_cer(data)
            cers[f.name] = {"bytes": data, "rfc": rfc, "nombre": nombre, "vig": vig}
        except Exception:
            st.error(f"❌ No se pudo leer **{f.name}** — ¿es un .cer válido?")
    else:
        keys[f.name] = data

pares = _emparejar(cers, keys) if (cers and keys) else []

# ─── 2. Contraseñas y NRP ────────────────────────────────────────────────────
if pares:
    st.subheader("2️⃣ Empresas detectadas")
    misma_pwd = len(pares) > 1 and st.checkbox(
        "Usar la misma contraseña para todas")
    pwd_comun = st.text_input("Contraseña de la e.firma (todas)", type="password",
                              key="imss_pwd_comun") if misma_pwd else None
    misma_nrp = len(pares) > 1 and st.checkbox(
        "NRP no requerido / igual para todas",
        help="Si las empresas tienen distinto NRP, déjalo en blanco y el portal lo detecta solo.")
    nrp_comun = st.text_input("NRP (todas)", key="imss_nrp_comun",
                               placeholder="Ej: A1234567890") if misma_nrp else None

    for i, par in enumerate(pares):
        info = cers[par["cer"]]
        c1, c2 = st.columns([3, 2])
        with c1:
            st.markdown(f"""
            <div style="background:#EFF6FF;border-left:4px solid #1E3A8A;
                        padding:8px 14px;border-radius:6px;margin:4px 0;">
              <div style="font-weight:700;color:#1E3A8A;">🏥 {info['rfc'] or '(RFC no leído)'}</div>
              <div>{info['nombre']}</div>
              <div style="color:#6B7280;font-size:.85rem;">
                📜 {par['cer']} &nbsp;·&nbsp; 🔑 {par['key']}
                &nbsp;·&nbsp; vigente hasta {info['vig']:%d/%m/%Y}
              </div>
            </div>""", unsafe_allow_html=True)
        with c2:
            pwd = pwd_comun if misma_pwd else st.text_input(
                f"Contraseña e.firma — {info['rfc']}", type="password",
                key=f"imss_pwd_{i}")
            nrp = nrp_comun if misma_nrp else st.text_input(
                f"NRP (opcional) — {info['rfc']}", key=f"imss_nrp_{i}",
                placeholder="A1234567890")
        empresas.append({"cer": info["bytes"], "key": keys[par["key"]],
                         "pwd": pwd or "", "nrp": nrp or "",
                         "rfc": info["rfc"], "nombre": info["nombre"]})
        toml_empresas[info["rfc"]] = {
            "nombre": info["nombre"], "cer": info["bytes"],
            "key": keys[par["key"]], "pwd": pwd or "", "nrp": nrp or "",
        }

    faltan_key = [c for c in cers if c not in {p["cer"] for p in pares}]
    for c in faltan_key:
        st.warning(f"⚠️ **{c}** no tiene archivo .key correspondiente.")
elif cers and not keys:
    st.info("Ya subiste los .cer — falta el archivo **.key** de cada empresa.")
elif keys and not cers:
    st.info("Ya subiste los .key — falta el archivo **.cer** de cada empresa.")

# ─── 2.5 Guardar empresas ─────────────────────────────────────────────────────
if toml_empresas:
    with st.expander("💾 Guardar estas empresas para no subirlas cada vez"):
        if _USE_SUPABASE:
            st.markdown("Guarda las e.firmas directamente en la base de datos.")
            if st.button("☁️ Guardar en Supabase", type="primary",
                         use_container_width=True, key="btn_sb_guardar_imss"):
                _sb_ok = True
                for _srfc, _sd in toml_empresas.items():
                    _svig = cers.get(
                        next((p["cer"] for p in pares
                              if cers[p["cer"]]["rfc"] == _srfc), ""), {}).get("vig")
                    if not _sb_save_empresa(
                        _srfc, _sd["nombre"], _sd["cer"], _sd["key"],
                        _sd["pwd"], _sd.get("nrp", ""), _svig
                    ): _sb_ok = False
                if _sb_ok:
                    st.success("✅ Empresa(s) guardadas."); _sb_load_empresas.clear()
        else:
            st.markdown("""
Copia el texto y pégalo en **Settings → Secrets** de Streamlit Cloud.
Las e.firmas quedan cifradas y **nunca** se publican en GitHub.
""")
            incluir_pwd = st.checkbox("Incluir contraseñas", value=True)
            _lineas = []
            for _rfc, _d in toml_empresas.items():
                _lineas.append(f"[imss_empresas.{_rfc}]")
                _lineas.append(f'nombre = "{_d["nombre"]}"')
                _lineas.append(f'cer_b64 = "{base64.b64encode(_d["cer"]).decode()}"')
                _lineas.append(f'key_b64 = "{base64.b64encode(_d["key"]).decode()}"')
                _lineas.append(f'password = "{_d["pwd"] if incluir_pwd else ""}"')
                _lineas.append(f'nrp = "{_d.get("nrp", "")}"')
                _lineas.append("")
            st.code("\n".join(_lineas), language="toml")

# ─── 3. Botón generar ────────────────────────────────────────────────────────
if empresas:
    st.subheader("3️⃣ Generar Opinión de Cumplimiento IMSS")
    sin_pwd = [e["rfc"] for e in empresas if not e["pwd"]]
    if sin_pwd:
        st.caption(f"✏️ Escribe la contraseña de: {', '.join(sin_pwd)}")

    if st.button("🏥 GENERAR OPINIÓN IMSS", type="primary",
                 use_container_width=True, disabled=bool(sin_pwd)):
        resultados = []
        barra = st.progress(0.0, text="Abriendo Buzón IMSS…")
        for i, emp in enumerate(empresas):
            barra.progress(i / len(empresas),
                           text=f"🏥 {emp['rfc']} — autenticando con e.firma…")
            try:
                r = _con_timeout(
                    lambda e=emp: _descargar_opinion_imss(
                        e["cer"], e["key"], e["pwd"], e.get("nrp", "")),
                    TIMEOUT)
            except _FutureTimeout:
                r = {"rfc": emp["rfc"], "nombre": emp["nombre"],
                     "pdf": None,
                     "error": "Timeout — el Buzón IMSS tardó demasiado. Intenta más tarde."}
            r["rfc"]    = r.get("rfc") or emp["rfc"]
            r["nombre"] = r.get("nombre") or emp["nombre"]
            resultados.append(r)
        barra.progress(1.0, text="✅ Proceso terminado")
        st.session_state["imss_resultados"] = resultados
        st.session_state["imss_timestamp"]  = datetime.now().strftime("%Y%m%d_%H%M")

# ─── 4. Resultados ───────────────────────────────────────────────────────────
resultados = st.session_state.get("imss_resultados")
if resultados:
    st.divider()
    st.subheader("📄 Resultados")
    ts      = st.session_state.get("imss_timestamp", datetime.now().strftime("%Y%m%d_%H%M"))
    zip_buf = BytesIO()
    hay_pdf = False

    with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for _i_res, r in enumerate(resultados):
            rfc = r.get("rfc") or "SIN_RFC"
            st.markdown(f"""
            <div style="background:#EFF6FF;border-left:4px solid #1E3A8A;
                        padding:8px 14px;border-radius:6px;margin:8px 0;">
              <div style="font-weight:700;color:#1E3A8A;font-size:1.05rem;">🏥 {rfc}</div>
              <div style="color:#374151;">{r.get('nombre') or ''}</div>
            </div>""", unsafe_allow_html=True)

            if r.get("error"):
                st.error(f"❌ {r['error']}")
                continue

            if r.get("pdf"):
                hay_pdf = True
                nombre_pdf = f"{rfc}_Opinion_Cumplimiento_IMSS_{ts}.pdf"
                zf.writestr(nombre_pdf, r["pdf"])
                st.download_button(
                    f"📥 Descargar Opinión IMSS — {rfc}",
                    data=r["pdf"], file_name=nombre_pdf,
                    mime="application/pdf",
                    key=f"dl_imss_{rfc}_{_i_res}",
                    use_container_width=True, type="primary",
                )
            else:
                st.warning("⚠️ No se obtuvo el PDF. Revisa los errores arriba.")

    if hay_pdf and len(resultados) > 1:
        st.divider()
        st.download_button(
            "🗜️ DESCARGAR TODO (ZIP)",
            data=zip_buf.getvalue(),
            file_name=f"IMSS_Opiniones_{ts}.zip",
            mime="application/zip",
            type="primary", use_container_width=True,
            key="dl_imss_zip",
        )

# ─── Ayuda ────────────────────────────────────────────────────────────────────
with st.expander("❓ ¿Qué necesito? / Preguntas frecuentes"):
    st.markdown("""
- **e.firma (FIEL) vigente** de cada empresa: archivo **.cer**, archivo **.key** y su **contraseña**.
  *(No sirve el CSD / sello digital — solo la e.firma personal del patrón o representante legal.)*
- **NRP (Número de Registro Patronal)** — opcional; el portal lo detecta automáticamente con la e.firma.
- **¿Qué genera este módulo?** La **Opinión de Cumplimiento de Obligaciones Fiscales en materia de
  Seguridad Social** (Art. 32-D CFF) del Buzón IMSS — válida por 15 días naturales.
- **¿Cómo guardo las empresas?** Sube las e.firmas una vez, abre *"💾 Guardar estas empresas…"*,
  copia el texto y pégalo en **Settings → Secrets** de Streamlit Cloud.
- **Portal usado:** `buzon.imss.gob.mx` (vigente desde oct-2025; reemplaza al Escritorio Virtual).
- Si el Buzón IMSS está en mantenimiento, vuelve a intentar más tarde.
- La opinión **positiva** = al corriente · **negativa** = adeudos · **sin opinión** = sin información en el sistema.
""")

st.markdown("---")
st.caption("Módulo Opinión IMSS · Buzón IMSS (Art. 32-D CFF) · Vigencia 15 días · Automatización con Playwright")
