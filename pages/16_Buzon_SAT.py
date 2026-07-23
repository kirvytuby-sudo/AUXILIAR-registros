"""
pages/16_Buzon_SAT.py — Buzón Tributario SAT
Consulta notificaciones electrónicas y mensajes de interés del SAT
para múltiples empresas. Autentica con e.firma (FIEL): .cer + .key + contraseña.
Portal: https://wwwmat.sat.gob.mx/iniciar-expediente/mis-notificaciones/
Sin captcha — login criptográfico con e.firma.
"""
import streamlit as st
import base64
import hashlib
import hmac
import json
import os
import subprocess
import sys
import tempfile
import threading
import time
from datetime import datetime
from io import BytesIO

st.set_page_config(
    page_title="Buzón SAT | Auxiliar de Registros",
    page_icon="🗂",
    layout="wide",
)

import _theme
_theme.aplicar_header("🗂 Buzón Tributario SAT",
                      "Notificaciones y Comunicados oficiales del SAT con e.firma")

# ── Constantes ────────────────────────────────────────────────────────────────
TIMEOUT = 180

# ── Playwright setup ──────────────────────────────────────────────────────────
@st.cache_resource(show_spinner=False)
def _preparar_playwright():
    try:
        from playwright.sync_api import sync_playwright  # noqa
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

# ── Auth helpers ──────────────────────────────────────────────────────────────
def _get_secret(key, default=None):
    try:
        return st.secrets.get(key, default)
    except Exception:
        return default

def _pw_verify(password: str, stored: str) -> bool:
    try:
        salt, hx = stored.split(":", 1)
        dk = hashlib.pbkdf2_hmac("sha256", password.encode(),
                                  salt.encode(), 260_000)
        return hmac.compare_digest(dk.hex(), hx)
    except Exception:
        return False

def _check_auth():
    users = _get_secret("buzon_users", {})
    if not users:
        return True  # sin configuración → abierto
    if st.session_state.get("buzon_auth"):
        return True
    with st.form("login_buzon"):
        st.subheader("🔐 Acceso — Buzón SAT")
        user = st.text_input("Usuario")
        pw   = st.text_input("Contraseña", type="password")
        ok   = st.form_submit_button("Entrar")
    if ok:
        stored = (users or {}).get(user, {}).get("hash", "")
        if stored and _pw_verify(pw, stored):
            st.session_state["buzon_auth"] = True
            st.rerun()
        else:
            st.error("Credenciales incorrectas.")
    return False

if not _check_auth():
    st.stop()

# ── Supabase helpers (empresas) ───────────────────────────────────────────────
@st.cache_resource(ttl=60)
def _sb_client():
    try:
        from supabase import create_client
        url  = _get_secret("SUPABASE_URL")
        key  = _get_secret("SUPABASE_KEY")
        if url and key:
            return create_client(url, key)
    except Exception:
        pass
    return None

def _sb_fernet():
    """Fernet con la misma clave que usa Constancia y Opinión SAT (tabla 'empresas')."""
    try:
        from cryptography.fernet import Fernet
        return Fernet(st.secrets["supabase"]["enc_key"].encode())
    except Exception:
        return None


@st.cache_data(ttl=30)
def _cargar_empresas_sat_docs(rfcs_ya_en_buzon: frozenset = frozenset()):
    """Lee la tabla 'empresas' (Constancia y Opinión SAT) y devuelve empresas
    en formato normalizado (cer_enc/key_enc/pwd_enc en base64 simple) con _sat_docs=True.
    Excluye RFCs que ya están en buzon_sat_empresas para evitar duplicados."""
    sb = _sb_client()
    f  = _sb_fernet()
    if not sb or not f:
        return []
    try:
        rows = sb.table("empresas").select("*").execute().data or []
    except Exception:
        return []
    result = []
    for row in rows:
        rfc = row.get("rfc", "")
        if not rfc or rfc in rfcs_ya_en_buzon:
            continue
        try:
            cer_bytes = f.decrypt(row["cer_enc"].encode())
            key_bytes = f.decrypt(row["key_enc"].encode())
            pwd_bytes = f.decrypt(row["pwd_enc"].encode())
            result.append({
                "rfc":       rfc,
                "nombre":    row.get("nombre", rfc),
                "cer_enc":   base64.b64encode(cer_bytes).decode(),
                "key_enc":   base64.b64encode(key_bytes).decode(),
                "pwd_enc":   base64.b64encode(pwd_bytes).decode(),
                "_sat_docs": True,
            })
        except Exception:
            pass
    return result


def _cargar_empresas():
    # 1. Buzón nativo
    sb = _sb_client()
    buzon = []
    if sb:
        try:
            res = sb.table("buzon_sat_empresas").select("*").execute()
            buzon = res.data or []
        except Exception:
            pass
    if not buzon:
        buzon = st.session_state.get("buzon_empresas_local", [])
    # 2. SAT Docs (Constancia y Opinión SAT) — solo RFCs nuevos
    rfcs_buzon = frozenset(e.get("rfc", "") for e in buzon)
    sat_docs = _cargar_empresas_sat_docs(rfcs_buzon)
    return buzon + sat_docs

def _guardar_empresa(rfc, nombre, cer_enc, key_enc, pwd_enc):
    emp = {"rfc": rfc, "nombre": nombre,
           "cer_enc": cer_enc, "key_enc": key_enc, "pwd_enc": pwd_enc}
    sb = _sb_client()
    if sb:
        try:
            sb.table("buzon_sat_empresas").upsert(emp, on_conflict="rfc").execute()
            return
        except Exception:
            pass
    lst = st.session_state.get("buzon_empresas_local", [])
    lst = [e for e in lst if e.get("rfc") != rfc]
    lst.append(emp)
    st.session_state["buzon_empresas_local"] = lst

def _eliminar_empresa(rfc):
    sb = _sb_client()
    if sb:
        try:
            sb.table("buzon_sat_empresas").delete().eq("rfc", rfc).execute()
            return
        except Exception:
            pass
    lst = st.session_state.get("buzon_empresas_local", [])
    st.session_state["buzon_empresas_local"] = [e for e in lst if e.get("rfc") != rfc]

# ── Encrypt / decrypt (base64 simple en web) ──────────────────────────────────
def _enc(data: bytes) -> str:
    return base64.b64encode(data).decode()

def _dec(b64: str) -> bytes:
    return base64.b64decode(b64)

# ── Script Playwright inline ──────────────────────────────────────────────────
_PLAYWRIGHT_LIST_SCRIPT = """
import sys, os, time, json
try:
    from playwright.sync_api import sync_playwright
except ImportError:
    print("ERR:playwright no disponible"); sys.exit(1)

CER = sys.argv[1]
KEY = sys.argv[2]
PWD = sys.argv[3]

SAT_LOGIN  = "https://wwwmat.sat.gob.mx/personas/iniciar-sesion"
SAT_NOTIFS = "https://wwwmat.sat.gob.mx/iniciar-expediente/mis-notificaciones/"
SAT_COMUNS = "https://wwwmat.sat.gob.mx/iniciar-expediente/mis-comunicados/"

def efirma_login(page):
    page.goto(SAT_LOGIN, timeout=60000); time.sleep(2)
    for sel in ["text=e.firma","a:has-text('e.firma')","button:has-text('e.firma')"]:
        try: page.click(sel, timeout=5000); time.sleep(1); break
        except Exception: pass
    try:
        ins = page.query_selector_all("input[type='file']")
        if ins: ins[0].set_input_files(CER); time.sleep(1)
        if len(ins)>=2: ins[1].set_input_files(KEY); time.sleep(1)
    except Exception as e: print(f"WARN:files: {e}")
    try: page.fill("input[type='password']", PWD)
    except Exception: pass
    for btn in ["button[type='submit']","button:has-text('Enviar')",
                "button:has-text('Ingresar')","input[type='submit']"]:
        try: page.click(btn, timeout=8000); time.sleep(4); break
        except Exception: pass
    if "iniciar-sesion" in page.url or "login" in page.url.lower():
        print("ERR:login fallido — verifica credenciales"); sys.exit(2)
    print("INFO:login OK")

def parse_table(page, tipo):
    items = []
    rows = page.query_selector_all("table tbody tr")
    for row in rows:
        cells = row.query_selector_all("td")
        if len(cells) < 2: continue
        texts = [c.inner_text().strip() for c in cells]
        link  = row.query_selector("a")
        href  = link.get_attribute("href") if link else ""
        nid   = href.split("/")[-1].split("?")[0] if href else ""
        items.append({"id": nid or str(abs(hash(texts[0]))),
                      "tipo": tipo,
                      "asunto": texts[1] if len(texts)>1 else texts[0],
                      "fecha":  texts[2] if len(texts)>2 else "",
                      "estado": texts[3] if len(texts)>3 else texts[-1],
                      "href":   href})
    return items

with sync_playwright() as pw:
    br  = pw.chromium.launch(headless=True)
    ctx = br.new_context(accept_downloads=True)
    pg  = ctx.new_page()
    try:
        efirma_login(pg)
        pg.goto(SAT_NOTIFS, timeout=60000); time.sleep(3)
        print("INFO:consultando notificaciones")
        notifs = parse_table(pg, "Notificación")
        pg.goto(SAT_COMUNS, timeout=60000); time.sleep(3)
        print("INFO:consultando comunicados")
        comuns = parse_table(pg, "Comunicado")
        all_items = notifs + comuns
        print(f"OK_LIST:{json.dumps(all_items, ensure_ascii=False)}")
    except SystemExit: raise
    except Exception as ex: print(f"ERR:{ex}")
    finally: br.close()
"""

_PLAYWRIGHT_PDF_SCRIPT = """
import sys, os, time
try:
    from playwright.sync_api import sync_playwright
except ImportError:
    print("ERR:playwright no disponible"); sys.exit(1)

CER      = sys.argv[1]
KEY      = sys.argv[2]
PWD      = sys.argv[3]
HREF     = sys.argv[4]
SAT_BASE = sys.argv[5]
OUT_PDF  = sys.argv[6]
SAT_LOGIN = "https://wwwmat.sat.gob.mx/personas/iniciar-sesion"

def efirma_login(page):
    page.goto(SAT_LOGIN, timeout=60000); time.sleep(2)
    for sel in ["text=e.firma","a:has-text('e.firma')","button:has-text('e.firma')"]:
        try: page.click(sel, timeout=5000); time.sleep(1); break
        except Exception: pass
    try:
        ins = page.query_selector_all("input[type='file']")
        if ins: ins[0].set_input_files(CER); time.sleep(1)
        if len(ins)>=2: ins[1].set_input_files(KEY); time.sleep(1)
    except Exception: pass
    try: page.fill("input[type='password']", PWD)
    except Exception: pass
    for btn in ["button[type='submit']","button:has-text('Enviar')",
                "button:has-text('Ingresar')","input[type='submit']"]:
        try: page.click(btn, timeout=8000); time.sleep(4); break
        except Exception: pass
    if "iniciar-sesion" in page.url or "login" in page.url.lower():
        print("ERR:login fallido"); sys.exit(2)

with sync_playwright() as pw:
    br  = pw.chromium.launch(headless=True)
    ctx = br.new_context(accept_downloads=True)
    pg  = ctx.new_page()
    try:
        efirma_login(pg)
        target = HREF if HREF.startswith("http") else (SAT_BASE + HREF)
        pg.goto(target, timeout=60000); time.sleep(3)
        pdf_saved = False
        for btn_sel in ["a:has-text('PDF')","a[href*='.pdf']",
                        "button:has-text('Descargar')","*:has-text('Ver / Acto')"]:
            try:
                with ctx.expect_download(timeout=20000) as dl_i:
                    pg.click(btn_sel, timeout=5000)
                dl_i.value.save_as(OUT_PDF)
                print(f"OK_PDF:{OUT_PDF}"); pdf_saved = True; break
            except Exception: pass
        if not pdf_saved:
            pg.pdf(path=OUT_PDF)
            print(f"OK_PDF:{OUT_PDF}")
    except SystemExit: raise
    except Exception as ex: print(f"ERR:{ex}")
    finally: br.close()
"""

def _run_playwright_list(cer_path, key_path, pwd):
    """Ejecuta el script de consulta. Devuelve (notifs_list, log_str)."""
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".py",
                                      mode="w", encoding="utf-8")
    tmp.write(_PLAYWRIGHT_LIST_SCRIPT)
    tmp.close()
    log_lines = []
    notifs = None
    try:
        proc = subprocess.Popen(
            [sys.executable, tmp.name, cer_path, key_path, pwd],
            stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
            text=True, encoding="utf-8", errors="replace")
        for line in proc.stdout:
            line = line.rstrip()
            if line.startswith("OK_LIST:"):
                try:
                    notifs = json.loads(line[8:])
                    log_lines.append(f"✅ {len(notifs)} item(s) recibidos")
                except Exception:
                    log_lines.append(f"⚠ parse error: {line}")
            elif line.startswith("ERR:"):
                log_lines.append(f"✖ {line[4:]}")
            elif line.startswith("INFO:"):
                log_lines.append(f"→ {line[5:]}")
            elif line.startswith("WARN:"):
                log_lines.append(f"⚠ {line[5:]}")
            elif line.strip():
                log_lines.append(line)
        proc.wait()
    finally:
        try:
            os.unlink(tmp.name)
        except Exception:
            pass
    return notifs, "\n".join(log_lines)

def _run_playwright_pdf(cer_path, key_path, pwd, href, sat_base, out_pdf):
    """Descarga el PDF de una notificación."""
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".py",
                                      mode="w", encoding="utf-8")
    tmp.write(_PLAYWRIGHT_PDF_SCRIPT)
    tmp.close()
    log_lines = []
    pdf_ok = False
    try:
        proc = subprocess.Popen(
            [sys.executable, tmp.name, cer_path, key_path, pwd,
             href, sat_base, out_pdf],
            stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
            text=True, encoding="utf-8", errors="replace")
        for line in proc.stdout:
            line = line.rstrip()
            if line.startswith("OK_PDF:"):
                log_lines.append(f"✅ PDF listo")
                pdf_ok = True
            elif line.startswith("ERR:"):
                log_lines.append(f"✖ {line[4:]}")
            elif line.startswith("INFO:"):
                log_lines.append(f"→ {line[5:]}")
            elif line.strip():
                log_lines.append(line)
        proc.wait()
    finally:
        try:
            os.unlink(tmp.name)
        except Exception:
            pass
    return pdf_ok, "\n".join(log_lines)

# ── Session state ─────────────────────────────────────────────────────────────
if "buzon_notifs"    not in st.session_state: st.session_state["buzon_notifs"]    = {}
if "buzon_log"       not in st.session_state: st.session_state["buzon_log"]       = {}
if "buzon_pdf_bytes" not in st.session_state: st.session_state["buzon_pdf_bytes"] = {}

# ── Layout ────────────────────────────────────────────────────────────────────
tab_consulta, tab_empresas = st.tabs(["📬 Consultar buzón", "🏢 Empresas"])

# ══════════════════════════════════════════════════════════════════════════════
# TAB 1: CONSULTAR
# ══════════════════════════════════════════════════════════════════════════════
with tab_consulta:
    empresas = _cargar_empresas()

    if not empresas:
        st.info("No hay empresas guardadas. Ve a la pestaña **🏢 Empresas** para agregar una.")
    else:
        # Resumen de notificaciones pendientes
        total_pend = 0
        for rfc_k, notifs_k in st.session_state["buzon_notifs"].items():
            total_pend += sum(1 for n in notifs_k
                              if "pendiente" in n.get("estado","").lower()
                              or "no leída"  in n.get("estado","").lower()
                              or "no leida"  in n.get("estado","").lower())
        if total_pend > 0:
            st.warning(f"⚠️  **{total_pend} notificación(es) pendiente(s)** en total")

        col_sel, col_btn = st.columns([3, 1])
        with col_sel:
            opciones = [f"{e.get('rfc','')}  —  {e.get('nombre','')}"
                        for e in empresas]
            sel_idx  = st.selectbox("Empresa", range(len(opciones)),
                                    format_func=lambda i: opciones[i])
        with col_btn:
            st.write("")
            st.write("")
            consultar_una = st.button("🔍 Consultar")

        col_todas, _ = st.columns([2, 3])
        with col_todas:
            consultar_todas = st.button("🔄 Consultar TODAS las empresas")

        emp_sel = empresas[sel_idx]
        rfc_sel = emp_sel.get("rfc", "")

        # ── Ejecutar consulta ───────────────────────────────────────────────
        def _consultar(emp):
            rfc = emp.get("rfc", "RFC")
            cer_b = _dec(emp["cer_enc"])
            key_b = _dec(emp["key_enc"])
            pwd_s = _dec(emp["pwd_enc"]).decode("utf-8", errors="replace")
            td = tempfile.mkdtemp()
            cer_f = os.path.join(td, f"{rfc}.cer")
            key_f = os.path.join(td, f"{rfc}.key")
            with open(cer_f, "wb") as f: f.write(cer_b)
            with open(key_f, "wb") as f: f.write(key_b)
            notifs, log = _run_playwright_list(cer_f, key_f, pwd_s)
            st.session_state["buzon_notifs"][rfc]  = notifs or []
            st.session_state["buzon_log"][rfc]     = log
            try:
                import shutil; shutil.rmtree(td)
            except Exception:
                pass

        if consultar_una:
            with st.spinner(f"Consultando buzón de {rfc_sel}…"):
                _consultar(emp_sel)
            st.rerun()

        if consultar_todas:
            for emp in empresas:
                with st.spinner(f"Consultando {emp.get('rfc','')}…"):
                    _consultar(emp)
            st.rerun()

        # ── Mostrar notificaciones ──────────────────────────────────────────
        notifs_rfc = st.session_state["buzon_notifs"].get(rfc_sel, None)

        if notifs_rfc is None:
            st.info("Haz clic en **🔍 Consultar** para ver las notificaciones.")
        elif not notifs_rfc:
            st.success("✅ Sin notificaciones ni comunicados pendientes.")
        else:
            pend = [n for n in notifs_rfc
                    if "pendiente" in n.get("estado","").lower()
                    or "no leída"  in n.get("estado","").lower()
                    or "no leida"  in n.get("estado","").lower()]
            otras = [n for n in notifs_rfc if n not in pend]

            if pend:
                st.error(f"🔴  **{len(pend)} pendiente(s)** — requieren atención")
                for n in pend:
                    with st.expander(f"🔴 [{n.get('tipo','')}]  {n.get('asunto','')}  —  {n.get('fecha','')}"):
                        st.write(f"**Estado:** {n.get('estado','')}")
                        _col1, _col2 = st.columns(2)
                        with _col1:
                            if st.button("⬇ Descargar PDF",
                                         key=f"dl_{n.get('id','')}"):
                                st.session_state["_buzon_dl_rfc"]   = rfc_sel
                                st.session_state["_buzon_dl_notif"] = n
                                st.rerun()

            if otras:
                st.info(f"📨  {len(otras)} comunicado(s) / notificación(es) leídos")
                for n in otras:
                    with st.expander(f"[{n.get('tipo','')}]  {n.get('asunto','')}  —  {n.get('fecha','')}"):
                        st.write(f"**Estado:** {n.get('estado','')}")
                        if st.button("⬇ Descargar PDF",
                                     key=f"dl_{n.get('id','')}"):
                            st.session_state["_buzon_dl_rfc"]   = rfc_sel
                            st.session_state["_buzon_dl_notif"] = n
                            st.rerun()

        # ── Procesar descarga PDF solicitada ────────────────────────────────
        if st.session_state.get("_buzon_dl_notif"):
            notif_dl = st.session_state.pop("_buzon_dl_notif")
            rfc_dl   = st.session_state.pop("_buzon_dl_rfc", rfc_sel)
            emp_dl   = next((e for e in empresas if e.get("rfc") == rfc_dl), None)
            if emp_dl:
                with st.spinner(f"Descargando PDF — {notif_dl.get('asunto','')}…"):
                    cer_b = _dec(emp_dl["cer_enc"])
                    key_b = _dec(emp_dl["key_enc"])
                    pwd_s = _dec(emp_dl["pwd_enc"]).decode("utf-8","replace")
                    td    = tempfile.mkdtemp()
                    cer_f = os.path.join(td, f"{rfc_dl}.cer")
                    key_f = os.path.join(td, f"{rfc_dl}.key")
                    out_f = os.path.join(td, "notif.pdf")
                    with open(cer_f,"wb") as f: f.write(cer_b)
                    with open(key_f,"wb") as f: f.write(key_b)
                    tipo  = notif_dl.get("tipo", "Notificación")
                    base  = ("https://wwwmat.sat.gob.mx/iniciar-expediente/mis-notificaciones/"
                             if tipo == "Notificación"
                             else "https://wwwmat.sat.gob.mx/iniciar-expediente/mis-comunicados/")
                    ok, log_pdf = _run_playwright_pdf(
                        cer_f, key_f, pwd_s,
                        notif_dl.get("href",""), base, out_f)
                    if ok and os.path.isfile(out_f):
                        with open(out_f,"rb") as f:
                            pdf_bytes = f.read()
                        safe = "".join(c for c in notif_dl.get("asunto","notif")
                                       if c.isalnum() or c in " _-")[:40]
                        fname = f"BUZON_{rfc_dl}_{safe}.pdf".replace(" ","_")
                        st.download_button(
                            label="📥 Descargar PDF",
                            data=pdf_bytes,
                            file_name=fname,
                            mime="application/pdf")
                    else:
                        st.error("No se pudo obtener el PDF.")
                        st.code(log_pdf)
                    try:
                        import shutil; shutil.rmtree(td)
                    except Exception:
                        pass

        # ── Log ─────────────────────────────────────────────────────────────
        log_rfc = st.session_state["buzon_log"].get(rfc_sel, "")
        if log_rfc:
            with st.expander("📋 Log de última consulta"):
                st.code(log_rfc)

# ══════════════════════════════════════════════════════════════════════════════
# TAB 2: EMPRESAS
# ══════════════════════════════════════════════════════════════════════════════
with tab_empresas:
    st.subheader("🏢 Empresas registradas")
    empresas_tab = _cargar_empresas()

    if not empresas_tab:
        st.info("No hay empresas registradas aún.")
    else:
        for emp in empresas_tab:
            col_r, col_n, col_x = st.columns([2, 4, 1])
            col_r.write(f"**{emp.get('rfc','')}**")
            sat_badge = "  `[↗SAT]`" if emp.get("_sat_docs") else ""
            col_n.markdown(emp.get("nombre","") + sat_badge)
            with col_x:
                if emp.get("_sat_docs"):
                    st.caption("↗SAT", help="Administra esta e.firma desde el módulo Constancia y Opinión SAT")
                else:
                    if st.button("🗑", key=f"del_{emp.get('rfc','')}",
                                 help="Eliminar empresa"):
                        _eliminar_empresa(emp.get("rfc",""))
                        st.rerun()

    st.divider()
    st.subheader("➕ Agregar / actualizar empresa")

    col_c, col_k = st.columns(2)
    with col_c:
        cer_file = st.file_uploader("Certificado .cer", type=["cer"],
                                     key="buzon_cer")
    with col_k:
        key_file = st.file_uploader("Llave privada .key", type=["key"],
                                     key="buzon_key")

    pwd_input = st.text_input("Contraseña de e.firma", type="password",
                               key="buzon_pwd")

    if st.button("💾 Guardar e.firma", type="primary"):
        if not cer_file or not key_file or not pwd_input:
            st.error("Completa .cer, .key y contraseña.")
        else:
            cer_bytes = cer_file.read()
            key_bytes = key_file.read()
            rfc_det   = os.path.splitext(cer_file.name)[0].upper()
            nombre_det = rfc_det
            try:
                from satcfdi.models import Signer
                sg = Signer.load(certificate=cer_bytes,
                                 key=key_bytes,
                                 password=pwd_input.encode())
                rfc_det    = sg.rfc
                nombre_det = getattr(sg, "legal_name", rfc_det)
            except Exception:
                pass
            _guardar_empresa(
                rfc_det, nombre_det,
                _enc(cer_bytes),
                _enc(key_bytes),
                _enc(pwd_input.encode("utf-8")))
            st.success(f"✅ E.firma de **{rfc_det}** guardada.")
            st.rerun()

    st.divider()
    st.caption(
        "⚠️ Las e.firmas se almacenan en Supabase cifradas en base64. "
        "Para mayor seguridad en producción, configura cifrado adicional en el servidor."
    )
