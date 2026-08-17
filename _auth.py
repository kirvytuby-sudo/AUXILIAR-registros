"""
_auth.py — Módulo de autenticación unificado para el Auxiliar de Registros.

Uso en cada página protegida:
    import _auth
    auth_user, auth_name, es_admin = _auth.require_login()

Si el usuario no está autenticado, muestra el formulario y llama st.stop().
Si no existen [sat_users] en Secrets, devuelve ("dev", "Desarrollador", True)
para no bloquear en entornos locales.

La sesión persiste entre páginas porque usa st.session_state con claves globales.
"""
import hashlib
import hmac
import secrets as _secrets_mod
import streamlit as st

# ── Claves de sesión ──────────────────────────────────────────────────────────
_KEY_USER = "sat_auth_user"
_KEY_NAME = "sat_auth_name"
_ADMIN    = "kirvy"


# ── Hash / verify ─────────────────────────────────────────────────────────────
def _pw_hash(password: str, salt: str) -> str:
    dk = hashlib.pbkdf2_hmac(
        "sha256", password.encode("utf-8"), salt.encode("utf-8"), 260_000
    )
    return dk.hex()


def _pw_verify(password: str, stored: str) -> bool:
    """stored = '<salt>:<hash_hex>'"""
    try:
        salt, expected = stored.split(":", 1)
        return hmac.compare_digest(_pw_hash(password, salt), expected)
    except Exception:
        return False


# ── Solicitudes pendientes (singleton entre sesiones) ─────────────────────────
@st.cache_resource
def _get_pendientes():
    return {"lista": []}


def _get_sat_users() -> dict | None:
    try:
        raw = st.secrets.get("sat_users")
        return dict(raw) if raw else None
    except Exception:
        return None


# ── Formulario de login ───────────────────────────────────────────────────────
def _mostrar_login(sat_users: dict, form_key: str) -> None:
    col_l, col_c, col_r = st.columns([1, 1.2, 1])
    with col_c:
        st.markdown("#### 🔐 Acceso al módulo")
        usr = st.text_input("Usuario", key=f"_auth_usr_{form_key}",
                            placeholder="tu usuario")
        pwd = st.text_input("Contraseña", type="password",
                            key=f"_auth_pwd_{form_key}",
                            placeholder="••••••••")
        if st.button("Entrar →", type="primary", use_container_width=True,
                     key=f"_auth_btn_{form_key}"):
            datos = sat_users.get(usr.strip().lower())
            if datos and _pw_verify(pwd, datos.get("password_hash", "")):
                st.session_state[_KEY_USER] = usr.strip().lower()
                st.session_state[_KEY_NAME] = datos.get("name", usr.upper())
                st.rerun()
            else:
                st.error("❌ Usuario o contraseña incorrectos.")

        st.markdown("---")
        with st.expander("📝 ¿No tienes cuenta? — Solicitar acceso"):
            with st.form(f"_auth_solicitud_{form_key}"):
                sol_nombre  = st.text_input("Nombre completo")
                sol_usuario = st.text_input("Usuario deseado",
                                            placeholder="sin espacios, minúsculas")
                sol_pwd1 = st.text_input("Contraseña", type="password")
                sol_pwd2 = st.text_input("Confirmar contraseña", type="password")
                submit = st.form_submit_button("📨 Enviar solicitud",
                                               use_container_width=True)
            if submit:
                u = sol_usuario.strip().lower().replace(" ", "_")
                if not all([sol_nombre.strip(), u, sol_pwd1]):
                    st.error("Completa todos los campos.")
                elif sol_pwd1 != sol_pwd2:
                    st.error("Las contraseñas no coinciden.")
                elif any(r["usuario"] == u for r in _get_pendientes()["lista"]):
                    st.warning("Ya hay una solicitud pendiente para ese usuario.")
                else:
                    s = _secrets_mod.token_hex(16)
                    _get_pendientes()["lista"].append({
                        "nombre":        sol_nombre.strip(),
                        "usuario":       u,
                        "password_hash": f"{s}:{_pw_hash(sol_pwd1, s)}",
                        "fecha":         __import__("datetime").datetime.now()
                                         .strftime("%d/%m/%Y %H:%M"),
                    })
                    st.success("✅ Solicitud enviada. El administrador la revisará.")


# ── Sidebar de usuario autenticado ────────────────────────────────────────────
def _sidebar_usuario(modulo: str) -> None:
    nombre = st.session_state.get(_KEY_NAME, "")
    with st.sidebar:
        st.markdown(f"👤 **{nombre}**")
        st.caption(modulo)
        if st.button("🚪 Cerrar sesión", key=f"_auth_logout_{modulo}"):
            st.session_state.pop(_KEY_USER, None)
            st.session_state.pop(_KEY_NAME, None)
            st.rerun()


# ── API pública ───────────────────────────────────────────────────────────────
def require_login(modulo: str = "SAT") -> tuple[str, str, bool]:
    """
    Verifica autenticación. Si no está logueado muestra el formulario y llama st.stop().

    Returns:
        (auth_user, auth_name, es_admin)
    """
    sat_users = _get_sat_users()

    # Sin configuración de usuarios → modo desarrollo, acceso libre
    if not sat_users:
        return "dev", "Desarrollador", True

    # Ya autenticado
    if st.session_state.get(_KEY_USER):
        auth_user = st.session_state[_KEY_USER]
        auth_name = st.session_state.get(_KEY_NAME, auth_user.upper())
        _sidebar_usuario(modulo)
        return auth_user, auth_name, (auth_user == _ADMIN)

    # No autenticado — mostrar formulario
    _mostrar_login(sat_users, form_key=modulo.lower().replace(" ", "_"))
    st.stop()
