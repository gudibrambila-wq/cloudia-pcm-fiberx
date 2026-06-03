"""
CloudIA PCM — wrapper Streamlit pra hospedar o dashboard HTML/JS no Streamlit Cloud.

Versão online (read-only) do CloudIA: injeta o `dados.json` inline no `index.html`
da V3 e renderiza num componente Streamlit em tela cheia. Edições continuam
sendo feitas só no PC principal — esta versão é só leitura (TV/consulta).

Autenticação: Microsoft Azure AD (FiberX), via `modules/velds_auth.py` —
restringe acesso a contas @fiberx.com.br.

Como subir:
  1. Garanta que velds_auth.py está em modules/
  2. Configure secrets no Streamlit Cloud (Settings → Secrets):
       [auth]
       client_id           = "GUID_DO_AZURE"
       server_metadata_url = "https://login.microsoftonline.com/<TENANT>/v2.0/.well-known/openid-configuration"
       redirect_uri        = "https://pcm-fiberx.streamlit.app/"
       cookie_secret       = "string_random_longa"
  3. Push pro repo → Streamlit Cloud rebuilda sozinho
"""
from pathlib import Path
import streamlit as st

# ── Page config (tem que vir ANTES de qualquer st.* de render) ─────────────
st.set_page_config(
    page_title="CloudIA PCM · FiberX",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# Esconde chrome do Streamlit (header/footer/menu) e remove padding
st.markdown(
    """
    <style>
      #MainMenu, footer, header { visibility: hidden; }
      .stApp > header { display: none; }
      .block-container { padding: 0 !important; max-width: 100% !important; }
      iframe { border: none !important; }
    </style>
    """,
    unsafe_allow_html=True,
)

# ── Autenticação Azure AD (FiberX) ─────────────────────────────────────────
from modules import velds_auth as _vauth

_AUTH_DOMAIN_PERMITIDO = "fiberx.com.br"


def _render_login(authorize_url: str = "") -> None:
    st.markdown(
        f"""
        <div style="text-align:center;padding:80px 20px;">
          <h1 style="color:#0078D4;margin-bottom:10px;">CloudIA PCM · FiberX</h1>
          <p style="color:#64748b;font-size:16px;margin-bottom:40px;">
            Entre com sua conta <b>@{_AUTH_DOMAIN_PERMITIDO}</b>
          </p>
          <a href="{authorize_url}" style="display:inline-block;padding:14px 32px;
             background:#0078D4;color:white;text-decoration:none;border-radius:6px;
             font-weight:600;font-size:15px;">
            🔐 Entrar com Microsoft
          </a>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.stop()


def _render_acesso_negado(email: str) -> None:
    st.error(f"❌ Acesso negado para `{email}`. Use uma conta @{_AUTH_DOMAIN_PERMITIDO}.")
    if st.button("Sair e tentar com outra conta"):
        _vauth.logout()
    st.stop()


_auth_ok, _user = _vauth.handle_auth_flow(
    domain_required=_AUTH_DOMAIN_PERMITIDO,
    render_login=_render_login,
    render_acesso_negado=_render_acesso_negado,
)

# Trava de segurança: se os secrets do Azure AD ainda não foram configurados,
# handle_auth_flow retorna (False, None) sem mostrar tela de login. Sem isso,
# o app abriria pra qualquer um que tivesse o link. Bloqueia até o admin
# configurar o [auth] no Streamlit Cloud (Settings → Secrets).
if not _auth_ok:
    st.markdown(
        """
        <div style="text-align:center;padding:80px 20px;">
          <h2 style="color:#dc2626;">🔧 Configuração incompleta</h2>
          <p style="color:#64748b;font-size:14px;margin-top:16px;">
            O CloudIA PCM está sendo configurado. Aguarde o admin terminar o setup
            do Azure AD e tente novamente em alguns minutos.
          </p>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.stop()

# ── Lê arquivos e injeta dados inline ──────────────────────────────────────
BASE = Path(__file__).parent
html = (BASE / "index.html").read_text(encoding="utf-8")
dados = (BASE / "dados.json").read_text(encoding="utf-8")

# Injeta dados como variável global ANTES do fetch original ser disparado.
# O index.html prefere window.__DADOS_INLINE__ quando existir (modo online).
html_injected = html.replace(
    "</head>",
    f"<script>window.__DADOS_INLINE__ = {dados};</script>\n</head>",
    1,
)

# ── Renderiza o app inteiro num iframe full-height ────────────────────────
st.components.v1.html(html_injected, height=1400, scrolling=True)
