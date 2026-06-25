"""
Auth Azure AD MINIMAL — adaptado do projeto velds-product-os do Bruno Fagundes.

Sem cookie manager, sem JS, sem meta refresh complexos. APENAS:
1. Tela de login com <a href> simples (mesma aba)
2. User vai pro Microsoft, autentica, volta pra /auth/callback?code=X&state=Y
3. App detecta code na URL, troca por token via POST, lê email do id_token
4. Salva user em st.session_state + assina vauth na URL pra persistir
5. Limpa query params + st.rerun
6. App reroda, lê session_state OU vauth da URL, libera

A sessão Streamlit é mantida via cookie automático do próprio Streamlit
(_streamlit_xsrf, etc) entre a ida pro Azure e a volta — desde que seja
a MESMA ABA (não nova aba).
"""
from __future__ import annotations

import base64
import hashlib
import hmac
import json
import secrets as _py_secrets
import time
import urllib.parse
from typing import Callable, Optional

import streamlit as st


# ── User token assinado (sobrevive na URL como ?vauth=...) ───────────────────

_VAUTH_PARAM = "vauth"
_VAUTH_TTL_HOURS = 12


def _sign_vauth_token(user: dict, secret: str) -> str:
    """JWT-like minimal: payload_b64.signature_b64"""
    exp = int(time.time()) + _VAUTH_TTL_HOURS * 3600
    payload = {**user, "exp": exp}
    payload_bytes = json.dumps(payload, separators=(",", ":")).encode("utf-8")
    payload_b64 = base64.urlsafe_b64encode(payload_bytes).decode("utf-8").rstrip("=")
    sig = hmac.new(secret.encode("utf-8"), payload_b64.encode("utf-8"),
                   hashlib.sha256).digest()
    sig_b64 = base64.urlsafe_b64encode(sig).decode("utf-8").rstrip("=")
    return f"{payload_b64}.{sig_b64}"


def _verify_vauth_token(token: str, secret: str) -> Optional[dict]:
    """Verifica HMAC e expiração. Retorna payload se válido."""
    try:
        payload_b64, sig_b64 = token.split(".")
        expected = hmac.new(secret.encode("utf-8"), payload_b64.encode("utf-8"),
                            hashlib.sha256).digest()
        expected_b64 = base64.urlsafe_b64encode(expected).decode("utf-8").rstrip("=")
        if not hmac.compare_digest(expected_b64, sig_b64):
            return None
        s = payload_b64 + "=" * (-len(payload_b64) % 4)
        payload = json.loads(base64.urlsafe_b64decode(s).decode("utf-8"))
        if payload.get("exp", 0) < time.time():
            return None
        return payload
    except Exception:
        return None


def _gen_pkce_pair() -> tuple[str, str]:
    """Gera (code_verifier, code_challenge) pra PKCE.
    code_verifier: random 64 chars base64url.
    code_challenge: base64url(SHA256(code_verifier)).
    Azure AD exige PKCE em apps configurados como SPA ou cross-origin."""
    verifier = _py_secrets.token_urlsafe(64)
    challenge = base64.urlsafe_b64encode(
        hashlib.sha256(verifier.encode("ascii")).digest()
    ).decode("ascii").rstrip("=")
    return verifier, challenge


def _read_auth_config() -> dict | None:
    """Lê o bloco [auth] dos secrets. Retorna None se ausente/incompleto."""
    try:
        cfg = dict(st.secrets.get("auth", {}) or {})
    except Exception:
        return None
    # client_secret é OPCIONAL — SPA/public clients não usam (só PKCE)
    required = ("client_id", "server_metadata_url", "redirect_uri")
    if not all(cfg.get(k) for k in required):
        return None
    return cfg


def _tenant_id_from_metadata_url(url: str) -> str:
    """Extrai tenant_id de https://login.microsoftonline.com/{tenant}/v2.0/..."""
    try:
        parts = url.split("/")
        idx = parts.index("login.microsoftonline.com")
        return parts[idx + 1]
    except Exception:
        import re as _re
        m = _re.search(r"login\.microsoftonline\.com/([^/]+)/", url)
        return m.group(1) if m else "common"


def _origin_from_redirect(redirect_uri: str) -> str:
    """Extrai 'https://host[:port]' de uma URL de redirect.
    Necessário no header Origin do POST pro token endpoint (validação SPA)."""
    p = urllib.parse.urlparse(redirect_uri)
    return f"{p.scheme}://{p.netloc}" if p.netloc else "https://localhost"


def _b64url_decode(s: str) -> bytes:
    s += "=" * (-len(s) % 4)
    return base64.urlsafe_b64decode(s.encode("utf-8"))


def _decode_id_token_claims(id_token: str) -> dict:
    """Decodifica payload do JWT (parte central). Sem validar assinatura
    porque recebemos via HTTPS direto do Azure no token endpoint."""
    try:
        parts = id_token.split(".")
        if len(parts) < 2:
            return {}
        return json.loads(_b64url_decode(parts[1]).decode("utf-8"))
    except Exception:
        return {}


def _build_authorize_url(cfg: dict, code_challenge: str, state: str) -> str:
    """Monta URL do endpoint /authorize do Azure AD com PKCE + state.
    state carrega o code_verifier (round-trip do OAuth) — sobrevive ao redirect
    sem precisar de cookie/session_state."""
    tenant = _tenant_id_from_metadata_url(cfg["server_metadata_url"])
    base = f"https://login.microsoftonline.com/{tenant}/oauth2/v2.0/authorize"
    params = {
        "client_id":             cfg["client_id"],
        "response_type":         "code",
        "redirect_uri":          cfg["redirect_uri"],
        "response_mode":         "query",
        "scope":                 "openid profile email User.Read",
        "prompt":                "select_account",
        "code_challenge":        code_challenge,
        "code_challenge_method": "S256",
        "state":                 state,
    }
    return f"{base}?{urllib.parse.urlencode(params)}"


def _exchange_code_for_token(cfg: dict, code: str, code_verifier: str) -> dict:
    """POST pro token endpoint do Azure pra trocar code por tokens (PKCE).
    Público client (SPA): NÃO envia client_secret — só PKCE como prova."""
    import requests as _req
    tenant = _tenant_id_from_metadata_url(cfg["server_metadata_url"])
    token_url = f"https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token"
    data = {
        "client_id":     cfg["client_id"],
        "code":          code,
        "redirect_uri":  cfg["redirect_uri"],
        "grant_type":    "authorization_code",
        "scope":         "openid profile email User.Read",
        "code_verifier": code_verifier,
    }
    # IMPORTANTE: NÃO incluir client_secret (Azure rejeita com AADSTS700025
    # se o app for registrado como SPA/public client).
    try:
        # Origin header dinâmico baseado no redirect_uri configurado.
        # Necessário pra Azure validar o SPA (CORS).
        origin = _origin_from_redirect(cfg["redirect_uri"])
        r = _req.post(token_url, data=data, timeout=15,
                      headers={"Origin": origin})
        return r.json() if r.content else {"error": f"empty_{r.status_code}"}
    except Exception as ex:
        return {"error": str(ex)}


def _read_pkce_cookie() -> str:
    """Lê o cookie velds_pkce_verifier (gravado via JS no clique do link)."""
    try:
        v = st.context.cookies.get("velds_pkce_verifier")
        if v:
            return v
    except Exception:
        pass
    try:
        from streamlit.web.server.websocket_headers import _get_websocket_headers
        headers = _get_websocket_headers() or {}
        cookie_header = headers.get("Cookie", "") or headers.get("cookie", "")
        for part in cookie_header.split(";"):
            k, _, val = part.strip().partition("=")
            if k == "velds_pkce_verifier" and val.strip():
                return val.strip()
    except Exception:
        pass
    return ""


def _email_domain_ok(email: str, domain_required) -> bool:
    """Aceita 1+ domínios. domain_required pode ser str ("velds.com.br"),
    str com vírgulas ("a.com,b.com") ou lista/tupla. Email já vem lowercased."""
    if isinstance(domain_required, str):
        domains = [d.strip() for d in domain_required.split(",") if d.strip()]
    else:
        domains = [str(d).strip() for d in domain_required if str(d).strip()]
    return any(email.endswith("@" + d.lower()) for d in domains)


def handle_auth_flow(
    *,
    domain_required,
    render_login: Callable,
    render_acesso_negado: Callable,
) -> tuple[bool, Optional[dict]]:
    """Fluxo de auth com PKCE. code_verifier guardado em COOKIE (gravado via
    JS no clique do link) — sobrevive ao redirect pro Azure e callback."""

    cfg = _read_auth_config()
    if not cfg:
        return False, None

    cookie_secret = cfg.get("cookie_secret", "")
    qp = st.query_params

    # ── 1) Já logado via ?vauth= na URL? (persiste através de reloads) ──
    vauth = qp.get(_VAUTH_PARAM, "")
    if vauth and cookie_secret:
        payload = _verify_vauth_token(vauth, cookie_secret)
        if payload:
            email = (payload.get("email") or "").lower().strip()
            if not _email_domain_ok(email, domain_required):
                render_acesso_negado(email or "(sem email)")
                st.stop()
            return True, {"email": email, "name": payload.get("name", "")}

    # ── 1b) Fallback: já logado via session_state (mesma sessão) ────────
    user_info = st.session_state.get("_velds_user")
    if user_info:
        email = (user_info.get("email") or "").lower().strip()
        if not _email_domain_ok(email, domain_required):
            render_acesso_negado(email or "(sem email)")
            st.stop()
        return True, user_info

    # ── 2) Voltando do callback Azure? ──────────────────────────────────
    code = qp.get("code", "")
    if code:
        # code_verifier vem no parâmetro state (round-trip do OAuth) -
        # NÃO depende de cookie nem session_state (que não sobreviveram
        # ao redirect no Streamlit Cloud)
        code_verifier = qp.get("state", "")
        if not code_verifier:
            # Fallback pra session_state/cookie (caso o state venha vazio)
            code_verifier = (
                _read_pkce_cookie()
                or st.session_state.get("_velds_pkce_verifier", "")
            )
        if not code_verifier:
            st.error(
                "❌ code_verifier perdido (state vazio). "
                "Limpe cookies e tente de novo."
            )
            for k in ("code", "state", "session_state"):
                if k in qp:
                    del qp[k]
            st.stop()
        with st.spinner("Validando login Microsoft..."):
            token_resp = _exchange_code_for_token(cfg, code, code_verifier)
        if "error" in token_resp:
            st.error(
                f"❌ Erro do Azure: `{token_resp.get('error')}`\n\n"
                f"`{token_resp.get('error_description', '')}`"
            )
            # Limpa URL pra não ficar em loop
            for k in ("code", "state", "session_state"):
                if k in qp:
                    del qp[k]
            st.stop()
        id_token = token_resp.get("id_token", "")
        claims = _decode_id_token_claims(id_token)
        email = (claims.get("email") or claims.get("preferred_username") or "").lower().strip()
        name = claims.get("name") or claims.get("given_name") or (email.split("@")[0] if email else "Usuário")
        if not email:
            st.error("❌ Token recebido mas sem email no payload.")
            st.stop()
        # Gera vauth token assinado e REDIRECIONA pra /?vauth=TOKEN
        # (não usa st.rerun porque session_state não persiste no Cloud).
        user_data = {"email": email, "name": name}
        st.session_state["_velds_user"] = user_data
        if cookie_secret:
            token = _sign_vauth_token(user_data, cookie_secret)
            # meta refresh com timeout 0 = redirect imediato pra URL com vauth
            redirect_url = f"/?{_VAUTH_PARAM}={urllib.parse.quote(token)}"
            st.markdown(
                f'<meta http-equiv="refresh" content="0; url={redirect_url}">'
                f'<div style="padding:40px;text-align:center;color:var(--text-md);">'
                f'<div style="font-size:1.2rem;margin-bottom:12px;">✓ Login OK, {name}</div>'
                f'<div>Carregando o CloudIA PCM...</div>'
                f'</div>',
                unsafe_allow_html=True,
            )
            st.stop()
        # Fallback se não tem cookie_secret (não deveria acontecer)
        for k in ("code", "state", "session_state"):
            if k in qp:
                del qp[k]
        st.rerun()

    # ── 3) Não logado: gera PKCE NOVO a cada render + mostra tela ──────
    # code_verifier viaja no parâmetro state do OAuth (round-trip):
    # /authorize?state=verifier → callback ?state=verifier
    # Sobrevive ao redirect sem precisar de cookie/session.
    verifier, challenge = _gen_pkce_pair()
    authorize_url = _build_authorize_url(cfg, challenge, state=verifier)
    try:
        render_login(authorize_url=authorize_url)
    except TypeError:
        render_login()
    st.stop()


def logout() -> None:
    """Limpa session state e remove o vauth da URL → força re-login."""
    for k in list(st.session_state.keys()):
        if k.startswith("_velds_"):
            del st.session_state[k]
    # Redireciona pra raiz limpa (sem vauth)
    st.markdown(
        '<meta http-equiv="refresh" content="0; url=/">',
        unsafe_allow_html=True,
    )
    st.stop()
