# CloudIA PCM — Streamlit Cloud (versão online read-only)

Wrapper Python (Streamlit) que serve o dashboard CloudIA da FiberX online, com autenticação Azure AD restrita a contas `@fiberx.com.br`.

**Estrutura:**

```
_deploy_streamlit/
├── app.py                  # Wrapper Streamlit (auth + injeta dados.json inline)
├── index.html              # CloudIA V3 (Opção C2)
├── dados.json              # Snapshot dos dados (regerado pelo gerar_dados.py)
├── requirements.txt        # streamlit + requests + PyJWT + cryptography
├── .gitignore
├── modules/
│   ├── __init__.py
│   └── velds_auth.py       # OAuth 2.0 + PKCE (cópia do velds-product-os)
└── assets/                 # Logos FiberX/Velds
```

## Deploy

### 0. Pré-requisitos

- Conta GitHub configurada (gudibrambila-wq)
- Repo criado: https://github.com/gudibrambila-wq/cloudia-pcm-fiberx
- Client ID do Azure AD obtido com o TI (Gláucio)
- `velds_auth.py` em mãos (do Bruno Fagundes)
- Conta no [share.streamlit.io](https://share.streamlit.io) linkada com o GitHub

### 1. Sobe pro GitHub

A partir da pasta `_deploy_streamlit/`:

```bash
cd "_deploy_streamlit"
git init
git remote add origin https://github.com/gudibrambila-wq/cloudia-pcm-fiberx.git
git add .
git commit -m "Initial commit: CloudIA PCM V3 com auth Azure AD"
git branch -M main
git push -u origin main
```

### 2. Deploy no Streamlit Cloud

1. Acesse [share.streamlit.io](https://share.streamlit.io) → "New app"
2. Repository: `gudibrambila-wq/cloudia-pcm-fiberx`
3. Branch: `main`
4. Main file path: `app.py`
5. Custom subdomain: `pcm-fiberx` → URL final: `https://pcm-fiberx.streamlit.app/`
6. Deploy. Aguarde ~2 min.

### 3. Secrets

Em `share.streamlit.io → app → Settings → Secrets`, cole:

```toml
[auth]
client_id           = "GUID_QUE_GLAUCIO_FORNECEU"
server_metadata_url = "https://login.microsoftonline.com/6a3b0155-0b3f-4492-8e83-6afdc2e35539/v2.0/.well-known/openid-configuration"
redirect_uri        = "https://pcm-fiberx.streamlit.app/"
cookie_secret       = "GERE_STRING_RANDOM_LONGA"
```

Gerar `cookie_secret`:
```bash
python -c "import secrets; print(secrets.token_urlsafe(48))"
```

Salve → Streamlit reinicia sozinho.

### 4. Testar

Em aba anônima: https://pcm-fiberx.streamlit.app

- Conta `@fiberx.com.br` → entra
- Conta externa (gmail) → "Acesso negado"

## Como atualizar `dados.json`

Toda vez que rodar `gerar_dados.py` em `CloudIA_V3/` localmente, o `dados.json` do `v3/` é atualizado. Pra refletir no app online, use o script `sync_deploy_streamlit.bat` na pasta raiz do CloudIA — ele copia, commita e faz push em uma só ação.

O Streamlit Cloud detecta o push e rebuilda em ~30s.

## Limitações da versão online

- **Read-only**: edições (PV, forecast, Kanban) não persistem entre sessões. Pra editar, use a versão local com `server.py`.
- **Dados estáticos**: `dados.json` é o snapshot do último push. Pra ver dados atualizados, fazer novo deploy.
- **Sem POST endpoints**: o Streamlit Cloud não suporta o fluxo `/api/save-overrides` do server.py local.
