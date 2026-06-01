# CloudIA PCM — Setup Azure Static Web Apps

Solicitante: Douglas Brambila (comercial@fiberx.com.br)
Objetivo: hospedar dashboard interno restrito aos colaboradores @fiberx.com.br

## O que é

Site estático (HTML/JS) com dashboard de PCM. Já está em repositório GitHub.
Precisamos publicá-lo como **Azure Static Web App** com login Microsoft restrito ao tenant da FiberX.

## Passos (para o TI / admin Azure)

### 1) Criar a Static Web App no Azure portal

1. Acessar [portal.azure.com](https://portal.azure.com) com conta admin
2. Procurar **"Static Web Apps"** → **+ Create**
3. Preencher:
   - **Subscription**: (a usada pela FiberX)
   - **Resource Group**: criar novo `rg-cloudia-pcm` (ou usar existente)
   - **Name**: `cloudia-pcm-fiberx`
   - **Plan type**: **Free**
   - **Region**: East US 2 (ou mais próximo)
   - **Source**: GitHub
   - **Sign in with GitHub** → autorizar a conta `gudibrambila-wq` (ou organização FiberX se houver)
   - **Organization**: gudibrambila-wq
   - **Repository**: `cloudia-pcm-fiberx` (Douglas vai criar)
   - **Branch**: `main`
   - **Build Presets**: Custom
   - **App location**: `/`
   - **Output location**: (deixar em branco)
4. **Review + create** → aguardar deploy

### 2) Registrar o app no Azure AD (Entra ID) para auth

1. Azure portal → **Microsoft Entra ID** → **App registrations** → **+ New registration**
2. Preencher:
   - **Name**: `CloudIA PCM`
   - **Supported account types**: **Accounts in this organizational directory only (FiberX only - Single tenant)**
   - **Redirect URI**: Web → `https://<NOME-DA-STATIC-WEB-APP>.azurestaticapps.net/.auth/login/aad/callback`
3. Após criar, anotar:
   - **Application (client) ID** → guardar
   - **Directory (tenant) ID** → guardar
4. Em **Certificates & secrets** → **+ New client secret** → criar (validade 24 meses):
   - Anotar o **Value** do secret (não vai mostrar depois)

### 3) Vincular auth ao Static Web App

1. Voltar pra **Static Web App** criada → **Configuration** → **Application settings**
2. Adicionar 2 variáveis:
   - `AAD_CLIENT_ID` = (Application client ID anotado no passo 2)
   - `AAD_CLIENT_SECRET` = (Value do secret anotado no passo 2)
3. **Save**

### 4) Atualizar `staticwebapp.config.json`

O arquivo já está no repo com placeholder. Substituir `COLOCAR_TENANT_ID_AQUI` pelo Directory (tenant) ID anotado.

Linha exata pra trocar:
```json
"openIdIssuer": "https://login.microsoftonline.com/COLOCAR_TENANT_ID_AQUI/v2.0"
```

→ commitar no GitHub. O Static Web App vai re-deployar automaticamente em 1-2 min.

### 5) Testar

1. Acessar `https://<NOME-DA-STATIC-WEB-APP>.azurestaticapps.net`
2. Deve redirecionar pra login Microsoft
3. Entrar com conta `@fiberx.com.br` → deve abrir o CloudIA
4. Entrar com conta de outro domínio → deve bloquear (porque tenant é Single)

### 6) (Opcional) Configurar domínio custom

Se quiser URL bonita tipo `cloudia.fiberx.com.br`:
1. Static Web App → **Custom domains** → **+ Add**
2. Adicionar `cloudia.fiberx.com.br`
3. Configurar CNAME no DNS da FiberX apontando pra `<NOME>.azurestaticapps.net`

## Custos

Plano **Free** do Azure Static Web Apps:
- 100 GB de banda/mês
- 0.5 GB de storage
- Custom domain gratuito
- SSL gratuito
- Auth gratuito

Mais que suficiente pra dashboard interno.

## Após setup pronto

Devolver pra Douglas:
- URL do site (`https://<NOME>.azurestaticapps.net` ou domínio custom)
- Confirmação de que login restringe ao tenant FiberX

Douglas vai atualizar o conteúdo fazendo push no GitHub — Static Web App re-deploya automaticamente.
