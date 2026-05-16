# Deploy gratuito da beta no Render

Esta beta agora esta preparada para um caminho gratuito mais realista:

- Web Service Flask no Render Free
- banco Postgres Free do Render
- sem depender de SQLite no disco local da hospedagem

## Identidade atual

- produto: `DigiAI Finance`
- subtitulo: `Controle financeiro pessoal inteligente`
- empresa: `DigiAI`
- suporte: `digiai.oficial@gmail.com`

## O que mudou tecnicamente

Antes, a beta dependia de SQLite local. Isso nao funciona bem em deploy gratuito porque o filesystem das web services gratuitas e efemero.

Agora, a beta aceita:

- SQLite local para desenvolvimento
- PostgreSQL via `DATABASE_URL` para deploy

Se a variavel `DATABASE_URL` estiver preenchida, o app usa Postgres automaticamente.
Se a variavel `NOTAFACIL_BACKEND_URL` estiver preenchida, a web passa a usar o backend oficial como fonte de verdade para:

- login
- cadastro
- plano Free/Pro
- contas
- categorias
- lancamentos
- configuracoes

## Arquivos prontos

- `render.yaml`
- `requirements.txt`
- `Procfile`
- `.env.example`
- `db_compat.py`

## O que o Render oferece na faixa gratuita

Segundo a documentacao oficial do Render:

- web services gratuitas existem
- Postgres gratuito existe
- o web service gratuito entra em idle sem trafego por 15 minutos
- bancos Postgres gratuitos expiram apos 30 dias se nao forem pagos
- bancos gratuitos nao possuem backup nativo

Referencias oficiais:

- [Deploy for Free](https://render.com/docs/free)
- [Create and Connect to Render Postgres](https://render.com/docs/databases)

## Como publicar

### 1. Suba esta pasta para um repositório Git

Pasta:

`C:\Users\ferre\OneDrive\Desktop\PROJETOS\CONTROLE DE DESPESA\NOTE NOVO\FINANCEIRO_BETA_WEB`

### 2. No Render, conecte o repositório

Voce pode usar Blueprint para o Render ler o `render.yaml`.

### 3. O blueprint cria:

- `digiai-finance-beta` como web service gratuito
- `digiai-finance-beta-db` como banco Postgres gratuito

### 4. Variáveis importantes

Ja definidas no blueprint:

- `FINANCEIRO_SECRET_KEY`
- `FLASK_DEBUG=false`
- `NOTAFACIL_COOKIE_SECURE=true`
- `NOTAFACIL_URL_SCHEME=https`
- `NOTAFACIL_APP_NAME=DigiAI Finance`
- `NOTAFACIL_APP_TAGLINE=Controle financeiro pessoal inteligente`
- `NOTAFACIL_COMPANY_NAME=DigiAI`
- `NOTAFACIL_SUPPORT_EMAIL=digiai.oficial@gmail.com`
- `DATABASE_URL` vindo do banco Postgres criado no Render

Depois que o backend estiver online, adicione manualmente no web service:

- `NOTAFACIL_BACKEND_URL=https://SEU-BACKEND.onrender.com`

### 5. Teste apos deploy

Abra estas rotas:

- `/login`
- `/cadastro`
- `/politica-de-privacidade`
- `/termos-de-uso`
- `/health`

Depois teste o fluxo:

1. criar conta
2. entrar
3. criar lancamento
4. sair
5. entrar de novo
6. confirmar se os dados continuam salvos
7. abrir o app mobile com a mesma conta e confirmar se os dados batem

## Limitacoes do plano gratuito

- o site pode demorar para acordar depois de ficar parado
- o banco gratuito expira apos 30 dias no Render Free
- nao ha backup nativo do banco gratuito

## Quando migrar para a versao mais forte

Quando a beta comecar a funcionar com usuarios reais, o ideal e evoluir para:

- Postgres pago ou provedor dedicado
- rotina de backup
- e-mail transacional
- app e exe publicados
- sincronizacao completa entre plataformas
