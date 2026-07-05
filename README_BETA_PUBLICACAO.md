# Dig.ai Finanças Beta Web

Esta pasta e uma versao separada do site para publicar uma beta publica sem mexer na versao principal.

Identidade atual desta beta:

- produto: `Dig.ai Finanças`
- subtitulo: `Controle financeiro pessoal inteligente`
- empresa: `Zattara Soluções Inteligentes`
- suporte: `digiai.oficial@gmail.com`

## Objetivo da beta

- manter somente o plano gratuito
- publicar sem cobranca
- salvar dados online no servidor
- exibir politica de privacidade e termos de uso
- usar e-mail institucional de suporte

## O que foi simplificado

- a beta funciona apenas no plano `free`
- temas liberados: `dark`, `light` e `ocean`
- a pagina de planos virou uma apresentacao da beta
- recursos Pro ficam reservados para a versao completa

## Como os dados ficam salvos online

Esta beta agora trabalha com dois modos:

- `SQLite` local para desenvolvimento
- `PostgreSQL` para deploy via `DATABASE_URL`
- `Backend API` compartilhado com app e futuras versões, quando `NOTAFACIL_BACKEND_URL` estiver configurada

Se `DATABASE_URL` estiver configurada, o app usa Postgres automaticamente.
Se `NOTAFACIL_BACKEND_URL` estiver configurada, o site passa a usar a mesma conta, o mesmo plano e os mesmos dados do backend oficial.

## E-mail institucional

O sistema ja esta preparado para mostrar o e-mail de suporte em:

- rodape
- politica de privacidade
- termos de uso
- perfil/configuracoes

Variavel:

- `NOTAFACIL_SUPPORT_EMAIL`

Email atual configurado para a beta:

- `digiai.oficial@gmail.com`

Importante: a criacao real da caixa de e-mail precisa ser feita no seu provedor de dominio/e-mail. O codigo apenas vincula e exibe esse contato.

## Execucao local

```powershell
$env:NOTAFACIL_BACKEND_URL="http://127.0.0.1:8000"
cd "C:\Users\ferre\OneDrive\Desktop\PROJETOS\CONTROLE DE DESPESA\NOTE NOVO\FINANCEIRO_BETA_WEB"
python app_pwa.py
```

Abrir:

- `http://127.0.0.1:5000`

Para rodar no modo integrado com a API, suba o backend antes:

```powershell
cd "C:\Users\ferre\OneDrive\Desktop\PROJETOS\CONTROLE DE DESPESA\NOTE NOVO\NOTAFACIL_BACKEND"
.\.venv\Scripts\Activate.ps1
python -m uvicorn app.main:app --host 0.0.0.0 --port 8000 --reload
```

## Rotas novas da beta

- `/politica-de-privacidade`
- `/termos-de-uso`
- `/health`

## Publicacao

Arquivos incluidos:

- `.env.example`
- `Procfile`
- `render.yaml`
- `db_compat.py`

## Observacao importante

Para a beta gratuita no Render, o caminho recomendado agora e usar Postgres gratuito em vez de SQLite local da hospedagem.

Para a versao final/completa, o ideal continua sendo:

- backend dedicado
- PostgreSQL mais robusto
- sincronizacao app/site/programa
- e-mail real de recuperacao
- plano Pro e pagamento
