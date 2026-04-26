# Versao Kivy offline

Foi criada uma versao Kivy separada em:

```text
main.py
```

Ela roda offline usando o mesmo `financeiro.db`.

## Testar no Windows

```powershell
python -m pip install -r requirements-kivy.txt
python main.py
```

## Gerar APK

Buildozer funciona melhor em Linux ou WSL. No Windows puro normalmente da erro.

No WSL/Linux, dentro da pasta do projeto:

```bash
pip install buildozer
buildozer android debug
```

O APK costuma aparecer em:

```text
bin/
```

## Observacao

Esta versao Kivy ja inclui login, cadastro, dashboard, adicionar/editar/excluir lancamentos, categorias, contas, planos e configuracoes basicas. Ela ainda nao substitui completamente toda a interface Flask/PWA, mas esta bem mais preparada para virar APK offline.
