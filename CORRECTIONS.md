# Correção v4.1 — GitHub Actions

## Falha corrigida

O workflow executava `npm test`, que descobria `tests/pptx.test.cjs`. Esse teste importa `adm-zip`, mas o pacote não estava declarado no `package.json`, causando `MODULE_NOT_FOUND`.

Foi adicionado `adm-zip` em `devDependencies` (`^0.5.16`). O workflow já executa `npm install`, portanto a dependência será instalada antes dos testes.

## O que os logs anteriores confirmaram

- `npm run typecheck`: passou.
- `npx expo install --check`: passou.
- 8 testes offline: passaram.
- O teste `tests/pptx.test.cjs` parou somente pela dependência ausente.
- A compilação Android ainda não foi iniciada nessa execução.

Esta alteração corrige a causa identificada no log. A compilação nativa precisa ser confirmada em uma nova execução do GitHub Actions.
