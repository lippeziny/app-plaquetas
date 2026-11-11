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


## v4.2 — correção da importação Excel no Android

- Substituída a leitura Base64 de planilhas por `ExpoFile.bytes()` e `XLSX.read(..., { type: 'array', dense: true })`.
- O parser deixou de transformar todas as abas em arrays ao mesmo tempo; agora converte somente a aba solicitada na prévia/importação.
- Incluído teste comparando a leitura por bytes com a leitura Base64 usada no teste legado.
- Incluída verificação para impedir regressão que volte a enviar a planilha como Base64 pela ponte nativa.
- Incluído limite preventivo de 50 MB por arquivo e 250 mil linhas na aba selecionada, com orientação de otimização em vez de deixar o Android encerrar o app por OOM.


## v4.2.1 — correção da importação de modelo PowerPoint

O fluxo de importação do modelo personalizado ainda chamava a função removida `readPickedFileBase64`. A chamada foi alinhada ao leitor binário atual: `bytesToBase64(await readPickedFileBytes(asset))`. Foi incluída uma asserção de regressão para exigir esse caminho e impedir referências à função antiga.


## v4.3 — importador incremental otimizado

- Substituída a importação em workbook inteiro por extração seletiva dos XMLs de XLSX e parser SAX em blocos de 256 KB.
- Mantidas apenas as colunas configuradas; padrão A (código), B (descrição) e D (quantidade).
- Strings compartilhadas são varridas em fluxo e persistidas no staging somente quando referenciadas pelos campos selecionados.
- Banco SQLite recebe lotes e usa índices de referência para reduzir o custo da procura de shared strings.
- A base anterior permanece intacta até a substituição transacional no fim; cancelamento/falha limpa apenas staging.
- Formatos numéricos simples como 000000 são aplicados para preservar zeros à esquerda em códigos.
- Decodificador UTF-8 em fluxo suporta sequências divididas entre blocos e não depende de TextDecoder global.
- Barra de progresso agora mostra etapa, percentual, linhas, tempo decorrido, média de linhas/s e ETA.
- Removido o limite rígido de 50 MB. O limite efetivo depende do armazenamento livre, tamanho descompactado e recursos do aparelho.
- Necessário rebuild nativo após instalar react-native-zip-archive (Android e iOS; mínimo iOS 15.5).


### v4.3 — otimização adicional da leitura de strings e progresso

- A relação de IDs compartilhados necessários é obtida com `getEachAsync()` no SQLite, sem materializar o resultado completo com `getAllAsync()`.
- IDs ordenados são armazenados em `Uint32Array` (4 bytes por ID) e o SAX parser captura apenas as strings referenciadas; a leitura antecipadamente assim que o último ID é localizado.
- Inserções de shared strings deixaram de executar uma verificação `EXISTS` contra todas as linhas a cada lote; o filtro acontece antes, no parser.
- Linhas completamente vazias nas colunas selecionadas não são armazenadas em staging. Os lotes de gravação da planilha foram ampliados para até 1.000 linhas e os lotes da etapa de consolidação para 2.000 linhas, reduzindo transações frequentes sem aumentar significativamente o pico de memória.
- Tempo decorrido e ETA são atualizados a cada segundo, inclusive durante etapas longas sem callback de progresso. O cursor de IDs também verifica o cancelamento periodicamente.
- O mapeamento inicial segue A = código, B = nome e D = quantidade; o usuário ainda pode ajustar as colunas na tela.


## v4.3.1 — Expo FileSystem SDK 54 API compatibility

- Removed the unsupported `FileMode` import from `expo-file-system`.
- Changed `File.open(FileMode.ReadOnly)` to `File.open()` as required by the project’s pinned Expo FileSystem typings.
- Kept incremental 256 KB `readBytes` reads and the UTF-8 streaming decoder unchanged.
- Added a regression assertion to the offline test suite so the unsupported call is not reintroduced.


## v4.4.0 — pesquisa e seleção direta do catálogo

- Adicionada busca de sugestões por prefixo para nome e código, com preenchimento do valor completo ao tocar em uma sugestão.
- A busca usa os índices `name_norm` e `code_norm` do SQLite e normalização sem acentos.
- Adicionado catálogo diretamente na etapa Itens: exibe nome e código com caixa de seleção, filtro por trecho do nome e paginação de 60 registros.
- A seleção de produtos permanece entre alterações do filtro e pode ser adicionada em lote às plaquetas; ao importar uma nova base, as seleções antigas são limpas.
- Ajustado o recorte das imagens na prévia conforme os metadados `a:srcRect` / `a:fillRect` do PowerPoint. O logotipo CKTR usa `cover` para respeitar o recorte horizontal original; imagens sem recorte continuam em `contain`.
- Incluídos testes de regressão para busca por nome/código, catálogo paginado, seleção por checkbox e modo de encaixe da imagem.

## v4.5.0 — detecção semântica de estoque, persistência e busca de produtos

- A aba padrão é escolhida automaticamente pelo nome: primeiro correspondência exata com "Estoque", depois nome iniciado por "Estoque" e depois qualquer nome que contenha "Estoque"; sem correspondência, usa a primeira aba.
- A prévia lê até 30 linhas físicas da aba para localizar cabeçalhos mesmo quando há títulos/linhas introdutórias antes da tabela.
- O mapeamento é inferido pelos rótulos das colunas, procurando Número/Código, Descrição/Nome/Item e Quantidade/Qtd/Estoque/Saldo. Mantém fallback A/B/D quando a planilha não oferece cabeçalhos reconhecíveis. O usuário pode remapear e cada aba tem o seu mapeamento persistido.
- A cópia original XLSX continua guardada na pasta privada do app. Na próxima abertura, o aplicativo reabre o mesmo arquivo, restaura a aba anteriormente escolhida e reaplica as colunas salvas.
- A seleção rápida de um único produto foi separada do campo de colagem em lote: o usuário escolhe Por código ou Por nome, digita, recebe sugestões diretamente da tabela completa de estoque, toca em uma sugestão e confirma o produto antes de adicioná-lo. A colagem múltipla continua em um campo separado.
- A busca do catálogo continua consultando a tabela SQLite completa `catalog`, sem ler a lista de plaquetas selecionadas. O filtro de nome é independente das seleções de checkbox.
- Preservada a correção de recorte/escala do logotipo CKTR da versão 4.4.0.
- A prévia de cabeçalhos foi ampliada para até 30 linhas físicas para localizar tabelas com cabeçalho abaixo de títulos de relatório.


## v4.5.1 — correção do teste do catálogo

- Corrigido o teste `catalog browser filters the complete imported stock database`: ele usava a variável `importer` sem declará-la no próprio escopo do teste.
- O teste agora lê explicitamente `offline/streaming-xlsx.ts` antes de validar a detecção de cabeçalhos e a prévia de até 30 linhas.
- Nenhuma lógica de produção do app foi alterada nesta correção; o filtro do catálogo continua consultando a tabela `catalog` completa.


## v4.5.2 — busca por nome/código e filtro do estoque

- A busca deixou de depender de `LIKE ... ESCAPE`: agora usa `instr()` sobre `name_norm` e `code_norm`, sem problemas com `%`, `_` e `\` digitados.
- O nome aceita várias palavras em qualquer ordem ("vassoura grande"), sem acento e sem diferenciar maiúsculas/minúsculas, e encontra o texto em qualquer parte do nome (antes o código só aceitava o começo).
- O código encontra o trecho em qualquer posição e ignora pontos, hifens, barras, espaços e zeros à esquerda ("12345" acha "1.234-5"; "123" acha "000123").
- O filtro "Todos os produtos do estoque" e as sugestões do autocomplete usam a mesma regra e ordenam: código exato, código que começa com o texto, nome que começa com o texto, demais.
- Novo `ensureCatalogSearchIndex()`: ao abrir o app, confere uma amostra do catálogo e reconstrói `code_norm`/`name_norm` se estiverem vazios ou desatualizados (por exemplo, base salva por versão antiga). Sem essas colunas, nenhuma busca encontrava nada.
- Erros da busca por sugestão agora aparecem na tela em vez de virarem "Nenhuma sugestão encontrada".
- Testes de regressão ajustados para a nova consulta. Não consegui rodar `npm test` neste ambiente (sem rede para instalar dependências); rode `npm install && npm test && npm run typecheck` antes do build.
