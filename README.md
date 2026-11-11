# Plaquetas by Lippe

Aplicativo Android/iOS para importar planilhas Excel, preparar plaquetas a partir de modelos PowerPoint, conferir o posicionamento dos textos, exportar arquivos e registrar contagens. **Depois de instalado, o aplicativo funciona sem servidor e sem conexão com a internet.**

## O que funciona localmente

- Importação de planilhas `.xlsx`, escolha da aba e mapeamento das colunas de código, descrição e quantidade. O arquivo é lido no aparelho.
- Pesquisa por código/nome, sugestões por prefixo nos dois campos e montagem da lista de plaquetas usando a base armazenada localmente. A etapa Itens também exibe um catálogo pesquisável por trecho do nome, com código, seleção por caixas e carregamento paginado.
- Modelo PowerPoint padrão incluído no próprio aplicativo; modelos personalizados são copiados para o armazenamento do app.
- Prévia em tempo real do slide carregado, reconstruída a partir das formas, caixas de texto, linhas e imagens que o aplicativo consegue interpretar no PPTX.
- Ajuste do texto de código ou nome por deslocamentos em milímetros: global para todas as posições do slide ou individual para uma posição da folha. Os ajustes são gravados no aparelho e usados na exportação.
- Geração de PowerPoint no aparelho, preservando o arquivo-modelo original e movendo as caixas de texto selecionadas.
- Geração local de PDF. **A saída PDF usa uma grade de impressão simplificada**, não é uma conversão completa do PowerPoint. Para manter todos os elementos gráficos e detalhes do modelo original, exporte o PPTX.
- Leitura de QR Code/código de barras e contagens +1/+10 processadas e armazenadas localmente.
- Cinco paletas claras inspiradas no Material You: azul, verde, roxo, âmbar e turquesa. A escolha fica salva no aparelho.

Os dados, planilhas, modelos e ajustes permanecem no armazenamento privado do aplicativo. Desinstalar o app ou limpar seus dados pode apagar essas informações. Faça cópias dos arquivos de origem e exportados.

## Prévia e ajustes de posicionamento

Na etapa **Modelo**, a prévia mostra o slide inteiro. Toque no campo de código ou nome para selecioná-lo. Escolha **Todas as plaquetas** para aplicar o deslocamento a todas as posições daquele tipo de texto, ou **Só uma plaqueta** para ajustar a posição escolhida na folha. Selecione a precisão do movimento (0,5 / 1 / 2 / 5 mm) e use as setas. O deslocamento aparece na prévia e é salvo automaticamente.

O ajuste individual é feito por posição física na folha: por exemplo, a posição 3 recebe o mesmo ajuste na posição 3 de cada folha gerada. Ele não altera apenas um item de uma lista longa.

A prévia vetorial reproduz os elementos visuais reconhecidos no arquivo. Efeitos avançados do PowerPoint, fontes que não existam no aparelho, transparências/gradientes especiais, rotações ou objetos complexos podem aparecer de forma aproximada na prévia; a exportação PPTX continua partindo do arquivo original. Confira uma folha de teste antes de imprimir em lote.

## Compilar Android e iOS pelo GitHub Actions

O workflow `.github/workflows/build-native.yml` instala dependências e compila os artefatos. O `.gitignore` permite versionar `package-lock.json`; depois do primeiro `npm install`, inclua esse arquivo no repositório para tornar as próximas instalações reproduzíveis. **O uso normal do app não precisa da internet; o processo de compilação no GitHub precisa acessar o npm, Gradle e os repositórios de dependências.**

1. Envie o projeto para a branch `main`.
2. Abra **Actions → Build Android and iOS**.
3. Baixe `plaquetas-android-apk` para Android.
4. O job `plaquetas-ios-simulator` gera um app para simulador, não para iPhone físico. Para um iPhone real, habilite `build_ios_ipa` e configure o certificado e perfil de provisionamento Apple descritos em `.github/workflows/README.md`.

## Seleção de produtos (v4.4)

Na etapa **Itens**, o app mostra os produtos da base local sem exigir uma busca inicial. A lista exibe apenas nome e código, com seleção por caixas. O campo de filtro procura termos em qualquer parte do nome e ignora diferenças entre letras acentuadas e sem acento (por exemplo, `calça` encontra `CALCA`). Os resultados são paginados em lotes de 60 para evitar carregar a base inteira na memória; use **Carregar mais produtos** para avançar. A seleção permanece marcada ao filtrar outros nomes e pode ser adicionada em lote às plaquetas.

Os campos **Por nome** e **Por código** oferecem sugestões conforme os primeiros caracteres digitados. Tocar em uma sugestão preenche o campo com o nome ou código completo; depois, use **Adicionar à lista**. Também continua sendo possível colar vários valores separados por quebra de linha, vírgula ou ponto e vírgula.

A prévia do PowerPoint respeita o modo de ajuste da imagem descrito no PPTX. Imagens com recorte definido pelo PowerPoint usam preenchimento com recorte, evitando que a imagem CKTR quadrada com grandes margens transparentes seja reduzida dentro da caixa horizontal do logotipo; imagens sem instrução de recorte continuam usando encaixe proporcional.

## Desenvolvimento

Requisitos: Node.js 22 e npm.

```bash
npm install
# Inclua o package-lock.json gerado no commit do projeto.
npm run typecheck
npm test
npm start
```

## Dependências principais

- Expo SDK 54 / React Native 0.81.
- `react-native-zip-archive` para listar o pacote XLSX e extrair somente os XMLs relevantes da aba escolhida.
- `saxes` para interpretar XML progressivamente em blocos, sem materializar a planilha inteira em memória.
- Expo FileSystem para leitura incremental de arquivos e armazenamento temporário privado.
- Expo SQLite para manter a base de produtos indexada e persistente, inclusive entre reinicializações.
- JSZip para inspecionar e gerar localmente arquivos PowerPoint.
- Expo Camera para leitura de códigos.
- Expo Print e Expo Sharing para criação e compartilhamento de PDF/arquivos.

## Rede e privacidade

A interface mobile não contém chamadas a uma API e não usa URL de backend. O manifesto Android bloqueia as permissões `INTERNET` e `ACCESS_NETWORK_STATE`. Planilhas, códigos, nomes, contagens e modelos são processados no aparelho. A câmera só é solicitada quando o usuário abre o leitor.


## Importação otimizada de planilhas grandes (v4.3)

O importador nativo não converte mais a planilha inteira para Base64 nem mantém um workbook completo na memória. Um arquivo `.xlsx` é aberto como pacote ZIP; o app extrai apenas o XML da aba escolhida, o catálogo de strings compartilhadas e os estilos. Os XMLs são lidos em blocos de 256 KB e o parser mantém apenas as colunas mapeadas (por padrão A = código, B = descrição e D = quantidade (ajustável)).

Os registros brutos são gravados em lotes de até 1.000 linhas no SQLite. Antes de ler as strings compartilhadas, o app consulta os IDs realmente usados nas colunas selecionadas em um cursor SQLite incremental; a lista de IDs fica em um vetor tipado compacto. O parser pula o texto das demais strings e encerra a leitura assim que encontra o último ID necessário. Em seguida, aplica máscaras numéricas simples (por exemplo, `000000` para preservar zeros à esquerda) e monta a nova base em tabelas temporárias indexadas. A base anterior só é substituída em uma transação depois que todas as etapas terminam. Cancelar ou falhar durante a leitura preserva o catálogo anterior.

A tela de importação mostra a etapa atual, progresso aproximado, quantidade de linhas, tempo decorrido atualizado a cada segundo, velocidade média e estimativa restante suavizada. A leitura pode ser cancelada. Não há mais um limite artificial de 50 MB; o tamanho prático passa a depender do armazenamento livre, da complexidade/expansão do XLSX e dos recursos do aparelho. Arquivos muito grandes ainda podem levar tempo porque o conteúdo XML e as strings precisam ser percorridos, mas a memória é mantida limitada por blocos e lotes.

O importador incremental exige uma nova compilação nativa, porque usa `react-native-zip-archive`. O pacote v9 requer iOS 15.5 ou superior e um rebuild Android/iOS após sua inclusão.


### Correção v4.2.1

A importação de modelos PowerPoint personalizados também lê o arquivo como bytes (`ExpoFile.bytes()`) e só converte para Base64 para o processamento interno do PPTX. Isso elimina a chamada obsoleta `readPickedFileBase64` que causava erro de TypeScript.

### Versão v4.3 — importador incremental

- Leitura seletiva do ZIP/XML de `.xlsx` em blocos de 256 KB; por padrão considera somente as colunas A, B e D, com mapeamento editável.
- IDs de strings compartilhadas coletados via cursor SQLite e representados em `Uint32Array`; o parser captura apenas textos referenciados e pode parar assim que encontra o último ID necessário.
- Gravação SQLite em lotes, índices de referência e substituição transacional da base.
- Preservação de códigos com zeros à esquerda em formatos numéricos simples.
- Tela de progresso com etapa, percentual, linhas lidas, duração atualizada a cada segundo, média de linhas por segundo, ETA suavizada e cancelamento.
- Sem limite fixo de 50 MB; o teto prático é limitado pelo espaço de armazenamento e capacidade do dispositivo.


### Build note for v4.3.1

The streaming XLSX reader uses the Expo SDK 54 `File.open()` signature with no mode argument. It continues reading the extracted XML in 256 KB chunks using `FileHandle.readBytes(length)`.

## Versão 4.5.0 — identificação automática da planilha de estoque

Ao escolher um arquivo XLSX, o aplicativo prioriza a aba cujo nome contém "Estoque" (correspondência exata, depois prefixo, depois ocorrência parcial). A detecção lê até 30 linhas iniciais para identificar a linha de cabeçalho e os campos Número/Código, Descrição/Nome/Item e Quantidade/Qtd/Estoque/Saldo. Se os cabeçalhos forem atípicos, a seleção manual de colunas continua disponível. As escolhas de aba e mapeamento ficam salvas por arquivo/aba, e a cópia do XLSX permanece no armazenamento privado do aparelho para ser restaurada na próxima abertura.

Na etapa Itens, a busca rápida individual agora é uma caixa de uma linha com sugestões por código ou nome, alimentadas pelo catálogo SQLite completo. Tocar numa sugestão confirma o produto e o botão “Adicionar produto selecionado” o inclui com código, descrição e quantidade. A área separada de colagem em lote continua disponível. O filtro do catálogo consulta o estoque importado completo, independentemente dos produtos já marcados para impressão.
