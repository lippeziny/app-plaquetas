# Plaquetas by Lippe

Aplicativo Android/iOS para importar planilhas Excel, preparar plaquetas a partir de modelos PowerPoint, conferir o posicionamento dos textos, exportar arquivos e registrar contagens. **Depois de instalado, o aplicativo funciona sem servidor e sem conexão com a internet.**

## O que funciona localmente

- Importação de planilhas `.xlsx`, escolha da aba e mapeamento das colunas de código, descrição e quantidade. O arquivo é lido no aparelho.
- Pesquisa por código/nome, sugestões e montagem da lista de plaquetas usando a base armazenada localmente.
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
- SheetJS (`xlsx`) para leitura local de planilhas.
- JSZip para inspeção e geração local de PowerPoint.
- Expo FileSystem para armazenamento privado no aparelho.
- Expo Camera para leitura de códigos.
- Expo Print e Expo Sharing para criação e compartilhamento de PDF/arquivos.

## Rede e privacidade

A interface mobile não contém chamadas a uma API e não usa URL de backend. O manifesto Android bloqueia as permissões `INTERNET` e `ACCESS_NETWORK_STATE`. Planilhas, códigos, nomes, contagens e modelos são processados no aparelho. A câmera só é solicitada quando o usuário abre o leitor.


## Importação de planilhas grandes (v4.2)

A importação Excel no Android lê o arquivo como `Uint8Array` usando `File.bytes()` do Expo FileSystem, em vez de converter o arquivo completo para Base64 pela ponte nativa. O parser usa `dense: true` e mantém o workbook em memória, convertendo para linhas somente a aba selecionada pelo usuário. Para evitar encerramento por falta de memória, o seletor orienta reduzir arquivos acima de 50 MB e abas acima de 250 mil linhas. Isso reduz cópias e consumo de memória durante a leitura.
