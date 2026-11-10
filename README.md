# Plaquetas by Lippe

Aplicativo universal (website Node.js + Android + iOS) para importar uma base Excel, preencher plaquetas a partir de códigos ou nomes, revisar cada registro e gerar arquivos PowerPoint/PDF prontos para impressão.

## O que já está incluído

- **Base de referência Excel:** aceita `.xlsx` com várias abas. O usuário escolhe a aba e mapeia as colunas de código, descrição/nome e quantidade. A leitura do arquivo usa streaming para não carregar todas as linhas simultaneamente. O valor da quantidade fica disponível como referência.
- **Lista por código ou nome:** aceita vários valores colados de uma vez. A busca por nome apresenta sugestões da base; a busca por código recupera a descrição e o estoque.
- **Correção individual:** cada plaqueta tem campos editáveis de código e nome, além de exclusão e visualização individual.
- **Dois modelos do PowerPoint enviado:** o arquivo original `PLAQUETAS_BY_FILLIPE.PPTX` está incluído como modelo padrão. A inspeção inicial identifica dois slides, cada um com oito pares de campos de nome/código.
- **Modelo substituível:** aceita outro `.pptx`. A detecção procura nomes de caixa como `name.item` e `cod.item`; também tenta reconhecer os campos visíveis no modelo atual. Quando não consegue identificar, a tela oferece um mapeamento por IDs das caixas de texto.
- **Exportação:** gera PPTX preservando os objetos e a formatação do slide selecionado, substituindo somente o conteúdo textual dos campos de nome e código. O PDF é convertido a partir desse PPTX usando LibreOffice no servidor.
- **Contagem experimental:** leitor de QR Code/código de barras, pesquisa do item e registro manual por toque de `+1` ou `+10`. Os registros acumulam por código.
- **Compilação pelo GitHub Actions:** workflow Android gera APK para aparelho físico; workflow iOS gera um pacote de simulador e inclui um job opcional para IPA assinado. O app agora mostra uma tela de recuperação caso ocorra um erro de renderização.

## Executar localmente

Requisitos: Node.js 22, npm e (para conversão de PDF) LibreOffice instalado no computador.

```bash
npm install
cp .env.example .env
npm run dev:api
```

Em outro terminal:

```bash
npm run start:web
```

O site abre pelo Expo. Em desenvolvimento web, o app tenta acessar `http://localhost:4000` como API. No aplicativo nativo, o endereço inicial aponta para `https://plaquetas-by-lippe.onrender.com`, conforme o serviço definido em `render.yaml`. Caso publique a API em outro endereço, abra a engrenagem no topo e informe a URL HTTPS pública. Para testes locais, use o IP LAN do computador e a porta `4000`; o endereço `10.0.2.2` é exclusivo do emulador Android e não deve ser usado em um telefone físico.

## Usar a planilha enviada

1. Abra **Banco de dados → Enviar Excel**.
2. Escolha a aba que contém os produtos (na planilha de referência enviada, a aba de interesse é `ESTOQUE SAP`).
3. Confira os mapeamentos. No formato exibido no print, normalmente são: `Nº do item` = código (coluna A), `Descrição do item` = nome (coluna B) e `Em estoque` = quantidade (coluna F). A tela permite alterar as colunas para outras planilhas.
4. Clique em **Importar esta aba como referência**.

O arquivo Excel enviado não é incorporado ao repositório por ser grande. O usuário faz o upload pelo próprio aplicativo e escolhe a aba na hora da importação.

## Como o modelo PPTX é tratado

O PowerPoint padrão fornecido é armazenado sem ser editado. Na exportação, o servidor cria uma cópia do arquivo e substitui apenas o texto dos campos de código e nome identificados. Fontes, tamanhos, bordas, imagens, posições, cores e demais caixas permanecem provenientes do slide original. Os slides que não correspondem ao modelo selecionado não são usados na saída; a página escolhida é duplicada quando a lista excede a capacidade de uma folha.

Para facilitar a troca de modelo, nomeie as caixas de texto repetidas no PowerPoint:

- `cod.item` para o código do produto;
- `name.item` para o nome/descrição do produto.

Use esses nomes como o **nome da forma** no Painel de Seleção do PowerPoint. Se preferir, também é possível inserir os IDs das caixas no mapeamento avançado que aparece quando a detecção automática não encontra os campos.

A quantidade de plaquetas por folha é detectada pelo número de campos de código e nome. O campo de folhas mostra a quantidade mínima calculada; um número maior pode ser informado para acrescentar folhas em branco. O sistema não reduz o número mínimo necessário para caber todos os itens.

## Publicar website e API no Render

O `Dockerfile` instala Node.js, LibreOffice e as dependências do app. O `render.yaml` configura um serviço web e um disco persistente para o banco SQLite.

1. Envie esta pasta para um repositório GitHub.
2. No Render, crie um serviço usando o Blueprint (`render.yaml`) ou conecte o repositório como serviço Docker.
3. O serviço executa `npm run build:web` durante a imagem e `npm run start:server` ao iniciar.
4. Configure a URL pública do serviço em `EXPO_PUBLIC_API_URL` para compilar os aplicativos nativos.

Para acionar deploy por GitHub Actions, crie o secret `RENDER_DEPLOY_HOOK` com o Deploy Hook do Render. A automação do deploy é opcional; o Render também pode acompanhar diretamente a branch `main`.

## Compilar Android e iOS pelo GitHub

O workflow `.github/workflows/build-native.yml` executa automaticamente em pushes na branch `main` ou manualmente em **Actions → Build Android and iOS → Run workflow**. Ele usa Node 22, Java 17 para Android, recria as pastas nativas com `expo prebuild --clean` e não tenta usar cache npm enquanto não existir um lockfile no repositório.

- **Android:** baixa o artefato `plaquetas-android-apk` com `app-release.apk`, instalável para testes em aparelhos Android. O APK de teste é assinado com a configuração de desenvolvimento do projeto e não é uma versão pronta para publicação na Play Store.
- **iOS:** baixa `plaquetas-ios-simulator`, um ZIP de `.app` compilado para o simulador iOS. Esse arquivo não instala em um iPhone físico.
- **iPhone físico/TestFlight:** o job opcional `ios-ipa` pode gerar `plaquetas-ios-ipa` quando a variável `BUILD_IOS_IPA` estiver como `true` e os certificados/perfil Apple descritos em `.github/workflows/README.md` estiverem cadastrados nos Secrets. É necessária uma conta Apple Developer e um perfil de distribuição válido.

Para que os apps nativos acessem o servidor publicado, crie em **Settings → Secrets and variables → Actions → Variables** a variável `API_URL` com a URL HTTPS pública da API. O workflow repassa essa variável para `EXPO_PUBLIC_API_URL` durante a compilação. Quando `API_URL` estiver ausente, usa o endereço Render já declarado no `render.yaml`.

**Limitação importante do iOS:** a compilação para simulador não exige assinatura e não instala em um iPhone físico. Para gerar IPA e distribuir via TestFlight/App Store, é necessária uma conta Apple Developer e certificados/perfil válidos. O workflow já contém um job opcional de IPA; consulte `.github/workflows/README.md` para cadastrar os Secrets e habilitá-lo. A versão SDK 54 está configurada temporariamente com a arquitetura antiga do React Native como medida de compatibilidade para investigar a falha de inicialização do APK anterior; o SDK 54 permite alternar essa configuração.

## Dados e segurança

- Os uploads temporários ficam em `server/uploads`; os arquivos gerados ficam em `server/outputs`.
- O banco SQLite fica em `server/data` localmente e em `/var/data` no Render, conforme o `render.yaml`.
- A base importada contém os códigos, descrições, quantidades e o nome da aba de origem; o arquivo inteiro não é necessário para cada consulta.
- Não há login nesta primeira versão. Para uso em produção com dados corporativos, configure autenticação e controle de acesso antes de disponibilizar o serviço publicamente.

## Estrutura

```text
App.tsx                     Aplicativo universal: web, Android e iOS
server/index.cjs            API Node/Express e endpoints de exportação
server/lib/workbook.cjs     Inspeção e importação streaming do Excel
server/lib/pptx.cjs         Detecção e preenchimento OOXML do PowerPoint
server/assets/               Modelo padrão enviado pelo usuário
server/db.cjs               Banco SQLite
.github/workflows/           Build Android/iOS e deploy opcional
Dockerfile                  API + website + LibreOffice
render.yaml                 Configuração do serviço Render
```

## Limitações conhecidas desta versão

- A prévia dentro do app é uma representação individual simplificada. A prévia completa para impressão é o PDF real gerado a partir do PowerPoint.
- A detecção dos campos em modelos novos funciona melhor com as caixas nomeadas `name.item` e `cod.item`. O mapeamento por IDs é a alternativa para modelos sem essa convenção.
- A função de contagem é experimental e não substitui uma auditoria/inventário validado. Os dados são salvos no banco do servidor.
- A versão inicial não tem múltiplos usuários, autenticação, permissões por filial ou sincronização offline.
