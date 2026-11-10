# Revisão offline e interface Material You

## Funcionamento offline

- Removida a dependência da API/Render na aplicação mobile. Importação de planilhas, pesquisa, lista de plaquetas, contagem, leitura de códigos, armazenamento e exportações são executados no aparelho.
- Modelo PowerPoint padrão incluído no pacote do app; modelos selecionados são copiados para o armazenamento privado.
- Permissões Android de internet e estado da rede bloqueadas no `app.json`.
- Dados locais guardam base importada, registros de contagem, modelo atual, tema e ajustes de posição.

## Prévia e posição

- O inspetor PPTX agora também extrai geometria, texto, cor, alinhamento, contorno e imagens dos objetos comuns do slide para montar a prévia.
- A tela mostra o slide completo e atualiza os campos à medida que o usuário altera os deslocamentos.
- Adicionado ajuste global e por posição individual (posição física na folha), separado para código e nome.
- Os deslocamentos em milímetros são salvos no armazenamento local e aplicados ao XML do PowerPoint na exportação.
- Mapeamento manual de IDs também é refletido na prévia e os IDs inválidos, duplicados ou sobrepostos são validados.
- O PDF usa grade simplificada e os mesmos deslocamentos; o PPTX é o formato que preserva o modelo original.

## Interface

- Fluxo de criação dividido em três passos: Base, Itens e Modelo.
- Paletas claras azul, verde, roxo, âmbar e turquesa inspiradas no Material You.
- Componentes com superfícies arredondadas, estados selecionados e controles de posição por setas.
- Tema aplicado também ao módulo de contagem.

## Compilação e validação

- Java 17 para o build Android.
- Dependências do Expo verificadas antes de compilar.
- Picker fixado em `2.11.1` para o SDK 54.
- `npm install` usado nos workflows até o `package-lock.json` ser adicionado ao repositório.
- O job de simulador iOS não gera IPA instalável em iPhone real; para isso é preciso assinar com certificado/perfil Apple.

Nesta revisão, a instalação de dependências do npm não respondeu dentro do tempo disponível nesta sessão. Portanto, não foi possível executar `npm run typecheck`, `npm test`, Gradle ou Xcode localmente. Os testes automatizados incluídos no repositório serão executados pelo GitHub Actions na próxima compilação.
