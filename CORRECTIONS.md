# Correções desta revisão

## Inicialização do aplicativo

- Corrigido um erro de inicialização encontrado na revisão do código: `useWindowDimensions` estava sendo importado de `react`, mas é exportado por `react-native`. A tela principal chamava esse hook ao montar, o que podia impedir o app de abrir.
- O Android físico não usa mais `10.0.2.2` como endereço padrão (esse host é específico do emulador).
- O endereço inicial do backend agora é o HTTPS que já está declarado no `render.yaml`: `https://plaquetas-by-lippe.onrender.com`.
- A configuração manual do servidor valida o formato da URL e a interface informa quando a API está indisponível.
- A verificação inicial do backend foi simplificada e a câmera só pede permissões quando o usuário abre o leitor.
- Foi incluído um limite de recuperação para erros de renderização do React; isso não captura falhas nativas fatais do Android.
- O projeto foi configurado para a arquitetura antiga do React Native no SDK 54 como tentativa de compatibilidade para o APK que não iniciava. Essa alteração exige uma nova compilação nativa.
- O app não solicita permissão de microfone porque só usa a câmera para ler códigos.

## GitHub Actions

- Corrigida a incompatibilidade vista nos logs de 10/11/2026: `@react-native-picker/picker` foi fixado em `2.11.1`, a versão indicada pelo Expo SDK 54; antes, o npm instalava `2.11.4` por causa do intervalo `^2.11.1`, e `expo install --check` parava os builds Android e iOS.
- Adicionada verificação de tipos TypeScript aos jobs Android, simulador iOS e IPA.
- Atualizadas as ações oficiais do GitHub para versões compatíveis com o runtime Node.js 24 dos runners atuais.
- Removido o cache npm que falhava sem `package-lock.json`. Os builds usam `npm install` até o lockfile ser adicionado.
- Expo Prebuild agora roda com `CI=1` e `--clean`, em vez da opção obsoleta `--non-interactive`.
- Android usa JDK 17 e compila APK release após limpar o Gradle.
- iOS descobre o workspace e o scheme reais em vez de presumir o nome do workspace.
- O artefato do simulador iOS é empacotado em ZIP para facilitar o download; ele não é instalável em iPhone físico.
- Mantido job opcional para IPA assinado via Apple Developer / TestFlight.

## Validação e limitações

- O JSON do Expo e o YAML do workflow foram validados por parser.
- O arquivo TSX foi analisado pelo parser do TypeScript sem erros de sintaxe.
- Não foi possível executar `npm install`, `npm test`, Gradle ou Xcode nesta sessão porque o acesso ao registry npm falhou por `EAI_AGAIN`; portanto, esta revisão não é apresentada como build nativo concluído. O workflow continuará usando `npm install` e gerará o lockfile de trabalho no runner; nenhum `package-lock.json` falso foi criado.
- Os logs mais recentes mostram que Android e simulador iOS pararam em `expo install --check` por causa do Picker fora da versão recomendada. Não mostram uma compilação nativa concluída nem incluem `logcat`; portanto, ainda não existe evidência de teste do APK corrigido em aparelho físico.
- O job `ios-ipa` não aparece nessa execução; somente o simulador foi acionado. Para IPA é preciso definir `BUILD_IOS_IPA=true` e cadastrar os Secrets/Variables Apple descritos no README.
