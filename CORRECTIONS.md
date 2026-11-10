# Correções desta revisão

## Inicialização do aplicativo

- O Android físico não usa mais `10.0.2.2` como endereço padrão (esse host é específico do emulador).
- O endereço inicial do backend agora é o HTTPS que já está declarado no `render.yaml`: `https://plaquetas-by-lippe.onrender.com`.
- A configuração manual do servidor valida o formato da URL e a interface informa quando a API está indisponível.
- A verificação inicial do backend foi simplificada e a câmera só pede permissões quando o usuário abre o leitor.
- Foi incluído um limite de recuperação para erros de renderização do React; isso não captura falhas nativas fatais do Android.
- O projeto foi configurado para a arquitetura antiga do React Native no SDK 54 como tentativa de compatibilidade para o APK que não iniciava. Essa alteração exige uma nova compilação nativa.
- O app não solicita permissão de microfone porque só usa a câmera para ler códigos.

## GitHub Actions

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
- Não foi possível executar `npm install`, `npm test`, Gradle ou Xcode nesta sessão porque o acesso ao registry npm falhou por `EAI_AGAIN`; portanto, esta revisão não é apresentada como build nativo concluído.
- O arquivo de log enviado continha somente a falha de configuração do `setup-node` por falta de lockfile. Ele não continha `logcat` nem a exceção nativa que explica o fechamento do APK. Se o APK atualizado ainda fechar, o `logcat` do Android será necessário para fechar o diagnóstico.
