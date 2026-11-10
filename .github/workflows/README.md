# Builds nativos pelo GitHub Actions

O workflow `build-native.yml` cria três artefatos independentes:

- `plaquetas-android-apk`: APK Android de release para instalação em aparelhos físicos.
- `plaquetas-ios-simulator`: ZIP com o `.app` para simulador; não instala em iPhone físico.
- `plaquetas-ios-ipa`: IPA assinado para distribuição via TestFlight/App Store Connect, somente quando o job estiver configurado.

## API usada pelos aplicativos

Em **Settings → Secrets and variables → Actions → Variables**, configure `API_URL` com a URL HTTPS pública do backend. Se ela não existir, o workflow usa `https://plaquetas-by-lippe.onrender.com`, endereço declarado em `render.yaml`.

## IPA para iPhone físico

O job é opcional. Ao executar manualmente o workflow, marque a opção **Gerar IPA assinado para iPhone físico**; alternativamente, configure a variável `BUILD_IOS_IPA` como `true`. Nos dois casos, os certificados e perfis Apple precisam estar cadastrados. Cadastre:

| Tipo | Nome | Conteúdo |
|---|---|---|
| Secret | `APPLE_TEAM_ID` | Team ID da conta Apple Developer |
| Secret | `IOS_CERTIFICATE_P12_BASE64` | Certificado de distribuição `.p12` convertido para Base64 |
| Secret | `IOS_CERTIFICATE_PASSWORD` | Senha do arquivo `.p12` |
| Secret | `IOS_PROVISIONING_PROFILE_BASE64` | Perfil de provisionamento compatível convertido para Base64 |
| Variable | `IOS_PROVISIONING_PROFILE_NAME` | Nome exato do perfil de provisionamento |
| Variable | `BUILD_IOS_IPA` | `true` |

O Bundle ID configurado é `com.lippe.plaquetas`. O certificado e o perfil precisam corresponder a esse identificador e ao mesmo Apple Team. O método configurado exporta para App Store Connect/TestFlight; ele não é um IPA Ad Hoc para instalação direta por sideload.

## Observações

- Não existe `package-lock.json` nesta cópia do projeto. Por isso, os jobs usam `npm install` e não habilitam cache npm. O workflow não falha mais no `setup-node` por ausência de lockfile.
- O Android usa Java 17.
- O pacote `@react-native-picker/picker` está fixado em `2.11.1`, versão recomendada pelo Expo SDK 54.
- Antes dos builds, o workflow verifica os tipos TypeScript, a compatibilidade dos pacotes Expo e os testes do projeto.
- As pastas `android/` e `ios/` são geradas novamente por `expo prebuild --clean`; elas não precisam ser versionadas.
- O APK gerado sem keystore próprio usa a assinatura de debug do runner. Para atualizar uma instalação anterior assinada com outra chave, desinstale a versão anterior antes de instalar o novo APK.
