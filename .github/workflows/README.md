# Builds nativos pelo GitHub Actions

O workflow `build-native.yml` produz estes artefatos:

- `plaquetas-android-apk`: APK Android de release para aparelhos físicos.
- `plaquetas-ios-simulator`: app `.app` compactado para simulador iOS; não instala em iPhone físico.
- `plaquetas-ios-ipa`: IPA assinado para iPhone real, somente quando a opção de build estiver ativada e os certificados Apple estiverem configurados.

## Aplicativo sem servidor

O app mobile é totalmente local após a instalação: não utiliza API, Render, banco de dados remoto nem variável `API_URL`. O Android bloqueia as permissões de rede. A conexão usada pelo GitHub Actions é apenas para baixar ferramentas e dependências durante a compilação; ela não é necessária para usar o aplicativo instalado.

## IPA para iPhone físico

Ao executar manualmente o workflow, marque **Gerar IPA assinado para iPhone físico**; alternativamente, configure a variável `BUILD_IOS_IPA` como `true`. Cadastre os seguintes itens em **Settings → Secrets and variables → Actions**:

| Tipo | Nome | Conteúdo |
|---|---|---|
| Secret | `APPLE_TEAM_ID` | Team ID da conta Apple Developer |
| Secret | `IOS_CERTIFICATE_P12_BASE64` | Certificado de distribuição `.p12` convertido para Base64 |
| Secret | `IOS_CERTIFICATE_PASSWORD` | Senha do arquivo `.p12` |
| Secret | `IOS_PROVISIONING_PROFILE_BASE64` | Perfil de provisionamento compatível convertido para Base64 |
| Variable | `IOS_PROVISIONING_PROFILE_NAME` | Nome exato do perfil de provisionamento |
| Variable | `BUILD_IOS_IPA` | `true` (quando quiser habilitar o IPA em todos os pushes) |

O Bundle ID é `com.lippe.plaquetas`. O certificado e o perfil precisam corresponder a esse identificador e ao mesmo Apple Team. O método atual exporta para App Store Connect/TestFlight.

## Observações de compilação

- Se o repositório ainda não incluir `package-lock.json`, os jobs usam `npm install` sem cache npm. Após o primeiro `npm install`, versione o `package-lock.json` para travar as dependências; então o workflow pode ser alterado para `npm ci` e cache npm.
- O Android usa Java 17.
- `@react-native-picker/picker` está fixado em `2.11.1`, compatível com o Expo SDK 54.
- Os jobs executam TypeScript, verificação de dependências do Expo e testes locais antes da compilação.
- As pastas `android/` e `ios/` são geradas por `expo prebuild --clean`.
- A assinatura do APK pode diferir entre execuções; para instalar um APK assinado por outra chave, desinstale a versão anterior primeiro.
