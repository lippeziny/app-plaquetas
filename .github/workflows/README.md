# GitHub Actions

- `build-native.yml`:
  - APK Android para testes (`plaquetas-android-apk`).
  - Aplicativo para simulador iOS, sem assinatura (`plaquetas-ios-simulator`).
  - Job opcional `ios-ipa`, habilitado quando a variável `BUILD_IOS_IPA` estiver como `true`; compila e assina o IPA para distribuição usando certificados Apple.
- `deploy-render.yml`: chama o Deploy Hook do Render quando o secret `RENDER_DEPLOY_HOOK` está configurado.

## Variáveis e secrets

Em **Settings → Secrets and variables → Actions**:

**Variables**
- `API_URL`: endereço HTTPS público do backend.
- `BUILD_IOS_IPA`: use `true` para habilitar o job de IPA.
- `IOS_PROVISIONING_PROFILE_NAME`: nome exato do perfil de provisionamento Apple.

**Secrets**
- `RENDER_DEPLOY_HOOK`: Deploy Hook opcional do Render.
- `APPLE_TEAM_ID`: Team ID da conta Apple Developer.
- `IOS_CERTIFICATE_P12_BASE64`: certificado de distribuição `.p12`, codificado em Base64.
- `IOS_CERTIFICATE_PASSWORD`: senha do arquivo `.p12`.
- `IOS_PROVISIONING_PROFILE_BASE64`: perfil de provisionamento `.mobileprovision`, codificado em Base64.

A versão para simulador não exige conta Apple Developer. O IPA para iPhone físico/TestFlight exige certificado e perfil válidos vinculados ao Bundle ID `com.lippe.plaquetas`.
