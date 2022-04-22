# Get iibrary files from SharePoint Online

Sharepoint Online의 라이브러리 목록에서 파일을 가져옵니다.
왕창 가장 빨리 가장 많이..

* 계정에 mfa가 걸려 있는 경우 실패합니다.

dotnet core 버전은 SharePointOnlineCredentials 를 지원하지 않아 다른 방식으로 인증되나,
현재는 실패하고 있네요... 왜 그럴까요?

## Technologies

사용된 기술

### AE.O365.GetFiles.CSApp 

- dotnet core 6
- [Microsoft.Extensions.Configuration.Json](https://www.nuget.org/packages/Microsoft.Extensions.Configuration.Json)
- [Microsoft.Extensions.DependencyInjection](https://www.nuget.org/packages/Microsoft.Extensions.DependencyInjection)
- [Microsoft.Extensions.Logging](https://www.nuget.org/packages/Microsoft.Extensions.Logging)
- [Microsoft.Extensions.Logging.Console](https://www.nuget.org/packages/Microsoft.Extensions.Logging.Console)
- [Microsoft.Extensions.Logging.Debug](https://www.nuget.org/packages/Microsoft.Extensions.Logging.Debug)
- [Microsoft.SharePointOnline.CSOM](https://www.nuget.org/packages/Microsoft.SharePointOnline.CSOM)
- [ByteSize](https://www.nuget.org/packages/ByteSize)
- [System.IdentityModel](https://www.nuget.org/packages/System.IdentityModel.Tokens.Jwt)