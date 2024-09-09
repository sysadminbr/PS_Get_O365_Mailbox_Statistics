# PS_Get_O365_Mailbox_Statistics
## Script Powershell para exibir estatísticas de uso de caixas de e-mail do Office365 (Microsoft365)

### Pré-Requisitos:
1. Instalar o gerenciador de pacotes .net/powershell NuGet com o comando abaixo (powershell como adm):
    ```
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    Install-Module -Name PowerShellGet -Force
    ```
2. Instalar o Módulo Exchange Online V2 com o comando (powershell como adm) abaixo:
    ```
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    Install-Module -Name ExchangeOnlineManagement
    ```
3. Criar um usuário novo no portal do office365, não atribuir nenhuma licença e marcar a função de "Administração Exchange" apenas.
Usar este usuário/senha para a conexão do script.

![image](https://user-images.githubusercontent.com/91758384/135780808-53b503a1-7e30-47e0-aca7-f567e2248bd3.png)

![image](https://user-images.githubusercontent.com/91758384/135780844-a48fd0c6-d764-4eea-b9ab-fd2803a21d75.png)
