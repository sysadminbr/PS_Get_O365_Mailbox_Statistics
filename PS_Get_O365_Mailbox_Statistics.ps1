<#
# CITRA IT - EXCELÊNCIA EM TI
# SCRIPT PARA OBTER STATISTICAS DE USO DAS CAIXAS DE E-MAIL NO OFFICE365
# AUTOR: luciano@citrait.com.br
# DATA: 01/10/2021
# Homologado para executar no Windows 10 ou Server 2012R2+
# EXAMPLO DE USO: Powershell -ExecutionPolicy ByPass -File C:\Scripts\PS_Get_O365_Mailbox_Statistics.ps1 AdminUser AdminPassword
#
# Referência: https://docs.microsoft.com/pt-br/powershell/exchange/exchange-online-powershell-v2?view=exchange-ps#install-and-maintain-the-exo-v2-module
# Pre-requisitos:
# 1. Instalar o gerenciador de pacotes .net/powershell NuGet com o comando abaixo (powershell como adm)
#     [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 ; Install-Module -Name PowerShellGet -Force
#
# 2. Instalar o Módulo Exchange Online V2 com o comando (powershell como adm) abaixo:
#     [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 ; Install-Module -Name ExchangeOnlineManagement
#
# 3. Criar um usuário novo no portal do office365, não atribuir nenhuma licença e marcar a função de "Administração Exchange" apenas.
# Usar teste usuário/senha para a conexão do script.
#>

#Requires -Version 5

Param(
	[Parameter(Mandatory=$True)] [string]$AdminUser,
    [Parameter(Mandatory=$True)] [string]$AdminPass
)


#
# Console log function
#
Function Log()
{
	Param([String] $text)
	$timestamp = Get-Date -Format G
	Write-Host -ForegroundColor Green "$timestamp`: $text"
	
}


#
# Transforming the plaintext provided password into a PSCredential
#
[Console]::Clear() # clear console if password passed as command line argument !
Log "===  Starting Script to Fech Office365 Mailbox Statistics  ==="
Log "Converting given password to secure password"
$user = $AdminUser
$pass = ConvertTo-SecureString -AsPlainText -Force $AdminPass
$UserCredential = New-Object System.Management.Automation.PSCredential($user, $pass)


#
# Connecting to Online Exchange
#
Log "Connecting to connect to Office365 via Powershell"
try{
	$ProgressPreference = "SilentlyContinue"
	Connect-ExchangeOnline -UserPrincipalName $user -credential $UserCredential -ShowBanner:$false
	$ProgressPreference = "Continue"
}catch{
    Write-Host -ForegroundColor RED "Error conecting to office365. Check your Credentials !"
    Write-Host $_.Exception.Message
    Exit
}


#
# Retrieving list of available mailboxes
#
Log "Retrieving Mailboxes..."
$Mailboxes = Get-MailBox | Where-Object { $_.RecipientTypeDetails -ne "DiscoveryMailbox"}



#
# Holds the accounts and statistics information
#
$table = @()


#
# Compiling the Statistics
#
Log "Compiling Statistics..."


#
# Holds progress bar percent processed
#
$processed = 0


$Mailboxes | ForEach-Object{
	$this_mailbox = $_
	$stats = Get-EXOMailboxStatistics -Identity $this_mailbox
	If($Mailboxes.Count) { Write-Progress -Activity $this_mailbox.PrimarySmtpAddress -PercentComplete ([math]::round((100/$Mailboxes.Count) * $processed)) }
	$table += [PSCustomObject]@{
		Email=$this_mailbox.PrimarySmtpAddress; 
		Quota=$this_mailbox.ProhibitSendReceiveQuota; 
		Inbox=$stats.TotalItemSize; 
		Deleted=$stats.TotalDeletedItemSize;
	}
	$processed += 1
		
}
Log "Statistics compiled !"

#
# Displaying the results...
#
$table | Sort-Object -Property Inbox -Descending | Out-GridView -Title "Listagem de E-mail | Office365" -Wait



#
# Nicely disconnect from Exchange Online 0/
#
Try{
	Disconnect-ExchangeOnline -Confirm:$false
}Catch{
	Write-Host -ForegroundColor RED "Error disconecting to Office365. Already disconnected ? "
    Write-Host $_.Exception.Message
    Exit
}
