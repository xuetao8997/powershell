$UserCredential = Get-Credential
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://partner.outlook.cn/PowerShell-LiveID/ -Credential $UserCredential -Authentication Basic -AllowRedirection
Import-PSSession $Session
