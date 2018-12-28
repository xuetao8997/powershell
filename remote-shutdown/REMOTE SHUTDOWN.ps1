Function Get-FileName($initialDirectory)
{   
 [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null

 $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
 $OpenFileDialog.initialDirectory = $initialDirectory
 $OpenFileDialog.filter = "All files (*.*)| *.*"
 $OpenFileDialog.ShowDialog() | Out-Null
 $OpenFileDialog.filename
} #end function Get-FileName
# *** Entry Point to Script ***
$file = Get-FileName
$computername = Import-Csv $file
foreach ($c in $computername)
{
$UserName = "username@domain.com"
$PlainPassword = "password"
$SecurePassword = $PlainPassword | ConvertTo-SecureString -AsPlainText -Force
$Credentials = New-Object System.Management.Automation.PSCredential -ArgumentList $UserName, $SecurePassword
Stop-Computer -Server $c.dnsname -Credential $Credentials -Force
}