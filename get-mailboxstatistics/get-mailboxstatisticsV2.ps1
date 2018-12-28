#[Microsoft.Exchange.Data.Directory.AdminSessionADSettings]::Instance
Add-PSSnapin Microsoft.Exchange*
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

$mailbox = Import-Csv $file
foreach ($m in $mailbox)
{
$m.database = (Get-MailboxStatistics -Identity $m.emailaddress).Database.Name
$m.size = (Get-MailboxStatistics -Identity $m.emailaddress).TotalItemSize.Value
$m.ProhibitSendReceiveQuota = (Get-Mailbox -Identity $m.emailaddress).ProhibitSendReceiveQuota.Value
}
$mailbox | Export-Csv -Path $env:USERPROFILE\Desktop\Finish.csv -NoTypeInformation
