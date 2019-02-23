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
$computername = Import-Csv $file

$From = Get-Date "02/21/2019"
$To = $From.AddDays(2)

foreach ($c in $computername)
{
#######################Get Send Email Count#######################

$a = Get-TransportService | Get-MessageTrackingLog -ResultSize Unlimited -Sender $c.EmailAddress -EventId RECEIVE -Source "STOREDRIVER" -Start $From -End $To
$d = 0
foreach ($r in $a)
{

	$b = (((((((($r.Recipients -notmatch "domain.com") -notmatch "domain.com") -notmatch "domain.com") -notmatch "domain.com") -notmatch "domain.com") -notmatch "domain.com") -notmatch "domain.com") -notmatch "domain.com").count
	$d = $d + $b
}
#######################Get Receive Email Count#######################
$e = Get-TransportService | Get-MessageTrackingLog -ResultSize Unlimited -Recipients $c.EmailAddress -EventId "DELIVER" -Start $From -End $To
$g = ((((((((($e.Sender -notmatch "domain.com") -notmatch "domain.com") -notmatch "domain.com") -notmatch "domain.com") -notmatch "domain.com") -notmatch "domain.com") -notmatch "domain.com") -notmatch "domain.com") -notmatch "domain.com").count
$c.InternalSend = $a.Recipients.count - $d
$c.ExternalSend = $d
$c.InternalReceive = $e.Sender.count - $g
$c.ExternalReceive = $g
}
$computername | Export-Csv D:\Script\Finish.csv -NoTypeInformation
