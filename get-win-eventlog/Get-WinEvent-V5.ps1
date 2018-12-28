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
$start = Get-Date
foreach ($c in $computername)
{
if ((Get-Service -ComputerName $c.Computername -Name RpcSs).Status -eq "Running")
{
$c.Event1000 = (Get-WinEvent -ComputerName $c.Computername -ea SilentlyContinue -FilterHashtable @{ProviderName= "Application Error";LogName = "Application"; StartTime = [DateTime]::Now.AddDays(-7)} | Where-Object{$_.ID -eq 1000 -and $_.Message -like "*OUTLOOK.EXE*"}).count
$c.Event1002 = (Get-WinEvent -ComputerName $c.Computername -ea SilentlyContinue -FilterHashtable @{ProviderName= "Application Error";LogName = "Application"; StartTime = [DateTime]::Now.AddDays(-7)} | Where-Object{$_.ID -eq 1002 -and $_.Message -like "*OUTLOOK.EXE*"}).count
}
else
{
$c.Event1000 = "Failed"
$c.Event1002 = "Failed"
}
}
$computername | Export-Csv .\Finish.csv -NoTypeInformation
$end = Get-Date
$min = ($end-$start).Minutes
$sec = ($end-$start).Seconds
Send-MailMessage -To "username@domain.com" -From "username@domain.com" -Subject "Finish Outlook Error $b" -Attachments ".\Finish.csv" -Body "Time is $min Minutes $sec Seconds." -SmtpServer "IP/DOMAIN" -Port 25
Remove-Item .\Finish.csv -Force
