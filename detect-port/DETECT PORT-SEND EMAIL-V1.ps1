$n = 0
do
{
$start = Get-Date
Start-Sleep -Seconds 60
$tcp = New-Object Net.Sockets.TcpClient
$tcp.Connect("x.x.x.x",25)
$end = Get-Date
$min = ($end-$start).Minutes
$sec = ($end-$start).Seconds
if ($tcp.Connected -eq $true)
{
Send-MailMessage -To "username@domain.com" -From "username@domain.com" -Subject "[SUCCESS]TELNET 25"  -Body "DURING $min MINUTES $sec SECONDS" -SmtpServer "domain.com" -Port 25
}
else
{
Send-MailMessage -To "username@domain.com" -From "username@domain.com" -Subject "[FAIL]TELNET 25"  -Body "DURING $min MINUTES $sec SECONDS" -SmtpServer "domain.com" -Port 25
}
$n++
}
while ($n -ne 0)