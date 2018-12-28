$n = 0
New-EventLog -LogName Application -Source Test
do
{
#$start = Get-Date
Start-Sleep -Seconds 60
$tcp = New-Object Net.Sockets.TcpClient
$tcp.Connect("x.x.x.x",25)
#$end = Get-Date
#$min = ($end-$start).Minutes
#$sec = ($end-$start).Seconds
if ($tcp.Connected -eq $true)
{
Write-EventLog -LogName Application -EventId 2500 -Source Test -Message “PORT 25 SUCCESS”
}
else
{
Write-EventLog -LogName Application -EventId 2501 -Source Test -Message “PORT 25 FAILED”
}
$n++
}
while ($n -ne 0)
#Remove-EventLog -Source Test