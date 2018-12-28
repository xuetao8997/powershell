Function Get-TCPResponse {
    <#
        .SYNOPSIS
            Tests TCP port of remote or local system and returns a response header
            if applicable
        .DESCRIPTION
            Tests TCP port of remote or local system and returns a response header
            if applicable
            If server has no default response, then Response property will be NULL
        .PARAMETER Computername
            Local or remote system to test connection
        .PARAMETER Port
            TCP Port to connect to
        .PARAMETER TCPTimeout
            Time until connection should abort
        .NOTES
            Name: Get-TCPResponse
            Author: Boe Prox
            Version History:
                1.0 -- 15 Jan 2014
                    -Initial build
        .INPUTS
            System.String
        .OUTPUTS
            Net.TCPResponse
        .EXAMPLE
        Get-TCPResponse -Computername Exchange1 -Port 25
        Computername : Exchange1
        Port         : 25
        IsOpen       : True
        Response     : 220 SMTP Server Ready
        Description
        -----------
        Checks port 25 of an exchange server and displays header response.
    #>
    [OutputType('Net.TCPResponse')]
    [cmdletbinding()]
    Param (
        [parameter(ValueFromPipeline,ValueFromPipelineByPropertyName)]
        [Alias('__Server','IPAddress','IP')]
        [string[]]$Computername = $env:Computername,
        [int[]]$Port = 902,
        [int]$TCPTimeout = 1000
    )
    Process {
        ForEach ($Computer in $Computername) {
            ForEach ($_port in $Port) {
                $stringBuilder = New-Object Text.StringBuilder
                $tcpClient = New-Object System.Net.Sockets.TCPClient
                $connect = $tcpClient.BeginConnect($Computer,$_port,$null,$null) 
                $wait = $connect.AsyncWaitHandle.WaitOne($TCPtimeout,$false) 
                If (-NOT $wait) {
                    $object = [pscustomobject] @{
                        Computername = $Computer
                        Port = $_Port
                        IsOpen = $False
                        Response = $Null
                    }
                } Else {
                    While ($True) {
                        #Let buffer
                        Start-Sleep -Milliseconds 1000
                        Write-Verbose "Bytes available: $($tcpClient.Available)"
                        If ([int64]$tcpClient.Available -gt 0) {
                            $stream = $TcpClient.GetStream()
                            $bindResponseBuffer = New-Object Byte[] -ArgumentList $tcpClient.Available
                            [Int]$response = $stream.Read($bindResponseBuffer, 0, $bindResponseBuffer.count)  
                            $Null = $stringBuilder.Append(($bindResponseBuffer | ForEach {[char][int]$_}) -join '')
                        } Else {
                            Break
                        }
                    } 
                    $object = [pscustomobject] @{
                        Computername = $Computer
                        Port = $_Port
                        IsOpen = $True
                        Response = $stringBuilder.Tostring()
                    }
                }
                $object.pstypenames.insert(0,'Net.TCPResponse')
                Write-Output $object
                If ($Stream) {
                    $stream.Close()
                    $stream.Dispose()
                }
                $tcpClient.Close()
                $tcpClient.Dispose()
            }
        }
    }
}
$n = 0
New-EventLog -LogName Application -Source Test
do
{
$start = Get-Date
Start-Sleep -Seconds 30
$resp = (Get-TCPResponse -Computername x.x.x.x -Port 25).Response
$end = Get-Date
$min = ($end-$start).Minutes
$sec = ($end-$start).Seconds
if ($resp -like "220 *")
{
Write-EventLog -LogName Application -EventId 2500 -Source Test -Message “$resp”
}
else
{
$tcp = New-Object Net.Sockets.TcpClient
$tcp.Connect("x.x.x.x",25)
if ($tcp.Connected -eq $true)
{
Write-EventLog -LogName Application -EventId 2501 -Source Test -Message “PORT 25 CONNECTED , BUT COULD NOT GET 220 RESPONSE” -EntryType Warning
}
else
{
Write-EventLog -LogName Application -EventId 2502 -Source Test -Message “PORT 25 DISCONNECTED , COULD NOT GET 220 RESPONSE” -EntryType Error
}
}
$n++
}
while ($n -ne 0)
#Remove-EventLog -Source Test