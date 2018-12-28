<#
Start this ps1 with taskschd.msc when use logon.
#>

do
{
If $DFS equal True, then exit loop
$DFS = Test-Path "\\GENSCRIPT.COM\DFS"
sleep(5)
}
while($DFS -eq $False)

#Re-mapping network drive
Remove-PSDrive -Name "O"
New-PSDrive -Name "O" -Root "THE PATH YOU NEED TO MAP" -Persist -PSProvider "FileSystem" -Scope "Global"
Remove-PSDrive -Name "P"
New-PSDrive -Name "P" -Root "THE PATH YOU NEED TO MAP" -Persist -PSProvider "FileSystem" -Scope "Global"
Remove-PSDrive -Name "Q"
New-PSDrive -Name "Q" -Root "THE PATH YOU NEED TO MAP" -Persist -PSProvider "FileSystem" -Scope "Global"
