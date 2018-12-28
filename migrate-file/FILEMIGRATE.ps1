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

$migrate = Import-Csv $file

foreach ($m in $migrate)
{
$n = 0
subst M: $m[$n].PATH_OLD
subst N: $m[$n].PATH_NEW
subst
robocopy M: N: *.* /e /copy:DAT /dcopy:T /mir /XF Thumbs.db *.ost *.tmp *.gho *.bak /MT:128 /r:30 /w:0 /tee /log:D:\LOGS\$m[$n].NAME_NEW.log
subst M: /D
subst N: /D
++$n
}
#$mailbox | Export-Csv -Path $env:USERPROFILE\Desktop\Finish.csv -NoTypeInformation
