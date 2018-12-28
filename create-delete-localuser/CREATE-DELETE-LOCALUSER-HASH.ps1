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

$hash = @{}
Import-Csv $file | foreach { $hash.add($_.MAC, $_.TAG) }
    foreach($obj in $hash.Keys)
    {
        "key = " + $obj ;
        "value = "+ $hash[$obj] ;
        if ($hash[$obj] -eq 1)
            {
                New-LocalUser -Name $obj -Password (ConvertTo-SecureString -AsPlainText "$obj" -Force) -PasswordNeverExpires -UserMayNotChangePassword
            }
        if ($hash[$obj] -eq 0)
            {
                Remove-LocalUser -Name $obj
            }
        if ($hash[$obj] -eq 2)
            {
                Disable-LocalUser -Name $obj
            }
    }