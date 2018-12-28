$filePath = "Source Path"
Get-ChildItem -Recurse -Force $filePath -ErrorAction SilentlyContinue | Where-Object {!$_.PSIsContainer -and  ($_.Name -like "*Key Word*") } | Copy-Item -Dest "Destination Path"
