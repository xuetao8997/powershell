$externalAuthority="*.domain.com"
 
$ServiceName = "00000002-0000-0ff1-ce00-000000000000";
 
$p = Get-MsolServicePrincipal -ServicePrincipalName $ServiceName;
 
$spn = [string]::Format("{0}/{1}", $ServiceName, $externalAuthority);
$p.ServicePrincipalNames.Add($spn);
 
Set-MsolServicePrincipal -ObjectID $p.ObjectId -ServicePrincipalNames $p.ServicePrincipalNames;