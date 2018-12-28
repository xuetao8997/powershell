# use first option if you want to impersonate, otherwise, grab your own credentials
#$s.Credentials = New-Object Net.NetworkCredential($username, $password, $domain)
#$s.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials
[Reflection.Assembly]::LoadFile("C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll")
$s = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2016_CU10)
$s.Credentials = New-Object Net.NetworkCredential('USERNAME', 'PASSWORD', 'DOMAIN.COM')
$s.AutodiscoverUrl("USERNAME@DOMAIN.COM")
$s.UseDefaultCredentials = $true

# discover the url from your email address
$s.AutodiscoverUrl(“lg-com-xuetao@genscript.com”)

# get a handle to the inbox
$inbox = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($s,[Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox)

#create a property set (to let us access the body & other details not available from the FindItems call)
$psPropertySet = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
$psPropertySet.RequestedBodyType = [Microsoft.Exchange.WebServices.Data.BodyType]::Text;

$items = $inbox.FindItems(5)

#set colours for Write-Color output
$colorsread = "Yellow","White"
$colorsunread = "Red","White"

# output unread count
Write-Color -Text "Unread count: ",$inbox.UnreadCount -Color $colorsread

foreach ($item in $items.Items)
{
  # load the property set to allow us to get to the body
  $item.load($psPropertySet)

  # colour our output
  If ($item.IsRead) { $colors = $colorsread } Else { $colors = $colorsunread }

  #format our body
  #replace any whitespace with a single space then get the 1st 100 chars
  $bod = $item.Body.Text -replace '\s+', ' '
  $bodCutOff = (100,$bod.Length | Measure-Object -Minimum).Minimum
  $bod = $bod.Substring(0,$bodCutOff)
  $bod = "$bod..."

  # output the results - first of all the From, Subject, References and Message ID
  write-host "====================================================================" -foregroundcolor White
  Write-Color "From:    ",$($item.From.Name) $colors
  Write-Color "Subject: ",$($item.Subject)   $colors
  Write-Color "Body:    ",$($bod)            $colors
  write-host "====================================================================" -foregroundcolor White
  ""
}


# display the newest 5 items
#$inbox.FindItems(5)
# display the unread items from the newest 5
#$inbox.FindItems(5) | ?{$_.IsRead -eq $False} | Select Subject, Sender, DateTimeSent | Format-Table -auto

# returns the number of unread items
# $inbox.UnreadCount


#see these URLs for more info
# EWS
# folder members: https://msdn.microsoft.com/en-us/library/microsoft.exchange.webservices.data.folder_members%28v=exchg.80%29.aspx
# exporting headers: http://www.stevieg.org/tag/how-to/
# read emails with EWS: https://social.technet.microsoft.com/Forums/en-US/3fbf8348-2945-43aa-a0bc-f3b1d34da27c/read-emails-with-ews?forum=exchangesvrdevelopment
# Powershell
# multi-color lines: http://stackoverflow.com/a/2688572



# download the Exchange Web Services Managed API 1.2.1 from
# http://www.microsoft.com/en-us/download/details.aspx?id=30141
# extract somewhere, e.g. ...
# msiexec /a C:\Users\YourUsername\Downloads\EwsManagedApi.msi /qb TARGETDIR=C:\Progs\EwsManagedApi
