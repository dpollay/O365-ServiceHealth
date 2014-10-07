function o365-ServiceHealth{

begin{
$htmlfile = $env:HOME + '\o365-ServiceHealth.html'
out-file $htmlfile
$Header = @"
<style>
TABLE {border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}
TH {border-width: 1px;padding: 3px;border-style: solid;border-color: black;background-color: #6495ED;}
TD {border-width: 1px;padding: 3px;border-style: solid;border-color: black;}
.odd  { background-color:#ffffff; }
.even { background-color:#dddddd; }
TR:Nth-Child(Even) {Background-Color: #dddddd;}
TR:Hover TD {Background-Color: #C1D5F8;}
</style>
<center>
<title>
O365 Service Health
</title>
"@

<#
    This code section was taken from Cam Murray's Technet Blog
    http://blogs.technet.com/b/cammurray/archive/2014/09/24/using-powershell-to-obtain-the-office365-dashboard.aspx
#>

#$cred = get-credential

$jsonPayload = (@{userName=$cred.username;password=$cred.GetNetworkCredential().password;} | convertto-json).tostring()
$cookie = (invoke-restmethod -contenttype "application/json" -method Post -uri "https://api.admin.microsoftonline.com/shdtenantcommunications.svc/Register" -body $jsonPayload).RegistrationCookie

$jsonPayload = (@{lastCookie=$cookie;locale="en-US";preferredEventTypes=@(0,1)} | convertto-json).tostring()
$notice = (invoke-restmethod -contenttype "application/json" -method Post -uri "https://api.admin.microsoftonline.com/shdtenantcommunications.svc/GetEvents" -body $jsonPayload)
}

process{ $events = $notice.Events

$ServiceHealth = $events |Sort-Object LastUpdatedTime -Descending | ForEach-Object{
    $information = $_ | select-object -Property StartTime, Status, ID, Title, Messages, LastUpdatedTime, PublishedTime
    $messagesSorted = $_.Messages |sort-object PublishedTime -Descending
    foreach($message in $MessagesSorted){
        $information.Messages = $message.MessageText
        $information.PublishedTime = $message.PublishedTime
        $information
    }
} | select @{name="Start Time";expression={([datetime]$information.StartTime).AddHours(-4)}}, `
    @{name="Status";expression={$information.Status}}, `
    @{name="Title";expression={$information.Title}}, `
    @{name="Update Time";expression={([datetime]$information.PublishedTime).AddHours(-4)}}, `
    @{name="Messages";expression={[string]$information.Messages}} `
    | ConvertTo-Html -Fragment -As Table | Out-String
    

$Report = ConvertTo-Html -Title "Office 365 Service Health" `
    -Head $header `
    -Body "$ServiceHealth"
}

end{$report | out-file $htmlfile ; Invoke-Expression $htmlfile}

}
o365-ServiceHealth