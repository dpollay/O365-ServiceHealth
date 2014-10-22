function o365-ServiceHealth{

begin{
$htmlfile = $env:HOME + '\o365-ServiceHealth.html'
out-file $htmlfile
Add-Type -AssemblyName System.Web
$Header = @"
<style>
TABLE {border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}
TH {border-width: 1px;padding: 3px;border-style: solid;border-color: #fff!important;background-color: #0072c6;font-family: SegoeUI-Regular-final,"Segoe UI",Segoe,Tahoma,Helvetica,Arial,Sans-Serif;font-size: 11px;color: #fff;text-align: left;text-transform: uppercase;letter-spacing: .01em;font-weight: normal;}
TD {border-width: 1px;padding: 3px;border-style: solid;border-color: #fff!important;font-family: SegoeUI-Regular-final,"Segoe UI",Segoe,Tahoma,Helvetica,Arial,Sans-Serif;font-size: 11px;color: #333;text-align: left;letter-spacing: .01em;font-weight: normal;}
#PRE {font-family: SegoeUI-Regular-final,"Segoe UI",Segoe,Tahoma,Helvetica,Arial,Sans-Serif;font-size: 11px;color: #333;text-align: left;letter-spacing: .01em;font-weight: bold;text-transform: uppercase;}
.odd  { background-color:#ffffff; }
.even { background-color:#dddddd; }
TR:Nth-Child(Even) {Background-Color: #dddddd;}
TR:Hover TD {Background-Color: #C1D5F8;}
</style>
<center>
<title>
O365 Service Health
</title>
<script src='http://code.jquery.com/jquery-1.5.1.min.js' type='text/javascript'></script>
 <script type='text/javascript'>
		   
             `$(document).ready(function()
             {
                 `$('.RowToClick').click(function ()
                 {
                     `$(this).nextAll('tr').each( function()
                     {
                         if (`$(this).is('.RowToClick'))
                        {
                           return false;
                        }
                        `$(this).toggle(350);
                     });
                 });
             });
 </script>
"@


<#
    This code section was taken from Cam Murray's Technet Blog
    http://blogs.technet.com/b/cammurray/archive/2014/09/24/using-powershell-to-obtain-the-office365-dashboard.aspx
#>

#$cred = get-credential

$jsonPayload = (@{userName=$cred.username;password=$cred.GetNetworkCredential().password;} | convertto-json).tostring()
$cookie = (invoke-restmethod -contenttype "application/json" -method Post -uri "https://api.admin.microsoftonline.com/shdtenantcommunications.svc/Register" -body $jsonPayload).RegistrationCookie

$jsonPayload = (@{lastCookie=$cookie;locale="en-US";preferredEventTypes=@(0,1);pastDays="30"} | convertto-json).tostring()
$notice = (invoke-restmethod -contenttype "application/json" -method Post -uri "https://api.admin.microsoftonline.com/shdtenantcommunications.svc/GetEvents" -body $jsonPayload)
}

process{ $events = $notice.Events | Sort-Object LastUpdatedTime -Descending

$servicehealth = foreach($event in $events){
 $information = $event | select-object -Property StartTime, Status, ID, Messages, LastUpdatedTime, AffectedServiceHealthStatus
 $messagesSorted = $event.Messages |sort-object PublishedTime -Descending
 $information.Messages = $null
 $pre = '<p id="pre">' + $information.AffectedServiceHealthStatus.ServiceName + " - " + $information.Status + "</p>"
 foreach($message in $messagesSorted.MessageText){
    $information.Messages += $message + '<br><br>'
 } 
 $information | select @{name="Start Time";expression={([datetime]$information.StartTime).AddHours(-4)}}, `
    @{name="Status";expression={$information.Status}}, `
    @{name="ID";expression={$information.ID}}, `
    @{name="Last Updated";expression={([datetime]$information.LastUpdatedTime).AddHours(-4)}}, `
    @{name="Messages";expression={$Information.messages}} `
    | ConvertTo-Html -Fragment -As Table -PreContent $pre
}

$servicehealth = $servicehealth -replace '<tr><th>', '<tr class="RowToClick"><th>' -replace '&lt;br&gt;', '<br>'
$Report = ConvertTo-Html `
    -Head $header `
    -Body "$ServiceHealth"

}

end{$report | out-file $htmlfile ; Invoke-Expression $htmlfile}

}
o365-ServiceHealth