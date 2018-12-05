#
.SYNOPSIS
    Compare the Exchange Online IP list with the locally saved list to detect changes.
.DESCRIPTION
    This script takes the Exchange Online IP list(https://support.content.office.net/en-us/static/O365IPAddresses.xml) and saves it locally. The next time, it compares the local list to the online one, to detect if there were changes. If there were, it sends a mail to the INIT Security group to notify them. If there weren't any changes, it still send a mail, just to say that it still works.    
    The Script is now Using the REST Webservice to retreive the IP's
    https://support.office.com/de-de/article/verwalten-von-office-365-endpunkten-99cab9d4-ef59-4207-9f2b-3728eb46bf9a?ui=de-DE&rs=de-DE&ad=DE#ID0EACAAA=4._Webdienst
    Author: a.bohren@icewolf.ch http://blog.icewolf.ch 
    V1.1 Andres Bohren - Initial Version
    
#>

#Define the path, where the file gets saved. It just takes the location of the script.
$path = Split-Path -parent $PSCommandPath

#Mailing related variables
[string]$smtpserver = "SMTPServer.domain.tld"
[string]$smtpfrom = "sender@foo.com"
[string]$smtpto = "recipient@foo.com"
$smtptochange = @("changes1@foo.com","changes2@foo.com")
[string]$body = ""

#Where the IPs get stored
[object[]]$changedIps
$removedIps = [System.Collections.ArrayList]@()
$addedIps = [System.Collections.ArrayList]@()

#Get the Version of the Last Change 
#$ClientRequestId = new-guid
#$ClientRequestId = "045bb3bb-dfcf-4359-a597-b16f6281b1ff"
#8b9387ca-f435-4296-ae88-a1fe669de6c4 ICE10
#$uri = "https://endpoints.office.com/version/O365Worldwide?ClientRequestId=$ClientRequestId"
#$Result = Invoke-RestMethod -Method GET -uri $uri
#$Lastchange = $Result.latest #2018033000
#Write-Host "LastChange: $Lastchange"
#$Lastchange | Out-File "C:\Scripts\TaskScheduledScripts\Daily-ExoIPchecker\o365ipversion.txt"

#Get Exchange Endpoints
#$uri = "https://endpoints.office.com/endpoints/O365Worldwide?ServiceAreas=Exchange&ClientRequestId=$ClientRequestId"
#$Result = Invoke-RestMethod -Method GET -uri $uri

#Exchange without Common --> IP's only
#$ips = ($Result | where {$_.serviceArea -eq "Exchange"}).ips

#Exchange without Common --> IPv4 IP's only
#$addresses = $ips | where {$_ -notmatch ":"}

#Check
if(Test-Path "$path\addresses.txt"){
    foreach($ip in (Compare-Object -ReferenceObject (Get-Content "$path\addresses.txt") -DifferenceObject $addresses)){
        if($ip.SideIndicator -eq "=>"){
            $addedIps.Add($ip.InputObject) | Out-Null
        }elseif($ip.SideIndicator -eq "<="){
            $removedIps.Add($ip.InputObject) | Out-Null
        }
    }
    #if there were changes, format the IPs
    if($addedIps.Count -ne "" -or $removedIps.Count -ne ""){
        if($addedIps.Count -ne ""){
            $addedIps = $addedIps | ForEach-Object { "$_<br />" }
            $addedIps = $addedIps -join "`n"
            $body = $body + "<strong><br />These IPs were added:<br /></strong>" + $addedIps
        }
        if($removedIps.Count -ne ""){
            $removedIps = $removedIps | ForEach-Object { "$_<br />" }
            $removedIps = $removedIps -join "`n"
            $body = $body + "<strong><br />These IPs were removed:<br /></strong>" + $removedIps
        }           
        #Send the update Mail 
        Send-MailMessage -SmtpServer $smtpserver -From $smtpfrom -To $smtptochange -Subject "EXO IP Checker - IP's changed Warning" -Body ("<span style='font-family:Arial;font-size:11pt'>There were some changes in the EXO IP list.<br />"+ $body +"</span>") -BodyAsHtml
        
		#Save the changes locally
        $addresses > $path\addresses.txt
    }else{
        #If there are no changes, send a mail so that the admins know, that the script still works
        Send-MailMessage -SmtpServer $smtpserver -From $smtpfrom -To $smtpto -Subject "EXO IP Checker - INFO" -Body ("<span style='font-family:Arial;font-size:11pt'>There were no changes, I am just letting you know that I still work.</span>") -BodyAsHtml
    }
}else{
    #If the file does not exist yet, create it
    $addresses > "$path\addresses.txt"
}
