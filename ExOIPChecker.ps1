###############################################################################
# ExOIPChecker - Exchange Online IP Checker
# Andres Bohren / www.icewolf.ch / blog.icewolf.ch / a.bohren@icewolf.ch
# Version 1.0 / 10.06.2019 - Initial Version 
# Version 1.1 / 08.05.2020 - Autoupdate ReceiveConnector 
# Version 1.2 / 24.06.2020 - Cleanup Script
###############################################################################
<#
.SYNOPSIS
    Compare the Exchange Online Protection IP list with the locally saved list to detect changes.
	The User must have Exchange Organisation Admin Permissions to Update Exchange Receive Connector
	
.DESCRIPTION
    This script takes the REST API to extract the Exchange Online IP's and saves them locally.
    https://docs.microsoft.com/en-us/office365/enterprise/urls-and-ip-address-ranges
    https://docs.microsoft.com/en-us/office365/enterprise/office-365-ip-web-service

	
	The next time, it compares the local list to the online one, to detect if there were changes. 
	If there are Changes, it sends a mail to the Change Recipients to notify them.
	If there weren't any changes, it still send a mail, just to say that the scipt was running.

.EXAMPLE
	./ExOIPChecker.ps1

.LINK
    Check https://github.com/BohrenAn/ExOIPChecker

#>

###############################################################################
# Connect ExchangeOnPrem
###############################################################################
Function Connect-ExchangeOnPrem
{
    PARAM 
    (
        [parameter(Mandatory=$true)][string]$OnPremExchangeServer
    )

    #Create PSSession
    $ExSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$OnPremExchangeServer/PowerShell/ -Authentication Kerberos
    #ImportPSSession
    Import-PSSession -Session $ExSession -DisableNameChecking | Out-Null
    #Nun kann man die Exchange Cmdlets des Remote Servers benutzen
}

###############################################################################
# Disconnect ExchangeOnPrem
###############################################################################
Function Disconnect-ExchangeOnPrem
{
    Get-PSSession | where {$_.ConfigurationName -eq "Microsoft.Exchange"} | Remove-PSSession
}

###############################################################################
# This Function Updates the 365 Receive Connector
# Only the EOP IP's are allowed
###############################################################################
Function Update-ExchangeConnector
{

    PARAM 
    (
        [array]$OnPremReceiveConnector
    )

		#Connect to ExchangeOnPrem
		Connect-ExchangeOnPrem -OnPremExchangeServer $OnPremExchangeServer
		
        Try {

            [Array]$EOPIPv4 = @()
            $EOPIPv4 = (Get-Content "$path\addresses.txt")
            #Localhost needed for Managed Availability
	        $EOPIPv4 += "127.0.0.1"
            $EOPIPv4 += $CustomRemoteIPRanges
            
            foreach ($ReceiveConnector in $OnPremReceiveConnector)
            {
                Set-ReceiveConnector -identity $ReceiveConnector -RemoteIPRanges $EOPIPv4
            }

		} catch {
			write-host "ERROR: An error has occurred: `r`n $_.Exception.Message" -foregroundColor Red
		} finally {
			#Disconnect to ExchangeOnPrem
			Disconnect-ExchangeOnPrem
		}
		
}

###############################################################################
# This Function Sends the Admin Mail
###############################################################################
Function Send-AdminMail
{

    If ($Changed -eq $true)
    {
        #Send the update Mail 
        Send-MailMessage -SmtpServer $smtpserver -From $smtpfrom -To $smtptochange -Subject "EXO IP Checker - IP's changed Warning" -Body ("<span style='font-family:Arial;font-size:11pt'>There were some changes in the EXO IP list.<br />"+ $body +"</span>") -BodyAsHtml
    } else {
        #If there are no changes, send a mail so that the admins know, that the script still works
        Send-MailMessage -SmtpServer $smtpserver -From $smtpfrom -To $smtpto -Subject "EXO IP Checker - INFO" -Body ("<span style='font-family:Arial;font-size:11pt'>There were no changes, I am just letting you know that I still work.</span>") -BodyAsHtml
    }
}

###############################################################################
# This Function gets the Exchange Online IP's from Office 365 REST API
###############################################################################
Function Check-ExOIPs
{
    #Where the IPs get stored
    [object[]]$changedIps
    $removedIps = [System.Collections.ArrayList]@()
    $addedIps = [System.Collections.ArrayList]@()
    
    #Create a GUID if not existent and save it in the Registry
    If ((Test-Path "HKCU:\Software\ExOIPChecker") -eq $false)
    {
        #$GUID = (new-guid).guid
	$GUID = ([guid]::NewGuid()).guid
        New-Item -Path HKCU:\Software\ExOIPChecker
        Set-ItemProperty -Path HKCU:\Software\ExOIPChecker -Name Guid -Value $GUID
	$ClientRequestId = $GUID
    } else {
        $ClientRequestId = (Get-ItemProperty -Path HKCU:\Software\ExOIPChecker -Name guid).guid
    }


    #Get Exchange Endpoints
    $uri = "https://endpoints.office.com/endpoints/worldwide?ServiceAreas=Exchange&NoIPv6=true&ClientRequestId=$ClientRequestId"
    Write-Host "DEBUG: URL: $uri"
    $Result = Invoke-RestMethod -Method GET -uri $uri

    #EOP IP's
    $addresses = ($Result | where {$_.urls -match "mail.protection.outlook.com"}).ips | Sort-Object -Unique

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

            #Save the changes locally
            $addresses > $path\addresses.txt
			return $true
			
        }else{
            return $false
        }
    }else{
        #If the file does not exist yet, create it
        $addresses > "$path\addresses.txt"
    }
}


###############################################################################
# Main Script
###############################################################################
#Define the path, where the file gets saved. It just takes the location of the script.
$path = Split-Path -parent $PSCommandPath

#Check Parameter RegisterScheduledTask
If ($RegisterScheduledTask -eq $true)
{
	#Write-Host "DEBUG: PSCommandPath > $PSCommandPath"
	Write-Host "INFO: Create-ScheduledTask"
	Create-ScheduledTask -Scriptpath $PSCommandPath
}

#Variables
[string]$smtpserver = "SERVER OR IP"
[string]$smtpfrom = "Sender@domain.tld"
[string]$smtpto = "recipient@domain.tld"
$smtptochange = @("Changes@domain.tld","recipient@domain.tld")
[string]$body = ""

[bool]$UpdateExchangeConnector = $true
[string]$OnPremExchangeServer = "ICESRV06"
[Array]$OnPremReceiveConnector = @()
$OnPremReceiveConnector = "ICESRV06\Default Frontend ICESRV06"
[Array]$CustomRemoteIPRanges = @()
$CustomRemoteIPRanges = "172.21.175.0/24"

#Check EOP IP's
$Changed = Check-ExOIPs
If ($Changed -eq $true)
{
    #Changes detected
    Write-Host "INFO: Changes detected"
} else {
    #No Changes Detected
    Write-Host "INFO: No Changes detected" -ForegroundColor Green
    get-Content "$path\addresses.txt"
}


If ($UpdateExchangeConnector -eq $true)
{
	Write-Host "INFO: Update Exchange Receive Connector"
        Update-ExchangeConnector -OnPremReceiveConnector $OnPremReceiveConnector
}

Write-Host "INFO: Sending Admin Mail"
#Call Function Send-AdminMail
Send-AdminMail
