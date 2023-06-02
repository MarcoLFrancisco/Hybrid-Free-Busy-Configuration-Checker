<#
.SYNOPSIS
.\FreeBusyChecker.ps1 
.DESCRIPTION
The script can be used to validate the Availability configuration of the following Exchange Server Versions: - Exchange Server 2013 - Exchange Server 2016 - Exchange Server 2019 - Exchange Online

Required Permissions:
    Organization Management
    Domain Admins (only necessary for the DCCoreRatio parameter)
Please make sure that the account used is a member of the Local Administrator group. This should be fulfilled on Exchange servers by being a member of the Organization Management group. However, if the group membership was adjusted or in case the script is executed on a non-Exchange system like a management server, you need to add your account to the Local Administrator group.

How To Run:
This script must be run as Administrator in Exchange Management Shell on an Exchange Server. You can provide no parameters and the script will just run against Exchnage On Premises and Exchange Online to query for OAuth and DAuth configuration setting. It will compare existing values with standard values and provide detail of what may not be correct.
Please take note that though this script may output that a specific setting is not a standard sertting, it does not mean that your configurations are incorrect. For exmaple, DNS may be configured with specific mapppings that this script can not evaluate.

.PARAMETER Auth
Allow you to choosse the authentication type to validate.
.PARAMETER Org
Allow you to choosse the organizartion type to validate.
.PARAMETER Pause
Pause after each test done..
.PARAMETER Help
Show help of this script.


.EXAMPLE
.\FreeBusyChecker.ps1 
This cmdlet will run Free Busy Checker script and check Availability OAuth and DAuth Configurations both for Exchange On Premises and Exchange Online.
.EXAMPLE
.\FreeBusyChecker.ps1 -Auth OAuth
This cmdlet will run the Free Busy Checker Script against for OAuth Availability Configurations only.
.EXAMPLE
.\FreeBusyChecker.ps1 -Auth DAuth
This cmdlet will run the Free Busy Checker Script against for DAuth Availability Configurations only.
.EXAMPLE
.\FreeBusyChecker.ps1 -Org ExchangeOnline
This cmdlet will run the Free Busy Checker Script for Exchange Online Availability Configurations only.
.EXAMPLE
.\FreeBusyChecker.ps1 -Org ExchangeOnPremise
This cmdlet will run the Free Busy Checker Script for Exchange On Premises OAuth and DAuth Availability Configurations only.
.EXAMPLE
.\FreeBusyChecker.ps1 -Org ExchangeOnPremise -Auth OAuth -Pause
This cmdlet will run the Free Busy Checker Script for Exchange On Premises Availability OAuth Configurations, pausing after each test done.
#>

#Exchange on Premise
#>
#region Properties and Parameters
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', '', Justification = 'Variables are being used')]
[CmdletBinding(DefaultParameterSetName = "FreeBusyInfo_OP", SupportsShouldProcess)]

param(
    [Parameter(Mandatory = $false, ParameterSetName = "Test")]
    [ValidateSet('DAuth', 'OAuth')]
    [string[]]$Auth,
    [Parameter(Mandatory = $false, ParameterSetName = "Test")]
    [ValidateSet('ExchangeOnPremise', 'ExchangeOnline')]
    [string[]]$Org,
    [Parameter(Mandatory = $false, ParameterSetName = "Test")]
    [switch]$Pause,
    [Parameter(Mandatory = $true, ParameterSetName = "Help")]
    [switch]$Help
)

Function ShowHelp {
    $bar
    Write-host -ForegroundColor Yellow "`n  Valid Input Option Parameters!"
    Write-Host -ForegroundColor White "`n  Paramater: Auth"
    Write-Host -ForegroundColor White "   Options  : DAuth; OAUth"
    Write-Host  "    DAuth             : DAuth Authentication"
    Write-Host  "    OAuth             : OAuth Authentication"
    Write-Host  "    Default Value     : No swith input means the script will collect both DAuth and OAuth Availability Configuration Detail"
    Write-Host -ForegroundColor White "`n  Paramater: Org"
    Write-Host -ForegroundColor White "   Options  : ExchangeOnPremise; Exchange Online"
    Write-Host  "    ExchangeOnPremise : Use ExchangeOnPremise parameter to collect Availability information in the Exchange On Premise Tenant"
    Write-Host  "    ExchangeOnline    : Use Exchange Online parameter to collect Availability information in the Exchange Online Tenant"
    Write-Host  "    Default Value     : No swith input means the script will collect both Exchange On Premise and Exchange OnlineAvailability configuration Detail"
    Write-Host -ForegroundColor White "`n  Paramater: Pause"
    Write-Host  "                 : Use the Pause parameter to use this script pausing after each test done."
    Write-Host -ForegroundColor White "`n  Paramater: Help"
    Write-Host  "                 : Use the Help parameter to use display valid parameter Options. `n`n"
}

If ($Help) {
    Write-Host $bar
    ShowHelp;
    $bar
    exit
}

Add-PSSnapin microsoft.exchange.management.powershell.snapin
import-module ActiveDirectory
Install-Module -Name ExchangeOnlineManagement
DisConnect-ExchangeOnline -Confirm:$False
Clear-Host
$countOrgRelIssues = (0)
$Global:FedTrust = $null
$Global:AutoDiscoveryVirtualDirectory = $null
$Global:OrgRel
$Global:SPDomainsOnprem
$AvailabilityAddressSpace = $null
$Global:WebServicesVirtualDirectory = $null
$ConsoleWidth = $Host.UI.RawUI.WindowSize.Width
$bar = "="
for ( $i = 1; $i -lt $ConsoleWidth; $i++) {
    $bar += "="
}
$logfile = "$PSScriptRoot\FreeBusyInfo_OP.txt"
$startingDate = (get-date -format yyyyMMdd_HHmmss)
$Logfile = [System.IO.Path]::GetFileNameWithoutExtension($logfile) + "_" + `
    $startingDate + ([System.IO.Path]::GetExtension($logfile))
$htmlfile = "$PSScriptRoot\FBCheckerOutput_$($startingDate).html"
Write-Host " `n`n "
Start-Transcript -path $LogFile -append
Write-Host $bar
Write-Host -foregroundcolor Green " `n  Free Busy Configuration Information Checker `n "
Write-Host -foregroundcolor White "   Version -1 `n "
Write-Host -foregroundcolor Green "  Loading Parameters..... `n "
#Parameter input
$UserOnline = get-remotemailbox -resultsize 1 -WarningAction SilentlyContinue
$UserOnline = $UserOnline.RemoteroutingAddress.smtpaddress
$ExchangeOnlineDomain = ($UserOnline -split "@")[1]

if ($ExchangeOnlineDomain -like "*.mail.onmicrosoft.com") {
    $ExchangeOnlineAltDomain = (($ExchangeOnlineDomain.Split(".")))[0] + ".onmicrosoft.com"
}

else {
    $ExchangeOnlineAltDomain = (($ExchangeOnlineDomain.Split(".")))[0] + ".mail.onmicrosoft.com"
}
# $UserOnPrem = get-mailbox -resultsize 1 -WarningAction SilentlyContinue | Where-Object { ($_.EmailAddresses -like "*" + $ExchangeOnlineDomain ) }
$temp = "*" + $ExchangeOnlineDomain
$UserOnPrem = get-mailbox -resultsize 2 -WarningAction SilentlyContinue -Filter 'EmailAddresses -like $temp -and HiddenFromAddressListsEnabled -eq $false'
$UserOnPrem = $UserOnPrem[1].PrimarySmtpAddress.Address
$Global:ExchangeOnPremDomain = ($UserOnPrem -split "@")[1]
$EWSVirtualDirectory = Get-WebServicesVirtualDirectory

if ($EWSVirtualDirectory.externalURL.AbsoluteUri.Count -gt 1) {
    $Global:ExchangeOnPremEWS = ($EWSVirtualDirectory.externalURL.AbsoluteUri)[0]
}

else {
    $Global:ExchangeOnPremEWS = ($EWSVirtualDirectory.externalURL.AbsoluteUri)
}

$ADDomain = Get-ADDomain
$ExchangeOnPremLocalDomain = $ADDomain.forest

if ([string]::IsNullOrWhitespace($ADDomain)) {
    $ExchangeOnPremLocalDomain = $exchangeOnPremDomain
}
$Script:fedinfoEOP = get-federationInformation -DomainName $ExchangeOnPremDomain  -BypassAdditionalDomainValidation -ErrorAction SilentlyContinue | Select-Object *
#endregion

#region Edit Parameters

Function UserOnlineCheck {
    Write-Host -foregroundcolor Green " Online Mailbox: $UserOnline"
    $UserOnlineCheck = Read-Host " Press the Enter key if OK or type an Exchange Online Email address and press the Enter key"
    if (![string]::IsNullOrWhitespace($UserOnlineCheck)) {
        $script:UserOnline = $UserOnlineCheck
    }
}

Function ExchangeOnlineDomainCheck {
    #$ExchangeOnlineDomain
    Write-Host -foregroundcolor Green " Exchange Online Domain: $ExchangeOnlineDomain"
    $ExchangeOnlineDomaincheck = $RH = Read-Host " Press Enter if OK or type in the Exchange Online Domain and press the Enter key."
    if (![string]::IsNullOrWhitespace($ExchangeOnlineDomaincheck)) {
        $script:ExchangeOnlineDomain = $ExchangeOnlineDomainCheck
    }
}

Function UseronpremCheck {
    Write-Host -foregroundcolor Green " On Premises Hybrid Mailbox: $Useronprem"
    $Useronpremcheck = $RH = Read-Host " Press Enter if OK or type in an Exchange OnPremises Hybrid email address and press the Enter key."
    if (![string]::IsNullOrWhitespace($Useronpremcheck)) {
        $script:Useronprem = $Useronpremcheck
    }
}

Function ExchangeOnPremDomainCheck {
    #$exchangeOnPremDomain
    Write-Host -foregroundcolor Green " On Premises Mail Domain: $exchangeOnPremDomain"
    $exchangeOnPremDomaincheck = $RH = Read-Host " Press Enter if OK or type in the Exchange On Premises Mail Domain and press the Enter key."
    if (![string]::IsNullOrWhitespace($exchangeOnPremDomaincheck)) {
        $script:exchangeOnPremDomain = $exchangeOnPremDomaincheck
    }
}

Function ExchangeOnPremEWSCheck {
    Write-Host -foregroundcolor Green " On Premises EWS External URL: $exchangeOnPremEWS"
    $exchangeOnPremEWScheck = $RH = Read-Host " Press Enter if OK or type in the Exchange On Premises EWS URL and press the Enter key."
    if (![string]::IsNullOrWhitespace($exchangeOnPremEWScheck)) {
        $exchangeOnPremEWS = $exchangeOnPremEWScheck
    }
}

Function ExchangeOnPremLocalDomainCheck {
    Write-Host -foregroundcolor Green " On Premises Root Domain: $exchangeOnPremLocalDomain  "
    if ([string]::IsNullOrWhitespace($exchangeOnPremLocalDomain)) {
        $exchangeOnPremLocalDomaincheck = Read-Host "Please type in the Active directory Root domain.
        Press Enter to use $exchangeOnPremDomain"
        if ([string]::IsNullOrWhitespace($ADDomain)) {
            $exchangeOnPremLocalDomain = $exchangeOnPremDomain
        }
        if ([string]::IsNullOrWhitespace($exchangeOnPremLocalDomain)) {
            $exchangeOnPremLocalDomain = $exchangeOnPremLocalDomaincheck
        }
    }
}

#endregion

#region Show Parameters
Function ShowParameters {
    #if (($exchangeOnlineDomain -ne $null) -And  ($exchangeOnPremDomain -ne $null) -And  ($exchangeOnlineDomain -ne $null) -And  ($exchangeOnPremEWS -ne $null) -And  ($useronline -ne $null) -And  ($useronprem -ne $null) ){
    Write-Host $bar
    write-host -foregroundcolor Green "  Loading modules for AD, Exchange"
    Write-Host $bar
    Write-Host   "  Color Scheme"
    Write-Host $bar
    Write-Host -ForegroundColor Red "  Look out for Red!"
    Write-Host -ForegroundColor Yellow "  Yellow - Example information or Links"
    Write-Host -ForegroundColor Green "  Green - In Summary Sections it means OK. Anywhere else it's just a visual aid."
    Write-Host $bar
    Write-Host   "  Parameters:"
    Write-Host $bar
    Write-Host  -ForegroundColor White " Log File Path:"
    Write-Host -foregroundcolor Green "  $PSScriptRoot\$Logfile"
    Write-Host  -ForegroundColor White " Office 365 Domain:"
    Write-Host -foregroundcolor Green "  $exchangeOnlineDomain"
    Write-Host  -ForegroundColor White " AD root Domain"
    Write-Host -foregroundcolor Green "  $exchangeOnPremLocalDomain"
    Write-Host -foregroundcolor White " Exchange On Premises Domain:  "
    Write-Host -foregroundcolor Green "  $exchangeOnPremDomain"
    Write-Host -ForegroundColor White " Exchange On Premises External EWS url:"
    Write-Host -foregroundcolor Green "  $exchangeOnPremEWS"
    Write-Host -ForegroundColor White " On Premises Hybrid Mailbox:"
    Write-Host -foregroundcolor Green "  $useronprem"
    Write-Host -ForegroundColor White " Exchange Online Mailbox:"
    Write-Host -foregroundcolor Green "  $userOnline"


    $script:html = "<!DOCTYPE html>
        <html>
        <head>
	        <title>Hybrid Free Busy Configration Checker</title>
	        <style>
              body {
               font-family: Courier;
               background-color: white;
              }
              table, th {
              border: 1px solid black;
              border-collapse: collapse;
              padding: 5px;
              font-family: Courier;
              background-color: white;
              table-layout: fixed;
            }
             td {
              border: 1px solid black;
              border-collapse: collapse;
              padding: 5px;
              font-family: Courier;
              background-color: white;
              width: 50%;
              max-width: 50%;
              word-wrap:break-word
            }
            th {
              background-color: blue;
              text-align: left;
            }
		        .green { color: green; }
		        .red { color: red; }
		        .yellow { color: yellow; }
		        .white { color: white; }
                .black { color: white; }
                .orange { color: orange; }
                
	        </style>
        </head>
        <body>
            <div class='Black'><h1>Hybrid Free Busy Configration Checker</h1></div>
	        
	        <div class='Black'><h2><b>Parameters:</b></h2></div>
	       
                    <div class='Black'><b>Log File Path:</b><span style='color:green;'> $PSScriptRoot\$Logfile</span></div>
	                <div class='Black'><b>Office 365 Domain:</b><span style='color:green;'>  $exchangeOnlineDomain</span></div>
                    <div class='Black'><b>AD root Domain: </b><span style='color:green;'>  $exchangeOnPremLocalDomain</span></div>
	                <div class='Black'><b>Exchange On Premises Domain: </b><span style='color:green;'>   $exchangeOnPremDomain</span></div>
	                <div class='Black'><b>Exchange On Premises External EWS url: </b><span style='color:green;'> $exchangeOnPremEWS</span></div>
	                <div class='Black'><b>On Premises Hybrid Mailbox: </b><span style='color:green;'>  $useronprem</span></div>
	                <div class='Black'><b>Exchange Online Mailbox: </b><span style='color:green;'>  $userOnline</span></div>
	            

            <div class='Black'><p></p></div>
            <div class='Black'><h2><b>Color Scheme:</b></h2></div>
	        <div class='red'><b>Look out for Red!</b></div>
	        <div class='orange'><b>Orange - Example information or Links</b></div>
	        <div class='green'><b>Green - Means OK</b></div>
            
             <div class='Black'><p></p></div>
            
            
            <div class='Black'><h2>Hybrid Free Busy Configuration:</h2></div>
            
           "
           #(get-date -format yyyyMMdd_HHmmss)
    $html | Out-File -FilePath $htmlfile
}
#}
#endregion

#region DAuth Functions

Function OrgRelCheck {
    Write-Host $bar
    Write-Host -foregroundcolor Green " Get-OrganizationRelationship  | Where{($_.DomainNames -like $ExchangeOnlineDomain )} | Select Identity,DomainNames,FreeBusy*,Target*,Enabled, ArchiveAccessEnabled"
    Write-Host $bar
    $OrgRel
    Write-Host $bar
    Write-Host  -foregroundcolor Green " Summary - Get-OrganizationRelationship"
    Write-Host $bar
    #$exchangeonlinedomain
    Write-Host  -foregroundcolor White   " Domain Names:"
    if ($orgrel.DomainNames -like $exchangeonlinedomain) {
        Write-Host -foregroundcolor Green "  Domain Names Include the $exchangeOnlineDomain Domain"
        $tdDomainNames = "Domain Names Include the $exchangeOnlineDomain Domain"
        $tdDomainNamesColor = "green"
        $tdDomainNamesfl = $tdDomainNames | Format-List
    }
    else {
        Write-Host -foregroundcolor Red "  Domain Names do Not Include the $exchangeOnlineDomain Domain"
        $tdDomainNames = "Domain Names do Not Include the $exchangeOnlineDomain Domain"
        $tdDomainNamesColor = "Red"
    }
    #FreeBusyAccessEnabled
    Write-Host -foregroundcolor White   " FreeBusyAccessEnabled:"
    if ($OrgRel.FreeBusyAccessEnabled -like "True" ) {
        Write-Host -foregroundcolor Green "  FreeBusyAccessEnabled is set to True"
        $tdFBAccessEnabled = "FreeBusyAccessEnabled is set to True"
        $tdFBAccessEnabledColor = "green"
    }
    else {
        Write-Host -foregroundcolor Red "  FreeBusyAccessEnabled : False"
        $tdFBAccessEnabled = "FreeBusyAccessEnabled is set to False"
        $tdFBAccessEnabledColor = "red"
        $countOrgRelIssues++
    }
    #FreeBusyAccessLevel
    Write-Host -foregroundcolor White   " FreeBusyAccessLevel:"
    if ($OrgRel.FreeBusyAccessLevel -like "AvailabilityOnly" ) {
        Write-Host -foregroundcolor Green "  FreeBusyAccessLevel is set to AvailabilityOnly"
        $tdFBAccessLevel = "FreeBusyAccessLevel is set to AvailabilityOnly"
        $tdFBAccessLevelColor = "green"
    }
    if ($OrgRel.FreeBusyAccessLevel -like "LimitedDetails" ) {
        Write-Host -foregroundcolor Green "  FreeBusyAccessLevel is set to LimitedDetails"
        $tdFBAccessLevel = "FreeBusyAccessLevel is set to  LimitedDetails"
        $tdFBAccessLevelColor = "green"
    }
    if ($OrgRel.FreeBusyAccessLevel -ne "LimitedDetails" -AND $OrgRel.FreeBusyAccessLevel -ne "AvailabilityOnly" ) {
        Write-Host -foregroundcolor Red "  FreeBusyAccessEnabled : False"
        $tdFBAccessLevel = "FreeBusyAccessEnabled : False"
        $tdFBAccessLevelColor = "Red"
        $countOrgRelIssues++
    }
    #TargetApplicationUri
    Write-Host -foregroundcolor White   " TargetApplicationUri:"
    if ($OrgRel.TargetApplicationUri -like "Outlook.com" ) {
        Write-Host -foregroundcolor Green "  TargetApplicationUri is Outlook.com"
        $tdTargetApplicationUri = "TargetApplicationUri is Outlook.com"
        $tdTargetApplicationUriColor = "green"
    }
    else {
        Write-Host -foregroundcolor Red "  TargetApplicationUri should be Outlook.com"
        $tdTargetApplicationUri = "TargetApplicationUri should be Outlook.com"
        $tdTargetApplicationUriColor = "red"
        $countOrgRelIssues++
    }
    #TargetOwaURL
    Write-Host -foregroundcolor White   " TargetOwaURL:"
    if ($OrgRel.TargetOwaURL -like "https://outlook.com/owa/$exchangeonlinedomain" -or $OrgRel.TargetOwaURL -like $Null) {
        if ($OrgRel.TargetOwaURL -like "http://outlook.com/owa/$exchangeonlinedomain") {
            Write-Host -foregroundcolor Green "  TargetOwaURL is http://outlook.com/owa/$exchangeonlinedomain. This is a possible standard value. TargetOwaURL can also be configured to be Blank."
            $tdOrgRelTargetOwaURL = " $($OrgRel.TargetOwaURL) - TargetOwaURL is http://outlook.com/owa/$exchangeonlinedomain. This is a possible standard value. TargetOwaURL can also be configured to be Blank."
            $tdOrgRelTargetOwaURLColor = "green"
        }
        if ($OrgRel.TargetOwaURL -like "https://outlook.office.com/mail") {
            Write-Host -foregroundcolor Green "  TargetOwaURL is https://outlook.office.com/mail. This is a possible standard value. TargetOwaURL can also be configured to be Blank or http://outlook.com/owa/$exchangeonlinedomain."
            $tdOrgRelTargetOwaURL = " $($OrgRel.TargetOwaURL) - TargetOwaURL is https://outlook.office.com/mail. TargetOwaURL can also be configured to be Blank or http://outlook.com/owa/$exchangeonlinedomain."
            $tdOrgRelTargetOwaURLColor = "green"
        }
        if ($OrgRel.TargetOwaURL -like $Null) {
            Write-Host -foregroundcolor Green "  TargetOwaURL is Blank, this is a standard value. "
            Write-Host  "  TargetOwaURL can also be configured to be https://outlook.com/owa/$exchangeonlinedomain or https://outlook.office.com/mail"
            $tdOrgRelTargetOwaURL = "$($OrgRel.TargetOwaURL) . TargetOwaURL is Blank, this is a standard value. TargetOwaURL can also be configured to be http://outlook.com/owa/$exchangeonlinedomain or http://outlook.office.com/mail. "
            $tdOrgRelTargetOwaURLColor = "green"
            if ($OrgRel.TargetOwaURL -like "https://outlook.com/owa/$exchangeonlinedomain") {
                Write-Host -foregroundcolor Green "  TargetOwaURL is https://outlook.com/owa/$exchangeonlinedomain. This is a possible standard value. TargetOwaURL can also be configured to be Blank or http://outlook.office.com/mail."
                $tdOrgRelTargetOwaURL = " $($OrgRel.TargetOwaURL) - TargetOwaURL is https://outlook.com/owa/$exchangeonlinedomain. This is a possible standard value. TargetOwaURL can also be configured to be Blank or http://outlook.office.com/mail."
                $tdOrgRelTargetOwaURLColor = "green"
            }
        
        }
    }
    else {
        Write-Host -foregroundcolor Red "  TargetOwaURL seems not to be Blank or https://outlook.com/owa/$exchangeonlinedomain. These are the standard values."
        $countOrgRelIssues++
        $tdOrgRelTargetOwaURL = "  TargetOwaURL seems not to be Blank or https://outlook.com/owa/$exchangeonlinedomain. These are the standard values."
        $tdOrgRelTargetOwaURLColor = "red"
    
    }
    #TargetSharingEpr
    Write-Host -foregroundcolor White   " TargetSharingEpr:"
    if ([string]::IsNullOrWhitespace($OrgRel.TargetSharingEpr) -or $OrgRel.TargetSharingEpr -eq "https://outlook.office365.com/EWS/Exchange.asmx ") {
        Write-Host -foregroundcolor Green "  TargetSharingEpr is ideally blank. this is the standard Value. "
        Write-Host  "  If it is set, it should be Office 365 EWS endpoint. Example: https://outlook.office365.com/EWS/Exchange.asmx "
        $tdTargetSharingEpr = "  TargetSharingEpr is ideally blank. this is the standard Value.
        If it is set, it should be Office 365 EWS endpoint. Example: https://outlook.office365.com/EWS/Exchange.asmx "
        $tdTargetSharingEprColor = "green"
    }
    else {
        Write-Host -foregroundcolor Red "  TargetSharingEpr should be blank or  https://outlook.office365.com/EWS/Exchange.asmx"
        Write-Host  "  If it is set, it should be Office 365 EWS endpoint.  Example: https://outlook.office365.com/EWS/Exchange.asmx "
        $tdTargetSharingEpr = "  TargetSharingEpr should be blank or  https://outlook.office365.com/EWS/Exchange.asmx
        If it is set, it should be Office 365 EWS endpoint.  Example: https://outlook.office365.com/EWS/Exchange.asmx "
        $tdTargetSharingEprColor = "red"
        $countOrgRelIssues++
    }
    #FreeBusyAccessScope
    Write-Host -ForegroundColor White  " FreeBusyAccessScope:"
    if ([string]::IsNullOrWhitespace($OrgRel.FreeBusyAccessScope)) {
        Write-Host -foregroundcolor Green "  FreeBusyAccessScope is blank, this is the standard Value. "
        $tdFreeBusyAccessScope = " FreeBusyAccessScope is blank, this is the standard Value."
        $tdFreeBusyAccessScopeColor = "green"
    }
    else {
        Write-Host -foregroundcolor Red "  FreeBusyAccessScope is should be Blank, that is the standard Value."
        $tdFreeBusyAccessScope = " FreeBusyAccessScope is should be Blank, that is the standard Value."
        $tdFreeBusyAccessScopeColor = "red"
        $countOrgRelIssues++
    }
    #TargetAutodiscoverEpr:
    $OrgRelTargetAutodiscoverEpr = $OrgRel.TargetAutodiscoverEpr
    If ([string]::IsNullOrWhitespace($OrgRelTargetAutodiscoverEpr))
        {
            $OrgRelTargetAutodiscoverEpr = "Blank" 
        }    
    Write-Host -foregroundcolor White   " TargetAutodiscoverEpr:"
    if ($OrgRel.TargetAutodiscoverEpr -like "https://autodiscover-s.outlook.com/autodiscover/autodiscover.svc/WSSecurity" ) {
        Write-Host -foregroundcolor Green "  TargetAutodiscoverEpr is correct"
        $tdTargetAutodiscoverEPR = " TargetAutodiscoverEpr is https://autodiscover-s.outlook.com/autodiscover/autodiscover.svc/WSSecurity"
        $tdTargetAutodiscoverEPRColor = "green"
    }
    else {
        Write-Host -foregroundcolor Red "  TargetAutodiscoverEpr is not correct. Should be https://autodiscover-s.outlook.com/autodiscover/autodiscover.svc/WSSecurity"
        $tdTargetAutodiscoverEPR = " TargetAutodiscoverEpr is $OrgRelTargetAutodiscoverEpr . Should be https://autodiscover-s.outlook.com/autodiscover/autodiscover.svc/WSSecurity"
        $tdTargetAutodiscoverEPRColor = "Red"
        $countOrgRelIssues++
    }
    #Enabled
    Write-Host -foregroundcolor White   " Enabled:"
    if ($OrgRel.enabled -like "True" ) {
        Write-Host -foregroundcolor Green "  Enabled is set to True"

        $tdFreeBusyEnabled = "$($OrgRel.enabled)"
        $tdFreeBusyEnabledColor = "green"


    }
    else {
        Write-Host -foregroundcolor Red "  Enabled is set to False."
        $countOrgRelIssues++

        $tdFreeBusyEnabled = "$($OrgRel.enabled) - Should be True."
        $tdFreeBusyEnabledColor = "red"
    }
    #if ($countOrgRelIssues -eq '0'){
    #Write-Host -foregroundcolor Green " Configurations Seem Correct"
    #}
    #else
    #{
    #Write-Host -foregroundcolor Red "  Configurations DO NOT Seem Correct"
    #}
    $OrgRelDomainNames = ""
    $OrgRelDomainNames = ""
    foreach ($domain in $OrgRel.DomainNames.Domain) {
        if ($OrgRelDomainNames -ne "") {
            $OrgRelDomainNames += "; "
        }
        $OrgRelDomainNames += $domain
    }
    $FreeBusyAccessEnabled = $OrgRel.FreeBusyAccessEnabled
    $FreeBusyAccessLevel = $OrgRel.FreeBusyAccessLevel
    $tdTargetOwaUrl = $OrgRel.TargetOwaURL
    $tdEnabled = $OrgRel.Enabled
    $script:html += "

     <div class='Black'><p></p></div>

             <div class='Black'><h2><b>`n Exchange On Premise Free Busy Configuration: `n</b></h2></div>

             <div class='Black'><p></p></div>

         <table style='width:100%'>
    <tr>
    <th colspan='2' style='text-align:center; color:white;'><b>Exchange On Premise DAuth Configuration</b></th>
    </tr>   
    <tr>
    <th colspan='2' style='color:white;'>Summary - Get-OrganizationRelationship</th>
    </tr>
    <tr>
    <td><b>Get-OrganizationRelationship</b></td>
    <td>
        <div> <b>Domain Names: </b> <span style='color:$tdDomainNamesColor'>$tdDomainNames</span></div> 
        <div> <b>FreeBusyAccessEnabled: </b> <span style='color:$tdFBAccessEnabledColor'>$tdFBAccessEnabled</span></div> 
        <div> <b>FreeBusyAccessLevel: </b> <span style='color:$tdFBAccessLevelColor'>$tdFBAccessLevel</span></div> 
        <div> <b>TargetApplicationUri: </b> <span style='color:$tdTargetApplicationUriColor'>$tdTargetApplicationUri</span></div> 
        <div> <b>TargetAutodiscoverEPR: </b> <span style='color:$tdTargetAutodiscoverEPRColor'>$tdTargetAutodiscoverEPR</span></div> 
        <div> <b>TargetOwaURL: </b> <span style='color:$tdOrgRelTargetOwaURLColor'>$tdOrgRelTargetOwaURL</span></div> 
        <div> <b>TargetSharingEpr: </b> <span style='color:$tdTargetSharingEprColor'>$tdTargetSharingEpr</span></div> 
        <div> <b>FreeBusyAccessScope: </b> <span style='color:$tdFreeBusyAccessScopeColor'>$tdFreeBusyAccessScope</span></div> 
        <div> <b>Enabled:</b> <span style='color:$tdFreeBusyEnabledColor'>$tdFreeBusyEnabled</span></div> 
    </td>

   
 </tr>
  "
    $html | Out-File -FilePath $htmlfile
    Write-Host -foregroundcolor Yellow "`n  Reference: https://learn.microsoft.com/en-us/exchange/create-an-organization-relationship-exchange-2013-help"
}

Function FedInfoCheck {
    Write-Host -foregroundcolor Green " Get-FederationInformation -DomainName $exchangeOnlineDomain  -BypassAdditionalDomainValidation | fl"
    Write-Host $bar
    $fedinfo = get-federationInformation -DomainName $exchangeOnlineDomain  -BypassAdditionalDomainValidation -ErrorAction SilentlyContinue | Select-Object *
    if (!$fedinfo) {
        $fedinfo = get-federationInformation -DomainName $exchangeOnlineDomain  -BypassAdditionalDomainValidation -ErrorAction SilentlyContinue | Select-Object *
    }
    
    $fedinfo
    
   
    Write-Host $bar
    Write-Host -foregroundcolor Green " Summary - Federation Information"
    Write-Host $bar
    #DomainNames
    Write-Host -foregroundcolor White   "  Domain Names: "
    if ($fedinfo.DomainNames -like "*$ExchangeOnlineDomain*") {
        Write-Host -foregroundcolor Green "   Domain Names include the Exchange Online Domain "$ExchangeOnlineDomain
        $tdDomainNamesColor = "green"
        $tdDomainNamesfl = "Domain Names include the Exchange Online Domain $ExchangeOnlineDomain"
    }
    else {
        Write-Host -foregroundcolor Red "   Domain Names seem not to include the Exchange Online Domain "$ExchangeOnlineDomain
        Write-Host  "   Domain Names: "$fedinfo.DomainNames
        $tdDomainNamesColor = "Red"
        $tdDomainNamesfl = "Domain Names seem not to include the Exchange Online Domain: $ExchangeOnlineDomain"
    }
    #TokenIssuerUris
    Write-Host  -foregroundcolor White  "  TokenIssuerUris: "
    if ($fedinfo.TokenIssuerUris -like "*urn:federation:MicrosoftOnline*") {
        Write-Host -foregroundcolor Green "  "  $fedinfo.TokenIssuerUris
        $tdTokenIssuerUrisColor = "green"
        $tdTokenIssuerUrisFL = $fedinfo.TokenIssuerUris
    }
    else {
        Write-Host "   " $fedinfo.TokenIssuerUris
        Write-Host  -foregroundcolor Red "   TokenIssuerUris should be urn:federation:MicrosoftOnline"
        $tdTokenIssuerUrisColor = "red"
        $tdTokenIssuerUrisFL = "   TokenIssuerUris should be urn:federation:MicrosoftOnline"
    }
    #TargetApplicationUri
    Write-Host -foregroundcolor White   "  TargetApplicationUri:"
    if ($fedinfo.TargetApplicationUri -like "Outlook.com") {
        Write-Host -foregroundcolor Green "  "$fedinfo.TargetApplicationUri
        $tdTargetApplicationUriColor = "green"
        $tdTargetApplicationUriFL = $fedinfo.TargetApplicationUri
    }
    else {
        Write-Host -foregroundcolor Red "   "$fedinfo.TargetApplicationUri
        Write-Host -foregroundcolor Red   "   TargetApplicationUri should be Outlook.com"
        $tdTargetApplicationUriColor = "red"
        $tdTargetApplicationUriFL = "   TargetApplicationUri should be Outlook.com"
    }
    #TargetAutodiscoverEpr
    Write-Host -foregroundcolor White   "  TargetAutodiscoverEpr:"
    if ($fedinfo.TargetAutodiscoverEpr -like "https://autodiscover-s.outlook.com/autodiscover/autodiscover.svc/WSSecurity") {
        Write-Host -foregroundcolor Green "   "$fedinfo.TargetAutodiscoverEpr
        $tdTargetAutodiscoverEprColor = "green"
        $tdTargetAutodiscoverEprFL = $fedinfo.TargetAutodiscoverEpr
    }
    else {
        Write-Host -foregroundcolor Red "   "$fedinfo.TargetAutodiscoverEpr
        Write-Host -foregroundcolor Red   " TargetAutodiscoverEpr should be https://autodiscover-s.outlook.com/autodiscover/autodiscover.svc/WSSecurity"
        $tdTargetAutodiscoverEprColor = "red"
        $tdTargetAutodiscoverEprFL = "   TargetAutodiscoverEpr should be https://autodiscover-s.outlook.com/autodiscover/autodiscover.svc/WSSecurity"
    }
    # Federation Information TargetApplicationUri vs Organization Relationship TargetApplicationUri
    Write-Host -ForegroundColor White "  Federation Information TargetApplicationUri vs Organization Relationship TargetApplicationUri "
    if ($fedinfo.TargetApplicationUri -like "Outlook.com") {
        if ($OrgRel.TargetApplicationUri -like $fedinfo.TargetApplicationUri) {
            Write-Host -foregroundcolor Green "   => Federation Information TargetApplicationUri matches the Organization Relationship TargetApplicationUri "
            Write-Host  "       Organization Relationship TargetApplicationUri:"  $OrgRel.TargetApplicationUri
            Write-Host  "       Federation Information TargetApplicationUri:   "  $fedinfo.TargetApplicationUri
            $tdFederationInformationTAColor = "green"
            $tdFederationInformationTAFL = " => Federation Information TargetApplicationUri matches the Organization Relationship TargetApplicationUri"
        }
        else {
            Write-Host -foregroundcolor Red "   => Federation Information TargetApplicationUri should be Outlook.com and match the Organization Relationship TargetApplicationUri "
            Write-Host  "       Organization Relationship TargetApplicationUri:"  $OrgRel.TargetApplicationUri
            Write-Host  "       Federation Information TargetApplicationUri:   "  $fedinfo.TargetApplicationUri
            $tdFederationInformationTAColor = "red"
            $tdFederationInformationTAFL = " => Federation Information TargetApplicationUri should be Outlook.com and match the Organization Relationship TargetApplicationUri"
        }
    }
    #TargetAutodiscoverEpr vs Organization Relationship TargetAutodiscoverEpr
    Write-Host -ForegroundColor White  "  Federation Information TargetAutodiscoverEpr vs Organization Relationship TargetAutodiscoverEpr "
    if ($OrgRel.TargetAutodiscoverEpr -like $fedinfo.TargetAutodiscoverEpr) {
        Write-Host -foregroundcolor Green "   => Federation Information TargetAutodiscoverEpr matches the Organization Relationship TargetAutodiscoverEpr "
        Write-Host  "       Organization Relationship TargetAutodiscoverEpr:"  $OrgRel.TargetAutodiscoverEpr
        Write-Host  "       Federation Information TargetAutodiscoverEpr:   "  $fedinfo.TargetAutodiscoverEpr
        $tdTargetAutodiscoverEprVSColor = "green"
        $tdTargetAutodiscoverEprVSFL = "=> Federation Information TargetAutodiscoverEpr matches the Organization Relationship TargetAutodiscoverEpr"
    }
    else {
        Write-Host -foregroundcolor Red "   => Federation Information TargetAutodiscoverEpr should match the Organization Relationship TargetAutodiscoverEpr"
        Write-Host  "       Organization Relationship TargetAutodiscoverEpr:"  $OrgRel.TargetAutodiscoverEpr
        Write-Host  "       Federation Information TargetAutodiscoverEpr:   "  $fedinfo.TargetAutodiscoverEpr
        $tdTargetAutodiscoverEprVSColor = "red"
        $tdTargetAutodiscoverEprVSFL = "=> Federation Information TargetAutodiscoverEpr should match the Organization Relationship TargetAutodiscoverEpr"
    }
    Write-Host -foregroundcolor Yellow "`n  Reference: https://learn.microsoft.com/en-us/exchange/configure-a-federation-trust-exchange-2013-help#what-do-you-need-to-know-before-you-begin"
    Write-Host $bar
    $FedInfoDomainNames = ""
    $FedInfoDomainNames = ""
    foreach ($domain in $FedInfo.DomainNames.Domain) {
        if ($FedInfoDomainNames -ne "") {
            $FedInfoDomainNames += "; "
        }
        $FedInfoDomainNames += $domain
    }
    $aux = $FedInfo.DomainNames
    $fedinfoTokenIssuerUris = $fedinfo.TokenIssuerUris
    $fedinfoTargetAutodiscoverEpr = $fedinfo.TargetAutodiscoverEpr
    $fedinfoTargetApplicationUri = $fedinfo.TargetApplicationUri



    $script:html += "
       
    <tr>
    <th colspan='2' style='color:white;'>Summary - Get-FederationInformation</th>
    </tr>
    <tr>
    <td><b>Get-FederationInformation -Domain $ExchangeOnPremDomain</b></td>
    <td>
        <div> <b>Domain Names: </b> <span style='color:$tdDomainNamesColor'>$tdDomainNamesFL</span></div> 
        <div> <b>TokenIssuerUris: </b> <span style='color:$tdTokenIssuerUrisColor'>$tdTokenIssuerUrisFL</span></div> 
        <div> <b>TargetApplicationUri: </b> <span style='color:$tdTargetApplicationUriColor'>$tdTargetApplicationUriFL</span></div> 
        <div> <b>TargetAutodiscoverEpr: </b> <span style='color:$tdTargetAutodiscoverEprColor'>$tdTargetAutodiscoverEprFL</span></div> 
        <div> <b>TargetApplicationUri - Federation Information vs Organization Relationship: </b> <span style='color:$tdTargetAutodiscoverEprVSColor'>$tdFederationInformationTAFL</span></div> 
        <div> <b>TargetAutodiscoverEpr - Federation Information vs Organization Relationship:</b> <span style='color:$tdTargetAutodiscoverEprVSColor'>$tdTargetAutodiscoverEprVSFL</span></div> 
         
    </td>

   
 </tr>
  "

    $html | Out-File -FilePath $htmlfile
}

Function FedTrustCheck {
    Write-Host -foregroundcolor Green " Get-FederationTrust | fl ApplicationUri,TokenIssuerUri,OrgCertificate,TokenIssuerCertificate,
    TokenIssuerPrevCertificate, TokenIssuerMetadataEpr,TokenIssuerEpr"
    Write-Host $bar
    $Global:fedtrust = Get-FederationTrust | Select-Object ApplicationUri, TokenIssuerUri, OrgCertificate, TokenIssuerCertificate, TokenIssuerPrevCertificate, TokenIssuerMetadataEpr, TokenIssuerEpr
    $fedtrust
    Write-Host $bar
    Write-Host -foregroundcolor Green " Summary - Federation Trust"
    Write-Host $bar
    $CurrentTime = get-date
    Write-Host -foregroundcolor White " Federation Trust Aplication Uri:"
    if ($fedtrust.ApplicationUri -like "FYDIBOHF25SPDLT.$ExchangeOnpremDomain") {
        Write-Host -foregroundcolor Green " " $fedtrust.ApplicationUri
        $tdfedtrustApplicationUriColor = "green"
        $tdfedtrustApplicationUriFL = $fedtrust.ApplicationUri
    }
    else {
        Write-Host -foregroundcolor Red "  Federation Trust Aplication Uri Should be "$fedtrust.ApplicationUri
        $tdfedtrustApplicationUriColor = "red"
        $tdfedtrustApplicationUriFL = "  Federation Trust Aplication Uri Should be $fedtrust.ApplicationUri"

    }
    #$fedtrust.TokenIssuerUri.AbsoluteUri
    Write-Host -foregroundcolor White " TokenIssuerUri:"
    if ($fedtrust.TokenIssuerUri.AbsoluteUri -like "urn:federation:MicrosoftOnline") {
        #Write-Host -foregroundcolor White "  TokenIssuerUri:"
        Write-Host -foregroundcolor Green " "$fedtrust.TokenIssuerUri.AbsoluteUri
        $tdfedtrustTokenIssuerUriColor = "green"
        $tdfedtrustTokenIssuerUriFL = $fedtrust.TokenIssuerUri.AbsoluteUri
    }
    else {
        Write-Host -foregroundcolor Red " Federation Trust TokenIssuerUri should be urn:federation:MicrosoftOnline"
        $tdfedtrustTokenIssuerUriColor = "red"
        $tdfedtrustTokenIssuerFL = " Federation Trust TokenIssuerUri is currently $fedtrust.TokenIssuerUri.AbsoluteUri but should be urn:federation:MicrosoftOnline"
    }
    Write-Host -foregroundcolor White " Federation Trust Certificate Expiracy:"
    if ($fedtrust.OrgCertificate.NotAfter.Date -gt $CurrentTime) {
        Write-Host -foregroundcolor Green "  Not Expired"
        Write-Host  "   - Expires on " $fedtrust.OrgCertificate.NotAfter.DateTime
        $tdfedtrustOrgCertificateNotAfterDateColor = "green"
        $tdfedtrustOrgCertificateNotAfterDateFL = $fedtrust.OrgCertificate.NotAfter.DateTime
    }
    else {
        Write-Host -foregroundcolor Red " Federation Trust Certificate is Expired on " $fedtrust.OrgCertificate.NotAfter.DateTime
        $tdfedtrustOrgCertificateNotAfterDateColor = "red"
        $tdfedtrustOrgCertificateNotAfterDateFL = $fedtrust.OrgCertificate.NotAfter.DateTime
    }
    Write-Host -foregroundcolor White " `Federation Trust Token Issuer Certificate Expiracy:"
    if ($fedtrust.TokenIssuerCertificate.NotAfter.DateTime -gt $CurrentTime) {
        Write-Host -foregroundcolor Green "  Not Expired"
        Write-Host  "   - Expires on " $fedtrust.TokenIssuerCertificate.NotAfter.DateTime
        $tdfedtrustTokenIssuerCertificateNotAfterDateTimeColor = "green"
        $tdfedtrustTokenIssuerCertificateNotAfterDateTimeFL = $fedtrust.TokenIssuerCertificate.NotAfter.DateTime
    }
    else {
        Write-Host -foregroundcolor Red "  Federation Trust TokenIssuerCertificate Expired on " $fedtrust.TokenIssuerCertificate.NotAfter.DateTime
        $tdfedtrustTokenIssuerCertificateNotAfterDateTimeColor = "red"
        $tdfedtrustTokenIssuerCertificateNotAfterDateTimeFL = $fedtrust.TokenIssuerCertificate.NotAfter.DateTime
    }
    #Write-Host -foregroundcolor White " Federation Trust Token Issuer Prev Certificate Expiracy:"
    #if ($fedtrust.TokenIssuerPrevCertificate.NotAfter.Date -gt $CurrentTime) {
    #Write-Host -foregroundcolor Green "  Not Expired"
    #Write-Host  "   - Expires on " $fedtrust.TokenIssuerPrevCertificate.NotAfter.DateTime
    #$tdfedtrustTokenIssuerPrevCertificateNotAfterDateColor = "green"
    #$tdfedtrustTokenIssuerPrevCertificateNotAfterDateFL = $fedtrust.TokenIssuerPrevCertificate.NotAfter.DateTime
    #}
    #else {
    #Write-Host -foregroundcolor Red "  Federation Trust TokenIssuerPrevCertificate Expired on " $fedtrust.TokenIssuerPrevCertificate.NotAfter.DateTime
    #$tdfedtrustTokenIssuerPrevCertificateNotAfterDateColor = "red"
    #$tdfedtrustTokenIssuerPrevCertificateNotAfterDateFL = $fedtrust.TokenIssuerPrevCertificate.NotAfter.DateTime
    #}
    $fedtrustTokenIssuerMetadataEpr = "https://nexus.microsoftonline-p.com/FederationMetadata/2006-12/FederationMetadata.xml"
    Write-Host -foregroundcolor White " `Token Issuer Metadata EPR:"
    if ($fedtrust.TokenIssuerMetadataEpr.AbsoluteUri -like $fedtrustTokenIssuerMetadataEpr) {
        Write-Host -foregroundcolor Green "  Token Issuer Metadata EPR is " $fedtrust.TokenIssuerMetadataEpr.AbsoluteUri
        #test if it can be reached
        $tdfedtrustTokenIssuerMetadataEprAbsoluteUriColor = "green"
        $tdfedtrustTokenIssuerMetadataEprAbsoluteUriFL = $fedtrust.TokenIssuerMetadataEpr.AbsoluteUri
    }
    else {
        Write-Host -foregroundcolor Red " Token Issuer Metadata EPR is Not " $fedtrust.TokenIssuerMetadataEpr.AbsoluteUri
        $tdfedtrustTokenIssuerMetadataEprAbsoluteUriColor = "red"
        $tdfedtrustTokenIssuerMetadataEprAbsoluteUriFL = $fedtrust.TokenIssuerMetadataEpr.AbsoluteUri
    }
    $fedtrustTokenIssuerEpr = "https://login.microsoftonline.com/extSTS.srf"
    Write-Host -foregroundcolor White " Token Issuer EPR:"
    if ($fedtrust.TokenIssuerEpr.AbsoluteUri -like $fedtrustTokenIssuerEpr) {
        Write-Host -foregroundcolor Green "  Token Issuer EPR is:" $fedtrust.TokenIssuerEpr.AbsoluteUri
        #test if it can be reached
        $tdfedtrustTokenIssuerEprAbsoluteUriColor = "green"
        $tdfedtrustTokenIssuerEprAbsoluteUriFL = $fedtrust.TokenIssuerEpr.AbsoluteUri
    }
    else {
        Write-Host -foregroundcolor Red "  Token Issuer EPR is Not:" $fedtrust.TokenIssuerEpr.AbsoluteUri
        $tdfedtrustTokenIssuerEprAbsoluteUriColor = "red"
        $tdfedtrustTokenIssuerEprAbsoluteUriFL = $fedtrust.TokenIssuerEpr.AbsoluteUri
    }
    $fedinfoTokenIssuerUris = $fedinfo.TokenIssuerUris
    $fedinfoTargetApplicationUri = $fedinfo.TargetApplicationUri
    $script:fedinfoTargetAutodiscoverEpr = $fedinfo.TargetAutodiscoverEpr
    
    
    
    $script:html += "
    <tr>
    <th colspan='2' style='color:white;'>Summary - Test-FederationTrust</th>
    </tr>
    <tr>
    <td><b>Get-FederationTrust | select ApplicationUri, TokenIssuerUri, OrgCertificate, TokenIssuerCertificate, TokenIssuerPrevCertificate, TokenIssuerMetadataEpr, TokenIssuerEpr</b></td>
    <td>
        <div> <b>Application Uri: </b> <span style='color:$tdfedtrustApplicationUricolor'>$tdfedtrustApplicationUriFL</span></div> 
        <div> <b>TokenIssuerUris: </b> <span style='color:$tdfedtrustTokenIssuerUriColor'>$tdfedtrustTokenIssuerUriFL</span></div> 
        <div> <b>Certificate Expiracy: </b> <span style='color:$tdfedtrustOrgCertificateNotAfterDateColor'>$tdfedtrustOrgCertificateNotAfterDateFL</span></div>
        <div> <b>Token Issuer Certificate Expiracy: </b> <span style='color:$tdfedtrustTokenIssuerCertificateNotAfterDateTimeColor'>$tdfedtrustTokenIssuerCertificateNotAfterDateTimeFL</span></div> 
        <div> <b>Token Issuer Metadata EPR:</b> <span style='color:$tdfedtrustTokenIssuerMetadataEprAbsoluteUriColor'>$tdfedtrustTokenIssuerMetadataEprAbsoluteUriFL</span></div> 
        <div> <b>Token Issuer EPR: </b> <span style='color:$tdfedtrustTokenIssuerEprAbsoluteUriColor'>$tdfedtrustTokenIssuerEprAbsoluteUriFL</span></div> 
         
    </td>
</tr>
  "
    
    $html | Out-File -FilePath $htmlfile
    Write-Host -foregroundcolor Yellow "`n  Reference: https://learn.microsoft.com/en-us/exchange/configure-a-federation-trust-exchange-2013-help"
}

Function AutoDVirtualDCheck {
    
    Write-Host $bar
    Write-Host -foregroundcolor Green " Get-AutodiscoverVirtualDirectory | Select Identity,Name,ExchangeVersion,*authentication*"
    Write-Host $bar
    $Global:AutoDiscoveryVirtualDirectory = Get-AutodiscoverVirtualDirectory | Select-Object Identity, Name, ExchangeVersion, *authentication*
    #Check if null or set
    #$AutoDiscoveryVirtualDirectory
    $Global:AutoDiscoveryVirtualDirectory
    $AutoDFL = $Global:AutoDiscoveryVirtualDirectory | Format-List
    $script:html += "<tr>
    <th colspan='2' style='color:white;'>Summary - Get-AutodiscoverVirtualDirectory</th>
    </tr>
    <tr>
    <td><b>Get-AutodiscoverVirtualDirectory | Select Identity,Name,ExchangeVersion,*authentication*</b></td>
    <td>"
    Write-Host $bar
    Write-Host -foregroundcolor Green " Summary - On-Prem Get-AutodiscoverVirtualDirectory"
    Write-Host $bar
    Write-Host -ForegroundColor White "  WSSecurityAuthentication:"
    if ($Global:AutoDiscoveryVirtualDirectory.WSSecurityAuthentication -eq "True") {
        foreach ( $ser in $Global:AutoDiscoveryVirtualDirectory) {
            Write-Host " $($ser.Identity) "
            Write-Host -ForegroundColor Green "  WSSecurityAuthentication: $($ser.WSSecurityAuthentication)"

            $AutodVDIdentity = $ser.Identity
            $AutodVDName = $ser.Name
            $AutodVDInternalAuthenticationMethods = $ser.InternalAuthenticationMethods
            $AutodVDExternalAuthenticationMethods = $ser.ExternalAuthenticationMethods
            $AutodVDWSAuthetication = $ser.WSSecurityAuthentication
            $AutodVDWSAutheticationColor = "green"
            $AutodVDWindowsAuthentication = $ser.WindowsAuthentication
            if ($AutodVDWindowsAuthentication -eq "True") {
                $AutodVDWindowsAuthenticationColor = "green"
            }
            else {
                $AutodVDWindowsAuthenticationColor = "red"
            }
            $AutodVDInternalNblBypassUrl = $ser.InternalNblBypassUrl
            $AutodVDInternalUrl = $ser.InternalUrl
            $AutodVDExternalUrl = $ser.ExternalUrl
            $script:html +=
            " <div><b>============================</b></div>
            <div><b>Identity:</b> $AutodVDIdentity</div>
            <div><b>Name:</b> $AutodVDName </div>
            <div><b>InternalAuthenticationMethods:</b> $AutodVDInternalAuthenticationMethods </div>
            <div><b>ExternalAuthenticationMethods:</b> $AutodVDExternalAuthenticationMethods </div>
            <div><b>WSAuthetication:</b> <span style='color:green'>$AutodVDWSAuthetication</span></div>
            <div><b>WindowsAuthentication:</b> <span style='color:green'>$AutodVDWindowsAuthentication</span></div>
            "
            
            $serWSSecurityAuthenticationColor = "Green"
        }
    }
    else {
        Write-Host -foregroundcolor Red " WSSecurityAuthentication is NOT correct."
        foreach ( $ser in $Global:AutoDiscoveryVirtualDirectory) {
            Write-Host " $($ser.Identity)"
            Write-Host -ForegroundColor Red "  WSSecurityAuthentication: $($ser.WSSecurityAuthentication)"
            $serWSSecurityAuthenticationColor = "Red"
            Write-Host " $($ser.Identity) "
            $AutodVDIdentity = $ser.Identity
            $AutodVDName = $ser.Name
            $AutodVDInternalAuthenticationMethods = $ser.InternalAuthenticationMethods
            $AutodVDExternalAuthenticationMethods = $ser.ExternalAuthenticationMethods
            $AutodVDWSAuthetication = $ser.WSSecurityAuthentication
            $AutodVDWSAutheticationColor = "green"
            $AutodVDWindowsAuthentication = $ser.WindowsAuthentication
            if ($AutodVDWindowsAuthentication -eq "True") {
                $AutodVDWindowsAuthenticationColor = "green"
            }
            else {
                $AutodVDWindowsAuthenticationColor = "red"
            }
            $AutodVDInternalNblBypassUrl = $ser.InternalNblBypassUrl
            $AutodVDInternalUrl = $ser.InternalUrl
            $AutodVDExternalUrl = $ser.ExternalUrl
            $script:html +=
            " <div><b>============================</b></div>
            <div><b>Identity:</b> $AutodVDIdentity</div>
            <div><b>Name:</b> $AutodVDName </div>
            <div><b>InternalAuthenticationMethods:</b> $AutodVDInternalAuthenticationMethods </div>
            <div><b>ExternalAuthenticationMethods:</b> $AutodVDExternalAuthenticationMethods </div>
            <div><b>WSAuthetication:</b> <span style='color:red'>$AutodVDWSAuthetication</span></div>
            <div><b>WindowsAuthentication:</b> <span style='color:$AutodVDWindowsAuthenticationColor'>$AutodVDWindowsAuthentication</span></div>
            "
            Write-Host -ForegroundColor Green "  WSSecurityAuthentication: $($ser.WSSecurityAuthentication)"
            $serWSSecurityAuthenticationColor = "Red"
        }
        Write-Host -foregroundcolor White "  Should be True "
    }
    Write-Host -ForegroundColor White "`n  WindowsAuthentication:"
    if ($Global:AutoDiscoveryVirtualDirectory.WindowsAuthentication -eq "True") {
        foreach ( $ser in $Global:AutoDiscoveryVirtualDirectory) {
            Write-Host " $($ser.Identity) "
            Write-Host -ForegroundColor Green "  WindowsAuthentication: $($ser.WindowsAuthentication)"
        }
    }
    else {
        Write-Host -foregroundcolor Red " WindowsAuthentication is NOT correct."
        foreach ( $ser in $Global:AutoDiscoveryVirtualDirectory) {
            Write-Host " $($ser.Identity)"
            Write-Host -ForegroundColor Red "  WindowsAuthentication: $($ser.WindowsAuthentication)"
        }
        Write-Host -foregroundcolor White "  Should be True "
    }
    Write-Host -foregroundcolor Yellow "`n  Reference: https://learn.microsoft.com/en-us/powershell/module/exchange/get-autodiscovervirtualdirectory?view=exchange-ps"
    $html | Out-File -FilePath $htmlfile
}

Function EWSVirtualDirectoryCheck {
    Write-Host -foregroundcolor Green " Get-WebServicesVirtualDirectory | Select Identity,Name,ExchangeVersion,*Authentication*,*url"
    Write-Host $bar
    $Global:WebServicesVirtualDirectory = Get-WebServicesVirtualDirectory | Select-Object Identity, Name, ExchangeVersion, *Authentication*, *url
    $Global:WebServicesVirtualDirectory
    Write-Host $bar
    Write-Host -foregroundcolor Green " Summary - Get-WebServicesVirtualDirectory"
    Write-Host $bar
    $script:html += "
    <tr>
    <th colspan='2' style='color:white;'>Summary - Get-WebServicesVirtualDirectory</th>
    </tr>
    <tr>
    <td><b> Get-WebServicesVirtualDirectory | Select Identity,Name,ExchangeVersion,*Authentication*,*url</b></td>
    <td >"
    Write-Host -foregroundcolor White "  WSSecurityAuthentication:"
    if ($Global:WebServicesVirtualDirectory.WSSecurityAuthentication -like "True") {
        foreach ( $EWS in $Global:WebServicesVirtualDirectory) {
            Write-Host " $($EWS.Identity)"
            Write-Host -ForegroundColor Green "  WSSecurityAuthentication: $($EWS.WSSecurityAuthentication) "
            $EWSVDIdentity = $EWS.Identity
            $EWSVDName = $EWS.Name
            $EWSVDInternalAuthenticationMethods = $EWS.InternalAuthenticationMethods
            $EWSVDExternalAuthenticationMethods = $EWS.ExternalAuthenticationMethods
            $EWSVDWSAuthetication = $EWS.WSSecurityAuthentication
            $EWSVDWSAutheticationColor = "green"
            $EWSVDWindowsAuthentication = $EWS.WindowsAuthentication
            if ($EWSVDWindowsAuthentication -eq "True") {
                $EWSVDWindowsAuthenticationColor = "green"
            }
            else {
                $EWSDWindowsAuthenticationColor = "red"
            }
            $EWSVDInternalNblBypassUrl = $EWS.InternalNblBypassUrl
            $EWSVDInternalUrl = $EWS.InternalUrl
            $EWSVDExternalUrl = $EWS.ExternalUrl
            $script:html +=
            " <div><b>============================</b></div>
            <div><b>Identity:</b> $EWSVDIdentity</div>
            <div><b>Name:</b> $EWSVDName </div>
            <div><b>InternalAuthenticationMethods:</b> $EWSVDInternalAuthenticationMethods </div>
            <div><b>ExternalAuthenticationMethods:</b> $EWSVDExternalAuthenticationMethods </div>
            <div><b>WSAuthetication:</b> <span style='color:green'>$EWSVDWSAuthetication</span></div>
            <div><b>WindowsAuthentication:</b> <span style='color:$EWSVDWindowsAuthenticationColor'>$EWSVDWindowsAuthentication</span></div>
            <div><b>InternalUrl:</b> $EWSVDInternalUrl </div>
            <div><b>ExternalUrl:</b> $EWSVDExternalUrl </div>  "
        }
    }
    else {
        Write-Host -foregroundcolor Red " WSSecurityAuthentication shoud be True."
        foreach ( $EWS in $Global:AutoDiscoveryVirtualDirectory) {
            Write-Host " $($EWS.Identity) "
            Write-Host -ForegroundColor Red "  WSSecurityAuthentication: $($ser.WSSecurityAuthentication) "
            $EWSVDIdentity = $EWS.Identity
            $EWSVDName = $EWS.Name
            $EWSVDInternalAuthenticationMethods = $EWS.InternalAuthenticationMethods
            $EWSVDExternalAuthenticationMethods = $EWS.ExternalAuthenticationMethods
            $EWSVDWSAuthetication = $EWS.WSSecurityAuthentication
            $EWSVDWSAutheticationColor = "green"
            $EWSVDWindowsAuthentication = $EWS.WindowsAuthentication
            if ($EWSVDWindowsAuthentication -eq "True") {
                $EWSVDWindowsAuthenticationColor = "green"
            }
            else {
                $EWSDWindowsAuthenticationColor = "red"
            }
            $EWSVDInternalNblBypassUrl = $EWS.InternalNblBypassUrl
            $EWSVDInternalUrl = $EWS.InternalUrl
            $EWSVDExternalUrl = $EWS.ExternalUrl
            $script:html +=
            " <div><b>============================</b></div>
            <div><b>Identity:</b> $EWSVDIdentity</div>
            <div><b>Name:</b> $EWSVDName </div>
            <div><b>InternalAuthenticationMethods:</b> $EWSVDInternalAuthenticationMethods </div>
            <div><b>ExternalAuthenticationMethods:</b> $EWSVDExternalAuthenticationMethods </div>
            <div><b>WSAuthetication:</b> <span style='color:red'>$EWSVDWSAuthetication</span></div>
            <div><b>WindowsAuthentication:</b> <span style='color:$EWSVDWindowsAuthenticationColor'>$EWSVDWindowsAuthentication</span></div>
            <div><b>InternalUrl:</b> $EWSVDInternalUrl </div>
            <div><b>ExternalUrl:</b> $EWSVDExternalUrl </div>  "
        }
        Write-Host -foregroundcolor White "  Should be True"
    }
    Write-Host -foregroundcolor White "`n  WindowsAuthentication:"
    if ($Global:WebServicesVirtualDirectory.WindowsAuthentication -like "True") {
        foreach ( $EWS in $Global:WebServicesVirtualDirectory) {
            Write-Host " $($EWS.Identity)"
            Write-Host -ForegroundColor Green "  WindowsAuthentication: $($EWS.WindowsAuthentication) "
        }
    }
    else {
        Write-Host -foregroundcolor Red " WindowsAuthentication shoud be True."
        foreach ( $EWS in $Global:AutoDiscoveryVirtualDirectory) {
            Write-Host " $($EWS.Identity) "
            Write-Host -ForegroundColor Red "  WindowsAuthentication: $($ser.WindowsAuthentication) "
        }
        Write-Host -foregroundcolor White "  Should be True"
    }
    $script:html += "
    </td>
    </tr>
    "
    $html | Out-File -FilePath $htmlfile
}

Function AvailabilityAddressSpaceCheck {
    $bar
    Write-Host -foregroundcolor Green " Get-AvailabilityAddressSpace $exchangeOnlineDomain | fl ForestName, UserName, UseServiceAccount, AccessMethod, ProxyUrl, Name"
    Write-Host $bar
    $AvailabilityAddressSpace = Get-AvailabilityAddressSpace $exchangeOnlineDomain -ErrorAction SilentlyContinue | Select-Object ForestName, UserName, UseServiceAccount, AccessMethod, ProxyUrl, Name
    If (!$AvailabilityAddressSpace) {
        $AvailabilityAddressSpace = Get-AvailabilityAddressSpace $exchangeOnlineDomain -ErrorAction SilentlyContinue | Select-Object ForestName, UserName, UseServiceAccount, AccessMethod, ProxyUrl, Name
    }
    $AvailabilityAddressSpace
    $tdAvailabilityAddressSpaceName = $AvailabilityAddressSpace.Name
    Write-Host $bar
    Write-Host -foregroundcolor Green " Summary - On-Prem Availability Address Space Check"
    Write-Host $bar
    Write-Host -foregroundcolor White " ForestName: "
    if ($AvailabilityAddressSpace.ForestName -like $ExchangeOnlineDomain) {
        Write-Host -foregroundcolor Green " " $AvailabilityAddressSpace.ForestName
        $tdAvailabilityAddressSpaceForestName = $AvailabilityAddressSpace.ForestName
        $tdAvailabilityAddressSpaceForestColor = "green"
    }
    else {
        Write-Host -foregroundcolor Red "  ForestName appears not to be correct."
        Write-Host -foregroundcolor White " Should contain the " $ExchaneOnlineDomain
        $tdAvailabilityAddressSpaceForestName = $AvailabilityAddressSpace.ForestName
        $tdAvailabilityAddressSpaceForestColor = "red"
    }
    Write-Host -foregroundcolor White " UserName: "
    if ($AvailabilityAddressSpace.UserName -like "") {
        Write-Host -foregroundcolor Green "  Blank"
        $tdAvailabilityAddressSpaceUserName = " Blank"
        $tdAvailabilityAddressSpaceUserNameColor = "green"
    }
    else {
        Write-Host -foregroundcolor Red " UserName is NOT correct. "
        Write-Host -foregroundcolor White "  Normally it should be blank"
        $tdAvailabilityAddressSpaceUserName = $AvailabilityAddressSpace.UserName
        $tdAvailabilityAddressSpaceUserNameColor = "red"
    }
    Write-Host -foregroundcolor White " UseServiceAccount: "
    if ($AvailabilityAddressSpace.UseServiceAccount -like "True") {
        Write-Host -foregroundcolor Green "  True"
        $tdAvailabilityAddressSpaceUseServiceAccount = $AvailabilityAddressSpace.UseServiceAccount
        $tAvailabilityAddressSpaceUseServiceAccountColor = "green"
    }
    else {
        Write-Host -foregroundcolor Red "  UseServiceAccount appears not to be correct."
        Write-Host -foregroundcolor White "  Should be True"
        $tdAvailabilityAddressSpaceUseServiceAccount = $AvailabilityAddressSpace.UseServiceAccount
        $tAvailabilityAddressSpaceUseServiceAccountColor = "red"
    }
    Write-Host -foregroundcolor White " AccessMethod:"
    if ($AvailabilityAddressSpace.AccessMethod -like "InternalProxy") {
        Write-Host -foregroundcolor Green "  InternalProxy"
        $tdAvailabilityAddressSpaceAccessMethod = $AvailabilityAddressSpace.AccessMethod
        $tdAvailabilityAddressSpaceAccessMethodColor = "green"
    }
    else {
        Write-Host -foregroundcolor Red " AccessMethod appears not to be correct."
        Write-Host -foregroundcolor White " Should be InternalProxy"
        $tdAvailabilityAddressSpaceAccessMethod = $AvailabilityAddressSpace.AccessMethod
        $tdAvailabilityAddressSpaceAccessMethodColor = "red"
    }
    Write-Host -foregroundcolor White " ProxyUrl: "
    $tdAvailabilityAddressSpaceProxyUrl = $AvailabilityAddressSpace.ProxyUrl
    if ([String]::Equals($tdAvailabilityAddressSpaceProxyUrl, $Global:ExchangeOnPremEWS, [StringComparison]::OrdinalIgnoreCase)) {
        Write-Host -foregroundcolor Green " "$AvailabilityAddressSpace.ProxyUrl
        #$tdAvailabilityAddressSpaceProxyUrl = $AvailabilityAddressSpace.ProxyUrl
        $tdAvailabilityAddressSpaceProxyUrlColor = "green"
    }
    else {
        Write-Host -foregroundcolor Red "  ProxyUrl appears not to be correct."
        Write-Host -foregroundcolor White "  Should be $Global:ExchangeOnPremEWS[0] and not $tdAvailabilityAddressSpaceProxyUrl"
        #$tdAvailabilityAddressSpaceProxyUrl = $AvailabilityAddressSpace.ProxyUrl
        $tdAvailabilityAddressSpaceProxyUrlColor = "red"
    }
    Write-Host -ForegroundColor Yellow "`n  Reference: https://learn.microsoft.com/en-us/powershell/module/exchange/get-availabilityaddressspace?view=exchange-ps"
    $script:html += "
    <tr>
    <th colspan='2' style='color:white;'>Summary - On-Premise Get-AvailabilityAddressSpace</th>
    </tr>
    <tr>
    <td><b> Get-AvailabilityAddressSpace $exchangeOnlineDomain | fl ForestName, UserName, UseServiceAccount, AccessMethod, ProxyUrl, Name</b></td>
    <td>
    <div> <b>Forest Name: </b> $tdAvailabilityAddressSpaceForestName</div>
    <div> <b>Name: </b> $tdAvailabilityAddressSpaceName</div>
    <div> <b>UserName: </b> <span style='color:$tdAvailabilityAddressSpaceUserNameColor'>$tdAvailabilityAddressSpaceUserName</span></div>
    <div> <b>Access Method: </b> <span style='color:$tdAvailabilityAddressSpaceAccessMethodColor'>$tdAvailabilityAddressSpaceAccessMethod</span></div>
    <div> <b>ProxyUrl: </b> <span style='color:$tdAvailabilityAddressSpaceProxyUrlColor'>$tdAvailabilityAddressSpaceProxyUrl</span></div>
    </td>
    </tr>"
    $html | Out-File -FilePath $htmlfile
}

Function TestFedTrust {
    Write-Host $bar
    $TestFedTrustFail = 0
    $a = Test-FederationTrust -UserIdentity $useronprem -verbose -ErrorAction silentlycontinue #fails the frist time on multiple ocasions so we have a gohst FedTrustCheck
    Write-Host -foregroundcolor Green  " Test-FederationTrust -UserIdentity $useronprem -verbose"
    Write-Host $bar
    $TestFedTrust = Test-FederationTrust -UserIdentity $useronprem -verbose -ErrorAction silentlycontinue
    $TestFedTrust
    $Script:html += "<tr>
    <th colspan='2' style='color:white;'><b>Summary - On Premise Test-FederationTrust</b></th>
    </tr>
    <tr>
    <td><b> Test-FederationTrust -UserIdentity $useronprem</b></td>
    <td>"
    $i = 0
    while ($i -lt $TestFedTrust.type.Count) {
        $test = $TestFedTrust.type[$i]
        $testType = $TestFedTrust.Type[$i]
        $testMessage = $TestFedTrust.Message[$i]
        $TestFedTrustID = $($TestFedTrust.ID[$i]) 
        if ($test -eq "Error") {
            # Write-Host " $($TestFedTrust.ID[$i]) "
            # Write-Host -foregroundcolor Red " $($TestFedTrust.Type[$i])  "
            # Write-Host " $($TestFedTrust.Message[$i]) "
            $Script:html += "
            
            <div> <span style='color:red'><b>$testType :</b></span> - <div> <b>$TestFedTrustID </b> - $testMessage  </div>
            "
            $TestFedTrustFail++
        }
        if ($test -eq "Success") {
            # Write-Host " $($TestFedTrust.ID[$i]) "
            # Write-Host -foregroundcolor Green " $($TestFedTrust.Type[$i])  "
            # Write-Host " $($TestFedTrust.Message[$i])  "
            $Script:html += "
            
            <div> <span style='color:green'><b>$testType :</b> </span> - <b>$TestFedTrustID </b> - $testMessage</div>"
        }
        $i++
    }

    if ($TestFedTrustFail -eq 0) {
        Write-Host -foregroundcolor Green " Federation Trust Successfully tested"
        $Script:html += "
        <p></p>
        <div class=´green´> <span style='color:green'> Federation Trust Successfully tested </span></div>"
    }
    else {
        Write-Host -foregroundcolor Red " Federation Trust test with Errors"
        $Script:html += "
        <p></p>
        <div class=´red´> <span style='color:red'> Federation Trust tested with Errors </span></div>"
    }



    #Write-Host $bar
    #Write-Host -foregroundcolor Green " Test-FederationTrustCertificate"
    #Write-Host $bar
    $TestFederationTrustCertificate = Test-FederationTrustCertificate -erroraction SilentlyContinue
    #$TestFederationTrustCertificate
    #Write-Host $bar
    if ($TestFederationTrustCertificate) {
    
        Write-Host $bar
        Write-Host -foregroundcolor Green " Test-FederationTrustCertificate"
        Write-Host $bar
        $TestFederationTrustCertificate
       
        $Script:html += "<tr>
                <th colspan='2' style='color:white;'><b>Summary - Test-FederationTrustCertificate</b></th>
                </tr>
                <tr>
                <td><b> Test-FederationTrustCertificate</b></td> 
                <td>"

     
        $j = 0
        while ($j -lt $TestFederationTrustCertificate.Count) {
    
            $TestFederationTrustCertificatej = "<div>" + $TestFederationTrustCertificate.site[$j] + "</div><div>" + $TestFederationTrustCertificate.state[$j] + "</div><div>" + $TestFederationTrustCertificate.Thumbprint[$j] + "</div>"
            $Script:html += "      
                $TestFederationTrustCertificatej
                
                "
            $j++
        }
        $Script:html += "</td>"
    }
    $html | Out-File -FilePath $htmlfile
}

Function TestOrgRel {
    $bar
    $TestFail = 0
    $OrgRelIdentity = $OrgRel.Identity
    
    $OrgRelTargetApplicationUri = $OrgRel.TargetApplicationUri


    if ( $OrgRelTargetApplicationUri -like "Oulook.com") {


    $Script:html += "<tr>
    <th colspan='2' style='color:white;'><b>Summary - Test-OrganizationRelationship</b></th>
    </tr>
    <tr>
    <td><b>Test-OrganizationRelationship -Identity $OrgRelIdentity  -UserIdentity $useronprem</b></td>
    <td>"
    Write-Host -foregroundcolor Green "Test-OrganizationRelationship -Identity $OrgRelIdentity  -UserIdentity $useronprem"
    #need to grab errors and provide alerts in error case
    Write-Host $bar
    $TestOrgRel = Test-OrganizationRelationship -Identity "$($OrgRelIdentity)"  -UserIdentity $useronprem -erroraction SilentlyContinue -warningaction SilentlyContinue
    #$TestOrgRel
    if ($TestOrgRel[16] -like "No Significant Issues to Report") {
        Write-Host -foregroundcolor Green "`n No Significant Issues to Report"
        $Script:html += "
        <div class='green'> <b>No Significant Issues to Report</b><div>"
    }
    else {
        Write-Host -foregroundcolor Red "`n Test Organization Relationship Completed with errors"
        $Script:html += "
        <div class='red'> <b>Test Organization Relationship Completed with errors</b><div>"
    }
    $TestOrgRel[0]
    $TestOrgRel[1]
    $i = 0
    while ($i -lt $TestOrgRel.Length) {
        $element = $TestOrgRel[$i]
        #if ($element.Contains("RESULT: Success.")) {
        if ($element -like "*RESULT: Success.*") {
            $TestOrgRelStep = $TestOrgRel[$i - 1]
            $TestOrgRelStep
            Write-Host -ForegroundColor Green "$element"
            if (![string]::IsNullOrWhitespace($TestOrgRelStep)) {
            $Script:html += "
            <div></b> <span style='color:black'> <b> $TestOrgRelStep :</b></span> <span style='color:green'>$element</span></div>"
        }
        }

        else {
            if ($element -like "*RESULT: Error*") {
                $TestOrgRelStep = $TestOrgRel[$i - 1]
                $TestOrgRelStep
                Write-Host -ForegroundColor Red "$element"
                if (![string]::IsNullOrWhitespace($TestOrgRelStep)) {    
                $Script:html += "
                <div></b> <span style='color:black'> <b> $TestOrgRelStep : </b></span> <span style='color:red'>$element</span></div>"
            }
            }
        }
        $i++
    }
}
else {
    Write-Host -foregroundcolor Green " Test-OrganizationRelationship -Identity $OrgRelIdentity  -UserIdentity $useronprem"
    #need to grab errors and provide alerts in error case
    Write-Host $bar
    $Script:html += "<tr>
    <th colspan='2' style='color:white;'><b>Summary - Test-OrganizationRelationship</b></th>
    </tr>
    <tr>
    <td><b>Test-OrganizationRelationship</b></td>
    <td>"
    Write-Host -foregroundcolor Red "`n Test-OrganizationRelationship can't be run if the Organization Relationship Target Application uri is not correct. Organization Relationship Target Application Uri should be Outlook.com"
    $Script:html += "
    <div class='red'> <b> Test-OrganizationRelationship can't be run if the Organization Relationship Target Application uri is not correct. Organization Relationship Target Application Uri should be Outlook.com</b><div>"
}


    
    Write-host -ForegroundColor Yellow "`n  Reference: https://techcommunity.microsoft.com/t5/exchange-team-blog/how-to-address-federation-trust-issues-in-hybrid-configuration/ba-p/1144285"
    #Write-Host $bar
    $Script:html += "</td>
    </tr>"
    $html | Out-File -FilePath $htmlfile
}

#endregion

#region OAuth Functions

Function IntraOrgConCheck {

    Write-Host -foregroundcolor Green " Get-IntraOrganizationConnector | Select Name,TargetAddressDomains,DiscoveryEndpoint,Enabled" 
    Write-Host $bar
    $IOC = $IntraOrgCon | Format-List
    $IOC
    $tdIntraOrgTargetAddressDomain = $IntraOrgCon.TargetAddressDomains
    $tdDiscoveryEndpoint = $IntraOrgCon.DiscoveryEndpoint
    $tdEnabled = $IntraOrgCon.Enabled

    Write-Host $bar
    Write-Host -foregroundcolor Green " Summary - Get-IntraOrganizationConnector" 
    Write-Host $bar
    $IntraOrgTargetAddressDomain = $IntraOrgCon.TargetAddressDomains.Domain
    $IntraOrgTargetAddressDomain = $IntraOrgTargetAddressDomain.Tolower()
    Write-Host -foregroundcolor White " Target Address Domains: " 
    if ($IntraOrgCon.TargetAddressDomains -like "*$ExchangeOnlineDomain*" -Or $IntraOrgCon.TargetAddressDomains -like "*$ExchangeOnlineAltDomain*" ) {
        Write-Host -foregroundcolor Green " " $IntraOrgCon.TargetAddressDomains
        $tdIntraOrgTargetAddressDomainColor = "green"
    }
    else {
        Write-Host -foregroundcolor Red " Target Address Domains appears not to be correct."
        Write-Host -foregroundcolor White " Should contain the $ExchangeOnlineDomain domain or the $ExchangeOnlineAltDomain domain."
        $tdIntraOrgTargetAddressDomainColor = "red"
    }

    Write-Host -foregroundcolor White " DiscoveryEndpoint: " 
    if ($IntraOrgCon.DiscoveryEndpoint -like "https://autodiscover-s.outlook.com/autodiscover/autodiscover.svc") {
        Write-Host -foregroundcolor Green "  https://autodiscover-s.outlook.com/autodiscover/autodiscover.svc" 
        $tdDiscoveryEndpointColor = "green"
    }
    else {
        Write-Host -foregroundcolor Red "  The DiscoveryEndpoint appears not to be correct. "
        Write-Host -foregroundcolor White "  It should represent the address of EXO autodiscover endpoint."
        Write-Host  "  Examples: https://autodiscover-s.outlook.com/autodiscover/autodiscover.svc; https://outlook.office365.com/autodiscover/autodiscover.svc "
        $tdDiscoveryEndpointColor = "red"
    }
    Write-Host -foregroundcolor White " Enabled: " 
    if ($IntraOrgCon.Enabled -like "True") { 
        Write-Host -foregroundcolor Green "  True "  
        $tdEnabledColor = "green"
    }
    else {
        Write-Host -foregroundcolor Red "  On-Prem Intra Organization Connector is not Enabled"
        Write-Host -foregroundcolor White "  In order to use OAuth it Should be True." 
        write-Host "  If it is set to False, the Organization Realtionship (DAuth) , if enabled, is used for the Hybrid Availability Sharing"
        $tdEnabledColor = "red"
    }


    Write-Host -ForegroundColor Yellow "https://techcommunity.microsoft.com/t5/exchange-team-blog/demystifying-hybrid-free-busy-what-are-the-moving-parts/ba-p/607704"


    # Build HTML table row
    if ($Auth -like "OAuth"){
        $Script:html += "
        <div class='Black'><p></p></div>
    
        <div class='Black'><h2><b>`n Exchange On Premise Free Busy Configuration: `n</b></h2></div>
    
        <div class='Black'><p></p></div>"
    }
        $Script:html += "

    <table style='width:100%'>

   <tr>
      <th colspan='2' style='text-align:center; color:white;'>Exchange On Premise OAuth Configuration</th>
    </tr>
    <tr>
      <th colspan='2' style='color:white;'>Summary - Get-IntraOrganizationConnector</th>
    </tr>

    <tr>
      <td><b>Get-IntraOrganizationConnector:</b></td>
      <td>
        <div><b>Target Address Domains:</b><span style='color: $tdIntraOrgTargetAddressDomainColor'>$($tdIntraOrgTargetAddressDomain)</span></div>
        <div><b>Discovery Endpoint:</b><span style='color: $tdDiscoveryEndpointColor;'>$($tdDiscoveryEndpoint)</span></div>
        <div><b>Enabled:</b><span style='color: $tdEnabledColor;'>$($tdEnabled)</span></div>
      </td>
    </tr>
  "

    $html | Out-File -FilePath $htmlfile

} 

Function AuthServerCheck {
    #Write-Host $bar
    Write-Host -foregroundcolor Green " Get-AuthServer | Select Name,IssuerIdentifier,TokenIssuingEndpoint,AuthMetadataUrl,Enabled"
    Write-Host $bar
    $AuthServer = Get-AuthServer | Where-Object { $_.Name -like "ACS*" } | Select-Object Name, IssuerIdentifier, TokenIssuingEndpoint, AuthMetadataUrl, Enabled
    $AuthServer
    $tdAuthServerIssuerIdentifier = $AuthServer.IssuerIdentifier
    $tdAuthServerTokenIssuingEndpoint = $AuthServer.TokenIssuingEndpoint 
    $tdAuthServerAuthMetadataUrl = $AuthServer.AuthMetadataUrl
    $tdAuthServerEnabled = $AuthServer.Enabled
    Write-Host $bar
    Write-Host -foregroundcolor Green " Summary - Auth Server"
    Write-Host $bar
    Write-Host -foregroundcolor White " IssuerIdentifier: "
    if ($AuthServer.IssuerIdentifier -like "00000001-0000-0000-c000-000000000000" ) {
        Write-Host -foregroundcolor Green " " $AuthServer.IssuerIdentifier
        $tdAuthServerIssuerIdentifierColor = "green"
       
    }
    else {
        Write-Host -foregroundcolor Red " IssuerIdentifier appears not to be correct."
        Write-Host -foregroundcolor White " Should be 00000001-0000-0000-c000-000000000000"
        $tdAuthServerIssuerIdentifierColor = "red"
    
    }
    Write-Host -foregroundcolor White " TokenIssuingEndpoint: "
    if ($AuthServer.TokenIssuingEndpoint -like "https://accounts.accesscontrol.windows.net/*" -and $AuthServer.TokenIssuingEndpoint -like "*/tokens/OAuth/2" ) {
        Write-Host -foregroundcolor Green " " $AuthServer.TokenIssuingEndpoint
        $tdAuthServerTokenIssuingEndpointColor = "green"
    
    
    }
    else {
        Write-Host -foregroundcolor Red " TokenIssuingEndpoint appears not to be correct."
        Write-Host -foregroundcolor White " Should be  https://accounts.accesscontrol.windows.net/<Cloud Tenant ID>/tokens/OAuth/2"
        $tdAuthServerTokenIssuingEndpointColor = "red"
    
    }
    Write-Host -foregroundcolor White " AuthMetadataUrl: "
    if ($AuthServer.AuthMetadataUrl -like "https://accounts.accesscontrol.windows.net/*" -and $AuthServer.TokenIssuingEndpoint -like "*/tokens/OAuth/2" ) {
        Write-Host -foregroundcolor Green " " $AuthServer.AuthMetadataUrl
        $tdAuthServerAuthMetadataUrlColor = "green"
    
    
    }
    else {
        Write-Host -foregroundcolor Red " AuthMetadataUrl appears not to be correct."
        Write-Host -foregroundcolor White " Should be  https://accounts.accesscontrol.windows.net/<Cloud Tenant ID>/metadata/json/1"
        $tdAuthServerAuthMetadataUrlColor = "red"
    
    
    }
    Write-Host -foregroundcolor White " Enabled: "
    if ($AuthServer.Enabled -like "True" ) {
        Write-Host -foregroundcolor Green " " $AuthServer.Enabled
        $tdAuthServerEnabledColor = "green"
    
    
    }
    else {
        Write-Host -foregroundcolor Red " Enalbed: False "
        Write-Host -foregroundcolor White " Should be True"
        $tdAuthServerEnabledColor = "red"
   
   
    }


    $Script:html += "
    <tr>
      <th colspan='2' style='color:white;'>Summary - Get-AuthServer</th>
    </tr>

    <tr>
      <td><b> Get-AuthServer | Select Name,IssuerIdentifier,TokenIssuingEndpoint,AuthMetadataUrl,Enabled</b></td>
      <td>
        <div><b>IssuerIdentifier:</b><span style='color: $tdAuthServerIssuerIdentifierColor'>$($tdAuthServerIssuerIdentifier)</span></div>
        <div><b>TokenIssuingEndpoint:</b><span style='color: $tdAuthServerTokenIssuingEndpointColor;'>$($tdAuthServerTokenIssuingEndpoint)</span></div>
        <div><b>AuthMetadataUrl:</b><span style='color: $tdAuthServerAuthMetadataUrlColor;'>$($tdAuthServerAuthMetadataUrl)</span></div>
        <div><b>Enabled:</b><span style='color: $tdAuthServerEnabledColor;'>$($tdAuthServerEnabled)</span></div>
      </td>
    </tr>
  "

    $html | Out-File -FilePath $htmlfile
}

Function PartnerApplicationCheck {
    #Write-Host $bar
    Write-Host -foregroundcolor Green " Get-PartnerApplication |  ?{$_.ApplicationIdentifier -eq '00000002-0000-0ff1-ce00-000000000000'
    -and $_.Realm -eq ''} | Select Enabled, ApplicationIdentifier, CertificateStrings, AuthMetadataUrl, Realm, UseAuthServer,
    AcceptSecurityIdentifierInformation, LinkedAccount, IssuerIdentifier, AppOnlyPermissions, ActAsPermissions, Name"
    Write-Host $bar
    $PartnerApplication = Get-PartnerApplication |  Where-Object { $_.ApplicationIdentifier -eq '00000002-0000-0ff1-ce00-000000000000' -and $_.Realm -eq '' } | Select-Object Enabled, ApplicationIdentifier, CertificateStrings, AuthMetadataUrl, Realm, UseAuthServer, AcceptSecurityIdentifierInformation, LinkedAccount, IssuerIdentifier, AppOnlyPermissions, ActAsPermissions, Name
    $PartnerApplication
    $tdPartnerApplicationEnabled = $PartnerApplication.Enabled
    $tdPartnerApplicationApplicationIdentifier = $PartnerApplication.ApplicationIdentifier
    $tdPartnerApplicationCertificateStrings = $PartnerApplication.CertificateStrings
    $tdPartnerApplicationAuthMetadataUrl = $PartnerApplication.AuthMetadataUrl
    $tdPartnerApplicationRealm = $PartnerApplication.Realm
    $tdPartnerApplicationUseAuthServer = $PartnerApplication.UseAuthServer
    $tdPartnerApplicationAcceptSecurityIdentifierInformation = $PartnerApplication.AcceptSecurityIdentifierInformation
    $tdPartnerApplicationLinkedAccount = $PartnerApplication.LinkedAccount
    $tdPartnerApplicationIssuerIdentifier = $PartnerApplication.IssuerIdentifier
    $tdPartnerApplicationAppOnlyPermissions = $PartnerApplication.AppOnlyPermissions
    $tdPartnerApplicationActAsPermissions = $PartnerApplication.ActAsPermissions
    $tdPartnerApplicationName = $PartnerApplication.Name
    
    Write-Host $bar
    Write-Host -foregroundcolor Green " Summary - Partner Application"
    Write-Host $bar
    Write-Host -foregroundcolor White " Enabled: "
    if ($PartnerApplication.Enabled -like "True" ) {
        Write-Host -foregroundcolor Green " " $PartnerApplication.Enabled
        $tdPartnerApplicationEnabledColor = "green"
    
    }
    else {
        Write-Host -foregroundcolor Red " Enabled: False "
        Write-Host -foregroundcolor White " Should be True"
        $tdPartnerApplicationEnabledColor = "red"
    
    
    }
    Write-Host -foregroundcolor White " ApplicationIdentifier: "
    if ($PartnerApplication.ApplicationIdentifier -like "00000002-0000-0ff1-ce00-000000000000" ) {
        Write-Host -foregroundcolor Green " " $PartnerApplication.ApplicationIdentifier
        $tdPartnerApplicationApplicationIdentifierColor = "green"
    
    
    }
    else {
        Write-Host -foregroundcolor Red " ApplicationIdentifier does not appear to be correct"
        Write-Host -foregroundcolor White " Should be 00000002-0000-0ff1-ce00-000000000000"
        $tdPartnerApplicationApplicationIdentifierColor = "red"
    
    
    }
    Write-Host -foregroundcolor White " AuthMetadataUrl: "
    if ([string]::IsNullOrWhitespace( $PartnerApplication.AuthMetadataUrl)) {
        Write-Host -foregroundcolor Green "  Blank"
        $tdPartnerApplicationAuthMetadataUrlColor = "green"
        $tdPartnerApplicationAuthMetadataUrl = "Blank"    
    
    }
    else {
        Write-Host -foregroundcolor Red " AuthMetadataUrl does not aooear correct"
        Write-Host -foregroundcolor White " Should be Blank"
        $tdPartnerApplicationAuthMetadataUrlColor = "red"
        $tdPartnerApplicationAuthMetadataUrl = " Should be Blank"
    }
    Write-Host -foregroundcolor White " Realm: "
    if ([string]::IsNullOrWhitespace( $PartnerApplication.Realm)) {
        Write-Host -foregroundcolor Green "  Blank"
        $tdPartnerApplicationRealmColor = "green"
        $tdPartnerApplicationRealm = "Blank"
    
    }
    else {
        Write-Host -foregroundcolor Red "  Realm does not appear to be correct"
        Write-Host -foregroundcolor White " Should be Blank"
        $tdPartnerApplicationRealmColor = "Red"
        $tdPartnerApplicationRealm = "Should be Blank"
    
    }
    Write-Host -foregroundcolor White " LinkedAccount: "
    if ($PartnerApplication.LinkedAccount -like "$exchangeOnPremDomain/Users/Exchange Online-ApplicationAccount" -or $PartnerApplication.LinkedAccount -like "$exchangeOnPremLocalDomain/Users/Exchange Online-ApplicationAccount"  ) {
        Write-Host -foregroundcolor Green " " $PartnerApplication.LinkedAccount
        $tdPartnerApplicationLinkedAccountColor = "green"
    
    }
    else {
        Write-Host -foregroundcolor Red "  LinkedAccount value does not appear to be correct"
        Write-Host -foregroundcolor White "  Should be $exchangeOnPremLocalDomain/Users/Exchange Online-ApplicationAccount"
        Write-Host "  If you value is empty, set it to correspond to the Exchange Online-ApplicationAccount which is located at the root of Users container in AD. After you make the change, reboot the servers."
        Write-Host "  Example: contoso.com/Users/Exchange Online-ApplicationAccount"
        $tdPartnerApplicationLinkedAccountColor = "red"
        $tdPartnerApplicationLinkedAccount 
    }


    $Script:html += "
    <tr>
      <th colspan='2' style='color:white;'>Summary - Get-PartnerApplication</th>
    </tr>
    <tr>
      <td><b> Get-PartnerApplication |  ?{$_.ApplicationIdentifier -eq '00000002-0000-0ff1-ce00-000000000000'
    -and $_.Realm -eq ''} | Select Enabled, ApplicationIdentifier, CertificateStrings, AuthMetadataUrl, Realm, UseAuthServer,
    AcceptSecurityIdentifierInformation, LinkedAccount, IssuerIdentifier, AppOnlyPermissions, ActAsPermissions, Name</b></td>
      <td>
        <div><b>Enabled:</b><span style='color: $tdPartnerApplicationEnabledColor'>$($tdPartnerApplicationEnabled)</span></div>
        <div><b>ApplicationIdentifier:</b><span style='color: $tdPartnerApplicationApplicationIdentifierColor;'>$($tdPartnerApplicationApplicationIdentifier)</span></div>
        <div><b>CertificateStrings:</b><span style='color: $tdPartnerApplicationCertificateStringsColor;'>$($tdPartnerApplicationCertificateStrings)</span></div>
        <div><b>AuthMetadataUrl:</b><span style='color: $tdPartnerApplicationAuthMetadataUrlColor;'>$($tdPartnerApplicationAuthMetadataUrl)</span></div>
        <div><b>Realm:</b><span style='color: $tdPartnerApplicationRealmColor'>$($tdPartnerApplicationRealm)</span></div>
        <div><b>LinkedAccount:</b><span style='color: $tdPartnerApplicationLinkedAccountColor;'>$($tdPartnerApplicationLinkedAccount)</span></div>
        <div><b>IssuerIdentifier:</b><span style='color: $tdPartnerApplicationEnabledColor'>$($tdPartnerApplicationEnabled)</span></div>
        <div><b>AppOnlyPermissions:</b><span style='color: $tdPartnerApplicationApplicationIdentifierColor;'>$($tdPartnerApplicationApplicationIdentifier)</span></div>
        <div><b>ActAsPermissions:</b><span style='color: $tdPartnerApplicationCertificateStringsColor;'>$($tdPartnerApplicationCertificateStrings)</span></div>
        <div><b>Name:</b><span style='color: $tdPartnerApplicationAuthMetadataUrlColor;'>$($tdPartnerApplicationAuthMetadataUrl)</span></div>

      </td>
    </tr>
  "

    $html | Out-File -FilePath $htmlfile
}

Function ApplicationAccounCheck {
    #Write-Host $bar
    Write-Host -foregroundcolor Green " Get-user '$exchangeOnPremLocalDomain/Users/Exchange Online-ApplicationAccount' | Select Name, RecipientType, RecipientTypeDetails, UserAccountControl"
    Write-Host $bar
    $ApplicationAccount = Get-user "$exchangeOnPremLocalDomain/Users/Exchange Online-ApplicationAccount"  | Select-Object Name, RecipientType, RecipientTypeDetails, UserAccountControl
    $ApplicationAccount
    $tdApplicationAccountRecipientType = $ApplicationAccount.RecipientType
    $tdApplicationAccountRecipientTypeDetails = $ApplicationAccount.RecipientTypeDetails
    $tdApplicationAccountUserAccountControl = $ApplicationAccount.UserAccountControl
    Write-Host $bar
    Write-Host -foregroundcolor Green " Summary - Application Account"
    Write-Host $bar
    Write-Host -foregroundcolor White " RecipientType: "
    if ($ApplicationAccount.RecipientType -like "User" ) {
        Write-Host -foregroundcolor Green " " $ApplicationAccount.RecipientType
        $tdApplicationAccountRecipientTypeColor = "green"
    }
    else {
        Write-Host -foregroundcolor Red " RecipientType value is $ApplicationAccount.RecipientType "
        Write-Host -foregroundcolor White " Should be User"
        $tdApplicationAccountRecipientTypeColor = "red"
    }
    Write-Host -foregroundcolor White " RecipientTypeDetails: "
    if ($ApplicationAccount.RecipientTypeDetails -like "LinkedUser" ) {
        Write-Host -foregroundcolor Green " " $ApplicationAccount.RecipientTypeDetails
        $tdApplicationAccountRecipientTypeDetailsColor = "green"
    }
    else {
        Write-Host -foregroundcolor Red " RecipientTypeDetails value is $ApplicationAccount.RecipientTypeDetails"
        Write-Host -foregroundcolor White " Should be LinkedUser"
        $tdApplicationAccountRecipientTypeDetailsColor = "red"
    }
    Write-Host -foregroundcolor White " UserAccountControl: "
    if ($ApplicationAccount.UserAccountControl -like "AccountDisabled, PasswordNotRequired, NormalAccount" ) {
        Write-Host -foregroundcolor Green " " $ApplicationAccount.UserAccountControl
        $tdApplicationAccountUserAccountControlColor = "green"
    }
    else {
        Write-Host -foregroundcolor Red " UserAccountControl value does not seem correct"
        Write-Host -foregroundcolor White " Should be AccountDisabled, PasswordNotRequired, NormalAccount"
        $tdApplicationAccountUserAccountControlColor = "red"
    }
     
    $Script:html += "
      <tr>
      <th colspan='2' style='color:white;'>Summary - Get-User ApplicationAccount</th>
    </tr>
    <tr>
      <td><b>  Get-user '$exchangeOnPremLocalDomain/Users/Exchange Online-ApplicationAccount' | Select Name, RecipientType, RecipientTypeDetails, UserAccountControl':</b></td>
      <td>
        <div><b>RecipientType:</b><span style='color: $tdApplicationAccountRecipientTypeColor'>$($tdApplicationAccountRecipientType)</span></div>
        <div><b>RecipientTypeDetails:</b><span style='color: $tdApplicationAccountRecipientTypeDetailsColor;'>$($tdApplicationAccountRecipientTypeDetails)</span></div>
        <div><b>UserAccountControl:</b><span style='color: $tdApplicationAccountUserAccountControlColor;'>$($tdApplicationAccountUserAccountControl)</span></div>
        
      </td>
    </tr>
  "

    $html | Out-File -FilePath $htmlfile
}

Function ManagementRoleAssignmentCheck {
    Write-Host -foregroundcolor Green " Get-ManagementRoleAssignment -RoleAssignee Exchange Online-ApplicationAccount | Select Name,Role -AutoSize"
    Write-Host $bar
    $ManagementRoleAssignment = Get-ManagementRoleAssignment -RoleAssignee "Exchange Online-ApplicationAccount"  | Select-Object Name, Role
    $M = $ManagementRoleAssignment | Out-String
    $M
   

    Write-Host $bar
    Write-Host -foregroundcolor Green " Summary - Management Role Assignment for the Exchange Online-ApplicationAccount"
    Write-Host $bar
    Write-Host -foregroundcolor White " Role: "
    if ($ManagementRoleAssignment.Role -like "*UserApplication*" ) {
        Write-Host -foregroundcolor Green "  UserApplication Role Assigned"
        $tdManagementRoleAssignmentUserApplication = " UserApplication Role Assigned"
        $tdManagementRoleAssignmentUserApplicationColor = "green"
    }
    else {
        Write-Host -foregroundcolor Red "  UserApplication Role not present for the Exchange Online-ApplicationAccount"
        $tdManagementRoleAssignmentUserApplication = " UserApplication Role not present"
        $tdManagementRoleAssignmentUserApplicationColor = "red"
    
    }
    if ($ManagementRoleAssignment.Role -like "*ArchiveApplication*" ) {
        Write-Host -foregroundcolor Green "  ArchiveApplication Role Assigned"
        $tdManagementRoleAssignmentArchiveApplication = " ArchiveApplication Role Assigned"
        $tdManagementRoleAssignmentArchiveApplicationColor = "green"
    
    }
    else {
        Write-Host -foregroundcolor Red "  ArchiveApplication Role not present for the Exchange Online-ApplicationAccount"
        $tdManagementRoleAssignmentArchiveApplication = " ArchiveApplication Role not Assigned"
        $tdManagementRoleAssignmentArchiveApplicationColor = "red"
    
    }
    if ($ManagementRoleAssignment.Role -like "*LegalHoldApplication*" ) {
        Write-Host -foregroundcolor Green "  LegalHoldApplication Role Assigned"
        $tdManagementRoleAssignmentLegalHoldApplication = " LegalHoldApplication Role Assigned"
        $tdManagementRoleAssignmentLegalHoldApplicationColor = "green"
    
    }
    else {
        Write-Host -foregroundcolor Red "  LegalHoldApplication Role not present for the Exchange Online-ApplicationAccount"
        $tdManagementRoleAssignmentLegalHoldApplication = " LegalHoldApplication Role Assigned"
        $tdManagementRoleAssignmentLegalHoldApplicationColor = "green"
    
    }
    if ($ManagementRoleAssignment.Role -like "*Mailbox Search*" ) {
        Write-Host -foregroundcolor Green "  Mailbox Search Role Assigned"
        $tdManagementRoleAssignmentMailboxSearch = " Mailbox Search Role Assigned"
        $tdManagementRoleAssignmentMailboxSearchColor = "green"
    
    }
    else {
        Write-Host -foregroundcolor Red "  Mailbox Search Role not present for the Exchange Online-ApplicationAccount"
        $tdManagementRoleAssignmentMailboxSearch = " Mailbox Search Role Not Assigned"
        $tdManagementRoleAssignmentMailboxSearchColor = "red"
    
    
    }
    if ($ManagementRoleAssignment.Role -like "*TeamMailboxLifecycleApplication*" ) {
        Write-Host -foregroundcolor Green "  TeamMailboxLifecycleApplication Role Assigned"
        $tdManagementRoleAssignmentTeamMailboxLifecycleApplication = " TeamMailboxLifecycleApplication Role Assigned"
        $tdManagementRoleAssignmentTeamMailboxLifecycleApplicationColor = "green"
    
    
    }
    else {
        Write-Host -foregroundcolor Red "  TeamMailboxLifecycleApplication Role not present for the Exchange Online-ApplicationAccount"
        $tdManagementRoleAssignmentTeamMailboxLifecycleApplication = " TeamMailboxLifecycleApplication Role Not Assigned"
        $tdManagementRoleAssignmentTeamMailboxLifecycleApplicationColor = "red"
    
    
    
    }
    if ($ManagementRoleAssignment.Role -like "*MailboxSearchApplication*" ) {
        Write-Host -foregroundcolor Green "  MailboxSearchApplication Role Assigned"
        $tdManagementRoleMailboxSearchApplication = " MailboxSearchApplication Role Assigned"
        $tdManagementRoleMailboxSearchApplicationColor = "green"
    
    
    
    }
    else {
        Write-Host -foregroundcolor Red "  MailboxSearchApplication Role not present for the Exchange Online-ApplicationAccount"
        $tdManagementRoleMailboxSearchApplication = " MailboxSearchApplication Role Not Assigned"
        $tdManagementRoleMailboxSearchApplicationColor = "red"
    
    
    }
    if ($ManagementRoleAssignment.Role -like "*MeetingGraphApplication*" ) {
        Write-Host -foregroundcolor Green "  MeetingGraphApplication Role Assigned"
        $tdManagementRoleMeetingGraphApplication = " MeetingGraphApplication Role Assigned"
        $tdManagementRoleMeetingGraphApplicationColor = "green"
        
    
    }
    else {
        Write-Host -foregroundcolor Red "  MeetingGraphApplication Role not present for the Exchange Online-ApplicationAccount"
        $tdManagementRoleMeetingGraphApplication = " MeetingGraphApplication Role Not Assigned"
        $tdManagementRoleMeetingGraphApplicationColor = "red"
        
    }

     

      
              
    $tdManagementRoleMeetingGraphApplication = " MailboxSearchApplication Role Assigned"
    $tdManagementRoleMeetingGraphApplicationColor = "green"

    
    $Script:html += "
      <tr>
      <th colspan='2' style='color:white;'>Summary - Get-ManagementRoleAssignment</th>
    </tr>
    <tr>
      <td><b>  Get-ManagementRoleAssignment -RoleAssignee Exchange Online-ApplicationAccount | Select Name,Role</b></td>
      <td>
        <div><b>UserApplication Role:</b><span style='color: $tdManagementRoleAssignmentUserApplicationColor'>$($tdManagementRoleAssignmentUserApplication)</span></div>
        <div><b>ArchiveApplication Role:</b><span style='color: $tdManagementRoleAssignmentArchiveApplicationColor;'>$($tdManagementRoleAssignmentArchiveApplication)</span></div>
        <div><b>LegalHoldApplication Role:</b><span style='color: $tdManagementRoleAssignmentLegalHoldApplicationColor;'>$($tdManagementRoleAssignmentLegalHoldApplication)</span></div>
        <div><b>Mailbox Search Role:</b><span style='color: $tdManagementRoleAssignmentMailboxSearchColor'>$($tdManagementRoleAssignmentMailboxSearch)</span></div>
        <div><b>TeamMailboxLifecycleApplication Role:</b><span style='color: $tdManagementRoleAssignmentTeamMailboxLifecycleApplicationColor;'>$($tdManagementRoleAssignmentTeamMailboxLifecycleApplication)</span></div>
        <div><b>MailboxSearchApplication Rolel:</b><span style='color: $tdManagementRoleMailboxSearchApplicationColor;'>$($tdManagementRoleMailboxSearchApplication)</span></div>
        <div><b>MeetingGraphApplication Role:</b><span style='color: $tdManagementRoleMeetingGraphApplicationColor;'>$($tdManagementRoleMeetingGraphApplication)</span></div>
        
      </td>
    </tr>
  "

    $html | Out-File -FilePath $htmlfile

}

Function AuthConfigCheck {
    Write-Host -foregroundcolor Green " Get-AuthConfig | Select *Thumbprint, ServiceName, Realm, Name"
    Write-Host $bar
    $AuthConfig = Get-AuthConfig | Select-Object *Thumbprint, ServiceName, Realm, Name
    $AC = $AuthConfig | Format-List
    $AC
    
    $tdAuthConfigName = $AuthConfig.Name
    


    Write-Host $bar
    Write-Host -foregroundcolor Green " Summary - Auth Config"
    Write-Host $bar
    if (![string]::IsNullOrWhitespace($AuthConfig.CurrentCertificateThumbprint)) {
        Write-HOst " Thumbprint: "$AuthConfig.CurrentCertificateThumbprint
        Write-Host -foregroundcolor Green " Certificate is Assigned"
        $tdAuthConfigCurrentCertificateThumbprint = $AuthConfig.CurrentCertificateThumbprint
        $tdAuthConfigCurrentCertificateThumbprintColor = "green"
    }
    else {
        Write-HOst " Thumbprint: "$AuthConfig.CurrentCertificateThumbprint
        Write-Host -foregroundcolor Red " No valid certificate Assigned "
        $tdAuthConfigCurrentCertificateThumbprintColor = "red"
        $tdAuthConfigCurrentCertificateThumbprint = "$AuthConfig.CurrentCertificateThumbprint - No valid certificate Assigned "
    }
    if ($AuthConfig.ServiceName -like "00000002-0000-0ff1-ce00-000000000000" ) {
        Write-HOst " ServiceName: "$AuthConfig.ServiceName
        Write-Host -foregroundcolor Green " Service Name Seems correct"
        $tdAuthConfigServiceNameColor = "green"
        $tdAuthConfigServiceName = $AuthConfig.ServiceName
    }
    else {
        Write-HOst " ServiceName: "$AuthConfig.ServiceName
        Write-Host -foregroundcolor Red " Service Name does not Seems correct. Should be 00000002-0000-0ff1-ce00-000000000000"
        $tdAuthConfigServiceNameColor = "red"
        $tdAuthConfigServiceName = "$AuthConfig.ServiceName  Should be 00000002-0000-0ff1-ce00-000000000000"
    }
    if ([string]::IsNullOrWhitespace($AuthConfig.Realm)) {
        Write-HOst " Realm: "
        Write-Host -foregroundcolor Green " Realm is Blank"
        $tdAuthConfigRealmColor = "green"
        $tdAuthConfigRealm = " Realm is Blank"
    }
    else {
        Write-HOst " Realm: "$AuthConfig.Realm
        Write-Host -foregroundcolor Red " Realm should be Blank"
        $tdAuthConfigRealmColor = "red"
        $tdAuthConfigRealm = "$tdAuthConfig.Realm - Realm should be Blank"

    }
    
    $Script:html += "
      <tr>
      <th colspan='2' style='color:white;'>Summary - Get-AuthConfig</th>
    </tr>
    <tr>
      <td><b>  Get-AuthConfig | Select-Object *Thumbprint, ServiceName, Realm, Name</b></td>
      <td>
        <div><b>Name:</b><span >$($tdAuthConfigName)</span></div>
        <div><b>Thumbprint:</b><span style='color: $tdAuthConfigCurrentCertificateThumbprintColor'>$($tdAuthConfigCurrentCertificateThumbprint)</span></div>
        <div><b>ServiceName:</b><span style='color:$tdAuthConfigServiceNameColor;'>$( $tdAuthConfigServiceName)</span></div>
        <div><b>Realm:</b><span style='color: $tdAuthConfigRealmColor;'>$($tdAuthConfigRealm)</span></div>
        
        
      </td>
    </tr>
  "

    $html | Out-File -FilePath $htmlfile


}

Function CurrentCertificateThumbprintCheck {
    $thumb = Get-AuthConfig | Select-Object CurrentCertificateThumbprint
    $thumbprint = $thumb.currentcertificateThumbprint
    #Write-Host $bar
    Write-Host -ForegroundColor Green " Get-ExchangeCertificate -Thumbprint $thumbprint | Select *"
    Write-Host $bar
    $CurrentCertificate = get-exchangecertificate $thumb.CurrentCertificateThumbprint | Select-Object *
    $CC = $CurrentCertificate | Format-List
    $CC





    Write-Host $bar
    Write-Host -foregroundcolor Green " Summary - Microsoft Exchange Server Auth Certificate"
    Write-Host $bar
    if ($CurrentCertificate.Issuer -like "CN=Microsoft Exchange Server Auth Certificate" ) {
        write-Host " Issuer: " $CurrentCertificate.Issuer
        Write-Host -foregroundcolor Green " Issuer is CN=Microsoft Exchange Server Auth Certificate"
        $tdCurrentCertificateIssuer = "   $($CurrentCertificate.Issuer) - Issuer is CN=Microsoft Exchange Server Auth Certificate"
        $tdCurrentCertificateIssuerColor = "green"


    }
    else {
        Write-Host -foregroundcolor Red "  Issuer is not CN=Microsoft Exchange Server Auth Certificate"
        $tdCurrentCertificateIssuer = "   $($CurrentCertificate.Issuer) - Issuer is Not CN=Microsoft Exchange Server Auth Certificate"
        $tdCurrentCertificateIssuerColor = "red"
    
    
    }
    if ($CurrentCertificate.Services -like "SMTP" ) {
        Write-Host " Services: " $CurrentCertificate.Services
        Write-Host -foregroundcolor Green "  Certificate enabled for SMTP"
        $tdCurrentCertificateServices = "  $($tdCurrentCertificate.Services) - Certificate enabled for SMTP"
        $tdCurrentCertificateServicesColor = "green"
    
    
    
    }
    else {
        Write-Host -foregroundcolor Red "  Certificate Not enabled for SMTP"
        $tdCurrentCertificateServices = "  $($tdCurrentCertificate.Services) - Certificate Not enabled for SMTP"
        $tdCurrentCertificateServicesColor = "red"
    
    
    }
    if ($CurrentCertificate.Status -like "Valid" ) {
        Write-Host " Status: " $CurrentCertificate.Status
        Write-Host -foregroundcolor Green "  Certificate is valid"
        $tdCurrentCertificateStatus = "  Certificate is valid"
        $tdCurrentCertificateStatusColor = "green"
    
    
    }
    else {
        Write-Host -foregroundcolor Red "  Certificate is not Valid"
        $tdCurrentCertificateStatus = "  Certificate is Not Valid"
        $tdCurrentCertificateStatusColor = "red"
    
    
    
    }
    if ($CurrentCertificate.Subject -like "CN=Microsoft Exchange Server Auth Certificate" ) {
        Write-Host " Subject: " $CurrentCertificate.Subject
        Write-Host -foregroundcolor Green "  Subject is CN=Microsoft Exchange Server Auth Certificate"
        $tdCurrentCertificateSubject = "  Subject is CN=Microsoft Exchange Server Auth Certificate"
        $tdCurrentCertificateSubjectColor = "green"
    
    }
    else {
        Write-Host -foregroundcolor Red "  Subject is not CN=Microsoft Exchange Server Auth Certificate"
        $tdCurrentCertificateSubject = "  $($CurrentCertificate.Subject) - Subject should be CN=Microsoft Exchange Server Auth Certificate"
        $tdCurrentCertificateSubjectColor = "red"
    
    
    }
    Write-Host -ForegroundColor White "`n Checking Exchange Auth Certificate Distribution `n"
    $CheckAuthCertDistribution = foreach ($name in (get-exchangeserver).name) { Get-Exchangecertificate -Thumbprint (Get-AuthConfig).CurrentCertificateThumbprint -server $name -ErrorAction SilentlyContinue | Select-Object Identity, thumbprint, services, subject }
    foreach ($serv in $CheckAuthCertDistribution) {
        $Servername = ($serv -split "\.")[0]
        Write-Host -ForegroundColor White  "  Server: " $servername
        #Write-Host  "   Thumbprint: " $Thumbprint
        if ($serv.Thumbprint -like $thumbprint) {
            Write-Host  "   Thumbprint: "$serv.Thumbprint
            Write-Host  "   Subject: "$serv.Subject
            $servIdentity = $serv.Identity
            $tdCheckAuthCertDistribution = "   <div>Certificate with Thumbprint: $($serv.Thumbprint) Subject: $($serv.Subject) is present in server $servIdentity</div>"
            $tdCheckAuthCertDistributionColor = "green"
    

        }
        if ($serv.Thumbprint -ne $thumbprint) {
            Write-Host -foregroundcolor Red "  Auth Certificate seems Not to be present in $Servername"
            $tdCheckAuthCertDistribution = "   Auth Certificate seems Not to be present in $Servername"
            $tdCheckAuthCertDistributionColor = "Red"
        
        
        }
    }

    $Script:html += "
      <tr>
      <th colspan='2' style='color:white;'>Summary - Get-ExchnageCertificate AuthCertificate</th>
    </tr>
    <tr>
      <td><b>  Get-ExchangeCertificate $thumb.CurrentCertificateThumbprint | Select-Object *</b></td>
      <td>
        <div><b>Issuer:</b><span style='color: $tdCurrentCertificateIssuerColor'>$($tdCurrentCertificateIssuer)</span></div>
        <div><b>Services:</b><span style='color: $tdCurrentCertificateServicesColor'>$($tdCurrentCertificateServices)</span></div>
        <div><b>Status:</b><span style='color:$tdCurrentCertificateStatusColor;'>$( $tdCurrentCertificateStatus)</span></div>
        <div><b>Subject:</b><span style='color: $tdCurrentCertificateSubjectColor;'>$($tdCurrentCertificateSubject)</span></div>
        <div><b>Distribution:</b><span style='color: $tdCheckAuthCertDistributionColor;'>$($tdCheckAuthCertDistribution)</span></div>
        
        
      </td>
    </tr>
  "

    $html | Out-File -FilePath $htmlfile


}

Function AutoDVirtualDCheckOauth {
    #Write-Host -foregroundcolor Green " `n On-Prem Autodiscover Virtual Directory `n "
    Write-Host -foregroundcolor Green " Get-AutodiscoverVirtualDirectory | Select Identity, Name,ExchangeVersion,*authentication*"
    Write-Host $bar
    $AutoDiscoveryVirtualDirectoryOAuth = Get-AutodiscoverVirtualDirectory | Select-Object Identity, Name, ExchangeVersion, *authentication*
    #Check if null or set
    $AD = $AutoDiscoveryVirtualDirectoryOAuth | Format-List
    $AD
    $script:html += "<tr>
    <th colspan='2' style='color:white;'>Summary - Get-AutodiscoverVirtualDirectory</th>
    </tr>
    <tr>
    <td><b>Get-AutodiscoverVirtualDirectory:</b></td>
    <td>"



    if ($Auth -contains "OAuth") {
    }
    Write-Host $bar
    Write-Host -foregroundcolor Green " Summary - Get-AutodiscoverVirtualDirectory"
    Write-Host $bar
    Write-Host -foregroundcolor White "  InternalAuthenticationMethods"
    if ($AutoDiscoveryVirtualDirectoryOAuth.InternalAuthenticationMethods -like "*OAuth*") {
        foreach ( $EWS in $AutoDiscoveryVirtualDirectoryOAuth) {
            Write-Host " $($EWS.Identity) "
            Write-Host -ForegroundColor Green "  InternalAuthenticationMethods Include OAuth Authentication Method "
            
            $AutodVDIdentity = $EWS.Identity
            $AutodVDName = $EWS.Name
            $AutodVDInternalAuthenticationMethods = $EWS.InternalAuthenticationMethods
            $AutodVDExternalAuthenticationMethods = $EWS.ExternalAuthenticationMethods
            $AutodVDWSAuthetication = $EWS.WSSecurityAuthentication
            $AutodVDWSAutheticationColor = "green"
            $AutodVDWindowsAuthentication = $EWS.WindowsAuthentication
            $AutodVDOAuthAuthentication = $EWS.OAuthAuthentication
            if ($AutodVDWindowsAuthentication -eq "True") {
                $AutodVDWindowsAuthenticationColor = "green"
            }
            else {
                $AutodVDWindowsAuthenticationColor = "red"
            }
            if ($AutodVDOAuthAuthentication -eq "True") {
                $AutodVDOAuthAuthenticationColor = "green"
            }
            else {
                $AutodVDOAuthAuthenticationColor = "red"
            }
            $AutodVDInternalNblBypassUrl = $EWS.InternalNblBypassUrl
            $AutodVDInternalUrl = $EWS.InternalUrl
            $AutodVDExternalUrl = $EWS.ExternalUrl
            $script:html +=
            " <div><b>============================</b></div>
            <div><b>Identity:</b> $AutodVDIdentity</div>
            <div><b>Name:</b> $AutodVDName </div>
            <div><b>InternalAuthenticationMethods:</b> $AutodVDInternalAuthenticationMethods </div>
            <div><b>ExternalAuthenticationMethods:</b> $AutodVDExternalAuthenticationMethods </div>
            <div><b>WSAuthetication:</b> <span style='color:green'>$AutodVDWSAuthetication</span></div>
            <div><b>WindowsAuthentication:</b> <span style='color:green'>$AutodVDWindowsAuthentication</span></div>
            <div><b>OAuthAuthentication:</b> <span style='color:$AutodVDOAuthAuthenticationColor'>$AutodVDOAuthAuthentication</span></div>
            "
            
        }
    }
    else {
        Write-Host -foregroundcolor Red "  InternalAuthenticationMethods seems not to include OAuth Authentication Method."
        $AutodVDIdentity = $EWS.Identity
        $AutodVDName = $EWS.Name
        $AutodVDInternalAuthenticationMethods = $EWS.InternalAuthenticationMethods
        $AutodVDExternalAuthenticationMethods = $EWS.ExternalAuthenticationMethods
        $AutodVDWSAuthetication = $EWS.WSSecurityAuthentication
        $AutodVDWSAutheticationColor = "green"
        $AutodVDOAuthAuthentication = $EWS.OAuthAuthentication
        $AutodVDWindowsAuthentication = $EWS.WindowsAuthentication
        if ($AutodVDWindowsAuthentication -eq "True") {
            $AutodVDWindowsAuthenticationColor = "green"
        }
        else {
            $AutodVDWindowsAuthenticationColor = "red"
        }
        $AutodVDInternalNblBypassUrl = $EWS.InternalNblBypassUrl
        $AutodVDInternalUrl = $EWS.InternalUrl
        $AutodVDExternalUrl = $EWS.ExternalUrl
        $script:html +=
        " <div><b>============================</b></div>
            <div><b>Identity:</b> $AutodVDIdentity</div>
            <div><b>Name:</b> $AutodVDName </div>
            <div><b>InternalAuthenticationMethods:</b> $AutodVDInternalAuthenticationMethods </div>
            <div><b>ExternalAuthenticationMethods:</b> $AutodVDExternalAuthenticationMethods </div>
            <div><b>WSAuthetication:</b> <span style='color:red'>$AutodVDWSAuthetication</span></div>
            <div><b>WindowsAuthentication:</b> <span style='color:$AutodVDWindowsAuthenticationColor'>$AutodVDWindowsAuthentication</span></div>
            <div><b>OAuthAuthentication:</b> <span style='color:$AutodVDOAuthAuthenticationColor'>$AutodVDOAuthAuthentication</span></div>
            "



    }
    Write-Host -foregroundcolor White "`n  ExternalAuthenticationMethods"
    if ($AutoDiscoveryVirtualDirectoryOAuth.ExternalAuthenticationMethods -like "*OAuth*") {
        foreach ( $EWS in $AutoDiscoveryVirtualDirectoryOAuth) {
            Write-Host " $($EWS.Identity) "
            Write-Host -ForegroundColor Green "  ExternalAuthenticationMethods Include OAuth Authentication Method "
        
        
        
        }
    }
    else {
        Write-Host -foregroundcolor Red "  ExternalAuthenticationMethods seems not to include OAuth Authentication Method."
        
        
    }
    Write-Host -ForegroundColor White "`n  WSSecurityAuthentication:"
    if ($AutoDiscoveryVirtualDirectoryOAuth.WSSecurityAuthentication -like "True") {
        #Write-Host -foregroundcolor Green " `n  " $Global:AutoDiscoveryVirtualDirectory.WSSecurityAuthentication
        foreach ( $ADVD in $AutoDiscoveryVirtualDirectoryOAuth) {
            Write-Host " $($ADVD.Identity) "
            Write-Host -ForegroundColor Green "  WSSecurityAuthentication: $($ADVD.WSSecurityAuthentication)"
        
        
        
         
        }
    }
    else {
        Write-Host -foregroundcolor Red "  WSSecurityAuthentication settings are NOT correct."
        foreach ( $ADVD in $AutoDiscoveryVirtualDirectoryOAuth) {
            Write-Host " $($ADVD.Identity) "
            Write-Host -ForegroundColor Red "  WSSecurityAuthentication: $($ADVD.WSSecurityAuthentication).  WSSecurityAuthentication setting should be True."
        



        }
        Write-Host -foregroundcolor White "  Should be True "
    }
    Write-Host -ForegroundColor White "`n  WindowsAuthentication:"
    if ($AutoDiscoveryVirtualDirectoryOAuth.WindowsAuthentication -eq "True") {
        foreach ( $ser in $AutoDiscoveryVirtualDirectoryOAuth) {
            Write-Host " $($ser.Identity) "
            Write-Host -ForegroundColor Green "  WindowsAuthentication: $($ser.WindowsAuthentication)"
            


        
        }
    }
    else {
        Write-Host -foregroundcolor Red " WindowsAuthentication is NOT correct."
        foreach ( $ser in $AutoDiscoveryVirtualDirectoryOAuth) {
            Write-Host " $($ser.Identity)"
            Write-Host -ForegroundColor Red "  WindowsAuthentication: $($ser.WindowsAuthentication)"
            


        
        }
        Write-Host -foregroundcolor White "  Should be True "
    }
    #Write-Host $bar


    $html | Out-File -FilePath $htmlfile



}

Function EWSVirtualDirectoryCheckOAuth {
    Write-Host -foregroundcolor Green " Get-WebServicesVirtualDirectory | Select Identity,Name,ExchangeVersion,*Authentication*,*url"
    Write-Host $bar
    $WebServicesVirtualDirectoryOAuth = Get-WebServicesVirtualDirectory | Select-Object Identity, Name, ExchangeVersion, *Authentication*, *url
    $W = $WebServicesVirtualDirectoryOAuth | Format-List
    $W

    $script:html += "
    <tr>
    <th colspan='2' style='color:white;'>Summary - Get-WebServicesVirtualDirectory</th>
    </tr>
    <tr>
    <td><b>Get-WebServicesVirtualDirectory | Select Identity,Name,ExchangeVersion,*Authentication*,*url</b></td>
    <td >"


    if ($Auth -contains "OAuth") {
    }
    Write-Host $bar
    Write-Host -foregroundcolor Green " Summary - On-Prem Get-WebServicesVirtualDirectory"
    Write-Host $bar
    Write-Host -foregroundcolor White "  InternalAuthenticationMethods"
    if ($WebServicesVirtualDirectoryOAuth.InternalAuthenticationMethods -like "*OAuth*") {
        foreach ( $EWS in $WebServicesVirtualDirectoryOAuth) {
            Write-Host " $($EWS.Identity) "
            Write-Host -ForegroundColor Green "  InternalAuthenticationMethods Include OAuth Authentication Method "


            $EWSVDIdentity = $EWS.Identity
            $EWSVDName = $EWS.Name
            $EWSVDInternalAuthenticationMethods = $EWS.InternalAuthenticationMethods
            $EWSVDExternalAuthenticationMethods = $EWS.ExternalAuthenticationMethods
            $EWSVDWSAuthetication = $EWS.WSSecurityAuthentication
            $EWSVDWSAutheticationColor = "green"
            $EWSVDWindowsAuthentication = $EWS.WindowsAuthentication
            $EWSVDOAuthAuthentication = $EWS.OAuthAuthentication
            if ($EWSVDWindowsAuthentication -eq "True") {
                $EWSVDWindowsAuthenticationColor = "green"
            }
            else {
                $EWSDWindowsAuthenticationColor = "red"
            }
            if ($EWSVDOAuthAuthentication -eq "True") {
                $EWSVDWOAuthAuthenticationColor = "green"
            }
            else {
                $EWSDOAuthAuthenticationColor = "red"
            }
            $EWSVDInternalNblBypassUrl = $EWS.InternalNblBypassUrl
            $EWSVDInternalUrl = $EWS.InternalUrl
            $EWSVDExternalUrl = $EWS.ExternalUrl
            $script:html +=
            " <div><b>============================</b></div>
            <div><b>Identity:</b> $EWSVDIdentity</div>
            <div><b>Name:</b> $EWSVDName </div>
            <div><b>InternalAuthenticationMethods:</b> $EWSVDInternalAuthenticationMethods </div>
            <div><b>ExternalAuthenticationMethods:</b> $EWSVDExternalAuthenticationMethods </div>
            <div><b>WSAuthetication:</b> <span style='color:green'>$EWSVDWSAuthetication</span></div>
            <div><b>WindowsAuthentication:</b> <span style='color:$EWSVDWindowsAuthenticationColor'>$EWSVDWindowsAuthentication</span></div>
            <div><b>OAuthAuthentication:</b> <span style='color:$EWSVDWOAuthAuthenticationColor'>$EWSVDOAuthAuthentication</span></div>
            <div><b>InternalUrl:</b> $EWSVDInternalUrl </div>
            <div><b>ExternalUrl:</b> $EWSVDExternalUrl </div>  "



        }
    }
    else {
        Write-Host -foregroundcolor Red "  InternalAuthenticationMethods seems not to include OAuth Authentication Method."
        $EWSVDIdentity = $EWS.Identity
        $EWSVDName = $EWS.Name
        $EWSVDInternalAuthenticationMethods = $EWS.InternalAuthenticationMethods
        $EWSVDExternalAuthenticationMethods = $EWS.ExternalAuthenticationMethods
        $EWSVDWSAuthetication = $EWS.WSSecurityAuthentication
        $EWSVDWSAutheticationColor = "green"
        $EWSVDWindowsAuthentication = $EWS.WindowsAuthentication
        $EWSVDOAuthAuthentication = $EWS.OAuthAuthentication
        if ($EWSVDWindowsAuthentication -eq "True") {
            $EWSVDWindowsAuthenticationColor = "green"
        }
        else {
            $EWSDWindowsAuthenticationColor = "red"
        }
        if ($EWSVDOAuthAuthentication -eq "True") {
            $EWSVDWOAuthAuthenticationColor = "green"
        }
        else {
            $EWSDOAuthAuthenticationColor = "red"
        }
        $EWSVDInternalNblBypassUrl = $EWS.InternalNblBypassUrl
        $EWSVDInternalUrl = $EWS.InternalUrl
        $EWSVDExternalUrl = $EWS.ExternalUrl
        $script:html +=
        " <div><b>============================</b></div>
            <div><b>Identity:</b> $EWSVDIdentity</div>
            <div><b>Name:</b> $EWSVDName </div>
            <div><b>InternalAuthenticationMethods:</b> $EWSVDInternalAuthenticationMethods </div>
            <div><b>ExternalAuthenticationMethods:</b> $EWSVDExternalAuthenticationMethods </div>
            <div><b>WSAuthetication:</b> <span style='color:red'>$EWSVDWSAuthetication</span></div>
            <div><b>WindowsAuthentication:</b> <span style='color:$EWSVDWindowsAuthenticationColor'>$EWSVDWindowsAuthentication</span></div>
            <div><b>OAuthAuthentication:</b> <span style='color:$EWSVDWOAuthAuthenticationColor'>$EWSVDOAuthAuthentication</span></div>
            <div><b>InternalUrl:</b> $EWSVDInternalUrl </div>
            <div><b>ExternalUrl:</b> $EWSVDExternalUrl </div>  "



    }
    Write-Host -foregroundcolor White "`n  ExternalAuthenticationMethods"
    if ($WebServicesVirtualDirectoryOAuth.ExternalAuthenticationMethods -like "*OAuth*") {
        foreach ( $EWS in $WebServicesVirtualDirectoryOAuth) {
            Write-Host " $($EWS.Identity) "
            Write-Host -ForegroundColor Green "  ExternalAuthenticationMethods Include OAuth Authentication Method "




        }
    }
    else {
        Write-Host -foregroundcolor Red "  ExternalAuthenticationMethods seems not to include OAuth Authentication Method."





    }
    Write-Host -ForegroundColor White "`n  WSSecurityAuthentication:"
    if ($WebServicesVirtualDirectoryOAuth.WSSecurityAuthentication -like "True") {
        foreach ( $EWS in $WebServicesVirtualDirectoryOAuth) {
            Write-Host " $($EWS.Identity) "
            Write-Host -ForegroundColor Green "  WSSecurityAuthentication: $($EWS.WSSecurityAuthentication) "




        }
    }
    else {
        Write-Host -foregroundcolor Red "  WSSecurityAuthentication is NOT correct."
        foreach ( $EWS in $WebServicesVirtualDirectoryOauth) {
            Write-Host " $($EWS.Identity) "
            Write-Host -ForegroundColor Red "  WSSecurityAuthentication: $($EWS.WSSecurityAuthentication)"




        }
        Write-Host -foregroundcolor White "  Should be True"
    }
    #Write-Host $bar
    Write-Host -ForegroundColor White "`n  WindowsAuthentication:"
    if ($WebServicesVirtualDirectoryOauth.WindowsAuthentication -eq "True") {
        foreach ( $ser in $WebServicesVirtualDirectoryOauth) {
            Write-Host " $($ser.Identity) "
            Write-Host -ForegroundColor Green "  WindowsAuthentication: $($ser.WindowsAuthentication)"



        }
    }
    else {
        Write-Host -foregroundcolor Red " WindowsAuthentication is NOT correct."
        foreach ( $ser in $WebServicesVirtualDirectoryOauth) {
            Write-Host " $($ser.Identity)"
            Write-Host -ForegroundColor Red "  WindowsAuthentication: $($ser.WindowsAuthentication)"





        }
        Write-Host -foregroundcolor White "  Should be True "
    }
    $html | Out-File -FilePath $htmlfile
}

Function AvailabilityAddressSpaceCheckOAuth {
    Write-Host -foregroundcolor Green " Get-AvailabilityAddressSpace $exchangeOnlineDomain | Select ForestName, UserName, UseServiceAccount, AccessMethod, ProxyUrl, Name"
    Write-Host $bar
    $AvailabilityAddressSpace = Get-AvailabilityAddressSpace $exchangeOnlineDomain | Select-Object ForestName, UserName, UseServiceAccount, AccessMethod, ProxyUrl, Name
    $AAS = $AvailabilityAddressSpace | Format-List
    $AAS
    if ($Auth -contains "OAuth") {
    }
    Write-Host $bar
    Write-Host -foregroundcolor Green " Summary - On-Prem Availability Address Space"
    Write-Host $bar
    Write-Host -foregroundcolor White " ForestName: "
    if ($AvailabilityAddressSpace.ForestName -like $ExchangeOnlineDomain) {
        Write-Host -foregroundcolor Green " "$AvailabilityAddressSpace.ForestName
        $tdAvailabilityAddressSpaceForestName = $AvailabilityAddressSpace.ForestName
        $tdAvailabilityAddressSpaceForestNameColor = "green"


    }
    else {
        Write-Host -foregroundcolor Red " ForestName is NOT correct. "
        Write-Host -foregroundcolor White " Should be $ExchaneOnlineDomain "
        $tdAvailabilityAddressSpaceForestName = $AvailabilityAddressSpace.ForestName
        $tdAvailabilityAddressSpaceForestNameColor = "red"


    }
    Write-Host -foregroundcolor White " UserName: "
    if ($AvailabilityAddressSpace.UserName -like "") {
        Write-Host -foregroundcolor Green "  Blank "
        $tdAvailabilityAddressSpaceUserName = "  Blank. This is the correct value. "
        $tdAvailabilityAddressSpaceUserNameColor = "green"


    }
    else {
        Write-Host -foregroundcolor Red "  UserName is NOT correct. "
        Write-Host -foregroundcolor White "  Should be blank "
        $tdAvailabilityAddressSpaceUserName = "  Blank. This is the correct value. "
        $tdAvailabilityAddressSpaceUserNameColor = "red"


    }
    Write-Host -foregroundcolor White " UseServiceAccount: "
    if ($AvailabilityAddressSpace.UseServiceAccount -like "True") {
        Write-Host -foregroundcolor Green "  True "
        $tdAvailabilityAddressSpaceUseServiceAccount = $AvailabilityAddressSpace.UseServiceAccount
        $tdAvailabilityAddressSpaceUseServiceAccountColor = "green"



    }
    else {
        Write-Host -foregroundcolor Red "  UseServiceAccount is NOT correct."
        Write-Host -foregroundcolor White "  Should be True "
        $tdAvailabilityAddressSpaceUseServiceAccount = "$($tAvailabilityAddressSpace.UseServiceAccount). Should be True"
        $tdAvailabilityAddressSpaceUseServiceAccountColor = "red"



    }
    Write-Host -foregroundcolor White " AccessMethod: "
    if ($AvailabilityAddressSpace.AccessMethod -like "InternalProxy") {
        Write-Host -foregroundcolor Green "  InternalProxy "
        $tdAvailabilityAddressSpaceAccessMethod = $AvailabilityAddressSpace.AccessMethod
        $tdAvailabilityAddressSpaceAccessMethodColor = "green"
    
    }
    else {
        Write-Host -foregroundcolor Red "  AccessMethod is NOT correct. "
        Write-Host -foregroundcolor White "  Should be InternalProxy "
        $tdAvailabilityAddressSpaceAccessMethod = $AvailabilityAddressSpace.AccessMethod
        $tdAvailabilityAddressSpaceAccessMethodColor = "red"
    
    
    }
    Write-Host -foregroundcolor White " ProxyUrl: "
    if ($AvailabilityAddressSpace.ProxyUrl -like $exchangeOnPremEWS) {
        Write-Host -foregroundcolor Green " "$AvailabilityAddressSpace.ProxyUrl
        $tdAvailabilityAddressSpaceProxyUrl = $AvailabilityAddressSpace.ProxyUrl
        $tdAvailabilityAddressSpaceProxyUrlColor = "green"
    
    
    }
    else {
        Write-Host -foregroundcolor Red "  ProxyUrl is NOT correct. "
        Write-Host -foregroundcolor White "  Should be $exchangeOnPremEWS"
        $tdAvailabilityAddressSpaceProxyUrl = $AvailabilityAddressSpace.ProxyUrl
        $tdAvailabilityAddressSpaceProxyUrlColor = "red"
    
    }


    $Script:html += "
      <tr>
      <th colspan='2' style='color:white;'>Summary - Get-AvailabilityAddressSpace</th>
    </tr>
    <tr>
      <td><b>  Get-AvailabilityAddressSpace $exchangeOnlineDomain | Select ForestName, UserName, UseServiceAccount, AccessMethod, ProxyUrl, Name</b></td>
      <td>
        <div><b>AddressSpaceForestName:</b><span style='color: $tdAvailabilityAddressSpaceForestNameColor'>$($tdAvailabilityAddressSpaceForestName)</span></div>
        <div><b>AddressSpaceUserName:</b><span style='color: $tdAvailabilityAddressSpaceUserNameColor'>$($tdAvailabilityAddressSpaceUserName)</span></div>
        <div><b>UseServiceAccount:</b><span style='color:$tdAvailabilityAddressSpaceUseServiceAccountColor;'>$( $tdAvailabilityAddressSpaceUseServiceAccount)</span></div>
        <div><b>AccessMethod:</b><span style='color: $tdAvailabilityAddressSpaceAccessMethodColor;'>$($tdAvailabilityAddressSpaceAccessMethod)</span></div>
        <div><b>ProxyUrl:</b><span style='color: $tdAvailabilityAddressSpaceProxyUrlColor;'>$($tdAvailabilityAddressSpaceProxyUrl)</span></div>
        
        
      </td>
    </tr>
  "

    $html | Out-File -FilePath $htmlfile
}

Function OAuthConnectivityCheck {
    Write-Host -foregroundcolor Green " Test-OAuthConnectivity -Service EWS -TargetUri https://outlook.office365.com/EWS/Exchange.asmx -Mailbox $useronprem"
    Write-Host $bar
    #$OAuthConnectivity = Test-OAuthConnectivity -Service EWS -TargetUri https://outlook.office365.com/EWS/Exchange.asmx -Mailbox $useronprem | fl
    #$OAuthConnectivity
    $OAuthConnectivity = Test-OAuthConnectivity -Service EWS -TargetUri https://outlook.office365.com/EWS/Exchange.asmx -Mailbox $useronprem
    if ($OAuthConnectivity.ResultType -eq 'Success' ) {
        #$OAuthConnectivity.ResultType
    }
    else {
        $OAuthConnectivity
    }
    #$OAC = $OAuthConnectivity | Format-List
    #$OAC
    #$bar
    #$OAuthConnectivity.Detail.FullId
    #$bar
    if ($OAuthConnectivity.Detail.FullId -like '*(401) Unauthorized*') {
        write-host -ForegroundColor Red "Error: The remote server returned an error: (401) Unauthorized"
        if ($OAuthConnectivity.Detail.FullId -like '*The user specified by the user-context in the token does not exist*') {
            write-host -ForegroundColor Yellow "The user specified by the user-context in the token does not exist"
            write-host "Please run Test-OAuthConnectivity with a different Exchange On Premises Mailbox"
        }
    }

    # Write-Host $bar
    #$OAuthConnectivity.detail.LocalizedString
    Write-Host -foregroundcolor Green " Summary - Test OAuth Connectivity"
    Write-Host $bar
    if ($OAuthConnectivity.ResultType -like "Success") {
        Write-Host -foregroundcolor Green "$($OAuthConnectivity.ResultType). OAuth Test was completed successfully "
        $OAuthConnectivityResultType = " OAuth Test was completed successfully "
        $OAuthConnectivityResultTypeColor = "green"
    }
    else {
        Write-Host -foregroundcolor Red " $OAuthConnectivity.ResultType - OAuth Test was completed with Error. "
        Write-Host -foregroundcolor White " Please rerun Test-OAuthConnectivity -Service EWS -TargetUri https://outlook.office365.com/EWS/Exchange.asmx -Mailbox <On Premises Mailbox> | fl to confirm the test failure"
        $OAuthConnectivityResultType = " <div>OAuth Test was completed with Error.</div><div>Please rerun Test-OAuthConnectivity -Service EWS -TargetUri https://outlook.office365.com/EWS/Exchange.asmx -Mailbox <On Premises Mailbox> | fl to confirm the test failure</div>"
        $OAuthConnectivityResultTypeColor = "red"


        
    }
    #Write-Host -foregroundcolor Green " Note:"
    #Write-Host -foregroundcolor Yellow " You can ignore the warning 'The SMTP address has no mailbox associated with it'"
    #Write-Host -foregroundcolor Yellow " when the Test-OAuthConnectivity returns a Success"
    Write-Host -foregroundcolor Green " Reference: "
    Write-Host -foregroundcolor White " Configure OAuth authentication between Exchange and Exchange Online organizations"
    Write-Host -foregroundcolor Yellow " https://technet.microsoft.com/en-us/library/dn594521(v=exchg.150).aspx"

    $Script:html += "
      <tr>
      <th colspan='2' style='color:white;'>Summary - Test-OAuthConnectivity</th>
    </tr>
    <tr>
      <td><b>  Test-OAuthConnectivity -Service EWS -TargetUri https://outlook.office365.com/EWS/Exchange.asmx -Mailbox $useronprem | fl</b></td>
      <td>
        <div><b>Result:</b><span style='color: $OAuthConnectivityResultTypeColor'> $OAuthConnectivityResultType</span></div>
        
        
        
      </td>
    </tr>
  "

    $html | Out-File -FilePath $htmlfile
}

#endregion

# EXO FUNCTIONS

#region ExoDauthFuntions

Function ExoOrgRelCheck () {
    Write-Host $bar
    Write-Host -foregroundcolor Green " Get-OrganizationRelationship  | Where{($_.DomainNames -like $ExchangeOnPremDomain )} | Select Identity,DomainNames,FreeBusy*,Target*,Enabled"
    Write-Host $bar
    $ExoOrgRel
    Write-Host $bar
    Write-Host  -foregroundcolor Green " Summary - Organization Relationship"
    Write-Host $bar
    Write-Host  " Domain Names:"
    #Write-Host  " Domain Names:"$exoOrgRel.DmainNames[0]
    #Write-Host  " Org Rel:"$exoOrgRel
    if ($exoOrgRel.DomainNames -like $ExchangeOnPremDomain) {
        Write-Host -foregroundcolor Green "  Domain Names Include the $ExchangeOnPremDomain Domain"
        $tdEXOOrgRelDomainNames = $exoOrgRel.DomainNames
        $tdEXOOrgRelDomainNamesColor = "green"

    }
    else {
        Write-Host -foregroundcolor Red "  Domain Names do Not Include the $ExchangeOnPremDomain Domain"
        $exoOrgRel.DomainNames

        $tdEXOOrgRelDomainNames = "$($exoOrgRel.DomainNames) - Domain Names do Not Include the $ExchangeOnPremDomain Domain"
        $tdEXOOrgRelDomainNamesColor = "green"


    }
    #FreeBusyAccessEnabled
    Write-Host  " FreeBusyAccessEnabled:"
    if ($exoOrgRel.FreeBusyAccessEnabled -like "True" ) {
        Write-Host -foregroundcolor Green "  FreeBusyAccessEnabled is set to True"
        $tdExoOrgRelFreeBusyAccessEnabled = "$($exoOrgRel.FreeBusyAccessEnabled)"
        $tdExoOrgRelFreeBusyAccessEnabledColor = "green" 


    }
    else {
        Write-Host -foregroundcolor Red "  FreeBusyAccessEnabled : False"
        #$countOrgRelIssues++
        $tdExoOrgRelFreeBusyAccessEnabled = "$($exoOrgRel.FreeBusyAccessEnabled). Free busy access is not enabled for the organization Relationship"
        $tdExoOrgRelFreeBusyAccessEnabledColor = "Red" 


    }
    #FreeBusyAccessLevel
    Write-Host  " FreeBusyAccessLevel:"
    if ($exoOrgRel.FreeBusyAccessLevel -like "AvailabilityOnly" ) {
        Write-Host -foregroundcolor Green "  FreeBusyAccessLevel is set to AvailabilityOnly"
        $tdExoOrgRelFreeBusyAccessLevel = "$($exoOrgRel.FreeBusyAccessLevel)"
        $tdExoOrgRelFreeBusyAccessLevelColor = "green" 



    }
    if ($exoOrgRel.FreeBusyAccessLevel -like "LimitedDetails" ) {
        Write-Host -foregroundcolor Green "  FreeBusyAccessLevel is set to LimitedDetails"
        $tdExoOrgRelFreeBusyAccessLevel = "$($exoOrgRel.FreeBusyAccessLevel)"
        $tdExoOrgRelFreeBusyAccessLevelColor = "green" 



    }
    #fix porque este else só respeita o if anterior
    if ($exoOrgRel.FreeBusyAccessLevel -NE "AvailabilityOnly" -AND $exoOrgRel.FreeBusyAccessLevel -NE "LimitedDetails") {
        Write-Host -foregroundcolor Red "  FreeBusyAccessEnabled : False"
        #$countOrgRelIssues++
        $tdExoOrgRelFreeBusyAccessLevel = "$($exoOrgRel.FreeBusyAccessLevel)"
        $tdExoOrgRelFreeBusyAccessLevelColor = "red" 



    }
    #TargetApplicationUri
    Write-Host  " TargetApplicationUri:"
   # Write-host $fedinfoTargetApplicationUri
    $a= "FYDIBOHF25SPDLT." + $ExchangeOnPremDomain
   #write-host $a 
    if ($exoOrgRel.TargetApplicationUri -like $fedtrust.ApplicationUri) {
        Write-Host -foregroundcolor Green "  TargetApplicationUri is" $fedtrust.ApplicationUri.originalstring
        $tdEXOOrgRelTargetApplicationUri = "  TargetApplicationUri is $($fedtrust.ApplicationUri.originalstring)"
        $tdEXOOrgRelTargetApplicationUriColor = "green"


    }
    else {
        Write-Host -foregroundcolor Red "  TargetApplicationUri should be " $a
        #$countOrgRelIssues++
        $tdEXOOrgRelTargetApplicationUri = "  TargetApplicationUri should be $a. Please check if Exchange On Premise Federation is correctly configured."
        $tdEXOOrgRelTargetApplicationUriColor = "red"



    }
    #TargetSharingEpr
    Write-Host  " TargetSharingEpr:"
    if ([string]::IsNullOrWhitespace($exoOrgRel.TargetSharingEpr)) {
        Write-Host -foregroundcolor Green "  TargetSharingEpr is blank. This is the standard Value."
        $tdEXOOrgRelTargetSharingEpr = "TargetSharingEpr is blank. This is the standard Value."
        $tdEXOOrgRelTargetSharingEprColor = "green"


    }
    else {
        Write-Host -foregroundcolor Red "  TargetSharingEpr should be blank. If it is set, it should be the On-Premises Exchange servers EWS ExternalUrl endpoint."
        #$countOrgRelIssues++
        $tdEXOOrgRelTargetSharingEpr = "  TargetSharingEpr should be blank. If it is set, it should be the On-Premises Exchange servers EWS ExternalUrl endpoint."
        $tdEXOOrgRelTargetSharingEprColor = "red"



    }
    #TargetAutodiscoverEpr:
    Write-Host  " TargetAutodiscoverEpr:"
    #Write-Host  "  OrgRel: " $exoOrgRel.TargetAutodiscoverEpr
    #Write-Host  "  FedInfo: " $fedinfoEOP
    #Write-Host  "  FedInfoEPR: " $fedinfoEOP.TargetAutodiscoverEpr
    if ($exoOrgRel.TargetAutodiscoverEpr -like $fedinfoEOP.TargetAutodiscoverEpr) {
        Write-Host -foregroundcolor Green "  TargetAutodiscoverEpr is" $exoOrgRel.TargetAutodiscoverEpr

        $tdexoOrgRelTargetAutodiscoverEpr = $exoOrgRel.TargetAutodiscoverEpr
        $tdexoOrgRelTargetAutodiscoverEprColor = "green" 

    }
    else {
        Write-Host -foregroundcolor Red "  TargetAutodiscoverEpr is not" $fedinfoEOP.TargetAutodiscoverEpr
        #$countOrgRelIssues++
        $tdexoOrgRelTargetAutodiscoverEpr = "  TargetAutodiscoverEpr is not $($fedinfoEOP.TargetAutodiscoverEpr)"
        $tdexoOrgRelTargetAutodiscoverEprColor = "red"



    }
    #Enabled
    Write-Host  " Enabled:"
    if ($exoOrgRel.enabled -like "True" ) {
        Write-Host -foregroundcolor Green "  Enabled is set to True"
        $tdEXOOrgRelEnabled = "  True"
        $tdEXOOrgRelEnabledColor = "green"


    }
    else {
        Write-Host -foregroundcolor Red "  Enabled is set to False."
        $tdEXOOrgRelEnabled = "  False"
        $tdEXOOrgRelEnabledColor = "red"






    }




    $script:html += "

    <div class='Black'><p></p></div>
    <div class='Black'><p></p></div>


     <tr>
        <th colspan='2' style='text-align:center; color:white;'>Exchange Online DAuth Configuration</th>
     </tr>
      <tr>
      <th colspan='2' style='color:white;'>Summary - Get-OrganizationRelationship</th>
    </tr>
    <tr>
      <td><b>  Get-OrganizationRelationship  | Where{($_.DomainNames -like $ExchangeOnPremDomain )} | Select Identity,DomainNames,FreeBusy*,Target*,Enabled</b></td>
      <td>
        <div><b>Domain Names:</b><span >$($tdEXOOrgRelDomainNames)</span></div>
        <div><b>FreeBusyAccessEnabled:</b><span style='color: $tdExoOrgRelFreeBusyAccessEnabledColor'>$($tdExoOrgRelFreeBusyAccessEnabled)</span></div>
        <div><b>FreeBusyAccessLevel::</b><span style='color:$tdEXOOrgRelFreeBusyAccessLevelColor;'>$( $tdEXOOrgRelFreeBusyAccessLevel)</span></div>
        <div><b>TargetApplicationUri:</b><span style='color: $tdEXOOrgRelTargetApplicationUriColor;'>$($tdEXOOrgRelTargetApplicationUri)</span></div>
        <div><b>TargetOwaURL:</b><span >$($tdEXOOrgRelTargetOwaUrl)</span></div>
        <div><b>TargetSharingEpr:</b><span style='color: $tdEXOOrgRelTargetSharingEprColor'>$($tdEXOOrgRelTargetSharingEpr)</span></div>
        <div><b>TargetAutodiscoverEpr:</b><span style='color:$tdEXOOrgRelFreeBusyAccessScopeColor;'>$( $tdEXOOrgRelFreeBusyAccessScope)</span></div>
        <div><b>Enabled:</b><span style='color: $tdEXOOrgRelEnabledColor;'>$($tdEXOOrgRelEnabled)</span></div>
        
        
      </td>
    </tr>
  "


    
    $html | Out-File -FilePath $htmlfile



}

Function EXOFedOrgIdCheck {
    Write-Host -foregroundcolor Green " Get-FederatedOrganizationIdentifier | select AccountNameSpace,Domains,Enabled"
    Write-Host $bar
    $exoFedOrgId = Get-FederatedOrganizationIdentifier | Select-Object AccountNameSpace, Domains, Enabled
    #$IntraOrgConCheck
    $efedorgid = $exoFedOrgId | Format-List
    $efedorgid
    Write-Host $bar
    Write-Host -foregroundcolor Green " Summary - Online Federated Organization Identifier"
    Write-Host $bar
    Write-Host -foregroundcolor White " Domains: "
    if ($exoFedOrgId.Domains -like "*$ExchangeOnlineDomain*") {
        Write-Host -foregroundcolor Green " " $exoFedOrgId.Domains
        $tdexoFedOrgIdDomains = $exoFedOrgId.Domains
        $tdexoFedOrgIdDomainsColor = "green" 

    }
    else {
        Write-Host -foregroundcolor Red " Domains are NOT correct."
        Write-Host -foregroundcolor White " Should contain the $ExchangeOnlinemDomain"
        $tdexoFedOrgIdDomains = "$($exoFedOrgId.Domains) . Domains Should contain the $ExchangeOnlinemDomain"
        $tdexoFedOrgIdDomainsColor = "red" 

    }
    Write-Host -foregroundcolor White " Enabled: "
    if ($exoFedOrgId.Enabled -like "True") {
        Write-Host -foregroundcolor Green "  True "
        $tdexoFedOrgIdEnabled = $exoFedOrgId.Enabled
        $tdexoFedOrgIdEnabledColor = "green"
    }
    else {
        Write-Host -foregroundcolor Red "  Enabled is NOT correct."
        Write-Host -foregroundcolor White " Should be True"
        $tdexoFedOrgIdEnabled = $exoFedOrgId.Enabled
        $tdexoFedOrgIdEnabledColor = "green"

    }



    

    $script:html += "
    <tr>
      <th colspan='2' style='color:white;'>Summary - Get-FederatedOrganizationIdentifier</th>
    </tr>
    <tr>
      <td><b>  Get-FederatedOrganizationIdentifier | select AccountNameSpace,Domains,Enabled</b></td>
      <td>
        <div><b>Domains:</b><span style='color: $tdexoFedOrgIdDomainsColor;'>$($tdexoFedOrgIdDomains)</span></div>
        <div><b>Enabled:</b><span style='color: $tdexoFedOrgIdEnabledColor;'>$($tdexoFedOrgIdEnabled)</span></div>
        
        
      </td>
    </tr>
  "


    
    $html | Out-File -FilePath $htmlfile


}

Function EXOTestOrgRelCheck {
    $exoIdentity = $ExoOrgRel.Identity
    
    $exoOrgRelTragetApplicationUri = $exoOrgRel.TargetApplicationUri
    $exoOrgRelTragetOWAurl = $ExoOrgRel.TargetOwaURL

    $script:html += "
        <tr>
            <th colspan='2' style='color:white;'>Summary - Test-OrganizationRelationship</th>
        </tr>
        <tr>
            <td><b>  Test-OrganizationRelationship -Identity $exoIdentity -UserIdentity $UserOnline</b></td>
        <td>"



    Write-Host -foregroundcolor Green " Test-OrganizationRelationship -Identity $exoIdentity -UserIdentity $UserOnline"
    Write-Host $bar

    if ((![string]::IsNullOrWhitespace($exoOrgRelTragetApplicationUri)) -and (![string]::IsNullOrWhitespace($exoOrgRelTragetOWAUrl))) {
        $exotestorgrel = Test-OrganizationRelationship -Identity $exoIdentity -UserIdentity $UserOnline -WarningAction SilentlyContinue
    
        $i = 2
    

        while ($i -lt $exotestorgrel.Length) {
            $element = $exotestorgrel[$i]
        
            $aux = "0"
  
            if ($element -like "*RESULT:*" -and $aux -like "0") {   
                $el = $element.TrimStart()
                if ($element -like "*Success.*") {
                    Write-Host -ForegroundColor Green "  $el"
                    $Script:html += "
            <div> <b> $exotestorgrelStep </b> <span style='color:green'>&emsp; $el</span>"
                    $aux = "1"
                }
                elseif ($element -like "*Error*" -or $element -like "*Unable*") {
                    $Script:html += "
                <div> <b> $exotestorgrelStep </b> <span style='color:red'>&emsp; $el</span>"
                    Write-Host -ForegroundColor Red "  $el"
                    $aux = "1"
                }
            }
            elseif ($aux -like "0" ) {
                if ($element -like "*STEP*" -or $element -like "*Complete*") {
                    Write-Host -ForegroundColor White "  $element"
                    $Script:html += "
               <p></p>
                <div> <b> $exotestorgrelStep </b> <span style='color:black'> $element</span></div>"
                    $aux = "1"
                }
                else {
                    $ID = $element.ID
                    $Status = $element.Status
                    $Description = $element.Description 
                    if (![string]::IsNullOrWhitespace($ID)) {
                        Write-Host -ForegroundColor White "`n  ID         : $ID"
                        $Script:html += "<div> <b>&emsp; ID &emsp;&emsp;&emsp;&emsp;&emsp;&nbsp;:</b> <span style='color:black'> $ID</span></div>"
                        if ($Status -like "*Success*") {
                            Write-Host -ForegroundColor White "  Status     : $Status"
                            $Script:html += "<div> <b>&emsp; Status &emsp;&emsp;&ensp;&ensp;:</b> <span style='color:green'> $Status</span></div>"
                        }
                    

                        if ($status -like "*error*") {
                            Write-Host -ForegroundColor White "  Status     : $Status"
                            $Script:html += "<div> <b>&emsp; Status &emsp;&emsp;&ensp;&ensp;:</b> <span style='color:red'> $Status</span></div>"
                        }

                        Write-Host -ForegroundColor White "  Description: $Description"
                        $Script:html += "<div> <b>&emsp; Description :</b> <span style='color:black'> $Description</span></div>"
                    }
                    #$element
                    $aux = "1"
                }
        
            }
       
        
        
    
            $i++
        }
    }
 
    
    elseif ((([string]::IsNullOrWhitespace($exoOrgRelTragetApplicationUri)) -and ([string]::IsNullOrWhitespace($exoOrgRelTragetOWAUrl)))) {
        <# Action when all if and elseif conditions are false #>
        Write-Host -ForegroundColor Red "  Error: Exchange Online Test-OrganizationRelationship cannot be run if the Organization Relationship TragetApplicationUri and TargetOwaURL are not set"
        $Script:html += "
    <div> <span style='color:red'>&emsp; Exchange Online Test-OrganizationRelationship cannot be run if the Organization Relationship TragetApplicationUri and TargetOwaURL are not set</span>"
    }
    elseif ((([string]::IsNullOrWhitespace($exoOrgRelTragetApplicationUri)) )) {
        <# Action when all if and elseif conditions are false #>
        Write-Host -ForegroundColor Red "  Error: Exchange Online Test-OrganizationRelationship cannot be run if the Organization Relationship TragetApplicationUri is not set"
        $Script:html += "
    <div> <span style='color:red'>&emsp; Exchange Online Test-OrganizationRelationship cannot be run if the Organization Relationship TragetApplicationUri is not set</span>"
    }
    elseif ((([string]::IsNullOrWhitespace($exoOrgRelTragetApplicationUri)) )) {
        <# Action when all if and elseif conditions are false #>
        Write-Host -ForegroundColor Red "  Error: Exchange Online Test-OrganizationRelationship cannot be run if the Organization Relationship TargetOwaURL is not set"
        $Script:html += "
    <div> <span style='color:red'>&emsp; Exchange Online Test-OrganizationRelationship cannot be run if the Organization Relationship TragetApplicationUri is not set</span>"
    }



    $Script:html += "</td>
    </tr>"
    
    $html | Out-File -FilePath $htmlfile

}

Function SharingPolicyCheck {
    Write-host $bar
    Write-Host -foregroundcolor Green " Get-SharingPolicy | select Domains,Enabled,Name,Identity"
    Write-Host $bar
    $Script:SPOnline = Get-SharingPolicy | Select-Object  Domains, Enabled, Name, Identity
    $SPOnline | Format-List
    
        #creating variables and setting uniform variable names
    $domain1 = (($SPOnline.domains[0] -split ":") -split " ")
    $domain2 = (($SPOnline.domains[1] -split ":") -split " ")
    $SPOnpremDomain1 = $SPOnprem.Domains.Domain[0]
    $SPOnpremAction1 = $SPOnprem.Domains.Actions[0]
    $SPOnpremDomain2 = $SPOnprem.Domains.Domain[1]
    $SPOnpremAction2 = $SPOnprem.Domains.Actions[1]
    $SPOnlineDomain1 = $domain1[0]
    $SPOnlineAction1 = $domain1[1]
    $SPOnlineDomain2 = $domain2[0]
    $SPOnlineAction2 = $domain2[1]


    Write-Host $bar
    Write-Host -foregroundcolor Green " Summary - Sharing Policy"
    Write-Host $bar
    Write-Host -foregroundcolor White " Exchange On Premises Sharing domains:`n"
    Write-Host -foregroundcolor White "  Domain:"
    Write-Host "   " $SPOnpremDomain1
    Write-Host -foregroundcolor White "  Action:"
    Write-Host "   " $SPOnpremAction1
    Write-Host -foregroundcolor White "  Domain:"
    Write-Host "   " $SPOnpremDomain2
    Write-Host -foregroundcolor White "  Action:"
    Write-Host "   " $SPOnpremAction2
    Write-Host -ForegroundColor White "`n  Exchange Online Sharing Domains: `n"
    Write-Host -foregroundcolor White "  Domain:"
    Write-Host "   " $SPOnlineDomain1
    Write-Host -foregroundcolor White "  Action:"
    Write-Host "   " $SPOnlineAction1
    Write-Host -foregroundcolor White "  Domain:"
    Write-Host "   " $SPOnlineDomain2
    Write-Host -foregroundcolor White "  Action:"
    Write-Host "   " $SPOnlineAction2
    #Write-Host $bar

    if ($SPOnpremDomain1 -eq $SPOnlineDomain1 -and $SPOnpremAction1 -eq $SPOnlineAction1)
    {
         if ($SPOnpremDomain2 -eq $SPOnlineDomain2 -and $SPOnpremAction2 -eq $SPOnlineAction2)
             { Write-Host -foregroundcolor Green "`n  Exchange Online Sharing Policy Domains match Exchange On Premise Sharing Policy Domains"
        $tdSharpingPolicyCheck = "`n  Exchange Online Sharing Policy matches Exchange On Premise Sharing Policy Domain"
        $tdSharpingPolicyCheckColor = "green"}

        else {Write-Host -foregroundcolor Red "`n   Sharing Domains appear not to be correct."
        Write-Host -foregroundcolor White "   Exchange Online Sharing Policy Domains appear not to match Exchange On Premise Sharing Policy Domains"
        $tdSharpingPolicyCheck = "`n  Exchange Online Sharing Policy Domains not match Exchange On Premise Sharing Policy Domains"
        $tdSharpingPolicyCheckColor = "red"}
    }
    elseif ($SPOnpremDomain1 -eq $SPOnlineDomain2 -and $SPOnpremAction1 -eq $SPOnlineAction2)
    { 
        if ($SPOnpremDomain2 -eq $SPOnlineDomain1 -and $SPOnpremAction2 -eq $SPOnlineAction1)
            { Write-Host -foregroundcolor Green "`n  Exchange Online Sharing Policy Domains match Exchange On Premise Sharing Policy Domains"
        $tdSharpingPolicyCheck = "`n  Exchange Online Sharing Policy matches Exchange On Premise Sharing Policy Domain"
        $tdSharpingPolicyCheckColor = "green"}

        else {Write-Host -foregroundcolor Red "`n   Sharing Domains appear not to be correct."
        Write-Host -foregroundcolor White "   Exchange Online Sharing Policy Domains appear not to match Exchange On Premise Sharing Policy Domains"
        $tdSharpingPolicyCheck = "`n  Exchange Online Sharing Policy Domains not match Exchange On Premise Sharing Policy Domains"
        $tdSharpingPolicyCheckColor = "red"}
    }
    else {
    Write-Host -foregroundcolor Red "`n   Sharing Domains appear not to be correct."
        Write-Host -foregroundcolor White "   Exchange Online Sharing Policy Domains appear not to match Exchange On Premise Sharing Policy Domains"
        $tdSharpingPolicyCheck = "`n  Exchange Online Sharing Policy Domains not match Exchange On Premise Sharing Policy Domains"
        $tdSharpingPolicyCheckColor = "red"
    }
    
    $bar

    $script:html += "
    <tr>
      <th colspan='2' style='color:white;'>Summary - Get-SharingPolicy</th>
    </tr>
    <tr>
      <td><b>  Get-SharingPolicy | select Domains,Enabled,Name,Identity</b></td>
      <td>
        <div><b>Exchange On Premises Sharing domains:<b></div> 
        <div><b>Domain:</b>$($SPOnprem.Domains.Domain[0])</div>
        <div><b>Action:</b>$($SPOnprem.Domains.Actions[0])</div>
        <div><b>Domain:</b>$($SPOnprem.Domains.Domain[1])</div>
        <div><b>Action:</b>$($SPOnprem.Domains.Actions[1])</div>
        <div><p></p></div> 
        <div><b>Exchange Online Sharing domains:<b></div> 
        <div><b>Domain:</b>$($domain1[0])</div>
        <div><b>Action:</b>$( $domain1[1])</div>
        <div><b>Domain:</b>$($domain2[0])</div>
        <div><b>Action:</b>$( $domain2[1])</div>
        <div><p></p></div> 
        <div><b>Sharing Policy - Exchange Online vs Exchange On Premise:<b></div> 
        <div><span style='color: $tdSharpingPolicyCheckColor;'>$($tdSharpingPolicyCheck)</span></div>
        
      </td>
    </tr>
  "


    
    $html | Out-File -FilePath $htmlfile
}

#endregion

#region ExoOauthFuntions

Function EXOIntraOrgConCheck {
    Write-Host -foregroundcolor Green " Get-IntraOrganizationConnector | Select TargetAddressDomains,DiscoveryEndpoint,Enabled"
    Write-Host $bar
    $exoIntraOrgCon = Get-IntraOrganizationConnector | Select-Object TargetAddressDomains, DiscoveryEndpoint, Enabled
    #$IntraOrgConCheck
    $IOC = $exoIntraOrgCon | Format-List
    $IOC
    Write-Host $bar
    Write-Host -foregroundcolor Green " Summary - Online Intra Organization Connector"
    Write-Host $bar
    Write-Host -foregroundcolor White " Target Address Domains: "
    if ($exoIntraOrgCon.TargetAddressDomains -like "*$ExchangeOnpremDomain*") {
        Write-Host -foregroundcolor Green " " $exoIntraOrgCon.TargetAddressDomains
        $tdexoIntraOrgConTargetAddressDomains = $exoIntraOrgCon.TargetAddressDomains
        $tdexoIntraOrgConTargetAddressDomainsColor = "green"
    
    }
    else {
        Write-Host -foregroundcolor Red " Target Address Domains is NOT correct."
        Write-Host -foregroundcolor White " Should contain the $ExchangeOnpremDomain"
        $tdexoIntraOrgConTargetAddressDomains = " $($exoIntraOrgCon.TargetAddressDomains) . Should contain the $ExchangeOnpremDomain"
        $tdexoIntraOrgConTargetAddressDomainsColor = "red"
    
    }
    Write-Host -foregroundcolor White " DiscoveryEndpoint: "
    if ($exoIntraOrgCon.DiscoveryEndpoint -like $EDiscoveryEndpoint.OnPremiseDiscoveryEndpoint) {
        Write-Host -foregroundcolor Green $exoIntraOrgCon.DiscoveryEndpoint
        $tdexoIntraOrgConDiscoveryEndpoints = $exoIntraOrgCon.DiscoveryEndpoint
        $tdexoIntraOrgConDiscoveryEndpointsColor = "green"
    
    
    
    }
    else {
        Write-Host -foregroundcolor Red " DiscoveryEndpoint is NOT correct. "
        Write-Host -foregroundcolor White "  Should be " $EDiscoveryEndpoint.OnPremiseDiscoveryEndpoint
        $tdexoIntraOrgConDiscoveryEndpoints = "$($exoIntraOrgCon.DiscoveryEndpoint) . Should be $($EDiscoveryEndpoint.OnPremiseDiscoveryEndpoint)"
        $tdexoIntraOrgConDiscoveryEndpointsColor = "red"
    
    
    
    }
    Write-Host -foregroundcolor White " Enabled: "
    if ($exoIntraOrgCon.Enabled -like "True") {
        Write-Host -foregroundcolor Green "  True "
        $tdexoIntraOrgConEnabled = "$($exoIntraOrgCon.Enabled)"
        $tdexoIntraOrgConEnabledColor = "green"
    
    
    
    }
    else {
        Write-Host -foregroundcolor Red "  False."
        Write-Host -foregroundcolor White " Should be True"
        $tdexoIntraOrgConEnabled = "$($exoIntraOrgCon.Enabled) . Should be True"
        $tdexoIntraOrgConEnabledColor = "red"
    
    
    }

    
    

    $script:html += "
    <tr>
      <th colspan='2' style='text-align:center; color:white;'><b>Exchange Online OAuth Configuration</b></th>
    </tr>
    <tr>
      <th colspan='2' style=' color:white;'><b>Summary - Get-IntraOrganizationConnector</b></th>
    </tr>
    <tr>
      <td><b>  Get-IntraOrganizationConnector | Select-Object TargetAddressDomains, DiscoveryEndpoint, Enabled</b></td>
      <td>
        <div><b>Target Address Domains:</b><span style='color: $tdexoIntraOrgConTargetAddressDomainsColor;'>' $($tdexoIntraOrgConTargetAddressDomains)'</span></div>
        <div><b>DiscoveryEndpoint:</b><span style='color: $tdexoIntraOrgConDiscoveryEndpointsColor;'>' $($tdexoIntraOrgConDiscoveryEndpoints)'</span></div>
        <div><b>Enabled:</b><span style='color:$tdexoIntraOrgConEnabledColor;'> $($tdexoIntraOrgConEnabled)</span></div>
        
      </td>
    </tr>
  "


    
    $html | Out-File -FilePath $htmlfile



}

Function EXOIntraOrgConfigCheck {
    Write-Host -foregroundcolor Green " Get-IntraOrganizationConfiguration | Select OnPremiseTargetAddresses"
    Write-Host $bar
    #fix because there can be multiple on prem or guid's
    #$exoIntraOrgConfig = Get-OnPremisesOrganization | select OrganizationGuid | Get-IntraOrganizationConfiguration | Select OnPremiseTargetAddresses
    $exoIntraOrgConfig = Get-OnPremisesOrganization | Select-Object OrganizationGuid | Get-IntraOrganizationConfiguration | Select-Object * | Where-Object { $_.OnPremiseTargetAddresses -like "*$ExchangeOnPremDomain*" }
    #$IntraOrgConCheck
    $IOConfig = $exoIntraOrgConfig | Format-List
    $IOConfig
    Write-Host $bar
    Write-Host -foregroundcolor Green " Summary - Online Intra Organization Configuration"
    Write-Host $bar
    Write-Host -foregroundcolor White " OnPremiseTargetAddresses: "
    if ($exoIntraOrgConfig.OnPremiseTargetAddresses -like "*$ExchangeOnpremDomain*") {
        Write-Host -foregroundcolor Green " " $exoIntraOrgConfig.OnPremiseTargetAddresses
        $tdexoIntraOrgConfigOnPremiseTargetAddresses = $exoIntraOrgConfig.OnPremiseTargetAddresses
        $tdexoIntraOrgConfigOnPremiseTargetAddressesColor = "green"
    
    }
    else {
        Write-Host -foregroundcolor Red " OnPremise Target Addressess are NOT correct."
        Write-Host -foregroundcolor White " Should contain the $ExchangeOnpremDomain"
        $tdexoIntraOrgConfigOnPremiseTargetAddresses = $exoIntraOrgConfig.OnPremiseTargetAddresses
        $tdexoIntraOrgConfigOnPremiseTargetAddressesColor = "red"
    }


    $script:html += "
    
    <tr>
      <th colspan='2' style=color:white;'><b>Summary - Get-IntraOrganizationConfiguration</b></th>
    </tr>
    <tr>
      <td><b>  Get-IntraOrganizationConfiguration | Select OnPremiseTargetAddresses</b></td>
      <td>
        <div><b>OnPremiseTargetAddresses:</b><span style='color: $tdexoIntraOrgConfigOnPremiseTargetAddressesColor;'>$($tdexoIntraOrgConfigOnPremiseTargetAddresses)</span></div>
        
      </td>
    </tr>
  "


    
    $html | Out-File -FilePath $htmlfile


}

Function EXOauthservercheck {
    Write-Host -foregroundcolor Green " Get-AuthServer -Identity 00000001-0000-0000-c000-000000000000 | select name,issueridentifier,enabled"
    Write-Host $bar
    $exoauthserver = Get-AuthServer -Identity 00000001-0000-0000-c000-000000000000 | Select-Object name, issueridentifier, enabled
    #$IntraOrgConCheck
    $authserver = $exoauthserver | Format-List
    $authserver
    $tdexoauthserverName = $exoauthserver.Name
    Write-Host $bar
    Write-Host -foregroundcolor Green " Summary - Exchange Online Authorization Server"
    Write-Host $bar
    Write-Host -foregroundcolor White " IssuerIdentifier: "
    if ($exoauthserver.IssuerIdentifier -like "00000001-0000-0000-c000-000000000000") {
        Write-Host -foregroundcolor Green " " $exoauthserver.IssuerIdentifier
        $tdexoauthserverIssuerIdentifier = $exoauthserver.IssuerIdentifier
        $tdexoauthserverIssuerIdentifierColor = "green"
        if ($exoauthserver.Enabled -like "True") {
            Write-Host -foregroundcolor Green "  True "
            $tdexoauthserverEnabled = $exoauthserver.Enabled
            $tdexoauthserverEnabledColor = "green"
        
        }
        else {
            Write-Host -foregroundcolor Red "  Enabled is NOT correct."
            Write-Host -foregroundcolor White " Should be True"
            $tdexoauthserverEnabled = "$($exoauthserver.Enabled) . Should be True"
            $tdexoauthserverEnabledColor = "red"
        
        }
    }
    else {
        Write-Host -foregroundcolor Red " Authorization Server object is NOT correct."
        Write-Host -foregroundcolor White " Enabled: "
        $tdexoauthserverIssuerIdentifier = "$($exoauthserver.IssuerIdentifier) - Authorization Server object should be 00000001-0000-0000-c000-000000000000"
        $tdexoauthserverIssuerIdentifierColor = "red"
        
        if ($exoauthserver.Enabled -like "True") {
            Write-Host -foregroundcolor Green "  True "
            $tdexoauthserverEnabled = $exoauthserver.Enabled
            $tdexoauthserverEnabledColor = "green"
        
        }
        else {
            Write-Host -foregroundcolor Red "  Enabled is NOT correct."
            Write-Host -foregroundcolor White " Should be True"
            $tdexoauthserverEnabled = "$($exoauthserver.Enabled) . Should be True"
            $tdexoauthserverEnabledColor = "red"
        
        }
    }

    $script:html += "
    
    <tr>
      <th colspan='2' style='color:white;'>Summary - Get-AuthServer</th>
    </tr>
    <tr>
      <td><b>  Get-AuthServer -Identity 00000001-0000-0000-c000-000000000000 | select name,issueridentifier,enabled</b></td>
      <td>
        <div><b>Name:</b><span style='color: $tdexoauthserverNameColor;'>$($tdexoauthserverName)</span></div>
        <div><b>IssuerIdentifier:</b><span style='color: $tdexoauthserverIssuerIdentifierColor;'>$($tdexoauthserverIssuerIdentifier)</span></div>
        <div><b>Enabled:</b><span style='color: $tdexoauthserverEnabledColor;'>$($tdexoauthserverEnabled)</span></div>
        
      </td>
    </tr>
  "


    
    $html | Out-File -FilePath $htmlfile

}

Function EXOtestoauthcheck {
    Write-Host -foregroundcolor Green " Test-OAuthConnectivity -Service EWS -TargetUri $Global:ExchangeOnPremEWS -Mailbox $useronline "
    Write-Host $bar
    $exotestoauth = Test-OAuthConnectivity -Service EWS -TargetUri $Global:ExchangeOnPremEWS -Mailbox $useronline 
    if ($exotestoauth.ResultType.Value -like 'Success' ) {
        #$exotestoauth.ResultType.Value
         
        $tdOAuthConnectivityResultType = "$($exotestoauth.ResultType.Value) - OAuth Test was completed successfully"
        $tdOAuthConnectivityResultTypeColor = "green"
    }
    else {
        $exotestoauth | Format-List
        $tdOAuthConnectivityResultType = "$($exotestoauth.ResultType) - OAuth Test was completed with Error. Please rerun Test-OAuthConnectivity -Service EWS -TargetUri <EWS target URI> -Mailbox <On Premises Mailbox> | fl to confirm the test failure"
        $tdOAuthConnectivityResultTypeColor = "red"
    }
    
    if ($exotestoauth.Detail.FullId -like '*(401) Unauthorized*') {
        write-host -ForegroundColor Red "The remote server returned an error: (401) Unauthorized"

    }

    if ($exotestoauth.Detail.FullId -like '*The user specified by the user-context in the token does not exist*') {
        write-host -ForegroundColor Yellow "The user specified by the user-context in the token does not exist"
        write-host "Please run Test-OAuthConnectivity with a different Exchange Online Mailbox"
        
        
        
    }

    if ($exotestoauth.Detail.FullId -like '*error_category="invalid_token"*') {
        write-host -ForegroundColor Yellow "This token profile 'S2SAppActAs' is not applicable for the current protocol"
        
        
        
        
    }
    #Write-Host $bar
    #$OAuthConnectivity.detail.LocalizedString
    Write-Host -foregroundcolor Green " Summary - Test-OAuthConnectivity"
    Write-Host $bar
    if ($exotestoauth.ResultType.value -like "Success") {
        Write-Host -foregroundcolor Green " OAuth Test was completed successfully "
        $tdOAuthConnectivityResultType = "  OAuth Test was completed successfully"
        $tdOAuthConnectivityResultTypeColor = "green"
        
    }
    else {
        Write-Host -foregroundcolor Red " OAuth Test was completed with Error. "
        Write-Host -foregroundcolor White " Please rerun Test-OAuthConnectivity -Service EWS -TargetUri <EWS target URI> -Mailbox <On Premises Mailbox> | fl to confirm the test failure"
        $tdOAuthConnectivityResultType = "$($exotestoauth.ResultType) - OAuth Test was completed with Error. Please rerun Test-OAuthConnectivity -Service EWS -TargetUri <EWS target URI> -Mailbox <On Premises Mailbox> | fl to confirm the test failure"
        $tdOAuthConnectivityResultTypeColor = "red"
        
        
    }
    
    
    #Write-Host -foregroundcolor Yellow "NOTE: You can ignore the warning 'The SMTP address has no mailbox associated with it'"
    #Write-Host -foregroundcolor Yellow " when the Test-OAuthConnectivity returns a Success"
    Write-Host -foregroundcolor Green "`n References: "
    Write-Host -foregroundcolor White " Configure OAuth authentication between Exchange and Exchange Online organizations"
    Write-Host -foregroundcolor Yellow " https://technet.microsoft.com/en-us/library/dn594521(v=exchg.150).aspx"


    $script:html += "
   
    <tr>
      <th colspan='2' style='color:white;'><b>Summary - Test-OAuthConnectivity</b></th>
    </tr>
    <tr>
      <td><b>  Test-OAuthConnectivity -Service EWS -TargetUri $($Global:ExchangeOnPremEWS) -Mailbox $useronline </b></td>
      <td>
        <div><b>Result:</b><span style='color: $tdOAuthConnectivityResultTypeColor;'>$($tdOAuthConnectivityResultType)</span></div>
        
      </td>
    </tr>
  "


    
    $html | Out-File -FilePath $htmlfile

}

#endregion
#cls
$IntraOrgCon = Get-IntraOrganizationConnector -WarningAction SilentlyContinue -ErrorAction SilentlyContinue | Select-Object Name, TargetAddressDomains, DiscoveryEndpoint, Enabled
#if($Auth -contains "DAuth" -and $IntraOrgCon.enabled -Like "True")

ShowParameters

if ($IntraOrgCon.enabled -Like "True") {
    Write-Host $bar
    Write-Host -foregroundcolor yellow "  Warning: Intra Organization Connector Enabled True `n  " 
    Write-Host -foregroundcolor White "    -> Free Busy Lookup is done using OAuth when the Intra Organization Connector is Enabled"
    Write-Host -foregroundcolor White "`n         This script can be Run using the -Auth paramenter to check for OAuth configurations only. `n `n         Example: ./FreeBusyChecker.ps1 -Auth OAuth"
    $Script:html += "<div><p></p></div><div>  <span style='color:orange;'><b>Warning: </b></span> Intra Organization Connector Enabled True `n <b>  -> Free Busy Lookup is done using OAuth when the Intra Organization Connector is Enabled<b></div>
     <div><p></p></div><div>  This script can be Run using the -Auth paramenter to check for OAuth configurations only. `n Example: ./FreeBusyChecker.ps1 -Auth OAuth </div> <div><p></p></div>

    "
    $html | Out-File -FilePath $htmlfile
}

if ($IntraOrgCon.enabled -Like "False") {
    Write-Host $bar
    Write-Host -foregroundcolor yellow "  Warning: Intra Organization Connector Enabled False -> Free Busy Lookup is done using DAuth (Organization Relationship) when the Intra Organization Connector is Disabled"
    Write-Host -foregroundcolor White "`n  This script can be Run using the -Auth paramenter to check for OAuth configurations only. `n  Example: ./FreeBusyChecker.ps1 -Auth OAuth"
    $Script:html += "<p>/p> <div>  <span style='color:yellow;'>Warning: </span> Intra Organization Connector Enabled True `n <b>       -> Free Busy Lookup is done using OAuth when the Intra Organization Connector is Enabled<b></div>
     <div><p></p></div>
     <div>  This script can be Run using the -Auth paramenter to check for OAuth configurations only. Example: ./FreeBusyChecker.ps1 -Auth OAuth</div> <div><p></p></div>
    "
    $html | Out-File -FilePath $htmlfile
}

do {
    #do while not Y or N
    Write-Host $bar
    $ParamOK = Read-Host " Are this values correct? Pess Y for YES and N for NO"
    $ParamOK = $ParamOK.ToUpper()
} while ($ParamOK -ne "Y" -AND $ParamOK -ne "N")
#cls
Write-Host $bar
if ($ParamOK -eq "N") {
    UserOnlineCheck
    ExchangeOnlineDomainCheck
    UseronpremCheck
    ExchangeOnPremDomainCheck
    ExchangeOnPremEWSCheck
    ExchangeOnPremLocalDomainCheck
}
# Free busy Lookup methods
$OrgRel = Get-OrganizationRelationship | Where-Object { ($_.DomainNames -like $ExchangeOnlineDomain) } | Select-Object Enabled, Identity, DomainNames, FreeBusy*, Target*

$EDiscoveryEndpoint = Get-IntraOrganizationConfiguration -WarningAction SilentlyContinue -ErrorAction SilentlyContinue | Select-Object OnPremiseDiscoveryEndpoint
$SPDomainsOnprem = Get-SharingPolicy -WarningAction SilentlyContinue -ErrorAction SilentlyContinue | Format-List Domains
$SPOnprem = Get-SharingPolicy  -WarningAction SilentlyContinue -ErrorAction SilentlyContinue | Select-Object *

if ($Org -contains 'ExchangeOnPremise' -or -not $Org) {
    #region DAutch Checks
    if ($Auth -contains "DAuth" -OR -not $Auth) {
        $StringTest = " Testing DAuth configuration "
        $side = ($ConsoleWidth - $StringTest.Length - 2) / 2
        $sideString = "*"
        for ( $i = 1; $i -lt $side; $i++) {
            $sideString += "*"
        }
        if ($ConsoleWidth % 2) {
            $fullString = "`n`n$sideString$StringTest$sideString**"
        }
        else {
            $fullString = "`n`n$sideString$StringTest$sideString*"
        }
        Write-Host -foregroundcolor Green $fullString 
        Write-Host $bar
        OrgRelCheck
        Write-Host $bar
        if ($pause) {
            $RH = Read- Host " Press Enter when ready to check the Federation Information Details."
            Write-Host $bar
        }
        FedInfoCheck
        if ($pause) {
            $RH = Read-Host " Press Enter when ready to check the Federation Trust configuration details. "
            Write-Host $bar
        }
        FedTrustCheck
        
        if ($pause) {
            $RH = Read-Host " Press Enter when ready to check the On-Prem Autodiscover Virtual Directory configuration details. "
            Write-Host $bar
        }
        AutoDVirtualDCheck
        Write-Host $bar
        if ($pause) {
            $RH = Read-Host " Press Enter when ready to check the On-Prem Web Services Virtual Directory configuration details. "
            Write-Host $bar
        }
        EWSVirtualDirectoryCheck
        if ($pause) {
            Write-Host $bar
            $RH = Read-Host " Press Enter when ready to  check the Availability Address Space configuration details. "
        }
        AvailabilityAddressSpaceCheck
        if ($pause) {
            Write-Host $bar
            $RH = Read-Host " Press Enter when ready to test the Federation Trust. "
        }
        #need to grab errors and provide alerts in error case
        TestFedTrust
        if ($pause) {
            Write-Host $bar
            $RH = Read-Host " Press Enter when ready to Test the Organization Relationship. "
        }
        TestOrgRel
    }
    #endregion
    #region OAuth Check
    if ($Auth -like "OAuth" -or -not $Auth) {
        if ($pause) {
            $RH = Read-Host " Press Enter when ready to check the OAuth configuration details. "
            Write-Host $bar
        }
        $StringTest = " Testing OAuth configuration "
        $side = ($ConsoleWidth - $StringTest.Length - 2) / 2
        $sideString = "*"
        for ( $i = 1; $i -lt $side; $i++) {
            $sideString += "*"
        }
        if ($ConsoleWidth % 2) {
            $fullString = "`n`n$sideString$StringTest$sideString**"
        }
        else {
            $fullString = "`n`n$sideString$StringTest$sideString*"
        }
        Write-Host -foregroundcolor Green $fullString 
        Write-Host $bar
        IntraOrgConCheck
        Write-Host $bar
        if ($pause) {
            $RH = Read-Host " Press Enter when ready to check the Auth Server configuration details. "
            Write-Host $bar
        }
        AuthServerCheck
        Write-Host $bar
        if ($pause) {
            $RH = Read-Host " Press Enter when ready to check the Partner Application configuration details. "
        }
        PartnerApplicationCheck
        Write-Host $bar
        if ($pause) {
            $RH = Read-Host " Press Enter when ready to check the Exchange Online-ApplicationAccount configuration details. "
        }
        ApplicationAccounCheck
        Write-Host $bar
        if ($pause) {
            $RH = Read-Host " Press Enter when ready to check the Management Role Assignments for the Exchange Online-ApplicationAccount. "
            Write-Host $bar
        }
        ManagementRoleAssignmentCheck
        Write-Host $bar
        if ($pause) {
            $RH = Read-Host " Press Enter when ready to check Auth configuration details. "
            Write-Host $bar
        }
        AuthConfigCheck
        Write-Host $bar
        if ($pause) {
            $RH = Read-Host " Press Enter when ready to check the Auth Certificate configuration details. "
            Write-Host $bar
        }
        CurrentCertificateThumbprintCheck
        Write-Host $bar
        if ($pause) {
            $RH = Read-Host " Press Enter when ready to  check the On Prem Autodiscover Virtual Directory configuration details. "
            Write-Host $bar
        }
        AutoDVirtualDCheckOAuth
        $AutoDiscoveryVirtualDirectoryOAuth
        Write-Host $bar
        if ($pause) {
            $RH = Read-Host " Press Enter when ready to check the On-Prem Web Services Virtual Directory configuration details. "
            Write-Host $bar
        }
        EWSVirtualDirectoryCheckOAuth
        Write-Host $bar
        if ($pause) {
            $RH = Read-Host " Press Enter when ready to check the AvailabilityAddressSpace configuration details. "
            Write-Host $bar
        }
        AvailabilityAddressSpaceCheckOAuth
        Write-Host $bar
        if ($pause -eq "True") {
            $RH = Read-Host " Press Enter when ready to test the OAuthConnectivity configuration details. "
            Write-Host $bar
        }
        OAuthConnectivityCheck
        Write-Host $bar
    }
    #$bar
    #endregion
}
# EXO Part
if ($Org -contains 'ExchangeOnline' -OR -not $Org) {
    #region ConnectExo
    #$bar
    Write-Host -ForegroundColor Green " Collecting Exchange Online Availability Information"
    #$bar
    #Exchange Online Management Shell
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    install-module AzureAD -AllowClobber
    #$CreateEXOPSSession = (Get-ChildItem -Path $env:userprofile -Filter CreateExoPSSession.ps1 -Recurse -ErrorAction SilentlyContinue -Force | Select -Last 1).DirectoryName. "$CreateEXOPSSession\CreateExoPSSession.ps1"
    #Import-Module $((Get-ChildItem -Path $($env:LOCALAPPDATA+"\Apps\2.0\") -Filter CreateExoPSSession.ps1 -Recurse ).FullName | Select-Object -Last 1)
    #Connect-EXOPSSession
    #Connect-EXOPSSession
    #RestV3 connection
    Install-Module -Name ExchangeOnlineManagement
    Connect-ExchangeOnline -ShowBanner:$false
    #Write-Host "========================================================="
    #Write-Host "Get-SharingPolicy | FL"
    #Write-Host "========================================================="
    #Get-SharingPolicy | FL
    # Variables
    $Script:ExoOrgRel = Get-OrganizationRelationship | Where-Object { ($_.DomainNames -like $ExchangeOnPremDomain ) } | Select-Object Enabled, Identity, DomainNames, FreeBusy*, Target*
    $ExoIntraOrgCon = Get-IntraOrganizationConnector | Select-Object Name, TargetAddressDomains, DiscoveryEndpoint, Enabled
    $targetadepr1 = ("https://autodiscover." + $ExchangeOnPremDomain + "/autodiscover/autodiscover.svc/WSSecurity")
    $targetadepr2 = ("https://" + $ExchangeOnPremDomain + "/autodiscover/autodiscover.svc/WSSecurity")
    $exofedinfo = get-federationInformation -DomainName $exchangeOnpremDomain  -BypassAdditionalDomainValidation -ErrorAction SilentlyContinue | Select-Object *
    
    $Script:html += "
     </table>

      <div class='Black'><p></p></div>

      <div class='Black'><p></p></div>

             <div class='Black'><h2><b>`n Exchange Online Free Busy Configuration: `n</b></h2></div>

             <div class='Black'><p></p></div>
             <div class='Black'><p></p></div>

     <table style='width:100%; margin-top:30px;'>

    "
    
    #endregion
    #region ExoDauthCheck
    if ($Auth -contains "DAuth" -or -not $Auth) {
        Write-Host $bar
        $StringTest = " Testing DAuth configuration "
        $side = ($ConsoleWidth - $StringTest.Length - 2) / 2
        $sideString = "*"
        for ( $i = 1; $i -lt $side; $i++) {
            $sideString += "*"
        }
        if ($ConsoleWidth % 2) {
            $fullString = "`n`n$sideString$StringTest$sideString**"
        }
        else {
            $fullString = "`n`n$sideString$StringTest$sideString*"
        }
        Write-Host -foregroundcolor Green $fullString 
        Write-Host $bar
        ExoOrgRelCheck
        Write-Host $bar
        if ($pause) {
            $RH = Read-Host " Press Enter when ready to check the Federation Organization Identifier configuration details. "
            Write-Host $bar
        }
        EXOFedOrgIdCheck
        Write-Host $bar
        if ($pause) {
            $RH = Read-Host " Press Enter when ready to check the Organization Relationship configuration details. "
            Write-Host $bar
        }
        EXOTestOrgRelCheck
        if ($pause) {
            Write-Host $bar
            $RH = Read-Host " Press Enter when ready to check the Sharing Policy configuration details. "
        }
        SharingPolicyCheck
    }

    #endregion
    #region ExoOauthCheck
    if ($Auth -contains "OAuth" -or -not $Auth) {
        $StringTest = " Testing OAuth configuration "
        $side = ($ConsoleWidth - $StringTest.Length - 2 ) / 2
        $sideString = "*"
        for ( $i = 1; $i -lt $side; $i++) {
            $sideString += "*"
        }
        if ($ConsoleWidth % 2) {
            $fullString = "`n`n$sideString$StringTest$sideString**"
        }
        else {
            $fullString = "`n`n$sideString$StringTest$sideString*"
        }
        Write-Host -foregroundcolor Green $fullString 
        Write-Host $bar
        ExoIntraOrgConCheck
        Write-Host $bar
        if ($pause) {
            $RH = Read-Host " Press Enter when ready to check the Organizationconfiguration details. "
            Write-Host $bar
        }
        EXOIntraOrgConfigCheck
        Write-Host $bar
        if ($pause) {
            $RH = Read-Host " Press Enter when ready to check the Authentication Server Authorization Details.  "
            Write-Host $bar
        }
        EXOauthservercheck
        Write-Host $bar
        if ($pause) {
            $RH = Read-Host " Press Enter when ready to test the OAuth Connectivity Details.  "
            Write-Host $bar
        }
        EXOtestoauthcheck
        Write-Host $bar
    }
    
    #endregion
    disConnect-ExchangeOnline  -Confirm:$False
    Write-Host -foregroundcolor Green " That is all for the Exchange Online Side"
    #Read-Host "Ctrl+C to exit. Enter to Exit."
    $bar
}
stop-transcript
#Read-Host " `n `n Ctrl+C to exit. Enter to Exit."
