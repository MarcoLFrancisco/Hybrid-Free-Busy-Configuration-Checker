#Exchange on Premise
#>
#region Properties and Parameters
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', '', Justification = 'Variables are being used')]
[CmdletBinding(DefaultParameterSetName = "FreeBusyInfo_OP", SupportsShouldProcess)]

param(
[Parameter(Mandatory = $false, ParameterSetName = "Auth")]
[string]$Auth,
[string]$Pause,
[string]$Org,
[string]$Help
)

Function ShowHelp {
    $bar
    Write-host -ForegroundColor Yellow "`n  Valid Input Option Parameters!"
    Write-Host -ForegroundColor White "`n  Paramater: Auth"
    Write-Host -ForegroundColor White "   Options  : DAuth; OAUth; Null"
    Write-Host  "    DAuth        : DAuth Authentication"
    Write-Host  "    OAuth        : OAuth Authentication"
    Write-Host  "    Default Value: Null. No swith input means the script will collect both DAuth and OAuth Availability Configuration Detail"
    Write-Host -ForegroundColor White "`n  Paramater: Org"
    Write-Host -ForegroundColor White "   Options  : EOP; EOL; Null"
    Write-Host  "    EOP          : Use EOP parameter to collect Availability information in the Exchange On Premise Tenant"
    Write-Host  "    EOL          : Use EOL parameter to collect Availability information in the Exchange Online Tenant"
    Write-Host  "    Default Value: Null. No swith input means the script will collect both Exchange On Premise and Exchange OnlineAvailability configuration Detail"
    Write-Host -ForegroundColor White "`n  Paramater: Pause"
    Write-Host -ForegroundColor White "   Options  : Null; True; False"
    Write-Host  "    True         : Use the True parameter to use this script pausing after each test done."
    Write-Host  "    False        : To use this script not pausing after each test done no Pause Parameter is needed."
    Write-Host  "    Default Value: False. `n"
    Write-Host -ForegroundColor White "`n  Paramater: Help"
    Write-Host -ForegroundColor White "   Options  : Null; True; False"
    Write-Host  "    True         : Use the True parameter to use display valid parameter Options. `n`n"
}

If ($Help -like "True") {
    Write-Host $bar
    ShowHelp;
    $bar
    exit
}

if (![string]::IsNullOrWhitespace($Auth)) {
    if ($Auth -notlike "DAuth" -And $Auth -notlike "OAuth") {
        Write-host -ForegroundColor Red "`n  Invalid Input Option Parameters!"
        ShowHelp
        exit
    }
}

if (![string]::IsNullOrWhitespace($Org)) {
    if ($Org -notlike "EOL" -AND $Org -notlike "EOP") {
        Write-host -ForegroundColor Red "`n  Invalid Input Option Parameters!"
        ShowHelp
        exit
    }
}

#Set-ExecutionPolicy " Unrestricted"  -Scope Process -Confirm:$false
#Set-ExecutionPolicy " Unrestricted"  -Scope CurrentUser -Confirm:$false
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
$bar = " =================================================================================================================="
$logfile = "$PSScriptRoot\FreeBusyInfo_OP.txt"
$Logfile = [System.IO.Path]::GetFileNameWithoutExtension($logfile) + "_" + `
(get-date -format yyyyMMdd_HHmmss) + ([System.IO.Path]::GetExtension($logfile))
Write-Host " `n`n "
Start-Transcript -path $LogFile -append
Write-Host $bar
Write-Host -foregroundcolor Green " `n  Free Busy Configuration Information Grabber `n "
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
$UserOnPrem = get-mailbox -resultsize 1 -WarningAction SilentlyContinue -Filter 'EmailAddresses -like $temp -and HiddenFromAddressListsEnabled -eq $false'
$UserOnPrem = $UserOnPrem.PrimarySmtpAddress.Address
$ExchangeOnPremDomain = ($UserOnPrem -split "@")[1]
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
        Write-Host -ForegroundColor Green "  Green - In SUMMARY Sections it means OK. Anywhere else it's just a visual aid."
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
	        <title>Show Parameters Output</title>
	        <style>
                table, th, td {
              border: 1px solid black;
              border-collapse: collapse;
              padding: 5px;
            }
            th {
              background-color: lightgray;
              text-align: left;
            }
		        .green { color: green; }
		        .red { color: red; }
		        .yellow { color: yellow; }
		        .white { color: white; }
                .black { color: white; }
	        </style>
        </head>
        <body>
	        <h1>Show Parameters Output</h1>
	        <hr />
	        <p class='green'>Loading modules for AD, Exchange</p>
	        <hr />
	        <h2>Color Scheme</h2>
	        <hr />
	        <p class='red'>Look out for Red!</p>
	        <p class='yellow'>Yellow - Example information or Links</p>
	        <p class='green'>Green - In SUMMARY Sections it means OK. Anywhere else it's just a visual aid.</p>
	        <hr />
	        <h2>Parameters:</h2>
	        <hr />
	        <p class='Black'>Log File Path:</p>
	        <p class='green'>$PSScriptRoot\$Logfile</p>
	        <p class='Black'>Office 365 Domain:</p>
            <p class='green'>$exchangeOnlineDomain</p>
	        <p class='Black'>AD root Domain</p>
	        <p class='green'>$exchangeOnPremLocalDomain</p>
	        <p class='Black'>Exchange On Premises Domain:</p>
	        <p class='green'>$exchangeOnPremDomain</p>
	        <p class='Black'>Exchange On Premises External EWS url:</p>
	        <p class='green'>$exchangeOnPremEWS</p>
	        <p class='Black'>On Premises Hybrid Mailbox:</p>
	        <p class='green'>$useronprem</p>
	        <p class='Black'>Exchange Online Mailbox:</p>
	        <p class='green'>$userOnline</p>"

        $html| Out-File -FilePath "$PSScriptRoot\ShowParameters1.html"

    }
#}
#endregion

#regionDAuth Functions

Function OrgRelCheck {
    Write-Host $bar
    Write-Host -foregroundcolor Green " Get-OrganizationRelationship  | Where{($_.DomainNames -like $ExchangeOnlineDomain )} | Select Identity,DomainNames,FreeBusy*,Target*,Enabled, ArchiveAccessEnabled"
    Write-Host $bar
    $OrgRel
    Write-Host $bar
    Write-Host  -foregroundcolor Green " SUMMARY - Organization Relationship"
    Write-Host $bar
    #$exchangeonlinedomain
    Write-Host  -foregroundcolor White   " Domain Names:"
    if ($orgrel.DomainNames -like $exchangeonlinedomain){
        Write-Host -foregroundcolor Green "  Domain Names Include the $exchangeOnlineDomain Domain"
        $tdDomainNames = "Domain Names Include the $exchangeOnlineDomain Domain"
        $tdDomainNamesColor = "green"
        $tdDomainNamesfl = $tdDomainNames | fl
    }
    else{
        Write-Host -foregroundcolor Red "  Domain Names do Not Include the $exchangeOnlineDomain Domain"
        $tdDomainNames = "Domain Names do Not Include the $exchangeOnlineDomain Domain"
        $tdDomainNamesColor = "Red"
    }
    #FreeBusyAccessEnabled
    Write-Host -foregroundcolor White   " FreeBusyAccessEnabled:"
    if ($OrgRel.FreeBusyAccessEnabled -like "True" ){
        Write-Host -foregroundcolor Green "  FreeBusyAccessEnabled is set to True"
        $tdFBAccessEnabled = "FreeBusyAccessEnabled is set to True"
        $tdFBAccessEnabledColor = "green"
    }
    else{
        Write-Host -foregroundcolor Red "  FreeBusyAccessEnabled : False"
        $tdFBAccessEnabled = "FreeBusyAccessEnabled is set to False"
        $tdFBAccessEnabledColor = "red"
        $countOrgRelIssues++
    }
    #FreeBusyAccessLevel
    Write-Host -foregroundcolor White   " FreeBusyAccessLevel:"
    if ($OrgRel.FreeBusyAccessLevel -like "AvailabilityOnly" ){
        Write-Host -foregroundcolor Green "  FreeBusyAccessLevel is set to AvailabilityOnly"
        $tdFBAccessLevel = "FreeBusyAccessLevel is set to AvailabilityOnly"
        $tdFBAccessLevelColor = "green"
    }
    if ($OrgRel.FreeBusyAccessLevel -like "LimitedDetails" ){
        Write-Host -foregroundcolor Green "  FreeBusyAccessLevel is set to LimitedDetails"
        $tdFBAccessLevel = "FreeBusyAccessLevel is set to  LimitedDetails"
        $tdFBAccessLevelColor = "green"
    }
    if ($OrgRel.FreeBusyAccessLevel -ne "LimitedDetails" -AND $OrgRel.FreeBusyAccessLevel -ne "AvailabilityOnly" )
    {
        Write-Host -foregroundcolor Red "  FreeBusyAccessEnabled : False"
        $tdFBAccessLevel = "FreeBusyAccessEnabled : False"
        $tdFBAccessLevelColor = "Red"
        $countOrgRelIssues++
    }
    #TargetApplicationUri
    Write-Host -foregroundcolor White   " TargetApplicationUri:"
    if ($OrgRel.TargetApplicationUri -like "Outlook.com" ){
        Write-Host -foregroundcolor Green "  TargetApplicationUri is Outlook.com"
        $tdTargetApplicationUri = "TargetApplicationUri is Outlook.com"
        $tdTargetApplicationUriColor = "green"
    }
    else {
        Write-Host -foregroundcolor Red "  TargetApplicationUri should be Outlook.com"
        $tdTargetApplicationUri = "TargetApplicationUri is Outlook.com"
        $tdTargetApplicationUriColor = "red"
        $countOrgRelIssues++
    }
    #TargetOwaURL
    Write-Host -foregroundcolor White   " TargetOwaURL:"
    if ($OrgRel.TargetOwaURL -like "http://outlook.com/owa/$exchangeonlinedomain" -or $OrgRel.TargetOwaURL -like $Null){
        if ($OrgRel.TargetOwaURL -like "http://outlook.com/owa/$exchangeonlinedomain"){
            Write-Host -foregroundcolor Green "  TargetOwaURL is http://outlook.com/owa/$exchangeonlinedomain. This is a possible standard value. TargetOwaURL can also be configured to be Blank."
        }
        if ($OrgRel.TargetOwaURL -like $Null){
            Write-Host -foregroundcolor Green "  TargetOwaURL is Blank, this is a standard value. "
            Write-Host  "  TargetOwaURL can also be configured to be http://outlook.com/owa/$exchangeonlinedomain"
        }
    }
    else{
        Write-Host -foregroundcolor Red "  TargetOwaURL seems not to be Blank or http://outlook.com/owa/$exchangeonlinedomain. These are the standard values."
        $countOrgRelIssues++
    }
    #TargetSharingEpr
    Write-Host -foregroundcolor White   " TargetSharingEpr:"
    if ([string]::IsNullOrWhitespace($OrgRel.TargetSharingEpr) -or $OrgRel.TargetSharingEpr -eq "https://outlook.office365.com/EWS/Exchange.asmx "){
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
    if ([string]::IsNullOrWhitespace($OrgRel.FreeBusyAccessScope)){
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
    Write-Host -foregroundcolor White   " TargetAutodiscoverEpr:"
    if ($OrgRel.TargetAutodiscoverEpr -like "https://autodiscover-s.outlook.com/autodiscover/autodiscover.svc/WSSecurity" ){
        Write-Host -foregroundcolor Green "  TargetAutodiscoverEpr is correct"
    }
    else
    {
        Write-Host -foregroundcolor Red "  TargetAutodiscoverEpr is not correct"
        $countOrgRelIssues++
    }
    #Enabled
    Write-Host -foregroundcolor White   " Enabled:"
    if ($OrgRel.enabled -like "True" ){
        Write-Host -foregroundcolor Green "  Enabled is set to True"
    }
    else
    {
        Write-Host -foregroundcolor Red "  Enabled is set to False."
        $countOrgRelIssues++
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
    $FreeBusyAccessLevel= $OrgRel.FreeBusyAccessLevel
    $tdTargetOwaUrl = $OrgRel.TargetOwaURL
    $tdEnabled = $OrgRel.Enabled
    $script:html += "
    <table style='width:100%'>
    <tr>
    <th colspan='2'>SUMMARY - Organization Relationship</th>
    </tr>
    <tr>
    <td><b>Get-OrganizationRelationship:</b></td>
    <td>
    <p> <b>Domain Names: </b> $OrgRelDomainNames</p>
    <p> <b>FreeBusyAccessEnabled: </b> $FreeBusyAccessEnabled</p>
    <p> <b>FreeBusyAccessLevel: </b> $FreeBusyAccessLevel</p>
    <p> <b>TargetApplicationUri: </b> $tdTargetApplicationUri</p>
    <p> <b>TargetOwaURL: </b> $tdTargetOwaUrl</p>
    <p> <b>TargetSharingEpr: </b> $tdTargetSharingEpr</p>
    <p> <b>FreeBusyAccessScope: </b> $tdFreeBusyAccessScope</p>
    <p> <b>Enabled:  $tdEnabled </b> </p>
    </td>
    </tr>
    <tr>
    <td>Domain Names:</td>
    <td class='$tdDomainNamesColor'>
    $tdDomainNames
    </td>
    </tr>
    <tr>
    <td>FreeBusyAccessEnabled:</td>
    <td class='$tdFBAccessEnabledColor'>
    $tdFBAccessEnabled
    </td>
    </tr>
    <tr>
    <td>FreeBusyAccessLevel:</td>
    <td class='$tdFBAccessLevelColor'>
    $tdFBAccessLevel
    </td>
    </td>
    </tr>
    <tr>
    <td>TargetApplicationUri:</td>
    <td class='$tdTargetApplicationUriColor'>
    $tdTargetApplicationUri
    </td>
    </tr>
    <tr>
    <td>TargetApplicationUri:</td>
    <td class='$(if ($OrgRel.TargetApplicationUri -like 'Outlook.com') { 'green' } else { 'red' })'>
    $(if ($OrgRel.TargetApplicationUri -like 'Outlook.com') {
        'TargetApplicationUri is Outlook.com'
        } else {
        'TargetApplicationUri should be Outlook.com'
    })
    </td>
    </tr>
    <tr>
    <td>TargetOwaURL:</td>
    <td class='$(if ($OrgRel.TargetOwaURL -like 'http://outlook.com/owa/'+$exchangeonlinedomain -or $OrgRel.TargetOwaURL -like $Null) { 'green' } else { 'red' })'>
    $(if ($OrgRel.TargetOwaURL -like 'http://outlook.com/owa/$exchangeonlinedomain') {
        'TargetOwaURL is http://outlook.com/owa/'+$exchangeonlinedomain +'. This is a possible standard value. TargetOwaURL can also be configured to be Blank.'
        } elseif ($OrgRel.TargetOwaURL -like $Null) {
        'TargetOwaURL is Blank, this is a standard value.'
        'TargetOwaURL can also be configured to be http://outlook.com/owa/'+$exchangeonlinedomain+''
        } else {
        'TargetOwaURL seems not to be Blank or http://outlook.com/owa/$exchangeonlinedomain. These are the standard values.'
    })
    </td>
    </tr>
    <tr>
    <td>TargetSharingEpr:</td>
    <td class='$tdTargetSharingEprColor'>
    $tdTargetSharingEpr
    </td>
    </tr>
    <tr>
    <td>FreeBusyAccessScope:</td>
    <td class='$tdFreeBusyAccessScopeColor'>
    $tdFreeBusyAccessScope
    </td>
    </tr>
    <tr>
    <td>TargetAutodiscoverEpr:</td>
    <td class='$(if ($OrgRel.TargetAutodiscoverEpr -like 'https://autodiscover-s.outlook.com/autodiscover/autodiscover.svc/WSSecurity') { 'green' } else { 'red' })'>
    $(if ($OrgRel.TargetAutodiscoverEpr -like 'https://autodiscover-s.outlook.com/autodiscover/autodiscover.svc/WSSecurity') {
        'TargetAutodiscoverEpr is correct'
        } else {
        'TargetAutodiscoverEpr is not correct'
    })
    </td>
    </tr>
    <tr>
    <td>Enabled:</td>
    <td class='$(if ($OrgRel.enabled -like 'True') { 'green' } else { 'red' })'>
    $(if ($OrgRel.enabled -like 'True') {
        'Enabled is set to True'
        } else {
        'Enabled is set to False.'
    })
    </td>
    </tr>"
    $html| Out-File -FilePath "$PSScriptRoot\ShowParameters.html"
    Write-Host -foregroundcolor Yellow "`n  Reference: https://learn.microsoft.com/en-us/exchange/create-an-organization-relationship-exchange-2013-help"
}

Function FedInfoCheck{
    Write-Host -foregroundcolor Green " Get-FederationInformation -DomainName $exchangeOnlineDomain  -BypassAdditionalDomainValidation | fl"
    Write-Host $bar
    $fedinfo = get-federationInformation -DomainName $exchangeOnlineDomain  -BypassAdditionalDomainValidation -ErrorAction SilentlyContinue| select *
    if (!$fedinfo){
        $fedinfo = get-federationInformation -DomainName $exchangeOnlineDomain  -BypassAdditionalDomainValidation -ErrorAction SilentlyContinue| select *
    }
    $fedinfo
    Write-Host $bar
    Write-Host -foregroundcolor Green " SUMMARY - Federation Information"
    Write-Host $bar
    #DomainNames
    Write-Host -foregroundcolor White   "  Domain Names: "
    if ($fedinfo.DomainNames -like "*$ExchangeOnlineDomain*"){
        Write-Host -foregroundcolor Green "   Domain Names include the Exchange Online Domain "$ExchangeOnlineDomain
        $tdDomainNamesColor = "green"
        $tdDomainNamesfl = "Domain Names include the Exchange Online Domain $ExchangeOnlineDomain"
    }
    else{
        Write-Host -foregroundcolor Red "   Domain Names seem not to include the Exchange Online Domain "$ExchangeOnlineDomain
        Write-Host  "   Domain Names: "$fedinfo.DomainNames
        $tdDomainNamesColor = "Red"
        $tdDomainNamesfl = "Domain Names seem not to include the Exchange Online Domain: $ExchangeOnlineDomain"
    }
    #TokenIssuerUris
    Write-Host  -foregroundcolor White  "  TokenIssuerUris: "
    if ($fedinfo.TokenIssuerUris -like "*urn:federation:MicrosoftOnline*"){
        Write-Host -foregroundcolor Green "  "  $fedinfo.TokenIssuerUris
        $tdTokenIssuerUrisColor = "green"
        $tdTokenIssuerUrisFL = $fedinfo.TokenIssuerUris
    }
    else{
        Write-Host "   " $fedinfo.TokenIssuerUris
        Write-Host  -foregroundcolor Red "   TokenIssuerUris should be urn:federation:MicrosoftOnline"
        $tdTokenIssuerUrisColor = "red"
        $tdTokenIssuerUrisFL = "   TokenIssuerUris should be urn:federation:MicrosoftOnline"
    }
    #TargetApplicationUri
    Write-Host -foregroundcolor White   "  TargetApplicationUri:"
    if ($fedinfo.TargetApplicationUri -like "Outlook.com"){
        Write-Host -foregroundcolor Green "  "$fedinfo.TargetApplicationUri
        $tdTargetApplicationUriColor = "green"
        $tdTargetApplicationUriFL = $fedinfo.TargetApplicationUri
    }
    else{
        Write-Host -foregroundcolor Red "   "$fedinfo.TargetApplicationUri
        Write-Host -foregroundcolor Red   "   TargetApplicationUri should be Outlook.com"
        $tdTargetApplicationUriColor = "red"
        $tdTargetApplicationUriFL =  "   TargetApplicationUri should be Outlook.com"
    }
    #TargetAutodiscoverEpr
    Write-Host -foregroundcolor White   "  TargetAutodiscoverEpr:"
    if ($OrgRel.TargetAutodiscoverEpr -like "https://autodiscover-s.outlook.com/autodiscover/autodiscover.svc/WSSecurity"){
        Write-Host -foregroundcolor Green "   "$fedinfo.TargetAutodiscoverEpr
        $tdTargetAutodiscoverEprColor = "green"
        $tdTargetAutodiscoverEprFL =    $fedinfo.TargetAutodiscoverEpr
    }
    else{
        Write-Host -foregroundcolor Red "   "$fedinfo.TargetAutodiscoverEpr
        Write-Host -foregroundcolor Red   "   TargetApplicationUri should be https://autodiscover-s.outlook.com/autodiscover/autodiscover.svc/WSSecurity"
        $tdTargetApplicationUriColor = "red"
        $tdTargetApplicationUriFL =  "TargetApplicationUri should be https://autodiscover-s.outlook.com/autodiscover/autodiscover.svc/WSSecurity"
    }
    # Federation Information TargetApplicationUri vs Organization Relationship TargetApplicationUri
    Write-Host -ForegroundColor White "  Federation Information TargetApplicationUri vs Organization Relationship TargetApplicationUri "
    if ($fedinfo.TargetApplicationUri -like "Outlook.com"){
        if ($OrgRel.TargetApplicationUri -like $fedinfo.TargetApplicationUri){
            Write-Host -foregroundcolor Green "   => Federation Information TargetApplicationUri matches the Organization Relationship TargetApplicationUri "
            Write-Host  "       Organization Relationship TargetApplicationUri:"  $OrgRel.TargetApplicationUri
            Write-Host  "       Federation Information TargetApplicationUri:   "  $fedinfo.TargetApplicationUri
            $tdFederationInformationTAColor = "green"
            $tdFederationInformationTAFL = " => Federation Information TargetApplicationUri matches the Organization Relationship TargetApplicationUri"
        }
        else{
            Write-Host -foregroundcolor Red "   => Federation Information TargetApplicationUri should be Outlook.com and match the Organization Relationship TargetApplicationUri "
            Write-Host  "       Organization Relationship TargetApplicationUri:"  $OrgRel.TargetApplicationUri
            Write-Host  "       Federation Information TargetApplicationUri:   "  $fedinfo.TargetApplicationUri
            $tdFederationInformationTAColor = "red"
            $tdFederationInformationTAFL = " => Federation Information TargetApplicationUri should be Outlook.com and match the Organization Relationship TargetApplicationUri"
        }
    }
    #TargetAutodiscoverEpr vs Organization Relationship TargetAutodiscoverEpr
    Write-Host -ForegroundColor White  "  Federation Information TargetAutodiscoverEpr vs Organization Relationship TargetAutodiscoverEpr "
    if ($OrgRel.TargetAutodiscoverEpr -like $fedinfo.TargetAutodiscoverEpr){
        Write-Host -foregroundcolor Green "   => Federation Information TargetAutodiscoverEpr matches the Organization Relationship TargetAutodiscoverEpr "
        Write-Host  "       Organization Relationship TargetAutodiscoverEpr:"  $OrgRel.TargetAutodiscoverEpr
        Write-Host  "       Federation Information TargetAutodiscoverEpr:   "  $fedinfo.TargetAutodiscoverEpr
        $tdTargetAutodiscoverEprVSColor = "green"
        $tdTargetAutodiscoverEprVSFL =    "=> Federation Information TargetAutodiscoverEpr matches the Organization Relationship TargetAutodiscoverEpr"
    }
    else
    {
        Write-Host -foregroundcolor Red "   => Federation Information TargetAutodiscoverEpr should match the Organization Relationship TargetAutodiscoverEpr"
        Write-Host  "       Organization Relationship TargetAutodiscoverEpr:"  $OrgRel.TargetAutodiscoverEpr
        Write-Host  "       Federation Information TargetAutodiscoverEpr:   "  $fedinfo.TargetAutodiscoverEpr
        $tdTargetAutodiscoverEprVSColor = "red"
        $tdTargetAutodiscoverEprVSFL =   "=> Federation Information TargetAutodiscoverEpr should match the Organization Relationship TargetAutodiscoverEpr"
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
    $Script:html += "< <tr>
    <th colspan='2'>SUMMARY - Federation Information</th>
    </tr>
    <tr>
    <td><b>Get-FederationInformation -Domain $ExchangeOnPremDomain</b></td>
    <td>
    <div><b>TargetApplicationUri:</b> $fedinfoTargetApplicationUri</div>
    <div><b>DomainNames:</b> $FedInfoDomainNames $aux </div>
    <div><b>TargetAutodiscoverEpr:</b> $fedinfoTargetAutodiscoverEpr </div>
    <div><b>TokenIssuerUris:</b> $fedinfoTokenIssuerUris </div>
    </td>
    </tr>
    <tr>
    <td>Domain Names:</td>
    <td class='$tdDomainNamesColor'>
    $tdDomainNamesFL
    </td>
    </tr>
    <tr>
    <td>TokenIssuerUris:</td>
    <td class='$tdTokenIssuerUrisColor'>
    $tdTokenIssuerUrisFL
    </td>
    </tr>
    <tr>
    <td>TargetApplicationUri:</td>
    </td>
    <td class='$tdTargetApplicationUriColor'>
    $tdTargetApplicationUriFL
    </td>
    </tr>
    <tr>
    <td>TargetAutodiscoverEpr:</td>
    <td class='$tdTargetAutodiscoverEprColor'>
    $tdTargetAutodiscoverEprFL
    </td>
    </td>
    </tr>
    <tr>
    <td>Federation Information TargetApplicationUri vs Organization Relationship TargetApplicationUri :</td>
    <td class='$tdFederationInformationTAColor'>
    $tdTokenIssuerUrisFL
    $tdFederationInformationTAFL
    </td>
    </td>
    <tr>
    <td>TargetAutodiscoverEpr vs Organization Relationship TargetAutodiscoverEpr :</td>
    <td class='$tdTargetAutodiscoverEprVSColor'>
    <div>$tdTargetAutodiscoverEpr</div>
    <div>$tdTargetAutodiscoverEprVSFL</div>
    </td>
    </td>
    </tr>
    "
    $html| Out-File -FilePath "$PSScriptRoot\ShowParameters.html"
}

Function FedTrustCheck{
    Write-Host -foregroundcolor Green " Get-FederationTrust | fl ApplicationUri,TokenIssuerUri,OrgCertificate,TokenIssuerCertificate,
    TokenIssuerPrevCertificate, TokenIssuerMetadataEpr,TokenIssuerEpr"
    Write-Host $bar
    $Global:fedtrust = Get-FederationTrust | select ApplicationUri,TokenIssuerUri,OrgCertificate,TokenIssuerCertificate,TokenIssuerPrevCertificate, TokenIssuerMetadataEpr,TokenIssuerEpr
    $fedtrust
    Write-Host $bar
    Write-Host -foregroundcolor Green " SUMMARY - Federation Trust"
    Write-Host $bar
    $CurrentTime = get-date
    Write-Host -foregroundcolor White " Federation Trust Aplication Uri:"
    if ($fedtrust.ApplicationUri -like  "FYDIBOHF25SPDLT.$ExchangeOnpremDomain"){
        Write-Host -foregroundcolor Green " " $fedtrust.ApplicationUri
        $tdfedtrustApplicationUriColor = "green"
        $tdfedtrustApplicationUriFL =    $fedtrust.ApplicationUri
    }
    else
    {
        Write-Host -foregroundcolor Red "  Federation Trust Aplication Uri Should be "$fedtrust.ApplicationUri
        $tdfedtrustApplicationUriColor = "red"
        $tdfedtrustApplicationUriFL =    "  Federation Trust Aplication Uri Should be $fedtrust.ApplicationUri"
    }
    #$fedtrust.TokenIssuerUri.AbsoluteUri
    Write-Host -foregroundcolor White " TokenIssuerUri:"
    if ($fedtrust.TokenIssuerUri.AbsoluteUri -like  "urn:federation:MicrosoftOnline"){
        #Write-Host -foregroundcolor White "  TokenIssuerUri:"
        Write-Host -foregroundcolor Green " "$fedtrust.TokenIssuerUri.AbsoluteUri
        $tdfedtrustTokenIssuerUriColor = "green"
        $tdfedtrustTokenIssuerUriFL =    $fedtrust.TokenIssuerUri.AbsoluteUri
    }
    else
    {
        Write-Host -foregroundcolor Red " Federation Trust TokenIssuerUri should be urn:federation:MicrosoftOnline"
        $tdfedtrustTokenIssuerUriColor = "red"
        $tdfedtrustTokenIssuerFL =    " Federation Trust TokenIssuerUri is currently $fedtrust.TokenIssuerUri.AbsoluteUri but should be urn:federation:MicrosoftOnline"
    }
    Write-Host -foregroundcolor White " Federation Trust Certificate Expiracy:"
    if ($fedtrust.OrgCertificate.NotAfter.Date -gt $CurrentTime){
        Write-Host -foregroundcolor Green "  Not Expired"
        Write-Host  "   - Expires on " $fedtrust.OrgCertificate.NotAfter.DateTime
        $tdfedtrustOrgCertificateNotAfterDateColor = "green"
        $tdfedtrustOrgCertificateNotAfterDateFL =    $fedtrust.OrgCertificate.NotAfter.DateTime
    }
    else
    {
        Write-Host -foregroundcolor Red " Federation Trust Certificate is Expired on " $fedtrust.OrgCertificate.NotAfter.DateTime
        $tdfedtrustOrgCertificateNotAfterDateColor = "red"
        $tdfedtrustOrgCertificateNotAfterDateFL =    $fedtrust.OrgCertificate.NotAfter.DateTime
    }
    Write-Host -foregroundcolor White " `Federation Trust Token Issuer Certificate Expiracy:"
    if ($fedtrust.TokenIssuerCertificate.NotAfter.DateTime -gt $CurrentTime){
        Write-Host -foregroundcolor Green "  Not Expired"
        Write-Host  "   - Expires on " $fedtrust.TokenIssuerCertificate.NotAfter.DateTime
        $tdfedtrustTokenIssuerCertificateNotAfterDateTimeColor = "green"
        $tdfedtrustTokenIssuerCertificateNotAfterDateTimeFL =    $fedtrust.TokenIssuerCertificate.NotAfter.DateTime
    }
    else
    {
        Write-Host -foregroundcolor Red "  Federation Trust TokenIssuerCertificate Expired on " $fedtrust.TokenIssuerCertificate.NotAfter.DateTime
        $tdfedtrustTokenIssuerCertificateNotAfterDateTimeColor = "red"
        $tdfedtrustTokenIssuerCertificateNotAfterDateTimeFL =   $fedtrust.TokenIssuerCertificate.NotAfter.DateTime
    }
    Write-Host -foregroundcolor White " Federation Trust Token Issuer Prev Certificate Expiracy:"
    if ($fedtrust.TokenIssuerPrevCertificate.NotAfter.Date -gt $CurrentTime){
        Write-Host -foregroundcolor Green "  Not Expired"
        Write-Host  "   - Expires on " $fedtrust.TokenIssuerPrevCertificate.NotAfter.DateTime
        $tdfedtrustTokenIssuerPrevCertificateNotAfterDateColor = "green"
        $tdfedtrustTokenIssuerPrevCertificateNotAfterDateFL =    $fedtrust.TokenIssuerPrevCertificate.NotAfter.DateTime
    }
    else
    {
        Write-Host -foregroundcolor Red "  Federation Trust TokenIssuerPrevCertificate Expired on " $fedtrust.TokenIssuerPrevCertificate.NotAfter.DateTime
        $tdfedtrustTokenIssuerPrevCertificateNotAfterDateColor = "red"
        $tdfedtrustTokenIssuerPrevCertificateNotAfterDateFL =    $fedtrust.TokenIssuerPrevCertificate.NotAfter.DateTime
    }
    $fedtrustTokenIssuerMetadataEpr = "https://nexus.microsoftonline-p.com/FederationMetadata/2006-12/FederationMetadata.xml"
    Write-Host -foregroundcolor White " `Token Issuer Metadata EPR:"
    if ($fedtrust.TokenIssuerMetadataEpr.AbsoluteUri -like $fedtrustTokenIssuerMetadataEpr){
        Write-Host -foregroundcolor Green "  Token Issuer Metadata EPR is " $fedtrust.TokenIssuerMetadataEpr.AbsoluteUri
        #test if it can be reached
        $tdfedtrustTokenIssuerMetadataEprAbsoluteUriColor = "green"
        $tdfedtrustTokenIssuerMetadataEprAbsoluteUriFL =    $fedtrust.TokenIssuerMetadataEpr.AbsoluteUri
    }
    else
    {
        Write-Host -foregroundcolor Red " Token Issuer Metadata EPR is Not " $fedtrust.TokenIssuerMetadataEpr.AbsoluteUri
        $tdfedtrustTokenIssuerMetadataEprAbsoluteUriColor = "red"
        $tdfedtrustTokenIssuerMetadataEprAbsoluteUriFL =    $fedtrust.TokenIssuerMetadataEpr.AbsoluteUri
    }
    $fedtrustTokenIssuerEpr = "https://login.microsoftonline.com/extSTS.srf"
    Write-Host -foregroundcolor White " Token Issuer EPR:"
    if ($fedtrust.TokenIssuerEpr.AbsoluteUri -like $fedtrustTokenIssuerEpr){
        Write-Host -foregroundcolor Green "  Token Issuer EPR is:" $fedtrust.TokenIssuerEpr.AbsoluteUri
        #test if it can be reached
        $tdfedtrustTokenIssuerEprAbsoluteUriColor = "green"
        $tdfedtrustTokenIssuerEprAbsoluteUriFL =    $fedtrust.TokenIssuerEpr.AbsoluteUri
    }
    else
    {
        Write-Host -foregroundcolor Red "  Token Issuer EPR is Not:" $fedtrust.TokenIssuerEpr.AbsoluteUri
        $tdfedtrustTokenIssuerEprAbsoluteUriColor = "red"
        $tdfedtrustTokenIssuerEprAbsoluteUriFL =    $fedtrust.TokenIssuerEpr.AbsoluteUri
    }
    $fedinfoTokenIssuerUris = $fedinfo.TokenIssuerUris
    $fedinfoTargetApplicationUri = $fedinfo.TargetApplicationUri
    $fedinfoTargetAutodiscoverEpr = $fedinfo.TargetAutodiscoverEpr
    $script:html += "
    <tr>
    <th colspan='2'>SUMMARY - Organization Test-FederationTrust</th>
    </tr>
    <tr>
    <td>Federation Trust Aplication Uri:</td>
    <td class='$tdfedtrustApplicationUricolor'>
    $tdfedtrustApplicationUriFL
    </td>
    </tr>
    <tr>
    <td>TokenIssuerUris:</td>
    <td class='$tdfedtrustTokenIssuerUriColor'>
    $tdfedtrustTokenIssuerUriFL
    </td>
    </tr>
    <tr>
    <td>Federation Trust Certificate Expiracy:</td>
    <td class='$tdfedtrustOrgCertificateNotAfterDateColor'>
    $tdfedtrustOrgCertificateNotAfterDateFL
    </td>
    </tr>
    <tr>
    <td>Federation Trust Token Issuer Certificate Expiracy:</td>
    <td class='$tdfedtrustTokenIssuerPrevCertificateNotAfterDateColor'>
    $tdfedtrustTokenIssuerPrevCertificateNotAfterDateFL
    </td>
    </tr>
    <tr>
    <td>Federation Trust Token Issuer Prev Certificate Expiracy:</td>
    <td class='$tdfedtrustTokenIssuerPrevCertificateNotAfterDateColor'>
    $tdfedtrustTokenIssuerPrevCertificateNotAfterDateFL
    </td>
    </tr>
    <tr>
    <td>Token Issuer Metadata EPR:</td>
    <td class='$tdfedtrustTokenIssuerMetadataEprAbsoluteUriColor'>
    $tdfedtrustTokenIssuerMetadataEprAbsoluteUriFL
    </td>
    </tr>
    <tr>
    <td>Token Issuer EPR:</td>
    <td class='$tdfedtrustTokenIssuerEprAbsoluteUriColor'>
    $tdfedtrustTokenIssuerEprAbsoluteUriFL
    </td>
    </tr>"
    $html| Out-File -FilePath "$PSScriptRoot\ShowParameters.html"
    Write-Host -foregroundcolor Yellow "`n  Reference: https://learn.microsoft.com/en-us/exchange/configure-a-federation-trust-exchange-2013-help"
}

Function AutoDVirtualDCheck{
    Write-Host -foregroundcolor Green " Get-AutodiscoverVirtualDirectory | Select Identity,Name,ExchangeVersion,*authentication*"
    Write-Host $bar
    $Global:AutoDiscoveryVirtualDirectory = Get-AutodiscoverVirtualDirectory | Select Identity,Name,ExchangeVersion,*authentication*
    #Check if null or set
    #$AutoDiscoveryVirtualDirectory
    $Global:AutoDiscoveryVirtualDirectory
    $AutoDFL = $Global:AutoDiscoveryVirtualDirectory | fl
    $script:html += "<tr>
    <th colspan='2'>On-Prem Autodiscover Virtual Directory</th>
    </tr>
    <tr>
    <td><b>Get-AutodiscoverVirtualDirectory:</b></td>
    <td>"
    Write-Host $bar
    Write-Host -foregroundcolor Green " SUMMARY - On-Prem Autodiscover Virtual Directory"
    Write-Host $bar
    Write-Host -ForegroundColor White "  WSSecurityAuthentication:"
    if ($Global:AutoDiscoveryVirtualDirectory.WSSecurityAuthentication -eq "True"){
        foreach( $ser in $Global:AutoDiscoveryVirtualDirectory) {
            Write-Host " $($ser.Identity) "
            $AutodVDIdentity = $ser.Identity
            $AutodVDName = $ser.Name
            $AutodVDInternalAuthenticationMethods = $ser.InternalAuthenticationMethods
            $AutodVDExternalAuthenticationMethods = $ser.ExternalAuthenticationMethods
            $AutodVDWSAuthetication = $ser.WSSecurityAuthentication
            $AutodVDWSAutheticationColor="green"
            $AutodVDWindowsAuthentication = $ser.WindowsAuthentication
            if ($AutodVDWindowsAuthentication -eq "True"){
                $AutodVDWindowsAuthenticationColor = "green"
            }
            else{
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
            Write-Host -ForegroundColor Green "  WSSecurityAuthentication: $($ser.WSSecurityAuthentication)"
            $serWSSecurityAuthenticationColor= "Green"
        }
    }
    else
    {
        Write-Host -foregroundcolor Red " WSSecurityAuthentication is NOT correct."
        foreach( $ser in $Global:AutoDiscoveryVirtualDirectory) {
            Write-Host " $($ser.Identity)"
            Write-Host -ForegroundColor Red "  WSSecurityAuthentication: $($ser.WSSecurityAuthentication)"
            $serWSSecurityAuthenticationColor= "Red"
            Write-Host " $($ser.Identity) "
            $AutodVDIdentity = $ser.Identity
            $AutodVDName = $ser.Name
            $AutodVDInternalAuthenticationMethods = $ser.InternalAuthenticationMethods
            $AutodVDExternalAuthenticationMethods = $ser.ExternalAuthenticationMethods
            $AutodVDWSAuthetication = $ser.WSSecurityAuthentication
            $AutodVDWSAutheticationColor="green"
            $AutodVDWindowsAuthentication = $ser.WindowsAuthentication
            if ($AutodVDWindowsAuthentication -eq "True"){
                $AutodVDWindowsAuthenticationColor = "green"
            }
            else{
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
            $serWSSecurityAuthenticationColor= "Red"
        }
        Write-Host -foregroundcolor White "  Should be True "
    }
    Write-Host -ForegroundColor White "`n  WindowsAuthentication:"
    if ($Global:AutoDiscoveryVirtualDirectory.WindowsAuthentication -eq "True"){
        foreach( $ser in $Global:AutoDiscoveryVirtualDirectory) {
            Write-Host " $($ser.Identity) "
            Write-Host -ForegroundColor Green "  WindowsAuthentication: $($ser.WindowsAuthentication)"
        }
    }
    else
    {
        Write-Host -foregroundcolor Red " WindowsAuthentication is NOT correct."
        foreach( $ser in $Global:AutoDiscoveryVirtualDirectory) {
            Write-Host " $($ser.Identity)"
            Write-Host -ForegroundColor Red "  WindowsAuthentication: $($ser.WindowsAuthentication)"
        }
        Write-Host -foregroundcolor White "  Should be True "
    }
    Write-Host -foregroundcolor Yellow "`n  Reference: https://learn.microsoft.com/en-us/powershell/module/exchange/get-autodiscovervirtualdirectory?view=exchange-ps"
    $html| Out-File -FilePath "$PSScriptRoot\ShowParameters.html"
}

Function EWSVirtualDirectoryCheck{
    Write-Host -foregroundcolor Green " Get-WebServicesVirtualDirectory | Select Identity,Name,ExchangeVersion,*Authentication*,*url"
    Write-Host $bar
    $Global:WebServicesVirtualDirectory = Get-WebServicesVirtualDirectory | Select Identity,Name,ExchangeVersion,*Authentication*,*url
    $Global:WebServicesVirtualDirectory
    Write-Host $bar
    Write-Host -foregroundcolor Green " SUMMARY - On-Prem Web Services Virtual Directory"
    Write-Host $bar
    $script:html += "
    <tr>
    <th colspan='2'>SUMMARY - On-Prem Web Services Virtual Directory</th>
    </tr>
    <tr>
    <td><b>Get-WebServicesVirtualDirectory:</b></td>
    <td >"
    Write-Host -foregroundcolor White "  WSSecurityAuthentication:"
    if ($Global:WebServicesVirtualDirectory.WSSecurityAuthentication -like  "True"){
        foreach( $EWS in $Global:WebServicesVirtualDirectory) {
            Write-Host " $($EWS.Identity)"
            Write-Host -ForegroundColor Green "  WSSecurityAuthentication: $($EWS.WSSecurityAuthentication) "
            $EWSVDIdentity = $EWS.Identity
            $EWSVDName = $EWS.Name
            $EWSVDInternalAuthenticationMethods = $EWS.InternalAuthenticationMethods
            $EWSVDExternalAuthenticationMethods = $EWS.ExternalAuthenticationMethods
            $EWSVDWSAuthetication = $EWS.WSSecurityAuthentication
            $EWSVDWSAutheticationColor="green"
            $EWSVDWindowsAuthentication = $EWS.WindowsAuthentication
            if ($EWSVDWindowsAuthentication -eq "True"){
                $EWSVDWindowsAuthenticationColor = "green"
            }
            else{
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
    else
    {
        Write-Host -foregroundcolor Red " WSSecurityAuthentication shoud be True."
        foreach( $EWS in $Global:AutoDiscoveryVirtualDirectory) {
            Write-Host " $($EWS.Identity) "
            Write-Host -ForegroundColor Red "  WSSecurityAuthentication: $($ser.WSSecurityAuthentication) "
            $EWSVDIdentity = $EWS.Identity
            $EWSVDName = $EWS.Name
            $EWSVDInternalAuthenticationMethods = $EWS.InternalAuthenticationMethods
            $EWSVDExternalAuthenticationMethods = $EWS.ExternalAuthenticationMethods
            $EWSVDWSAuthetication = $EWS.WSSecurityAuthentication
            $EWSVDWSAutheticationColor="green"
            $EWSVDWindowsAuthentication = $EWS.WindowsAuthentication
            if ($EWSVDWindowsAuthentication -eq "True"){
                $EWSVDWindowsAuthenticationColor = "green"
            }
            else{
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
    if ($Global:WebServicesVirtualDirectory.WindowsAuthentication -like  "True"){
        foreach( $EWS in $Global:WebServicesVirtualDirectory) {
            Write-Host " $($EWS.Identity)"
            Write-Host -ForegroundColor Green "  WindowsAuthentication: $($EWS.WindowsAuthentication) "
        }
    }
    else
    {
        Write-Host -foregroundcolor Red " WindowsAuthentication shoud be True."
        foreach( $EWS in $Global:AutoDiscoveryVirtualDirectory) {
            Write-Host " $($EWS.Identity) "
            Write-Host -ForegroundColor Red "  WindowsAuthentication: $($ser.WindowsAuthentication) "
        }
        Write-Host -foregroundcolor White "  Should be True"
    }
    $script:html += "
    </td>
    </tr>
    "
    $html| Out-File -FilePath "$PSScriptRoot\ShowParameters.html"
}

Function AvailabilityAddressSpaceCheck{
    $bar
    Write-Host -foregroundcolor Green " Get-AvailabilityAddressSpace $exchangeOnlineDomain | fl ForestName, UserName, UseServiceAccount, AccessMethod, ProxyUrl, Name"
    Write-Host $bar
    $AvailabilityAddressSpace = Get-AvailabilityAddressSpace $exchangeOnlineDomain -ErrorAction SilentlyContinue| select ForestName, UserName, UseServiceAccount, AccessMethod, ProxyUrl, Name
    If (!$AvailabilityAddressSpace){
        $AvailabilityAddressSpace = Get-AvailabilityAddressSpace $exchangeOnlineDomain -ErrorAction SilentlyContinue| select ForestName, UserName, UseServiceAccount, AccessMethod, ProxyUrl, Name
    }
    $AvailabilityAddressSpace
    $tdAvailabilityAddressSpaceName = $AvailabilityAddressSpace.Name
    Write-Host $bar
    Write-Host -foregroundcolor Green " SUMMARY - On-Prem Availability Address Space Check"
    Write-Host $bar
    Write-Host -foregroundcolor White " ForestName: "
    if ($AvailabilityAddressSpace.ForestName -like  $ExchangeOnlineDomain){
        Write-Host -foregroundcolor Green " " $AvailabilityAddressSpace.ForestName
        $tdAvailabilityAddressSpaceForestName = $AvailabilityAddressSpace.ForestName
        $tdAvailabilityAddressSpaceForestColor = "green"
    }
    else
    {
        Write-Host -foregroundcolor Red "  ForestName appears not to be correct."
        Write-Host -foregroundcolor White " Should contain the " $ExchaneOnlineDomain
        $tdAvailabilityAddressSpaceForestName = $AvailabilityAddressSpace.ForestName
        $tdAvailabilityAddressSpaceForestColor = "red"
    }
    Write-Host -foregroundcolor White " UserName: "
    if ($AvailabilityAddressSpace.UserName -like  ""){
        Write-Host -foregroundcolor Green "  Blank"
        $tdAvailabilityAddressSpaceUserName = $AvailabilityAddressSpace.UserName
        $tdAvailabilityAddressSpaceUserNameColor = "green"
    }
    else
    {
        Write-Host -foregroundcolor Red " UserName is NOT correct. "
        Write-Host -foregroundcolor White "  Normally it should be blank"
        $tdAvailabilityAddressSpaceUserName = $AvailabilityAddressSpace.UserName
        $tdAvailabilityAddressSpaceUserNameColor = "red"
    }
    Write-Host -foregroundcolor White " UseServiceAccount: "
    if ($AvailabilityAddressSpace.UseServiceAccount -like  "True"){
        Write-Host -foregroundcolor Green "  True"
        $tdAvailabilityAddressSpaceUseServiceAccount = $AvailabilityAddressSpace.UseServiceAccount
        $tAvailabilityAddressSpaceUseServiceAccountColor = "green"
    }
    else
    {
        Write-Host -foregroundcolor Red "  UseServiceAccount appears not to be correct."
        Write-Host -foregroundcolor White "  Should be True"
        $tdAvailabilityAddressSpaceUseServiceAccount = $AvailabilityAddressSpace.UseServiceAccount
        $tAvailabilityAddressSpaceUseServiceAccountColor = "red"
    }
    Write-Host -foregroundcolor White " AccessMethod:"
    if ($AvailabilityAddressSpace.AccessMethod -like  "InternalProxy"){
        Write-Host -foregroundcolor Green "  InternalProxy"
        $tdAvailabilityAddressSpaceAccessMethod = $AvailabilityAddressSpace.AccessMethod
        $tdAvailabilityAddressSpaceAccessMethodColor = "green"
    }
    else
    {
        Write-Host -foregroundcolor Red " AccessMethod appears not to be correct."
        Write-Host -foregroundcolor White " Should be InternalProxy"
        $tdAvailabilityAddressSpaceAccessMethod = $AvailabilityAddressSpace.AccessMethod
        $tdAvailabilityAddressSpaceAccessMethodColor = "red"
    }
    Write-Host -foregroundcolor White " ProxyUrl: "
    $tdAvailabilityAddressSpaceProxyUrl = $AvailabilityAddressSpace.ProxyUrl
    if ([String]::Equals($tdAvailabilityAddressSpaceProxyUrl, $Global:ExchangeOnPremEWS, [StringComparison]::OrdinalIgnoreCase))
    {
        Write-Host -foregroundcolor Green " "$AvailabilityAddressSpace.ProxyUrl
        #$tdAvailabilityAddressSpaceProxyUrl = $AvailabilityAddressSpace.ProxyUrl
        $tdAvailabilityAddressSpaceProxyUrlColor = "green"
    }
    else
    {
        Write-Host -foregroundcolor Red "  ProxyUrl appears not to be correct."
        Write-Host -foregroundcolor White "  Should be $Global:ExchangeOnPremEWS[0] and not $tdAvailabilityAddressSpaceProxyUrl"
        #$tdAvailabilityAddressSpaceProxyUrl = $AvailabilityAddressSpace.ProxyUrl
        $tdAvailabilityAddressSpaceProxyUrlColor = "red"
    }
    Write-Host -ForegroundColor Yellow "  Reference: https://learn.microsoft.com/en-us/powershell/module/exchange/get-availabilityaddressspace?view=exchange-ps"
    $script:html += "
    <tr>
    <th colspan='2'>SUMMARY - On-Prem Availability Address Space</th>
    </tr>
    <tr>
    <td><b>Get-AvailabilityAddressSpace:</b></td>
    <td>
    <div> <b>Forest Name: </b> $tdAvailabilityAddressSpaceForestName</div>
    <div> <b>Name: </b> $tdAvailabilityAddressSpaceName</div>
    <div> <b>UserName: </b> <span style='color:$tdAvailabilityAddressSpaceUserNameColor'>$tdAvailabilityAddressSpaceUserName</span></div>
    <div> <b>Access Method: </b> <span style='color:$tdAvailabilityAddressSpaceAccessMethodColor'>$tdAvailabilityAddressSpaceAccessMethod</span></div>
    <div> <b>ProxyUrl: </b> <span style='color:$tdAvailabilityAddressSpaceProxyUrlColor'>$tdAvailabilityAddressSpaceProxyUrl</span></div>
    </td>
    </tr>"
    $html| Out-File -FilePath "$PSScriptRoot\ShowParameters.html"
}

Function TestFedTrust{
    Write-Host $bar
    $TestFedTrustFail = 0
    $a = Test-FederationTrust -UserIdentity $useronprem -verbose -ErrorAction silentlycontinue #fails the frist time on multiple ocasions so we have a gohst FedTrustCheck
    Write-Host -foregroundcolor Green  " Test-FederationTrust -UserIdentity $useronprem -verbose"
    Write-Host $bar
    $TestFedTrust = Test-FederationTrust -UserIdentity $useronprem -verbose -ErrorAction silentlycontinue
    $TestFedTrust
    $Script:html += "<tr>
    <th colspan='2'><b>Test Federation Trust</b></th>
    </tr>
    <tr>
    <td><b>Test-FederationTrust:</b></td>
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
            $Script:html +="
            
            <div> <span style='color:red'><b>$testType :</b></span> - <div> <b>$TestFedTrustID </b> - $testMessage  </div>
            "
            $TestFedTrustFail++
        }
        if ($test -eq "Success") {
            # Write-Host " $($TestFedTrust.ID[$i]) "
            # Write-Host -foregroundcolor Green " $($TestFedTrust.Type[$i])  "
            # Write-Host " $($TestFedTrust.Message[$i])  "
            $Script:html +="
            
            <div> <span style='color:green'><b>$testType :</b> </span> - <b>$TestFedTrustID </b> - $testMessage</div>"
        }
        $i++
    }

    if ($TestFedTrustFail -eq  0){
        Write-Host -foregroundcolor Green " Federation Trust Successfully tested"
        $Script:html +="
        <p></p>
        <div class=´green´> <span style='color:green'> Federation Trust Successfully tested </span></div>"
    }
    else  {
        Write-Host -foregroundcolor Red " Federation Trust test with Errors"
        $Script:html +="
        <p></p>
        <div class=´red´> <span style='color:red'> Federation Trust tested with Errors </span></div>"
    }
    $html| Out-File -FilePath "$PSScriptRoot\ShowParameters.html"
}

Function TestOrgRel{
    $bar
    $TestFail = 0
    $OrgRelIdentity=$OrgRel.Identity
    $Script:html += "<tr>
    <th colspan='2'><b>Test Organization Relationship</b></th>
    </tr>
    <tr>
    <td><b>Test-OrganizationRelationship:</b></td>
    <td>"
    Write-Host -foregroundcolor Green "Test-OrganizationRelationship -Identity $OrgRelIdentity  -UserIdentity $useronprem"
    #need to grab errors and provide alerts in error case
    Write-Host $bar
    $TestOrgRel = Test-OrganizationRelationship -Identity $OrgRelIdentity  -UserIdentity $useronprem -erroraction SilentlyContinue -warningaction SilentlyContinue
    #$TestOrgRel
    if ($TestOrgRel[16] -like "No Significant Issues to Report")
    {
        Write-Host -foregroundcolor Green "`n No Significant Issues to Report"
        $Script:html +="
        <div class='green'> <b>No Significant Issues to Report</b><div>"
    }
    else
    {
        Write-Host -foregroundcolor Red "`n Test Organization Relationship Completed with errors"
        $Script:html +="
        <div class='red'> <b>Test Organization Relationship Completed with errors</b><div>"
    }
    $TestOrgRel[0]
    $TestOrgRel[1]
    $i = 0
    while ($i -lt $TestOrgRel.Length) {
        $element = $TestOrgRel[$i]
        if ($element.Contains("RESULT: Success.")) {
            $TestOrgRelStep=$TestOrgRel[$i-1]
            $TestOrgRelStep
            Write-Host -ForegroundColor Green "$element"
            $Script:html +="
            <div> <b> $TestOrgRelStep :</b> <span style='color:green'>$element</span>"
        }
        else  {
            if ($element.Contains("RESULT: Error")) {
                $TestOrgRelStep=$TestOrgRel[$i-1]
                $TestOrgRelStep
                Write-Host -ForegroundColor Red "$element"
                $Script:html +="
                div> <b> $TestOrgRelStep : </b> <span style='color:red'>$element</span>"
            }
        }
        $i++
    }
    if ($TestFail -eq "0"){
        #Write-Host -foregroundcolor Green " Organization Relationship Successfully tested `n "
        $Script:html +="
        $testType : $testMessage"
    }
    else  {
        $Script:html +="
        Organization Relationship test with Errors"
        #Check this an that
    }
    Write-host -ForegroundColor Yellow "`n  Reference: https://techcommunity.microsoft.com/t5/exchange-team-blog/how-to-address-federation-trust-issues-in-hybrid-configuration/ba-p/1144285"
    Write-Host $bar
    $Script:html +=  "</td>
    </tr>"
    $html| Out-File -FilePath "$PSScriptRoot\ShowParameters.html"
}

#endregion

#region OAuth Functions

Function IntraOrgConCheck {
    Write-Host -foregroundcolor Green " Get-IntraOrganizationConnector | Selecct Name,TargetAddressDomains,DiscoveryEndpoint,Enabled"
    Write-Host $bar
    $IOC = $IntraOrgCon | Format-List
    $IOC
    Write-Host $bar
    Write-Host -foregroundcolor Green " SUMMARY - On-Prem Intra Organization Connector"
    Write-Host $bar
    $IntraOrgTargetAddressDomain = $IntraOrgCon.TargetAddressDomains.Domain
    $IntraOrgTargetAddressDomain = $IntraOrgTargetAddressDomain.Tolower()
    Write-Host -foregroundcolor White " Target Address Domains: "
    if ($IntraOrgCon.TargetAddressDomains -like "*$ExchangeOnlineDomain*" -Or $IntraOrgCon.TargetAddressDomains -like "*$ExchangeOnlineAltDomain*" ) {
        Write-Host -foregroundcolor Green " " $IntraOrgCon.TargetAddressDomains
    }
    else {
        Write-Host -foregroundcolor Red " Target Address Domains appears not to be correct."
        Write-Host -foregroundcolor White " Should contain the $ExchangeOnlineDomain domain or the $ExchangeOnlineAltDomain domain."
    }
    Write-Host -foregroundcolor White " DiscoveryEndpoint: "
    if ($IntraOrgCon.DiscoveryEndpoint -like "https://autodiscover-s.outlook.com/autodiscover/autodiscover.svc") {
        Write-Host -foregroundcolor Green "  https://autodiscover-s.outlook.com/autodiscover/autodiscover.svc"
    }
    else {
        Write-Host -foregroundcolor Red "  The DiscoveryEndpoint appears not to be correct. "
        Write-Host -foregroundcolor White "  It should represent the address of EXO autodiscover endpoint."
        Write-Host  "  Examples: https://autodiscover-s.outlook.com/autodiscover/autodiscover.svc; https://outlook.office365.com/autodiscover/autodiscover.svc "
    }
    Write-Host -foregroundcolor White " Enabled: "
    if ($IntraOrgCon.Enabled -like "True") {
        Write-Host -foregroundcolor Green "  True "
    }
    else {
        Write-Host -foregroundcolor Red "  On-Prem Intra Organization Connector is not Enabled"
        Write-Host -foregroundcolor White "  In order to use OAuth it Should be True."
        write-Host "  If it is set to False, the Organization Realtionship (DAuth) , if enabled, is used for the Hybrid Availability Sharing"
    }
    Write-Host -ForegroundColor Yellow "https://techcommunity.microsoft.com/t5/exchange-team-blog/demystifying-hybrid-free-busy-what-are-the-moving-parts/ba-p/607704"
}

Function AuthServerCheck {
    #Write-Host $bar
    Write-Host -foregroundcolor Green " Get-AuthServer | Select Name,IssuerIdentifier,TokenIssuingEndpoint,AuthMetadataUrl,Enabled"
    Write-Host $bar
    $AuthServer = Get-AuthServer | Where-Object { $_.Name -like "ACS*" } | Select-Object Name, IssuerIdentifier, TokenIssuingEndpoint, AuthMetadataUrl, Enabled
    $AuthServer
    Write-Host $bar
    Write-Host -foregroundcolor Green " SUMMARY - Auth Server"
    Write-Host $bar
    Write-Host -foregroundcolor White " IssuerIdentifier: "
    if ($AuthServer.IssuerIdentifier -like "00000001-0000-0000-c000-000000000000" ) {
        Write-Host -foregroundcolor Green " " $AuthServer.IssuerIdentifier
    }
    else {
        Write-Host -foregroundcolor Red " IssuerIdentifier appears not to be correct."
        Write-Host -foregroundcolor White " Should be 00000001-0000-0000-c000-000000000000"
    }
    Write-Host -foregroundcolor White " TokenIssuingEndpoint: "
    if ($AuthServer.TokenIssuingEndpoint -like "https://accounts.accesscontrol.windows.net/*" -and $AuthServer.TokenIssuingEndpoint -like "*/tokens/OAuth/2" ) {
        Write-Host -foregroundcolor Green " " $AuthServer.TokenIssuingEndpoint
    }
    else {
        Write-Host -foregroundcolor Red " TokenIssuingEndpoint appears not to be correct."
        Write-Host -foregroundcolor White " Should be  https://accounts.accesscontrol.windows.net/<Cloud Tenant ID>/tokens/OAuth/2"
    }
    Write-Host -foregroundcolor White " AuthMetadataUrl: "
    if ($AuthServer.AuthMetadataUrl -like "https://accounts.accesscontrol.windows.net/*" -and $AuthServer.TokenIssuingEndpoint -like "*/tokens/OAuth/2" ) {
        Write-Host -foregroundcolor Green " " $AuthServer.AuthMetadataUrl
    }
    else {
        Write-Host -foregroundcolor Red " AuthMetadataUrl appears not to be correct."
        Write-Host -foregroundcolor White " Should be  https://accounts.accesscontrol.windows.net/<Cloud Tenant ID>/metadata/json/1"
    }
    Write-Host -foregroundcolor White " Enabled: "
    if ($AuthServer.Enabled -like "True" ) {
        Write-Host -foregroundcolor Green " " $AuthServer.Enabled
    }
    else {
        Write-Host -foregroundcolor Red " Enalbed: False "
        Write-Host -foregroundcolor White " Should be True"
    }
}

Function PartnerApplicationCheck {
    Write-Host $bar
    Write-Host -foregroundcolor Green " Get-PartnerApplication |  ?{$_.ApplicationIdentifier -eq '00000002-0000-0ff1-ce00-000000000000'
    -and $_.Realm -eq ''} | Select Enabled, ApplicationIdentifier, CertificateStrings, AuthMetadataUrl, Realm, UseAuthServer,
    AcceptSecurityIdentifierInformation, LinkedAccount, IssuerIdentifier, AppOnlyPermissions, ActAsPermissions, Name"
    Write-Host $bar
    $PartnerApplication = Get-PartnerApplication |  Where-Object { $_.ApplicationIdentifier -eq '00000002-0000-0ff1-ce00-000000000000' -and $_.Realm -eq '' } | Select-Object Enabled, ApplicationIdentifier, CertificateStrings, AuthMetadataUrl, Realm, UseAuthServer, AcceptSecurityIdentifierInformation, LinkedAccount, IssuerIdentifier, AppOnlyPermissions, ActAsPermissions, Name
    $PartnerApplication
    Write-Host $bar
    Write-Host -foregroundcolor Green " SUMMARY - Partner Application"
    Write-Host $bar
    Write-Host -foregroundcolor White " Enabled: "
    if ($PartnerApplication.Enabled -like "True" ) {
        Write-Host -foregroundcolor Green " " $PartnerApplication.Enabled
    }
    else {
        Write-Host -foregroundcolor Red " Enabled: False "
        Write-Host -foregroundcolor White " Should be True"
    }
    Write-Host -foregroundcolor White " ApplicationIdentifier: "
    if ($PartnerApplication.ApplicationIdentifier -like "00000002-0000-0ff1-ce00-000000000000" ) {
        Write-Host -foregroundcolor Green " " $PartnerApplication.ApplicationIdentifier
    }
    else {
        Write-Host -foregroundcolor Red " ApplicationIdentifier does not appear to be correct"
        Write-Host -foregroundcolor White " Should be 00000002-0000-0ff1-ce00-000000000000"
    }
    Write-Host -foregroundcolor White " AuthMetadataUrl: "
    if ([string]::IsNullOrWhitespace( $PartnerApplication.AuthMetadataUrl)) {
        Write-Host -foregroundcolor Green "  Blank"
    }
    else {
        Write-Host -foregroundcolor Red " AuthMetadataUrl does not aooear correct"
        Write-Host -foregroundcolor White " Should be Blank"
    }
    Write-Host -foregroundcolor White " Realm: "
    if ([string]::IsNullOrWhitespace( $PartnerApplication.Realm)) {
        Write-Host -foregroundcolor Green "  Blank"
    }
    else {
        Write-Host -foregroundcolor Red "  Realm does not appear to be correct"
        Write-Host -foregroundcolor White " Should be Blank"
    }
    Write-Host -foregroundcolor White " LinkedAccount: "
    if ($PartnerApplication.LinkedAccount -like "$exchangeOnPremDomain/Users/Exchange Online-ApplicationAccount" -or $PartnerApplication.LinkedAccount -like "$exchangeOnPremLocalDomain/Users/Exchange Online-ApplicationAccount"  ) {
        Write-Host -foregroundcolor Green " " $PartnerApplication.LinkedAccount
    }
    else {
        Write-Host -foregroundcolor Red "  LinkedAccount value does not appear to be correct"
        Write-Host -foregroundcolor White "  Should be $exchangeOnPremLocalDomain/Users/Exchange Online-ApplicationAccount"
        Write-Host "  If you value is empty, set it to correspond to the Exchange Online-ApplicationAccount which is located at the root of Users container in AD. After you make the change, reboot the servers."
        Write-Host "  Example: contoso.com/Users/Exchange Online-ApplicationAccount"
    }
}

Function ApplicationAccounCheck {
    Write-Host $bar
    Write-Host -foregroundcolor Green " Get-user '$exchangeOnPremLocalDomain/Users/Exchange Online-ApplicationAccount' | Select Name, RecipientType, RecipientTypeDetails, UserAccountControl"
    Write-Host $bar
    $ApplicationAccount = Get-user "$exchangeOnPremLocalDomain/Users/Exchange Online-ApplicationAccount"  | Select-Object Name, RecipientType, RecipientTypeDetails, UserAccountControl
    $ApplicationAccount
    Write-Host $bar
    Write-Host -foregroundcolor Green " SUMMARY - Application Account"
    Write-Host $bar
    Write-Host -foregroundcolor White " RecipientType: "
    if ($ApplicationAccount.RecipientType -like "User" ) {
        Write-Host -foregroundcolor Green " " $ApplicationAccount.RecipientType
    }
    else {
        Write-Host -foregroundcolor Red " RecipientType value is $ApplicationAccount.RecipientType "
        Write-Host -foregroundcolor White " Should be User"
    }
    Write-Host -foregroundcolor White " RecipientTypeDetails: "
    if ($ApplicationAccount.RecipientTypeDetails -like "LinkedUser" ) {
        Write-Host -foregroundcolor Green " " $ApplicationAccount.RecipientTypeDetails
    }
    else {
        Write-Host -foregroundcolor Red " RecipientTypeDetails value is $ApplicationAccount.RecipientTypeDetails"
        Write-Host -foregroundcolor White " Should be LinkedUser"
    }
    Write-Host -foregroundcolor White " UserAccountControl: "
    if ($ApplicationAccount.UserAccountControl -like "AccountDisabled, PasswordNotRequired, NormalAccount" ) {
        Write-Host -foregroundcolor Green " " $ApplicationAccount.UserAccountControl
    }
    else {
        Write-Host -foregroundcolor Red " UserAccountControl value does not seem correct"
        Write-Host -foregroundcolor White " Should be AccountDisabled, PasswordNotRequired, NormalAccount"
    }
}

Function ManagementRoleAssignmentCheck {
    Write-Host -foregroundcolor Green " Get-ManagementRoleAssignment -RoleAssignee Exchange Online-ApplicationAccount | Select Name,Role -AutoSize"
    Write-Host $bar
    $ManagementRoleAssignment = Get-ManagementRoleAssignment -RoleAssignee "Exchange Online-ApplicationAccount"  | Select-Object Name, Role
    $M = $ManagementRoleAssignment | Out-String
    $M
    Write-Host $bar
    Write-Host -foregroundcolor Green " SUMMARY - Management Role Assignment for the Exchange Online-ApplicationAccount"
    Write-Host $bar
    Write-Host -foregroundcolor White " Role: "
    if ($ManagementRoleAssignment.Role -like "*UserApplication*" ) {
        Write-Host -foregroundcolor Green "  UserApplication Role Assigned"
    }
    else {
        Write-Host -foregroundcolor Red "  UserApplication Role not present for the Exchange Online-ApplicationAccount"
    }
    if ($ManagementRoleAssignment.Role -like "*ArchiveApplication*" ) {
        Write-Host -foregroundcolor Green "  ArchiveApplication Role Assigned"
    }
    else {
        Write-Host -foregroundcolor Red "  ArchiveApplication Role not present for the Exchange Online-ApplicationAccount"
    }
    if ($ManagementRoleAssignment.Role -like "*LegalHoldApplication*" ) {
        Write-Host -foregroundcolor Green "  LegalHoldApplication Role Assigned"
    }
    else {
        Write-Host -foregroundcolor Red "  LegalHoldApplication Role not present for the Exchange Online-ApplicationAccount"
    }
    if ($ManagementRoleAssignment.Role -like "*Mailbox Search*" ) {
        Write-Host -foregroundcolor Green "  Mailbox Search Role Assigned"
    }
    else {
        Write-Host -foregroundcolor Red "  Mailbox Search Role not present for the Exchange Online-ApplicationAccount"
    }
    if ($ManagementRoleAssignment.Role -like "*TeamMailboxLifecycleApplication*" ) {
        Write-Host -foregroundcolor Green "  TeamMailboxLifecycleApplication Role Assigned"
    }
    else {
        Write-Host -foregroundcolor Red "  TeamMailboxLifecycleApplication Role not present for the Exchange Online-ApplicationAccount"
    }
    if ($ManagementRoleAssignment.Role -like "*MailboxSearchApplication*" ) {
        Write-Host -foregroundcolor Green "  MailboxSearchApplication Role Assigned"
    }
    else {
        Write-Host -foregroundcolor Red "  MailboxSearchApplication Role not present for the Exchange Online-ApplicationAccount"
    }
    if ($ManagementRoleAssignment.Role -like "*MeetingGraphApplication*" ) {
        Write-Host -foregroundcolor Green "  MeetingGraphApplication Role Assigned"
    }
    else {
        Write-Host -foregroundcolor Red "  MeetingGraphApplication Role not present for the Exchange Online-ApplicationAccount"
    }
}

Function AuthConfigCheck {
    Write-Host -foregroundcolor Green " Get-AuthConfig | Select *Thumbprint, ServiceName, Realm, Name"
    Write-Host $bar
    $AuthConfig = Get-AuthConfig | Select-Object *Thumbprint, ServiceName, Realm, Name
    $AC = $AuthConfig | Format-List
    $AC
    Write-Host $bar
    Write-Host -foregroundcolor Green " SUMMARY - Auth Config"
    Write-Host $bar
    if (![string]::IsNullOrWhitespace($AuthConfig.CurrentCertificateThumbprint)) {
        Write-HOst " Thumbprint: "$AuthConfig.CurrentCertificateThumbprint
        Write-Host -foregroundcolor Green " Certificate is Assigned"
    }
    else {
        Write-HOst " Thumbprint: "$AuthConfig.CurrentCertificateThumbprint
        Write-Host -foregroundcolor Red " No valid certificate Assigned "
    }
    if ($AuthConfig.ServiceName -like "00000002-0000-0ff1-ce00-000000000000" ) {
        Write-HOst " ServiceName: "$AuthConfig.ServiceName
        Write-Host -foregroundcolor Green " Service Name Seems correct"
    }
    else {
        Write-HOst " ServiceName: "$AuthConfig.ServiceName
        Write-Host -foregroundcolor Red " Service Name does not Seems correct. Should be 00000002-0000-0ff1-ce00-000000000000"
    }
    if ([string]::IsNullOrWhitespace($AuthConfig.Realm)) {
        Write-HOst " Realm: "
        Write-Host -foregroundcolor Green " Realm is Blank"
    }
    else {
        Write-HOst " Realm: "$AuthConfig.Realm
        Write-Host -foregroundcolor Red " Realm should be Blank"
    }
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
    Write-Host -foregroundcolor Green " SUMMARY - Microsoft Exchange Server Auth Certificate"
    Write-Host $bar
    if ($CurrentCertificate.Issuer -like "CN=Microsoft Exchange Server Auth Certificate" ) {
        write-Host " Issuer: " $CurrentCertificate.Issuer
        Write-Host -foregroundcolor Green "  Issuer is CN=Microsoft Exchange Server Auth Certificate"
    }
    else {
        Write-Host -foregroundcolor Red "  Issuer is not CN=Microsoft Exchange Server Auth Certificate"
    }
    if ($CurrentCertificate.Services -like "SMTP" ) {
        Write-Host " Services: " $CurrentCertificate.Services
        Write-Host -foregroundcolor Green "  Certificate enabled for SMTP"
    }
    else {
        Write-Host -foregroundcolor Red "  Certificate Not enabled for SMTP"
    }
    if ($CurrentCertificate.Status -like "Valid" ) {
        Write-Host " Status: " $CurrentCertificate.Status
        Write-Host -foregroundcolor Green "  Certificate is valid"
    }
    else {
        Write-Host -foregroundcolor Red "  Certificate is not Valid"
    }
    if ($CurrentCertificate.Subject -like "CN=Microsoft Exchange Server Auth Certificate" ) {
        Write-Host " Subject: " $CurrentCertificate.Subject
        Write-Host -foregroundcolor Green "  Subject is CN=Microsoft Exchange Server Auth Certificate"
    }
    else {
        Write-Host -foregroundcolor Red "  Subject is not CN=Microsoft Exchange Server Auth Certificate"
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
        }
        if ($serv.Thumbprint -ne $thumbprint) {
            Write-Host -foregroundcolor Red "  Auth Certificate seems not to be present in "$servername
        }
    }
}

Function AutoDVirtualDCheckOauth {
    #Write-Host -foregroundcolor Green " `n On-Prem Autodiscover Virtual Directory `n "
    Write-Host -foregroundcolor Green " Get-AutodiscoverVirtualDirectory | Select Identity, Name,ExchangeVersion,*authentication*"
    Write-Host $bar
    $AutoDiscoveryVirtualDirectoryOAuth = Get-AutodiscoverVirtualDirectory | Select-Object Identity, Name, ExchangeVersion, *authentication*
    #Check if null or set
    $AD = $AutoDiscoveryVirtualDirectoryOAuth | Format-List
    $AD
    if ($Auth -like "OAuth") {
    }
    Write-Host $bar
    Write-Host -foregroundcolor Green " SUMMARY - On-Prem Autodiscover Virtual Directory"
    Write-Host $bar
    Write-Host -foregroundcolor White "  InternalAuthenticationMethods"
    if ($AutoDiscoveryVirtualDirectoryOAuth.InternalAuthenticationMethods -like "*OAuth*") {
        foreach ( $EWS in $AutoDiscoveryVirtualDirectoryOAuth) {
            Write-Host " $($EWS.Identity) "
            Write-Host -ForegroundColor Green "  InternalAuthenticationMethods Include OAuth Authentication Method "
        }
    }
    else {
        Write-Host -foregroundcolor Red "  InternalAuthenticationMethods seems not to include OAuth Authentication Method."
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
        Write-Host -foregroundcolor Red "  WSSecurityAuthentication setting is NOT correct."
        foreach ( $ADVD in $AutoDiscoveryVirtualDirectoryOAuth) {
            Write-Host " $($ADVD.Identity) "
            Write-Host -ForegroundColor Red "  WSSecurityAuthentication: $($ADVD.WSSecurityAuthentication)"
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
}

Function EWSVirtualDirectoryCheckOAuth {
    Write-Host -foregroundcolor Green " Get-WebServicesVirtualDirectory | Select Identity,Name,ExchangeVersion,*Authentication*,*url"
    Write-Host $bar
    $WebServicesVirtualDirectoryOAuth = Get-WebServicesVirtualDirectory | Select-Object Identity, Name, ExchangeVersion, *Authentication*, *url
    $W = $WebServicesVirtualDirectoryOAuth | Format-List
    $W
    if ($Auth -like "OAuth") {
    }
    Write-Host $bar
    Write-Host -foregroundcolor Green " SUMMARY - On-Prem Web Services Virtual Directory"
    Write-Host $bar
    Write-Host -foregroundcolor White "  InternalAuthenticationMethods"
    if ($WebServicesVirtualDirectoryOAuth.InternalAuthenticationMethods -like "*OAuth*") {
        foreach ( $EWS in $WebServicesVirtualDirectoryOAuth) {
            Write-Host " $($EWS.Identity) "
            Write-Host -ForegroundColor Green "  InternalAuthenticationMethods Include OAuth Authentication Method "
        }
    }
    else {
        Write-Host -foregroundcolor Red "  InternalAuthenticationMethods seems not to include OAuth Authentication Method."
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
}

Function AvailabilityAddressSpaceCheckOAuth {
    Write-Host -foregroundcolor Green " Get-AvailabilityAddressSpace $exchangeOnlineDomain | Select ForestName, UserName, UseServiceAccount, AccessMethod, ProxyUrl, Name"
    Write-Host $bar
    $AvailabilityAddressSpace = Get-AvailabilityAddressSpace $exchangeOnlineDomain | Select-Object ForestName, UserName, UseServiceAccount, AccessMethod, ProxyUrl, Name
    $AAS = $AvailabilityAddressSpace | Format-List
    $AAS
    if ($Auth -like "OAuth") {
    }
    Write-Host $bar
    Write-Host -foregroundcolor Green " SUMMARY - On-Prem Availability Address Space"
    Write-Host $bar
    Write-Host -foregroundcolor White " ForestName: "
    if ($AvailabilityAddressSpace.ForestName -like $ExchangeOnlineDomain) {
        Write-Host -foregroundcolor Green " "$AvailabilityAddressSpace.ForestName
    }
    else {
        Write-Host -foregroundcolor Red " ForestName is NOT correct. "
        Write-Host -foregroundcolor White " Should be $ExchaneOnlineDomain "
    }
    Write-Host -foregroundcolor White " UserName: "
    if ($AvailabilityAddressSpace.UserName -like "") {
        Write-Host -foregroundcolor Green "  Blank "
    }
    else {
        Write-Host -foregroundcolor Red "  UserName is NOT correct. "
        Write-Host -foregroundcolor White "  Should be blank "
    }
    Write-Host -foregroundcolor White " UseServiceAccount: "
    if ($AvailabilityAddressSpace.UseServiceAccount -like "True") {
        Write-Host -foregroundcolor Green "  True "
    }
    else {
        Write-Host -foregroundcolor Red "  UseServiceAccount is NOT correct."
        Write-Host -foregroundcolor White "  Should be True "
    }
    Write-Host -foregroundcolor White " AccessMethod: "
    if ($AvailabilityAddressSpace.AccessMethod -like "InternalProxy") {
        Write-Host -foregroundcolor Green "  InternalProxy "
    }
    else {
        Write-Host -foregroundcolor Red "  AccessMethod is NOT correct. "
        Write-Host -foregroundcolor White "  Should be InternalProxy "
    }
    Write-Host -foregroundcolor White " ProxyUrl: "
    if ($AvailabilityAddressSpace.ProxyUrl -like $exchangeOnPremEWS) {
        Write-Host -foregroundcolor Green " "$AvailabilityAddressSpace.ProxyUrl
    }
    else {
        Write-Host -foregroundcolor Red "  ProxyUrl is NOT correct. "
        Write-Host -foregroundcolor White "  Should be $exchangeOnPremEWS"
    }
}

Function OAuthConnectivityCheck {
    Write-Host -foregroundcolor Green " Test-OAuthConnectivity -Service EWS -TargetUri https://outlook.office365.com/EWS/Exchange.asmx -Mailbox $useronprem | fl"
    Write-Host $bar
    #$OAuthConnectivity = Test-OAuthConnectivity -Service EWS -TargetUri https://outlook.office365.com/EWS/Exchange.asmx -Mailbox $useronprem | fl
    #$OAuthConnectivity
    $OAuthConnectivity = Test-OAuthConnectivity -Service EWS -TargetUri https://outlook.office365.com/EWS/Exchange.asmx -Mailbox $useronprem
    if ($OAuthConnectivity.ResultType -eq 'Sucess' ) {
        $OAuthConnectivity.ResultType
        } else {
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
        Write-Host $bar
        #$OAuthConnectivity.detail.LocalizedString
        Write-Host -foregroundcolor Green " SUMMARY - Test OAuth COnnectivity"
        Write-Host $bar
        if ($OAuthConnectivity.ResultType -like "Success") {
            Write-Host -foregroundcolor Green " OAuth Test was completed successfully "
        }
        else {
            Write-Host -foregroundcolor Red " OAuth Test was completed with Error. "
            Write-Host -foregroundcolor White " Please rerun Test-OAuthConnectivity -Service EWS -TargetUri https://outlook.office365.com/EWS/Exchange.asmx -Mailbox <On Premises Mailbox> | fl to confirm the test failure"
        }
    }
    Write-Host -foregroundcolor Green " Note:"
    Write-Host -foregroundcolor Yellow " You can ignore the warning 'The SMTP address has no mailbox associated with it'"
    Write-Host -foregroundcolor Yellow " when the Test-OAuthConnectivity returns a Success"
    Write-Host -foregroundcolor Green " Reference: "
    Write-Host -foregroundcolor White " Configure OAuth authentication between Exchange and Exchange Online organizations"
    Write-Host -foregroundcolor Yellow " https://technet.microsoft.com/en-us/library/dn594521(v=exchg.150).aspx"
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
    }
    else {
        Write-Host -foregroundcolor Red "  Domain Names do Not Include the $ExchangeOnPremDomain Domain"
        $exoOrgRel.DomainNames
    }
    #FreeBusyAccessEnabled
    Write-Host  " FreeBusyAccessEnabled:"
    if ($exoOrgRel.FreeBusyAccessEnabled -like "True" ) {
        Write-Host -foregroundcolor Green "  FreeBusyAccessEnabled is set to True"
    }
    else {
        Write-Host -foregroundcolor Red "  FreeBusyAccessEnabled : False"
        #$countOrgRelIssues++
    }
    #FreeBusyAccessLevel
    Write-Host  " FreeBusyAccessLevel:"
    if ($exoOrgRel.FreeBusyAccessLevel -like "AvailabilityOnly" ) {
        Write-Host -foregroundcolor Green "  FreeBusyAccessLevel is set to AvailabilityOnly"
    }
    if ($exoOrgRel.FreeBusyAccessLevel -like "LimitedDetails" ) {
        Write-Host -foregroundcolor Green "  FreeBusyAccessLevel is set to LimitedDetails"
    }
    #fix porque este else só respeita o if anterior
    if ($exoOrgRel.FreeBusyAccessLevel -NE "AvailabilityOnly" -AND $exoOrgRel.FreeBusyAccessLevel -NE "LimitedDetails") {
        Write-Host -foregroundcolor Red "  FreeBusyAccessEnabled : False"
        #$countOrgRelIssues++
    }
    #TargetApplicationUri
    Write-Host  " TargetApplicationUri:"
    if ($exoOrgRel.TargetApplicationUri -like $fedtrust.ApplicationUri) {
        Write-Host -foregroundcolor Green "  TargetApplicationUri is" $fedtrust.ApplicationUri.originalstring
    }
    else {
        Write-Host -foregroundcolor Red "  TargetApplicationUri should be " $fedtrust.ApplicationUri.originalstring
        #$countOrgRelIssues++
    }
    #TargetSharingEpr
    Write-Host  " TargetSharingEpr:"
    if ([string]::IsNullOrWhitespace($exoOrgRel.TargetSharingEpr)) {
        Write-Host -foregroundcolor Green "  TargetSharingEpr is blank. This is the standard Value."
    }
    else {
        Write-Host -foregroundcolor Red "  TargetSharingEpr should be blank. If it is set, it should be the On-Premises Exchange servers EWS ExternalUrl endpoint."
        #$countOrgRelIssues++
    }
    #TargetAutodiscoverEpr:
    Write-Host  " TargetAutodiscoverEpr:"
    #Write-Host  "  OrgRel: " $exoOrgRel.TargetAutodiscoverEpr
    #Write-Host  "  FedInfo: " $fedinfoEOP
    #Write-Host  "  FedInfoEPR: " $fedinfoEOP.TargetAutodiscoverEpr
    if ($exoOrgRel.TargetAutodiscoverEpr -like $fedinfoEOP.TargetAutodiscoverEpr) {
        Write-Host -foregroundcolor Green "  TargetAutodiscoverEpr is" $exoOrgRel.TargetAutodiscoverEpr
    }
    else {
        Write-Host -foregroundcolor Red "  TargetAutodiscoverEpr is not" $fedinfoEOP.TargetAutodiscoverEpr
        #$countOrgRelIssues++
    }
    #Enabled
    Write-Host  " Enabled:"
    if ($exoOrgRel.enabled -like "True" ) {
        Write-Host -foregroundcolor Green "  Enabled is set to True"
    }
    else {
        Write-Host -foregroundcolor Red "  Enabled is set to False."
    }
}

Function EXOFedOrgIdCheck {
    Write-Host -foregroundcolor Green " Get-FederatedOrganizationIdentifier | select AccountNameSpace,Domains,Enabled"
    Write-Host $bar
    $exoFedOrgId = Get-FederatedOrganizationIdentifier | Select-Object AccountNameSpace, Domains, Enabled
    #$IntraOrgConCheck
    $efedorgid = $exoFedOrgId | Format-List
    $efedorgid
    Write-Host $bar
    Write-Host -foregroundcolor Green " SUMMARY - Online Federated Organization Identifier"
    Write-Host $bar
    Write-Host -foregroundcolor White " Domains: "
    if ($exoFedOrgId.Domains -like "*$ExchangeOnlineDomain*") {
        Write-Host -foregroundcolor Green " " $exoFedOrgId.Domains
    }
    else {
        Write-Host -foregroundcolor Red " Domains are NOT correct."
        Write-Host -foregroundcolor White " Should contain the $ExchangeOnlinemDomain"
    }
    Write-Host -foregroundcolor White " Enabled: "
    if ($exoFedOrgId.Enabled -like "True") {
        Write-Host -foregroundcolor Green "  True "
    }
    else {
        Write-Host -foregroundcolor Red "  Enabled is NOT correct."
        Write-Host -foregroundcolor White " Should be True"
    }
}

Function EXOTestOrgRelCheck {
    $exoIdentity = $ExoOrgRel.Identity
    Write-Host -foregroundcolor Green " Test-OrganizationRelationship -Identity $exoIdentity -UserIdentity $UserOnline"
    Write-Host $bar
    $exotestorgrel = Test-OrganizationRelationship -Identity $exoIdentity -UserIdentity $UserOnline
    $exotor = $exotestorgrel | Format-List
    $exotor
}

Function SharingPolicyCheck {
    Write-host $bar
    Write-Host -foregroundcolor Green " Get-SharingPolicy | select Domains,Enabled,Name,Identity"
    Write-Host $bar
    $Script:SPOnline = Get-SharingPolicy | Select-Object  Domains, Enabled, Name, Identity
    $SPOnline | Format-List
    Write-Host $bar
    Write-Host -foregroundcolor Green " SUMMARY - Sharing Policy"
    Write-Host $bar
    Write-Host -foregroundcolor White " Exchange On Premises Sharing domains:`n"
    #for
    Write-Host -foregroundcolor White "  Domain:"
    Write-Host "   " $SPOnprem.Domains.Domain[0]
    Write-Host -foregroundcolor White "  Action:"
    Write-Host "   " $SPOnprem.Domains.Actions[0]
    Write-Host -foregroundcolor White "  Domain:"
    Write-Host "   " $SPOnprem.Domains.Domain[1]
    Write-Host -foregroundcolor White "  Action:"
    Write-Host "   " $SPOnprem.Domains.Actions[1]
    Write-Host -ForegroundColor White "`n  Exchange Online Sharing Domains: `n"
    $domain1 = (($SPOnline.domains[0] -split ":") -split " ")
    $domain2 = (($SPOnline.domains[1] -split ":") -split " ")
    Write-Host -foregroundcolor White "  Domain:"
    Write-Host "   " $domain1[0]
    Write-Host -foregroundcolor White "  Action:"
    Write-Host "   " $domain1[1]
    Write-Host -foregroundcolor White "  Domain:"
    Write-Host "   " $domain2[0]
    Write-Host -foregroundcolor White "  Action:"
    Write-Host "   " $domain2[1]
    #Write-Host $bar
    if ((($domain1[0]) -eq ($SPOnprem.Domains.Domain[0]) -OR (($domain1[0]) -eq ($SPOnprem.Domains.Domain[1]))) -AND (($domain2[0]) -eq ($SPOnprem.Domains.Domain[0]) -OR (($domain2[0]) -eq ($SPOnprem.Domains.Domain[1]))) -AND (($domain1[1]) -eq ($SPOnprem.Domains.Actions[0]) -OR (($domain1[1]) -eq ($SPOnprem.Domains.Actions[1]))) -AND (($domain2[1]) -eq ($SPOnprem.Domains.Actions[0]) -OR (($domain1[1]) -eq ($SPOnprem.Domains.Actions[1])))  ) {
        Write-Host -foregroundcolor Green "`n  Exchange Online Sharing Domains match Exchange On Premise Sharing Policy Domain"
    }
    else {
        Write-Host -foregroundcolor Red "`n   Sharing Domains appear not to be correct."
        Write-Host -foregroundcolor White "   Exchange Online Sharing Domains appear not to match Exchange On Premise Sharing Policy Domains"
    }
    $bar
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
    Write-Host -foregroundcolor Green " SUMMARY - Online Intra Organization Connector"
    Write-Host $bar
    Write-Host -foregroundcolor White " Target Address Domains: "
    if ($exoIntraOrgCon.TargetAddressDomains -like "*$ExchangeOnpremDomain*") {
        Write-Host -foregroundcolor Green " " $exoIntraOrgCon.TargetAddressDomains
    }
    else {
        Write-Host -foregroundcolor Red " Target Address Domains is NOT correct."
        Write-Host -foregroundcolor White " Should contain the $ExchangeOnpremDomain"
    }
    Write-Host -foregroundcolor White " DiscoveryEndpoint: "
    if ($exoIntraOrgCon.DiscoveryEndpoint -like $EDiscoveryEndpoint.OnPremiseDiscoveryEndpoint) {
        Write-Host -foregroundcolor Green $exoIntraOrgCon.DiscoveryEndpoint
    }
    else {
        Write-Host -foregroundcolor Red " DiscoveryEndpoint is NOT correct. "
        Write-Host -foregroundcolor White "  Should be " $EDiscoveryEndpoint.OnPremiseDiscoveryEndpoint
    }
    Write-Host -foregroundcolor White " Enabled: "
    if ($exoIntraOrgCon.Enabled -like "True") {
        Write-Host -foregroundcolor Green "  True "
    }
    else {
        Write-Host -foregroundcolor Red "  Enabled is NOT correct."
        Write-Host -foregroundcolor White " Should be True"
    }
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
    Write-Host -foregroundcolor Green " SUMMARY - Online Intra Organization Configuration"
    Write-Host $bar
    Write-Host -foregroundcolor White " OnPremiseTargetAddresses: "
    if ($exoIntraOrgConfig.OnPremiseTargetAddresses -like "*$ExchangeOnpremDomain*") {
        Write-Host -foregroundcolor Green " " $exoIntraOrgConfig.OnPremiseTargetAddresses
    }
    else {
        Write-Host -foregroundcolor Red " OnPremise Target Addressess are NOT correct."
        Write-Host -foregroundcolor White " Should contain the $ExchangeOnpremDomain"
    }
}

Function EXOauthservercheck {
    Write-Host -foregroundcolor Green " Get-AuthServer -Identity 00000001-0000-0000-c000-000000000000 | select name,issueridentifier,enabled"
    Write-Host $bar
    $exoauthserver = Get-AuthServer -Identity 00000001-0000-0000-c000-000000000000 | Select-Object name, issueridentifier, enabled
    #$IntraOrgConCheck
    $authserver = $exoauthserver | Format-List
    $authserver
    Write-Host $bar
    Write-Host -foregroundcolor Green " SUMMARY - Exchange Online Authorization Server"
    Write-Host $bar
    Write-Host -foregroundcolor White " IssuerIdentifier: "
    if ($exoauthserver.IssuerIdentifier -like "00000001-0000-0000-c000-000000000000") {
        Write-Host -foregroundcolor Green " " $exoauthserver.IssuerIdentifier
    }
    else {
        Write-Host -foregroundcolor Red " Authorization Server object is NOT correct."
        Write-Host -foregroundcolor White " Enabled: "
        if ($exoauthserver.Enabled -like "True") {
            Write-Host -foregroundcolor Green "  True "
        }
        else {
            Write-Host -foregroundcolor Red "  Enabled is NOT correct."
            Write-Host -foregroundcolor White " Should be True"
        }
    }
}

Function EXOtestoauthcheck {
    Write-Host -foregroundcolor Green " Test-OAuthConnectivity -Service EWS -TargetUri $Global:ExchangeOnPremEWS -Mailbox $useronline "
    Write-Host $bar
    $exotestoauth = Test-OAuthConnectivity -Service EWS -TargetUri $Global:ExchangeOnPremEWS -Mailbox $useronline
    if ($exotestoauth.ResultType -eq 'Sucess' ) {
        $exotestoauth.ResultType
        } else {
        $exotestoauth
    }
    #$exoOAC = $exotestoauth | Format-List
    #$exoOAC
    #$bar
    #$exotestoauth.Detail.FullId
    #$bar
    if ($exotestoauth.Detail.FullId -like '*(401) Unauthorized*') {
        write-host -ForegroundColor Red "The remote server returned an error: (401) Unauthorized"
        if ($exotestoauth.Detail.FullId -like '*The user specified by the user-context in the token does not exist*') {
            write-host -ForegroundColor Yellow "The user specified by the user-context in the token does not exist"
            write-host "Please run Test-OAuthConnectivity with a different Exchange Online Mailbox"
        }
        if ($exotestoauth.Detail.FullId -like '*error_category="invalid_token"*') {
            write-host -ForegroundColor Yellow "This token profile 'S2SAppActAs' is not applicable for the current protocol"
        }
        Write-Host $bar
        #$OAuthConnectivity.detail.LocalizedString
        Write-Host -foregroundcolor Green " SUMMARY - Test OAuth COnnectivity"
        Write-Host $bar
        if ($OAuthConnectivity.ResultType -like "Success") {
            Write-Host -foregroundcolor Green " OAuth Test was completed successfully "
        }
        else {
            Write-Host -foregroundcolor Red " OAuth Test was completed with Error. "
            Write-Host -foregroundcolor White " Please rerun Test-OAuthConnectivity -Service EWS -TargetUri <EWS target URI> -Mailbox <On Premises Mailbox> | fl to confirm the test failure"
        }
    }
    Write-Host -foregroundcolor Green " Note:"
    Write-Host -foregroundcolor Yellow " You can ignore the warning 'The SMTP address has no mailbox associated with it'"
    Write-Host -foregroundcolor Yellow " when the Test-OAuthConnectivity returns a Success"
    Write-Host -foregroundcolor Green " Reference: "
    Write-Host -foregroundcolor White " Configure OAuth authentication between Exchange and Exchange Online organizations"
    Write-Host -foregroundcolor Yellow " https://technet.microsoft.com/en-us/library/dn594521(v=exchg.150).aspx"
}

#endregion
#cls
ShowParameters
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
$IntraOrgCon = Get-IntraOrganizationConnector -WarningAction SilentlyContinue -ErrorAction SilentlyContinue | Select-Object Name, TargetAddressDomains, DiscoveryEndpoint, Enabled
$EDiscoveryEndpoint = Get-IntraOrganizationConfiguration -WarningAction SilentlyContinue -ErrorAction SilentlyContinue | Select-Object OnPremiseDiscoveryEndpoint
$SPDomainsOnprem = Get-SharingPolicy -WarningAction SilentlyContinue -ErrorAction SilentlyContinue | Format-List Domains
$SPOnprem = Get-SharingPolicy  -WarningAction SilentlyContinue -ErrorAction SilentlyContinue | Select-Object *
#if($Auth -like "DAuth" -and $IntraOrgCon.enabled -Like "True")
#{
    #Write-Host $bar
    #Write-Host -foregroundcolor yellow "  Warning: Intra Organization Connector is Enabled -> Free Busy Lookup is done using OAuth"
    #Write-Host $bar
#}
if ($Org -like "EOP" -OR [string]::IsNullOrWhitespace($Organization)) {
    #region DAutch Checks
    if ($Auth -like "dauth" -OR [string]::IsNullOrWhitespace($Auth)) {
        Write-Host -foregroundcolor Green " `n `n ************************************TestingDAuth configuration************************************************* `n `n "
        OrgRelCheck
        Write-Host $bar
        if ($pause -eq "True") {
            $RH = Read-Host " Press Enter when ready to check the Federation Information Details."
            Write-Host $bar
            #$pause = "True"
            #$pause
        }
        FedInfoCheck
        #$pause
        if ($pause -eq "True") {
            $RH = Read-Host " Press Enter when ready to check the Federation Trust configuration details. "
            Write-Host $bar
            #$pause = "True"
        }
        FedTrustCheck
        Write-Host $bar
        Write-Host -foregroundcolor Green " Test-FederationTrustCertificate"
        Write-Host $bar
        Test-FederationTrustCertificate
        Write-Host $bar
        if ($pause -eq "True") {
            $RH = Read-Host " Press Enter when ready to check the On-Prem Autodiscover Virtual Directory configuration details. "
            Write-Host $bar
            #$pause = "True"
        }
        AutoDVirtualDCheck
        Write-Host $bar
        if ($pause -eq "True") {
            $RH = Read-Host " Press Enter when ready to check the On-Prem Web Services Virtual Directory configuration details. "
            Write-Host $bar
            #$pause = "True"
        }
        EWSVirtualDirectoryCheck
        if ($pause -eq "True") {
            Write-Host $bar
            $RH = Read-Host " Press Enter when ready to  check the Availability Address Space configuration details. "
            #Write-Host $bar
            #$pause = "True"
        }
        AvailabilityAddressSpaceCheck
        if ($pause -eq "True") {
            Write-Host $bar
            $RH = Read-Host " Press Enter when ready to test the Federation Trust. "
            #Write-Host $bar
            #$pause = "True"
        }
        #need to grab errors and provide alerts in error case
        TestFedTrust
        if ($pause -eq "True") {
            Write-Host $bar
            $RH = Read-Host " Press Enter when ready to Test the Organization Relationship. "
            #Write-Host $bar
            #$pause = "True"
        }
        TestOrgRel
    }
    #endregion
    #region OAuth Check
    if ($Auth -like "OAuth" -OR [string]::IsNullOrWhitespace($Auth)) {
        if ($pause -eq "True") {
            $RH = Read-Host " Press Enter when ready to check the OAuth configuration details. "
            Write-Host $bar
            #$pause = "True"
        }
        Write-Host -foregroundcolor Green " `n `n ************************************TestingOAuth configuration************************************************* `n `n "
        Write-Host $bar
        IntraOrgConCheck
        Write-Host $bar
        if ($pause -eq "True") {
            $RH = Read-Host " Press Enter when ready to check the Auth Server configuration details. "
            Write-Host $bar
            #$pause = "True"
        }
        AuthServerCheck
        Write-Host $bar
        if ($pause -eq "True") {
            $RH = Read-Host " Press Enter when ready to check the Partner Application configuration details. "
            #Write-Host $bar
            #$pause = "True"
        }
        PartnerApplicationCheck
        Write-Host $bar
        if ($pause -eq "True") {
            $RH = Read-Host " Press Enter when ready to check the Exchange Online-ApplicationAccount configuration details. "
            #Write-Host $bar
            #$pause = "True"
        }
        ApplicationAccounCheck
        Write-Host $bar
        if ($pause -eq "True") {
            $RH = Read-Host " Press Enter when ready to check the Management Role Assignments for the Exchange Online-ApplicationAccount. "
            Write-Host $bar
            #$pause = "True"
        }
        ManagementRoleAssignmentCheck
        Write-Host $bar
        if ($pause -eq "True") {
            $RH = Read-Host " Press Enter when ready to check Auth configuration details. "
            Write-Host $bar
            #$pause = "True"
        }
        AuthConfigCheck
        Write-Host $bar
        if ($pause -eq "True") {
            $RH = Read-Host " Press Enter when ready to check the Auth Certificate configuration details. "
            Write-Host $bar
            #$pause = "True"
        }
        CurrentCertificateThumbprintCheck
        Write-Host $bar
        if ($pause -eq "True") {
            $RH = Read-Host " Press Enter when ready to  check the On Prem Autodiscover Virtual Directory configuration details. "
            Write-Host $bar
            #$pause = "True"
        }
        AutoDVirtualDCheckOAuth
        $AutoDiscoveryVirtualDirectoryOAuth
        Write-Host $bar
        if ($pause -eq "True") {
            $RH = Read-Host " Press Enter when ready to check the On-Prem Web Services Virtual Directory configuration details. "
            Write-Host $bar
            #$pause = "True"
        }
        EWSVirtualDirectoryCheckOAuth
        Write-Host $bar
        if ($pause -eq "True") {
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
if ($Org -like "EOL" -OR [string]::IsNullOrWhitespace($Organization)) {
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
    #endregion
    #region ExoDauthCheck
    if ($Auth -like "dauth" -OR [string]::IsNullOrWhitespace($Auth)) {
        Write-Host $bar
        Write-Host -foregroundcolor Green " `n `n ************************************Testing DAuth configuration************************************************* `n `n "
        ExoOrgRelCheck
        Write-Host $bar
        if ($pause -eq "True") {
            $RH = Read-Host " Press Enter when ready to check the Federation Organization Identifier configuration details. "
            Write-Host $bar
        }
        EXOFedOrgIdCheck
        Write-Host $bar
        if ($pause -eq "True") {
            $RH = Read-Host " Press Enter when ready to check the Organization Relationship configuration details. "
            Write-Host $bar
        }
        EXOTestOrgRelCheck
        if ($pause -eq "True") {
            Write-Host $bar
            $RH = Read-Host " Press Enter when ready to check the Sharing Policy configuration details. "
        }
        SharingPolicyCheck
    }
    #endregion
    #region ExoOauthCheck
    if ($Auth -like "oauth" -OR [string]::IsNullOrWhitespace($Auth)) {
        Write-Host -foregroundcolor Green " `n `n ************************************Testing OAuth configuration************************************************* `n `n "
        Write-Host $bar
        ExoIntraOrgConCheck
        Write-Host $bar
        if ($pause -eq "True") {
            $RH = Read-Host " Press Enter when ready to check the Organizationconfiguration details. "
            Write-Host $bar
        }
        EXOIntraOrgConfigCheck
        Write-Host $bar
        if ($pause -eq "True") {
            $RH = Read-Host " Press Enter when ready to check the Authentication Server Authorization Details.  "
            Write-Host $bar
        }
        EXOauthservercheck
        Write-Host $bar
        if ($pause -eq "True") {
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
