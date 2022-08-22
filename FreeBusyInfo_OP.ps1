#Exchange on Premise

#>


#region Properties and Parameters
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', '', Justification = 'Variables are being used')]
[CmdletBinding(DefaultParameterSetName = "FreeBusyInfo_OP", SupportsShouldProcess)]
param(
    [Parameter(Mandatory = $false, ParameterSetName = "Auth")]
    [string]$Auth,
    [Parameter(Mandatory = $false, ParameterSetName = "ConfigurationOnly")]
    [switch]$ConfigurationOnly
)


Set-ExecutionPolicy " Unrestricted"  -Scope Process -Confirm:$false
Set-ExecutionPolicy " Unrestricted"  -Scope CurrentUser -Confirm:$false
Add-PSSnapin microsoft.exchange.management.powershell.snapin
import-module ActiveDirectory 
cls
$countOrgRelIssues = "0"
$FedTrust = $null
$Global:AutoDiscoveryVirtualDirectory = $null
$Global:OrgRel
$AvailabilityAddressSpace = $null
$Global:WebServicesVirtualDirectory = $null
$bar = " ==================================================================================================================" 
$barspace = " ================================================================================================================== `n " 
$spacebar = " `n ==================================================================================================================" 
$spacebarspace = " `n ================================================================================================================== `n "
$logfile = "$PSScriptRoot\FreeBusyInfo_OP.txt" 
$Logfile = [System.IO.Path]::GetFileNameWithoutExtension($logfile) + "_"  + `
        (get-date -format yyyyMMdd_HHmmss) + ([System.IO.Path]::GetExtension($logfile))
Write-Host " `n`n " 
Start-Transcript -path $LogFile -append
Write-Host -foregroundcolor Green " `n`n Free Busy Configuration Information Grabber `n "
Write-Host -foregroundcolor White " Version -1 `n " 
Write-Host -foregroundcolor Green " Loading Parameters..... `n "

#Parameter input

$UserOnline = get-remotemailbox -resultsize 1 -WarningAction SilentlyContinue
$UserOnline = $UserOnline.RemoteroutingAddress.smtpaddress 
$ExchangeOnlineDomain = ($UserOnline -split "@")[1]

if ($ExchangeOnlineDomain -like "*mail.onmicrosoft.com"){
$ExchangeOnlineAltDomain = (($ExchangeOnlineDomain.Split(".")))[0]+".onmicrosoft.com"
}
else{
$ExchangeOnlineAltDomain = (($ExchangeOnlineDomain.Split(".")))[0]+".mail.onmicrosoft.com"
}

$UserOnPrem = get-mailbox -resultsize 1 -WarningAction SilentlyContinue | Where{($_.EmailAddresses -like "*"+$ExchangeOnlineDomain )}
$UserOnPrem = $UserOnPrem.PrimarySmtpAddress.Address
$ExchangeOnPremDomain = ($UserOnPrem -split "@")[1]
$EWSVirtualDirectory = Get-WebServicesVirtualDirectory
$ExchangeOnPremEWS = ($EWSVirtualDirectory.externalURL.AbsoluteUri)[0]
$ADDomain=Get-ADDomain
$ExchangeOnPremLocalDomain=$ADDomain.forest
if ([string]::IsNullOrWhitespace($ADDomain)){
$ExchangeOnPremLocalDomain = $exchangeOnPremDomain

}

#endregion

Function Get-Sumary {


Write-Host -foregroundcolor Green " `n Sumary - Free Busy Configuration GLobal View (non standard values will show up in Red. Standard Values in Green)" 
Write-Host $spacebar
  if ($IntraOrgConEnabled.enabled -Like "True" )
    {
    Write-Host -foregroundcolor Green " Intra Organization Connector is Enabled" 
    if ($OrgRel.enabled -like "True" )
        {
            Write-Host -foregroundcolor Green " Organization Relationship is Enabled for Hybrid Use. `n " 
            Write-Host -foregroundcolor White " Intra Organization Connector takes precedence over Organization Relationship. `n "  
            Write-Host -foregroundcolor Green " - Free Busy Lookup Configured Methods:"  
            Write-Host -foregroundcolor White " 
            - Intra Organization Connector
            - Authentication Method -> oAuth
            - Intra Organization Connector is Configured/Enabled for Hybrid use.
            - Organization Relationship is Enabled for Hybrid Use. `n " 
           
            Write-Host -foregroundcolor Green " => Free Busy Lookup From On Premise to Exchange Online is done using Intra Org Connector `n" 
        }
    }
else
    {
    Write-Host " Intra Organization Connector is NOT Enabled." 
    if ($OrgRel.enabled -like "True" )
        {
            Write-Host -foregroundcolor Green " `n Organization Relationship is Enabled for Hybrid Use." 
            Write-Host -foregroundcolor Green " - Free Busy Lookup Method:"  
            Write-Host -foregroundcolor White " 
            - Organization Relationship
            - Authentication Method -> dAuth
            - Intra Organization Connector is NOT Configured/Enabled for Hybrid use.
            - Organization Relationship is Enabled for Hybrid Use. `n "   
            Write-Host -foregroundcolor Green " => Free Busy Lookup From On Premise to Exchange Online is done using Organization Relationship `n" 
        }
        else
            {
            Write-Host -foregroundcolor Red " 
            - Organization Relationship is NOT Enabled or correctly configured for Hybrid Use.
            - Intra Org Connector is NOT Enabled or configured for Hybrid use. `n`n "  
            Write-Host -foregroundcolor Red " => Free Busy Lookup From On Premise to Exchange Online is NOT correctly Configured for Hybrid Lookup `n" 
            }
    
    }
  Write-Host $barspace
}


#region Edit Parameters

Function UserOnlineCheck{
Write-Host -foregroundcolor Green " `n`n Online Mailbox: $UserOnline `n`n " 
$UserOnlineCheck = Read-Host " Press the Enter key if OK or type an Exchange Online Email address and press the Enter key `n "
if (![string]::IsNullOrWhitespace($UserOnlineCheck))
{
    $script:UserOnline = $UserOnlineCheck
} 
}

Function ExchangeOnlineDomainCheck{

#$ExchangeOnlineDomain
Write-Host -foregroundcolor Green " `n`n Exchange Online Domain: $ExchangeOnlineDomain `n`n " 
$ExchangeOnlineDomaincheck = Read-Host " Press enter if OK or type in the Exchange Online Domain and press the Enter key."

if (![string]::IsNullOrWhitespace($ExchangeOnlineDomaincheck))
{
    $script:ExchangeOnlineDomain = $ExchangeOnlineDomainCheck
} 
}

Function UseronpremCheck {
Write-Host -foregroundcolor Green " `n`n On Premises Hybrid Mailbox: $Useronprem `n`n " 
$Useronpremcheck = Read-Host " Press Enter if OK or type in an Exchange OnPremises Hybrid email address and press the Enter key."

if (![string]::IsNullOrWhitespace($Useronpremcheck))
{
    $script:Useronprem = $Useronpremcheck
} 
}

Function ExchangeOnPremDomainCheck {
#$exchangeOnPremDomain
Write-Host -foregroundcolor Green " `n`n On Premises Mail Domain: $exchangeOnPremDomain `n`n" 
$exchangeOnPremDomaincheck = Read-Host " Press enter if OK or type in the Exchange On Premises Mail Domain and press the Enter key."

if (![string]::IsNullOrWhitespace($exchangeOnPremDomaincheck))
{
    $script:exchangeOnPremDomain = $exchangeOnPremDomaincheck
} 

}

Function ExchangeOnPremEWSCheck{

Write-Host -foregroundcolor Green " `n`n On Premises EWS External URL: $exchangeOnPremEWS `n`n " 

$exchangeOnPremEWScheck = Read-Host " Press enter if OK or type in the Exchange On Premises EWS URL and press the Enter key."

if (![string]::IsNullOrWhitespace($exchangeOnPremEWScheck))
{
   $exchangeOnPremEWS = $exchangeOnPremEWScheck
} 
}

Function ExchangeOnPremLocalDomainCheck{
Write-Host -foregroundcolor Green " `n`n On Premises Root Domain: $exchangeOnPremLocalDomain `n`n " 

if ([string]::IsNullOrWhitespace($exchangeOnPremLocalDomain)){
$exchangeOnPremLocalDomaincheck = Read-Host "Please type in the Active directory Root domain.
Press Enter to use $exchangeOnPremDomain" 
if ([string]::IsNullOrWhitespace($ADDomain)){
$exchangeOnPremLocalDomain = $exchangeOnPremDomain
}
if ([string]::IsNullOrWhitespace($exchangeOnPremLocalDomain)){
$exchangeOnPremLocalDomain = $exchangeOnPremLocalDomaincheck
}


}
}
#endregion

#region Show Parameters

Function ShowParameters{
Write-Host $bar
$doublespace
Write-Host -foregroundcolor Green "          Exchange On Premise to Office365 Free Busy Configuration Information Graber" 
$doublespace 
Write-Host $bar
write-host -foregroundcolor Green " Loading modules for AD, Exchange" 

Write-Host $bar
Write-Host   "  Color Scheme"
Write-Host $barspace
Write-Host -ForegroundColor Red "  Look out for Red!"
Write-Host -ForegroundColor Yellow "  Yellow - Example information or Links"
Write-Host -ForegroundColor Green "  Green - In Sumary Sections it means OK. Anywhere else it's just a visual aid."
Write-Host $spacebar
Write-Host   "  Parameters:"
Write-Host $barspace
Write-Host " Log File Path:" 

Write-Host -foregroundcolor Green "  $PSScriptRoot\$Logfile `n "
Write-Host " Office 365 Domain:"
Write-Host -foregroundcolor Green "  $exchangeOnlineDomain `n "
Write-Host " AD root Domain"
Write-Host -foregroundcolor Green "  $exchangeOnPremLocalDomain `n "
Write-Host -foregroundcolor White " Exchange On Premises Domain:  "
Write-Host -foregroundcolor Green "  $exchangeOnPremDomain `n "
Write-Host " Exchange On Premises External EWS url:"
Write-Host -foregroundcolor Green "  $exchangeOnPremEWS `n "
Write-Host " On Premises Hybrid Mailboxr:"
Write-Host -foregroundcolor Green "  $useronprem `n "
Write-Host " Exchange Online Mailbox:"
Write-Host -foregroundcolor Green "  $userOnline `n "
}

Function Update-HealthChecker {

Write-Verbose "Updating HealthChecker"
.\HealthChecker.ps1 -ScriptUpdateOnly

}
#endregion

#regionDAuth Functions

Function OrgRelSTV{


Write-Host $spacebar
Write-Host -foregroundcolor Green "  Get-OrganizationRelationship  CONFIGURATION EXAMPLE" 
Write-Host $barspace
Write-Host -foregroundcolor Yellow "  Enabled               : True" 
Write-Host -ForegroundColor Yellow "  DomainNames           : {contoso.com, contoso.mail.onmicrosoft.com}" 
Write-Host -foregroundcolor Yellow "  FreeBusyAccessEnabled : True" 
Write-Host -foregroundcolor Yellow "  FreeBusyAccessLevel   : LimitedDetails" 
Write-Host -foregroundcolor Yellow "  FreeBusyAccessScope   :" 
Write-Host -foregroundcolor Yellow "  TargetApplicationUri  : outlook.com" 
Write-Host -foregroundcolor Yellow "  TargetSharingEpr      :" 
Write-Host -foregroundcolor Yellow "  TargetOwaURL          : http://outlook.com/owa/contoso.com" 
Write-Host -foregroundcolor Yellow "  TargetAutodiscoverEpr : https://autodiscover-s.outlook.com/autodiscover/autodiscover.svc/WSSecurity" 
Write-Host $spacebar
Write-Host -foregroundcolor Green "  Standard Configuration Values"  
Write-Host $barspace
Write-Host -foregroundcolor Green "  DomainNames `n "  
Write-Host -foregroundcolor White " - DomainNames should represent Exchange Online email Domain Name. `n "  
Write-Host -foregroundcolor Yellow  "  Example: contoso.mail.onmicrosoft.com `n " 
Write-Host -foregroundcolor Green "  TargetAutodiscoverEpr `n "
Write-Host -foregroundcolor White " - TargetAutodiscoverEpr should be the Exchange Online autodiscover endpoint. `n` " 
Write-Host -foregroundcolor Yellow "  Example: https://autodiscover-s.outlook.com/autodiscover/autodiscover.svc/WSSecurity `n "  
Write-Host -foregroundcolor White " - To find the value from TargetAutodiscoverEpr the follwoing command gives the necessary output: `n " 
Write-Host -foregroundcolor Yellow "  Get-FederationInformation -DomainName contoso.mail.onmicrosoft.com  -BypassAdditionalDomainValidation | fl `n " 
Write-Host -foregroundcolor Green "  TargetSharingEPR `n " 
Write-Host -foregroundcolor White "  - TargetSharingEPR ideally is left blank. If it is set, it should be Office 365 EWS endpoint. `n " 
Write-Host -foregroundcolor Yellow "  Example: https://outlook.office365.com/EWS/Exchange.asmx `n " 
Write-Host -foregroundcolor Green "  TargetApplicationURI `n " 
Write-Host -foregroundcolor White "  - TargetApplicationURI should be outlook.com If this is for organization relationship of the cloud tenant. "  
Write-Host -foregroundcolor White "  - For non-cloud organization relationship it must match (Get-FederationTrust).ApplicationUri of the 
 On-Premise Trust of the Trusted Organization. `n " 
Write-Host -foregroundcolor Green "  FreeBusyAccessEnabled `n "
Write-Host -foregroundcolor White " - FreeBusyAccessEnabled must be set to True. `n " 
Write-Host -foregroundcolor Green "  FreeBusyAccessLevel `n " 
Write-Host -foregroundcolor White " - FreeBusyAccessLevel should be either AvailabilityOnly or LimitedDetails. 
 AvailabilityOnly:Free/busy access with time only LimitedDetails:Free/busy access with time, 
 subject, and location For more information, see curated:FreeBusyAccessLevel and MailboxFolderPermission `n " 
Write-Host -foregroundcolor Green "  FreeBusyAccessScope `n " 
Write-Host -foregroundcolor White " - FreeBusyAccessScope is typically blank. 
 The FreeBusyAccessScope parameter specifies a security distribution group in the internal organization that contains users that can have their free/busy information accessed by an external organization. `n " 
Write-Host -foregroundcolor Green "  ArchiveAccessEnabled `n " 
Write-Host -foregroundcolor White "  - ArchiveAccessEnable must be True. `n " 
Write-Host -foregroundcolor Green "  Enabled `n " 
Write-Host -foregroundcolor White " - Enabled must be True. `n " 
Write-Host -foregroundcolor Green " - To find the value from TargetAutodiscoverEpr the follwoing command gives the necessary output: `n "  
Write-Host -foregroundcolor Yellow "  Get-FederationInformation -DomainName tenantname.mail.onmicrosoft.com  -BypassAdditionalDomainValidation | fl `n " 



}

Function OrgRelCheck (){

Write-Host -foregroundcolor Green "  ************** Testing Free Busy Lookup with Organization Relationship - DAuth (WITHOUT OAuth)***************** `n " 
Write-Host $barspace
Write-Host -foregroundcolor Green "  Checking Organization Relationship Configuration `n " 
Write-Host -foregroundcolor Green "  Get-OrganizationRelationship  | Where{($_.DomainNames -like $ExchangeOnlineDomain )} | Select Identity,DomainNames,FreeBusy*,Target*,Enabled" 
Write-Host $bar
$OrgRel

Write-Host $bar
Write-Host  -foregroundcolor Green " Sumary - Organization Relationship (non standard values will show up in Red. Standard Values in Green)" 
Write-Host $barspace

#$exchangeonlinedomain
Write-Host  " Domains:" 
if ($orgrel.DomainNames -like $exchangeonlinedomain){
Write-Host -foregroundcolor Green "  Domain Names Include the $exchangeOnlineDomain Domain `n " 
}
else
{
Write-Host -foregroundcolor Red "  Domain Names do Not Include the $exchangeOnlineDomain Domain `n " 

}


#FreeBusyAccessEnabled
Write-Host  " FreeBusyAccessEnabled:" 
if ($OrgRel.FreeBusyAccessEnabled -like "True" ){
Write-Host -foregroundcolor Green "  FreeBusyAccessEnabled is set to True `n " 

}
else
{
Write-Host -foregroundcolor Red "  FreeBusyAccessEnabled : False `n " 
$countOrgRelIssues++
}



#FreeBusyAccessLevel
Write-Host  " FreeBusyAccessLevel:" 
if ($OrgRel.FreeBusyAccessLevel -like "AvailabilityOnly" ){
Write-Host -foregroundcolor Green "  FreeBusyAccessLevel is set to AvailabilityOnly `n " 
}

if ($OrgRel.FreeBusyAccessLevel -like "LimitedDetails" ){
Write-Host -foregroundcolor Green "  FreeBusyAccessLevel is set to LimitedDetails `n " 
}

else
{
Write-Host -foregroundcolor Red "  FreeBusyAccessEnabled : False `n " 
$countOrgRelIssues++
}



#TargetApplicationUri
Write-Host  " TargetApplicationUri:" 
 if ($OrgRel.TargetApplicationUri -like "Outlook.com" ){
Write-Host -foregroundcolor Green "  TargetApplicationUri is Outlook.com `n " 
}
else
{
Write-Host -foregroundcolor Red "  TargetApplicationUri IS NOT Outlook.com `n " 
$countOrgRelIssues++

}


#TargetOwaURL
Write-Host  " TargetOwaURL:" 
if ($OrgRel.TargetOwaURL -like "http://outlook.com/owa/$exchangeonpremdomain"){
Write-Host -foregroundcolor Green "  TargetOwaURL is http://outlook.com/owa/$exchangeonpremdomain `n " 
}
else
{
Write-Host -foregroundcolor Red "  TargetOwaURL IS NOT http://outlook.com/owa/$exchangeonpremdomain `n " 
$countOrgRelIssues++

}




#TargetSharingEpr
Write-Host  " TargetSharingEpr:" 
if ([string]::IsNullOrWhitespace($OrgRel.TargetSharingEpr)){
Write-Host -foregroundcolor Green "  TargetSharingEpr is blank. this is the standard Value.`n " 
}
else
{
Write-Host -foregroundcolor Red "  TargetSharingEpr is NOT blank. 
If it is set, it should be Office 365 EWS endpoint `n " 
$countOrgRelIssues++

}



#TargetAutodiscoverEpr:
Write-Host  " TargetAutodiscoverEpr:" 
if ($OrgRel.TargetAutodiscoverEpr -like "https://autodiscover-s.outlook.com/autodiscover/autodiscover.svc/WSSecurity" ){
Write-Host -foregroundcolor Green "  TargetAutodiscoverEpr is correct `n " 
}
else
{
Write-Host -foregroundcolor Red "  TargetAutodiscoverEpr is not correct `n " 
$countOrgRelIssues++

}


#Enabled

Write-Host  " Enabled:" 
if ($OrgRel.enabled -like "True" ){
Write-Host -foregroundcolor Green "  Enabled is set to True" 

}
else
{
Write-Host -foregroundcolor Red "  Enabled is set to False." 
$countOrgRelIssues++
}
$doublespace

if ($countOrgRelIssues -eq '0'){
Write-Host -foregroundcolor Green "   `n `n      Configurations Seem Correct" 
}
else
{
Write-Host -foregroundcolor Red "   `n `n      Configurations DO NOT Seem Correct" 
}

}

Function FedInfoSTV{


Write-Host $spacebar
Write-Host -foregroundcolor Green "  Get-FederationInformation  CONFIGURATION EXAMPLE" 
Write-Host $barspace
Write-Host -foregroundcolor Yellow "  TargetApplicationUri  : Outlook.com" 
Write-Host -ForegroundColor Yellow "  DomainNames           : {contoso.com, contoso.mail.onmicrosoft.com, contoso.nail.onmicrosoft.com}" 
Write-Host -foregroundcolor Yellow "  TargetAutodiscoverEpr : https://autodiscover-s.outlook.com/autodiscover/autodiscover.svc/WSSecurity" 
Write-Host -foregroundcolor Yellow "  TokenIssuerUris       : {urn:federation:MicrosoftOnline}" 
Write-Host $spacebar
Write-Host -foregroundcolor Green "  Standard Configuration Values"  
Write-Host $barspace
Write-Host -foregroundcolor Green "  TargetApplicationURI `n " 
Write-Host -foregroundcolor White "  - TargetApplicationURI should be Outlook.com. `n "  
Write-Host -foregroundcolor Green "  DomainNames `n "  
Write-Host -foregroundcolor White " - DomainNames should contain the Exchange Online .onmicrosoft.com Domains. `n "  
Write-Host -foregroundcolor Yellow  "  Example: contoso.mail.onmicrosoft.com, contoso.mail.onmicrosoft.com, contoso.com `n " 
Write-Host -foregroundcolor Green "  TargetAutodiscoverEpr `n "
Write-Host -foregroundcolor White " - TargetAutodiscoverEpr should be the Exchange Online autodiscover endpoint. Should also match the Organizatin Relationship TargetAutodiscoverEpr.`n` " 
Write-Host -foregroundcolor Yellow "  Example: https://autodiscover-s.outlook.com/autodiscover/autodiscover.svc/WSSecurity `n "  
#test if they match
$spacebar


}

Function FedInfoCheck{

Write-Host -foregroundcolor Green " Checking FederationInformation for the Exchange Online Domain `n" 
Write-Host -foregroundcolor Green " Get-FederationInformation -DomainName $exchangeOnlineDomain  -BypassAdditionalDomainValidation | fl" 
Write-Host $spacebarspace
$fedinfo = get-federationInformation -DomainName $exchangeOnlineDomain  -BypassAdditionalDomainValidation | select *
$fedinfo

Write-Host $spacebar

Write-Host -foregroundcolor Green " SUMARY - Federation Information (non standard values will show up in Red. Standard Values in Green)" 
Write-Host $barspace


#TargetApplicationUri
if ($fedinfo.TargetApplicationUri -eq "outlook.com"){
Write-Host  " TargetApplicationUri: "
Write-Host -foregroundcolor Green " "$fedinfo.TargetApplicationUri
}
else{
Write-Host " TargetApplicationUri: "
Write-Host -foregroundcolor Red $fedinfo.TargetApplicationUri
Write-Host -foregroundcolor Red   " TargetApplicationUri should be Outlook.com"
}

#DomainNames
if ($fedinfo.DomainNames -like "*$ExchangeOnlineDomain*"){
Write-Host  "`n Domain Names: "
#if not null
Write-Host -foregroundcolor Green " "$fedinfo.DomainNames
}
else{
Write-Host " n Domain Names: "
Write-Host -foregroundcolor Red " "$fedinfo.DomainNames
Write-Host -foregroundcolor Red  " DomainNames should contain $ExchangeOnlineDomain"
}

#TargetAutodiscoverEpr
if ($OrgRel.TargetAutodiscoverEpr -like $fedinfo.TargetAutodiscoverEpr){

Write-Host  "`n TargetAutodiscoverEpr: "
Write-Host -foregroundcolor Green " "$fedinfo.TargetAutodiscoverEpr
Write-Host -foregroundcolor Green "`n Federation Information TargetAutodiscoverEpr matches the Organization Relationship TargetAutodiscoverEpr " 
}
else
{
Write-Host -foregroundcolor Red " `n =>  Federation Information TargetAutodiscoverEpr DOES NOT MATCH the Organization Relationship TargetAutodiscoverEpr`n " 

Write-Host " Organization Relationship:"  $OrgRel.TargetAutodiscoverEpr
Write-Host " `n Federation Information:   "  $fedinfo.TargetAutodiscoverEpr
}

#TokenIssuerUris
if ($fedinfo.TokenIssuerUris -like "*urn:federation:MicrosoftOnline*"){
Write-Host  " `n TokenIssuerUris: "
Write-Host -foregroundcolor Green " "  $fedinfo.TokenIssuerUris
}
else{
Write-Host "`n TokenIssuerUris: " $fedinfo.TokenIssuerUris
Write-Host  -foregroundcolor Red " `n TokenIssuerUris should be urn:federation:MicrosoftOnline"
}

Write-Host $spacebar
}

Function FedTrustSTV{

Write-Host $spacebar
Write-Host -foregroundcolor Green "  Standard Configuration Values "  
Write-Host $barspace
Write-Host -foregroundcolor Green "  ApplicationUri `n "  
Write-Host -foregroundcolor White " - This should be the unique identifier for the domain. Example: FYDIBOHF25SPDLT.contoso.com `n " 
Write-Host -foregroundcolor Green "  TokenIssuerUri `n " 
Write-Host -foregroundcolor White "  - This must be urn:federation:MicrosoftOnline. `n " 
Write-Host -foregroundcolor Green "  OrgCertificate `n " 
Write-Host -foregroundcolor White " - This is typically a self-signed certificate. Verify the certificate referenced has not expired," 
Write-Host -foregroundcolor White " - Use Test-FederationTrustCertificate to validate if the certificate exists on all Exchange 2010 
and 2013/2016 MBX and CAS servers.`n " 
Write-Host -foregroundcolor Green " TokenIssuerCertificate `n " 
Write-Host -foregroundcolor White "  - Verify the certificate has not expired. If expired, use Get-FederationTrust | Set-FederationTrust -RefreshMetadata `n " 
Write-Host -foregroundcolor Green " TokenIssuerPrevCertificate `n " 
Write-Host -foregroundcolor White "  - Verify the certificate has not expired. If expired, use Get-FederationTrust | Set-FederationTrust -RefreshMetadata `n " 
Write-Host -foregroundcolor Green " TokenIssuerMetadataEpr `n " 
Write-Host -foregroundcolor White " - The value is https://nexus.microsoftonline-p.com/federationmetadata/2006-12/federationmetadata.xml . 
Verify ALL Exchange 2010 and 2013/2016 MBX and `n "  
Write-Host -foregroundcolor White " CAS servers can access the Url. If the servers require outbound proxy server for internet access, 
verify InternetWebProxy is set from Get-ExchangeServer. `n "  
Write-Host -foregroundcolor White " The proxy cannot require authentication from the Exchange servers. If you paste the URL in IE, 
you should either see the metadata xml or be prompted to download"  
Write-Host -foregroundcolor White "  the xml file TokenIssuerMetadataEpr.png See curated:Outbound Access to MFG\AuthServer for more information. `n " 
Write-Host -foregroundcolor Green " TokenIssuerEpr `n " 
Write-Host -foregroundcolor White "  - The value is https://login.microsoftonline.com/extSTS.srf . Verify ALL Exchange 2010 and 2013 MBX and CAS servers can 
access the Url."  
Write-Host -foregroundcolor White "  If the servers require outbound proxy server for internet access, verify InternetWebProxy is set from Get-ExchangeServer. " 
Write-Host -foregroundcolor White "  The proxy cannot require authentication from the Exchange servers. If you paste the URL in IE, you should be prompted 
to download the file. `n " 
}

Function FedTrustCheck{
Write-Host -foregroundcolor Green " Get-FederationTrust | fl ApplicationUri,TokenIssuerUri,OrgCertificate,TokenIssuerCertificate,
TokenIssuerPrevCertificate, TokenIssuerMetadataEpr,TokenIssuerEpr" 
Write-Host $barspace
$script:fedtrust = Get-FederationTrust | select ApplicationUri,TokenIssuerUri,OrgCertificate,TokenIssuerCertificate,TokenIssuerPrevCertificate, TokenIssuerMetadataEpr,TokenIssuerEpr
$fedtrust
Write-Host $spacebar
Write-Host -foregroundcolor Green " `n SUMARY - Federation Trust (non standard values will show up in Red. Standard Values in Green)" 
Write-Host $barspace
$CurrentTime = get-date
Write-Host -foregroundcolor White " `n Federation Trust Aplication Uri:" 
if ($script:fedtrust.ApplicationUri -like  "FYDIBOHF25SPDLT.$ExchangeOnpremDomain"){

Write-Host -foregroundcolor Green " " $fedtrust.ApplicationUri

}
else
{ 
Write-Host -foregroundcolor Red " `n Federation Trust Aplication Uri is NOT correct. "
Write-Host -foregroundcolor White " Should be "$fedtrust.ApplicationUri

}

#$fedtrust.TokenIssuerUri.AbsoluteUri
Write-Host -foregroundcolor White " `n TokenIssuerUri:"
if ($fedtrust.TokenIssuerUri.AbsoluteUri -like  "urn:federation:MicrosoftOnline"){
#Write-Host -foregroundcolor White "  TokenIssuerUri:"
Write-Host -foregroundcolor Green " "$fedtrust.TokenIssuerUri.AbsoluteUri

}
else
{

Write-Host -foregroundcolor Red " `n  Federation Trust TokenIssuerUri is NOT urn:federation:MicrosoftOnline `n"

}
Write-Host -foregroundcolor White " Federation Trust Certificate Expiracy:"
if ($fedtrust.OrgCertificate.NotAfter.Date -gt $CurrentTime){

Write-Host -foregroundcolor Green "  Not Expired - Expires on " $fedtrust.OrgCertificate.NotAfter.DateTime

}
else
{

Write-Host -foregroundcolor Red "  `n  Federation Trust Certificate is Expired on " $fedtrust.OrgCertificate.NotAfter.DateTime

}
Write-Host -foregroundcolor White " `n  Federation Trust Token Issuer Certificate Expiracy:"
if ($fedtrust.OrgCertificate.NotAfter.Date -gt $CurrentTime){

Write-Host -foregroundcolor Green "  Federation Trust TokenIssuerCertificate Not Expired - Expires on " $fedtrust.TokenIssuerCertificate.NotAfter.DateTime

}
else
{
Write-Host -foregroundcolor Red "  Federation Trust TokenIssuerCertificate Expired on " $fedtrust.TokenIssuerCertificate.NotAfter.DateTime
}
Write-Host -foregroundcolor White " `n Federation Trust Token Issuer Prev Certificate Expiracy:"
if ($fedtrust.TokenIssuerPrevCertificate.NotAfter.Date -gt $CurrentTime){
Write-Host -foregroundcolor Green "  Federation Trust TokenIssuerPrevCertificate Not Expired - Expires on " $fedtrust.TokenIssuerPrevCertificate.NotAfter.DateTime 
}
else
{ 
Write-Host -foregroundcolor Red "  Federation Trust TokenIssuerPrevCertificate Expired on " $fedtrust.TokenIssuerPrevCertificate.NotAfter.DateTime

}
$fedtrustTokenIssuerMetadataEpr = "https://nexus.microsoftonline-p.com/FederationMetadata/2006-12/FederationMetadata.xml"
Write-Host -foregroundcolor White " `n Token Issuer Metadata EPR:"
if ($fedtrust.TokenIssuerMetadataEpr.AbsoluteUri -like $fedtrustTokenIssuerMetadataEpr){
Write-Host -foregroundcolor Green "  Token Issuer Metadata EPR is " $fedtrust.TokenIssuerMetadataEpr.AbsoluteUri
}
else
{
Write-Host -foregroundcolor Red " Token Issuer Metadata EPR is Not " $fedtrust.TokenIssuerMetadataEpr.AbsoluteUri
}
$fedtrustTokenIssuerEpr = "https://login.microsoftonline.com/extSTS.srf"
Write-Host -foregroundcolor White " `n Token Issuer EPR:"
if ($fedtrust.TokenIssuerEpr.AbsoluteUri -like $fedtrustTokenIssuerEpr){
Write-Host -foregroundcolor Green "  Token Issuer EPR is:" $fedtrust.TokenIssuerEpr.AbsoluteUri
}
else
{
Write-Host -foregroundcolor Red "  Token Issuer EPR is Not:" $fedtrust.TokenIssuerEpr.AbsoluteUri
}




}

Function AutoDVirtualDSTV{
Write-Host -foregroundcolor Green "  Standard Configuration Values `n "  
Write-Host $barspace
Write-Host -foregroundcolor Green " Authentication `n "  

Write-Host -foregroundcolor White " - The *AuthenticationMethods must include WSSecurity. Verify WSSecurityAuthentication and WindowsAuthentication 
are set to True. By default, BasicAuthentication is also enabled. If you use the -ADPropertiesOnly switch, the *Authentication would be blank. 
Check ExternalAuthenticationMethods instead." 
Write-Host $spacebar
Write-Host -foregroundcolor Green " On-Prem Autodiscover Virtual Directory Configuration Example" 
Write-Host $barspace
Write-Host -foregroundcolor Yellow " Name                          : Autodiscover (Default Web Site)" 
Write-Host -foregroundcolor Yellow " ExchangeVersion               : 0.10 (14.0.100.0)" 
Write-Host -foregroundcolor Yellow " InternalAuthenticationMethods : {Basic, Ntlm, WindowsIntegrated, WSSecurity}" 
Write-Host -foregroundcolor Yellow " ExternalAuthenticationMethods : {Basic, Ntlm, WindowsIntegrated, WSSecurity}" 
Write-Host -foregroundcolor Yellow " LiveIdSpNegoAuthentication    : False" 
Write-Host -foregroundcolor Yellow " WSSecurityAuthentication      : True" 
Write-Host -foregroundcolor Yellow " LiveIdBasicAuthentication     : False" 
Write-Host -foregroundcolor Yellow " BasicAuthentication           : True" 
Write-Host -foregroundcolor Yellow " DigestAuthentication          : False" 
Write-Host -foregroundcolor Yellow " WindowsAuthentication         : True `n"    
}

Function AutoDVirtualDCheck{
Write-Host -foregroundcolor Green " `n On-Prem Autodiscover Virtual Directory `n "  
Write-Host -foregroundcolor Green " Get-AutodiscoverVirtualDirectory | Select Identity,Name,ExchangeVersion,*authentication*" 


Write-Host $spacebarspace

$Global:AutoDiscoveryVirtualDirectory = Get-AutodiscoverVirtualDirectory | Select Identity,Name,ExchangeVersion,*authentication*  

#Check if null or set
#$AutoDiscoveryVirtualDirectory
$Global:AutoDiscoveryVirtualDirectory

Write-Host $spacebar

Write-Host -foregroundcolor Green " SUMARY - On-Prem Autodiscover Virtual Directory (non standard values will show up in Red. Standard Values in Green)" 

Write-Host $barspace

#Write-Host -foregroundcolor White "  WSSecurityAuthentication:" 

if ($Global:AutoDiscoveryVirtualDirectory.WSSecurityAuthentication -eq "True"){
foreach( $ser in $Global:AutoDiscoveryVirtualDirectory) { 
Write-Host " $($ser.Identity) "
Write-Host -ForegroundColor Green "  WSSecurityAuthentication: $($ser.WSSecurityAuthentication) `n  " 
}

}
else
{
Write-Host -foregroundcolor Red " `n  WSSecurityAuthentication is NOT correct."
foreach( $ser in $Global:AutoDiscoveryVirtualDirectory) { 
Write-Host " $($ser.Identity)"
Write-Host -ForegroundColor Red "  WSSecurityAuthentication: $($ser.WSSecurityAuthentication) `n  " 
}

Write-Host -foregroundcolor White "  Should be True `n"
}

}

Function EWSVirtualDirectorySTV{
Write-Host $bar
Write-Host -foregroundcolor Green "  Standard Configuration Values"  
Write-Host $barspace
Write-Host -foregroundcolor Green "  Authentication `n " 
Write-Host -foregroundcolor White " - The *AuthenticationMethods must include WSSecurity. Verify WSSecurityAuthentication and WindowsAuthentication 
are set to True. If you use the -ADPropertiesOnly switch, the *Authentication would be blank. Check ExternalAuthenticationMethods instead. `n " 
Write-Host -foregroundcolor Green "  ExternalUrl `n " 
Write-Host -foregroundcolor White " - The should be the public FQDN of the hybrid server. Example: https://server.contoso.com/EWS/Exchange.asmx `n " 
Write-Host -foregroundcolor Green "  Notes:	`n " 
Write-Host -foregroundcolor Yellow " - Depending on your topology with single site, multiple sites, internet facing site or non-internet facing site, 
your web services virtual directory ExternalUrl may or may not be set. Please consult the proper deployment guide for reference." 
Write-Host -foregroundcolor Yell " - The hybrid servers must have the ExternalUrl set, it cannot be blank. `n " 
Write-Host $spacebar
Write-Host -foregroundcolor Green "  On-Prem Web Services Virtual Directory Configuration Example " 
Write-Host $barspace
Write-Host -foregroundcolor Yellow "  Name                          : EWS (Default Web Site)" 
Write-Host -foregroundcolor Yellow "  ExchangeVersion               : 0.10 (14.0.100.0)" 
Write-Host -foregroundcolor Yellow "  CertificateAuthentication     :" 
Write-Host -foregroundcolor Yellow "  InternalAuthenticationMethods : {Ntlm, WindowsIntegrated, WSSecurity}" 
Write-Host -foregroundcolor Yellow "  ExternalAuthenticationMethods : {Ntlm, WindowsIntegrated, WSSecurity}" 
Write-Host -foregroundcolor Yellow "  LiveIdSpNegoAuthentication    : False" 
Write-Host -foregroundcolor Yellow "  WSSecurityAuthentication      : True" 
Write-Host -foregroundcolor Yellow "  LiveIdBasicAuthentication     : False" 
Write-Host -foregroundcolor Yellow "  BasicAuthentication           : False" 
Write-Host -foregroundcolor Yellow "  DigestAuthentication          : False" 
Write-Host -foregroundcolor Yellow "  WindowsAuthentication         : True" 
Write-Host -foregroundcolor Yellow "  InternalNLBBypassUrl          : https://server01.contoso.com/EWS/Exchange.asmx" 
Write-Host -foregroundcolor Yellow "  InternalUrl                   : https://server01.contoso.com/EWS/Exchange.asmx" 
Write-Host -foregroundcolor Yellow "  ExternalUrl                   : https://server.contoso.com/EWS/Exchange.asmx `n" 
}

Function EWSVirtualDirectoryCheck{


Write-Host -foregroundcolor Green " `n On-Prem Web Services Virtual Directory `n " 
Write-Host -foregroundcolor Green " Get-WebServicesVirtualDirectory | Select Identity,Name,ExchangeVersion,*Authentication*,*url `n " 
Write-Host $spacebarspace
$Global:WebServicesVirtualDirectory = Get-WebServicesVirtualDirectory | Select Identity,Name,ExchangeVersion,*Authentication*,*url
$Global:WebServicesVirtualDirectory
Write-Host $spacebar
Write-Host -foregroundcolor Green " SUMARY - On-Prem Web Services Virtual Directory (non standard values will show up in Red. Standard Values in Green)" 
Write-Host $barspace
#Write-Host -foregroundcolor White "  WSSecurityAuthentication: `n " 
if ($Global:WebServicesVirtualDirectory.WSSecurityAuthentication -like  "True"){

foreach( $EWS in $Global:WebServicesVirtualDirectory) { 
Write-Host " $($EWS.Identity)"
Write-Host -ForegroundColor Green "  WSSecurityAuthentication: $($EWS.WSSecurityAuthentication) `n  " 
}

}
else
{
Write-Host -foregroundcolor Red " `n  WSSecurityAuthentication is NOT correct."
foreach( $EWS in $Global:AutoDiscoveryVirtualDirectory) { 
Write-Host " $($EWS.Identity) "
Write-Host -ForegroundColor Red "  WSSecurityAuthentication: $($ser.WSSecurityAuthentication) `n  " 
}
Write-Host -foregroundcolor White "  Should be True `n"
}

}

Function AvailabilityAddressSpaceSTV{

Write-Host $spacebar
Write-Host -foregroundcolor Green  " Configuration Example" 
Write-Host $barspace
Write-Host -foregroundcolor Yellow " ForestName        : contoso.mail.onmicrosoft.com" 
Write-Host -foregroundcolor Yellow " UserName          :" 
Write-Host -foregroundcolor Yellow " UseServiceAccount : True" 
Write-Host -foregroundcolor Yellow " AccessMethod      : InternalProxy" 
Write-Host -foregroundcolor Yellow " ProxyUrl          : https://server01.contoso.com/EWS/Exchange.asmx" 
Write-Host -foregroundcolor Yellow " Name              : contoso.mail.onmicrosoft.com" 
Write-Host $spacebar
Write-Host  "  Standard Configuration Values"
Write-Host $spacebar
Write-Host -foregroundcolor Green "  ForestName `n " 

Write-Host -foregroundcolor White "  - The should be the contoso.mail.onmicrosoft.com  domain name. This should also match the domain 
name of RemoteRoutingAddress of remote mailboxes. `n "  
Write-Host "  Example: contoso.mail.onmicrosoft.com `n "  
Write-Host -foregroundcolor Green "  UserName `n " 
Write-Host -foregroundcolor White "  - This should be blank. `n " 
Write-Host -foregroundcolor Green "  UseServiceAccount `n "  
Write-Host -foregroundcolor White "  - This must be True. `n " 
Write-Host -foregroundcolor Green "  AccessMethod `n " 
Write-Host -foregroundcolor White "  - This should be InternalProxy. `n "  
Write-Host -foregroundcolor Green "  ProxyUrl `n " 
Write-Host -foregroundcolor White "  - This should be the Exchange 2010 (Hybrid 2010) or 2013/2016 (Hybrid 2013/2016) Exchange Web 
Services Virtual Directory url. The address could be the internal FQDN or load balancing EWS url. Example: https://server01.contoso.com/EWS/Exchange.asmx `n " 
}

Function AvailabilityAddressSpaceCheck{
Write-Host -foregroundcolor Green " Get-AvailabilityAddressSpace $exchangeOnlineDomain | fl ForestName, UserName, UseServiceAccount, AccessMethod, ProxyUrl, Name"
Write-Host $barspace
$script:AvailabilityAddressSpace = Get-AvailabilityAddressSpace $exchangeOnlineDomain | select ForestName, UserName, UseServiceAccount, AccessMethod, ProxyUrl, Name
$AvailabilityAddressSpace
Write-Host $spacebar
Write-Host -foregroundcolor Green " SUMARY - On-Prem Availability Address Space (non standard values will show up in Red. Standard Values in Green)" 
Write-Host $barspace
Write-Host -foregroundcolor White " ForestName: " 
if ($AvailabilityAddressSpace.ForestName -like  $ExchangeOnlineDomain){
Write-Host -foregroundcolor Green " " $AvailabilityAddressSpace.ForestName
}
else
{
Write-Host -foregroundcolor Red " ForestName is NOT correct. `n "
Write-Host -foregroundcolor White " Should be $ExchaneOnlineDomain "
}

Write-Host -foregroundcolor White " `n UserName: " 
if ($AvailabilityAddressSpace.UserName -like  ""){
Write-Host -foregroundcolor Green "  Blank `n " 
}
else
{
Write-Host -foregroundcolor Red " `n  UserName is NOT correct. "
Write-Host -foregroundcolor White "  Should be blank `n "
}
Write-Host -foregroundcolor White " UseServiceAccount: " 
if ($AvailabilityAddressSpace.UseServiceAccount -like  "True"){ 
Write-Host -foregroundcolor Green "  True `n "  
}
else
{
Write-Host -foregroundcolor Red "  UseServiceAccount is NOT correct."
Write-Host -foregroundcolor White " `n  Should be True `n"
}



Write-Host -foregroundcolor White " AccessMethod:" 
if ($AvailabilityAddressSpace.AccessMethod -like  "InternalProxy"){
Write-Host -foregroundcolor Green " InternalProxy `n " 
}
else
{
Write-Host -foregroundcolor Red " AccessMethod is NOT correct. `n "
Write-Host -foregroundcolor White " Should be InternalProxy `n "
}

Write-Host -foregroundcolor White " ProxyUrl: " 
if ($AvailabilityAddressSpace.ProxyUrl -like  $exchangeOnPremEWS){
Write-Host -foregroundcolor Green "  "$AvailabilityAddressSpace.ProxyUrl 

}
else
{
Write-Host -foregroundcolor Red "  ProxyUrl is NOT correct. `n "
Write-Host -foregroundcolor White "  Should be $exchangeOnPremEWS"
}
#flata o ews
#Write-Host $spacebar
}

Function TestFedTrust{
$TestFedTrustFail = 0
$a = Test-FederationTrust -UserIdentity $useronprem -verbose  #fails the frist time on multiple ocasions
Write-Host -foregroundcolor Green  " Test-FederationTrust -UserIdentity $useronprem -verbose"  
Write-Host $bar
$script:TestFedTrust = Test-FederationTrust -UserIdentity $useronprem -verbose  
$script:TestFedTrust


foreach( $test in $TestFedTrust.type) { 

#$test
if ($test -ne "Success"){
Write-Host -foregroundcolor Red " $($test.Type)  "
Write-Host " $($test.Message)  "
$TestFedTrustFail++
}




}

if ($TestFedTrustFail -eq  0){
Write-Host -foregroundcolor Green " Federation Trust Successfully tested `n "
}
else  {

Write-Host -foregroundcolor Red " Federation Trust test with Errors `n "
#Check this an that

}



#Write-Host $bar
}

Function TestOrgRel{
$TestFail = "0"
Write-Host -foregroundcolor Green "Test-OrganizationRelationship -Identity "On-premises to O365*"  -UserIdentity $useronprem" 
#need to grab errors and provide alerts in error case 
Write-Host $barspace
#this test needs to be more effective and Identity passed as variable
$TestOrgRel = Test-OrganizationRelationship -Identity "On-premises to O365*"  -UserIdentity $useronprem 
$TestOrgRel




foreach( $test in $TestOrgRel.type) { 

#$test
if ($test -ne "Success"){
Write-Host -foregroundcolor Red " $($test.Type)  "
Write-Host " $($test.Message)  "
$TestFail++
}




}
if ($TestFail -eq "0"){
Write-Host -foregroundcolor Green " Organization Relationship Successfully tested `n "
}
else  {

Write-Host -foregroundcolor Red " Organization Relationship test with Errors `n "
#Check this an that

}

Write-Host $spacebarspace

}
#endregion

#region OAuth Functions

Function IntraOrgConSTV{
Write-Host $bar
Write-Host -foregroundcolor Green " Configuration Example" 
Write-Host $bar
Write-Host -foregroundcolor Yellow  " `n Name                 : ExchangeHybridOnPremisesToOnline"
Write-Host -foregroundcolor Yellow " TargetAddressDomains : {contoso.mail.onmicrosoft.com}"
Write-Host -foregroundcolor Yellow " DiscoveryEndpoint    : https://outlook.office365.com/autodiscover/autodiscover.svc"
Write-Host -foregroundcolor Yellow " Enabled              : True `n"
Write-Host $bar
Write-Host -foregroundcolor Green " Standard Configuration Values" 
Write-Host $bar
Write-Host -ForegroundColor Green " `n TargetAddressDomains: `n" 
Write-Host -ForegroundColor White " - This should be customer's contoso.mail.onmicroosft.com domain name." 
Write-Host -foregroundcolor Yellow Example: "contoso.mail.onmicrosoft.com `n "
Write-Host -ForegroundColor White " TargetDiscoveryEndpoint: `n"
Write-Host -ForegroundColor White " - This should be the address of EXO autodiscover endpoint." 
Write-Host -foregroundcolor Yellow "Example: https://outlook.office365.com/autodiscover/autodiscover.svc `n" 
Write-Host -ForegroundColor Green "Enabled:"
Write-Host -ForegroundColor White " - This must be True."
}

Function IntraOrgConCheck{

Write-Host $bar
Write-Host -foregroundcolor Green " Get-IntraOrganizationConnector | FL Name,TargetAddressDomains,DiscoveryEndpoint,Enabled" 
Write-Host $bar

$IntraOrgConCheck = Get-IntraOrganizationConnector | FL Name,TargetAddressDomains,DiscoveryEndpoint,Enabled
$IntraOrgConCheck
Write-Host $spacebar
Write-Host -foregroundcolor Green " SUMARY - On-Prem Intra Organization Connector (non standard values will show up in Red. Standard Values in Green)" 
Write-Host $barspace

$IntraOrgTargetAddressDomain = $IntraOrgCon.TargetAddressDomains.Domain
$IntraOrgTargetAddressDomain = $IntraOrgTargetAddressDomain.Tolower()
Write-Host -foregroundcolor White " Target Address Domains: " 
if ($IntraOrgCon.TargetAddressDomains-like  "*$ExchangeOnlineDomain*" -Or $IntraOrgCon.TargetAddressDomains -like  "*$ExchangeOnlineAltDomain*" ){
Write-Host -foregroundcolor Green " " $IntraOrgCon.TargetAddressDomains
}
else
{
Write-Host -foregroundcolor Red " Target Address Domains is NOT correct. `n "
Write-Host -foregroundcolor White " Should contain the $ExchangeOnlineDomain domain or the $ExchangeOnlineAltDomain"
}

Write-Host -foregroundcolor White " `n DiscoveryEndpoint: " 
if ($IntraOrgCon.DiscoveryEndpoint -like  "https://autodiscover-s.outlook.com/autodiscover/autodiscover.svc"){
Write-Host -foregroundcolor Green "  https://autodiscover-s.outlook.com/autodiscover/autodiscover.svc `n " 
}
else
{
Write-Host -foregroundcolor Red " `n  DiscoveryEndpoint are NOT correct. "
Write-Host -foregroundcolor White "  Should be https://autodiscover-s.outlook.com/autodiscover/autodiscover.svc `n "
}
Write-Host -foregroundcolor White " Enabled: " 
if ($IntraOrgCon.Enabled -like  "True"){ 
Write-Host -foregroundcolor Green "  True `n "  
}
else
{
Write-Host -foregroundcolor Red "  Enabled is NOT correct."
Write-Host -foregroundcolor White " Should be True `n"
}




}

Function AuthServerSTV{

Write-Host $bar
Write-Host -foregroundcolor Green " Configuration Example" 
Write-Host $barspace
write-Host -foregroundcolor Yellow " Name                 : WindowsAzureACS"
write-Host -foregroundcolor Yellow " IssuerIdentifier     : 00000001-0000-0000-c000-000000000000"
write-Host -foregroundcolor Yellow " TokenIssuingEndpoint : https://accounts.accesscontrol.windows.net/cb52ada4-5045-4d00-a59a-c7896ef052a1/tokens/OAuth/2"
write-Host -foregroundcolor Yellow " AuthMetadataUrl      : https://accounts.accesscontrol.windows.net/contoso.com/metadata/json/1"
write-Host -foregroundcolor Yellow " Enabled              : True `n "
Write-Host $bar
Write-Host -foregroundcolor Green " Standard Configuration Values" 
Write-Host $barspace
Write-Host -foregroundcolor Green " AuthMetadataUrl:" 
Write-Host -ForegroundColor White " - This should be https://accounts.accesscontrol.windows.net/ <hybrid_domain>/metadata/json/1. "
write-Host -foregroundcolor White " The <hybrid domain> referenced should be the primary SMTP domain of cloud mailboxes." 
Write-Host -ForegroundColor White  " If there are multiple SMTP domains for different mailboxes, any primary SMTP should be ok." 
Write-Host -ForegroundColor Yellow  "`n Example: https://accounts.accesscontrol.windows.net/contoso.com/metadata/json/1 `n "
Write-Host -foregroundcolor Green " TokenIssuingEndpoint: `n "
Write-Host -foregroundcolor White " - The GUID referenced in the URL is the Company Id/Tenant Id of the Office 365 tenant. You can retrieve the value from ViewPoint "
Write-Host -foregroundcolor White " -> tenant info -> Company Id. See curated:Outbound Access to MFG\AuthServer for more information. `n "
Write-Host -foregroundcolor Green " Enabled: `n "  
Write-Host -foregroundcolor White " - This must be True. `n"
}

Function AuthServerCheck{

Write-Host $bar
Write-Host -foregroundcolor Green " Get-AuthServer | Select Name,IssuerIdentifier,TokenIssuingEndpoint,AuthMetadataUrl,Enabled" 
Write-Host $bar
$AuthServer = Get-AuthServer | Where {$_.Name -like "ACS*"} | Select Name,IssuerIdentifier,TokenIssuingEndpoint,AuthMetadataUrl,Enabled
$AuthServer

Write-Host $spacebar
Write-Host -foregroundcolor Green " SUMARY - Auth Server (non standard values will show up in Red. Standard Values in Green)" 
Write-Host $barspace

Write-Host -foregroundcolor White " IssuerIdentifier: " 
if ($AuthServer.IssuerIdentifier -like  "00000001-0000-0000-c000-000000000000" ){
Write-Host -foregroundcolor Green " " $AuthServer.IssuerIdentifier
}
else
{
Write-Host -foregroundcolor Red " IssuerIdentifier is NOT correct. `n "
Write-Host -foregroundcolor White " Should be 00000001-0000-0000-c000-000000000000"
}

Write-Host -foregroundcolor White " TokenIssuingEndpoint: " 
if ($AuthServer.TokenIssuingEndpoint -like  "https://accounts.accesscontrol.windows.net/*"  -and $AuthServer.TokenIssuingEndpoint -like  "*/tokens/OAuth/2" ){
Write-Host -foregroundcolor Green " " $AuthServer.TokenIssuingEndpoint
}
else
{
Write-Host -foregroundcolor Red " TokenIssuingEndpoint is NOT correct. `n "
Write-Host -foregroundcolor White " Should be  https://accounts.accesscontrol.windows.net/<Cloud Tenant ID>/tokens/OAuth/2"
}


Write-Host -foregroundcolor White " AuthMetadataUrl: " 
if ($AuthServer.AuthMetadataUrl -like  "https://accounts.accesscontrol.windows.net/*"  -and $AuthServer.TokenIssuingEndpoint -like  "*/tokens/OAuth/2" ){
Write-Host -foregroundcolor Green " " $AuthServer.AuthMetadataUrl
}
else
{
Write-Host -foregroundcolor Red " AuthMetadataUrl is NOT correct. `n "
Write-Host -foregroundcolor White " Should be  https://accounts.accesscontrol.windows.net/<Cloud Tenant ID>/metadata/json/1"
}



Write-Host -foregroundcolor White " Enabled: " 
if ($AuthServer.Enabled -like  "True" ){
Write-Host -foregroundcolor Green " " $AuthServer.Enabled
}
else
{
Write-Host -foregroundcolor Red " Enalbed: False "
Write-Host -foregroundcolor White " Should be True"
}



}

Function PartnerApplicationSTV{
Write-Host $spacebar
Write-Host -foregroundcolor Green " Configuration Example  " 
Write-Host $barspace
Write-Host -foregroundcolor Yellow " Enabled                             : True"
Write-Host -foregroundcolor Yellow  " ApplicationIdentifier               : 00000002-0000-0ff1-ce00-000000000000"
Write-Host -foregroundcolor Yellow  " CertificateStrings                  : {}"
Write-Host -foregroundcolor Yellow  " AuthMetadataUrl                     :"
Write-Host -foregroundcolor Yellow  " Realm                               :"
Write-Host -foregroundcolor Yellow  " UseAuthServer                       : True"
Write-Host -foregroundcolor Yellow  " AcceptSecurityIdentifierInformation : False"
Write-Host -foregroundcolor Yellow  " LinkedAccount                       : contoso.com/Users/Exchange Online-ApplicationAccount"
Write-Host -foregroundcolor Yellow  " IssuerIdentifier                    :"
Write-Host -foregroundcolor Yellow  " AppOnlyPermissions                  :"
Write-Host -foregroundcolor Yellow  " ActAsPermissions                    :"
Write-Host -foregroundcolor Yellow  " Name                                : Exchange Online `n " 
Write-Host $bar
Write-Host -foregroundcolor Green " Standard Configuration Values" 
Write-Host $bar
Write-Host -foregroundcolor Green " `n ApplicationIdentifier"
Write-Host -foregroundcolor White " - There should already be a partner application with ApplicationIdentifier = 00000002-0000-0ff1-ce00-000000000000 and with a blank empty Realm. `n "
Write-Host -foregroundcolor Green " AuthMetadataUrl"
Write-Host -foregroundcolor White " - This should be blank. `n "
Write-Host -foregroundcolor Green " Enabled"
Write-Host -foregroundcolor White " - This must be True. `n "
Write-Host -foregroundcolor Green " LinkedAccount"
Write-Host -foregroundcolor White " - This references a linked user account. Verify the user account is present. If you value is empty, set it back to Exchange Online-ApplicationAccount"
Write-Host -foregroundcolor White " which is located at the root of Users container in AD. After you make the change, reboot the servers. `n "
Write-Host -foregroundcolor Yellow " Example: contoso.com/Users/Exchange  Online-ApplicationAccount" 
Write-Host -foregroundcolor Yellow " **Note:**The account has 5 RBAC assignments at on-premises: `n"
}

Function PartnerApplicationCheck{
Write-Host $bar
Write-Host -foregroundcolor Green " Get-PartnerApplication |  ?{$_.ApplicationIdentifier -eq '00000002-0000-0ff1-ce00-000000000000'
 -and $_.Realm -eq ''} | fl Enabled, ApplicationIdentifier, CertificateStrings, AuthMetadataUrl, Realm, UseAuthServer, 
 AcceptSecurityIdentifierInformation, LinkedAccount, IssuerIdentifier, AppOnlyPermissions, ActAsPermissions, Name" 
Write-Host $bar

$PartnerApplication = Get-PartnerApplication |  ?{$_.ApplicationIdentifier -eq '00000002-0000-0ff1-ce00-000000000000' -and $_.Realm -eq ''} | Select Enabled, ApplicationIdentifier, CertificateStrings, AuthMetadataUrl, Realm, UseAuthServer, AcceptSecurityIdentifierInformation, LinkedAccount, IssuerIdentifier, AppOnlyPermissions, ActAsPermissions, Name
$PartnerApplication
Write-Host $spacebar
Write-Host -foregroundcolor Green " SUMARY - Partner Application (non standard values will show up in Red. Standard Values in Green)" 
Write-Host $barspace


Write-Host -foregroundcolor White " Enabled: " 
if ($PartnerApplication.Enabled -like  "True" ){
Write-Host -foregroundcolor Green " " $PartnerApplication.Enabled
}
else
{
Write-Host -foregroundcolor Red " Enabled: False  `n "
Write-Host -foregroundcolor White " Should be True"
}

Write-Host -foregroundcolor White "`n ApplicationIdentifier: " 
if ($PartnerApplication.ApplicationIdentifier -like  "00000002-0000-0ff1-ce00-000000000000" ){
Write-Host -foregroundcolor Green " " $PartnerApplication.ApplicationIdentifier
}
else
{
Write-Host -foregroundcolor Red " ApplicationIdentifier does not seem correct  `n "
Write-Host -foregroundcolor White " Should be 00000002-0000-0ff1-ce00-000000000000"
}

Write-Host -foregroundcolor White "`n AuthMetadataUrl: " 
if ([string]::IsNullOrWhitespace( $PartnerApplication.AuthMetadataUrl)){
Write-Host -foregroundcolor Green "  Blank" 
}
else
{
Write-Host -foregroundcolor Red " AuthMetadataUrl does not seem correct  `n "
Write-Host -foregroundcolor White " Should be Blank"
}



Write-Host -foregroundcolor White "`n Realm: " 
if ([string]::IsNullOrWhitespace( $PartnerApplication.Realm)){
Write-Host -foregroundcolor Green "  Blank"
}
else
{
Write-Host -foregroundcolor Red "  Realm does not seem correct  `n "
Write-Host -foregroundcolor White " Should be Blank"
}


Write-Host -foregroundcolor White "`n LinkedAccount: " 
if ($PartnerApplication.LinkedAccount -like  "$exchangeOnPremDomain/Users/Exchange Online-ApplicationAccount" -or $PartnerApplication.LinkedAccount -like  "$exchangeOnPremLocalDomain/Users/Exchange Online-ApplicationAccount"  ){
Write-Host -foregroundcolor Green " " $PartnerApplication.LinkedAccount
}
else
{
Write-Host -foregroundcolor Red " LinkedAccount value does not seem correct  `n "
Write-Host -foregroundcolor White " Should be $exchangeOnPremLocalDomain/Users/Exchange Online-ApplicationAccount"
}
}



Function ApplicationAccountSTV{
Write-Host $bar
Write-Host -foregroundcolor Green " Configuration Example" 
Write-Host $barspace
Write-Host -foregroundcolor Yellow " Name                 : Exchange Online-ApplicationAccount"
Write-Host -foregroundcolor Yellow " RecipientType        : User"
Write-Host -foregroundcolor Yellow " RecipientTypeDetails : LinkedUser"
Write-Host -foregroundcolor Yellow " UserAccountControl   : AccountDisabled, PasswordNotRequired, NormalAccount `n"
}

Function ApplicationAccounCheck{

Write-Host $bar
Write-Host -foregroundcolor Green " Get-user '$exchangeOnPremLocalDomain/Users/Exchange Online-ApplicationAccount' | Select Name, RecipientType, 
RecipientTypeDetails, UserAccountControl" 
Write-Host $bar
$ApplicationAccount = Get-user "$exchangeOnPremLocalDomain/Users/Exchange Online-ApplicationAccount"  | Select Name, RecipientType, RecipientTypeDetails, UserAccountControl
$ApplicationAccount

Write-Host $spacebar
Write-Host -foregroundcolor Green " SUMARY - Application Account (non standard values will show up in Red. Standard Values in Green)" 
Write-Host $barspace

Write-Host -foregroundcolor White " RecipientType: " 
if ($ApplicationAccount.RecipientType -like  "User" ){
Write-Host -foregroundcolor Green " " $ApplicationAccount.RecipientType
}
else
{
Write-Host -foregroundcolor Red " RecipientType value is $ApplicationAccount.RecipientType `n "
Write-Host -foregroundcolor White " Should be User"
}

Write-Host -foregroundcolor White " RecipientTypeDetails: " 
if ($ApplicationAccount.RecipientTypeDetails -like  "LinkedUser" ){
Write-Host -foregroundcolor Green " " $ApplicationAccount.RecipientTypeDetails
}
else
{
Write-Host -foregroundcolor Red " RecipientTypeDetails value is $ApplicationAccount.RecipientTypeDetails `n "
Write-Host -foregroundcolor White " Should be LinkedUser"
}


Write-Host -foregroundcolor White " UserAccountControl: " 
if ($ApplicationAccount.UserAccountControl -like  "AccountDisabled, PasswordNotRequired, NormalAccount" ){
Write-Host -foregroundcolor Green " " $ApplicationAccount.UserAccountControl
}
else
{
Write-Host -foregroundcolor Red " UserAccountControl value does not seem correct `n "
Write-Host -foregroundcolor White " Should be AccountDisabled, PasswordNotRequired, NormalAccount"
}

}



Function ManagementRoleAssignmentSTV{
Write-Host $spacebar
Write-Host -foregroundcolor Green " Configuration Example" 
Write-Host $barspace
Write-Host -foregroundcolor Yellow "   Name                                                             Role"
Write-Host -foregroundcolor Yellow "   ----                                                             ----"
Write-Host -foregroundcolor Yellow "   UserApplication-Exchange Online-ApplicationAccount               UserApplication"
Write-Host -foregroundcolor Yellow "   ArchiveApplication-Exchange Online-ApplicationAccount            ArchiveApplication"
Write-Host -foregroundcolor Yellow "   LegalHoldApplication-Exchange Online-ApplicationAccount          LegalHoldApplication"
Write-Host -foregroundcolor Yellow "   Mailbox Search-Exchange Online-ApplicationAccount                Mailbox Search"
Write-Host -foregroundcolor Yellow "   TeamMailboxLifecycleApplication-Exchange Online-ApplicationAccou TeamMailboxLifecycleApplication"
Write-Host -foregroundcolor Yellow "   MailboxSearchApplication-Exchange Online-ApplicationAccount      MailboxSearchApplication `n "
}

Function AuthConfigSTV{

Write-Host $spacebar
Write-Host -foregroundcolor Green " Configuration Example" 
Write-Host $barspace
Write-Host -foregroundcolor Yellow "  CurrentCertificateThumbprint  : 8BB44C52F03EF30AEAB27528A1019F0887A7A383"
Write-Host -foregroundcolor Yellow "  PreviousCertificateThumbprint :"
Write-Host -foregroundcolor Yellow "  NextCertificateThumbprint     :"
Write-Host -foregroundcolor Yellow "  ServiceName                   : 00000002-0000-0ff1-ce00-000000000000"
Write-Host -foregroundcolor Yellow "  Realm                         :"
Write-Host -foregroundcolor Yellow "  Name                          : Auth Configuration `n "
}

Function CurrentCertificateThumbprintSTV{
Write-Host $bar
Write-Host -foregroundcolor Green " Configuration Example" 
Write-Host $barspace
Write-Host -foregroundcolor Yellow "  AccessRules        : {System.Security.AccessControl.CryptoKeyAccessRule,"
Write-Host -foregroundcolor Yellow "                        System.Security.AccessControl.CryptoKeyAccessRule,"
Write-Host -foregroundcolor Yellow "                        System.Security.AccessControl.CryptoKeyAccessRule}"
Write-Host -foregroundcolor Yellow "  CertificateDomains : {}"
Write-Host -foregroundcolor Yellow "  HasPrivateKey      : True"
Write-Host -foregroundcolor Yellow "  IsSelfSigned       : True"
Write-Host -foregroundcolor Yellow "  Issuer             : CN=Microsoft Exchange Server Auth Certificate"
Write-Host -foregroundcolor Yellow "  NotAfter           : 9/13/2017 5:26:01 PM"
Write-Host -foregroundcolor Yellow "  NotBefore          : 10/9/2012 5:26:01 PM"
Write-Host -foregroundcolor Yellow "  PublicKeySize      : 2048"
Write-Host -foregroundcolor Yellow "  RootCAType         : None"
Write-Host -foregroundcolor Yellow "  SerialNumber       : 310196A39840EB964D5E2FD29D3B431C"
Write-Host -foregroundcolor Yellow "  Services           : SMTP"
Write-Host -foregroundcolor Yellow "  Status             : Valid"
Write-Host -foregroundcolor Yellow "  Subject            : CN=Microsoft Exchange Server Auth Certificate"
Write-Host -foregroundcolor Yellow "  Thumbprint         : 8BB44C52F03EF30AEAB27528A1019F0887A7A383 `n " 
Write-Host $bar
Write-Host -foregroundcolor Green " Standard Configuration Values" 
Write-Host $barspace
Write-Host -foregroundcolor Green " ServiceName"
Write-Host -foregroundcolor White "  - This should be 00000002-0000-0ff1-ce00-000000000000. `n "
Write-Host -foregroundcolor Green " Realm"
Write-Host -foregroundcolor White "  - This should be blank. `n " 
Write-Host -foregroundcolor Green " CurrentCertificateThumbprint"
Write-Host -foregroundcolor White "  - The certificate referenced should exist on all of your Exchange 2013/2016 servers. `n"

}




Function AutoDVirtualDCheckOauth{
Write-Host -foregroundcolor Green " `n On-Prem Autodiscover Virtual Directory `n "  
Write-Host -foregroundcolor Green " Get-AutodiscoverVirtualDirectory | FL Identity, Name,ExchangeVersion,*authentication*" 


Write-Host $spacebarspace

$AutoDiscoveryVirtualDirectoryOAuth = Get-AutodiscoverVirtualDirectory | FL Identity,Name,ExchangeVersion,*authentication*  

#Check if null or set
$AutoDiscoveryVirtualDirectoryOAuth
Write-Host $spacebar

Write-Host -foregroundcolor Green " SUMARY - On-Prem Autodiscover Virtual Directory (non standard values will show up in Red. Standard Values in Green)" 

Write-Host $barspace

Write-Host -foregroundcolor White "  WSSecurityAuthentication:" 

if ($Global:AutoDiscoveryVirtualDirectory.WSSecurityAuthentication -like  "True"){
#Write-Host -foregroundcolor Green " `n  " $Global:AutoDiscoveryVirtualDirectory.WSSecurityAuthentication
foreach( $ser in $Global:AutoDiscoveryVirtualDirectory) { 
Write-Host " $($ser.Identity) `n "
Write-Host -ForegroundColor Green " WSSecurityAuthentication: $($ser.WSSecurityAuthentication) `n  " 
}


}
else
{
Write-Host -foregroundcolor Red " `n WSSecurityAuthentication setting is NOT correct."
foreach( $ser in $Global:AutoDiscoveryVirtualDirectory) { 
Write-Host " $($ser.Identity) `n "
Write-Host -ForegroundColor Red " WSSecurityAuthentication: $($ser.WSSecurityAuthentication) `n  " 
}

Write-Host $spacebar
}
}

Function EWSVirtualDirectoryCheckOAuth{


Write-Host -foregroundcolor Green " `n On-Prem Web Services Virtual Directory `n " 
Write-Host -foregroundcolor Green " Get-WebServicesVirtualDirectory | fl Name,ExchangeVersion,*Authentication*,*url `n " 
Write-Host $spacebarspace

$WebServicesVirtualDirectoryOAuth = Get-WebServicesVirtualDirectory | fl Name,ExchangeVersion,*Authentication*,*url

$WebServicesVirtualDirectoryOAuth
Write-Host $spacebar
Write-Host -foregroundcolor Green " SUMARY - On-Prem Web Services Virtual Directory (non standard values will show up in Red. Standard Values in Green)" 
Write-Host $barspace
Write-Host -foregroundcolor White "  WSSecurityAuthentication: `n " 
if ($Global:WebServicesVirtualDirectory.WSSecurityAuthentication -like  "True"){
#Write-Host -foregroundcolor Green "  " $Global:WebServicesVirtualDirectory.WSSecurityAuthentication

foreach( $EWS in $Global:WebServicesVirtualDirectory) { 
Write-Host " $($EWS.Identity) `n "
Write-Host -ForegroundColor Green " WSSecurityAuthentication: $($EWS.WSSecurityAuthentication) `n  " 
}

}
else
{
Write-Host -foregroundcolor Red " `n WSSecurityAuthentication is NOT correct."
foreach( $EWS in $Global:AutoDiscoveryVirtualDirectory) { 
Write-Host " $($EWS.Identity) `n "
Write-Host -ForegroundColor Red " WSSecurityAuthentication: $($ser.WSSecurityAuthentication) `n  " 
}
Write-Host -foregroundcolor White "  Should be True `n"
}


Write-Host $spacebar
}

Function AvailabilityAddressSpaceSTV{


Write-Host $spacebar
Write-Host -foregroundcolor Green " Configuration Example"  
Write-Host $barspace
Write-Host -foregroundcolor Yellow " ForestName        : contoso.mail.onmicrosoft.com" 
Write-Host -foregroundcolor Yellow " UserName          :" 
Write-Host -foregroundcolor Yellow " UseServiceAccount : True" 
Write-Host -foregroundcolor Yellow " AccessMethod      : InternalProxy" 
Write-Host -foregroundcolor Yellow " ProxyUrl          : https://server01.contoso.com/EWS/Exchange.asmx" 
Write-Host -foregroundcolor Yellow " Name              : contoso.mail.onmicrosoft.com `n " 
Write-Host $bar
Write-Host -ForegroundColor Green "  Standard Configuration Values" 
Write-Host $barspace
Write-Host -foregroundcolor Green "  ForestName `n " 
Write-Host -foregroundcolor White "  - The should be the contoso.mail.onmicrosoft.com  domain name. This should also match the 
domain name of RemoteRoutingAddress of remote mailboxes. `n "  
" Example: contoso.mail.onmicrosoft.com `n "  
Write-Host -foregroundcolor Green "  UserName `n " 
Write-Host -foregroundcolor White "  - This should be blank. `n " 
Write-Host -foregroundcolor Green "  UserServiceAccount `n "  
Write-Host -foregroundcolor White "  - This must be True. `n " 
Write-Host -foregroundcolor Green "  AccessMethod `n " 
Write-Host -foregroundcolor White "  - This should be InternalProxy. `n " 
Write-Host -foregroundcolor Green "  ProxyUrl `n " 
Write-Host -foregroundcolor White "  - This should be the Exchange 2010 (Hybrid 2010) or 2013/2016 (Hybrid 2013/2016) Exchange 
Web Services Virtual Directory url. The address could be the internal FQDN or load balancing EWS url. Example: https://server01.contoso.com/EWS/Exchange.asmx `n " 
}

Function AvailabilityAddressSpaceCheckOAuth{
Write-Host -foregroundcolor Green "Get-AvailabilityAddressSpace $exchangeOnlineDomain | fl ForestName, UserName, UseServiceAccount, AccessMethod, ProxyUrl, Name"
#Write-Host $barspace
$AvailabilityAddressSpaceOAuth = Get-AvailabilityAddressSpace $exchangeOnlineDomain | fl ForestName, UserName, UseServiceAccount, AccessMethod, ProxyUrl, Name
$AvailabilityAddressSpaceOAuth
Write-Host $spacebar
Write-Host -foregroundcolor Green " SUMARY - On-Prem Availability Address Space (non standard values will show up in Red. Standard Values in Green)" 
Write-Host $barspace
Write-Host -foregroundcolor White "  ForestName: `n " 
if ($AvailabilityAddressSpace.ForestName -like  $ExchangeOnlineDomain){
Write-Host -foregroundcolor Green " `n " $AvailabilityAddressSpace.ForestName
}
else
{
Write-Host -foregroundcolor Red " `n ForestName is NOT correct. `n "
Write-Host -foregroundcolor White "  Should be $ExchaneOnlineDomain `n "
}

Write-Host -foregroundcolor White "  UserName: `n " 
if ($AvailabilityAddressSpace.UserName -like  ""){
Write-Host -foregroundcolor Green "   Blank `n " 
}
else
{
Write-Host -foregroundcolor Red " `n UserName is NOT correct. `n "
Write-Host -foregroundcolor White " `n  Should be blank `n "
}
Write-Host -foregroundcolor White " `n  UseServiceAccount: `n " 
if ($AvailabilityAddressSpace.UseServiceAccount -like  "True"){ 
Write-Host -foregroundcolor Green " `n   True `n "  
}
else
{
Write-Host -foregroundcolor Red " `n UseServiceAccount is NOT correct."
Write-Host -foregroundcolor White " `n  Should be True `n"
}



Write-Host -foregroundcolor White "  AccessMethod: `n " 
if ($AvailabilityAddressSpace.AccessMethod -like  "InternalProxy"){
Write-Host -foregroundcolor Green "   InternalProxy `n " 
}
else
{
Write-Host -foregroundcolor Red "   AccessMethod is NOT correct. `n "
Write-Host -foregroundcolor White "   Should be InternalProxy `n "
}

Write-Host -foregroundcolor White "  ProxyUrl: `n " 
if ($AvailabilityAddressSpace.ProxyUrl -like  $exchangeOnPremEWS){
Write-Host -foregroundcolor Green "   "$AvailabilityAddressSpace.ProxyUrl 

}
else
{
Write-Host -foregroundcolor Red "   ProxyUrl is NOT correct. `n "
Write-Host -foregroundcolor White "   Should be $exchangeOnPremEWS"
}
#flata o ews
#Write-Host $spacebar
}



Function OAuthConnectivitySTV{
Write-Host $spacebar
Write-Host -foregroundcolor Green " Standard Configuration Values"
Write-Host $barspace
Write-Host -foregroundcolor Green " Note:" 
Write-Host -foregroundcolor Yellow " You can ignore the warning 'The SMTP address has no mailbox associated with it'" 
Write-Host -foregroundcolor Yellow " when the Test-OAuthConnectivity returns a Success `n "
Write-Host -foregroundcolor Green " Reference: "
Write-Host -foregroundcolor White " Configure OAuth authentication between Exchange and Exchange Online organizations `n "
Write-Host -foregroundcolor Yellow " https://technet.microsoft.com/en-us/library/dn594521(v=exchg.150).aspx" 
}
#endregion

cls
ShowParameters


do{
#do while not Y or N
Write-Host $bar
$ParamOK = Read-Host " Are this values correct? Pess Y for YES and N for NO"
$ParamOK = $ParamOK.ToUpper()
} while ($ParamOK -ne "Y" -AND $ParamOK -ne "N")

#cls
Write-Host $bar

if ($ParamOK -eq "N"){

UserOnlineCheck
ExchangeOnlineDomainCheck
UseronpremCheck
ExchangeOnPremDomainCheck
ExchangeOnPremEWSCheck
ExchangeOnPremLocalDomainCheck


}


# Free busy Lookup methods

$OrgRel = Get-OrganizationRelationship | Where{($_.DomainNames -like $ExchangeOnlineDomain )} | select Enabled,Identity,DomainNames,FreeBusy*,Target*
#$OrgRel
#$IntraOrgConEnabled = Get-IntraOrganizationConnector | select Enabled
$IntraOrgCon = Get-IntraOrganizationConnector | Select Name,TargetAddressDomains,DiscoveryEndpoint,Enabled

if ([string]::IsNullOrWhitespace($Auth))
{

Get-Sumary;

}

if($Auth -like "DAuth" -and $IntraOrgCon.enabled -Like "True")
{
Write-Host $bar
Write-Host -foregroundcolor yellow "  Warning: Intra Organization Connector is Enabled -> Free Busy Lookup is done using OAuth"
Write-Host $bar
}

#if ($nobrakes -ne "NB")
#{
#$nobrakes = Read-Host " Press Enter when ready to Grab Configuration Details. Ctrl+C to exit. Type NB for no Brakes"   
#Write-Host $bar
#}


#region DAutch Checks
if ($Auth -like "dauth" -OR [string]::IsNullOrWhitespace($Auth))

{


OrgRelCheck


if (!$ConfigurationOnly)
{
OrgRelSTV
}
Write-Host $spacebar
if ($nobrakes -ne "NB")
{
$nobrakes = Read-Host " Press Enter when ready to get Federation Information Details. Ctrl+C to exit. Type NB for no Brakes "   
Write-Host $bar
}

FedInfoCheck
if (!$ConfigurationOnly)
{
FedInfoSTV
}



if ($nobrakes -ne "NB")
{
$nobrakes = Read-Host " Press Enter when ready to get Federation Trust Details. Ctrl+C to exit. Type NB for no Brakes"   
Write-Host $bar
}

FedTrustCheck

if (!$ConfigurationOnly)
{
FedTrustSTV
}
Write-Host $bar
Write-Host -foregroundcolor Green " `n Test-FederationTrustCertificate" 
Write-Host $barspace
Test-FederationTrustCertificate


Write-Host $spacebar
if ($nobrakes -ne "NB")
{
$nobrakes = Read-Host " Press Enter when ready to check On-Prem Autodiscover Virtual Directory Details. Ctrl+C to exit. Type NB for no Brakes"   
Write-Host $bar
}   

AutoDVirtualDCheck

if (!$ConfigurationOnly)
{
AutoDVirtualDSTV

}


Write-Host $spacebar
if ($nobrakes -ne "NB")
{
$nobrakes = Read-Host " Press Enter when ready to Grab On-Prem Web Services Virtual Directory. Ctrl+C to exit. Type NB for no Brakes"   
Write-Host $bar
}

EWSVirtualDirectoryCheck

if (!$ConfigurationOnly)
{
EWSVirtualDirectorySTV

}


if ($nobrakes -ne "NB")
{
Write-Host $bar
$nobrakes = Read-Host "`n Press Enter when ready to  check the Availability Address Space configuration. Ctrl+C to exit. Type NB for no Brakes"   
Write-Host $bar
}
AvailabilityAddressSpaceCheck

If (!$ConfigurationOnly)
{
AvailabilityAddressSpaceSTV

}


if ($nobrakes -ne "NB")
{
Write-Host $bar
$nobrakes = Read-Host " Press Enter when ready to Test-FederationTrust. Ctrl+C to exit. Type NB for no Brakes"   
Write-Host $bar
}
#need to grab errors and provide alerts in error case 
TestFedTrust




if ($nobrakes -ne "NB")
{
Write-Host $bar
$nobrakes = Read-Host " Press Enter when ready to Test the OrganizationRelationship. Ctrl+C to exit. Type NB for no Brakes"   
Write-Host $bar
}
TestOrgRel


}

#endregion


#region OAuth Check

if ($Auth -like "OAuth" -OR [string]::IsNullOrWhitespace($Auth))

{
if ($nobrakes -ne "NB")
{
$nobrakes = Read-Host " Press Enter when ready to Grab OAuth Configuration Details. Ctrl+C to exit. Type NB for no Brakes"   
Write-Host $bar
}

Write-Host -foregroundcolor Green " `n `n ************************************TestingOAuth configuration************************************************* `n `n " 
Write-Host $spacebar




IntraOrgConCheck

#Get-IntraOrganizationConnector | Select Name,TargetAddressDomains,DiscoveryEndpoint,Enabled

If (!$ConfigurationOnly)
{

IntraOrgConSTV


}


Write-Host $spacebar




if ($nobrakes -ne "NB")
{
$nobrakes = Read-Host " Press Enter when ready to Grab the Auth Server Details. Ctrl+C to exit. Type NB for no Brakes"   
Write-Host $bar
}


 


AuthServerCheck
If (!$ConfigurationOnly)
{

AuthServerSTV



}


Write-Host $spacebar

if ($nobrakes -ne "NB")
{
$nobrakes = Read-Host " Press Enter when ready to Grab the Partner Application Details. Ctrl+C to exit. Type NB for no Brakes"   
Write-Host $bar
}


PartnerApplicationCheck
If (!$ConfigurationOnly)
{
PartnerApplicationSTV
}
Write-Host $bar
if ($nobrakes -ne "NB")
{
$nobrakes = Read-Host " Press Enter when ready to Check the Exchange Online-ApplicationAccount. Ctrl+C to exit. Type NB for no Brakes"   
Write-Host $bar
}
  


ApplicationAccounCheck
If (!$ConfigurationOnly)
{
ApplicationAccountSTV
}

Write-Host $bar
if ($nobrakes -ne "NB")
{
$nobrakes = Read-Host " Press Enter when ready to check the ManagementRoleAssignment of the Exchange Online-ApplicationAccount . Ctrl+C to exit. Type NB for no Brakes"   
Write-Host $bar
}
  
Write-Host -foregroundcolor Green "Get-ManagementRoleAssignment -RoleAssignee "Exchange Online-ApplicationAccount"  | ft Name,Role -AutoSize" 
Write-Host $bar
$ManagementRoleAssignment = Get-ManagementRoleAssignment -RoleAssignee "Exchange Online-ApplicationAccount"  | ft Name,Role -AutoSize
$ManagementRoleAssignment
If (!$ConfigurationOnly)
{



ManagementRoleAssignmentSTV


}
Write-Host $bar
if ($nobrakes -ne "NB")
{
$nobrakes = Read-Host " Press Enter when ready to Grab Auth config Details. Ctrl+C to exit. Type NB for no Brakes"   
Write-Host $bar
}
  
Write-Host -foregroundcolor Green " Get-AuthConfig | fl *Thumbprint, ServiceName, Realm, Name" 
Write-Host $bar
$AuthConfig = Get-AuthConfig | fl *Thumbprint, ServiceName, Realm, Name
$AuthConfig


If (!$ConfigurationOnly)
{

AuthConfigSTV


}

Write-Host $bar
if ($nobrakes -ne "NB")
{
$nobrakes = Read-Host " Press Enter when ready to Grab information for the Auth Certificate. Ctrl+C to exit. Type NB for no Brakes"   
Write-Host $bar
}
$thumb = Get-AuthConfig | Select CurrentCertificateThumbprint  
Write-Host $bar
Write-Host -ForegroundColor Green "Get-ExchangeCertificate -Thumbprint $thumb.CurrentCertificateThumbprint | fl" 
Write-Host $barspace
$CurrentCertificate = get-exchangecertificate $thumb.CurrentCertificateThumbprint | fl
$CurrentCertificate
If (!$ConfigurationOnly)
{

CurrentCertificateThumbprintSTV

}
Write-Host $bar
if ($nobrakes -ne "NB")
{
$nobrakes = Read-Host " Press Enter when ready to  check the On Prem Autodiscover Virtual Directory Configuration. Ctrl+C to exit. Type NB for no Brakes"   
Write-Host $bar
}

AutoDVirtualDCheckOAuth

$AutoDiscoveryVirtualDirectoryOAuth


If (!$ConfigurationOnly)
{
AutoDVirtualDSTV

}
Write-Host $bar
if ($nobrakes -ne "NB")
{
$nobrakes = Read-Host " Press Enter when ready to On-Prem Web Services Virtual Directory. Ctrl+C to exit. Type NB for no Brakes"   
Write-Host $bar
}

EWSVirtualDirectoryCheckOAuth

If (!$ConfigurationOnly)
{
EWSVirtualDirectorySTV
}
Write-Host $bar
if ($nobrakes -ne "NB")
{
$nobrakes = Read-Host " Press Enter when ready to Grab AvailabilityAddressSpace. Ctrl+C to exit. Type NB for no Brakes"   
Write-Host $bar
}

Write-Host $bar

AvailabilityAddressSpaceCheckOAuth

If (!$ConfigurationOnly)
{


AvailabilityAddressSpaceSTV


}
Write-Host $bar
if ($nobrakes -ne "NB")
{
$nobrakes = Read-Host " Press Enter when ready to test the Test-OAuthConnectivity. Ctrl+C to exit. Type NB for no Brakes"   
Write-Host $bar
}

Write-Host -foregroundcolor Green " Test-OAuthConnectivity -Service EWS -TargetUri https://outlook.office365.com/EWS/Exchange.asmx -Mailbox $useronprem | fl" 
Write-Host $barspace


$OAuthConnectivity = Test-OAuthConnectivity -Service EWS -TargetUri https://outlook.office365.com/EWS/Exchange.asmx -Mailbox $useronprem | fl
$OAuthConnectivity


If (!$ConfigurationOnly)
{


OAuthConnectivitySTV

}
Write-Host $bar
}

#endregion

Write-Host -foregroundcolor Green " `n That is all for the On Premise Side `n " 
#$exo = Read-Host "Ctrl+C to exit. Enter to Exit." 
#$exo = Read-Host "Ctrl+C to exit. Enter to Exit.
# Type EXO to check the configuration on the Exchange online Side
# Type HC to run the HealthChecker Script.
 
# Make sure you have FreeBusyInfo_EXP.ps1 and Healthchecker.ps1 files in the same folder where this script resides."   



if ($exo -eq "EXO"){
Write-Host $bar
write-host -ForegroundColor Blue "Running Exchange Online Free Busy Configuration Graber"
Write-Host $bar
start powershell {.\FreeBusyInfo_EXO.ps1}


$exo = Read-Host "Type HC to run the HealthChecker Script. Ctrl+C to exit. Enter to Exit. 
 
 To run the HealthChecker please make sure you have the Healthchecker.ps1 file in the same folder where this script resides."   

 if ($exo -eq "HC"){
Write-Host $bar
write-host -ForegroundColor Green "Running the HealthChecker Script"
Write-Host $bar
.\HealthChecker.ps1

$exo = Read-Host " Ctrl+C to exit. Enter to Exit." 
 }

}
Write-Host $barspace



if ($exo -eq "HC"){
Write-Host $bar
write-host -ForegroundColor Green "Running the HealthChecker Script"
Write-Host $bar
.\HealthChecker.ps1

$exo = Read-Host " Type EXO to run the Free Busy configuration checker on Exchange Online. 
Make sure you have FreeBusyInfo_EXO.ps1 file in the same folder where this script resides.

Ctrl+C to exit. Enter to Exit."  
  
if ($exo -eq "EXO"){   

Write-Host $bar
start powershell {.\FreeBusyInfo_EXO.ps1}

}

}
stop-transcript
Read-Host " `n `n Ctrl+C to exit. Enter to Exit."
