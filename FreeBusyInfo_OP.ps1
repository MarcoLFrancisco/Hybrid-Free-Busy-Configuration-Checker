 #Exchange on Premise

#>


#region Properties and Parameters
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', '', Justification = 'Variables are being used')]
[CmdletBinding(DefaultParameterSetName = "FreeBusyInfo_OP", SupportsShouldProcess)]
param(
    [Parameter(Mandatory = $false, ParameterSetName = "Auth")]
    [string]$Auth,
    [switch]$ConfigurationOnly,
    [string]$pause,
    [string]$Organization
)


Set-ExecutionPolicy " Unrestricted"  -Scope Process -Confirm:$false
Set-ExecutionPolicy " Unrestricted"  -Scope CurrentUser -Confirm:$false
Add-PSSnapin microsoft.exchange.management.powershell.snapin
import-module ActiveDirectory 
cls
$countOrgRelIssues = (0)
$Global:FedTrust = $null
$Global:AutoDiscoveryVirtualDirectory = $null
$Global:OrgRel
$Global:SPDomainsOnprem
$AvailabilityAddressSpace = $null
$Global:WebServicesVirtualDirectory = $null
$bar = " ==================================================================================================================" 
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
$Global:ExchangeOnPremEWS = ($EWSVirtualDirectory.externalURL.AbsoluteUri)[0]
$ADDomain=Get-ADDomain
$ExchangeOnPremLocalDomain=$ADDomain.forest
if ([string]::IsNullOrWhitespace($ADDomain)){
$ExchangeOnPremLocalDomain = $exchangeOnPremDomain
}


#endregion

#region Edit Parameters

Function UserOnlineCheck{
Write-Host -foregroundcolor Green " Online Mailbox: $UserOnline" 
$UserOnlineCheck = Read-Host " Press the Enter key if OK or type an Exchange Online Email address and press the Enter key"
if (![string]::IsNullOrWhitespace($UserOnlineCheck))
{
    $script:UserOnline = $UserOnlineCheck
} 
}

Function ExchangeOnlineDomainCheck{

#$ExchangeOnlineDomain
Write-Host -foregroundcolor Green " Exchange Online Domain: $ExchangeOnlineDomain" 
$ExchangeOnlineDomaincheck = Read-Host " Press enter if OK or type in the Exchange Online Domain and press the Enter key."

if (![string]::IsNullOrWhitespace($ExchangeOnlineDomaincheck))
{
    $script:ExchangeOnlineDomain = $ExchangeOnlineDomainCheck
} 
}

Function UseronpremCheck {
Write-Host -foregroundcolor Green " On Premises Hybrid Mailbox: $Useronprem" 
$Useronpremcheck = Read-Host " Press Enter if OK or type in an Exchange OnPremises Hybrid email address and press the Enter key."

if (![string]::IsNullOrWhitespace($Useronpremcheck))
{
    $script:Useronprem = $Useronpremcheck
} 
}

Function ExchangeOnPremDomainCheck {
#$exchangeOnPremDomain
Write-Host -foregroundcolor Green " On Premises Mail Domain: $exchangeOnPremDomain" 
$exchangeOnPremDomaincheck = Read-Host " Press enter if OK or type in the Exchange On Premises Mail Domain and press the Enter key."

if (![string]::IsNullOrWhitespace($exchangeOnPremDomaincheck))
{
    $script:exchangeOnPremDomain = $exchangeOnPremDomaincheck
} 

}

Function ExchangeOnPremEWSCheck{

Write-Host -foregroundcolor Green " On Premises EWS External URL: $exchangeOnPremEWS" 

$exchangeOnPremEWScheck = Read-Host " Press enter if OK or type in the Exchange On Premises EWS URL and press the Enter key."

if (![string]::IsNullOrWhitespace($exchangeOnPremEWScheck))
{
   $exchangeOnPremEWS = $exchangeOnPremEWScheck
} 
}

Function ExchangeOnPremLocalDomainCheck{
Write-Host -foregroundcolor Green " On Premises Root Domain: $exchangeOnPremLocalDomain  " 

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
write-host -foregroundcolor Green " Loading modules for AD, Exchange" 

Write-Host $bar
Write-Host   "  Color Scheme"
Write-Host $bar
Write-Host -ForegroundColor Red "  Look out for Red!"
Write-Host -ForegroundColor Yellow "  Yellow - Example information or Links"
Write-Host -ForegroundColor Green "  Green - In SUMMARY Sections it means OK. Anywhere else it's just a visual aid."
Write-Host $bar
Write-Host   "  Parameters:"
Write-Host $bar
Write-Host " Log File Path:" 

Write-Host -foregroundcolor Green "  $PSScriptRoot\$Logfile"
Write-Host " Office 365 Domain:"
Write-Host -foregroundcolor Green "  $exchangeOnlineDomain"
Write-Host " AD root Domain"
Write-Host -foregroundcolor Green "  $exchangeOnPremLocalDomain"
Write-Host -foregroundcolor White " Exchange On Premises Domain:  "
Write-Host -foregroundcolor Green "  $exchangeOnPremDomain"
Write-Host " Exchange On Premises External EWS url:"
Write-Host -foregroundcolor Green "  $exchangeOnPremEWS"
Write-Host " On Premises Hybrid Mailbox:"
Write-Host -foregroundcolor Green "  $useronprem"
Write-Host " Exchange Online Mailbox:"
Write-Host -foregroundcolor Green "  $userOnline"
}

#endregion

#regionDAuth Functions
Function OrgRelCheck (){
Write-Host $bar
Write-Host -foregroundcolor Green " Get-OrganizationRelationship  | Where{($_.DomainNames -like $ExchangeOnlineDomain )} | Select Identity,DomainNames,FreeBusy*,Target*,Enabled" 
Write-Host $bar
$OrgRel
Write-Host $bar
Write-Host  -foregroundcolor Green " SUMMARY - Organization Relationship (non standard values will show up in Red. Standard Values in Green)" 
Write-Host $bar
#$exchangeonlinedomain
Write-Host  " Domain Names:" 
if ($orgrel.DomainNames -like $exchangeonlinedomain){
Write-Host -foregroundcolor Green "  Domain Names Include the $exchangeOnlineDomain Domain" 
}
else
{
Write-Host -foregroundcolor Red "  Domain Names do Not Include the $exchangeOnlineDomain Domain" 

}
#FreeBusyAccessEnabled
Write-Host  " FreeBusyAccessEnabled:" 
if ($OrgRel.FreeBusyAccessEnabled -like "True" ){
Write-Host -foregroundcolor Green "  FreeBusyAccessEnabled is set to True" 
}
else
{
Write-Host -foregroundcolor Red "  FreeBusyAccessEnabled : False" 
$countOrgRelIssues++
}
#FreeBusyAccessLevel
Write-Host  " FreeBusyAccessLevel:" 
if ($OrgRel.FreeBusyAccessLevel -like "AvailabilityOnly" ){
Write-Host -foregroundcolor Green "  FreeBusyAccessLevel is set to AvailabilityOnly" 
}
if ($OrgRel.FreeBusyAccessLevel -like "LimitedDetails" ){
Write-Host -foregroundcolor Green "  FreeBusyAccessLevel is set to LimitedDetails" 
}
else
{
Write-Host -foregroundcolor Red "  FreeBusyAccessEnabled : False" 
$countOrgRelIssues++
}
#TargetApplicationUri
Write-Host  " TargetApplicationUri:" 
 if ($OrgRel.TargetApplicationUri -like "Outlook.com" ){
Write-Host -foregroundcolor Green "  TargetApplicationUri is Outlook.com" 
}
else
{
Write-Host -foregroundcolor Red "  TargetApplicationUri should be Outlook.com" 
$countOrgRelIssues++
}
#TargetOwaURL
Write-Host  " TargetOwaURL:" 
if ($OrgRel.TargetOwaURL -like "http://outlook.com/owa/$exchangeonpremdomain"){
Write-Host -foregroundcolor Green "  TargetOwaURL is http://outlook.com/owa/$exchangeonpremdomain" 
}
else
{
Write-Host -foregroundcolor Red "  TargetOwaURL IS NOT http://outlook.com/owa/$exchangeonpremdomain" 
$countOrgRelIssues++
}
#TargetSharingEpr
Write-Host  " TargetSharingEpr:" 
if ([string]::IsNullOrWhitespace($OrgRel.TargetSharingEpr)){
Write-Host -foregroundcolor Green "  TargetSharingEpr is blank. this is the standard Value." 
}
else
{
Write-Host -foregroundcolor Red "  TargetSharingEpr should be blank. If it is set, it should be Office 365 EWS endpoint" 
$countOrgRelIssues++
}
#TargetAutodiscoverEpr:
Write-Host  " TargetAutodiscoverEpr:" 
if ($OrgRel.TargetAutodiscoverEpr -like "https://autodiscover-s.outlook.com/autodiscover/autodiscover.svc/WSSecurity" ){
Write-Host -foregroundcolor Green "  TargetAutodiscoverEpr is correct" 
}
else
{
Write-Host -foregroundcolor Red "  TargetAutodiscoverEpr is not correct" 
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
#if ($countOrgRelIssues -eq '0'){
#Write-Host -foregroundcolor Green " Configurations Seem Correct" 
#}
#else
#{
#Write-Host -foregroundcolor Red "  Configurations DO NOT Seem Correct" 
#}

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
Write-Host  "`Domain Names: "
#if not null
Write-Host -foregroundcolor Green " "$fedinfo.DomainNames
}
else{
Write-Host " Domain Names: "
Write-Host -foregroundcolor Red " "$fedinfo.DomainNames
Write-Host -foregroundcolor Red  " DomainNames should contain $ExchangeOnlineDomain"
}

#TargetAutodiscoverEpr
if ($OrgRel.TargetAutodiscoverEpr -like $fedinfo.TargetAutodiscoverEpr){

Write-Host  " TargetAutodiscoverEpr: "
Write-Host -foregroundcolor Green " "$fedinfo.TargetAutodiscoverEpr
Write-Host  " Federation Information TargetAutodiscoverEpr vs Organization Relationship TargetAutodiscoverEpr " 
Write-Host -foregroundcolor Green "  Federation Information TargetAutodiscoverEpr matches the Organization Relationship TargetAutodiscoverEpr " 
}
else
{
Write-Host  " Federation Information TargetAutodiscoverEpr vs Organization Relationship TargetAutodiscoverEpr " 
Write-Host -foregroundcolor Red " =>  Federation Information TargetAutodiscoverEpr DOES NOT MATCH the Organization Relationship TargetAutodiscoverEpr" 

Write-Host " Organization Relationship:"  $OrgRel.TargetAutodiscoverEpr
Write-Host " Federation Information:   "  $fedinfo.TargetAutodiscoverEpr
}

#TokenIssuerUris
if ($fedinfo.TokenIssuerUris -like "*urn:federation:MicrosoftOnline*"){
Write-Host  " TokenIssuerUris: "
Write-Host -foregroundcolor Green " "  $fedinfo.TokenIssuerUris
}
else{
Write-Host " TokenIssuerUris: " $fedinfo.TokenIssuerUris
Write-Host  -foregroundcolor Red " TokenIssuerUris should be urn:federation:MicrosoftOnline"
}




Write-Host $bar
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
}
else
{ 
Write-Host -foregroundcolor Red "  Federation Trust Aplication Uri is NOT correct. "
Write-Host -foregroundcolor White "  Should be "$fedtrust.ApplicationUri
}
#$fedtrust.TokenIssuerUri.AbsoluteUri
Write-Host -foregroundcolor White " TokenIssuerUri:"
if ($fedtrust.TokenIssuerUri.AbsoluteUri -like  "urn:federation:MicrosoftOnline"){
#Write-Host -foregroundcolor White "  TokenIssuerUri:"
Write-Host -foregroundcolor Green " "$fedtrust.TokenIssuerUri.AbsoluteUri
}
else
{
Write-Host -foregroundcolor Red " Federation Trust TokenIssuerUri is NOT urn:federation:MicrosoftOnline"
}
Write-Host -foregroundcolor White " Federation Trust Certificate Expiracy:"
if ($fedtrust.OrgCertificate.NotAfter.Date -gt $CurrentTime){
Write-Host -foregroundcolor Green "  Not Expired - Expires on " $fedtrust.OrgCertificate.NotAfter.DateTime
}
else
{
Write-Host -foregroundcolor Red " Federation Trust Certificate is Expired on " $fedtrust.OrgCertificate.NotAfter.DateTime
}
Write-Host -foregroundcolor White " `Federation Trust Token Issuer Certificate Expiracy:"
if ($fedtrust.OrgCertificate.NotAfter.Date -gt $CurrentTime){
Write-Host -foregroundcolor Green "  Federation Trust TokenIssuerCertificate Not Expired - Expires on " $fedtrust.TokenIssuerCertificate.NotAfter.DateTime
}
else
{
Write-Host -foregroundcolor Red "  Federation Trust TokenIssuerCertificate Expired on " $fedtrust.TokenIssuerCertificate.NotAfter.DateTime
}
Write-Host -foregroundcolor White " Federation Trust Token Issuer Prev Certificate Expiracy:"
if ($fedtrust.TokenIssuerPrevCertificate.NotAfter.Date -gt $CurrentTime){
Write-Host -foregroundcolor Green " Federation Trust TokenIssuerPrevCertificate Not Expired - Expires on " $fedtrust.TokenIssuerPrevCertificate.NotAfter.DateTime 
}
else
{ 
Write-Host -foregroundcolor Red "  Federation Trust TokenIssuerPrevCertificate Expired on " $fedtrust.TokenIssuerPrevCertificate.NotAfter.DateTime

}
$fedtrustTokenIssuerMetadataEpr = "https://nexus.microsoftonline-p.com/FederationMetadata/2006-12/FederationMetadata.xml"
Write-Host -foregroundcolor White " `Token Issuer Metadata EPR:"
if ($fedtrust.TokenIssuerMetadataEpr.AbsoluteUri -like $fedtrustTokenIssuerMetadataEpr){
Write-Host -foregroundcolor Green "  Token Issuer Metadata EPR is " $fedtrust.TokenIssuerMetadataEpr.AbsoluteUri
}
else
{
Write-Host -foregroundcolor Red " Token Issuer Metadata EPR is Not " $fedtrust.TokenIssuerMetadataEpr.AbsoluteUri
}
$fedtrustTokenIssuerEpr = "https://login.microsoftonline.com/extSTS.srf"
Write-Host -foregroundcolor White " Token Issuer EPR:"
if ($fedtrust.TokenIssuerEpr.AbsoluteUri -like $fedtrustTokenIssuerEpr){
Write-Host -foregroundcolor Green "  Token Issuer EPR is:" $fedtrust.TokenIssuerEpr.AbsoluteUri
}
else
{
Write-Host -foregroundcolor Red "  Token Issuer EPR is Not:" $fedtrust.TokenIssuerEpr.AbsoluteUri
}
}

Function AutoDVirtualDCheck{
Write-Host -foregroundcolor Green " Get-AutodiscoverVirtualDirectory | Select Identity,Name,ExchangeVersion,*authentication*" 
Write-Host $bar
$Global:AutoDiscoveryVirtualDirectory = Get-AutodiscoverVirtualDirectory | Select Identity,Name,ExchangeVersion,*authentication*  
#Check if null or set
#$AutoDiscoveryVirtualDirectory
$Global:AutoDiscoveryVirtualDirectory
Write-Host $bar
Write-Host -foregroundcolor Green " SUMMARY - On-Prem Autodiscover Virtual Directory" 
Write-Host $bar
if ($Global:AutoDiscoveryVirtualDirectory.WSSecurityAuthentication -eq "True"){
foreach( $ser in $Global:AutoDiscoveryVirtualDirectory) { 
Write-Host " $($ser.Identity) "
Write-Host -ForegroundColor Green "  WSSecurityAuthentication: $($ser.WSSecurityAuthentication)" 
}
}
else
{
Write-Host -foregroundcolor Red " WSSecurityAuthentication is NOT correct."
foreach( $ser in $Global:AutoDiscoveryVirtualDirectory) { 
Write-Host " $($ser.Identity)"
Write-Host -ForegroundColor Red "  WSSecurityAuthentication: $($ser.WSSecurityAuthentication)" 
}

Write-Host -foregroundcolor White "  Should be True "
}

}

Function EWSVirtualDirectoryCheck{
Write-Host -foregroundcolor Green " Get-WebServicesVirtualDirectory | Select Identity,Name,ExchangeVersion,*Authentication*,*url" 
Write-Host $bar
$Global:WebServicesVirtualDirectory = Get-WebServicesVirtualDirectory | Select Identity,Name,ExchangeVersion,*Authentication*,*url
$Global:WebServicesVirtualDirectory
Write-Host $bar
Write-Host -foregroundcolor Green " SUMMARY - On-Prem Web Services Virtual Directory" 
Write-Host $bar
#Write-Host -foregroundcolor White "  WSSecurityAuthentication: `n " 
if ($Global:WebServicesVirtualDirectory.WSSecurityAuthentication -like  "True"){

foreach( $EWS in $Global:WebServicesVirtualDirectory) { 
Write-Host " $($EWS.Identity)"
Write-Host -ForegroundColor Green "  WSSecurityAuthentication: $($EWS.WSSecurityAuthentication) " 
}

}
else
{
Write-Host -foregroundcolor Red " WSSecurityAuthentication is NOT correct."
foreach( $EWS in $Global:AutoDiscoveryVirtualDirectory) { 
Write-Host " $($EWS.Identity) "
Write-Host -ForegroundColor Red "  WSSecurityAuthentication: $($ser.WSSecurityAuthentication) " 
}
Write-Host -foregroundcolor White "  Should be True"
}

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
Write-Host $bar
Write-Host -foregroundcolor Green " SUMMARY - On-Prem Availability Address Space" 
Write-Host $bar
Write-Host -foregroundcolor White " ForestName: " 
if ($AvailabilityAddressSpace.ForestName -like  $ExchangeOnlineDomain){
Write-Host -foregroundcolor Green " " $AvailabilityAddressSpace.ForestName
}
else
{
Write-Host -foregroundcolor Red "  ForestName is NOT correct."
Write-Host -foregroundcolor White " Should be $ExchaneOnlineDomain "
}
Write-Host -foregroundcolor White " UserName: " 
if ($AvailabilityAddressSpace.UserName -like  ""){
Write-Host -foregroundcolor Green "  Blank" 
}
else
{
Write-Host -foregroundcolor Red " UserName is NOT correct. "
Write-Host -foregroundcolor White "  Should be blank"
}
Write-Host -foregroundcolor White " UseServiceAccount: " 
if ($AvailabilityAddressSpace.UseServiceAccount -like  "True"){ 
Write-Host -foregroundcolor Green "  True"  
}
else
{
Write-Host -foregroundcolor Red "  UseServiceAccount is NOT correct."
Write-Host -foregroundcolor White "  Should be True"
}
Write-Host -foregroundcolor White " AccessMethod:" 
if ($AvailabilityAddressSpace.AccessMethod -like  "InternalProxy"){
Write-Host -foregroundcolor Green "  InternalProxy" 
}
else
{
Write-Host -foregroundcolor Red " AccessMethod is NOT correct."
Write-Host -foregroundcolor White " Should be InternalProxy"
}
Write-Host -foregroundcolor White " ProxyUrl: " 
if ($AvailabilityAddressSpace.ProxyUrl -like  $exchangeOnPremEWS){
Write-Host -foregroundcolor Green " "$AvailabilityAddressSpace.ProxyUrl 
}
else
{
Write-Host -foregroundcolor Red "  ProxyUrl is NOT correct."
Write-Host -foregroundcolor White "  Should be $exchangeOnPremEWS"
}

}

Function TestFedTrust{
Write-Host $bar
$TestFedTrustFail = 0
$a = Test-FederationTrust -UserIdentity $useronprem -verbose  #fails the frist time on multiple ocasions
Write-Host -foregroundcolor Green  " Test-FederationTrust -UserIdentity $useronprem -verbose"  
Write-Host $bar
$TestFedTrust = Test-FederationTrust -UserIdentity $useronprem -verbose -ErrorAction silentlycontinue
$TestFedTrust


foreach( $test in $TestFedTrust.type) { 

#$test
if ($test -ne "Success"){
Write-Host -foregroundcolor Red " $($test.Type)  "
Write-Host " $($test.Message)  "
$TestFedTrustFail++
}




}

if ($TestFedTrustFail -eq  0){
Write-Host -foregroundcolor Green " Federation Trust Successfully tested"
}
else  {

Write-Host -foregroundcolor Red " Federation Trust test with Errors"
#Check this an that

}



#Write-Host $bar
}

Function TestOrgRel{
$bar
$TestFail = 0
Write-Host -foregroundcolor Green "Test-OrganizationRelationship -Identity "On-premises to O365*"  -UserIdentity $useronprem" 
#need to grab errors and provide alerts in error case 
Write-Host $bar
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
#Write-Host -foregroundcolor Green " Organization Relationship Successfully tested `n "
}
else  {

#Write-Host -foregroundcolor Red " Organization Relationship test with Errors `n "
#Check this an that

}

Write-Host $bar

}
#endregion

#region OAuth Functions


Function IntraOrgConCheck{

Write-Host -foregroundcolor Green " Get-IntraOrganizationConnector | Selecct Name,TargetAddressDomains,DiscoveryEndpoint,Enabled" 
Write-Host $bar
#$IntraOrgConCheck = Get-IntraOrganizationConnector | fl Name,TargetAddressDomains,DiscoveryEndpoint,Enabled
#$IntraOrgConCheck
$IOC=$IntraOrgCon | fl
$IOC
Write-Host $bar
Write-Host -foregroundcolor Green " SUMMARY - On-Prem Intra Organization Connector" 
Write-Host $bar
$IntraOrgTargetAddressDomain = $IntraOrgCon.TargetAddressDomains.Domain
$IntraOrgTargetAddressDomain = $IntraOrgTargetAddressDomain.Tolower()
Write-Host -foregroundcolor White " Target Address Domains: " 
if ($IntraOrgCon.TargetAddressDomains-like  "*$ExchangeOnlineDomain*" -Or $IntraOrgCon.TargetAddressDomains -like  "*$ExchangeOnlineAltDomain*" ){
Write-Host -foregroundcolor Green " " $IntraOrgCon.TargetAddressDomains
}
else
{
Write-Host -foregroundcolor Red " Target Address Domains is NOT correct."
Write-Host -foregroundcolor White " Should contain the $ExchangeOnlineDomain domain or the $ExchangeOnlineAltDomain"
}

Write-Host -foregroundcolor White " DiscoveryEndpoint: " 
if ($IntraOrgCon.DiscoveryEndpoint -like  "https://autodiscover-s.outlook.com/autodiscover/autodiscover.svc"){
Write-Host -foregroundcolor Green "  https://autodiscover-s.outlook.com/autodiscover/autodiscover.svc" 
}
else
{
Write-Host -foregroundcolor Red " `DiscoveryEndpoint are NOT correct. "
Write-Host -foregroundcolor White "  Should be https://autodiscover-s.outlook.com/autodiscover/autodiscover.svc"
}
Write-Host -foregroundcolor White " Enabled: " 
if ($IntraOrgCon.Enabled -like  "True"){ 
Write-Host -foregroundcolor Green "  True "  
}
else
{
Write-Host -foregroundcolor Red "  Enabled is NOT correct."
Write-Host -foregroundcolor White " Should be True"
}




}

Function AuthServerCheck{

#Write-Host $bar
Write-Host -foregroundcolor Green " Get-AuthServer | Select Name,IssuerIdentifier,TokenIssuingEndpoint,AuthMetadataUrl,Enabled" 
Write-Host $bar
$AuthServer = Get-AuthServer | Where {$_.Name -like "ACS*"} | Select Name,IssuerIdentifier,TokenIssuingEndpoint,AuthMetadataUrl,Enabled
$AuthServer

Write-Host $bar
Write-Host -foregroundcolor Green " SUMMARY - Auth Server" 
Write-Host $bar

Write-Host -foregroundcolor White " IssuerIdentifier: " 
if ($AuthServer.IssuerIdentifier -like  "00000001-0000-0000-c000-000000000000" ){
Write-Host -foregroundcolor Green " " $AuthServer.IssuerIdentifier
}
else
{
Write-Host -foregroundcolor Red " IssuerIdentifier is NOT correct."
Write-Host -foregroundcolor White " Should be 00000001-0000-0000-c000-000000000000"
}

Write-Host -foregroundcolor White " TokenIssuingEndpoint: " 
if ($AuthServer.TokenIssuingEndpoint -like  "https://accounts.accesscontrol.windows.net/*"  -and $AuthServer.TokenIssuingEndpoint -like  "*/tokens/OAuth/2" ){
Write-Host -foregroundcolor Green " " $AuthServer.TokenIssuingEndpoint
}
else
{
Write-Host -foregroundcolor Red " TokenIssuingEndpoint is NOT correct."
Write-Host -foregroundcolor White " Should be  https://accounts.accesscontrol.windows.net/<Cloud Tenant ID>/tokens/OAuth/2"
}


Write-Host -foregroundcolor White " AuthMetadataUrl: " 
if ($AuthServer.AuthMetadataUrl -like  "https://accounts.accesscontrol.windows.net/*"  -and $AuthServer.TokenIssuingEndpoint -like  "*/tokens/OAuth/2" ){
Write-Host -foregroundcolor Green " " $AuthServer.AuthMetadataUrl
}
else
{
Write-Host -foregroundcolor Red " AuthMetadataUrl is NOT correct."
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

Function PartnerApplicationCheck{
Write-Host $bar
Write-Host -foregroundcolor Green " Get-PartnerApplication |  ?{$_.ApplicationIdentifier -eq '00000002-0000-0ff1-ce00-000000000000'
 -and $_.Realm -eq ''} | Select Enabled, ApplicationIdentifier, CertificateStrings, AuthMetadataUrl, Realm, UseAuthServer, 
 AcceptSecurityIdentifierInformation, LinkedAccount, IssuerIdentifier, AppOnlyPermissions, ActAsPermissions, Name" 
Write-Host $bar

$PartnerApplication = Get-PartnerApplication |  ?{$_.ApplicationIdentifier -eq '00000002-0000-0ff1-ce00-000000000000' -and $_.Realm -eq ''} | Select Enabled, ApplicationIdentifier, CertificateStrings, AuthMetadataUrl, Realm, UseAuthServer, AcceptSecurityIdentifierInformation, LinkedAccount, IssuerIdentifier, AppOnlyPermissions, ActAsPermissions, Name
$PartnerApplication
Write-Host $bar
Write-Host -foregroundcolor Green " SUMMARY - Partner Application" 
Write-Host $bar


Write-Host -foregroundcolor White " Enabled: " 
if ($PartnerApplication.Enabled -like  "True" ){
Write-Host -foregroundcolor Green " " $PartnerApplication.Enabled
}
else
{
Write-Host -foregroundcolor Red " Enabled: False "
Write-Host -foregroundcolor White " Should be True"
}

Write-Host -foregroundcolor White " ApplicationIdentifier: " 
if ($PartnerApplication.ApplicationIdentifier -like  "00000002-0000-0ff1-ce00-000000000000" ){
Write-Host -foregroundcolor Green " " $PartnerApplication.ApplicationIdentifier
}
else
{
Write-Host -foregroundcolor Red " ApplicationIdentifier does not seem correct"
Write-Host -foregroundcolor White " Should be 00000002-0000-0ff1-ce00-000000000000"
}

Write-Host -foregroundcolor White " AuthMetadataUrl: " 
if ([string]::IsNullOrWhitespace( $PartnerApplication.AuthMetadataUrl)){
Write-Host -foregroundcolor Green "  Blank" 
}
else
{
Write-Host -foregroundcolor Red " AuthMetadataUrl does not seem correct"
Write-Host -foregroundcolor White " Should be Blank"
}



Write-Host -foregroundcolor White " Realm: " 
if ([string]::IsNullOrWhitespace( $PartnerApplication.Realm)){
Write-Host -foregroundcolor Green "  Blank"
}
else
{
Write-Host -foregroundcolor Red "  Realm does not seem correct"
Write-Host -foregroundcolor White " Should be Blank"
}


Write-Host -foregroundcolor White " LinkedAccount: " 
if ($PartnerApplication.LinkedAccount -like  "$exchangeOnPremDomain/Users/Exchange Online-ApplicationAccount" -or $PartnerApplication.LinkedAccount -like  "$exchangeOnPremLocalDomain/Users/Exchange Online-ApplicationAccount"  ){
Write-Host -foregroundcolor Green " " $PartnerApplication.LinkedAccount
}
else
{
Write-Host -foregroundcolor Red " LinkedAccount value does not seem correct"
Write-Host -foregroundcolor White " Should be $exchangeOnPremLocalDomain/Users/Exchange Online-ApplicationAccount"
}
}

Function ApplicationAccounCheck{

Write-Host $bar
Write-Host -foregroundcolor Green " Get-user '$exchangeOnPremLocalDomain/Users/Exchange Online-ApplicationAccount' | Select Name, RecipientType, 
RecipientTypeDetails, UserAccountControl" 
Write-Host $bar
$ApplicationAccount = Get-user "$exchangeOnPremLocalDomain/Users/Exchange Online-ApplicationAccount"  | Select Name, RecipientType, RecipientTypeDetails, UserAccountControl
$ApplicationAccount

Write-Host $bar
Write-Host -foregroundcolor Green " SUMMARY - Application Account" 
Write-Host $bar

Write-Host -foregroundcolor White " RecipientType: " 
if ($ApplicationAccount.RecipientType -like  "User" ){
Write-Host -foregroundcolor Green " " $ApplicationAccount.RecipientType
}
else
{
Write-Host -foregroundcolor Red " RecipientType value is $ApplicationAccount.RecipientType "
Write-Host -foregroundcolor White " Should be User"
}

Write-Host -foregroundcolor White " RecipientTypeDetails: " 
if ($ApplicationAccount.RecipientTypeDetails -like  "LinkedUser" ){
Write-Host -foregroundcolor Green " " $ApplicationAccount.RecipientTypeDetails
}
else
{
Write-Host -foregroundcolor Red " RecipientTypeDetails value is $ApplicationAccount.RecipientTypeDetails"
Write-Host -foregroundcolor White " Should be LinkedUser"
}


Write-Host -foregroundcolor White " UserAccountControl: " 
if ($ApplicationAccount.UserAccountControl -like  "AccountDisabled, PasswordNotRequired, NormalAccount" ){
Write-Host -foregroundcolor Green " " $ApplicationAccount.UserAccountControl
}
else
{
Write-Host -foregroundcolor Red " UserAccountControl value does not seem correct"
Write-Host -foregroundcolor White " Should be AccountDisabled, PasswordNotRequired, NormalAccount"
}

}

Function ManagementRoleAssignmentCheck{

Write-Host -foregroundcolor Green " Get-ManagementRoleAssignment -RoleAssignee Exchange Online-ApplicationAccount | Select Name,Role -AutoSize" 
Write-Host $bar
$ManagementRoleAssignment = Get-ManagementRoleAssignment -RoleAssignee "Exchange Online-ApplicationAccount"  | Select Name,Role 
$M= $ManagementRoleAssignment | Out-String
$M
Write-Host $bar
Write-Host -foregroundcolor Green " SUMMARY - Management Role Assignment for the Exchange Online-ApplicationAccount (non standard values will show up in Red. Standard Values in Green)" 
Write-Host $bar
Write-Host -foregroundcolor White " Role: " 
if ($ManagementRoleAssignment.Role -like "*UserApplication*" ){
Write-Host -foregroundcolor Green "  UserApplication Role Assigned" 
}
else
{
Write-Host -foregroundcolor Red "  UserApplication Role not present for the Exchange Online-ApplicationAccount"

}
if ($ManagementRoleAssignment.Role -like "*ArchiveApplication*" ){
Write-Host -foregroundcolor Green "  ArchiveApplication Role Assigned" 
}
else
{
Write-Host -foregroundcolor Red "  ArchiveApplication Role not present for the Exchange Online-ApplicationAccount"

}

if ($ManagementRoleAssignment.Role -like "*LegalHoldApplication*" ){
Write-Host -foregroundcolor Green "  LegalHoldApplication Role Assigned" 
}
else
{
Write-Host -foregroundcolor Red "  LegalHoldApplication Role not present for the Exchange Online-ApplicationAccount"

}

if ($ManagementRoleAssignment.Role -like "*Mailbox Search*" ){
Write-Host -foregroundcolor Green "  Mailbox Search Role Assigned" 
}
else
{
Write-Host -foregroundcolor Red "  Mailbox Search Role not present for the Exchange Online-ApplicationAccount"

}


if ($ManagementRoleAssignment.Role -like "*TeamMailboxLifecycleApplication*" ){
Write-Host -foregroundcolor Green "  TeamMailboxLifecycleApplication Role Assigned" 
}
else
{
Write-Host -foregroundcolor Red "  TeamMailboxLifecycleApplication Role not present for the Exchange Online-ApplicationAccount"

}


if ($ManagementRoleAssignment.Role -like "*MailboxSearchApplication*" ){
Write-Host -foregroundcolor Green "  MailboxSearchApplication Role Assigned" 
}
else
{
Write-Host -foregroundcolor Red "  MailboxSearchApplication Role not present for the Exchange Online-ApplicationAccount"

}


if ($ManagementRoleAssignment.Role -like "*MeetingGraphApplication*" ){
Write-Host -foregroundcolor Green "  MeetingGraphApplication Role Assigned" 
}
else
{
Write-Host -foregroundcolor Red "  MeetingGraphApplication Role not present for the Exchange Online-ApplicationAccount"

}

}

Function AuthConfigCheck{

Write-Host -foregroundcolor Green " Get-AuthConfig | Select *Thumbprint, ServiceName, Realm, Name" 
Write-Host $bar
$AuthConfig = Get-AuthConfig | Select *Thumbprint, ServiceName, Realm, Name
$AC=$AuthConfig | fl
$AC
Write-Host $bar
Write-Host -foregroundcolor Green " SUMMARY - Auth Config (non standard values will show up in Red. Standard Values in Green)" 
Write-Host $bar
if (![string]::IsNullOrWhitespace($AuthConfig.CurrentCertificateThumbprint)) {
Write-HOst " Thumbprint: "$AuthConfig.CurrentCertificateThumbprint 
Write-Host -foregroundcolor Green " Certificate is Assigned"

}
else
{
Write-HOst " Thumbprint: "$AuthConfig.CurrentCertificateThumbprint 
Write-Host -foregroundcolor Red " No valid certificate Assigned "
}


if ($AuthConfig.ServiceName -like "00000002-0000-0ff1-ce00-000000000000" ){
Write-HOst " ServiceName: "$AuthConfig.ServiceName 
Write-Host -foregroundcolor Green " Service Name Seems correct" 
}
else
{
Write-HOst " ServiceName: "$AuthConfig.ServiceName
Write-Host -foregroundcolor Red " Service Name does not Seems correct. Should be 00000002-0000-0ff1-ce00-000000000000"

}


if ([string]::IsNullOrWhitespace($AuthConfig.Realm)) {
Write-HOst " Realm: "
Write-Host -foregroundcolor Green " Realm is Blank" 
}
else
{
Write-HOst " Realm: "$AuthConfig.Realm
Write-Host -foregroundcolor Red " Realm should be Blank"

}

}

Function CurrentCertificateThumbprintCheck{

$thumb = Get-AuthConfig | Select CurrentCertificateThumbprint  
$thumbprint = $thumb.currentcertificateThumbprint
#Write-Host $bar
Write-Host -ForegroundColor Green " Get-ExchangeCertificate -Thumbprint $thumbprint | Select *" 
Write-Host $bar
$CurrentCertificate = get-exchangecertificate $thumb.CurrentCertificateThumbprint | Select *
$CC = $CurrentCertificate | fl
$CC
Write-Host $bar
Write-Host -foregroundcolor Green " SUMMARY - Microsoft Exchange Server Auth Certificate (non standard values will show up in Red. Standard Values in Green)" 
Write-Host $bar
if ($CurrentCertificate.Issuer -like "CN=Microsoft Exchange Server Auth Certificate" ){
write-Host " Issuer: " $CurrentCertificate.Issuer
Write-Host -foregroundcolor Green "  Issuer is CN=Microsoft Exchange Server Auth Certificate" 
}
else
{
Write-Host -foregroundcolor Red "  Issuer is not CN=Microsoft Exchange Server Auth Certificate"
}
if ($CurrentCertificate.Services -like "SMTP" ){
Write-Host " Services: " $CurrentCertificate.Services
Write-Host -foregroundcolor Green "  Certificate enabled for SMTP" 
}
else
{
Write-Host -foregroundcolor Red "  Certificate Not enabled for SMTP"
}
if ($CurrentCertificate.Status -like "Valid" ){
Write-Host " Status: " $CurrentCertificate.Status
Write-Host -foregroundcolor Green "  Certificate is valid" 
}
else
{
Write-Host -foregroundcolor Red "  Certificate is not Valid"
}
if ($CurrentCertificate.Subject -like "CN=Microsoft Exchange Server Auth Certificate" ){
Write-Host " Subject: " $CurrentCertificate.Subject
Write-Host -foregroundcolor Green "  Subject is CN=Microsoft Exchange Server Auth Certificate" 
}
else
{
Write-Host -foregroundcolor Red "  Subject is not CN=Microsoft Exchange Server Auth Certificate"
}
}

Function AutoDVirtualDCheckOauth{
#Write-Host -foregroundcolor Green " `n On-Prem Autodiscover Virtual Directory `n "  
Write-Host -foregroundcolor Green " Get-AutodiscoverVirtualDirectory | Select Identity, Name,ExchangeVersion,*authentication*" 
Write-Host $bar
$AutoDiscoveryVirtualDirectoryOAuth = Get-AutodiscoverVirtualDirectory | Select Identity,Name,ExchangeVersion,*authentication*  
#Check if null or set
$AD=$AutoDiscoveryVirtualDirectoryOAuth | fl
$AD
if ($Auth -like "OAuth"){
}
Write-Host $bar
Write-Host -foregroundcolor Green " SUMMARY - On-Prem Autodiscover Virtual Directory" 
Write-Host $bar
if ($AutoDiscoveryVirtualDirectoryOAuth.WSSecurityAuthentication -like  "True"){
#Write-Host -foregroundcolor Green " `n  " $Global:AutoDiscoveryVirtualDirectory.WSSecurityAuthentication
foreach( $ADVD in $AutoDiscoveryVirtualDirectoryOAuth) { 
Write-Host " $($ADVD.Identity) "
Write-Host -ForegroundColor Green " WSSecurityAuthentication: $($ADVD.WSSecurityAuthentication)" 
}
}
else
{
Write-Host -foregroundcolor Red " WSSecurityAuthentication setting is NOT correct."
foreach( $ADVD in $AutoDiscoveryVirtualDirectoryOAuth) { 
Write-Host " $($ADVD.Identity) "
Write-Host -ForegroundColor Red " WSSecurityAuthentication: $($ADVD.WSSecurityAuthentication)" 
}
}
#Write-Host $bar
}

Function EWSVirtualDirectoryCheckOAuth{
Write-Host -foregroundcolor Green " Get-WebServicesVirtualDirectory | Select Identity,Name,ExchangeVersion,*Authentication*,*url" 
Write-Host $bar
$WebServicesVirtualDirectoryOAuth = Get-WebServicesVirtualDirectory | select Identity, Name,ExchangeVersion,*Authentication*,*url
$W= $WebServicesVirtualDirectoryOAuth | fl
$W
if ($Auth -like "OAuth"){
}
Write-Host $bar
Write-Host -foregroundcolor Green " SUMMARY - On-Prem Web Services Virtual Directory" 
Write-Host $bar
if ($WebServicesVirtualDirectoryOAuth.WSSecurityAuthentication -like  "True"){
foreach( $EWS in $WebServicesVirtualDirectoryOAuth) { 
Write-Host " $($EWS.Identity) "
Write-Host -ForegroundColor Green " WSSecurityAuthentication: $($EWS.WSSecurityAuthentication) " 
}

}
else
{
Write-Host -foregroundcolor Red " WSSecurityAuthentication is NOT correct."
foreach( $EWS in $WebServicesVirtualDirectoryOauth) { 
Write-Host " $($EWS.Identity) "
Write-Host -ForegroundColor Red " WSSecurityAuthentication: $($EWS.WSSecurityAuthentication)" 
}
Write-Host -foregroundcolor White "  Should be True"
}
#Write-Host $bar
}

Function AvailabilityAddressSpaceCheckOAuth{
Write-Host -foregroundcolor Green " Get-AvailabilityAddressSpace $exchangeOnlineDomain | Select ForestName, UserName, UseServiceAccount, AccessMethod, ProxyUrl, Name"
Write-Host $bar
$AvailabilityAddressSpace = Get-AvailabilityAddressSpace $exchangeOnlineDomain | Select ForestName,UserName,UseServiceAccount,AccessMethod,ProxyUrl,Name
$AAS = $AvailabilityAddressSpace | fl
$AAS

if ($Auth -like "OAuth"){
}
Write-Host $bar
Write-Host -foregroundcolor Green " SUMMARY - On-Prem Availability Address Space" 
Write-Host $bar
Write-Host -foregroundcolor White " ForestName: " 
if ($AvailabilityAddressSpace.ForestName -like  $ExchangeOnlineDomain){
Write-Host -foregroundcolor Green " "$AvailabilityAddressSpace.ForestName
}
else
{
Write-Host -foregroundcolor Red " ForestName is NOT correct. "
Write-Host -foregroundcolor White " Should be $ExchaneOnlineDomain "
}
Write-Host -foregroundcolor White " UserName: " 
if ($AvailabilityAddressSpace.UserName -like  ""){
Write-Host -foregroundcolor Green " Blank " 
}
else
{
Write-Host -foregroundcolor Red " UserName is NOT correct. "
Write-Host -foregroundcolor White "Should be blank "
}
Write-Host -foregroundcolor White " UseServiceAccount: " 
if ($AvailabilityAddressSpace.UseServiceAccount -like  "True"){ 
Write-Host -foregroundcolor Green " True "  
}
else
{
Write-Host -foregroundcolor Red " UseServiceAccount is NOT correct."
Write-Host -foregroundcolor White " Should be True "
}
Write-Host -foregroundcolor White " AccessMethod: " 
if ($AvailabilityAddressSpace.AccessMethod -like  "InternalProxy"){
Write-Host -foregroundcolor Green " InternalProxy " 
}
else
{
Write-Host -foregroundcolor Red " AccessMethod is NOT correct. "
Write-Host -foregroundcolor White " Should be InternalProxy "
}
Write-Host -foregroundcolor White " ProxyUrl: " 
if ($AvailabilityAddressSpace.ProxyUrl -like  $exchangeOnPremEWS){
Write-Host -foregroundcolor Green " "$AvailabilityAddressSpace.ProxyUrl 
}
else
{
Write-Host -foregroundcolor Red " ProxyUrl is NOT correct. "
Write-Host -foregroundcolor White " Should be $exchangeOnPremEWS"
}
#falta o ews
#Write-Host $bar
}

Function OAuthConnectivityCheck{


Write-Host -foregroundcolor Green " Test-OAuthConnectivity -Service EWS -TargetUri https://outlook.office365.com/EWS/Exchange.asmx -Mailbox $useronprem | fl" 
Write-Host $bar
#$OAuthConnectivity = Test-OAuthConnectivity -Service EWS -TargetUri https://outlook.office365.com/EWS/Exchange.asmx -Mailbox $useronprem | fl
#$OAuthConnectivity
$OAuthConnectivity = Test-OAuthConnectivity -Service EWS -TargetUri https://outlook.office365.com/EWS/Exchange.asmx -Mailbox $useronprem | Select *
$OAC = $OAuthConnectivity | fl
$OAC
$bar
$OAuthConnectivity.Detail.FullId
$bar
if ($OAuthConnectivity.Detail.FullId -like '*(401) Unauthorized*'){
write-host -ForegroundColor Red "Error: (401) Unauthorized"
if ($OAuthConnectivity.Detail.FullId -like 'The user specified by the user-context in the token does not exist*'){
write-host "The user specified by the user-context in the token does not exist"
write-host "Please run Test-OAuthConnectivity with a different Exchange On Premises Mailbox"

}
   
Write-Host $bar
#$OAuthConnectivity.detail.LocalizedString
Write-Host -foregroundcolor Green " SUMMARY - Test OAuth COnnectivity" 
Write-Host $bar
if ($OAuthConnectivity.ResultType -like  "Success"){
Write-Host -foregroundcolor Green " OAuth Test was completed successfully " 
}
else
{
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

#region ExoDauthFuntions (1 fix)

Function ExoOrgRelCheck (){
Write-Host $bar
Write-Host -foregroundcolor Green " Get-OrganizationRelationship  | Where{($_.DomainNames -like $ExchangeOnPremDomain )} | Select Identity,DomainNames,FreeBusy*,Target*,Enabled" 
Write-Host $bar
$ExoOrgRel
Write-Host $bar
Write-Host  -foregroundcolor Green " Summary - Organization Relationship" 
Write-Host $bar
Write-Host  " Domain Names:" 
if ($exoOrgRel.DmainNames -like $exchangeonpremdomain){
Write-Host -foregroundcolor Green "  Domain Names Include the $exchangeOnpremDomain Domain" 
}
else
{
Write-Host -foregroundcolor Red "  Domain Names do Not Include the $exchangeOnpremDomain Domain" 
}

#FreeBusyAccessEnabled

Write-Host  " FreeBusyAccessEnabled:" 
if ($exoOrgRel.FreeBusyAccessEnabled -like "True" ){
Write-Host -foregroundcolor Green "  FreeBusyAccessEnabled is set to True" 
}
else
{
Write-Host -foregroundcolor Red "  FreeBusyAccessEnabled : False" 
#$countOrgRelIssues++
}


#FreeBusyAccessLevel
Write-Host  " FreeBusyAccessLevel:" 
if ($exoOrgRel.FreeBusyAccessLevel -like "AvailabilityOnly" ){
Write-Host -foregroundcolor Green "  FreeBusyAccessLevel is set to AvailabilityOnly" 
}
if ($exoOrgRel.FreeBusyAccessLevel -like "LimitedDetails" ){
Write-Host -foregroundcolor Green "  FreeBusyAccessLevel is set to LimitedDetails" 
}

#fix porque este else só respeita o if anterior
if ($exoOrgRel.FreeBusyAccessLevel -NE "AvailabilityOnly" -AND $exoOrgRel.FreeBusyAccessLevel -NE "LimitedDetails")
{
Write-Host -foregroundcolor Red "  FreeBusyAccessEnabled : False" 
#$countOrgRelIssues++
}
#TargetApplicationUri
Write-Host  " TargetApplicationUri:" 

 if ($exoOrgRel.TargetApplicationUri -like $fedtrust.ApplicationUri){
Write-Host -foregroundcolor Green "  TargetApplicationUri is" $fedtrust.ApplicationUri.originalstring 
}
else
{
Write-Host -foregroundcolor Red "  TargetApplicationUri should be " $fedtrust.ApplicationUri.originalstring 
#$countOrgRelIssues++
}

#TargetSharingEpr
Write-Host  " TargetSharingEpr:" 
if ([string]::IsNullOrWhitespace($exoOrgRel.TargetSharingEpr)){
Write-Host -foregroundcolor Green "  TargetSharingEpr is blank. This is the standard Value." 
}
else
{
Write-Host -foregroundcolor Red "  TargetSharingEpr should be blank. If it is set, it should be the On-Premises Exchange servers EWS ExternalUrl endpoint." 
#$countOrgRelIssues++
}
#TargetAutodiscoverEpr:

Write-Host  " TargetAutodiscoverEpr:"
 
if ($exoOrgRel.TargetAutodiscoverEpr -like $exofedinfo.TargetAutodiscoverEpr){
Write-Host -foregroundcolor Green "  TargetAutodiscoverEpr is" $exofedinfo.TargetAutodiscoverEpr.OriginalString 
}
else
{
Write-Host -foregroundcolor Red "  TargetAutodiscoverEpr is not" $exofedinfo.TargetAutodiscoverEpr.OriginalString 
#$countOrgRelIssues++
}
#Enabled
Write-Host  " Enabled:" 
if ($exoOrgRel.enabled -like "True" ){
Write-Host -foregroundcolor Green "  Enabled is set to True" 
}
else
{
Write-Host -foregroundcolor Red "  Enabled is set to False." 
}
}

Function EXOFedOrgIdCheck{
Write-Host -foregroundcolor Green " Get-FederatedOrganizationIdentifier | select AccountNameSpace,Domains,Enabled" 
Write-Host $bar
$exoFedOrgId = Get-FederatedOrganizationIdentifier | select AccountNameSpace,Domains,Enabled
#$IntraOrgConCheck
$efedorgid=$exoFedOrgId | fl
$efedorgid
Write-Host $bar
Write-Host -foregroundcolor Green " SUMMARY - Online Federated Organization Identifier" 
Write-Host $bar
Write-Host -foregroundcolor White " Domains: " 
if ($exoFedOrgId.Domains -like  "*$ExchangeOnlineDomain*"){
Write-Host -foregroundcolor Green " " $exoFedOrgId.Domains
}
else
{
Write-Host -foregroundcolor Red " Domains are NOT correct."
Write-Host -foregroundcolor White " Should contain the $ExchangeOnlinemDomain"
}
Write-Host -foregroundcolor White " Enabled: " 
if ($exoFedOrgId.Enabled -like  "True"){ 
Write-Host -foregroundcolor Green "  True "  
}
else
{
Write-Host -foregroundcolor Red "  Enabled is NOT correct."
Write-Host -foregroundcolor White " Should be True"
}
}

Function EXOTestOrgRelCheck{

#fix 2
#troquei isto Write-Host -foregroundcolor Green " Test-OrganizationRelationship -Identity 'O365 to On-premises*' -UserIdentity $UserOnline" 
#por isto
Write-Host -foregroundcolor Green " Test-OrganizationRelationship -Identity $ExoOrgRel.Identity -UserIdentity $UserOnline"  
Write-Host $bar

#Write-Host -ForegroundColor Green $ExoOrgRel.Identity

#troquei isto $exotestorgrel= Test-OrganizationRelationship -Identity 'O365 to On-premises*' -UserIdentity $UserOnline
$exotestorgrel= Test-OrganizationRelationship -Identity $ExoOrgRel.Identity -UserIdentity $UserOnline
$exotor = $exotestorgrel | fl
$exotor
$bar
$exotestorgrel.Detail.FullId
#$bar
}

Function SharingPolicyCheck{

Write-Host -foregroundcolor Green " Get-SharingPolicy | select *" 
Write-Host $bar
$Script:SPOnline= Get-SharingPolicy | select *
$SPOnline
Write-Host $bar
Write-Host -foregroundcolor Green " SUMMARY - Sharing Policy" 
Write-Host $bar


Write-Host -foregroundcolor White " Exchange On Premises Sharing domains:"
Write-Host -foregroundcolor White " Domain:"
$SPOnprem.Domains.Domain[0]
Write-Host -foregroundcolor White " Action:"
$SPOnprem.Domains.Actions[0]
Write-Host -foregroundcolor White " Domain:"
$SPOnprem.Domains.Domain[1]
Write-Host -foregroundcolor White " Action:"
$SPOnprem.Domains.Actions[1]
Write-Host -ForegroundColor White "Exchange OnLine Sharing Domains"
$domain1=(($SPOnline.domains[0] -split ":") -split " ")
$domain2=(($SPOnline.domains[1] -split ":") -split " ")
Write-Host -foregroundcolor White "Domain:" 
Write-Host " " $domain1[0] 
Write-Host -foregroundcolor White "Action:" 
Write-Host " " $domain1[1]
Write-Host -foregroundcolor White "Domain:" 
Write-Host " " $domain2[0] 
Write-Host -foregroundcolor White "Action:" 
Write-Host " " $domain2[1]
Write-Host $bar


Write-Host -foregroundcolor White " Domains: " 
if ($SPDomainsOnline -like $SPDomainsOnprem){

Write-Host -foregroundcolor Green " " $SPDomainsOnline.Domains
}
else
{
Write-Host -foregroundcolor Red " Domains are NOT correct."
Write-Host -foregroundcolor White " Exchange Online Sharing domains: "
write-Host -foregroundcolor White "Domain:" 
Write-Host " " $domain1[0] 
Write-Host -foregroundcolor White "Action:" 
Write-Host " " $domain1[1]
Write-Host -foregroundcolor White "Domain:" 
Write-Host " " $domain2[0] 
Write-Host -foregroundcolor White "Action:" 
Write-Host " " $domain2[1]
Write-Host -foregroundcolor White " Exchange On Premises Sharing domains:"
Write-Host -foregroundcolor White " Domain:"
Write-Host " " $SPOnprem.Domains.Domain[0]
Write-Host -foregroundcolor White " Action:"
Write-Host " " $SPOnprem.Domains.Actions[0]
Write-Host -foregroundcolor White " Domain:"
Write-Host " " $SPOnprem.Domains.Domain[1]
Write-Host -foregroundcolor White " Action:"
Write-Host " " $SPOnprem.Domains.Actions[1]
Write-Host -foregroundcolor Yellow "Exchange Online Sharing domains should match the Exchange On Premises Sharing Domains"
}
}



 #endregion 

 #region ExoOauthFuntions


 Function EXOIntraOrgConCheck{

Write-Host -foregroundcolor Green " Get-IntraOrganizationConnector | Select TargetAddressDomains,DiscoveryEndpoint,Enabled" 
Write-Host $bar
$exoIntraOrgCon = Get-IntraOrganizationConnector | Select TargetAddressDomains,DiscoveryEndpoint,Enabled
#$IntraOrgConCheck
$IOC=$exoIntraOrgCon | fl
$IOC
Write-Host $bar
Write-Host -foregroundcolor Green " SUMMARY - Online Intra Organization Connector" 
Write-Host $bar

Write-Host -foregroundcolor White " Target Address Domains: " 
if ($exoIntraOrgCon.TargetAddressDomains -like  "*$ExchangeOnpremDomain*"){
Write-Host -foregroundcolor Green " " $exoIntraOrgCon.TargetAddressDomains
}
else
{
Write-Host -foregroundcolor Red " Target Address Domains is NOT correct."
Write-Host -foregroundcolor White " Should contain the $ExchangeOnpremDomain"
}

Write-Host -foregroundcolor White " DiscoveryEndpoint: " 
if ($exoIntraOrgCon.DiscoveryEndpoint -like $EDiscoveryEndpoint.OnPremiseDiscoveryEndpoint){
Write-Host -foregroundcolor Green $exoIntraOrgCon.DiscoveryEndpoint
}
else
{
Write-Host -foregroundcolor Red " DiscoveryEndpoint is NOT correct. "
Write-Host -foregroundcolor White "  Should be " $EDiscoveryEndpoint.OnPremiseDiscoveryEndpoint
}
Write-Host -foregroundcolor White " Enabled: " 
if ($exoIntraOrgCon.Enabled -like  "True"){ 
Write-Host -foregroundcolor Green "  True "  
}
else
{
Write-Host -foregroundcolor Red "  Enabled is NOT correct."
Write-Host -foregroundcolor White " Should be True"
}
}


Function EXOIntraOrgConfigCheck{

Write-Host -foregroundcolor Green " Get-IntraOrganizationConfiguration | Select OnPremiseTargetAddresses" 
Write-Host $bar
$exoIntraOrgConfig = Get-IntraOrganizationConfiguration | Select OnPremiseTargetAddresses
#$IntraOrgConCheck
$IOConfig=$exoIntraOrgConfig | fl
$IOConfig
Write-Host $bar
Write-Host -foregroundcolor Green " SUMMARY - Online Intra Organization Configuration" 
Write-Host $bar

Write-Host -foregroundcolor White " OnPremiseTargetAddresses: " 
if ($exoIntraOrgConfig.OnPremiseTargetAddresses -like  "*$ExchangeOnpremDomain*"){
Write-Host -foregroundcolor Green " " $exoIntraOrgConfig.OnPremiseTargetAddresses
}
else
{
Write-Host -foregroundcolor Red " OnPremise Target Addressess are NOT correct."
Write-Host -foregroundcolor White " Should contain the $ExchangeOnpremDomain"
}
}


Function EXOauthservercheck{

Write-Host -foregroundcolor Green " Get-AuthServer -Identity 00000001-0000-0000-c000-000000000000 | select name,issueridentifier,enabled" 
Write-Host $bar
$exoauthserver = Get-AuthServer -Identity 00000001-0000-0000-c000-000000000000 | select name,issueridentifier,enabled
#$IntraOrgConCheck
$authserver=$exoauthserver | fl
$authserver
Write-Host $bar
Write-Host -foregroundcolor Green " SUMMARY - Exchange Online Authorization Server" 
Write-Host $bar

Write-Host -foregroundcolor White " IssuerIdentifier: " 
if ($exoauthserver.IssuerIdentifier -like  "00000001-0000-0000-c000-000000000000"){
Write-Host -foregroundcolor Green " " $exoauthserver.IssuerIdentifier
}
else
{
Write-Host -foregroundcolor Red " Authorization Server object is NOT correct."

Write-Host -foregroundcolor White " Enabled: " 
if ($exoauthserver.Enabled -like  "True"){ 
Write-Host -foregroundcolor Green "  True "  
}
else
{
Write-Host -foregroundcolor Red "  Enabled is NOT correct."
Write-Host -foregroundcolor White " Should be True"
}
}
}

Function EXOtestoauthcheck{
Write-Host -foregroundcolor Green " Test-OAuthConnectivity -Service EWS -TargetUri $Global:ExchangeOnPremEWS -Mailbox $useronline " 
Write-Host $bar

$exotestoauth = Test-OAuthConnectivity -Service EWS -TargetUri $Global:ExchangeOnPremEWS -Mailbox $useronline | select *
$exoOAC = $exotestoauth | fl
$exoOAC
$bar
$exotestoauth.Detail.FullId
$bar
if ($exotestoauth.Detail.FullId -like '*(401) Unauthorized*'){
write-host -ForegroundColor Red "Error: (401) Unauthorized"
if ($exotestoauth.Detail.FullId -like 'The user specified by the user-context in the token does not exist*'){
write-host "The user specified by the user-context in the token does not exist"
write-host "Please run Test-OAuthConnectivity with a different Exchange Online Mailbox"

}
   
Write-Host $bar
#$OAuthConnectivity.detail.LocalizedString
Write-Host -foregroundcolor Green " SUMMARY - Test OAuth COnnectivity" 
Write-Host $bar
if ($OAuthConnectivity.ResultType -like  "Success"){
Write-Host -foregroundcolor Green " OAuth Test was completed successfully " 
}
else
{
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
$IntraOrgCon = Get-IntraOrganizationConnector | Select Name,TargetAddressDomains,DiscoveryEndpoint,Enabled
#if ($Auth -like "OAuth" -OR [string]::IsNullOrWhitespace($Auth))
#{
$EDiscoveryEndpoint = Get-IntraOrganizationConfiguration | select OnPremiseDiscoveryEndpoint

#fix 3
$SPDomainsOnprem = Get-SharingPolicy | fl Domains
$SPOnprem = Get-SharingPolicy | Select *
#}

#if ([string]::IsNullOrWhitespace($Auth))
#{
#Get-SUMMARY;
#}

if($Auth -like "DAuth" -and $IntraOrgCon.enabled -Like "True")
{
Write-Host $bar
Write-Host -foregroundcolor yellow "  Warning: Intra Organization Connector is Enabled -> Free Busy Lookup is done using OAuth"
Write-Host $bar
}

if ($Organization -like "Onprem" -OR [string]::IsNullOrWhitespace($Organization))

{
#region DAutch Checks
if ($Auth -like "dauth" -OR [string]::IsNullOrWhitespace($Auth))

{
Write-Host -foregroundcolor Green " `n `n ************************************TestingDAuth configuration************************************************* `n `n " 

OrgRelCheck
Write-Host $bar
if ($pause -eq "True")
{
$pause = Read-Host " Press Enter when ready to get Federation Information Details.  "   
Write-Host $bar
}
FedInfoCheck
if ($pause -eq "True")
{
$pause = Read-Host " Press Enter when ready to get Federation Trust Details. "   
Write-Host $bar
}
FedTrustCheck
Write-Host $bar
Write-Host -foregroundcolor Green " Test-FederationTrustCertificate" 
Write-Host $bar
Test-FederationTrustCertificate
Write-Host $bar
if ($pause -eq "True")
{
$pause = Read-Host " Press Enter when ready to check On-Prem Autodiscover Virtual Directory Details. "   
Write-Host $bar
}   
AutoDVirtualDCheck
Write-Host $bar
if ($pause -eq "True")
{
$pause = Read-Host " Press Enter when ready to Grab On-Prem Web Services Virtual Directory. "   
Write-Host $bar
}
EWSVirtualDirectoryCheck
if ($pause -eq "True")
{
Write-Host $bar
$pause = Read-Host " Press Enter when ready to  check the Availability Address Space configuration. "   
Write-Host $bar
}
AvailabilityAddressSpaceCheck
if ($pause -eq "True")
{
Write-Host $bar
$pause = Read-Host " Press Enter when ready to Test-FederationTrust. "   
Write-Host $bar
}
#need to grab errors and provide alerts in error case 
TestFedTrust
if ($pause -eq "True")
{
Write-Host $bar
$pause = Read-Host " Press Enter when ready to Test the OrganizationRelationship. "   
Write-Host $bar
}
TestOrgRel
}


#endregion


#region OAuth Check

if ($Auth -like "OAuth" -OR [string]::IsNullOrWhitespace($Auth))

{
if ($pause -eq "True")
{
$pause = Read-Host " Press Enter when ready to Grab OAuth Configuration Details. "   
Write-Host $bar
}
Write-Host -foregroundcolor Green " `n `n ************************************TestingOAuth configuration************************************************* `n `n " 
Write-Host $bar
IntraOrgConCheck
Write-Host $bar
if ($pause -eq "True"){
$pause = Read-Host " Press Enter when ready to Grab the Auth Server Details. "   
Write-Host $bar
}
AuthServerCheck
if ($pause -eq "True"){
$pause = Read-Host " Press Enter when ready to Grab the Partner Application Details. "   
Write-Host $bar
}
PartnerApplicationCheck
Write-Host $bar
if ($pause -eq "True")
{
$pause = Read-Host " Press Enter when ready to Check the Exchange Online-ApplicationAccount. "   
Write-Host $bar
}
ApplicationAccounCheck
Write-Host $bar
if ($pause -eq "True")
{
$pause = Read-Host " Press Enter when ready to check the ManagementRoleAssignment of the Exchange Online-ApplicationAccount . "   
Write-Host $bar
}
ManagementRoleAssignmentCheck
Write-Host $bar
if ($pause -eq "True")
{
$pause = Read-Host " Press Enter when ready to Grab Auth config Details. "   
Write-Host $bar
}
AuthConfigCheck
Write-Host $bar
if ($pause -eq "True")
{
$pause = Read-Host " Press Enter when ready to Grab information for the Auth Certificate. "   
Write-Host $bar
}
CurrentCertificateThumbprintCheck
Write-Host $bar
if ($pause -eq "True")
{
$pause = Read-Host " Press Enter when ready to  check the On Prem Autodiscover Virtual Directory Configuration. "   
Write-Host $bar
}
AutoDVirtualDCheckOAuth
$AutoDiscoveryVirtualDirectoryOAuth
Write-Host $bar
if ($pause -eq "True")
{
$pause = Read-Host " Press Enter when ready to On-Prem Web Services Virtual Directory. "   
Write-Host $bar
}
EWSVirtualDirectoryCheckOAuth
Write-Host $bar
if ($pause -eq "True")
{
$pause = Read-Host " Press Enter when ready to Grab AvailabilityAddressSpace. "   
Write-Host $bar
}
AvailabilityAddressSpaceCheckOAuth
Write-Host $bar
if ($pause -eq "True")
{
$pause = Read-Host " Press Enter when ready to test the Test-OAuthConnectivity. "   
Write-Host $bar
}
OAuthConnectivityCheck
Write-Host $bar
}
#$bar
#endregion
}


# EXO Part

if ($Organization -like "Online" -OR [string]::IsNullOrWhitespace($Organization))

{
#region ConnectExo

Write-Host " Exchange Online Info"
$bar

#Exchange Online Management Shell 
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
install-module AzureAD -AllowClobber
#$CreateEXOPSSession = (Get-ChildItem -Path $env:userprofile -Filter CreateExoPSSession.ps1 -Recurse -ErrorAction SilentlyContinue -Force | Select -Last 1).DirectoryName. "$CreateEXOPSSession\CreateExoPSSession.ps1"
#Import-Module $((Get-ChildItem -Path $($env:LOCALAPPDATA+"\Apps\2.0\") -Filter CreateExoPSSession.ps1 -Recurse ).FullName | Select-Object -Last 1)
#Connect-EXOPSSession 
#Connect-EXOPSSession

#RestV3 connection

Install-Module -Name ExchangeOnlineManagement
Connect-ExchangeOnline 

Write-Host "========================================================="
Write-Host "Get-SharingPolicy | FL"
Write-Host "========================================================="
Get-SharingPolicy | FL



# Variables
$ExoOrgRel = Get-OrganizationRelationship | Where{($_.DomainNames -like $ExchangeOnPremDomain )} | select Enabled,Identity,DomainNames,FreeBusy*,Target*
$ExoIntraOrgCon = Get-IntraOrganizationConnector | Select Name,TargetAddressDomains,DiscoveryEndpoint,Enabled
$targetadepr1=("https://autodiscover." + $ExchangeOnPremDomain +"/autodiscover/autodiscover.svc/WSSecurity")
$targetadepr2=("https://" + $ExchangeOnPremDomain +"/autodiscover/autodiscover.svc/WSSecurity")
$exofedinfo = get-federationInformation -DomainName $exchangeOnpremDomain  -BypassAdditionalDomainValidation -ErrorAction SilentlyContinue| select *

 #endregion

  
  #region ExoDauthCheck
  if ($Auth -like "dauth" -OR [string]::IsNullOrWhitespace($Auth))

{
Write-Host $bar
Write-Host -foregroundcolor Green " `n `n ************************************Testing DAuth configuration************************************************* `n `n " 

ExoOrgRelCheck
Write-Host $bar
if ($pause -eq "True")
{
$pause = Read-Host " Press Enter when ready to get Federation Organization Identifier Details.  "   
Write-Host $bar
}
EXOFedOrgIdCheck
Write-Host $bar
if ($pause -eq "True")
{
$pause = Read-Host " Press Enter when ready to test Organization Relationship.  "   
Write-Host $bar
}
EXOTestOrgRelCheck
if ($pause -eq "True")
{
$pause = Read-Host " Press Enter when ready to get Sharing Policy Details.  "   
Write-Host $bar
}
SharingPolicyCheck

}
 #endregion 

  #region ExoOauthCheck
    if ($Auth -like "oauth" -OR [string]::IsNullOrWhitespace($Auth))

{

Write-Host -foregroundcolor Green " `n `n ************************************Testing OAuth configuration************************************************* `n `n " 
Write-Host $bar
ExoIntraOrgConCheck

Write-Host $bar
if ($pause -eq "True")
{
$pause = Read-Host " Press Enter when ready to get Organization Configuration Details.  "   
Write-Host $bar
}

EXOIntraOrgConfigCheck
Write-Host $bar
if ($pause -eq "True")
{
$pause = Read-Host " Press Enter when ready to get Server Authorization Details.  "   
Write-Host $bar
}

EXOauthservercheck
Write-Host $bar
if ($pause -eq "True")
{
$pause = Read-Host " Press Enter when ready to get OAuth Connectivity Details.  "   
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
