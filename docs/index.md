# Hybrid Free Busy Configuration Check

##Description

To View this Project at GitHub! [GitHub Repository](https://github.com/MarcoLFrancisco/Hybrid-Free-Busy-Configuration-Checker)

To Download the latest release: [FreeBusyChecker.ps1](https://github.com/MarcoLFrancisco/Hybrid-Free-Busy-Configuration-Checker/releases/download/Version1/FreeBusyChecker.ps1)

To Provide Feedback about this tool: [Feedback Form](https://forms.office.com/pages/responsepage.aspx?id=v4j5cvGGr0GRqy180BHbR2LVru-UswhJmHot_XEUrVVURFVMRkE5VUg4QUU0MEpNRjgxUExPVlBVOS4u)



- This script does not make changes to current settings. It collects relevant configuration information regarding Hybrid Free Busy configurations on Exchange On Premises Servers and on Exchange Online, both for OAuth and DAuth.

- This is a Beta Version. Please doublecheck on any information provided by this script before procceding to address any changes to your envoironment. Be advised that there may be incorrections in the provided output.

##Use 

Collects OAuth and DAuth Hybrid Availability Configuration Settings Both for Exchange On Premises and Exchange Online  


##Exchange Shell Output

![image](https://user-images.githubusercontent.com/3670637/215355627-ee99b28d-1753-4461-8cef-969340cbc7a3.png)

##TXT Output

![image](https://user-images.githubusercontent.com/98214653/235616232-b0d66185-ec5f-4ff7-a81a-f7250f9accc1.png)

##HTML Output
![image](https://github.com/MarcoLFrancisco/Hybrid-Free-Busy-Configuration-Checker/assets/98214653/7f4443c4-41d1-40a3-b97a-f872c850505f)

##Supported Exchange Server Versions

The script can be used to validate the Availability configuration of the following Exchange Server Versions: - Exchange Server 2013 - Exchange Server 2016 - Exchange Server 2019 - Exchange Online


##Required Permissions

- Organization Management
- Domain Admins (only necessary for the DCCoreRatio parameter)


Please make sure that the account used is a member of the Local Administrator group. This should be fulfilled on Exchange servers by being a member of the Organization Management group. However, if the group membership was adjusted or in case the script is executed on a non-Exchange system like a management server, you need to add your account to the Local Administrator group. 

##Other Pre Requisites

- AD management Tools:
  
  If not available, they can be installed with the following command:
    
    Install-windowsfeature -name AD-Domain-Services -IncludeManagementTools 
    
    Imports and Installs the following Modules (if not available):
      PSSnapin: microsoft.exchange.management.powershell.snapin
      Module  : ActiveDirectory Module 
      Module  : ExchangeOnlineManagement Module 

- Exchange Online Powershell Module:

  To Download: [Exchange Online Powershell Module](https://learn.microsoft.com/en-us/powershell/exchange/exchange-online-powershell-v2?view=exchange-ps  )
                
##Syntax

      FreeBusyChecker.ps1
        [-Auth <string[]>]
        [-Org <string>]
        [-Pause]
  
  
##How To Run

- This script must be run as Administrator in Exchange Management Shell on an Exchange Server. You can provide no parameters and the script will just run against Exchnage On Premises and Exchange Online to query for OAuth and DAuth configuration setting. It will compare existing values with standard values and provide detail of what may not be correct. 

- Please take note that though this script may output that a specific setting is not a standard sertting, it does not mean that your configurations are incorrect. For exmaple, DNS may be configured with specific mapppings that this script can not evaluate.



##Valid Input Option Parameters

  Paramater               : Auth
    Options               : All; DAuth; OAUth; Null
        
        All               : Collects information both for OAuth and DAuth;
        DAuth             : DAuth Authentication
        OAuth             : OAuth Authentication
        Default Value.    : Null. No swith input means the script will collect information for the current used method. If OAuth is enabled only OAuth is checked.


  Paramater               : Org
    Options               : ExchangeOnPremise; ExchangeOnline; Null
    
        ExchangeOnPremise : Use ExchangeOnPremise parameter to collect Availability information in the Exchange On Premise Tenant
        ExchangeOnline    : Use ExchangeOnline parameter to collect Availability information in the Exchange Online Tenant
        Default Value.    : Null. No swith input means the script will collect both Exchange On Premise and Exchange OnlineAvailability configuration Detail


  Paramater               : Pause
    Options               : Null; True; False
    
        True              : Use the True parameter to use this script pausing after each test done.
        False             : To use this script not pausing after each test done no Pause Parameter is needed.
        Default Value.    : False.


  Paramater               : Help
  
    Options               : Null; True; False

        True              : Use the $True parameter to use display valid parameter Options.



##Command Examples


- This cmdlet will run Free Busy Checker script and check Availability for Exchange On Premises and Exchange Online for the currently used method, OAuth or DAuth. If OAuth is enabled OAUth is checked. If OAUth is not enabled, DAuth Configurations are collected.

            PS C:\> .\FreeBusyChecker.ps1


- This cmdlet will run Free Busy Checker script and check Availability OAuth and DAuth Configurations both for Exchange On Premises and Exchange Online.

            PS C:\> .\FreeBusyChecker.ps1 -Auth All

- This cmdlet will run the Free Busy Checker Script against for OAuth Availability Configurations only.

            PS C:\> .\FreeBusyChecker.ps1 -Auth OAuth

- This cmdlet will run the Free Busy Checker Script against for DAuth Availability Configurations only.

            PS C:\> .\FreeBusyChecker.ps1 -Auth DAuth

- This cmdlet will run the Free Busy Checker Script for Exchange Online Availability Configurations only.

            PS C:\> .\FreeBusyChecker.ps1 -Org ExchangeOnline

- This cmdlet will run the Free Busy Checker Script for Exchange On Premises OAuth and DAuth Availability Configurations only.

            PS C:\> .\FreeBusyChecker.ps1 -Org ExchangeOnPremise

- This cmdlet will run the Free Busy Checker Script for Exchange On Premises Availability OAuth Configurations, pausing after each test done.

            PS C:\> .\FreeBusyChecker.ps1 -Org ExchangeOnPremise -Auth OAuth -Pause $True

