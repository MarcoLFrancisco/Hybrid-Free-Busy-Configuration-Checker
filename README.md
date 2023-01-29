# On-Premises-Hybrid-Free-Busy-Configuration-Checker
Collects On Premises Hybrid Configuration Settings 


This script is still under development. It is not a stable finished version.

This script does not make changes to corrent settings. It collects relevant configuration information regarding Hybrid Free Busy configurations on Exchange On Premises Servers and on Exchange Online, both for OAuth and DAuth.

This is a Beta Version. Please doublecheck on any information provided by this script before procceding to address any changes to your envoironment. Be advised that there may be incorrections in the provided output.


Initial Screen Output:

![image](https://user-images.githubusercontent.com/3670637/215355627-ee99b28d-1753-4461-8cef-969340cbc7a3.png)

Supported Exchange Server Versions:

        The script can be used to validate the Availability configuration of the following Exchange Server Versions: - Exchange Server 2013 - Exchange Server 2016 - Exchange Server 2019 - Exchange Online

Required Permissions:

                Organization Management
                Domain Admins (only necessary for the DCCoreRatio parameter)


        Please make sure that the account used is a member of the Local Administrator group. This should be fulfilled on Exchange servers by being a        member of the Organization Management group. However, if the group membership was adjusted or in case the script is executed on a non-Exchange system like a management server, you need to add your account to the Local Administrator group. You also need to be a member of the following groups:

                

Syntax:

      FreeBusyChecker.ps1
        [-Auth <string[]>]
        [-Org <string>]
        [-Pause]
  
How To Run:

      This script must be run as Administrator in Exchange Management Shell on an Exchange Server. You can provide no parameters and the script will just run against the local server and provide the detail output of the configuration of the server.



Valid Input Option Parameters:

  Paramater: Auth
    Options  : DAuth; OAUth; Null
        DAuth        : DAuth Authentication
        OAuth        : OAuth Authentication
        Default Value: Null. No swith input means the script will collect both DAuth and OAuth Availability Configuration Detail

  Paramater: Org
    Options  : EOP; EOL; Null
        EOP          : Use EOP parameter to collect Availability information in the Exchange On Premise Tenant
        EOL          : Use EOL parameter to collect Availability information in the Exchange Online Tenant
        Default Value: Null. No swith input means the script will collect both Exchange On Premise and Exchange OnlineAvailability configuration Detail

  Paramater: Pause
    Options  : Null; True; False
        True         : Use the True parameter to use this script pausing after each test done.
        False        : To use this script not pausing after each test done no Pause Parameter is needed.
        Default Value: False.

  Paramater: Help
    Options  : Null; True; False
        True         : Use the True parameter to use display valid parameter Options.



Examples:


  This cmdlet with run Health Checker Script by default and check Organization Availability Configurations both for Exchange On Premises and Exchange Online.

            PS C:\> .\FreeBusyChecker.ps1

        This cmdlet will run the Health Checker Script against for OAuth Availability Configurations only.

            PS C:\> .\FreeBusyChecker.ps1 -Auth OAuth

        This cmdlet will run the Health Checker Script against for DAuth Availability Configurations only.

            PS C:\> .\FreeBusyChecker.ps1 -Auth DAuth

        This cmdlet will run the Health Checker Script for Exchange Online Availability Configurations only.

            PS C:\> .\FreeBusyChecker.ps1 -Org EOL

        This cmdlet will run the Health Checker Script for Exchange On Premises Availability Configurations only.

            PS C:\> .\FreeBusyChecker.ps1 -Org EOP

        This cmdlet will run the Health Checker Script for Exchange On Premises Availability OAuth Configurations, Pausing after each test done.

            PS C:\> .\FreeBusyChecker.ps1 -Org EOP -Auth OAuth -Pause $True



    
    
    
    
    
    
    
    
