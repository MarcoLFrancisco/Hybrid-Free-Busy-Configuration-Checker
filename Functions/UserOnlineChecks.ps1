
function UserOnlineCheck {
    Write-Host -ForegroundColor Green "Online Mailbox: $UserOnline"
    Write-Host "Press the Enter key if OK or type an Exchange Online Email address and press the Enter key"
    $UserOnlineCheck = [System.Console]::ReadLine()
    if (![string]::IsNullOrWhitespace($UserOnlineCheck)) {
        $script:UserOnline = $UserOnlineCheck
    }
}



function ExchangeOnlineDomainCheck {
    #$ExchangeOnlineDomain
    Write-Host -ForegroundColor Green " Exchange Online Domain: $ExchangeOnlineDomain"
    Write-Host " Press Enter if OK or type in the Exchange Online Domain and press the Enter key."
    $ExchangeOnlineDomainCheck = [System.Console]::ReadLine()
    if (![string]::IsNullOrWhitespace($ExchangeOnlineDomainCheck)) {
        $script:ExchangeOnlineDomain = $ExchangeOnlineDomainCheck
    }
}