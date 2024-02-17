<# function UserOnPremCheck {
    Write-Host -ForegroundColor Green " On Premises Hybrid Mailbox: $UserOnPrem"
    Write-Host " Press Enter if OK or type in an Exchange OnPremises Hybrid email address and press the Enter key."
    $UserOnPremCheck = [System.Console]::ReadLine()
    if (![string]::IsNullOrWhitespace($UserOnPremCheck)) {
        $script:UserOnPrem = $UserOnPremCheck
    }
}

function ExchangeOnPremDomainCheck {
    #$exchangeOnPremDomain
    Write-Host -ForegroundColor Green " On Premises Mail Domain: $exchangeOnPremDomain"
    Write-Host " Press Enter if OK or type in the Exchange On Premises Mail Domain and press the Enter key."
    $exchangeOnPremDomainCheck = [System.Console]::ReadLine()
    if (![string]::IsNullOrWhitespace($exchangeOnPremDomainCheck)) {
        $script:exchangeOnPremDomain = $exchangeOnPremDomainCheck
    }
}

function ExchangeOnPremEWSCheck {
    Write-Host -ForegroundColor Green " On Premises EWS External URL: $exchangeOnPremEWS"
    Write-Host " Press Enter if OK or type in the Exchange On Premises EWS URL and press the Enter key."
    $exchangeOnPremEWSCheck = [System.Console]::ReadLine()
    if (![string]::IsNullOrWhitespace($exchangeOnPremEWSCheck)) {
        $exchangeOnPremEWS = $exchangeOnPremEWSCheck
    }
}

function ExchangeOnPremLocalDomainCheck {
    Write-Host -ForegroundColor Green " On Premises Root Domain: $exchangeOnPremLocalDomain  "
    Write-Host " Press Enter if OK or type in the Exchange On Premises Root Domain."
    $exchangeOnPremLocalDomain = [System.Console]::ReadLine()
    if ([string]::IsNullOrWhitespace($ADDomain)) {
        $exchangeOnPremLocalDomain = $exchangeOnPremDomain
    }
    if ([string]::IsNullOrWhitespace($exchangeOnPremLocalDomain)) {
        $exchangeOnPremLocalDomain = $exchangeOnPremLocalDomainCheck
    }
} #>