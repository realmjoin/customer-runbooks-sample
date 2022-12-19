#Requires -Modules @{ModuleName = "RealmJoin.RunbookHelper"; ModuleVersion = "0.6.0" }

param(
    [ValidateScript( { Use-RJInterface -DisplayName "Minimum days left in Cert/Token to be healthy." } )]
    [int] $Days = 30
)

Connect-RjRbGraph

$minDate = (get-date) + (New-TimeSpan -Day $Days)

$applePushCerts = Invoke-RjRbRestMethodGraph -Resource "/deviceManagement/applePushNotificationCertificate" -ErrorAction SilentlyContinue
if ($applePushCerts) {
    #"## Apple Push Notification Certs found."
    foreach ($ApplePushCert in $applePushCerts) {
        "## Apple Push Cert '$($ApplePushCert.appleIdentifier)' expiry is/was on $((get-date -date $ApplePushCert.expirationDateTime).ToShortDateString())."
        "## -> $(((get-date -Date $ApplePushCert.expirationDateTime) - (get-date)).Days) days left."
        if(((get-date -Date $ApplePushCert.expirationDateTime) - $minDate) -le 0) {
            "## ALERT - Days left is below limit!"
        }
        ""
    }
}


$vppTokens = Invoke-RjRbRestMethodGraph -Resource "/deviceAppManagement/vppTokens" -ErrorAction SilentlyContinue
if ($vppTokens) {
    #"## VPP Tokens found."
    foreach ($token in $vppTokens) {
        if ($token.state -ne 'valid') {
            "## VPP Token for '$($token.appleId)' is not valid."
            "## ALERT - VPP Token not valid!"
        } else {
            "## VPP Token for '$($token.appleId)' expiry is/was on $((get-date -date $token.expirationDateTime).ToShortDateString())."
            "## -> $(((get-date -Date $token.expirationDateTime) - (get-date)).Days) days left."
            if (((get-date -date $token.expirationDateTime) - $minDate) -le 0 ) {
                "## ALERT - Days left is below limit!"
            }
        }
        ""
    }
}

$depSettings = Invoke-RjRbRestMethodGraph -Resource "/deviceManagement/depOnboardingSettings" -Beta -ErrorAction SilentlyContinue
if ($depSettings) {
    #"## DEP Settings found."
    foreach ($token in $depSettings) {
        "## DEP Settings for '$($token.appleIdentifier)' expiry is/was on $((get-date -date $token.tokenExpirationDateTime).ToShortDateString())."
        "## -> $(((get-date -Date $token.tokenExpirationDateTime) - (get-date)).Days) days left."
        if (((get-date -date $token.tokenExpirationDateTime) - $minDate) -le 0) {
            "## ALERT - Days left is below limit!"
        }
        ""
    }
}