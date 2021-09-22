Connect-PnPOnline -Url "https://[TENANTNAME]-admin.sharepoint.com"
$users = Get-PnPAzureADUser -Filter "accountEnabled eq true and UserType eq 'Member'" -EndIndex $null
$subscriptionStatus = @()

$users | % {

    try {
        $userStatus = Get-PnPSubscribeSharePointNewsDigest -Account $_.UserPrincipalName -ErrorAction Stop
        $isSubscribed = $userStatus.Enabled
    } 
    catch {
        if ($Error[0].Exception.Message -like "*(401)*") {
            $isSubscribed = "n/a"
        }
        if ($Error[0].Exception.Message -like "*(403)*") {
            $isSubscribed = "Access Denied"
        }
        if ($Error[0].Exception.Message -like "*Sequence contains no matching element*") {
            $isSubscribed = $true
        }
    }

    $userSubscriptionStatus = New-Object PSObject
    $userSubscriptionStatus | Add-Member NoteProperty UPN $_.UserPrincipalName
    $userSubscriptionStatus | Add-Member NoteProperty IsSubscribed $isSubscribed
    $subscriptionStatus += $userSubscriptionStatus

    $subscribedCount = ($subscriptionStatus | select isSubscribed | ? isSubscribed -eq $true).Count
    $unubscribedCount = ($subscriptionStatus | select isSubscribed | ? isSubscribed -eq $false).Count

    Write-Host "Subscribed: $subscribedCount"
    Write-Host "Unsubscribed: $unubscribedCount"

}

