# Update-akoyaGOSharePointSitePermission.ps1
# Version: 0.1
#
# Raises the BCO akoyaGO Integration application's EXISTING grant on one
# SharePoint site collection to Manage (for example from Write). For a site with
# no grant yet, run Grant-akoyaGOSharePointSitePermission.ps1 instead.
#
# Beginner-friendly usage directly from the akoyaGO public repository:
#   iex ((New-Object System.Net.WebClient).DownloadString('https://raw.githubusercontent.com/akoyago/public/refs/heads/main/scripts/deployment/Update-akoyaGOSharePointSitePermission.ps1'))
#
# Optional: supply the site URL on the same line instead of being prompted:
#   $global:akoyaGOSiteUrl = 'https://contoso.sharepoint.com/sites/example'; iex ((New-Object System.Net.WebClient).DownloadString('https://raw.githubusercontent.com/akoyago/public/refs/heads/main/scripts/deployment/Update-akoyaGOSharePointSitePermission.ps1'))
#
# Why a temporary app: Microsoft Graph only lets an app-only connection with
# Sites.FullControl.All APPLICATION permission update or delete an existing site
# grant. A signed-in administrator's token cannot. So the script:
#   1. creates a temporary app registration in the signed-in tenant,
#   2. gives it Microsoft Graph Sites.FullControl.All with admin consent,
#   3. uses it to change the grant to Manage and reads the result back,
#   4. deletes the temporary app, also from deleted items, even when a step fails.
# The temporary app's secret expires after one hour in case deletion cannot finish.
#
# It runs in an isolated child scope so Invoke-Expression cannot collide with
# variables or functions already present in the user's PowerShell session.

& {
$SiteUrl = [string](Get-Variable `
    -Name "akoyaGOSiteUrl" `
    -Scope Global `
    -ValueOnly `
    -ErrorAction SilentlyContinue)
Remove-Variable -Name "akoyaGOSiteUrl" -Scope Global -ErrorAction SilentlyContinue

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$applicationId = "a86b9632-42bf-4dfe-83c8-bbc95145504b"
$applicationDisplayName = "BCO akoyaGO Integration"
$microsoftGraphAppId = "00000003-0000-0000-c000-000000000000"
$requiredRole = "manage"
# Site roles that include everything Manage allows. A grant with one of these is left as it is.
$higherRoles = @("fullcontrol", "owner")
$temporaryAppPermission = "Sites.FullControl.All"
$requiredGraphScopes = @("Application.ReadWrite.All", "AppRoleAssignment.ReadWrite.All")
$requiredModule = "Microsoft.Graph.Authentication"
$graph = "https://graph.microsoft.com/v1.0"
$temporaryAppName = "akoyaGO temporary site grant update $((Get-Date).ToUniversalTime().ToString('yyyy-MM-dd HH:mm')) UTC"
$secretExpiry = (Get-Date).ToUniversalTime().AddHours(1)
$graphConnected = $false
$temporaryApp = $null

function Write-Section {
    param([Parameter(Mandatory = $true)][string]$Text)
    Write-Host ""
    Write-Host "=== $Text ===" -ForegroundColor Cyan
}

function Stop-WithInstructions {
    param(
        [Parameter(Mandatory = $true)][string]$Problem,
        [Parameter(Mandatory = $true)][string[]]$Fix
    )

    Write-Host ""
    Write-Host "FAILED: $Problem" -ForegroundColor Red
    Write-Host ""
    Write-Host "How to fix it:" -ForegroundColor Yellow
    foreach ($step in $Fix) {
        Write-Host "  - $step" -ForegroundColor Yellow
    }
    throw $Problem
}

function Get-ObjectPropertyValue {
    param(
        [AllowNull()][object]$InputObject,
        [Parameter(Mandatory = $true)][string]$Name
    )

    if ($null -eq $InputObject) {
        return $null
    }

    if ($InputObject -is [System.Collections.IDictionary]) {
        if ($InputObject.Contains($Name)) {
            return $InputObject[$Name]
        }
        return $null
    }

    $property = $InputObject.PSObject.Properties[$Name]
    if ($null -ne $property) {
        return $property.Value
    }

    return $null
}

# Graph often puts its JSON error body (with the request ID) in ErrorDetails instead of the exception message.
function Get-ErrorText {
    param([Parameter(Mandatory = $true)][System.Management.Automation.ErrorRecord]$ErrorRecord)

    $body = [string](Get-ObjectPropertyValue -InputObject $ErrorRecord.ErrorDetails -Name "Message")
    if ([string]::IsNullOrWhiteSpace($body) -or $body -eq $ErrorRecord.Exception.Message) {
        return $ErrorRecord.Exception.Message
    }
    return "$($ErrorRecord.Exception.Message) $body"
}

# Graph call as the temporary app. Invoke-MgGraphRequest stays signed in as the administrator for cleanup.
function Invoke-AppGraph {
    param(
        [Parameter(Mandatory = $true)][string]$Token,
        [Parameter(Mandatory = $true)][string]$Method,
        [Parameter(Mandatory = $true)][string]$Uri,
        [object]$Body
    )

    $request = @{
        Method = $Method
        Uri = $Uri
        Headers = @{ Authorization = "Bearer $Token" }
        ErrorAction = "Stop"
    }
    if ($null -ne $Body) {
        $request.ContentType = "application/json"
        $request.Body = $Body | ConvertTo-Json -Depth 5
    }
    Invoke-RestMethod @request
}

# Newly created Entra objects can take a little while to reach every directory replica.
function Invoke-WithRetry {
    param(
        [Parameter(Mandatory = $true)][scriptblock]$Action,
        [int]$Attempts = 6,
        [int]$DelaySeconds = 10
    )

    for ($attempt = 1; ; $attempt++) {
        try {
            return & $Action
        }
        catch {
            if ($attempt -ge $Attempts) {
                throw
            }
            Start-Sleep -Seconds $DelaySeconds
        }
    }
}

function Test-IsTargetGrant {
    param([AllowNull()][object]$Permission)

    $identities = @(Get-ObjectPropertyValue -InputObject $Permission -Name "grantedToIdentitiesV2") +
        @(Get-ObjectPropertyValue -InputObject $Permission -Name "grantedToIdentities")
    $matched = @($identities | Where-Object {
        $application = Get-ObjectPropertyValue -InputObject $_ -Name "application"
        $null -ne $application -and (Get-ObjectPropertyValue -InputObject $application -Name "id") -eq $applicationId
    })
    return $matched.Count -gt 0
}

function Get-GrantRoles {
    param([AllowNull()][object]$Permission)

    return @(@(Get-ObjectPropertyValue -InputObject $Permission -Name "roles") |
        Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } |
        ForEach-Object { ([string]$_).ToLowerInvariant() })
}

try {
    Write-Section "Checking PowerShell"

    $isPowerShellCore = $PSVersionTable.PSEdition -eq "Core"
    if ($PSVersionTable.PSVersion -lt [version]"5.1" -or
        ($isPowerShellCore -and $PSVersionTable.PSVersion -lt [version]"7.0")) {
        Stop-WithInstructions `
            -Problem "Windows PowerShell 5.1 or PowerShell 7 is required. You are running $($PSVersionTable.PSVersion)." `
            -Fix @(
                "Install PowerShell 7 with: winget install --id Microsoft.PowerShell --source winget",
                "Open PowerShell 7 or run 'pwsh', then run this script again."
            )
    }

    if (-not [Environment]::Is64BitProcess) {
        Stop-WithInstructions `
            -Problem "This script is running in 32-bit PowerShell." `
            -Fix @(
                "Close this window.",
                "Open 64-bit Windows PowerShell 5.1 or 64-bit PowerShell 7 and rerun the script."
            )
    }

    if (-not $isPowerShellCore) {
        # Older Windows PowerShell defaults may not negotiate the TLS version
        # required by PowerShell Gallery, Microsoft sign-in and Microsoft Graph.
        [Net.ServicePointManager]::SecurityProtocol = `
            [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
    }

    Write-Host "$($PSVersionTable.PSEdition) PowerShell $($PSVersionTable.PSVersion) 64-bit: OK" -ForegroundColor Green

    Write-Section "Validating the site URL"

    if ([string]::IsNullOrWhiteSpace($SiteUrl)) {
        Write-Host "Enter the complete SharePoint site URL (the client's default site for Dynamics documents)." -ForegroundColor Yellow
        Write-Host "Example: https://contoso.sharepoint.com/sites/foundation" -ForegroundColor DarkGray
        $SiteUrl = Read-Host "SharePoint site URL"
    }

    $siteUri = $null
    if ([string]::IsNullOrWhiteSpace($SiteUrl) -or
        -not [Uri]::TryCreate($SiteUrl.Trim(), [UriKind]::Absolute, [ref]$siteUri) -or
        $siteUri.Scheme -ne "https" -or
        $siteUri.Host -notmatch "\.sharepoint\.com$" -or
        -not [string]::IsNullOrWhiteSpace($siteUri.Query) -or
        -not [string]::IsNullOrWhiteSpace($siteUri.Fragment) -or
        -not [string]::IsNullOrWhiteSpace($siteUri.UserInfo)) {
        Stop-WithInstructions `
            -Problem "'$SiteUrl' is not an HTTPS SharePoint Online site URL." `
            -Fix @(
                "Copy the site URL from SharePoint in your browser, without anything after '?' or '#'.",
                "Expected format: https://tenant.sharepoint.com/sites/site-name"
            )
    }

    $sitePath = $siteUri.AbsolutePath.TrimEnd("/")
    if ([string]::IsNullOrWhiteSpace($sitePath)) {
        $sitePath = "/"
    }
    $normalizedSiteUrl = "https://$($siteUri.Host)$sitePath"
    Write-Host "Site: $normalizedSiteUrl" -ForegroundColor Green

    Write-Section "Checking Microsoft Graph module"

    $installedModule = Get-Module -ListAvailable -Name $requiredModule |
        Sort-Object Version -Descending |
        Select-Object -First 1

    if ($null -eq $installedModule) {
        Write-Host "$requiredModule is not installed. Installing it for the current user..." -ForegroundColor Yellow
        try {
            Install-Module -Name $requiredModule -Repository PSGallery -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
        }
        catch {
            Write-Host $_.Exception.Message -ForegroundColor Red
            Stop-WithInstructions `
                -Problem "PowerShell could not install $requiredModule." `
                -Fix @(
                    "Make sure this computer can reach https://www.powershellgallery.com.",
                    "Run: Install-Module Microsoft.Graph.Authentication -Scope CurrentUser -Force",
                    "If prompted to trust PSGallery, answer Y."
                )
        }
    }

    try {
        Import-Module $requiredModule -ErrorAction Stop
    }
    catch {
        Write-Host $_.Exception.Message -ForegroundColor Red
        Stop-WithInstructions `
            -Problem "PowerShell found $requiredModule but could not load it." `
            -Fix @(
                "Close all PowerShell windows and reopen the same PowerShell edition.",
                "Run: Update-Module Microsoft.Graph.Authentication",
                "For Windows PowerShell 5.1, confirm .NET Framework 4.7.2 or newer is installed. PowerShell 7 is recommended."
            )
    }

    Write-Host "$requiredModule $((Get-Module -Name $requiredModule | Select-Object -First 1).Version): OK" -ForegroundColor Green

    Write-Section "Signing in to Microsoft Graph"
    Write-Host "A browser window may open." -ForegroundColor Yellow
    Write-Host "Sign in as a Global Administrator of the tenant that owns $($siteUri.Host)." -ForegroundColor Yellow
    Write-Host "Approve the requested permissions ($($requiredGraphScopes -join ', ')) if asked." -ForegroundColor Yellow

    try {
        Connect-MgGraph -Scopes $requiredGraphScopes -ContextScope Process -NoWelcome -ErrorAction Stop
        $graphConnected = $true
    }
    catch {
        Write-Host $_.Exception.Message -ForegroundColor Red
        Stop-WithInstructions `
            -Problem "Microsoft Graph sign-in failed." `
            -Fix @(
                "Run the script again and complete the browser sign-in.",
                "Use a Global Administrator account in the tenant that owns $($siteUri.Host).",
                "If no browser opens, run: Connect-MgGraph -Scopes '$($requiredGraphScopes -join "','")' -UseDeviceAuthentication"
            )
    }

    $graphContext = Get-MgContext
    $missingScopes = @($requiredGraphScopes | Where-Object { $null -eq $graphContext -or $graphContext.Scopes -notcontains $_ })
    if ($missingScopes.Count -gt 0) {
        Stop-WithInstructions `
            -Problem "The sign-in did not include $($missingScopes -join ', ')." `
            -Fix @(
                "Sign in as a Global Administrator and approve the requested permissions.",
                "Run Disconnect-MgGraph, then rerun this script."
            )
    }

    $tenantId = [string]$graphContext.TenantId
    Write-Host "Signed in as: $($graphContext.Account)" -ForegroundColor Green
    Write-Host "Tenant ID:    $tenantId" -ForegroundColor Green

    Write-Host ""
    Write-Host "This script will:" -ForegroundColor Yellow
    Write-Host "  1. Create a temporary app registration '$temporaryAppName' in this tenant." -ForegroundColor Yellow
    Write-Host "  2. Give it Microsoft Graph $temporaryAppPermission (application permission, admin consent)." -ForegroundColor Yellow
    Write-Host "     Its secret expires at $($secretExpiry.ToString('yyyy-MM-dd HH:mm')) UTC." -ForegroundColor Yellow
    Write-Host "  3. Use it to raise the existing grant for '$applicationDisplayName' ($applicationId)" -ForegroundColor Yellow
    Write-Host "     on the site collection that contains $normalizedSiteUrl to Manage." -ForegroundColor Yellow
    Write-Host "  4. Delete the temporary app, also from deleted applications." -ForegroundColor Yellow
    Write-Host "Nothing else is changed. If the grant is already Manage or higher, it is left as it is." -ForegroundColor Yellow
    Write-Host ""

    $confirmation = Read-Host "Type RAISE to continue"
    if ($confirmation -cne "RAISE") {
        Write-Host "No changes were made." -ForegroundColor Yellow
        return
    }

    Write-Section "Creating the temporary app"

    try {
        $graphServicePrincipal = Invoke-MgGraphRequest `
            -Method GET `
            -Uri "$graph/servicePrincipals(appId='$microsoftGraphAppId')?`$select=id,appRoles" `
            -ErrorAction Stop
        $fullControlRole = @($graphServicePrincipal.appRoles | Where-Object {
            $_.value -eq $temporaryAppPermission -and @($_.allowedMemberTypes) -contains "Application"
        }) | Select-Object -First 1
        if ($null -eq $fullControlRole) {
            throw "Microsoft Graph's service principal has no $temporaryAppPermission application role."
        }

        $temporaryApp = Invoke-MgGraphRequest `
            -Method POST `
            -Uri "$graph/applications" `
            -ContentType "application/json" `
            -Body (@{
                displayName = $temporaryAppName
                signInAudience = "AzureADMyOrg"
                requiredResourceAccess = @(@{
                    resourceAppId = $microsoftGraphAppId
                    resourceAccess = @(@{ id = $fullControlRole.id; type = "Role" })
                })
            } | ConvertTo-Json -Depth 5) `
            -ErrorAction Stop
        Write-Host "Created app registration: $temporaryAppName (app ID $($temporaryApp.appId))" -ForegroundColor Green

        $secret = Invoke-WithRetry {
            Invoke-MgGraphRequest `
                -Method POST `
                -Uri "$graph/applications/$($temporaryApp.id)/addPassword" `
                -ContentType "application/json" `
                -Body (@{ passwordCredential = @{ displayName = "temporary"; endDateTime = $secretExpiry.ToString("o") } } | ConvertTo-Json -Depth 3) `
                -ErrorAction Stop
        }

        $temporaryServicePrincipal = Invoke-WithRetry {
            Invoke-MgGraphRequest `
                -Method POST `
                -Uri "$graph/servicePrincipals" `
                -ContentType "application/json" `
                -Body (@{ appId = $temporaryApp.appId } | ConvertTo-Json) `
                -ErrorAction Stop
        }

        Invoke-WithRetry {
            Invoke-MgGraphRequest `
                -Method POST `
                -Uri "$graph/servicePrincipals/$($graphServicePrincipal.id)/appRoleAssignedTo" `
                -ContentType "application/json" `
                -Body (@{
                    principalId = $temporaryServicePrincipal.id
                    resourceId = $graphServicePrincipal.id
                    appRoleId = $fullControlRole.id
                } | ConvertTo-Json) `
                -ErrorAction Stop
        } | Out-Null
        Write-Host "Granted Microsoft Graph $temporaryAppPermission to the temporary app." -ForegroundColor Green
    }
    catch {
        Write-Host (Get-ErrorText -ErrorRecord $_) -ForegroundColor DarkRed
        Stop-WithInstructions `
            -Problem "The temporary app could not be created or given $temporaryAppPermission." `
            -Fix @(
                "Sign in as a Global Administrator. Other admin roles cannot grant Microsoft Graph application permissions.",
                "Run Disconnect-MgGraph, then rerun this script and approve the requested permissions.",
                "If it still fails, give akoyaGO support the red error text above."
            )
    }

    Write-Section "Waiting for the temporary app to become active"
    Write-Host "Microsoft Entra can take a few minutes to apply a new permission." -ForegroundColor Yellow

    $tokenUri = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"
    $tokenBody = @{
        client_id = $temporaryApp.appId
        client_secret = $secret.secretText
        scope = "https://graph.microsoft.com/.default"
        grant_type = "client_credentials"
    }
    $lookupUri = '{0}/sites/{1}:{2}?$select=id,displayName,webUrl,siteCollection' -f $graph, $siteUri.Host, $sitePath
    $maximumAttempts = 24
    $site = $null
    $appToken = $null
    $lastError = ""

    for ($attempt = 1; $attempt -le $maximumAttempts -and $null -eq $site; $attempt++) {
        try {
            # A fresh token each time: a token issued before the permission applies never gains it.
            $appToken = (Invoke-RestMethod -Method POST -Uri $tokenUri -Body $tokenBody -ErrorAction Stop).access_token
            $site = Invoke-AppGraph -Token $appToken -Method GET -Uri $lookupUri
        }
        catch {
            $lastError = Get-ErrorText -ErrorRecord $_
            if ($lastError -match "itemNotFound") {
                Stop-WithInstructions `
                    -Problem "Microsoft Graph could not find $normalizedSiteUrl in tenant $tenantId." `
                    -Fix @(
                        "Open $normalizedSiteUrl in a browser and copy the site URL from the address bar.",
                        "Make sure you signed in to the tenant that owns the site, then rerun the script."
                    )
            }
            if ($attempt -lt $maximumAttempts) {
                Write-Host "Not active yet (attempt $attempt of $maximumAttempts). Waiting 15 seconds..." -ForegroundColor DarkGray
                Start-Sleep -Seconds 15
            }
        }
    }

    if ($null -eq $site) {
        Write-Host $lastError -ForegroundColor DarkRed
        Stop-WithInstructions `
            -Problem "The temporary app did not become active within six minutes." `
            -Fix @(
                "Rerun the script. It creates a new temporary app each time and deletes it at the end.",
                "If it fails again, give akoyaGO support the red error text above."
            )
    }

    Write-Host "Resolved site: $($site.webUrl)" -ForegroundColor Green

    if ($null -eq (Get-ObjectPropertyValue -InputObject $site -Name "siteCollection")) {
        # A subsite. A Graph site ID is hostname,siteCollectionId,webId and the grant is on the site collection.
        $siteIdParts = ([string]$site.id).Split(",")
        $siteCollectionId = [guid]::Empty
        if ($siteIdParts.Count -ne 3 -or -not [guid]::TryParse($siteIdParts[1], [ref]$siteCollectionId)) {
            Stop-WithInstructions `
                -Problem "The containing site collection could not be determined from Graph site ID '$($site.id)'." `
                -Fix @("Enter the site collection's root URL from SharePoint admin center > Active sites and rerun the script.")
        }

        $site = Invoke-AppGraph `
            -Token $appToken `
            -Method GET `
            -Uri ('{0}/sites/{1},{2}?$select=id,displayName,webUrl,siteCollection' -f $graph, $siteUri.Host, $siteCollectionId)
        Write-Host "Entered URL is a subsite. Containing site collection: $($site.webUrl)" -ForegroundColor Yellow
    }

    Write-Section "Checking the existing grant"

    $permissionsUri = "$graph/sites/$($site.id)/permissions"
    $grants = @((Invoke-AppGraph -Token $appToken -Method GET -Uri $permissionsUri).value |
        Where-Object { Test-IsTargetGrant -Permission $_ })

    if ($grants.Count -eq 0) {
        Stop-WithInstructions `
            -Problem "'$applicationDisplayName' has no grant on $($site.webUrl), so there is nothing to raise." `
            -Fix @("Run Grant-akoyaGOSharePointSitePermission.ps1 (KA-01525) for this site instead. It creates a Manage grant.")
    }

    if ($grants.Count -gt 1) {
        Stop-WithInstructions `
            -Problem "More than one grant for '$applicationDisplayName' was found on $($site.webUrl). The script will not guess which one to change." `
            -Fix @("Give akoyaGO support the site URL. Permission IDs: $(($grants | ForEach-Object { $_.id }) -join ', ')")
    }

    # The single-permission GET returns the roles reliably; the list sometimes leaves them out.
    $permissionUri = "$permissionsUri/$($grants[0].id)"
    $grant = Invoke-AppGraph -Token $appToken -Method GET -Uri $permissionUri
    $roles = @(Get-GrantRoles -Permission $grant)
    if ($roles.Count -eq 0) {
        Stop-WithInstructions `
            -Problem "Microsoft Graph did not return the role of the existing grant (permission ID $($grant.id))." `
            -Fix @("No change was made. Give akoyaGO support the site URL and permission ID.")
    }

    Write-Host "Current role: $($roles -join ', ')" -ForegroundColor Green

    if ($roles -contains $requiredRole -or @($roles | Where-Object { $higherRoles -contains $_ }).Count -gt 0) {
        Write-Host "The grant already allows everything Manage allows. No change was needed." -ForegroundColor Green
    }
    else {
        Write-Section "Raising the grant to Manage"
        Invoke-AppGraph -Token $appToken -Method PATCH -Uri $permissionUri -Body @{ roles = @($requiredRole) } | Out-Null

        $grant = Invoke-AppGraph -Token $appToken -Method GET -Uri $permissionUri
        $roles = @(Get-GrantRoles -Permission $grant)
        if ($roles -notcontains $requiredRole) {
            Stop-WithInstructions `
                -Problem "After the update, Microsoft Graph reports the role as '$($roles -join ', ')', not Manage." `
                -Fix @("Rerun the script once. If the result is unchanged, give akoyaGO support the site URL and permission ID $($grant.id).")
        }
    }

    Write-Host ""
    Write-Host "SUCCESS" -ForegroundColor Green
    Write-Host "  Site collection:         $($site.webUrl)" -ForegroundColor Green
    Write-Host "  Application:             $applicationDisplayName" -ForegroundColor Green
    Write-Host "  Application (client) ID: $applicationId" -ForegroundColor Green
    Write-Host "  Permission:              $($roles -join ', ')" -ForegroundColor Green
    Write-Host "  Permission ID:           $($grant.id)" -ForegroundColor DarkGray
}
catch {
    Write-Host ""
    Write-Host "Technical error:" -ForegroundColor DarkRed
    Write-Host $_.Exception.Message -ForegroundColor DarkRed
}
finally {
    if ($null -ne $temporaryApp) {
        Write-Section "Deleting the temporary app"
        $deleted = $false
        try {
            Invoke-MgGraphRequest -Method DELETE -Uri "$graph/applications/$($temporaryApp.id)" -ErrorAction Stop | Out-Null
            $deleted = $true
            Write-Host "Deleted '$temporaryAppName'. It can no longer sign in." -ForegroundColor Green
        }
        catch {
            Write-Host (Get-ErrorText -ErrorRecord $_) -ForegroundColor DarkRed
            Write-Host "IMPORTANT: the temporary app could not be deleted and still has $temporaryAppPermission." -ForegroundColor Red
            Write-Host "Delete it now: Microsoft Entra admin center > App registrations > All applications >" -ForegroundColor Red
            Write-Host "  '$temporaryAppName' (application ID $($temporaryApp.appId)) > Delete." -ForegroundColor Red
        }

        if ($deleted) {
            try {
                Invoke-WithRetry {
                    Invoke-MgGraphRequest -Method DELETE -Uri "$graph/directory/deletedItems/$($temporaryApp.id)" -ErrorAction Stop | Out-Null
                } | Out-Null
                Write-Host "Removed it from deleted applications." -ForegroundColor Green
            }
            catch {
                Write-Host "It is still listed under App registrations > Deleted applications for 30 days; you can remove it there." -ForegroundColor Yellow
            }
        }
    }

    if ($graphConnected) {
        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    }
}
}
