# Grant-akoyaGOSharePointSitePermission.ps1
# Version: 0.10
#
# Beginner-friendly usage directly from the akoyaGO public repository:
#   iex ((New-Object System.Net.WebClient).DownloadString('https://raw.githubusercontent.com/akoyago/public/refs/heads/main/scripts/deployment/Grant-akoyaGOSharePointSitePermission.ps1'))
#
# Optional: supply the site URL on the same line instead of being prompted:
#   $global:akoyaGOSiteUrl = 'https://contoso.sharepoint.com/sites/example'; iex ((New-Object System.Net.WebClient).DownloadString('https://raw.githubusercontent.com/akoyago/public/refs/heads/main/scripts/deployment/Grant-akoyaGOSharePointSitePermission.ps1'))
#
# The script grants the fixed target enterprise application identified below
# Manage access to one SharePoint site through the Sites.Selected permission model.
# It runs in an isolated child scope so Invoke-Expression cannot collide with
# variables or functions already present in the user's PowerShell session.

& {
$SiteUrl = [string](Get-Variable `
    -Name "akoyaGOSiteUrl" `
    -Scope Global `
    -ValueOnly `
    -ErrorAction SilentlyContinue)
Remove-Variable -Name "akoyaGOSiteUrl" -Scope Global -ErrorAction SilentlyContinue
$Force = $false

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$applicationId = "a86b9632-42bf-4dfe-83c8-bbc95145504b"
$applicationDisplayName = "BCO akoyaGO Integration"
$microsoftGraphResourceAppId = "00000003-0000-0000-c000-000000000000"
$sitesSelectedApplicationPermissionId = "883ea226-0bf2-4a8f-9f9d-92c9162a727d"
$sitesFullControlDelegatedPermissionId = "5a54b8b3-347c-476d-8f8e-42d5c7424d29"
$requiredRole = "manage"
$requiredGraphScope = "Sites.FullControl.All"
$requiredModule = "Microsoft.Graph.Authentication"
$minimumWindowsPowerShellVersion = [version]"5.1"
$minimumPowerShellCoreVersion = [version]"7.0"
$graphConnected = $false
$graphClientId = $null

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

function Write-GraphFailureDetails {
    param(
        [Parameter(Mandatory = $true)][System.Management.Automation.ErrorRecord]$ErrorRecord,
        [Parameter(Mandatory = $true)][string]$Method,
        [Parameter(Mandatory = $true)][string]$RequestUri
    )

    Write-Host "Graph request: $Method $RequestUri" -ForegroundColor DarkRed
    if (-not [string]::IsNullOrWhiteSpace($graphClientId)) {
        Write-Host "Sign-in client application ID: $graphClientId" -ForegroundColor DarkRed
    }
    Write-Host $ErrorRecord.Exception.Message -ForegroundColor Red

    # Invoke-MgGraphRequest often puts Graph's JSON error body (including the
    # request ID) in ErrorDetails instead of the exception message.
    $errorBody = [string](Get-ObjectPropertyValue -InputObject $ErrorRecord.ErrorDetails -Name "Message")
    if (-not [string]::IsNullOrWhiteSpace($errorBody) -and
        $errorBody -ne $ErrorRecord.Exception.Message) {
        Write-Host "Graph error body: $errorBody" -ForegroundColor DarkRed
    }
}

function Get-GraphErrorGuidance {
    param(
        [Parameter(Mandatory = $true)][System.Management.Automation.ErrorRecord]$ErrorRecord,
        [Parameter(Mandatory = $true)][string]$NormalizedSiteUrl,
        [Parameter(Mandatory = $true)]
        [ValidateSet("FindSite", "ListPermissions", "CreatePermission")]
        [string]$Operation
    )

    $message = $ErrorRecord.Exception.Message + " " + [string](Get-ObjectPropertyValue `
        -InputObject $ErrorRecord.ErrorDetails `
        -Name "Message")
    $statusCode = $null

    $response = Get-ObjectPropertyValue -InputObject $ErrorRecord.Exception -Name "Response"
    $responseStatusCode = Get-ObjectPropertyValue -InputObject $response -Name "StatusCode"
    if ($null -ne $responseStatusCode) {
        $statusCode = [int]$responseStatusCode
    }
    else {
        $exceptionStatusCode = Get-ObjectPropertyValue -InputObject $ErrorRecord.Exception -Name "ResponseStatusCode"
        if ($null -ne $exceptionStatusCode) {
            $statusCode = [int]$exceptionStatusCode
        }
    }

    if ($statusCode -eq 401 -or $message -match "Unauthorized|InvalidAuthenticationToken") {
        return @(
            "Close this PowerShell window, open a new 64-bit Windows PowerShell 5.1 or PowerShell 7 window, and run the script again.",
            "When the browser opens, sign in to the Microsoft 365 tenant that owns $NormalizedSiteUrl.",
            "Do not use a guest account from a different tenant."
        )
    }

    if ($statusCode -eq 403 -or $message -match "accessDenied|Access denied|Forbidden|Authorization_RequestDenied") {
        if ($Operation -eq "FindSite") {
            return @(
                "The failure occurred while resolving the site, before Graph inspected or changed any grant for target application ID $applicationId.",
                "The signed-in account is a tenant administrator but does not currently have access to the site collection containing $NormalizedSiteUrl.",
                "In SharePoint admin center > Active sites, select the site collection, open Membership > Site admins, and add the exact account printed under 'Signed in as'.",
                "Rerun this script after the site-admin change has propagated. The site-admin assignment can be removed after this script succeeds.",
                "For unattended operation without temporarily granting the user site access, use a dedicated provisioning application with Microsoft Graph Sites.FullControl.All APPLICATION permission and certificate authentication; that is a separate, higher-privilege deployment model."
            )
        }

        return @(
            "Check the 'Signed in as' account printed above. The browser account and the PowerShell account can be different.",
            "Use an account that is a Global Administrator or SharePoint Administrator and also has site-admin access to the site collection containing $NormalizedSiteUrl.",
            "Global Administrator and SharePoint Administrator roles do not automatically grant access to every SharePoint site.",
            "In the SharePoint admin center, open Active sites, select this site, open Membership, and add the signed-in account as a Site admin.",
            "Approve/admin-consent the Microsoft Graph delegated permission $requiredGraphScope when prompted.",
            "Run Disconnect-MgGraph, rerun the script, and select the site-admin account in the sign-in window."
        )
    }

    if ($statusCode -eq 404 -or $message -match "itemNotFound|Not Found|Resource.*not.*found") {
        return @(
            "Open $NormalizedSiteUrl in a browser and verify that the site exists.",
            "Copy the URL from the browser address bar and rerun the script with that exact URL.",
            "Make sure you signed in to the tenant that owns the site."
        )
    }

    if ($statusCode -eq 400 -or $message -match "BadRequest|invalidRequest") {
        if ($Operation -eq "FindSite") {
            return @(
                "Verify that the exact request URL printed above contains only the SharePoint site path, with no page, document, query string, or sharing-link suffix.",
                "Open $NormalizedSiteUrl and copy the site home-page URL, then rerun the script.",
                "Give support the Graph error body and request ID printed above. This failure occurred before the target enterprise application was evaluated."
            )
        }

        if ($Operation -eq "ListPermissions") {
            return @(
                "This GET request only reads existing site grants; the target application's permissions do not cause this request to return Bad Request.",
                "The script resolved the entered URL to its containing site-collection root before making this request.",
                "Retry once and give support the exact Graph request, error body, and request ID printed above. Changing the signed-in user's Entra role will not repair a malformed or unsupported request."
            )
        }

        return @(
            "Target enterprise application: '$applicationDisplayName'; application (client) ID: $applicationId. Search Entra ID > Enterprise applications by this ID; its object ID is different in every tenant.",
            "Required API permission on that target application: Microsoft Graph (resource app ID $microsoftGraphResourceAppId) > Sites.Selected APPLICATION permission (role ID $sitesSelectedApplicationPermissionId), with tenant admin consent.",
            "The script's sign-in client is a separate enterprise application. Its application ID is printed above; it needs Microsoft Graph $requiredGraphScope DELEGATED permission (scope ID $sitesFullControlDelegatedPermissionId), with tenant admin consent.",
            "Do not use an Entra object ID in place of application (client) ID $applicationId.",
            "Then rerun the script."
        )
    }

    return @(
        "Read the technical error printed above these instructions.",
        "Confirm that $NormalizedSiteUrl opens in a browser.",
        "Confirm that you are using 64-bit Windows PowerShell 5.1 or PowerShell 7 and have internet access to login.microsoftonline.com and graph.microsoft.com.",
        "Rerun the script. If it still fails, give your administrator the complete red error text."
    )
}

try {
    Write-Section "Checking PowerShell"

    $isWindowsPowerShell = $PSVersionTable.PSEdition -eq "Desktop"
    $isPowerShellCore = $PSVersionTable.PSEdition -eq "Core"

    if (-not $isWindowsPowerShell -and -not $isPowerShellCore) {
        Stop-WithInstructions `
            -Problem "This PowerShell edition is not supported: $($PSVersionTable.PSEdition) $($PSVersionTable.PSVersion)." `
            -Fix @(
                "Install PowerShell 7 with: winget install --id Microsoft.PowerShell --source winget",
                "Open PowerShell 7 or run 'pwsh'.",
                "Run this script again from the PowerShell 7 window."
            )
    }

    if ($isWindowsPowerShell -and $PSVersionTable.PSVersion -lt $minimumWindowsPowerShellVersion) {
        Stop-WithInstructions `
            -Problem "Windows PowerShell 5.1 or newer is required. You are running $($PSVersionTable.PSVersion)." `
            -Fix @(
                "Install Windows Management Framework 5.1, or install PowerShell 7 with: winget install --id Microsoft.PowerShell --source winget",
                "Open the updated PowerShell and rerun the script."
            )
    }

    if ($isPowerShellCore -and $PSVersionTable.PSVersion -lt $minimumPowerShellCoreVersion) {
        Stop-WithInstructions `
            -Problem "PowerShell Core 6 is not supported. PowerShell 7.0 or newer is required. You are running $($PSVersionTable.PSVersion)." `
            -Fix @(
                "Install the current PowerShell 7 release with: winget install --id Microsoft.PowerShell --source winget",
                "Open PowerShell 7 or run 'pwsh', then rerun the script."
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

    if ($isWindowsPowerShell) {
        $netFrameworkRelease = $null
        try {
            $netFrameworkRelease = Get-ItemPropertyValue `
                -Path "HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full" `
                -Name Release `
                -ErrorAction Stop
        }
        catch {
            Stop-WithInstructions `
                -Problem "The script could not confirm that .NET Framework 4.7.2 or newer is installed." `
                -Fix @(
                    "Install .NET Framework 4.7.2 or newer from Microsoft, or use PowerShell 7 instead.",
                    "After installation, restart Windows and rerun the script."
                )
        }

        if ([int]$netFrameworkRelease -lt 461808) {
            Stop-WithInstructions `
                -Problem "Windows PowerShell requires .NET Framework 4.7.2 or newer for Microsoft Graph." `
                -Fix @(
                    "Install .NET Framework 4.7.2 or newer from Microsoft, or use PowerShell 7 instead.",
                    "After installation, restart Windows and rerun the script."
                )
        }

        # Older Windows PowerShell defaults may not negotiate the TLS version
        # required by PowerShell Gallery and Microsoft Graph.
        [Net.ServicePointManager]::SecurityProtocol = `
            [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
    }

    Write-Host "$($PSVersionTable.PSEdition) PowerShell $($PSVersionTable.PSVersion) 64-bit: OK" -ForegroundColor Green

    Write-Section "Validating the site URL"

    if ([string]::IsNullOrWhiteSpace($SiteUrl)) {
        Write-Host "Enter the complete SharePoint site URL." -ForegroundColor Yellow
        Write-Host "Example: https://contoso.sharepoint.com/sites/foundation" -ForegroundColor DarkGray
        $SiteUrl = Read-Host "SharePoint site URL"
    }

    if ([string]::IsNullOrWhiteSpace($SiteUrl)) {
        Stop-WithInstructions `
            -Problem "A SharePoint site URL was not provided." `
            -Fix @(
                "Run the command again and enter the complete SharePoint site URL when prompted.",
                "Example: https://contoso.sharepoint.com/sites/foundation"
            )
    }

    $siteUri = $null
    if (-not [Uri]::TryCreate($SiteUrl.Trim(), [UriKind]::Absolute, [ref]$siteUri)) {
        Stop-WithInstructions `
            -Problem "'$SiteUrl' is not a valid complete URL." `
            -Fix @(
                "Use the complete URL, for example: https://contoso.sharepoint.com/sites/example",
                "Put the URL in quotation marks when running the script."
            )
    }

    if ($siteUri.Scheme -ne "https" -or $siteUri.Host -notmatch "\.sharepoint\.(com|us|de|cn)$") {
        Stop-WithInstructions `
            -Problem "The URL must be an HTTPS SharePoint Online site URL." `
            -Fix @(
                "Copy the site URL from SharePoint in your browser.",
                "Expected format: https://tenant.sharepoint.com/sites/site-name"
            )
    }

    if (-not [string]::IsNullOrWhiteSpace($siteUri.Query) -or
        -not [string]::IsNullOrWhiteSpace($siteUri.Fragment) -or
        -not [string]::IsNullOrWhiteSpace($siteUri.UserInfo)) {
        Stop-WithInstructions `
            -Problem "The site URL cannot contain a query string, fragment, or embedded username." `
            -Fix @(
                "Remove everything beginning with '?' or '#'.",
                "Use only the site address, such as https://contoso.sharepoint.com/sites/example."
            )
    }

    $sitePath = $siteUri.AbsolutePath.TrimEnd("/")
    if ([string]::IsNullOrWhiteSpace($sitePath)) {
        $sitePath = "/"
    }

    $normalizedSiteUrl = if ($sitePath -eq "/") {
        "https://$($siteUri.Host)/"
    }
    else {
        "https://$($siteUri.Host)$sitePath"
    }

    Write-Host "Site: $normalizedSiteUrl" -ForegroundColor Green

    Write-Section "Checking Microsoft Graph module"

    $installedModule = Get-Module -ListAvailable -Name $requiredModule |
        Sort-Object Version -Descending |
        Select-Object -First 1

    if ($null -eq $installedModule) {
        Write-Host "$requiredModule is not installed. Installing it for the current user..." -ForegroundColor Yellow
        try {
            Install-Module `
                -Name $requiredModule `
                -Repository PSGallery `
                -Scope CurrentUser `
                -Force `
                -AllowClobber `
                -ErrorAction Stop
        }
        catch {
            Write-Host $_.Exception.Message -ForegroundColor Red
            Stop-WithInstructions `
                -Problem "PowerShell could not install $requiredModule." `
                -Fix @(
                    "Make sure this computer can reach https://www.powershellgallery.com.",
                    "Run: Install-Module Microsoft.Graph.Authentication -Scope CurrentUser -Force",
                    "If prompted to trust PSGallery, answer Y.",
                    "For Windows PowerShell 5.1 module-install errors, first run: Install-Module PowerShellGet -Scope CurrentUser -Force -AllowClobber",
                    "If your company uses a proxy, ask IT to allow PowerShell Gallery access."
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
                "For Windows PowerShell 5.1, confirm .NET Framework 4.7.2 or newer is installed.",
                "Then rerun this script. PowerShell 7 is recommended if loading still fails."
            )
    }

    $loadedModule = Get-Module -Name $requiredModule | Select-Object -First 1
    Write-Host "$requiredModule $($loadedModule.Version): OK" -ForegroundColor Green

    Write-Section "Signing in to Microsoft Graph"
    Write-Host "A browser window may open." -ForegroundColor Yellow
    Write-Host "Sign in with a Global Administrator or SharePoint Administrator for $($siteUri.Host)." -ForegroundColor Yellow
    Write-Host "The account must also be a site admin for the target site collection; tenant admin roles alone do not grant site access." -ForegroundColor Yellow

    try {
        Connect-MgGraph `
            -Scopes $requiredGraphScope `
            -ContextScope Process `
            -NoWelcome `
            -ErrorAction Stop
        $graphConnected = $true
    }
    catch {
        Write-Host $_.Exception.Message -ForegroundColor Red
        Stop-WithInstructions `
            -Problem "Microsoft Graph sign-in failed." `
            -Fix @(
                "Run the script again and complete the browser sign-in.",
                "Use a Global Administrator or SharePoint Administrator account in the target tenant.",
                "If no browser opens, run: Connect-MgGraph -Scopes '$requiredGraphScope' -UseDeviceAuthentication"
            )
    }

    $graphContext = Get-MgContext
    if ($null -eq $graphContext -or $graphContext.Scopes -notcontains $requiredGraphScope) {
        Stop-WithInstructions `
            -Problem "The Graph access token does not contain the required $requiredGraphScope permission." `
            -Fix @(
                "A Global Administrator must grant admin consent for Microsoft Graph delegated $requiredGraphScope.",
                "Disconnect with Disconnect-MgGraph and rerun this script.",
                "Accept the permission request during browser sign-in."
            )
    }

    $graphClientId = [string]$graphContext.ClientId
    Write-Host "Signed in as:      $($graphContext.Account)" -ForegroundColor Green
    Write-Host "Tenant ID:         $($graphContext.TenantId)" -ForegroundColor Green
    Write-Host "Sign-in client ID: $graphClientId" -ForegroundColor DarkGray

    Write-Section "Finding the SharePoint site"

    $lookupUri = 'https://graph.microsoft.com/v1.0/sites/{0}:{1}?$select=id,displayName,webUrl,siteCollection' -f `
        $siteUri.Host, $sitePath

    try {
        $site = Invoke-MgGraphRequest -Method GET -Uri $lookupUri -ErrorAction Stop
    }
    catch {
        Write-GraphFailureDetails -ErrorRecord $_ -Method "GET" -RequestUri $lookupUri
        $guidance = Get-GraphErrorGuidance -ErrorRecord $_ -NormalizedSiteUrl $normalizedSiteUrl -Operation "FindSite"
        Stop-WithInstructions -Problem "Microsoft Graph could not find or access the SharePoint site." -Fix $guidance
    }

    if ([string]::IsNullOrWhiteSpace([string]$site.id)) {
        Stop-WithInstructions `
            -Problem "Microsoft Graph returned the site without a site ID." `
            -Fix @(
                "Verify that $normalizedSiteUrl is a SharePoint site rather than a page, document, or sharing link.",
                "Copy the site home-page URL and rerun the script."
            )
    }

    Write-Host "Resolved site: $($site.webUrl)" -ForegroundColor Green
    Write-Host "Site ID:       $($site.id)" -ForegroundColor DarkGray

    $requestedSite = $site
    $requestedSiteWasSubsite = $false
    $siteCollection = Get-ObjectPropertyValue -InputObject $site -Name "siteCollection"
    if ($null -eq $siteCollection) {
        # A Graph site ID is hostname,siteCollectionId,webId. Addressing a site
        # with only hostname and siteCollectionId returns that collection's root.
        $siteIdParts = ([string]$site.id).Split(",")
        $siteCollectionId = [guid]::Empty
        if ($siteIdParts.Count -ne 3 -or
            -not [guid]::TryParse($siteIdParts[1], [ref]$siteCollectionId)) {
            Stop-WithInstructions `
                -Problem "The entered URL is a subsite, but its containing site collection could not be determined from Graph site ID '$($site.id)'." `
                -Fix @(
                    "Give support the entered URL and Graph site ID printed above.",
                    "As a workaround, enter the containing site collection's root URL from SharePoint admin center > Active sites."
                )
        }

        $siteCollectionLookupUri = 'https://graph.microsoft.com/v1.0/sites/{0},{1}?$select=id,displayName,webUrl,siteCollection' -f `
            $siteUri.Host, $siteCollectionId

        try {
            $siteCollectionRoot = Invoke-MgGraphRequest `
                -Method GET `
                -Uri $siteCollectionLookupUri `
                -ErrorAction Stop
        }
        catch {
            Write-GraphFailureDetails -ErrorRecord $_ -Method "GET" -RequestUri $siteCollectionLookupUri
            $guidance = Get-GraphErrorGuidance `
                -ErrorRecord $_ `
                -NormalizedSiteUrl $normalizedSiteUrl `
                -Operation "FindSite"
            Stop-WithInstructions `
                -Problem "Microsoft Graph found the subsite but could not resolve its containing site-collection root." `
                -Fix $guidance
        }

        $resolvedSiteCollection = Get-ObjectPropertyValue `
            -InputObject $siteCollectionRoot `
            -Name "siteCollection"
        if ($null -eq $resolvedSiteCollection) {
            Stop-WithInstructions `
                -Problem "Microsoft Graph did not identify '$($siteCollectionRoot.webUrl)' as a site-collection root." `
                -Fix @(
                    "Give support the entered URL, resolved URL, and Graph site IDs printed above.",
                    "As a workaround, enter the site collection's root URL from SharePoint admin center > Active sites."
                )
        }

        $site = $siteCollectionRoot
        $requestedSiteWasSubsite = $true
        Write-Host "Entered URL is a SharePoint subsite." -ForegroundColor Yellow
        Write-Host "Containing site collection: $($site.webUrl)" -ForegroundColor Yellow
        Write-Host "Site collection ID:         $($site.id)" -ForegroundColor DarkGray
    }

    if (-not $Force) {
        Write-Host ""
        Write-Host "This will grant the following application permission:" -ForegroundColor Yellow
        Write-Host "  Target enterprise application: $applicationDisplayName" -ForegroundColor Yellow
        Write-Host "  Application (client) ID:        $applicationId" -ForegroundColor Yellow
        Write-Host "  Microsoft Graph Sites.Selected application role ID:" -ForegroundColor Yellow
        Write-Host "                                $sitesSelectedApplicationPermissionId" -ForegroundColor Yellow
        Write-Host "  Requested site:             $($requestedSite.webUrl)" -ForegroundColor Yellow
        Write-Host "  Grant scope (site collection): $($site.webUrl)" -ForegroundColor Yellow
        Write-Host "  Permission:  Manage" -ForegroundColor Yellow
        if ($requestedSiteWasSubsite) {
            Write-Host "  IMPORTANT: Sites.Selected cannot be scoped to only this subsite." -ForegroundColor Yellow
            Write-Host "             Manage access applies to the entire containing site collection." -ForegroundColor Yellow
        }
        Write-Host ""

        $confirmation = Read-Host "Type GRANT to continue"
        if ($confirmation -cne "GRANT") {
            Write-Host "No changes were made." -ForegroundColor Yellow
            return
        }
    }

    Write-Section "Checking existing site permissions"

    $permissionsUri = "https://graph.microsoft.com/v1.0/sites/$($site.id)/permissions"

    try {
        $permissionResponse = Invoke-MgGraphRequest -Method GET -Uri $permissionsUri -ErrorAction Stop
    }
    catch {
        Write-GraphFailureDetails -ErrorRecord $_ -Method "GET" -RequestUri $permissionsUri
        $guidance = Get-GraphErrorGuidance -ErrorRecord $_ -NormalizedSiteUrl $normalizedSiteUrl -Operation "ListPermissions"
        Stop-WithInstructions -Problem "Microsoft Graph could not read the site's application permissions." -Fix $guidance
    }

    $existingPermissions = @($permissionResponse.value | Where-Object {
        $candidate = $_
        $v2Match = @($candidate.grantedToIdentitiesV2) | Where-Object {
            $null -ne $_.application -and $_.application.id -eq $applicationId
        }
        $legacyMatch = @($candidate.grantedToIdentities) | Where-Object {
            $null -ne $_.application -and $_.application.id -eq $applicationId
        }
        $v2Match.Count -gt 0 -or $legacyMatch.Count -gt 0
    })

    if ($existingPermissions.Count -gt 1) {
        Stop-WithInstructions `
            -Problem "More than one existing permission was found for target application ID $applicationId. The script will not guess which record to change." `
            -Fix @(
                "Have a SharePoint administrator review the permission records on $($site.webUrl).",
                "Remove duplicate grants, leaving one grant for application ID $applicationId.",
                "Then rerun the script."
            )
    }

    if ($existingPermissions.Count -eq 1) {
        $permission = $existingPermissions[0]
        $currentRoles = @($permission.roles | ForEach-Object { ([string]$_).ToLowerInvariant() })

        if ($currentRoles -contains $requiredRole) {
            Write-Host "The application already has Manage permission on this site. No change was needed." -ForegroundColor Green
        }
        else {
            Stop-WithInstructions `
                -Problem "Target application ID $applicationId already has role(s) '$($permission.roles -join ', ')' instead of Manage. Microsoft Graph does not support updating this grant with the script's delegated user token." `
                -Fix @(
                    "Have an administrator remove or update permission ID $($permission.id) using an app-only Microsoft Graph connection with Sites.FullControl.All APPLICATION permission.",
                    "Target enterprise application: '$applicationDisplayName'; application (client) ID: $applicationId.",
                    "After the old grant is removed, rerun this script to create the Manage grant."
                )
        }
    }
    else {
        Write-Host "No existing grant was found. Creating Manage permission..." -ForegroundColor Yellow

        $createBody = @{
            roles = @($requiredRole)
            grantedToIdentities = @(
                @{
                    application = @{
                        id = $applicationId
                        displayName = $applicationDisplayName
                    }
                }
            )
        } | ConvertTo-Json -Depth 5

        try {
            $permission = Invoke-MgGraphRequest `
                -Method POST `
                -Uri $permissionsUri `
                -ContentType "application/json" `
                -Body $createBody `
                -ErrorAction Stop
        }
        catch {
            Write-GraphFailureDetails -ErrorRecord $_ -Method "POST" -RequestUri $permissionsUri
            $guidance = Get-GraphErrorGuidance -ErrorRecord $_ -NormalizedSiteUrl $normalizedSiteUrl -Operation "CreatePermission"
            Stop-WithInstructions -Problem "Microsoft Graph could not grant Manage permission to target application ID $applicationId." -Fix $guidance
        }
    }

    Write-Section "Verifying the result"

    try {
        # Listing the permission collection supports delegated
        # Sites.FullControl.All. GET /permissions/{permission-id} does not.
        $verificationResponse = Invoke-MgGraphRequest -Method GET -Uri $permissionsUri -ErrorAction Stop
    }
    catch {
        Write-GraphFailureDetails -ErrorRecord $_ -Method "GET" -RequestUri $permissionsUri
        Stop-WithInstructions `
            -Problem "The grant operation completed, but the script could not verify the saved permission." `
            -Fix @(
                "Rerun this script with the same site URL; it is safe to rerun.",
                "The script will detect the existing grant instead of creating a duplicate."
            )
    }

    $verifiedMatches = @($verificationResponse.value | Where-Object {
        $candidate = $_
        $v2Match = @($candidate.grantedToIdentitiesV2) | Where-Object {
            $null -ne $_.application -and $_.application.id -eq $applicationId
        }
        $legacyMatch = @($candidate.grantedToIdentities) | Where-Object {
            $null -ne $_.application -and $_.application.id -eq $applicationId
        }
        $v2Match.Count -gt 0 -or $legacyMatch.Count -gt 0
    })

    if ($verifiedMatches.Count -ne 1) {
        Stop-WithInstructions `
            -Problem "Verification found $($verifiedMatches.Count) permission records for target application ID $applicationId; exactly one was expected." `
            -Fix @(
                "Rerun this script once; it is safe to rerun.",
                "If the result is unchanged, have a SharePoint administrator review the site's application permission records."
            )
    }

    $verifiedPermission = $verifiedMatches[0]

    $verifiedRoles = @($verifiedPermission.roles | ForEach-Object { ([string]$_).ToLowerInvariant() })
    $verifiedApplications = @($verifiedPermission.grantedToIdentitiesV2) + @($verifiedPermission.grantedToIdentities)
    $applicationVerified = @($verifiedApplications | Where-Object {
        $null -ne $_.application -and $_.application.id -eq $applicationId
    }).Count -gt 0

    if ($verifiedRoles -notcontains $requiredRole -or -not $applicationVerified) {
        Stop-WithInstructions `
            -Problem "The saved permission did not match the expected application and Manage role." `
            -Fix @(
                "Rerun the script once.",
                "If it fails again, give your administrator the site URL and permission ID shown below: $($permission.id)"
            )
    }

    Write-Host ""
    Write-Host "SUCCESS" -ForegroundColor Green
    Write-Host "  Requested site:            $($requestedSite.webUrl)" -ForegroundColor Green
    Write-Host "  Granted site collection:   $($site.webUrl)" -ForegroundColor Green
    Write-Host "  Target application:          $applicationDisplayName" -ForegroundColor Green
    Write-Host "  Application (client) ID:     $applicationId" -ForegroundColor Green
    Write-Host "  Permission:  Manage" -ForegroundColor Green
    Write-Host "  Permission ID: $($verifiedPermission.id)" -ForegroundColor DarkGray
    Write-Host ""
    Write-Host "The script is safe to rerun for this site." -ForegroundColor Green
}
catch {
    Write-Host ""
    Write-Host "Technical error:" -ForegroundColor DarkRed
    Write-Host $_.Exception.Message -ForegroundColor DarkRed
    return
}
finally {
    if ($graphConnected) {
        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    }
}
}
