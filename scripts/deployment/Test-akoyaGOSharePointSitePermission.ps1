# Test-akoyaGOSharePointSitePermission.ps1
# Version: 0.1
#
# Read-only companion to Grant-akoyaGOSharePointSitePermission.ps1.
# It reports which applications have been granted access to one SharePoint
# site collection through the Sites.Selected permission model, and whether the
# akoyaGO integration application has the expected Manage grant.
# It never creates, changes, or removes anything.
#
# Beginner-friendly usage directly from the akoyaGO public repository:
#   iex ((New-Object System.Net.WebClient).DownloadString('https://raw.githubusercontent.com/akoyago/public/refs/heads/main/scripts/deployment/Test-akoyaGOSharePointSitePermission.ps1'))
#
# Optional: supply the site URL on the same line instead of being prompted:
#   $global:akoyaGOSiteUrl = 'https://contoso.sharepoint.com/sites/example'; iex ((New-Object System.Net.WebClient).DownloadString('https://raw.githubusercontent.com/akoyago/public/refs/heads/main/scripts/deployment/Test-akoyaGOSharePointSitePermission.ps1'))
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
$expectedRole = "manage"
# Roles that include everything Manage allows.
$higherRoles = @("fullcontrol", "owner")
$requiredGraphScope = "Sites.FullControl.All"
$requiredModule = "Microsoft.Graph.Authentication"
$minimumWindowsPowerShellVersion = [version]"5.1"
$minimumPowerShellCoreVersion = [version]"7.0"
$graphConnected = $false

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
    Write-Host "COULD NOT CHECK: $Problem" -ForegroundColor Red
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
        [Parameter(Mandatory = $true)][string]$RequestUri
    )

    Write-Host "Graph request: GET $RequestUri" -ForegroundColor DarkRed
    Write-Host $ErrorRecord.Exception.Message -ForegroundColor Red
    $errorBody = [string](Get-ObjectPropertyValue -InputObject $ErrorRecord.ErrorDetails -Name "Message")
    if (-not [string]::IsNullOrWhiteSpace($errorBody) -and
        $errorBody -ne $ErrorRecord.Exception.Message) {
        Write-Host "Graph error body: $errorBody" -ForegroundColor DarkRed
    }
}

# Returns the applications named on one permission record. Graph returns the
# same identities in grantedToIdentitiesV2 and the deprecated
# grantedToIdentities, so they are de-duplicated by application ID.
function Get-PermissionApplications {
    param([AllowNull()][object]$Permission)

    $found = @{}
    foreach ($propertyName in @("grantedToIdentitiesV2", "grantedToIdentities")) {
        foreach ($identity in @(Get-ObjectPropertyValue -InputObject $Permission -Name $propertyName)) {
            $application = Get-ObjectPropertyValue -InputObject $identity -Name "application"
            $id = [string](Get-ObjectPropertyValue -InputObject $application -Name "id")
            if ([string]::IsNullOrWhiteSpace($id)) {
                continue
            }
            $key = $id.ToLowerInvariant()
            if (-not $found.ContainsKey($key)) {
                $found[$key] = [pscustomobject]@{
                    Id          = $id
                    DisplayName = [string](Get-ObjectPropertyValue -InputObject $application -Name "displayName")
                }
            }
        }
    }

    return @($found.Values)
}

function Get-PermissionRoles {
    param([AllowNull()][object]$Permission)

    return @(
        @(Get-ObjectPropertyValue -InputObject $Permission -Name "roles") |
            Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } |
            ForEach-Object { ([string]$_).ToLowerInvariant() }
    )
}

try {
    Write-Section "Checking PowerShell"

    $isWindowsPowerShell = $PSVersionTable.PSEdition -eq "Desktop"
    $isPowerShellCore = $PSVersionTable.PSEdition -eq "Core"

    if (($isWindowsPowerShell -and $PSVersionTable.PSVersion -lt $minimumWindowsPowerShellVersion) -or
        ($isPowerShellCore -and $PSVersionTable.PSVersion -lt $minimumPowerShellCoreVersion) -or
        (-not $isWindowsPowerShell -and -not $isPowerShellCore)) {
        Stop-WithInstructions `
            -Problem "This PowerShell is not supported: $($PSVersionTable.PSEdition) $($PSVersionTable.PSVersion)." `
            -Fix @(
                "Use Windows PowerShell 5.1 or PowerShell 7.",
                "Install PowerShell 7 with: winget install --id Microsoft.PowerShell --source winget"
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
            -Fix @("Use the complete URL, for example: https://contoso.sharepoint.com/sites/example")
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
                    "Run: Install-Module Microsoft.Graph.Authentication -Scope CurrentUser -Force"
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
                "Then rerun this script. PowerShell 7 is recommended if loading still fails."
            )
    }

    Write-Host "${requiredModule}: OK" -ForegroundColor Green

    Write-Section "Signing in to Microsoft Graph"
    Write-Host "A browser window may open." -ForegroundColor Yellow
    Write-Host "Sign in with a Global Administrator or SharePoint Administrator for $($siteUri.Host)." -ForegroundColor Yellow
    Write-Host "The account must also be a site admin for the target site collection." -ForegroundColor Yellow
    Write-Host "This script only reads. It makes no changes." -ForegroundColor Yellow

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
                "Microsoft Graph requires delegated $requiredGraphScope even to read a site's application grants.",
                "A Global Administrator must grant admin consent for it, then rerun this script."
            )
    }

    Write-Host "Signed in as: $($graphContext.Account)" -ForegroundColor Green
    Write-Host "Tenant ID:    $($graphContext.TenantId)" -ForegroundColor Green

    Write-Section "Finding the SharePoint site"

    $lookupUri = 'https://graph.microsoft.com/v1.0/sites/{0}:{1}?$select=id,displayName,webUrl,siteCollection' -f `
        $siteUri.Host, $sitePath

    try {
        $site = Invoke-MgGraphRequest -Method GET -Uri $lookupUri -ErrorAction Stop
    }
    catch {
        Write-GraphFailureDetails -ErrorRecord $_ -RequestUri $lookupUri
        Stop-WithInstructions `
            -Problem "Microsoft Graph could not find or access the SharePoint site." `
            -Fix @(
                "Open $normalizedSiteUrl in a browser and verify that the site exists and the URL is the site home page.",
                "Make sure you signed in to the tenant that owns the site.",
                "If the error is Access denied: in SharePoint admin center > Active sites, select the site, open Membership > Site admins, and add the account printed under 'Signed in as'."
            )
    }

    $siteId = [string](Get-ObjectPropertyValue -InputObject $site -Name "id")
    if ([string]::IsNullOrWhiteSpace($siteId)) {
        Stop-WithInstructions `
            -Problem "Microsoft Graph returned the site without a site ID." `
            -Fix @("Verify that $normalizedSiteUrl is a SharePoint site rather than a page, document, or sharing link.")
    }

    $requestedSiteUrl = [string](Get-ObjectPropertyValue -InputObject $site -Name "webUrl")
    Write-Host "Resolved site: $requestedSiteUrl" -ForegroundColor Green

    # Sites.Selected grants live on the site collection root. Graph does not
    # support listing permissions on a subsite, so resolve the root first.
    if ($null -eq (Get-ObjectPropertyValue -InputObject $site -Name "siteCollection")) {
        $siteIdParts = $siteId.Split(",")
        $siteCollectionId = [guid]::Empty
        if ($siteIdParts.Count -ne 3 -or
            -not [guid]::TryParse($siteIdParts[1], [ref]$siteCollectionId)) {
            Stop-WithInstructions `
                -Problem "The entered URL is a subsite, but its containing site collection could not be determined from Graph site ID '$siteId'." `
                -Fix @("Enter the containing site collection's root URL from SharePoint admin center > Active sites.")
        }

        $siteCollectionLookupUri = 'https://graph.microsoft.com/v1.0/sites/{0},{1}?$select=id,displayName,webUrl,siteCollection' -f `
            $siteUri.Host, $siteCollectionId

        try {
            $site = Invoke-MgGraphRequest -Method GET -Uri $siteCollectionLookupUri -ErrorAction Stop
        }
        catch {
            Write-GraphFailureDetails -ErrorRecord $_ -RequestUri $siteCollectionLookupUri
            Stop-WithInstructions `
                -Problem "Microsoft Graph found the subsite but could not resolve its containing site-collection root." `
                -Fix @("Enter the containing site collection's root URL from SharePoint admin center > Active sites.")
        }

        $siteId = [string](Get-ObjectPropertyValue -InputObject $site -Name "id")
        Write-Host "Entered URL is a SharePoint subsite." -ForegroundColor Yellow
        Write-Host "Checking its site collection: $(Get-ObjectPropertyValue -InputObject $site -Name 'webUrl')" -ForegroundColor Yellow
    }

    $siteCollectionUrl = [string](Get-ObjectPropertyValue -InputObject $site -Name "webUrl")

    Write-Section "Reading application permissions on the site"

    $permissionsUri = "https://graph.microsoft.com/v1.0/sites/$siteId/permissions"
    $permissions = @()
    $pageUri = $permissionsUri
    while (-not [string]::IsNullOrWhiteSpace($pageUri)) {
        try {
            $page = Invoke-MgGraphRequest -Method GET -Uri $pageUri -ErrorAction Stop
        }
        catch {
            Write-GraphFailureDetails -ErrorRecord $_ -RequestUri $pageUri
            Stop-WithInstructions `
                -Problem "Microsoft Graph could not read the site's application permissions." `
                -Fix @(
                    "Use an account that is a Global Administrator or SharePoint Administrator and also a site admin of $siteCollectionUrl.",
                    "Approve the Microsoft Graph delegated permission $requiredGraphScope when prompted.",
                    "Run Disconnect-MgGraph and rerun the script."
                )
        }
        $permissions += @(Get-ObjectPropertyValue -InputObject $page -Name "value")
        $pageUri = [string](Get-ObjectPropertyValue -InputObject $page -Name "@odata.nextLink")
    }

    $rows = @()
    foreach ($permission in $permissions) {
        $applications = @(Get-PermissionApplications -Permission $permission)
        if ($applications.Count -eq 0) {
            continue
        }

        $permissionId = [string](Get-ObjectPropertyValue -InputObject $permission -Name "id")
        $roles = @(Get-PermissionRoles -Permission $permission)
        $roleSource = "list"

        if ($roles.Count -eq 0 -and -not [string]::IsNullOrWhiteSpace($permissionId)) {
            # Graph has been observed to omit 'roles' when listing a site's
            # permissions. Reading the single permission returns them, but
            # Microsoft documents that call as application-only, so a failure
            # here is expected in some tenants and is reported, not fatal.
            $roleSource = "unknown"
            $singleUri = "$permissionsUri/$([Uri]::EscapeDataString($permissionId))"
            try {
                $single = Invoke-MgGraphRequest -Method GET -Uri $singleUri -ErrorAction Stop
                $roles = @(Get-PermissionRoles -Permission $single)
                if ($roles.Count -gt 0) {
                    $roleSource = "single"
                }
            }
            catch {
                $roles = @()
            }
        }

        foreach ($application in $applications) {
            $rows += [pscustomobject]@{
                Application   = $application.DisplayName
                ApplicationId = $application.Id
                Roles         = if ($roles.Count -gt 0) { $roles -join ", " } else { "(not returned by Graph)" }
                RoleList      = $roles
                RoleSource    = $roleSource
                PermissionId  = $permissionId
            }
        }
    }

    Write-Host "Site collection: $siteCollectionUrl" -ForegroundColor Green
    Write-Host "Site ID:         $siteId" -ForegroundColor DarkGray
    Write-Host ""

    if ($rows.Count -eq 0) {
        Write-Host "No application has a Sites.Selected grant on this site collection." -ForegroundColor Yellow
    }
    else {
        Write-Host "Applications granted access to this site collection:" -ForegroundColor Cyan
        foreach ($row in $rows) {
            $marker = if ($row.ApplicationId -eq $applicationId) { " <-- akoyaGO" } else { "" }
            Write-Host ""
            Write-Host "  Application:    $($row.Application)$marker"
            Write-Host "  Application ID: $($row.ApplicationId)"
            Write-Host "  Role(s):        $($row.Roles)"
            Write-Host "  Permission ID:  $($row.PermissionId)" -ForegroundColor DarkGray
        }
    }

    Write-Section "Result for $applicationDisplayName"

    $targetRows = @($rows | Where-Object { $_.ApplicationId -eq $applicationId })
    $targetRoles = @($targetRows | ForEach-Object { $_.RoleList } | Where-Object { $null -ne $_ } | Select-Object -Unique)

    Write-Host "  Target application:      $applicationDisplayName"
    Write-Host "  Application (client) ID: $applicationId"
    Write-Host "  Site collection:         $siteCollectionUrl"
    Write-Host "  Expected role:           Manage"
    Write-Host ""

    if ($targetRows.Count -eq 0) {
        Write-Host "NOT GRANTED" -ForegroundColor Red
        Write-Host "  The application has no grant on this site collection." -ForegroundColor Red
        Write-Host "  Run Grant-akoyaGOSharePointSitePermission.ps1 for this site." -ForegroundColor Yellow
    }
    elseif ($targetRoles.Count -eq 0) {
        Write-Host "GRANT FOUND, ROLE NOT CONFIRMED" -ForegroundColor Yellow
        Write-Host "  A grant exists for the application, but Microsoft Graph did not return its role to this sign-in." -ForegroundColor Yellow
        Write-Host "  The grant script only ever creates a Manage grant, so this is normally fine." -ForegroundColor Yellow
        Write-Host "  To see the role, read permission ID $($targetRows[0].PermissionId) with an app-only" -ForegroundColor Yellow
        Write-Host "  Microsoft Graph connection that has the Sites.FullControl.All application permission." -ForegroundColor Yellow
    }
    elseif ($targetRoles -contains $expectedRole) {
        Write-Host "VERIFIED" -ForegroundColor Green
        Write-Host "  The application has Manage permission on this site collection." -ForegroundColor Green
    }
    elseif (@($targetRoles | Where-Object { $higherRoles -contains $_ }).Count -gt 0) {
        Write-Host "GRANTED WITH A HIGHER ROLE" -ForegroundColor Yellow
        Write-Host "  The application has role(s) '$($targetRoles -join ', ')', which include everything Manage allows." -ForegroundColor Yellow
        Write-Host "  This was not set by the akoyaGO grant script, which only creates Manage." -ForegroundColor Yellow
    }
    else {
        Write-Host "WRONG ROLE" -ForegroundColor Red
        Write-Host "  The application has role(s) '$($targetRoles -join ', ')' instead of Manage." -ForegroundColor Red
        Write-Host "  Have an administrator remove permission ID $($targetRows[0].PermissionId), then rerun the grant script." -ForegroundColor Yellow
    }

    if ($targetRows.Count -gt 1) {
        Write-Host ""
        Write-Host "  WARNING: $($targetRows.Count) separate grants exist for this application; exactly one is expected." -ForegroundColor Yellow
    }

    Write-Host ""
    Write-Host "This confirms the site grant only. The application also needs the Sites.Selected" -ForegroundColor DarkGray
    Write-Host "application permission with admin consent, shown in Microsoft Entra admin center >" -ForegroundColor DarkGray
    Write-Host "Enterprise applications > $applicationDisplayName > Permissions." -ForegroundColor DarkGray
    Write-Host "No changes were made." -ForegroundColor DarkGray
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
