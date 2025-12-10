[CmdletBinding()]
Param(
    # gather permission requests but don't create any AppId nor ServicePrincipal
    [switch] $DryRun = $false,
    # Azure environment name for Microsoft Graph
    [ValidateSet("Global", "USGov", "USGovDoD", "Germany", "China")]
    [string] $AzureEnvironment = "Global",

    [ValidateSet(
        "UnitedStates",
        "Preview(UnitedStates)",
        "Europe",
        "EMEA",
        "Asia",
        "Australia",
        "Japan",
        "SouthAmerica",
        "India",
        "Canada",
        "UnitedKingdom",
        "France"
    )]
    [string] $TenantLocation = "UnitedStates"
)

function ensureModules {
    $dependencies = @(
        @{ 
            Name = "Microsoft.Graph.Applications"
            Version = [Version]"2.0.0"
            InstallWith = "Install-Module -Name Microsoft.Graph.Applications -Scope CurrentUser -Force"
        },
        @{ 
            Name = "Microsoft.Graph.Authentication"
            Version = [Version]"2.0.0"
            InstallWith = "Install-Module -Name Microsoft.Graph.Authentication -Scope CurrentUser -Force"
        }
    )
    $missingDependencies = $false
    $dependencies | ForEach-Object -Process {
        $moduleName = $_.Name
        $deps = (Get-Module -ListAvailable -Name $moduleName | Sort-Object -Descending -Property Version)
        if ($null -eq $deps) {
            Write-Host @"
ERROR: Required module not installed; install from PowerShell prompt with:
>>  $($_.InstallWith) -MinimumVersion $($_.Version)
"@
            $missingDependencies = $true
            return
        }
        $dep = $deps[0]
        if ($dep.Version -lt $_.Version) {
            Write-Host @"
ERROR: Required module installed but does not meet minimal required version:
       found: $($dep.Version), required: >= $($_.Version); to fix, please run:
>>  Update-Module $($_.Name) -Scope CurrentUser -RequiredVersion $($_.Version)
"@
            $missingDependencies = $true
            return
        }
        Import-Module $moduleName -MinimumVersion $_.Version
    }
    if ($missingDependencies) {
        throw "Missing required dependencies!"
    }
}

function checkIsElevated {
    $id = [System.Security.Principal.WindowsIdentity]::GetCurrent()
    $p = New-Object System.Security.Principal.WindowsPrincipal($id)
    if ($p.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)) { 
        Write-Output $true 
    }
    else { 
        Write-Output $false 
    }
}

function connectGraph {
    Write-Host @"

Connecting to Microsoft Graph.
Please log in, using your Dynamics365 / Power Platform tenant ADMIN credentials:

"@
    try {
        # Required scopes for creating service principals
        Connect-MgGraph -Scopes "Application.ReadWrite.All", "Directory.ReadWrite.All" -Environment $AzureEnvironment -NoWelcome -ErrorAction Stop
    }
    catch {
        throw "Failed to login: $($_.Exception.Message)"
    }
    return Get-MgContext
}

function reconnectGraph {
    try {
        $context = Get-MgContext
        if ($null -eq $context) {
            $context = connectGraph
        }
        elseif ($context.Environment -ne $AzureEnvironment) {
            Disconnect-MgGraph | Out-Null
            $context = connectGraph
        }
    }
    catch {
        $context = connectGraph
    }
    
    $tenantId = $context.TenantId

    Write-Host @"
Connected to tenant: $($tenantId)

"@
    return $tenantId
}

function getAppConsentUri($tenantId) {
    "https://login.microsoftonline.com/$tenantId/oauth2/authorize?client_id=a86b9632-42bf-4dfe-83c8-bbc95145504b&response_type=code&redirect_uri=https://manager.onakoyago.com/AdminConsent&nonce=doesntmatter&resource=https://graph.microsoft.com&prompt=admin_consent"
}

# Main script execution

$hasAdmin = checkIsElevated
if (!$hasAdmin) {
    throw "This action requires administrator privileges."
}

# validate required modules are installed
ensureModules

$ErrorActionPreference = "Stop"

$tenantId = reconnectGraph
$context = Get-MgContext
$spnDisplayName = "BCO akoyaGO Integration"
$appId = "a86b9632-42bf-4dfe-83c8-bbc95145504b"

if (!$DryRun) {
    # Check if SPN already exists
    $existingSpn = Get-MgServicePrincipal -Filter "appId eq '$appId'" -ErrorAction SilentlyContinue
    
    if ($null -ne $existingSpn) {
        Write-Host "Service Principal already exists with objectId: $($existingSpn.Id)"
        $spnId = $existingSpn.Id
    }
    else {
        $spnParams = @{
            AppId = $appId
            DisplayName = $spnDisplayName
            AccountEnabled = $true
            AppRoleAssignmentRequired = $true
            Tags = @("WindowsAzureActiveDirectoryIntegratedApp")
        }
        
        $spn = New-MgServicePrincipal -BodyParameter $spnParams
        $spnId = $spn.Id
        Write-Host "Created SPN '$spnDisplayName' with objectId: $spnId"
    }
}
else {
    Write-Host "Skipping SPN creation because DryRun is 'true'"
}

Write-Host @"

#################################################################

Copy and paste the following URL in a browser to grant consent:

"@

Write-Host $(getAppConsentUri $tenantId)

Write-Host @"

#################################################################


Done.

"@

# Disconnect when finished
Disconnect-MgGraph | Out-Null
