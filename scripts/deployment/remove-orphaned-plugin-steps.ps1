<#
.SYNOPSIS
    Removes orphaned plugin steps and plugin types from Dynamics 365 environment.

.DESCRIPTION
    This script compares plugin steps in the exported plugin-steps.json file with the target
    Dynamics 365 environment and removes any steps that exist in the target but not in the export.
    Also removes orphaned plugin types that exist in the target but not in the export.
    Only processes steps for the AkoyaGo.Plugins assembly.
    Comparison is done based on logical names/properties, not GUIDs.

.PARAMETER EnvironmentUrl
    The URL of the target Dynamics 365 environment (e.g., https://org.crm.dynamics.com)

.PARAMETER ClientId
    The Azure AD Application (Client) ID for service principal authentication

.PARAMETER ClientSecret
    The Client Secret for the service principal

.PARAMETER PluginStepsJsonPath
    Path to the plugin-steps.json file (defaults to looking in artifacts directory)

.EXAMPLE
    .\Remove-OrphanedPluginSteps.ps1 -EnvironmentUrl "https://org.crm.dynamics.com" `
        -ClientId "your-client-id" `
        -ClientSecret "your-client-secret" `
        -PluginStepsJsonPath "$(System.ArtifactsDirectory)/drop/plugin-steps.json"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)]
    [string]$EnvironmentUrl,
    
    [Parameter(Mandatory=$true)]
    [string]$ClientId,
    
    [Parameter(Mandatory=$true)]
    [string]$ClientSecret,
    
    [Parameter(Mandatory=$false)]
    [string]$PluginStepsJsonPath
)

# =============================================================================
# HANDLE DEFAULT PATH WITH BETTER DETECTION
# =============================================================================

if ([string]::IsNullOrEmpty($PluginStepsJsonPath)) {
    Write-Host "No explicit path provided, searching for plugin-steps.json..." -ForegroundColor Gray
    
    # Try multiple common locations
    $searchPaths = @(
        "$env:SYSTEM_ARTIFACTSDIRECTORY\drop\plugin-steps.json",
        "$env:SYSTEM_ARTIFACTSDIRECTORY\plugin-steps.json",
        "$env:SYSTEM_DEFAULTWORKINGDIRECTORY\drop\plugin-steps.json",
        "$env:SYSTEM_DEFAULTWORKINGDIRECTORY\plugin-steps.json",
        "$env:AGENT_RELEASEDIRECTORY\drop\plugin-steps.json",
        "$env:BUILD_ARTIFACTSTAGINGDIRECTORY\plugin-steps.json"
    )
    
    $foundPath = $null
    
    foreach ($path in $searchPaths) {
        if ($path -and (Test-Path $path)) {
            $foundPath = $path
            Write-Host "Found plugin-steps.json at: $foundPath" -ForegroundColor Green
            break
        }
    }
    
    if (-not $foundPath) {
        # Last resort: search recursively
        Write-Host "Standard paths not found, searching recursively..." -ForegroundColor Yellow
        
        $searchRoots = @($env:SYSTEM_ARTIFACTSDIRECTORY, $env:SYSTEM_DEFAULTWORKINGDIRECTORY, "D:\a")
        
        foreach ($root in $searchRoots) {
            if ($root -and (Test-Path $root)) {
                $found = Get-ChildItem -Path $root -Recurse -Filter "plugin-steps.json" -ErrorAction SilentlyContinue | Select-Object -First 1
                if ($found) {
                    $foundPath = $found.FullName
                    Write-Host "Found plugin-steps.json at: $foundPath" -ForegroundColor Green
                    break
                }
            }
        }
    }
    
    if ($foundPath) {
        $PluginStepsJsonPath = $foundPath
    }
    else {
        Write-Error "Could not find plugin-steps.json in any standard location."
        Write-Host "`nSearched paths:" -ForegroundColor Yellow
        $searchPaths | ForEach-Object { Write-Host "  $_" -ForegroundColor Gray }
        exit 1
    }
}

Write-Host "Using plugin-steps.json path: $PluginStepsJsonPath" -ForegroundColor Cyan

# =============================================================================
# SCRIPT START
# =============================================================================

Write-Host "==================================================" -ForegroundColor Cyan
Write-Host "  Remove Orphaned Plugin Steps and Types" -ForegroundColor Cyan
Write-Host "==================================================" -ForegroundColor Cyan
Write-Host ""

# =============================================================================
# IMPORT REQUIRED MODULES
# =============================================================================


# Check if module is installed, and install if not found
if (-not (Get-Module -ListAvailable -Name Microsoft.Xrm.Data.PowerShell)) {
    Write-Host "Microsoft.Xrm.Data.PowerShell module not found. Installing..." -ForegroundColor Yellow
    try {
        Install-Module Microsoft.Xrm.Data.PowerShell -Scope CurrentUser -Force -AllowClobber
        Write-Host "Microsoft.Xrm.Data.PowerShell module installed successfully." -ForegroundColor Green
    }
    catch {
        Write-Host "Failed to install Microsoft.Xrm.Data.PowerShell module: $_" -ForegroundColor Red
        Write-Host "You may need to run: Install-Module Microsoft.Xrm.Data.PowerShell -Scope CurrentUser" -ForegroundColor Yellow
        exit 1
    }
}

Import-Module Microsoft.Xrm.Data.PowerShell


# =============================================================================
# CONNECT TO DYNAMICS 365
# =============================================================================

try {
    Write-Host "`nConnecting to Dynamics 365 environment: $EnvironmentUrl" -ForegroundColor Cyan
    
    # Connect using service principal (no TenantId required)
    $connection = Connect-CrmOnline -ServerUrl $EnvironmentUrl `
        -ClientSecret $ClientSecret `
        -OAuthClientId $ClientId
    
    if ($connection.IsReady) {
        Write-Host "[OK] Successfully connected to Dynamics 365" -ForegroundColor Green
    }
    else {
        throw "Connection is not ready"
    }
}
catch {
    Write-Error "Failed to connect to Dynamics 365: $_"
    exit 1
}

# =============================================================================
# LOAD PLUGIN STEPS FROM JSON FILE
# =============================================================================

try {
    Write-Host "`nLoading plugin steps from JSON file: $PluginStepsJsonPath" -ForegroundColor Cyan
    
    if (-not (Test-Path $PluginStepsJsonPath)) {
        throw "Plugin steps JSON file not found at: $PluginStepsJsonPath"
    }
    
    $jsonContent = Get-Content $PluginStepsJsonPath -Raw | ConvertFrom-Json
    
    $assemblyName = $jsonContent.metadata.assemblyName
    $totalStepsInExport = $jsonContent.metadata.totalSteps
    
    Write-Host "[OK] Loaded plugin steps for assembly: $assemblyName" -ForegroundColor Green
    Write-Host "  Total steps in export: $totalStepsInExport" -ForegroundColor Gray
    
    # Build lookup sets based on logical identifiers (not GUIDs)
    $exportedStepKeys = @{}
    foreach ($step in $jsonContent.pluginSteps) {
        # Create a unique key for each step based on logical properties
        $stepKey = "$($step.pluginTypeName)|$($step.primaryEntity)|$($step.message)|$($step.stage)|$($step.rank)"
        $exportedStepKeys[$stepKey] = $true
    }
    
    Write-Host "  Extracted $($exportedStepKeys.Count) unique step identifiers from JSON" -ForegroundColor Gray
    
    # Extract unique plugin type names (not IDs)
    $exportedPluginTypeNames = @{}
    foreach ($step in $jsonContent.pluginSteps) {
        $exportedPluginTypeNames[$step.pluginTypeName] = $true
    }
    
    Write-Host "  Extracted $($exportedPluginTypeNames.Count) unique plugin type names from JSON" -ForegroundColor Gray
}
catch {
    Write-Error "Failed to load plugin steps from JSON: $_"
    if ($connection -and $connection.IsReady) {
        $connection.Dispose()
    }
    exit 1
}

# =============================================================================
# GET PLUGIN STEPS FROM TARGET ENVIRONMENT
# =============================================================================

try {
    Write-Host "`nRetrieving plugin steps from target environment for assembly: $assemblyName" -ForegroundColor Cyan
    
    # Query for plugin assembly
    $assemblyQuery = @"
<fetch>
  <entity name='pluginassembly'>
    <attribute name='pluginassemblyid' />
    <attribute name='name' />
    <filter>
      <condition attribute='name' operator='eq' value='$assemblyName' />
    </filter>
  </entity>
</fetch>
"@
    
    $assemblyResult = Get-CrmRecordsByFetch -conn $connection -Fetch $assemblyQuery
    
    if ($assemblyResult.CrmRecords.Count -eq 0) {
        Write-Warning "Assembly '$assemblyName' not found in target environment"
        Write-Host "`n[OK] Script completed - No assembly found to process" -ForegroundColor Green
        $connection.Dispose()
        exit 0
    }
    
    $assemblyId = $assemblyResult.CrmRecords[0].pluginassemblyid
    Write-Host "  Found assembly ID: $assemblyId" -ForegroundColor Gray
    
    # Query for plugin types in this assembly
    $pluginTypeQuery = @"
<fetch>
  <entity name='plugintype'>
    <attribute name='plugintypeid' />
    <attribute name='typename' />
    <filter>
      <condition attribute='pluginassemblyid' operator='eq' value='$assemblyId' />
    </filter>
  </entity>
</fetch>
"@
    
    $pluginTypes = Get-CrmRecordsByFetch -conn $connection -Fetch $pluginTypeQuery
    
    if ($pluginTypes.CrmRecords.Count -eq 0) {
        Write-Warning "No plugin types found for assembly '$assemblyName'"
        Write-Host "`n[OK] Script completed - No plugin types found" -ForegroundColor Green
        $connection.Dispose()
        exit 0
    }
    
    $pluginTypeIds = $pluginTypes.CrmRecords | ForEach-Object { $_.plugintypeid }
    Write-Host "  Found $($pluginTypeIds.Count) plugin type(s)" -ForegroundColor Gray
    
    # Build filter for plugin types
    $pluginTypeIdFilter = ($pluginTypeIds | ForEach-Object { "<value>$_</value>" }) -join ""
    
    # Query for all steps related to these plugin types - need to get more attributes for comparison
    $stepsQuery = @"
<fetch>
  <entity name='sdkmessageprocessingstep'>
    <attribute name='sdkmessageprocessingstepid' />
    <attribute name='name' />
    <attribute name='plugintypeid' />
    <attribute name='stage' />
    <attribute name='mode' />
    <attribute name='rank' />
    <link-entity name='plugintype' from='plugintypeid' to='plugintypeid' alias='pt'>
      <attribute name='typename' />
    </link-entity>
    <link-entity name='sdkmessagefilter' from='sdkmessagefilterid' to='sdkmessagefilterid' link-type='outer' alias='smf'>
      <attribute name='primaryobjecttypecode' />
    </link-entity>
    <link-entity name='sdkmessage' from='sdkmessageid' to='sdkmessageid' alias='sm'>
      <attribute name='name' />
    </link-entity>
    <filter>
      <condition attribute='plugintypeid' operator='in'>
        $pluginTypeIdFilter
      </condition>
    </filter>
  </entity>
</fetch>
"@
    
    $stepsResult = Get-CrmRecordsByFetch -conn $connection -Fetch $stepsQuery
    
    Write-Host "[OK] Found $($stepsResult.CrmRecords.Count) plugin step(s) in target environment" -ForegroundColor Green
    
    $targetSteps = $stepsResult.CrmRecords
}
catch {
    Write-Error "Failed to retrieve plugin steps from environment: $_"
    if ($connection -and $connection.IsReady) {
        $connection.Dispose()
    }
    exit 1
}

# =============================================================================
# ANALYZE AND REMOVE ORPHANED STEPS
# =============================================================================

Write-Host "`nAnalyzing plugin steps..." -ForegroundColor Cyan

# Helper function to convert stage value to text
function Get-StageText {
    param($stageValue)
    switch ($stageValue) {
        10 { return "Pre-validation" }
        20 { return "Pre-operation" }
        40 { return "Post-operation" }
        default { return "Unknown" }
    }
}

$orphanedSteps = @()

foreach ($step in $targetSteps) {
    # Build the same key format as the export
    $typename = $step."pt.typename"
    $primaryEntity = $step."smf.primaryobjecttypecode"
    if ([string]::IsNullOrEmpty($primaryEntity)) { $primaryEntity = "" }
    $message = $step."sm.name"
    $stageText = Get-StageText -stageValue $step.stage
    $rank = $step.rank
    
    $stepKey = "$typename|$primaryEntity|$message|$stageText|$rank"
    
    if (-not $exportedStepKeys.ContainsKey($stepKey)) {
        $orphanedSteps += $step
        Write-Verbose "Orphaned step found: $stepKey"
    }
}

if ($orphanedSteps.Count -eq 0) {
    Write-Host "[OK] No orphaned plugin steps found. All steps in target match the export." -ForegroundColor Green
}
else {
    # Orphaned steps found - proceed with removal
    Write-Host "`nFound $($orphanedSteps.Count) orphaned plugin step(s) to remove:" -ForegroundColor Yellow
}

$stepsRemoved = 0
$stepsFailed = 0

foreach ($step in $orphanedSteps) {
    Write-Host "  - $($step.name) (ID: $($step.sdkmessageprocessingstepid))" -ForegroundColor Yellow
    
    try {
        Remove-CrmRecord -conn $connection `
            -EntityLogicalName "sdkmessageprocessingstep" `
            -Id $step.sdkmessageprocessingstepid `
            -ErrorAction Stop
        
        Write-Host "    [OK] Removed successfully" -ForegroundColor Green
        $stepsRemoved++
    }
    catch {
        Write-Host "    [FAIL] Failed to remove: $_" -ForegroundColor Red
        $stepsFailed++
    }
}

# =============================================================================
# ANALYZE AND REMOVE ORPHANED PLUGIN TYPES
# =============================================================================

Write-Host "`nAnalyzing plugin types..." -ForegroundColor Cyan

$orphanedPluginTypes = @()

foreach ($pluginType in $pluginTypes.CrmRecords) {
    $typename = $pluginType.typename
    
    if (-not $exportedPluginTypeNames.ContainsKey($typename)) {
        $orphanedPluginTypes += $pluginType
        Write-Verbose "Orphaned type found: $typename"
    }
}

if ($orphanedPluginTypes.Count -eq 0) {
    Write-Host "[OK] No orphaned plugin types found. All types in target match the export." -ForegroundColor Green
}
else {
    # Orphaned types found - proceed with removal
    Write-Host "`nFound $($orphanedPluginTypes.Count) orphaned plugin type(s) to remove:" -ForegroundColor Yellow
}

$typesRemoved = 0
$typesFailed = 0

foreach ($pluginType in $orphanedPluginTypes) {
    Write-Host "  - $($pluginType.typename) (ID: $($pluginType.plugintypeid))" -ForegroundColor Yellow
    
    try {
        Remove-CrmRecord -conn $connection `
            -EntityLogicalName "plugintype" `
            -Id $pluginType.plugintypeid `
            -ErrorAction Stop
        
        Write-Host "    [OK] Removed successfully" -ForegroundColor Green
        $typesRemoved++
    }
    catch {
        Write-Host "    [FAIL] Failed to remove: $_" -ForegroundColor Red
        $typesFailed++
    }
}

# =============================================================================
# SUMMARY AND CLEANUP
# =============================================================================

Write-Host ""
Write-Host "==================================================" -ForegroundColor Cyan
Write-Host "  Summary" -ForegroundColor Cyan
Write-Host "==================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Plugin Steps:" -ForegroundColor Cyan
Write-Host "  Total orphaned steps found: $($orphanedSteps.Count)" -ForegroundColor $(if ($orphanedSteps.Count -gt 0) { "Yellow" } else { "Green" })
Write-Host "  Successfully removed: $stepsRemoved" -ForegroundColor Green
Write-Host "  Failed to remove: $stepsFailed" -ForegroundColor $(if ($stepsFailed -gt 0) { "Red" } else { "Green" })
Write-Host ""
Write-Host "Plugin Types:" -ForegroundColor Cyan
Write-Host "  Total orphaned types found: $($orphanedPluginTypes.Count)" -ForegroundColor $(if ($orphanedPluginTypes.Count -gt 0) { "Yellow" } else { "Green" })
Write-Host "  Successfully removed: $typesRemoved" -ForegroundColor Green
Write-Host "  Failed to remove: $typesFailed" -ForegroundColor $(if ($typesFailed -gt 0) { "Red" } else { "Green" })
Write-Host ""

$totalFailures = $stepsFailed + $typesFailed

if ($totalFailures -gt 0) {
    Write-Warning "Some plugin steps or types could not be removed. Check the logs above for details."
    
    if ($connection -and $connection.IsReady) {
        Write-Host "`nDisconnecting from Dynamics 365..." -ForegroundColor Gray
        $connection.Dispose()
    }
    exit 1
}

Write-Host "[OK] Script completed successfully" -ForegroundColor Green

if ($connection -and $connection.IsReady) {
    Write-Host "`nDisconnecting from Dynamics 365..." -ForegroundColor Gray
    $connection.Dispose()
}

exit 0
