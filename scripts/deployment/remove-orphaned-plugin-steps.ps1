<#
.SYNOPSIS
    Removes orphaned plugin steps from Dynamics 365 environment.

.DESCRIPTION
    This script compares plugin steps in the exported plugin-steps.json file with the target
    Dynamics 365 environment and removes any steps that exist in the target but not in the export.
    Only processes steps for the AkoyaGo.Plugins assembly.

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
# HANDLE DEFAULT PATH
# =============================================================================

if ([string]::IsNullOrEmpty($PluginStepsJsonPath)) {
    # Use Azure DevOps default path if not specified
    if ($env:SYSTEM_ARTIFACTSDIRECTORY) {
        $PluginStepsJsonPath = Join-Path $env:SYSTEM_ARTIFACTSDIRECTORY "drop\plugin-steps.json"
        Write-Host "Using default artifact path: $PluginStepsJsonPath" -ForegroundColor Gray
    }
    else {
        Write-Error "PluginStepsJsonPath parameter is required when not running in Azure DevOps pipeline"
        exit 1
    }
}

# =============================================================================
# SCRIPT START
# =============================================================================

Write-Host "==================================================" -ForegroundColor Cyan
Write-Host "  Remove Orphaned Plugin Steps" -ForegroundColor Cyan
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
    
    # Extract step IDs from JSON
    $exportedStepIds = $jsonContent.pluginSteps | ForEach-Object { $_.sdkmessageprocessingstepid }
    
    Write-Host "  Extracted $($exportedStepIds.Count) step IDs from JSON" -ForegroundColor Gray
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
    
    # Query for all steps related to these plugin types
    $stepsQuery = @"
<fetch>
  <entity name='sdkmessageprocessingstep'>
    <attribute name='sdkmessageprocessingstepid' />
    <attribute name='name' />
    <attribute name='plugintypeid' />
    <attribute name='stage' />
    <attribute name='mode' />
    <attribute name='rank' />
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

$orphanedSteps = @()

foreach ($step in $targetSteps) {
    $stepId = $step.sdkmessageprocessingstepid
    
    if ($stepId -notin $exportedStepIds) {
        $orphanedSteps += $step
    }
}

if ($orphanedSteps.Count -eq 0) {
    Write-Host "[OK] No orphaned plugin steps found. All steps in target match the export." -ForegroundColor Green
    
    Write-Host "`n==================================================" -ForegroundColor Cyan
    Write-Host "  Summary" -ForegroundColor Cyan
    Write-Host "==================================================" -ForegroundColor Cyan
    Write-Host "Total orphaned steps found: 0" -ForegroundColor Green
    Write-Host "Successfully removed: 0" -ForegroundColor Green
    Write-Host "Failed to remove: 0" -ForegroundColor Green
    Write-Host ""
    Write-Host "[OK] Script completed successfully" -ForegroundColor Green
    
    if ($connection -and $connection.IsReady) {
        Write-Host "`nDisconnecting from Dynamics 365..." -ForegroundColor Gray
        $connection.Dispose()
    }
    exit 0
}

# Orphaned steps found - proceed with removal
Write-Host "`nFound $($orphanedSteps.Count) orphaned plugin step(s) to remove:" -ForegroundColor Yellow

$removed = 0
$failed = 0

foreach ($step in $orphanedSteps) {
    Write-Host "  - $($step.name) (ID: $($step.sdkmessageprocessingstepid))" -ForegroundColor Yellow
    
    try {
        Remove-CrmRecord -conn $connection `
            -EntityLogicalName "sdkmessageprocessingstep" `
            -Id $step.sdkmessageprocessingstepid `
            -ErrorAction Stop
        
        Write-Host "    [OK] Removed successfully" -ForegroundColor Green
        $removed++
    }
    catch {
        Write-Host "    [FAIL] Failed to remove: $_" -ForegroundColor Red
        $failed++
    }
}

# =============================================================================
# SUMMARY AND CLEANUP
# =============================================================================

Write-Host ""
Write-Host "==================================================" -ForegroundColor Cyan
Write-Host "  Summary" -ForegroundColor Cyan
Write-Host "==================================================" -ForegroundColor Cyan
Write-Host "Total orphaned steps found: $($orphanedSteps.Count)" -ForegroundColor $(if ($orphanedSteps.Count -gt 0) { "Yellow" } else { "Green" })
Write-Host "Successfully removed: $removed" -ForegroundColor Green
Write-Host "Failed to remove: $failed" -ForegroundColor $(if ($failed -gt 0) { "Red" } else { "Green" })
Write-Host ""

if ($failed -gt 0) {
    Write-Warning "Some plugin steps could not be removed. Check the logs above for details."
    
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
