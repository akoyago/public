<#
.SYNOPSIS
    Removes orphaned plugin steps and plugin types from Dynamics 365 environment.

.DESCRIPTION
    This script compares plugin steps in the exported plugin-steps.json file with the target
    Dynamics 365 environment and removes any steps that exist in the target but not in the export.
    Also removes orphaned plugin types that exist in the target but not in the export.
    Only processes steps for the AkoyaGo.Plugins assembly.
    Steps are compared by GUID, plugin types are compared by typename.
    
    Additionally, if the assembly version matches AND there are missing plugin types in the target,
    the script will re-register the assembly from solution-managed.zip.

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
    $exportedAssemblyVersion = $jsonContent.metadata.assemblyVersion
    
    Write-Host "[OK] Loaded plugin steps for assembly: $assemblyName" -ForegroundColor Green
    Write-Host "  Total steps in export: $totalStepsInExport" -ForegroundColor Gray
    Write-Host "  Assembly version in export: $exportedAssemblyVersion" -ForegroundColor Gray
    
    # Extract step IDs from JSON and convert to lowercase strings for comparison
    $exportedStepIds = $jsonContent.pluginSteps | ForEach-Object { $_.sdkmessageprocessingstepid.ToString().ToLower() }
    
    Write-Host "  Extracted $($exportedStepIds.Count) step IDs from JSON" -ForegroundColor Gray
    
    # Extract unique plugin type names (not IDs) for comparison by logical name
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
    <attribute name='version' />
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
    $targetAssemblyVersion = $assemblyResult.CrmRecords[0].version
    Write-Host "  Found assembly ID: $assemblyId" -ForegroundColor Gray
    Write-Host "  Assembly version in target: $targetAssemblyVersion" -ForegroundColor Gray
    
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
# CHECK FOR MISSING PLUGIN TYPES AND RE-REGISTER IF NEEDED
# =============================================================================

Write-Host "`nChecking for missing plugin types..." -ForegroundColor Cyan

# Build hashtable of existing plugin types in target
$targetPluginTypeNames = @{}
foreach ($pluginType in $pluginTypes.CrmRecords) {
    $targetPluginTypeNames[$pluginType.typename] = $true
}

# Find missing plugin types
$missingPluginTypes = @()
foreach ($typeName in $exportedPluginTypeNames.Keys) {
    if (-not $targetPluginTypeNames.ContainsKey($typeName)) {
        $missingPluginTypes += $typeName
    }
}

if ($missingPluginTypes.Count -gt 0) {
    Write-Host "  Found $($missingPluginTypes.Count) missing plugin type(s) in target:" -ForegroundColor Yellow
    foreach ($missingType in $missingPluginTypes) {
        Write-Host "    - $missingType" -ForegroundColor Yellow
    }
    
    # Check if versions match
    if ($targetAssemblyVersion -eq $exportedAssemblyVersion) {
        Write-Host "`n  Version match detected: Target=$targetAssemblyVersion, Export=$exportedAssemblyVersion" -ForegroundColor Yellow
        Write-Host "  Initiating assembly re-registration..." -ForegroundColor Yellow
        
        # Find solution-managed.zip
        $solutionZipPath = $null
        $searchPaths = @(
            "$env:SYSTEM_ARTIFACTSDIRECTORY\drop\solution-managed.zip",
            "$env:SYSTEM_ARTIFACTSDIRECTORY\solution-managed.zip",
            "$env:SYSTEM_DEFAULTWORKINGDIRECTORY\drop\solution-managed.zip",
            "$env:SYSTEM_DEFAULTWORKINGDIRECTORY\solution-managed.zip",
            "$env:AGENT_RELEASEDIRECTORY\drop\solution-managed.zip",
            "$env:BUILD_ARTIFACTSTAGINGDIRECTORY\solution-managed.zip"
        )
        
        foreach ($path in $searchPaths) {
            if ($path -and (Test-Path $path)) {
                $solutionZipPath = $path
                Write-Host "  Found solution-managed.zip at: $solutionZipPath" -ForegroundColor Green
                break
            }
        }
        
        if (-not $solutionZipPath) {
            # Search recursively
            $searchRoots = @($env:SYSTEM_ARTIFACTSDIRECTORY, $env:SYSTEM_DEFAULTWORKINGDIRECTORY, "D:\a")
            foreach ($root in $searchRoots) {
                if ($root -and (Test-Path $root)) {
                    $found = Get-ChildItem -Path $root -Recurse -Filter "solution-managed.zip" -ErrorAction SilentlyContinue | Select-Object -First 1
                    if ($found) {
                        $solutionZipPath = $found.FullName
                        Write-Host "  Found solution-managed.zip at: $solutionZipPath" -ForegroundColor Green
                        break
                    }
                }
            }
        }
        
        if ($solutionZipPath) {
            try {
                # Extract and search for AkoyaGo.Plugins.dll from solution-managed.zip
                Write-Host "  Extracting solution archive to locate plugin assembly..." -ForegroundColor Cyan
                
                Add-Type -AssemblyName System.IO.Compression.FileSystem
                $tempExtractPath = Join-Path $env:TEMP "PluginExtract_$(Get-Date -Format 'yyyyMMddHHmmss')"
                New-Item -ItemType Directory -Path $tempExtractPath -Force | Out-Null
                
                [System.IO.Compression.ZipFile]::ExtractToDirectory($solutionZipPath, $tempExtractPath)
                
                # Search for the DLL in multiple possible locations within the solution
                $possiblePaths = @(
                    "PluginAssemblies\AkoyaGo.Plugins.dll",
                    "PluginAssemblies\AkoyaGo.Plugins\AkoyaGo.Plugins.dll",
                    "Plugins\AkoyaGo.Plugins.dll"
                )
                
                $pluginDllPath = $null
                
                foreach ($relativePath in $possiblePaths) {
                    $testPath = Join-Path $tempExtractPath $relativePath
                    if (Test-Path $testPath) {
                        $pluginDllPath = $testPath
                        Write-Host "  Found DLL at: $relativePath" -ForegroundColor Gray
                        break
                    }
                }
                
                # If not found in standard locations, search recursively
                if (-not $pluginDllPath) {
                    Write-Host "  Searching recursively for AkoyaGo.Plugins.dll..." -ForegroundColor Gray
                    $foundDll = Get-ChildItem -Path $tempExtractPath -Recurse -Filter "AkoyaGo.Plugins.dll" -ErrorAction SilentlyContinue | Select-Object -First 1
                    if ($foundDll) {
                        $pluginDllPath = $foundDll.FullName
                        $relativePath = $pluginDllPath.Substring($tempExtractPath.Length + 1)
                        Write-Host "  Found DLL at: $relativePath" -ForegroundColor Gray
                    }
                }
                
                if (-not $pluginDllPath) {
                    Write-Warning "  AkoyaGo.Plugins.dll not found in solution-managed.zip"
                    Write-Host "  Searched locations:" -ForegroundColor Gray
                    foreach ($path in $possiblePaths) {
                        Write-Host "    - $path" -ForegroundColor Gray
                    }
                    
                    # List contents of PluginAssemblies folder if it exists
                    $pluginAssembliesPath = Join-Path $tempExtractPath "PluginAssemblies"
                    if (Test-Path $pluginAssembliesPath) {
                        Write-Host "  Contents of PluginAssemblies folder:" -ForegroundColor Gray
                        Get-ChildItem -Path $pluginAssembliesPath -Recurse | ForEach-Object {
                            $relativePath = $_.FullName.Substring($tempExtractPath.Length + 1)
                            Write-Host "    - $relativePath" -ForegroundColor Gray
                        }
                    }
                }
                else {
                    Write-Host "  [OK] Located AkoyaGo.Plugins.dll" -ForegroundColor Green
                    
                    # Read DLL as base64
                    $dllBytes = [System.IO.File]::ReadAllBytes($pluginDllPath)
                    $dllBase64 = [System.Convert]::ToBase64String($dllBytes)
                    
                    Write-Host "  Updating plugin assembly in target environment..." -ForegroundColor Cyan
                    
                    # Update the plugin assembly record with new content
                    $updateFields = @{
                        "content" = $dllBase64
                    }
                    
                    Set-CrmRecord -conn $connection `
                        -EntityLogicalName "pluginassembly" `
                        -Id $assemblyId `
                        -Fields $updateFields
                    
                    Write-Host "  [OK] Plugin assembly re-registered successfully" -ForegroundColor Green
                    Write-Host "  Missing plugin types should now be available in the target environment" -ForegroundColor Green
                }
                
                # Cleanup temp directory
                Remove-Item -Path $tempExtractPath -Recurse -Force -ErrorAction SilentlyContinue
            }
            catch {
                Write-Warning "  Failed to re-register plugin assembly: $_"
                Write-Host "  Continuing with orphaned step/type removal..." -ForegroundColor Yellow
            }
        }
        else {
            Write-Warning "  solution-managed.zip not found in artifacts - cannot re-register assembly"
            Write-Host "  Continuing with orphaned step/type removal..." -ForegroundColor Yellow
        }
    }
    else {
        Write-Host "  Version mismatch detected: Target=$targetAssemblyVersion, Export=$exportedAssemblyVersion" -ForegroundColor Gray
        Write-Host "  Skipping re-registration (versions must match)" -ForegroundColor Gray
    }
}
else {
    Write-Host "  [OK] All plugin types from export exist in target environment" -ForegroundColor Green
}

# =============================================================================
# ANALYZE AND REMOVE ORPHANED STEPS
# =============================================================================

Write-Host "`nAnalyzing plugin steps..." -ForegroundColor Cyan

$orphanedSteps = @()

foreach ($step in $targetSteps) {
    # Convert GUID to lowercase string for comparison
    $stepId = $step.sdkmessageprocessingstepid.ToString().ToLower()
    
    if ($stepId -notin $exportedStepIds) {
        $orphanedSteps += $step
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
Write-Host "Assembly Re-registration:" -ForegroundColor Cyan
Write-Host "  Missing plugin types detected: $($missingPluginTypes.Count)" -ForegroundColor $(if ($missingPluginTypes.Count -gt 0) { "Yellow" } else { "Green" })
if ($missingPluginTypes.Count -gt 0 -and $targetAssemblyVersion -eq $exportedAssemblyVersion) {
    Write-Host "  Re-registration attempted: Yes" -ForegroundColor Yellow
}
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
