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

.PARAMETER TenantId
    The Azure AD Tenant ID

.PARAMETER PluginStepsJsonPath
    Path to the plugin-steps.json file (defaults to looking in artifacts directory)

.EXAMPLE
    .\Remove-OrphanedPluginSteps.ps1 -EnvironmentUrl "https://org.crm.dynamics.com" `
        -ClientId "your-client-id" `
        -ClientSecret "your-client-secret" `
        -TenantId "your-tenant-id" `
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
    
    [Parameter(Mandatory=$true)]
    [string]$TenantId,
    
    [Parameter(Mandatory=$false)]
    [string]$PluginStepsJsonPath = "$(System.ArtifactsDirectory)/drop/plugin-steps.json"
)

# Import required modules
try {
    Import-Module Microsoft.Xrm.Data.PowerShell -ErrorAction Stop
    Write-Host "✓ Microsoft.Xrm.Data.PowerShell module loaded successfully" -ForegroundColor Green
}
catch {
    Write-Error "Failed to load Microsoft.Xrm.Data.PowerShell module. Please install it using: Install-Module Microsoft.Xrm.Data.PowerShell"
    exit 1
}

# Function to connect to Dynamics 365
function Connect-ToCRM {
    param(
        [string]$Url,
        [string]$AppId,
        [string]$Secret,
        [string]$Tenant
    )
    
    try {
        Write-Host "Connecting to Dynamics 365 environment: $Url" -ForegroundColor Cyan
        
        # Create secure string for client secret
        $SecureSecret = ConvertTo-SecureString $Secret -AsPlainText -Force
        
        # Connect using service principal
        $conn = Connect-CrmOnline -ServerUrl $Url `
            -ClientSecret $SecureSecret `
            -OAuthClientId $AppId `
            -TenantId $Tenant
        
        if ($conn.IsReady) {
            Write-Host "✓ Successfully connected to Dynamics 365" -ForegroundColor Green
            return $conn
        }
        else {
            throw "Connection is not ready"
        }
    }
    catch {
        Write-Error "Failed to connect to Dynamics 365: $_"
        exit 1
    }
}

# Function to load plugin steps from JSON
function Get-PluginStepsFromJson {
    param([string]$JsonPath)
    
    try {
        Write-Host "Loading plugin steps from JSON file: $JsonPath" -ForegroundColor Cyan
        
        if (-not (Test-Path $JsonPath)) {
            throw "Plugin steps JSON file not found at: $JsonPath"
        }
        
        $jsonContent = Get-Content $JsonPath -Raw | ConvertFrom-Json
        
        $assemblyName = $jsonContent.metadata.assemblyName
        Write-Host "✓ Loaded plugin steps for assembly: $assemblyName" -ForegroundColor Green
        Write-Host "  Total steps in export: $($jsonContent.metadata.totalSteps)" -ForegroundColor Gray
        
        # Extract step IDs
        $stepIds = $jsonContent.pluginSteps | ForEach-Object { $_.sdkmessageprocessingstepid }
        
        return @{
            AssemblyName = $assemblyName
            StepIds = $stepIds
            Steps = $jsonContent.pluginSteps
        }
    }
    catch {
        Write-Error "Failed to load plugin steps from JSON: $_"
        exit 1
    }
}

# Function to get plugin steps from target environment
function Get-PluginStepsFromEnvironment {
    param(
        [Microsoft.Xrm.Tooling.Connector.CrmServiceClient]$Connection,
        [string]$AssemblyName
    )
    
    try {
        Write-Host "Retrieving plugin steps from target environment for assembly: $AssemblyName" -ForegroundColor Cyan
        
        # Query for plugin assembly
        $assemblyQuery = @"
<fetch>
  <entity name='pluginassembly'>
    <attribute name='pluginassemblyid' />
    <attribute name='name' />
    <filter>
      <condition attribute='name' operator='eq' value='$AssemblyName' />
    </filter>
  </entity>
</fetch>
"@
        
        $assemblyResult = Get-CrmRecordsByFetch -conn $Connection -Fetch $assemblyQuery
        
        if ($assemblyResult.CrmRecords.Count -eq 0) {
            Write-Warning "Assembly '$AssemblyName' not found in target environment"
            return @()
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
        
        $pluginTypes = Get-CrmRecordsByFetch -conn $Connection -Fetch $pluginTypeQuery
        
        if ($pluginTypes.CrmRecords.Count -eq 0) {
            Write-Warning "No plugin types found for assembly '$AssemblyName'"
            return @()
        }
        
        $pluginTypeIds = $pluginTypes.CrmRecords | ForEach-Object { $_.plugintypeid }
        Write-Host "  Found $($pluginTypeIds.Count) plugin type(s)" -ForegroundColor Gray
        
        # Query for all steps related to these plugin types
        $pluginTypeIdFilter = ($pluginTypeIds | ForEach-Object { "<value>$_</value>" }) -join ""
        
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
        
        $stepsResult = Get-CrmRecordsByFetch -conn $Connection -Fetch $stepsQuery
        
        Write-Host "✓ Found $($stepsResult.CrmRecords.Count) plugin step(s) in target environment" -ForegroundColor Green
        
        return $stepsResult.CrmRecords
    }
    catch {
        Write-Error "Failed to retrieve plugin steps from environment: $_"
        throw
    }
}

# Function to remove orphaned plugin steps
function Remove-OrphanedSteps {
    param(
        [Microsoft.Xrm.Tooling.Connector.CrmServiceClient]$Connection,
        [array]$TargetSteps,
        [array]$ExportedStepIds
    )
    
    Write-Host "`nAnalyzing plugin steps..." -ForegroundColor Cyan
    
    $orphanedSteps = @()
    
    foreach ($step in $TargetSteps) {
        $stepId = $step.sdkmessageprocessingstepid
        
        if ($stepId -notin $ExportedStepIds) {
            $orphanedSteps += $step
        }
    }
    
    if ($orphanedSteps.Count -eq 0) {
        Write-Host "✓ No orphaned plugin steps found. All steps in target match the export." -ForegroundColor Green
        return @{
            Total = 0
            Removed = 0
            Failed = 0
        }
    }
    
    Write-Host "`nFound $($orphanedSteps.Count) orphaned plugin step(s) to remove:" -ForegroundColor Yellow
    
    $removed = 0
    $failed = 0
    
    foreach ($step in $orphanedSteps) {
        Write-Host "  - $($step.name) (ID: $($step.sdkmessageprocessingstepid))" -ForegroundColor Yellow
        
        try {
            Remove-CrmRecord -conn $Connection `
                -EntityLogicalName "sdkmessageprocessingstep" `
                -Id $step.sdkmessageprocessingstepid `
                -ErrorAction Stop
            
            Write-Host "    ✓ Removed successfully" -ForegroundColor Green
            $removed++
        }
        catch {
            Write-Host "    ✗ Failed to remove: $_" -ForegroundColor Red
            $failed++
        }
    }
    
    return @{
        Total = $orphanedSteps.Count
        Removed = $removed
        Failed = $failed
    }
}

# Main script execution
try {
    Write-Host "==================================================" -ForegroundColor Cyan
    Write-Host "  Remove Orphaned Plugin Steps" -ForegroundColor Cyan
    Write-Host "==================================================" -ForegroundColor Cyan
    Write-Host ""
    
    # Connect to Dynamics 365
    $connection = Connect-ToCRM -Url $EnvironmentUrl `
        -AppId $ClientId `
        -Secret $ClientSecret `
        -Tenant $TenantId
    
    Write-Host ""
    
    # Load exported plugin steps
    $exportData = Get-PluginStepsFromJson -JsonPath $PluginStepsJsonPath
    
    Write-Host ""
    
    # Get current plugin steps from target environment
    $targetSteps = Get-PluginStepsFromEnvironment -Connection $connection `
        -AssemblyName $exportData.AssemblyName
    
    Write-Host ""
    
    # Remove orphaned steps
    $result = Remove-OrphanedSteps -Connection $connection `
        -TargetSteps $targetSteps `
        -ExportedStepIds $exportData.StepIds
    
    Write-Host ""
    Write-Host "==================================================" -ForegroundColor Cyan
    Write-Host "  Summary" -ForegroundColor Cyan
    Write-Host "==================================================" -ForegroundColor Cyan
    Write-Host "Total orphaned steps found: $($result.Total)" -ForegroundColor $(if ($result.Total -gt 0) { "Yellow" } else { "Green" })
    Write-Host "Successfully removed: $($result.Removed)" -ForegroundColor Green
    Write-Host "Failed to remove: $($result.Failed)" -ForegroundColor $(if ($result.Failed -gt 0) { "Red" } else { "Green" })
    Write-Host ""
    
    if ($result.Failed -gt 0) {
        Write-Warning "Some plugin steps could not be removed. Check the logs above for details."
        exit 1
    }
    
    Write-Host "✓ Script completed successfully" -ForegroundColor Green
}
catch {
    Write-Error "Script failed with error: $_"
    Write-Error $_.ScriptStackTrace
    exit 1
}
finally {
    if ($connection -and $connection.IsReady) {
        Write-Host "`nDisconnecting from Dynamics 365..." -ForegroundColor Gray
        $connection.Dispose()
    }
}
