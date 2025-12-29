<#
.SYNOPSIS
    Generates a JSON file containing all registered plugin steps for the akoyaGO Plugins assembly.

.DESCRIPTION
    This script queries a Dynamics 365 environment and exports all plugin steps, including
    their configuration, filtering attributes, and images, to a JSON file for documentation
    and deployment verification purposes.

.PARAMETER EnvironmentUrl
    The URL of your Dynamics 365 environment (e.g., https://yourorg.crm.dynamics.com)

.PARAMETER OutputPath
    The path where the JSON file will be saved. Defaults to ../deployment/plugin-steps.json

.PARAMETER ClientId
    The Client ID (Application ID) of the Azure AD app registration for service principal authentication

.PARAMETER ClientSecret
    The Client Secret of the Azure AD app registration for service principal authentication

.PARAMETER TenantId
    The Tenant ID (Directory ID) of your Azure AD tenant

.EXAMPLE
    .\generate-plugin-steps.ps1 -EnvironmentUrl "https://yourorg.crm.dynamics.com"

.EXAMPLE
    .\generate-plugin-steps.ps1 -EnvironmentUrl "https://yourorg.crm.dynamics.com" -ClientId "12345678-1234-1234-1234-123456789012" -ClientSecret "your-secret" -TenantId "87654321-4321-4321-4321-210987654321" -OutputPath "$(Build.ArtifactStagingDirectory)/plugin-steps.json"

.NOTES
    Requires: Microsoft.Xrm.Data.PowerShell module
    Install with: Install-Module Microsoft.Xrm.Data.PowerShell -Scope CurrentUser
    
    For pipeline use, provide ClientId, ClientSecret, and TenantId for non-interactive authentication.
#>

param(
    [Parameter(Mandatory=$false)]
    [string]$EnvironmentUrl,
    
    [Parameter(Mandatory=$false)]
    [string]$OutputPath = "$PSScriptRoot\..\..\deployment\plugin-steps.json",
    
    [Parameter(Mandatory=$false)]
    [string]$ClientId,
    
    [Parameter(Mandatory=$false)]
    [string]$ClientSecret,
    
    [Parameter(Mandatory=$false)]
    [string]$TenantId
)

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

# Connect to Dynamics 365
Write-Host "Connecting to Dynamics 365..." -ForegroundColor Cyan

if ($ClientId -and $ClientSecret -and $TenantId) {
    # Service Principal authentication for pipeline
    Write-Host "Using Service Principal authentication" -ForegroundColor Yellow
    if ($EnvironmentUrl) {
        Write-Host "Environment: $EnvironmentUrl" -ForegroundColor Cyan
    }
    Write-Host "Tenant ID: $TenantId" -ForegroundColor Cyan
    
    try {
        # Build connection string for OAuth with Client Secret
        if (-not $EnvironmentUrl) {
            Write-Host "EnvironmentUrl is required when using Service Principal authentication" -ForegroundColor Red
            exit 1
        }

        $connectionString = @"
AuthType=ClientSecret;
Url=$EnvironmentUrl;
ClientId=$ClientId;
ClientSecret=$ClientSecret;
"@
        
        $conn = Get-CrmConnection -ConnectionString $connectionString
    }
    catch {
        Write-Host "Failed to connect using Service Principal: $_" -ForegroundColor Red
        Write-Host "Error Details: $($_.Exception.Message)" -ForegroundColor Red
        exit 1
    }
} else {
    # Interactive mode for local testing
    Write-Host "Using Interactive authentication" -ForegroundColor Yellow
    try {
        if ($EnvironmentUrl) {
            $conn = Get-CrmConnection -ServerUrl $EnvironmentUrl -InteractiveMode
        } else {
            $conn = Get-CrmConnection -InteractiveMode
        }
    }
    catch {
        Write-Host "Failed to connect interactively: $_" -ForegroundColor Red
        Write-Host "Error Details: $($_.Exception.Message)" -ForegroundColor Red
        exit 1
    }
}

if (-not $conn.IsReady) {
    Write-Host "Failed to connect to Dynamics 365" -ForegroundColor Red
    exit 1
}

Write-Host "Connected to: $($conn.ConnectedOrgFriendlyName)" -ForegroundColor Green

# Query plugin assembly to get current version
$assemblyFetch = @"
<fetch>
  <entity name='pluginassembly'>
    <attribute name='name'/>
    <attribute name='version'/>
    <attribute name='pluginassemblyid'/>
    <filter>
      <condition attribute='name' operator='eq' value='AkoyaGo.Plugins'/>
    </filter>
  </entity>
</fetch>
"@

Write-Host "`nQuerying akoyaGO Plugins assembly..." -ForegroundColor Cyan
$assemblyResult = Get-CrmRecordsByFetch -conn $conn -Fetch $assemblyFetch

if ($assemblyResult.CrmRecords.Count -eq 0) {
    Write-Host "Assembly 'akoyaGO Plugins' not found in this environment" -ForegroundColor Red
    exit 1
}

$assembly = $assemblyResult.CrmRecords[0]
Write-Host "Found assembly version: $($assembly.version)" -ForegroundColor Green

# Query all plugin steps for the assembly
$fetchXml = @"
<fetch>
  <entity name='sdkmessageprocessingstep'>
    <attribute name='sdkmessageprocessingstepid'/>
    <attribute name='plugintypeid'/>
    <attribute name='name'/>
    <attribute name='stage'/>
    <attribute name='rank'/>
    <attribute name='mode'/>
    <attribute name='filteringattributes'/>
    <attribute name='configuration'/>
    <attribute name='description'/>
    <attribute name='asyncautodelete'/>
    <attribute name='statecode'/>
    <attribute name='impersonatinguserid'/>
    <link-entity name='sdkmessagefilter' from='sdkmessagefilterid' to='sdkmessagefilterid' alias='filter' link-type='outer'>
      <attribute name='primaryobjecttypecode'/>
    </link-entity>
    <link-entity name='sdkmessage' from='sdkmessageid' to='sdkmessageid' alias='message'>
      <attribute name='name'/>
    </link-entity>
    <link-entity name='plugintype' from='plugintypeid' to='plugintypeid' alias='type'>
      <attribute name='typename'/>
      <attribute name='friendlyname'/>
      <link-entity name='pluginassembly' from='pluginassemblyid' to='pluginassemblyid' alias='assembly'>
        <attribute name='name'/>
        <attribute name='version'/>
        <filter>
          <condition attribute='name' operator='eq' value='AkoyaGo.Plugins'/>
        </filter>
      </link-entity>
    </link-entity>
    <link-entity name='systemuser' from='systemuserid' to='impersonatinguserid' alias='impuser' link-type='outer'>
      <attribute name='fullname'/>
      <attribute name='domainname'/>
      <attribute name='systemuserid'/>
      <attribute name='applicationid'/>
    </link-entity>
    <order attribute='name' ascending='true'/>
  </entity>
</fetch>
"@

Write-Host "`nQuerying plugin steps..." -ForegroundColor Cyan
$steps = Get-CrmRecordsByFetch -conn $conn -Fetch $fetchXml

Write-Host "Found $($steps.CrmRecords.Count) plugin steps" -ForegroundColor Green

# Build the JSON structure
$pluginStepsData = [ordered]@{
    metadata = [ordered]@{
        solution = "akoyaGO Solution"
        assemblyName = $assembly.name
        assemblyVersion = $assembly.version
        lastUpdated = (Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
        environment = $conn.ConnectedOrgFriendlyName
        generatedBy = $env:USERNAME
        totalSteps = $steps.CrmRecords.Count
    }
    pluginSteps = @()
}

$stepCount = 0
foreach ($step in $steps.CrmRecords) {
    $stepCount++
    Write-Host "Processing step $stepCount of $($steps.CrmRecords.Count): $($step.name)" -ForegroundColor Yellow

    # Escape the GUID for XML safety
    $escapedStepId = [System.Security.SecurityElement]::Escape($step.sdkmessageprocessingstepid)
    
    # Query images for this step
    $imagesFetch = @"
<fetch>
  <entity name='sdkmessageprocessingstepimage'>
    <attribute name='sdkmessageprocessingstepimageid'/>
    <attribute name='name'/>
    <attribute name='entityalias'/>
    <attribute name='imagetype'/>
    <attribute name='attributes'/>
    <attribute name='messagepropertyname'/>
    <filter>
      <condition attribute='sdkmessageprocessingstepid' operator='eq' value='$escapedStepId'/>
    </filter>
    <order attribute='imagetype' ascending='true'/>
  </entity>
</fetch>
"@
    
    $images = Get-CrmRecordsByFetch -conn $conn -Fetch $imagesFetch
    
    $imageArray = @()
    foreach ($image in $images.CrmRecords) {
        $imageType = switch ($image.imagetype) {
            0 { "PreImage" }
            1 { "PostImage" }
            2 { "Both" }
            default { "Unknown" }
        }
        
        $imageArray += [ordered]@{
            name = $image.name
            entityAlias = $image.entityalias
            imageType = $imageType
            messagePropertyName = if ($image.messagepropertyname) { $image.messagepropertyname } else { "Id" }
            attributes = if ($image.attributes) { $image.attributes -split ',' | Sort-Object } else { @() }
            sdkmessageprocessingstepimageid = $image.sdkmessageprocessingstepimageid
        }
    }
    
    # Map stage - handle both numeric values and string labels
    $stageName = if ($step.stage -is [int]) {
        # Numeric value
        switch ($step.stage) {
            10 { "Pre-validation" }
            20 { "Pre-operation" }
            40 { "Post-operation" }
            default { "Unknown" }
        }
    } else {
        # String label - normalize to consistent format
        switch ($step.stage) {
            "Pre-validation" { "Pre-validation" }
            "Pre-operation" { "Pre-operation" }
            "Post-operation" { "Post-operation" }
            "MainOperation" { "MainOperation" }
            default { $step.stage }  # Keep original value if not recognized
        }
    }
    
    # Map mode - handle both numeric values and string labels
    $modeName = if ($step.mode -is [int]) {
        # Numeric value
        switch ($step.mode) {
            0 { "Synchronous" }
            1 { "Asynchronous" }
            default { "Unknown" }
        }
    } else {
        # String label - normalize to consistent format
        switch ($step.mode) {
            "Synchronous" { "Synchronous" }
            "Asynchronous" { "Asynchronous" }
            "Sync" { "Synchronous" }
            "Async" { "Asynchronous" }
            default { $step.mode }  # Keep original value if not recognized
        }
    }
    
    # Map state - handle both numeric values and string labels
    $state = if ($step.statecode -is [int]) {
        # Numeric value
        switch ($step.statecode) {
            0 { "Enabled" }
            1 { "Disabled" }
            default { "Unknown" }
        }
    } else {
        # String label - normalize to consistent format
        switch ($step.statecode) {
            "Enabled" { "Enabled" }
            "Disabled" { "Disabled" }
            "Active" { "Enabled" }
            "Inactive" { "Disabled" }
            default { $step.statecode }  # Keep original value if not recognized
        }
    }
    
    # Determine the run-as user (impersonating user)
    $runAsUser = if ($step.impersonatinguserid) {
        [ordered]@{
            systemUserId = $step.'impuser.systemuserid'
            fullName = $step.'impuser.fullname'
            domainName = $step.'impuser.domainname'
            applicationId = if ($step.'impuser.applicationid') { $step.'impuser.applicationid' } else { $null }
        }
    } else {
        [ordered]@{
            systemUserId = $null
            fullName = "Calling User"
            domainName = $null
            applicationId = $null
        }
    }
    
    Write-Host "  Stage: $($step.stage) -> $stageName | Mode: $($step.mode) -> $modeName | State: $($step.statecode) -> $state | Run As: $($runAsUser.fullName)" -ForegroundColor DarkGray
    
    $stepData = [ordered]@{
        name = $step.name
        pluginTypeName = $step.'type.typename'
        plugintypeid = $step.plugintypeid
        primaryEntity = if ($step.'filter.primaryobjecttypecode' -and $step.'filter.primaryobjecttypecode' -ne "none") { $step.'filter.primaryobjecttypecode' } else { "" }
        message = $step.'message.name'
        stage = $stageName
        mode = $modeName
        rank = $step.rank
        state = $state
        asyncAutoDelete = if ($step.asyncautodelete) { $step.asyncautodelete } else { $false }
        runAsUser = $runAsUser
        filteringAttributes = if ($step.filteringattributes) { $step.filteringattributes -split ',' | Sort-Object } else { @() }
        description = if ($step.description) { $step.description } else { "" }
        configuration = if ($step.configuration) { $step.configuration } else { "" }
        secureConfiguration = ""  # Can't retrieve secure config via API
        images = $imageArray
        sdkmessageprocessingstepid = $step.sdkmessageprocessingstepid
    }
    
    $pluginStepsData.pluginSteps += $stepData
}

# Ensure output directory exists
$outputDir = Split-Path -Parent $OutputPath
if (-not (Test-Path $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
    Write-Host "`nCreated directory: $outputDir" -ForegroundColor Green
}

# Export to JSON with proper formatting
Write-Host "`nExporting to JSON..." -ForegroundColor Cyan
$json = $pluginStepsData | ConvertTo-Json -Depth 10
$json | Out-File -FilePath $OutputPath -Encoding UTF8 -Force

Write-Host "`n========================================" -ForegroundColor Green
Write-Host "SUCCESS: Plugin steps exported" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host "File: $OutputPath" -ForegroundColor Cyan
Write-Host "Assembly: $($assembly.name) v$($assembly.version)" -ForegroundColor Cyan
Write-Host "Total steps: $($pluginStepsData.pluginSteps.Count)" -ForegroundColor Cyan
Write-Host "Environment: $($conn.ConnectedOrgFriendlyName)" -ForegroundColor Cyan
Write-Host "========================================`n" -ForegroundColor Green

# Display summary by entity
$stepsByEntity = $pluginStepsData.pluginSteps | Group-Object -Property primaryEntity
Write-Host "Steps by Entity:" -ForegroundColor Yellow
foreach ($group in $stepsByEntity | Sort-Object Name) {
    Write-Host "  $($group.Name): $($group.Count) steps" -ForegroundColor White

}


