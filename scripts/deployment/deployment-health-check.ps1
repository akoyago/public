<#
.SYNOPSIS
    Health check script to validate Dynamics 365 deployment.

.DESCRIPTION
    This script validates that web resources and plugin steps in the target Dynamics 365 environment
    match the exported solution and plugin-steps.json file from the build artifacts.
    Attempts to fix mismatches and reports any unresolvable issues.
    Web resource validation is limited to HTML and JavaScript files only.

.PARAMETER EnvironmentUrl
    The URL of the target Dynamics 365 environment (e.g., https://org.crm.dynamics.com)

.PARAMETER ClientId
    The Azure AD Application (Client) ID for service principal authentication

.PARAMETER ClientSecret
    The Client Secret for the service principal

.PARAMETER ArtifactDirectory
    Path to the directory containing solution-managed.zip and plugin-steps.json

.EXAMPLE
    .\HealthCheck-Deployment.ps1 -EnvironmentUrl "https://org.crm.dynamics.com" `
        -ClientId "your-client-id" `
        -ClientSecret "your-client-secret"
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
    [string]$ArtifactDirectory
)

# =============================================================================
# INITIALIZE VARIABLES
# =============================================================================

$script:failures = @()
$script:warnings = @()
$script:successes = @()
$script:fixes = @()

# =============================================================================
# HELPER FUNCTIONS
# =============================================================================

function Write-StatusMessage {
    param(
        [string]$Message,
        [ValidateSet('Info', 'Success', 'Warning', 'Error')]
        [string]$Type = 'Info'
    )
    
    $color = switch ($Type) {
        'Info' { 'Cyan' }
        'Success' { 'Green' }
        'Warning' { 'Yellow' }
        'Error' { 'Red' }
    }
    
    $prefix = switch ($Type) {
        'Info' { '[INFO]' }
        'Success' { '[OK]' }
        'Warning' { '[WARN]' }
        'Error' { '[FAIL]' }
    }
    
    Write-Host "$prefix $Message" -ForegroundColor $color
}

function Add-Failure {
    param([string]$Message)
    $script:failures += $Message
    Write-StatusMessage $Message -Type Error
}

function Add-Warning {
    param([string]$Message)
    $script:warnings += $Message
    Write-StatusMessage $Message -Type Warning
}

function Add-Success {
    param([string]$Message)
    $script:successes += $Message
}

function Add-Fix {
    param([string]$Message)
    $script:fixes += $Message
    Write-StatusMessage $Message -Type Success
}

function Find-ArtifactDirectory {
    if ([string]::IsNullOrEmpty($ArtifactDirectory)) {
        Write-StatusMessage "No artifact directory provided, searching for artifacts..." -Type Info
        
        $searchPaths = @(
            $env:SYSTEM_ARTIFACTSDIRECTORY,
            $env:SYSTEM_DEFAULTWORKINGDIRECTORY,
            "D:\a"
        )
        
        foreach ($root in $searchPaths) {
            if ($root -and (Test-Path $root)) {
                $solutionFile = Get-ChildItem -Path $root -Recurse -Filter "solution-managed.zip" -ErrorAction SilentlyContinue | Select-Object -First 1
                if ($solutionFile) {
                    $foundDir = $solutionFile.DirectoryName
                    Write-StatusMessage "Found artifacts at: $foundDir" -Type Success
                    return $foundDir
                }
            }
        }
        
        throw "Could not find artifact directory containing solution-managed.zip"
    }
    else {
        if (Test-Path $ArtifactDirectory) {
            return $ArtifactDirectory
        }
        else {
            throw "Specified artifact directory does not exist: $ArtifactDirectory"
        }
    }
}

function Extract-WebResourcesFromSolution {
    param([string]$SolutionPath)
    
    Write-StatusMessage "Extracting web resources from solution file..." -Type Info
    
    $tempExtractPath = Join-Path $env:TEMP "SolutionExtract_$(Get-Random)"
    New-Item -ItemType Directory -Path $tempExtractPath -Force | Out-Null
    
    try {
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        [System.IO.Compression.ZipFile]::ExtractToDirectory($SolutionPath, $tempExtractPath)
        
        # Look for customizations.xml or solution.xml
        $customizationsPath = Join-Path $tempExtractPath "customizations.xml"
        $solutionXmlPath = Join-Path $tempExtractPath "solution.xml"
        
        $xmlPath = $null
        if (Test-Path $customizationsPath) {
            $xmlPath = $customizationsPath
        }
        elseif (Test-Path $solutionXmlPath) {
            $xmlPath = $solutionXmlPath
        }
        else {
            Write-StatusMessage "No customizations.xml or solution.xml found in solution" -Type Warning
            return @()
        }
        
        Write-StatusMessage "Reading solution metadata from: $xmlPath" -Type Info
        
        # Load the XML
        [xml]$solutionXml = Get-Content $xmlPath
        
        # Find WebResources node
        $webResourceNodes = $solutionXml.SelectNodes("//WebResource")
        
        if ($null -eq $webResourceNodes -or $webResourceNodes.Count -eq 0) {
            Write-StatusMessage "No web resources found in solution metadata" -Type Warning
            return @()
        }
        
        Write-StatusMessage "Found $($webResourceNodes.Count) web resources in solution metadata" -Type Info
        
        $webResources = @()
        
        foreach ($wrNode in $webResourceNodes) {
            $name = $wrNode.Name
            $type = [int]$wrNode.WebResourceType
            $fileName = $wrNode.FileName
            
            # Only process HTML (1) and JavaScript (3) types
            if ($type -ne 1 -and $type -ne 3) {
                continue
            }
            
            # Check if name matches our criteria
            if ($name -notmatch '^akoya_' -and $name -notmatch '^Akoya_' -and $name -notmatch 'ccount_main_library') {
                continue
            }
            
            # Build full path to the file in the extracted solution
            $fullFilePath = Join-Path $tempExtractPath $fileName
            
            if (-not (Test-Path $fullFilePath)) {
                Add-Warning "Web resource file not found: $fileName for $name"
                continue
            }
            
            $typeDesc = if ($type -eq 1) { "HTML" } else { "JavaScript" }
            Write-StatusMessage "  Found $typeDesc web resource: $name" -Type Info
            
            $webResources += @{
                Name = $name
                FilePath = $fullFilePath
                Content = [System.IO.File]::ReadAllBytes($fullFilePath)
                Base64Content = [Convert]::ToBase64String([System.IO.File]::ReadAllBytes($fullFilePath))
                Type = $type
            }
        }
        
        Write-StatusMessage "Extracted $($webResources.Count) matching HTML/JS web resources" -Type Success
        return $webResources
    }
    finally {
        if (Test-Path $tempExtractPath) {
            Remove-Item -Path $tempExtractPath -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}

function Get-WebResourceFromEnvironment {
    param(
        [object]$Connection,
        [string]$Name
    )
    
    $fetchXml = @"
<fetch>
  <entity name='webresource'>
    <attribute name='webresourceid' />
    <attribute name='name' />
    <attribute name='content' />
    <attribute name='webresourcetype' />
    <filter>
      <condition attribute='name' operator='eq' value='$Name' />
    </filter>
  </entity>
</fetch>
"@
    
    $result = Get-CrmRecordsByFetch -conn $Connection -Fetch $fetchXml
    
    if ($result.CrmRecords.Count -gt 0) {
        return $result.CrmRecords[0]
    }
    
    return $null
}

function Remove-WebResourceSolutionLayers {
    param(
        [object]$Connection,
        [string]$WebResourceId,
        [string]$WebResourceName
    )
    
    Write-StatusMessage "  Checking for solution layers on web resource: $WebResourceName" -Type Info
    
    try {
        # Query for solution components
        $fetchXml = @"
<fetch>
  <entity name='msdyn_solutioncomponentsummary'>
    <attribute name='msdyn_objectid' />
    <attribute name='msdyn_componentlogicalname' />
    <attribute name='msdyn_solutionid' />
    <filter>
      <condition attribute='msdyn_objectid' operator='eq' value='$WebResourceId' />
      <condition attribute='msdyn_componentlogicalname' operator='eq' value='webresource' />
    </filter>
    <link-entity name='solution' from='solutionid' to='msdyn_solutionid' alias='sol'>
      <attribute name='uniquename' />
      <attribute name='ismanaged' />
      <filter>
        <condition attribute='ismanaged' operator='eq' value='0' />
      </filter>
    </link-entity>
  </entity>
</fetch>
"@
        
        $layers = Get-CrmRecordsByFetch -conn $Connection -Fetch $fetchXml
        
        if ($layers.CrmRecords.Count -gt 0) {
            Write-StatusMessage "  Found $($layers.CrmRecords.Count) unmanaged solution layer(s) for $WebResourceName" -Type Warning
            
            foreach ($layer in $layers.CrmRecords) {
                $solutionName = $layer.'sol.uniquename'
                
                # Skip if it's the Active solution or Default solution
                if ($solutionName -eq 'Active' -or $solutionName -eq 'Default') {
                    continue
                }
                
                Write-StatusMessage "  Attempting to remove layer from solution: $solutionName" -Type Info
                
                try {
                    # Remove the component from the solution
                    $removeRequest = @{
                        ComponentId = $WebResourceId
                        ComponentType = 61 # Web Resource
                        SolutionUniqueName = $solutionName
                    }
                    
                    # Use RemoveSolutionComponent request
                    $request = New-Object 'Microsoft.Crm.Sdk.Messages.RemoveSolutionComponentRequest'
                    $request.ComponentId = [guid]$WebResourceId
                    $request.ComponentType = 61
                    $request.SolutionUniqueName = $solutionName
                    
                    $response = $Connection.Execute($request)
                    Add-Fix "Removed web resource layer from solution: $solutionName"
                }
                catch {
                    Add-Warning "Could not remove layer from solution $solutionName`: $_"
                }
            }
        }
    }
    catch {
        # If we can't check layers, just continue
        Write-StatusMessage "  Could not check solution layers (may not be available in this environment)" -Type Info
    }
}

function Compare-WebResources {
    param(
        [object]$Connection,
        [array]$SolutionWebResources
    )
    
    Write-StatusMessage "`nValidating Web Resources (HTML and JavaScript only)..." -Type Info
    
    if ($SolutionWebResources.Count -eq 0) {
        Write-StatusMessage "No HTML or JavaScript web resources found matching criteria" -Type Info
        return
    }
    
    # Sort web resources by Type first, then by Name for better readability
    $sortedWebResources = $SolutionWebResources | Sort-Object -Property @{Expression={$_.Type}; Ascending=$true}, @{Expression={$_.Name}; Ascending=$true}
    
    Write-StatusMessage "Processing $($sortedWebResources.Count) web resources..." -Type Info
    
    foreach ($wr in $sortedWebResources) {
        $typeDesc = if ($wr.Type -eq 1) { "HTML" } else { "JavaScript" }
        Write-StatusMessage "Checking $typeDesc web resource: $($wr.Name)" -Type Info
        
        $targetWr = Get-WebResourceFromEnvironment -Connection $Connection -Name $wr.Name
        
        if ($null -eq $targetWr) {
            Add-Warning "Web resource '$($wr.Name)' not found in target environment - attempting to create"
            
            # Attempt to create the web resource
            try {
                $newWr = @{
                    name = $wr.Name
                    content = $wr.Base64Content
                    webresourcetype = $wr.Type
                    displayname = $wr.Name
                }
                
                $newId = New-CrmRecord -conn $Connection -EntityLogicalName webresource -Fields $newWr
                Add-Fix "Created missing web resource: $($wr.Name)"
            }
            catch {
                Add-Failure "Failed to create web resource '$($wr.Name)': $_"
            }
        }
        else {
            # Compare content
            $targetContent = $targetWr.content
            
            if ($targetContent -ne $wr.Base64Content) {
                Add-Warning "Web resource '$($wr.Name)' content mismatch - attempting to fix"
                
                # First, remove any custom solution layers
                Remove-WebResourceSolutionLayers -Connection $Connection -WebResourceId $targetWr.webresourceid -WebResourceName $wr.Name
                
                # Now update the content
                try {
                    $updateFields = @{
                        webresourceid = $targetWr.webresourceid
                        content = $wr.Base64Content
                    }
                    
                    Set-CrmRecord -conn $Connection -EntityLogicalName webresource -Id $targetWr.webresourceid -Fields $updateFields
                    Write-StatusMessage "  Updated web resource content: $($wr.Name)" -Type Info
                    
                    # Re-verify the content after update
                    Start-Sleep -Seconds 2  # Give the system a moment to process
                    
                    $verifyWr = Get-WebResourceFromEnvironment -Connection $Connection -Name $wr.Name
                    
                    if ($null -eq $verifyWr) {
                        Add-Failure "Web resource '$($wr.Name)' disappeared after update"
                    }
                    elseif ($verifyWr.content -ne $wr.Base64Content) {
                        Add-Failure "Web resource '$($wr.Name)' content still does not match after update - may have active customization layers"
                    }
                    else {
                        Add-Fix "Updated and verified web resource content: $($wr.Name)"
                    }
                }
                catch {
                    Add-Failure "Failed to update web resource '$($wr.Name)': $_"
                }
            }
            else {
                Add-Success "Web resource '$($wr.Name)' matches"
            }
        }
    }
}

function Get-PluginAssemblyInfo {
    param(
        [object]$Connection,
        [string]$AssemblyName
    )
    
    $fetchXml = @"
<fetch>
  <entity name='pluginassembly'>
    <attribute name='pluginassemblyid' />
    <attribute name='name' />
    <attribute name='version' />
    <filter>
      <condition attribute='name' operator='eq' value='$AssemblyName' />
    </filter>
  </entity>
</fetch>
"@
    
    $result = Get-CrmRecordsByFetch -conn $Connection -Fetch $fetchXml
    
    if ($result.CrmRecords.Count -gt 0) {
        return $result.CrmRecords[0]
    }
    
    return $null
}

function Get-PluginStepFromEnvironment {
    param(
        [object]$Connection,
        [string]$StepId
    )
    
    $fetchXml = @"
<fetch>
  <entity name='sdkmessageprocessingstep'>
    <attribute name='sdkmessageprocessingstepid' />
    <attribute name='name' />
    <attribute name='description' />
    <attribute name='configuration' />
    <attribute name='mode' />
    <attribute name='rank' />
    <attribute name='stage' />
    <attribute name='statecode' />
    <attribute name='impersonatinguserid' />
    <attribute name='asyncautodelete'/>
    <attribute name='plugintypeid' />
    <filter>
      <condition attribute='sdkmessageprocessingstepid' operator='eq' value='$StepId' />
    </filter>
    <link-entity name='systemuser' from='systemuserid' to='impersonatinguserid' link-type='outer' alias='impuser'>
      <attribute name='systemuserid' />
    </link-entity>
    <link-entity name='sdkmessagefilter' from='sdkmessagefilterid' to='sdkmessagefilterid' link-type='outer' alias='filter'>
      <attribute name='primaryobjecttypecode' />
    </link-entity>
    <link-entity name='sdkmessage' from='sdkmessageid' to='sdkmessageid' link-type='inner' alias='message'>
      <attribute name='name' />
    </link-entity>
    <link-entity name='sdkmessageprocessingstepimage' from='sdkmessageprocessingstepid' to='sdkmessageprocessingstepid' link-type='outer' alias='image'>
      <attribute name='sdkmessageprocessingstepimageid' />
      <attribute name='name' />
      <attribute name='entityalias' />
      <attribute name='imagetype' />
      <attribute name='messagepropertyname' />
    </link-entity>
  </entity>
</fetch>
"@
    
    try {
        $result = Get-CrmRecordsByFetch -conn $Connection -Fetch $fetchXml
        
        if ($result.CrmRecords.Count -gt 0) {
            $record = $result.CrmRecords[0]
            
            # Extract the impersonating user ID from the linked entity
            if ($record.'impuser.systemuserid') {
                $record | Add-Member -NotePropertyName "impersonatinguserid_value" -NotePropertyValue $record.'impuser.systemuserid' -Force
            }
            
            # Extract primary entity name
            if ($record.'filter.primaryobjecttypecode') {
                $record | Add-Member -NotePropertyName "primaryentityname" -NotePropertyValue $record.'filter.primaryobjecttypecode' -Force
            }
            
            # Extract SDK message name
            if ($record.'message.name') {
                $record | Add-Member -NotePropertyName "sdkmessagename" -NotePropertyValue $record.'message.name' -Force
            }
            
            # Build the images array from linked entity results
            $images = @()
            foreach ($rec in $result.CrmRecords) {
                if ($rec.'image.sdkmessageprocessingstepimageid') {
                    $image = @{
                        sdkmessageprocessingstepimageid = $rec.'image.sdkmessageprocessingstepimageid'
                        name = $rec.'image.name'
                        entityAlias = $rec.'image.entityalias'
                        imageType = $rec.'image.imagetype'
                        messagePropertyName = $rec.'image.messagepropertyname'
                    }
                    
                    # Only add unique images (in case of multiple link-entity matches)
                    if (-not ($images | Where-Object { $_.sdkmessageprocessingstepimageid -eq $image.sdkmessageprocessingstepimageid })) {
                        $images += $image
                    }
                }
            }

            # Add images array to the record
            $record | Add-Member -NotePropertyName "images" -NotePropertyValue $images -Force
            
            return $record
        }
    }
    catch {
        # Step doesn't exist
    }
    
    return $null
}

function Get-SystemUserByApplicationId {
    param(
        [object]$Connection,
        [string]$ApplicationId
    )
    
    $fetchXml = @"
<fetch>
  <entity name='systemuser'>
    <attribute name='systemuserid' />
    <attribute name='applicationid' />
    <attribute name='fullname' />
    <filter>
      <condition attribute='applicationid' operator='eq' value='$ApplicationId' />
    </filter>
  </entity>
</fetch>
"@
    
    $result = Get-CrmRecordsByFetch -conn $Connection -Fetch $fetchXml
    
    if ($result.CrmRecords.Count -gt 0) {
        return $result.CrmRecords[0]
    }
    
    return $null
}

function Compare-PluginStepProperty {
    param(
        [object]$SourceStep,
        [object]$TargetStep,
        [string]$PropertyName
    )
    
    $sourceValue = $SourceStep.$PropertyName
    $targetValue = $TargetStep.$PropertyName
    
    # Normalize null and empty strings to be treated as equal
    $normalizedSource = if ([string]::IsNullOrEmpty($sourceValue)) { "" } else { $sourceValue.ToString() }
    $normalizedTarget = if ([string]::IsNullOrEmpty($targetValue)) { "" } else { $targetValue.ToString() }
    
    # Compare normalized values
    return $normalizedSource -eq $normalizedTarget
}

function Update-PluginStepInEnvironment {
    param(
        [object]$Connection,
        [object]$SourceStep,
        [object]$TargetStep
    )
    
    $updateFields = @{
        sdkmessageprocessingstepid = $TargetStep.sdkmessageprocessingstepid
    }
    
    # Map string properties
    if ($SourceStep.name) { $updateFields['name'] = $SourceStep.name }
    if ($SourceStep.description) { $updateFields['description'] = $SourceStep.description }
    if ($SourceStep.configuration) { $updateFields['configuration'] = $SourceStep.configuration }
    
    # Map rank (integer)
    if ($SourceStep.rank) { $updateFields['rank'] = [int]$SourceStep.rank }
    
    # Map mode (OptionSetValue)
    if ($SourceStep.mode) {
        $modeValue = if ($SourceStep.mode -eq 'Synchronous') { 0 } else { 1 }
        $updateFields['mode'] = New-CrmOptionSetValue -Value $modeValue
    }
    
    # Map stage (OptionSetValue)
    if ($SourceStep.stage) {
        $stageValue = switch ($SourceStep.stage) {
            'Pre-validation' { 10 }
            'Pre-operation' { 20 }
            'Post-operation' { 40 }
            default { 40 }
        }
        $updateFields['stage'] = New-CrmOptionSetValue -Value $stageValue
    }
    
    # Map state (StateCode - also needs OptionSetValue)
    if ($SourceStep.state) {
        $stateValue = if ($SourceStep.state -eq 'Enabled') { 0 } else { 1 }
        $updateFields['statecode'] = New-CrmOptionSetValue -Value $stateValue
    }
    
    # Map asyncAutoDelete (boolean)
    if ($SourceStep.asyncAutoDelete) {
        $updateFields['asyncautodelete'] = if ($SourceStep.asyncAutoDelete -eq 'Yes') { $true } else { $false }
    }
    
    # Handle RunAsUser (EntityReference)
    if ($SourceStep.runAsUser.applicationId) {
        $expectedUser = Get-SystemUserByApplicationId -Connection $Connection -ApplicationId $SourceStep.runAsUser.applicationId
        if ($expectedUser) {
            $updateFields['impersonatinguserid'] = New-CrmEntityReference -EntityLogicalName 'systemuser' -Id $expectedUser.systemuserid
        }
        else {
            Add-Warning "Could not find service user with applicationId: $($SourceStep.runAsUser.applicationId) for step: $($SourceStep.name)"
        }
    }
    elseif ($SourceStep.runAsUser.systemUserId) {
        $updateFields['impersonatinguserid'] = New-CrmEntityReference -EntityLogicalName 'systemuser' -Id $SourceStep.runAsUser.systemUserId
    }
    else {
        # Running as calling user - set impersonatinguserid to null
        $updateFields['impersonatinguserid'] = $null
    }
    
    Set-CrmRecord -conn $Connection -EntityLogicalName sdkmessageprocessingstep -Id $TargetStep.sdkmessageprocessingstepid -Fields $updateFields
    
    # Handle images synchronization
    # 1. Remove extra images that exist in target but not in source
    if ($TargetStep.images -and $TargetStep.images.Count -gt 0) {
        foreach ($targetImage in $TargetStep.images) {
            # Check if this image exists in the source
            $sourceImage = $null
            if ($SourceStep.images) {
                $sourceImage = $SourceStep.images | Where-Object { $_.name -eq $targetImage.name }
            }
            
            if (-not $sourceImage) {
                # Image exists in target but not in source - remove it
                Write-StatusMessage "  Removing extra image '$($targetImage.name)' from step '$($SourceStep.name)'" -Type Info
                
                try {
                    Remove-CrmRecord -conn $Connection -EntityLogicalName sdkmessageprocessingstepimage -Id $targetImage.sdkmessageprocessingstepimageid
                    Add-Fix "Removed extra image '$($targetImage.name)' from step '$($SourceStep.name)'"
                }
                catch {
                    Add-Failure "Failed to remove image '$($targetImage.name)' from step '$($SourceStep.name)': $_"
                }
            }
        }
    }
    
    # 2. Create missing images that exist in source but not in target
    if ($SourceStep.images -and $SourceStep.images.Count -gt 0) {
        foreach ($sourceImage in $SourceStep.images) {
            # Check if this image exists in the target
            $targetImage = $null
            if ($TargetStep.images) {
                $targetImage = $TargetStep.images | Where-Object { $_.name -eq $sourceImage.name }
            }
            
            if (-not $targetImage) {
                # Image is missing in target - create it
                Write-StatusMessage "  Creating missing image '$($sourceImage.name)' for step '$($SourceStep.name)'" -Type Info
                
                try {
                    # Convert imageType text to numeric value
                    $imageTypeValue = switch ($sourceImage.imageType) {
                        'PreImage' { 0 }
                        'PostImage' { 1 }
                        'Both' { 2 }
                        default { 0 }  # Default to PreImage
                    }
                    
                    $imageFields = @{
                        sdkmessageprocessingstepid = New-CrmEntityReference -EntityLogicalName 'sdkmessageprocessingstep' -Id $TargetStep.sdkmessageprocessingstepid
                        name = $sourceImage.name
                        entityalias = $sourceImage.entityAlias
                        imagetype = New-CrmOptionSetValue -Value $imageTypeValue
                        messagepropertyname = $sourceImage.messagePropertyName
                        attributes = ""  # Empty string for all attributes
                    }
                    
                    $newImageId = New-CrmRecord -conn $Connection -EntityLogicalName sdkmessageprocessingstepimage -Fields $imageFields
                    Add-Fix "Created missing image '$($sourceImage.name)' for step '$($SourceStep.name)'"
                }
                catch {
                    Add-Failure "Failed to create image '$($sourceImage.name)' for step '$($SourceStep.name)': $_"
                }
            }
        }
    }
}

function Validate-PluginStepImages {
    param(
        [object]$SourceStep
    )
    
    $issues = @()
    
    # Check if Update or Delete step requires PreImage
    if ($SourceStep.message -eq 'Update' -or $SourceStep.message -eq 'Delete') {
        # Check if images array exists and has at least one PreImage
        $hasPreImage = $false
        
        if ($SourceStep.images -and $SourceStep.images.Count -gt 0) {
            foreach ($image in $SourceStep.images) {
                if ($image.name -eq 'PreImage') {
                    $hasPreImage = $true
                    break
                }
            }
        }
        
        if (-not $hasPreImage) {
            $issues += "Missing PreImage for $($SourceStep.message) step"
        }
    }
    
    return $issues
}

function Compare-PluginStepImages {
    param(
        [object]$Connection,
        [object]$SourceStep,
        [object]$TargetStep
    )
    
    $mismatch = $false
    
    # Compare number of images
    $sourceImageCount = if ($SourceStep.images) { $SourceStep.images.Count } else { 0 }
    $targetImageCount = if ($TargetStep.images) { $TargetStep.images.Count } else { 0 }
    
    if ($sourceImageCount -ne $targetImageCount) {
        Add-Warning "Image count mismatch on step '$($SourceStep.name)': Expected $sourceImageCount images, Found $targetImageCount images"
        $mismatch = $true
    }
    
    # Compare each image
    foreach ($sourceImage in $SourceStep.images) {
        # Find matching image in target by name
        $targetImage = $TargetStep.images | Where-Object { $_.name -eq $sourceImage.name }
        
        if (-not $targetImage) {
            Add-Warning "Missing image '$($sourceImage.name)' on step '$($SourceStep.name)'"
            $mismatch = $true
            continue
        }
        
        # Compare image properties
        if ($sourceImage.entityAlias -ne $targetImage.entityAlias) {
            Add-Warning "Image entityAlias mismatch on step '$($SourceStep.name)', image '$($sourceImage.name)': Expected '$($sourceImage.entityAlias)', Found '$($targetImage.entityAlias)'"
            $mismatch = $true
        }
        
        if ($sourceImage.messagePropertyName -ne $targetImage.messagePropertyName) {
            Add-Warning "Image messagePropertyName mismatch on step '$($SourceStep.name)', image '$($sourceImage.name)': Expected '$($sourceImage.messagePropertyName)', Found '$($targetImage.messagePropertyName)'"
            $mismatch = $true
        }
    }
    
    return $mismatch
}

function Compare-PluginSteps {
    param(
        [object]$Connection,
        [string]$PluginStepsJsonPath,
        [string]$AssemblyName
    )
    
    Write-StatusMessage "`nValidating Plugin Steps..." -Type Info
    
    # Load JSON
    $jsonContent = Get-Content $PluginStepsJsonPath -Raw | ConvertFrom-Json
    $expectedVersion = $jsonContent.metadata.assemblyVersion
    
    # Check assembly exists and version matches
    $assembly = Get-PluginAssemblyInfo -Connection $Connection -AssemblyName $AssemblyName
    
    if ($null -eq $assembly) {
        Add-Failure "Plugin assembly '$AssemblyName' not found in target environment"
        return
    }
    
    if ($assembly.version -ne $expectedVersion) {
        Add-Failure "Plugin assembly version mismatch. Expected: $expectedVersion, Found: $($assembly.version)"
        return
    }
    
    Write-StatusMessage "Assembly '$AssemblyName' version $($assembly.version) validated" -Type Success
    
    # Check each step
    foreach ($sourceStep in $jsonContent.pluginSteps) {
        $stepId = $sourceStep.sdkmessageprocessingstepid
        Write-StatusMessage "Checking plugin step: $($sourceStep.name)" -Type Info
        
        # Validate PreImage requirements for Update and Delete steps
        $imageIssues = Validate-PluginStepImages -SourceStep $sourceStep
        if ($imageIssues.Count -gt 0) {
            foreach ($issue in $imageIssues) {
                Add-Warning "$($sourceStep.name): $issue"
            }
        }
        
        $targetStep = Get-PluginStepFromEnvironment -Connection $Connection -StepId $stepId

        if ($null -eq $targetStep) {
            Add-Warning "Plugin step '$($sourceStep.name)' (ID: $stepId) not found in target environment - attempting to create"
            
            # Attempt to create the plugin step
            $created = New-PluginStepInEnvironment -Connection $Connection -SourceStep $sourceStep
            
            if ($created) {
                # Successfully created, move to next step
                continue
            }
            else {
                # Creation failed, error already logged
                continue
            }
        }
        
        # ===== VALIDATE IMMUTABLE PROPERTIES =====

        # These properties cannot be updated - if they don't match, the step needs to be recreated manually
        $immutableMismatch = $false
        
        # Compare plugintypeid (immutable)
        if ($sourceStep.plugintypeid -ne $targetStep.plugintypeid) {
            $immutableMismatch = $true
            Add-Failure "CRITICAL: Plugin Type ID mismatch on step '$($sourceStep.name)': Expected '$($sourceStep.plugintypeid)', Found '$($targetStep.plugintypeid)' - STEP MUST BE RECREATED"
        }
        
        # Compare primaryEntity (immutable) - treat null, empty string, and "none" as equal
        $sourcePrimaryEntity = if ([string]::IsNullOrEmpty($sourceStep.primaryEntity) -or $sourceStep.primaryEntity -eq 'none') { "" } else { $sourceStep.primaryEntity }
        $targetPrimaryEntity = if ([string]::IsNullOrEmpty($targetStep.primaryentityname) -or $targetStep.primaryentityname -eq 'none') { "" } else { $targetStep.primaryentityname }
        if ($sourcePrimaryEntity -ne $targetPrimaryEntity) {
            $immutableMismatch = $true
            Add-Failure "CRITICAL: Primary Entity mismatch on step '$($sourceStep.name)': Expected '$($sourceStep.primaryEntity)', Found '$($targetStep.primaryentityname)' - STEP MUST BE RECREATED"
        }

        # Compare message (immutable) - need to get message name from target
        # The target returns the message name, JSON has the message name
        if ($sourceStep.message -ne $targetStep.sdkmessagename) {
            $immutableMismatch = $true
            Add-Failure "CRITICAL: Message mismatch on step '$($sourceStep.name)': Expected '$($sourceStep.message)', Found '$($targetStep.sdkmessagename)' - STEP MUST BE RECREATED"
        }
        
        # If immutable properties don't match, skip updating mutable properties
        if ($immutableMismatch) {
            Add-Failure "Step '$($sourceStep.name)' has critical mismatches and cannot be auto-fixed. Manual intervention required."
            continue
        }

        # ===== COMPARE MUTABLE PROPERTIES =====

        # Compare string properties (using null/empty normalization)
        $mismatch = $false
        $propertiesToCheck = @('name', 'description', 'configuration')
        
        foreach ($prop in $propertiesToCheck) {
            if (-not (Compare-PluginStepProperty -SourceStep $sourceStep -TargetStep $targetStep -PropertyName $prop)) {
                $mismatch = $true
                Add-Warning "Property mismatch on step '$($sourceStep.name)': $prop (Expected: '$($sourceStep.$prop)', Found: '$($targetStep.$prop)')"
            }
        }
        
        # Compare rank (convert both to int)
        $expectedRank = [int]$sourceStep.rank
        $actualRank = [int]$targetStep.rank
        if ($actualRank -ne $expectedRank) {
            $mismatch = $true
            Add-Warning "Rank mismatch on step '$($sourceStep.name)': Expected $expectedRank, Found $actualRank"
        }
        
        # Compare mode (target returns text label like "Asynchronous" or "Synchronous")
        $expectedMode = $sourceStep.mode  # Already text in JSON
        $actualMode = $targetStep.mode.ToString()
        if ($actualMode -ne $expectedMode) {
            $mismatch = $true
            Add-Warning "Mode mismatch on step '$($sourceStep.name)': Expected $expectedMode, Found $actualMode"
        }
        
        # Compare stage (target returns text label like "Pre-validation" or "Pre-operation" or "Post-operation")
        $expectedStage = $sourceStep.stage
        $actualStage = $targetStep.stage
        if ($actualStage -ne $expectedStage) {
            $mismatch = $true
            Add-Warning "Stage mismatch on step '$($sourceStep.name)': Expected $($sourceStep.stage), Found $($targetStep.stage)"
        }       
 
        # Compare state (target returns text label like "Enabled" or "Disabled")
        $expectedState = $sourceStep.state  # Already text in JSON
        $actualState = $targetStep.statecode.ToString()
        if ($actualState -ne $expectedState) {
            $mismatch = $true
            Add-Warning "State mismatch on step '$($sourceStep.name)': Expected $expectedState, Found $actualState"
        }
        
        # Compare asyncAutoDelete (normalize null to "No", then compare text-to-text)
        $expectedAsyncAutoDelete = if ([string]::IsNullOrEmpty($sourceStep.asyncAutoDelete)) { "No" } else { $sourceStep.asyncAutoDelete.ToString() }
        $actualAsyncAutoDelete = if ([string]::IsNullOrEmpty($targetStep.asyncautodelete)) { "No" } else { $targetStep.asyncautodelete.ToString() }
        if ($actualAsyncAutoDelete -ne $expectedAsyncAutoDelete) {
            $mismatch = $true
            Add-Warning "AsyncAutoDelete mismatch on step '$($sourceStep.name)': Expected $expectedAsyncAutoDelete, Found $actualAsyncAutoDelete"
        }        

        # Check RunAsUser (impersonation)
        $runAsUserMismatch = $false

        if ($sourceStep.runAsUser.applicationId) {
            # Source specifies a user by applicationId
            $expectedUser = Get-SystemUserByApplicationId -Connection $Connection -ApplicationId $sourceStep.runAsUser.applicationId
            if ($expectedUser) {
                # Compare GUIDs as strings (normalize to lowercase)
                $expectedUserId = $expectedUser.systemuserid.ToString().ToLower()
                $actualUserId = if ($targetStep.impersonatinguserid_value) { $targetStep.impersonatinguserid_value.ToString().ToLower() } else { "" }
                if ($actualUserId -ne $expectedUserId) {
                    $runAsUserMismatch = $true
                    Add-Warning "RunAsUser mismatch on step '$($sourceStep.name)': Expected user with appId $($sourceStep.runAsUser.applicationId) (userId: $expectedUserId), Found: $actualUserId"
                }
            }
            else {
                Add-Warning "Could not find service user with applicationId: $($sourceStep.runAsUser.applicationId) for step: $($SourceStep.name)"
            }
        }
        elseif ($sourceStep.runAsUser.systemUserId) {
            # Source specifies a user by systemUserId
            $expectedUserId = $sourceStep.runAsUser.systemUserId.ToString().ToLower()
            $actualUserId = if ($targetStep.impersonatinguserid_value) { $targetStep.impersonatinguserid_value.ToString().ToLower() } else { "" }
    
            if ($actualUserId -ne $expectedUserId) {
                $runAsUserMismatch = $true
                Add-Warning "RunAsUser mismatch on step '$($sourceStep.name)': Expected systemUserId $expectedUserId, Found: $actualUserId"
            }
        }
        else {
            # Running as calling user - impersonatinguserid should be null or empty
            if ($targetStep.impersonatinguserid_value) {
                $runAsUserMismatch = $true
                Add-Warning "RunAsUser mismatch on step '$($sourceStep.name)': Should run as calling user (no impersonation), but found userId: $($targetStep.impersonatinguserid_value)"
            }
        }

        if ($runAsUserMismatch) {
            $mismatch = $true
        }
        
        # Compare images (PreImages, PostImages, etc.)
        $imagesMismatch = Compare-PluginStepImages -Connection $Connection -SourceStep $sourceStep -TargetStep $targetStep
        if ($imagesMismatch) {
            $mismatch = $true
        }
        
        if ($mismatch) {
            try {
                Update-PluginStepInEnvironment -Connection $Connection -SourceStep $sourceStep -TargetStep $targetStep
                Add-Fix "Updated plugin step: $($sourceStep.name)"
            }
            catch {
                Add-Failure "Failed to update plugin step '$($sourceStep.name)': $_"
            }
        }
        else {
            Add-Success "Plugin step '$($sourceStep.name)' matches"
        }

    }
}

function Get-SdkMessageId {
    param(
        [object]$Connection,
        [string]$MessageName
    )
    
    $fetchXml = @"
<fetch top='1'>
  <entity name='sdkmessage'>
    <attribute name='sdkmessageid' />
    <attribute name='name' />
    <filter>
      <condition attribute='name' operator='eq' value='$MessageName' />
    </filter>
  </entity>
</fetch>
"@
    
    $result = Get-CrmRecordsByFetch -conn $Connection -Fetch $fetchXml
    
    if ($result.CrmRecords.Count -gt 0) {
        return $result.CrmRecords[0].sdkmessageid
    }
    
    return $null
}

function Get-SdkMessageFilterId {
    param(
        [object]$Connection,
        [string]$MessageName,
        [string]$PrimaryEntity
    )
    
    # For messages that don't have entity filters (like custom actions on "none")
    if ([string]::IsNullOrEmpty($PrimaryEntity) -or $PrimaryEntity -eq 'none') {
        return $null
    }
    
    $fetchXml = @"
<fetch top='1'>
  <entity name='sdkmessagefilter'>
    <attribute name='sdkmessagefilterid' />
    <attribute name='primaryobjecttypecode' />
    <link-entity name='sdkmessage' from='sdkmessageid' to='sdkmessageid' alias='msg'>
      <filter>
        <condition attribute='name' operator='eq' value='$MessageName' />
      </filter>
    </link-entity>
    <filter>
      <condition attribute='primaryobjecttypecode' operator='eq' value='$PrimaryEntity' />
    </filter>
  </entity>
</fetch>
"@
    
    $result = Get-CrmRecordsByFetch -conn $Connection -Fetch $fetchXml
    
    if ($result.CrmRecords.Count -gt 0) {
        return $result.CrmRecords[0].sdkmessagefilterid
    }
    
    return $null
}

function Get-PluginTypeId {
    param(
        [object]$Connection,
        [string]$PluginTypeIdFromJson
    )
    
    # If we have the GUID from JSON, try to look up directly
    $fetchXml = @"
<fetch top='1'>
  <entity name='plugintype'>
    <attribute name='plugintypeid' />
    <attribute name='typename' />
    <filter>
      <condition attribute='plugintypeid' operator='eq' value='$PluginTypeIdFromJson' />
    </filter>
  </entity>
</fetch>
"@
    
    $result = Get-CrmRecordsByFetch -conn $Connection -Fetch $fetchXml
    
    if ($result.CrmRecords.Count -gt 0) {
        return $result.CrmRecords[0].plugintypeid
    }
    
    return $null
}

# =============================================================================
# MAIN SCRIPT EXECUTION
# =============================================================================

Write-Host "==================================================" -ForegroundColor Cyan
Write-Host "  Dynamics 365 Deployment Health Check" -ForegroundColor Cyan
Write-Host "==================================================" -ForegroundColor Cyan
Write-Host ""

try {
    # Import required module
    Import-Module Microsoft.Xrm.Data.PowerShell -ErrorAction Stop
    Write-StatusMessage "Microsoft.Xrm.Data.PowerShell module loaded" -Type Success
}
catch {
    Write-StatusMessage "Failed to load Microsoft.Xrm.Data.PowerShell module. Install it using: Install-Module Microsoft.Xrm.Data.PowerShell" -Type Error
    exit 1
}

# Find artifact directory
try {
    $artifactDir = Find-ArtifactDirectory
    Write-StatusMessage "Using artifact directory: $artifactDir" -Type Info
}
catch {
    Write-StatusMessage $_.Exception.Message -Type Error
    exit 1
}

# Locate required files
$solutionPath = Join-Path $artifactDir "solution-managed.zip"
$pluginStepsPath = Join-Path $artifactDir "plugin-steps.json"

if (-not (Test-Path $solutionPath)) {
    Write-StatusMessage "solution-managed.zip not found at: $solutionPath" -Type Error
    exit 1
}

if (-not (Test-Path $pluginStepsPath)) {
    Write-StatusMessage "plugin-steps.json not found at: $pluginStepsPath" -Type Error
    exit 1
}

Write-StatusMessage "Found solution-managed.zip" -Type Success
Write-StatusMessage "Found plugin-steps.json" -Type Success

# Connect to Dynamics 365
Write-StatusMessage "`nConnecting to Dynamics 365: $EnvironmentUrl" -Type Info

try {
    $connection = Connect-CrmOnline -ServerUrl $EnvironmentUrl `
        -ClientSecret $ClientSecret `
        -OAuthClientId $ClientId
    
    if ($connection.IsReady) {
        Write-StatusMessage "Connected to Dynamics 365" -Type Success
    }
    else {
        throw "Connection is not ready"
    }
}
catch {
    Write-StatusMessage "Failed to connect to Dynamics 365: $_" -Type Error
    exit 1
}

# Extract and validate web resources
try {
    $webResources = Extract-WebResourcesFromSolution -SolutionPath $solutionPath
    Compare-WebResources -Connection $connection -SolutionWebResources $webResources
}
catch {
    Write-StatusMessage "Error during web resource validation: $_" -Type Error
    Add-Failure "Web resource validation failed: $_"
}

# Validate plugin steps
try {
    Compare-PluginSteps -Connection $connection -PluginStepsJsonPath $pluginStepsPath -AssemblyName "AkoyaGo.Plugins"
}
catch {
    Write-StatusMessage "Error during plugin step validation: $_" -Type Error
    Add-Failure "Plugin step validation failed: $_"
}

# Disconnect
if ($connection -and $connection.IsReady) {
    $connection.Dispose()
}

# =============================================================================
# FINAL REPORT
# =============================================================================

Write-Host "`n==================================================" -ForegroundColor Cyan
Write-Host "  Health Check Summary" -ForegroundColor Cyan
Write-Host "==================================================" -ForegroundColor Cyan

Write-Host "`nSuccesses: $($script:successes.Count)" -ForegroundColor Green
Write-Host "Fixes Applied: $($script:fixes.Count)" -ForegroundColor Yellow
Write-Host "Warnings: $($script:warnings.Count)" -ForegroundColor Yellow
Write-Host "Failures: $($script:failures.Count)" -ForegroundColor Red

if ($script:fixes.Count -gt 0) {
    Write-Host "`nFixes Applied:" -ForegroundColor Yellow
    foreach ($fix in $script:fixes) {
        Write-Host "  - $fix" -ForegroundColor Yellow
    }
}

if ($script:warnings.Count -gt 0) {
    Write-Host "`nWarnings:" -ForegroundColor Yellow
    foreach ($warning in $script:warnings) {
        Write-Host "  - $warning" -ForegroundColor Yellow
    }
}

if ($script:failures.Count -gt 0) {
    Write-Host "`nFailures:" -ForegroundColor Red
    foreach ($failure in $script:failures) {
        Write-Host "  - $failure" -ForegroundColor Red
    }
    
    Write-Host "`n[FAIL] Health check completed with $($script:failures.Count) unresolved issue(s)" -ForegroundColor Red
    throw "Health check failed with $($script:failures.Count) unresolved issue(s)"
}
else {
    Write-Host "`n[OK] Health check completed successfully" -ForegroundColor Green

}

function New-PluginStepInEnvironment {
    param(
        [object]$Connection,
        [object]$SourceStep
    )
    
    Write-StatusMessage "  Creating new plugin step: $($SourceStep.name)" -Type Info
    
    try {
        # Get SDK Message ID
        $sdkMessageId = Get-SdkMessageId -Connection $Connection -MessageName $SourceStep.message
        if (-not $sdkMessageId) {
            throw "SDK Message '$($SourceStep.message)' not found"
        }
        
        # Get SDK Message Filter ID (if applicable)
        $sdkMessageFilterId = $null
        if (![string]::IsNullOrEmpty($SourceStep.primaryEntity) -and $SourceStep.primaryEntity -ne 'none') {
            $sdkMessageFilterId = Get-SdkMessageFilterId -Connection $Connection -MessageName $SourceStep.message -PrimaryEntity $SourceStep.primaryEntity
            if (-not $sdkMessageFilterId) {
                throw "SDK Message Filter for message '$($SourceStep.message)' and entity '$($SourceStep.primaryEntity)' not found"
            }
        }
        
        # Verify Plugin Type exists
        $pluginTypeId = Get-PluginTypeId -Connection $Connection -PluginTypeIdFromJson $SourceStep.plugintypeid
        if (-not $pluginTypeId) {
            throw "Plugin Type with ID '$($SourceStep.plugintypeid)' not found in target environment"
        }
        
        # Build the step fields
        $stepFields = @{
            name = $SourceStep.name
            sdkmessageid = New-CrmEntityReference -EntityLogicalName 'sdkmessage' -Id $sdkMessageId
            plugintypeid = New-CrmEntityReference -EntityLogicalName 'plugintype' -Id $pluginTypeId
        }
        
        # Add optional fields
        if ($SourceStep.description) {
            $stepFields['description'] = $SourceStep.description
        }
        
        if ($SourceStep.configuration) {
            $stepFields['configuration'] = $SourceStep.configuration
        }
        
        # Add rank
        $stepFields['rank'] = [int]$SourceStep.rank
        
        # Add mode (Synchronous = 0, Asynchronous = 1)
        $modeValue = if ($SourceStep.mode -eq 'Synchronous') { 0 } else { 1 }
        $stepFields['mode'] = New-CrmOptionSetValue -Value $modeValue
        
        # Add stage
        $stageValue = switch ($SourceStep.stage) {
            'Pre-validation' { 10 }
            'Pre-operation' { 20 }
            'Post-operation' { 40 }
            default { 40 }
        }
        $stepFields['stage'] = New-CrmOptionSetValue -Value $stageValue
        
        # Add message filter if applicable
        if ($sdkMessageFilterId) {
            $stepFields['sdkmessagefilterid'] = New-CrmEntityReference -EntityLogicalName 'sdkmessagefilter' -Id $sdkMessageFilterId
        }
        
        # Add asyncAutoDelete for async steps
        if ($SourceStep.mode -eq 'Asynchronous' -and $SourceStep.asyncAutoDelete) {
            $stepFields['asyncautodelete'] = if ($SourceStep.asyncAutoDelete -eq 'Yes') { $true } else { $false }
        }
        
        # Handle RunAsUser (impersonation)
        if ($SourceStep.runAsUser.applicationId) {
            $expectedUser = Get-SystemUserByApplicationId -Connection $Connection -ApplicationId $SourceStep.runAsUser.applicationId
            if ($expectedUser) {
                $stepFields['impersonatinguserid'] = New-CrmEntityReference -EntityLogicalName 'systemuser' -Id $expectedUser.systemuserid
            }
            else {
                Add-Warning "Could not find service user with applicationId: $($SourceStep.runAsUser.applicationId) for new step: $($SourceStep.name)"
            }
        }
        elseif ($SourceStep.runAsUser.systemUserId) {
            $stepFields['impersonatinguserid'] = New-CrmEntityReference -EntityLogicalName 'systemuser' -Id $SourceStep.runAsUser.systemUserId
        }
        
        # Create the step with the specific GUID from JSON
        $newStepId = New-CrmRecord -conn $Connection -EntityLogicalName sdkmessageprocessingstep -Fields $stepFields -Guid $SourceStep.sdkmessageprocessingstepid
        
        Write-StatusMessage "  Created plugin step with ID: $newStepId" -Type Success
        
        # Create images if defined
        if ($SourceStep.images -and $SourceStep.images.Count -gt 0) {
            foreach ($sourceImage in $SourceStep.images) {
                Write-StatusMessage "  Creating image '$($sourceImage.name)' for new step" -Type Info
                
                try {
                    # Convert imageType text to numeric value
                    $imageTypeValue = switch ($sourceImage.imageType) {
                        'PreImage' { 0 }
                        'PostImage' { 1 }
                        'Both' { 2 }
                        default { 0 }
                    }
                    
                    $imageFields = @{
                        sdkmessageprocessingstepid = New-CrmEntityReference -EntityLogicalName 'sdkmessageprocessingstep' -Id $newStepId
                        name = $sourceImage.name
                        entityalias = $sourceImage.entityAlias
                        imagetype = New-CrmOptionSetValue -Value $imageTypeValue
                        messagepropertyname = $sourceImage.messagePropertyName
                        attributes = ""  # Empty string for all attributes
                    }
                    
                    $newImageId = New-CrmRecord -conn $Connection -EntityLogicalName sdkmessageprocessingstepimage -Fields $imageFields
                    Write-StatusMessage "  Created image '$($sourceImage.name)'" -Type Success
                }
                catch {
                    Add-Warning "Failed to create image '$($sourceImage.name)' for new step '$($SourceStep.name)': $_"
                }
            }
        }
        
        # Set the state (Enabled/Disabled) after creation
        if ($SourceStep.state -eq 'Disabled') {
            try {
                $setStateRequest = @{
                    EntityMoniker = New-CrmEntityReference -EntityLogicalName 'sdkmessageprocessingstep' -Id $newStepId
                    State = 1  # Disabled
                    Status = 2  # Disabled status
                }
                
                Set-CrmRecordState -conn $Connection -EntityLogicalName sdkmessageprocessingstep -Id $newStepId -StateCode 1 -StatusCode 2
                Write-StatusMessage "  Set step state to Disabled" -Type Success
            }
            catch {
                Add-Warning "Failed to set state to Disabled for new step '$($SourceStep.name)': $_"
            }
        }
        
        Add-Fix "Created new plugin step: $($SourceStep.name)"
        return $true
    }
    catch {
        Add-Failure "Failed to create plugin step '$($SourceStep.name)': $_"
        return $false
    }
}
