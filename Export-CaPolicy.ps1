<#
	.SYNOPSIS
		Conditional Access Export Utility
	.DESCRIPTION
    Exports CA Policy to HTML

	.PARAMETER TenantID
		Optional. The Azure AD tenant ID to connect to. If not specified, connects to the default tenant.

	.PARAMETER PolicyID
		Optional. A specific Conditional Access policy ID (GUID) to export. If not specified, all policies are exported.

	.PARAMETER Csv
		Switch. Generate a CSV file in normalized/flat format for detailed analysis.

	.PARAMETER CsvPivot
		Switch. Generate a pivot-friendly CSV (wide format) for ad-hoc aggregation in Excel / BI tools.

	.EXAMPLE
		PS> .\Export-CaPolicy.ps1
		Exports all Conditional Access policies to HTML format (default).

	.EXAMPLE
		PS> .\Export-CaPolicy.ps1 -Csv
		Exports all policies to both HTML and normalized CSV format.

	.EXAMPLE
		PS> .\Export-CaPolicy.ps1 -CsvPivot
		Exports all policies to HTML and pivot-friendly CSV format.

	.EXAMPLE
		PS> .\Export-CaPolicy.ps1 -PolicyID "12345678-1234-1234-1234-123456789012" -Csv -CsvPivot
		Exports a specific policy to HTML, normalized CSV, and pivot CSV formats.

	.NOTES
		Douglas Baker
		@dougsbaker

    CONTRIBUTORS
		Andres Bohren
    @andresbohren

    Output report uses open source components for HTML formatting
      - bootstrap - MIT License - https://getbootstrap.com/docs/4.0/about/license/
      - fontawesome - CC BY 4.0 License - https://fontawesome.com/license/free

    ############################################################################
    This sample script is not supported under any standard support program or service.
    This sample script is provided AS IS without warranty of any kind.
    This work is licensed under a Creative Commons Attribution 4.0 International License
    https://creativecommons.org/licenses/by-nc-sa/4.0/
    ############################################################################

#>

[CmdletBinding()]
# Suppress long line warnings for embedded HTML
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidLongLines', '')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAlignAssignmentStatement', '')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseConsistentIndentation', '')]
param (
  [Parameter()]
  [String]$TenantID,
  [Parameter()]
  [String]$PolicyID,
  [switch]$Csv,
  [switch]$CsvPivot
)

# Reference parameters to satisfy analyzer usage checks
$null = $TenantID

# UTILITY FUNCTIONS

function Write-Info {
  <#
  .SYNOPSIS
    Writes informational messages to the console
  .PARAMETER Message
    The message to display
  #>
  param([string]$Message)
  Write-Information -MessageData $Message -InformationAction Continue
}

function Write-Warn {
  <#
  .SYNOPSIS
    Writes warning messages to the console
  .PARAMETER Message
    The warning message to display
  #>
  param([string]$Message)
  Write-Warning $Message
}

function Write-Err {
  <#
  .SYNOPSIS
    Writes error messages to the console
  .PARAMETER Message
    The error message to display
  #>
  param([string]$Message)
  Write-Error -Message $Message
}

function Format-PolicyStatus {
  <#
  .SYNOPSIS
    Formats policy status for better readability in exports
  .PARAMETER Status
    The raw policy status from Microsoft Graph
  .OUTPUTS
    String - Formatted status for display
  #>
  param([string]$Status)
  switch ($Status) {
    'enabledForReportingButNotEnforced' { return 'reporting only' }
    default { return $Status }
  }
}

function Get-RecommendedPolicyName {
  <#
  .SYNOPSIS
    Generates a recommended policy name based on enterprise naming conventions
  .DESCRIPTION
    Creates a standardized policy name using the format: [Prefix]-[Scope]-[Condition]-[Control]-[ID]
    Based on the policy's configuration and intended purpose.
  .PARAMETER Policy
    The Conditional Access policy object from Microsoft Graph
  .OUTPUTS
    String - Recommended policy name following enterprise naming conventions
  .EXAMPLE
    Get-RecommendedPolicyName -Policy $caPolicy
    Returns: "CA-AllUsers-HighRisk-RequireMFA-001"
  #>
  param([Parameter(Mandatory)]$Policy)

  # Prefix - Standard CA prefix
  $prefix = "CA"

  # Scope - Determine target user scope
  $scope = "Unknown"
  $users = $Policy.conditions.users
  if ($users.includeUsers -contains "All") {
    if ($users.excludeUsers -and $users.excludeUsers.Count -gt 0) {
      $scope = "AllUsers-Exceptions"
    } else {
      $scope = "AllUsers"
    }
  } elseif ($users.includeUsers -contains "GuestsOrExternalUsers") {
    $scope = "Guests"
  } elseif ($users.includeRoles -and $users.includeRoles.Count -gt 0) {
    # Check for privileged admin roles
    $adminRoles = @('Global Administrator', 'Privileged Role Administrator', 'Security Administrator', 'Conditional Access Administrator')
    $hasAdminRole = $false
    foreach ($roleId in $users.includeRoles) {
      if ($RoleMap.ContainsKey($roleId) -and $RoleMap[$roleId] -in $adminRoles) {
        $hasAdminRole = $true
        break
      }
    }
    $scope = if ($hasAdminRole) { "PrivAdmins" } else { "Roles" }
  } elseif ($users.includeGroups -and $users.includeGroups.Count -gt 0) {
    $scope = "Groups"
  } elseif ($users.includeUsers -and $users.includeUsers.Count -gt 0) {
    $scope = "Users"
  }

  # Condition - Determine primary triggering condition
  $condition = @()
  $conditions = $Policy.conditions

  # Risk-based conditions (highest priority)
  if ($conditions.userRiskLevels -and $conditions.userRiskLevels.Count -gt 0) {
    if ($conditions.userRiskLevels -contains "high") {
      $condition += "HighUserRisk"
    } else {
      $condition += "UserRisk"
    }
  }
  if ($conditions.signInRiskLevels -and $conditions.signInRiskLevels.Count -gt 0) {
    if ($conditions.signInRiskLevels -contains "high") {
      $condition += "HighSignInRisk"
    } else {
      $condition += "SignInRisk"
    }
  }

  # Location-based conditions
  if ($conditions.locations.excludeLocations -and $conditions.locations.excludeLocations.Count -gt 0) {
    $condition += "UntrustedLocation"
  }
  if ($conditions.locations.includeLocations -and $conditions.locations.includeLocations.Count -gt 0) {
    $condition += "SpecificLocation"
  }

  # Client app conditions
  if ($conditions.clientAppTypes -contains "exchangeActiveSync" -or
      $conditions.clientAppTypes -contains "other") {
    $condition += "LegacyAuth"
  }

  # Authentication flows conditions
  if ($conditions.authenticationFlows -and $conditions.authenticationFlows.transferMethods -and $conditions.authenticationFlows.transferMethods.Count -gt 0) {
    $condition += "AuthFlows"
  }

  # Platform conditions
  if ($conditions.platforms.includePlatforms -and $conditions.platforms.includePlatforms.Count -gt 0) {
    $platforms = $conditions.platforms.includePlatforms
    if ($platforms -contains "android" -and $platforms -contains "iOS") {
      $condition += "MobileDevices"
    } elseif ($platforms.Count -eq 1) {
      $condition += $platforms[0].Substring(0,1).ToUpper() + $platforms[0].Substring(1)
    } else {
      $condition += "Platforms"
    }
  }

  # Device conditions
  if ($conditions.devices.includeDevices -and $conditions.devices.includeDevices.Count -gt 0) {
    if ($conditions.devices.includeDevices -contains "All") {
      $condition += "AllDevices"
    } else {
      $condition += "Devices"
    }
  }

  # Device filter conditions
  if ($conditions.devices.deviceFilter -and $conditions.devices.deviceFilter.rule) {
    $condition += "DeviceFilter"
  }

  # Application conditions
  if ($conditions.applications.includeApplications -and $conditions.applications.includeApplications.Count -gt 0) {
    $apps = $conditions.applications.includeApplications
    if ($apps -contains "All") {
      $condition += "AllApps"
    } elseif ($apps -contains "Office365") {
      $condition += "Office365"
    } else {
      $condition += "CloudApps"
    }
  }

  # Default condition if none detected
  if ($condition.Count -eq 0) {
    $condition += "General"
  }

  # Control - Determine primary enforcement action
  $control = "Unknown"
  $grantControls = $Policy.grantControls
  $sessionControls = $Policy.sessionControls

  if ($grantControls.builtInControls -contains "Block") {
    $control = "Block"
  } elseif ($grantControls.builtInControls -contains "Mfa") {
    if ($grantControls.authenticationStrength -and $grantControls.authenticationStrength.displayName) {
      $control = "RequireAuthStrength"
    } elseif ($grantControls.builtInControls.Count -eq 1) {
      $control = "RequireMFA"
    } else {
      $control = "RequireMFA-Plus"
    }
  } elseif ($grantControls.builtInControls -contains "CompliantDevice") {
    $control = "RequireCompliantDevice"
  } elseif ($grantControls.builtInControls -contains "DomainJoinedDevice") {
    $control = "RequireDomainJoined"
  } elseif ($grantControls.builtInControls -contains "ApprovedApplication") {
    $control = "RequireApprovedApp"
  } elseif ($grantControls.builtInControls -contains "CompliantApplication") {
    $control = "RequireCompliantApp"
  } elseif ($grantControls.builtInControls -contains "PasswordChange") {
    $control = "RequirePasswordChange"
  } elseif ($grantControls.termsOfUse -and $grantControls.termsOfUse.Count -gt 0) {
    $control = "RequireToU"
  } else {
    $control = "Grant"
  }

  # Add session control modifiers
  $sessionModifiers = @()
  if ($sessionControls.signInFrequency -and $sessionControls.signInFrequency.isEnabled) {
    $sessionModifiers += "SignInFreq"
  }
  if ($sessionControls.continuousAccessEvaluation -and $sessionControls.continuousAccessEvaluation.mode -eq "strictEnforcement") {
    $sessionModifiers += "CAE"
  }
  if ($sessionControls.persistentBrowser -and $sessionControls.persistentBrowser.isEnabled) {
    $sessionModifiers += "PersistentBrowser"
  }
  if ($sessionControls.applicationEnforcedRestrictions -and $sessionControls.applicationEnforcedRestrictions.isEnabled) {
    $sessionModifiers += "AppRestrictions"
  }
  if ($sessionControls.cloudAppSecurity -and $sessionControls.cloudAppSecurity.isEnabled) {
    $sessionModifiers += "CloudAppSec"
  }

  # Append session modifiers to control
  if ($sessionModifiers.Count -gt 0) {
    $control += "-" + ($sessionModifiers -join "-")
  }

  # Add status modifier if reporting only
  if ($Policy.state -eq "enabledForReportingButNotEnforced") {
    $control += "-ReportOnly"
  }

  # ID - Simple incremental ID (could be enhanced with actual policy counting)
  $id = "001"

  # Combine components - limit condition to first 3 for readability with enhanced naming
  $conditionStr = ($condition | Select-Object -First 3) -join "-"
  $recommendedName = "$prefix$id-$scope-$conditionStr-$control"

  # Ensure name doesn't exceed reasonable length (max 100 characters for enhanced naming)
  if ($recommendedName.Length -gt 100) {
    # Truncate condition part if too long
    $maxConditionLength = 100 - $prefix.Length - $id.Length - $scope.Length - $control.Length - 3 # 3 hyphens
    if ($conditionStr.Length -gt $maxConditionLength -and $maxConditionLength -gt 0) {
      $conditionStr = $conditionStr.Substring(0, $maxConditionLength)
    }
    $recommendedName = "$prefix$id-$scope-$conditionStr-$control"
  }

  return $recommendedName
}

function Convert-IdListToName {
  <#
  .SYNOPSIS
    Convert a list of IDs to their friendly names when present in a lookup map
  .DESCRIPTION
    For each element in List, if the Map contains that key the mapped value is output; otherwise the original value.
    Null / empty input yields an empty array.
  .PARAMETER List
    Collection of IDs or values
  .PARAMETER Map
    Hashtable keyed by ID with friendly values
  .EXAMPLE
    Convert-IdListToName -List $policy.conditions.users.includeUsers -Map $UserMap
  #>
  [CmdletBinding()]
  param([string[]]$List, [hashtable]$Map)
  if (-not $List) { return @() }
  return $List | ForEach-Object { if ($Map.ContainsKey($_)) { $Map[$_] } else { $_ } }
}

function Test-ModuleInstalled {
  <#
  .SYNOPSIS
    Checks if PowerShell modules are installed
  .PARAMETER ModuleNames
    Array of module names to check
  .OUTPUTS
    String[] - Array of missing module names
  #>
  param([string[]]$ModuleNames)
  $missing = @(); foreach ($m in $ModuleNames) { if (-not (Get-Module -ListAvailable -Name $m)) { $missing += $m } }; return $missing
}

# Helper function for safe Graph API calls
function Invoke-SafeGet {
  <#
  .SYNOPSIS
    Executes Graph API calls with error suppression for graceful failure handling
  .DESCRIPTION
    Wraps Graph API calls to prevent script termination when individual entities cannot be resolved.
    Common scenarios include deleted users, inaccessible applications, or permission limitations.
  .PARAMETER ScriptBlock
    The Graph API command to execute safely
  .OUTPUTS
    Object - The result of the API call, or $null if an error occurred
  .EXAMPLE
    $user = Invoke-SafeGet { Get-MgUser -UserId $userId }
  #>
  param([Parameter(Mandatory)][ScriptBlock]$ScriptBlock)
  try { & $ScriptBlock } catch { Write-Verbose ('Invoke-SafeGet suppressed error: {0}' -f $_.Exception.Message); return $null }
}

function Initialize-GraphModule {
  <#
.SYNOPSIS
  Ensure required PowerShell modules are installed and loadable for the current user.
.DESCRIPTION
  Verifies presence of Microsoft Graph modules used by this script and installs them to CurrentUser scope when missing.
  Attempts to trust PSGallery and install the NuGet provider when necessary. Imports modules after install.
.PARAMETER RequiredModules
  The list of module names to validate/install. Defaults to Microsoft Graph modules used by this script.
.EXAMPLE
  Ensure-GraphModules
#>
  [CmdletBinding()]
  param(
    [string[]]$RequiredModules = @(
      'Microsoft.Graph.Authentication',
      'Microsoft.Graph.Identity.DirectoryManagement',
      'Microsoft.Graph.Identity.SignIns'
    )
  )

  try {
    # Prefer TLS 1.2 for gallery operations (safe no-op on newer PowerShell)
    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
  }
  catch {
    Write-Verbose 'Failed to set SecurityProtocol to TLS 1.2; proceeding with defaults.'
  }

  # Ensure NuGet provider exists (for Install-Module)
  try {
    $nuget = Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue
    if (-not $nuget) {
      Write-Info 'Installing NuGet package provider (CurrentUser)'
      Install-PackageProvider -Name NuGet -Scope CurrentUser -Force -MinimumVersion '2.8.5.201' -ErrorAction Stop | Out-Null
    }
  }
  catch {
    Write-Warn ('Failed to install NuGet provider: {0}' -f $_.Exception.Message)
  }

  # Ensure PSGallery is available and trusted
  try {
    $repo = Get-PSRepository -Name 'PSGallery' -ErrorAction Stop
    if ($repo.InstallationPolicy -ne 'Trusted') {
      try { Set-PSRepository -Name 'PSGallery' -InstallationPolicy Trusted -ErrorAction Stop } catch { Write-Warn 'Could not set PSGallery as Trusted. You may be prompted during install.' }
    }
  }
  catch {
    Write-Warn 'PowerShell Gallery (PSGallery) not found. Module installation may fail until the repository is available.'
  }

  foreach ($m in $RequiredModules) {
    $installed = Get-Module -ListAvailable -Name $m
    if (-not $installed) {
      Write-Info ('Installing module: {0} (CurrentUser)' -f $m)
      try {
        Install-Module -Name $m -Scope CurrentUser -AllowClobber -Force -ErrorAction Stop
      }
      catch {
        Write-Warn ("Failed to install module '{0}': {1}" -f $m, $_.Exception.Message)
      }
    }
    # Import (best-effort)
    #try {
      #Import-Module -Name $m -Force -ErrorAction Stop
    #}
    #catch {
      #Write-Warn ("Failed to import module '{0}': {1}" -f $m, $_.Exception.Message)
    #}
  }
}

# Graph API connection and authentication functions

function Test-GraphConnected {
  <#
  .SYNOPSIS
    Tests if the current session is connected to Microsoft Graph
  .DESCRIPTION
    Verifies Graph connectivity by attempting to call Get-MgOrganization
  .OUTPUTS
    Boolean - True if connected, False if not connected
  #>
  try { Get-MgOrganization -ErrorAction Stop | Out-Null; return $true } catch { return $false }
}

function Get-CurrentGraphScope {
  <#
  .SYNOPSIS
    Gets the current Graph API scopes for the active connection
  .DESCRIPTION
    Retrieves the permission scopes granted to the current Graph session
  .OUTPUTS
    String[] - Array of scope names, empty array if not connected
  #>
  try { (Get-MgContext).Scopes } catch { @() }
}

function Connect-GraphContext {
  <#
  .SYNOPSIS
    Establishes connection to Microsoft Graph with required permissions
  .DESCRIPTION
    Connects to Microsoft Graph API with the specified scopes. Validates existing connection
    and only reconnects if missing required scopes or not currently connected.
  .PARAMETER RequiredScopes
    Array of permission scopes needed for the script to function properly
  .EXAMPLE
    Connect-GraphContext -RequiredScopes @('Policy.Read.All', 'Directory.Read.All')
  #>
  param([string[]]$RequiredScopes = @('Policy.Read.All', 'Directory.Read.All', 'Application.Read.All', 'Agreement.Read.All'))

  Initialize-GraphModule
  $connected = Test-GraphConnected
  $current = Get-CurrentGraphScope
  $still = @(); foreach ($s in $RequiredScopes) { if ($current -notcontains $s) { $still += $s } }

  if (-not $connected -or $still) {
    Write-Info 'Connecting to Microsoft Graph...'
    try {
      if ($TenantID) {
        Connect-MgGraph -Scopes $RequiredScopes -TenantId $TenantID -ErrorAction Stop | Out-Null
      }
      else {
        Connect-MgGraph -Scopes $RequiredScopes -ErrorAction Stop | Out-Null
      }
    }
    catch { Write-Err "Unable to connect to Microsoft Graph: $_"; throw }
  }

  if (-not (Test-GraphConnected)) { Write-Err 'Failed to connect to Microsoft Graph.'; exit 1 }
  $current = Get-CurrentGraphScope
  if ($still) { Write-Warn "Connected but missing scopes: $($still -join ', ')" }
  Write-Info "Connected scopes: $($current -join ', ')"
}

# MAIN SCRIPT EXECUTION

# Initialize variables and paths
$ExportLocation = $PSScriptRoot; if (!$ExportLocation) { $ExportLocation = $PWD }
$HTMLExport = $true
$CsvExport = $Csv
$CsvPivotExport = $CsvPivot

# Connect to Microsoft Graph with required permissions
Connect-GraphContext

# Get tenant information for report headers
$TenantData = Get-MgOrganization
$TenantName = $TenantData.DisplayName
$date = Get-Date
Write-Info "Connected: $TenantName tenant"

# Generate timestamped filename for outputs
$baseName = "CAExport_${TenantName}_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
$FileName = "$baseName.html"
$CsvFileName = "$baseName.csv"
$CsvPivotFileName = "$baseName-pivot.csv"

# CONDITIONAL ACCESS POLICY COLLECTION


Write-Info 'Exporting: CA Policy'
try {
  if ($PolicyID) {
    # Export specific policy by ID
    $CAPolicy = Get-MgIdentityConditionalAccessPolicy -ConditionalAccessPolicyId $PolicyID
    if (-not $CAPolicy) {
      Write-Err "Policy with ID '$PolicyID' not found. Exiting."
      exit 1
    }
  }
  else {
    # Export all policies in the tenant
    $CAPolicy = Get-MgIdentityConditionalAccessPolicy -All
  }
}
catch {
  Write-Err ('Failed to retrieve policies: {0}' -f $_.Exception.Message)
  Write-Err 'Cannot continue without policies. Exiting.'
  exit 1
}

if (-not $CAPolicy -or $CAPolicy.Count -eq 0) {
  Write-Err 'No Conditional Access policies found in tenant. Exiting.'
  exit 1
}

Write-Info "Successfully retrieved $($CAPolicy.Count) Conditional Access $(if($CAPolicy.Count -eq 1){'policy'}else{'policies'})"

$TenantData = Get-MgOrganization
$TenantName = $TenantData.DisplayName
$date = Get-Date

# ENTITY RESOLUTION - Convert GUIDs to Display Names


Write-Info "Extracting: Names from Guid's"

# Collect unique IDs from all policies for efficient bulk lookups
$userIds = [System.Collections.Generic.HashSet[string]]::new()
$groupIds = [System.Collections.Generic.HashSet[string]]::new()
$roleIds = [System.Collections.Generic.HashSet[string]]::new()

# Parse all policies to extract unique entity IDs
foreach ($p in $CAPolicy) {
  $c = $p.conditions
  if ($c.users) {
    # Extract user IDs (skip built-in identifiers like 'All', 'None')
    foreach ($i in @($c.users.includeUsers)) { if ($i -and $i -notin @('All', 'None', 'GuestsOrExternalUsers') -and ([Guid]::TryParse($i, [ref][Guid]::Empty))) { [void]$userIds.Add($i) } }
    foreach ($i in @($c.users.excludeUsers)) { if ($i -and ([Guid]::TryParse($i, [ref][Guid]::Empty))) { [void]$userIds.Add($i) } }

    # Extract group IDs
    foreach ($i in @($c.users.includeGroups)) { if ($i -and ([Guid]::TryParse($i, [ref][Guid]::Empty))) { [void]$groupIds.Add($i) } }
    foreach ($i in @($c.users.excludeGroups)) { if ($i -and ([Guid]::TryParse($i, [ref][Guid]::Empty))) { [void]$groupIds.Add($i) } }

    # Extract role IDs
    foreach ($i in @($c.users.includeRoles)) { if ($i -and ([Guid]::TryParse($i, [ref][Guid]::Empty))) { [void]$roleIds.Add($i) } }
    foreach ($i in @($c.users.excludeRoles)) { if ($i -and ([Guid]::TryParse($i, [ref][Guid]::Empty))) { [void]$roleIds.Add($i) } }
  }
}

# Build unified lookup hashtable for all entities
$mgObjectsLookup = @{}

# Create separate lookup maps for different entity types
$UserMap = @{}
$GroupMap = @{}
$RoleMap = @{}

# Resolve Users - Convert user GUIDs to display names
foreach ($id in $userIds) {
  $obj = Invoke-SafeGet { Get-MgUser -UserId $id -Property Id, DisplayName }
  if ($obj) {
    $mgObjectsLookup[$id] = $obj.DisplayName
    $UserMap[$id] = $obj.DisplayName
  }
}

# Resolve Groups - Convert group GUIDs to display names
foreach ($id in $groupIds) {
  $obj = Invoke-SafeGet { Get-MgGroup -GroupId $id -Property Id, DisplayName }
  if ($obj) {
    $mgObjectsLookup[$id] = $obj.DisplayName
    $GroupMap[$id] = $obj.DisplayName
  }
}

# ================================================================================================
# APPLICATION RESOLUTION
# ================================================================================================

# Collect all application IDs referenced in policies
$AppIds = @()
foreach ($policy in $CAPolicy) {
  if ($policy.Conditions -and $policy.Conditions.Applications) {
    $AppIds += $policy.Conditions.Applications.IncludeApplications
    $AppIds += $policy.Conditions.Applications.ExcludeApplications
  }
}
# Filter to valid GUIDs and remove duplicates
$AppIds = $AppIds | Where-Object { $_ -and ([Guid]::TryParse($_, [ref][Guid]::Empty)) } | Sort-Object -Unique

# Build application display name lookup table
$MGAppsLookup = @{}
foreach ($AppId in $AppIds) {
  if ([Guid]::TryParse($AppId, [ref][Guid]::Empty)) {
    $obj = Invoke-SafeGet { Get-MgServicePrincipal -ServicePrincipalId $AppId -Property Id, DisplayName, AppId }
    if ($obj) { $MGAppsLookup[$AppId] = $obj.DisplayName }
  }
}


# UPDATE POLICIES WITH RESOLVED NAMES


# Replace application GUIDs with display names in policy objects
#"<div class='tooltip-container'>" + $obj.DisplayName +"<span class='tooltip-text'>App Id:"+ $obj.AppId +"</span></div>"

foreach ($policy in $CAPolicy) {
  if (-not $policy.Conditions -or -not $policy.Conditions.Applications) { continue }

  # Update excluded applications
  for ($i = 0; $i -lt $policy.Conditions.Applications.ExcludeApplications.Count; $i++) {
    $AppId = $policy.Conditions.Applications.ExcludeApplications[$i]
    if ($MGAppsLookup.ContainsKey($AppId)) {
      $policy.Conditions.Applications.ExcludeApplications[$i] = $MGAppsLookup[$AppId]
    }
  }

  # Update included applications
  for ($i = 0; $i -lt $policy.Conditions.Applications.IncludeApplications.Count; $i++) {
    $AppId = $policy.Conditions.Applications.IncludeApplications[$i]
    if ($MGAppsLookup.ContainsKey($AppId)) {
      $policy.Conditions.Applications.IncludeApplications[$i] = $MGAppsLookup[$AppId]
    }
  }
}



# NAMED LOCATIONS RESOLUTION

# Get all named locations and build lookup table
$mgLoc = Get-MgIdentityConditionalAccessNamedLocation
$MGLocLookup = @{}
foreach ($obj in $mgLoc) {
  $MGLocLookup[$obj.Id] = $obj.DisplayName
}

# Replace location GUIDs with display names in policy objects
foreach ($policy in $CAPolicy) {
  if (-not $policy.Conditions -or -not $policy.Conditions.Locations) { continue }

  # Update included locations
  for ($i = 0; $i -lt $policy.Conditions.Locations.IncludeLocations.Count; $i++) {
    $LocId = $policy.Conditions.Locations.IncludeLocations[$i]
    if ($MGLocLookup.ContainsKey($LocId)) {
      $policy.Conditions.Locations.IncludeLocations[$i] = $MGLocLookup[$LocId]
    }
  }

  # Update excluded locations
  for ($i = 0; $i -lt $policy.Conditions.Locations.ExcludeLocations.Count; $i++) {
    $LocId = $policy.Conditions.Locations.ExcludeLocations[$i]
    if ($MGLocLookup.ContainsKey($LocId)) {
      $policy.Conditions.Locations.ExcludeLocations[$i] = $MGLocLookup[$LocId]
    }
  }
}


# TERMS OF USE RESOLUTION

# Get all terms of use agreements and build lookup table
$mgTou = Get-MgAgreement
$MGTouLookup = @{}
foreach ($obj in $mgTou) {
  $MGTouLookup[$obj.Id] = $obj.DisplayName
}

# Replace terms of use GUIDs with display names in policy grant controls
foreach ($policy in $CAPolicy) {
  if (-not $policy.GrantControls -or -not $policy.GrantControls.TermsOfUse) { continue }

  for ($i = 0; $i -lt $policy.GrantControls.TermsOfUse.Count; $i++) {
    $TouId = $policy.GrantControls.TermsOfUse[$i]
    if ($MGTouLookup.ContainsKey($TouId)) {
      $policy.GrantControls.TermsOfUse[$i] = $MGTouLookup[$TouId]
    }
  }
}


# ROLE RESOLUTION - Administrative Roles

# Build comprehensive role lookup table from multiple sources
$roleLookup = @{}

# Static fallback for common Azure AD role template IDs (used if API calls fail)
$commonRoleTemplates = @{
  '62e90394-69f5-4237-9190-012177145e10' = 'Global Administrator'
  'f28a1f50-f6e7-4571-818b-6a12f2af6b6c' = 'SharePoint Administrator'
  '29232cdf-9323-42fd-ade2-1d097af3e4de' = 'Exchange Administrator'
  'b1be1c3e-b65d-4f19-8427-f6fa0d97feb9' = 'Conditional Access Administrator'
  '729827e3-9c14-49f7-bb1b-9608f156bbb8' = 'Helpdesk Administrator'
  'b0f54661-2d74-4c50-afa3-1ec803f12efe' = 'Billing Administrator'
  'fe930be7-5e62-47db-91af-98c3a49a38b1' = 'User Administrator'
  'c4e39bd9-1100-46d3-8c65-fb160da0071f' = 'Authentication Administrator'
  '9b895d92-2cd3-44c7-9d02-a6ac2d5ea5c3' = 'Application Administrator'
  '158c047a-c907-4556-b7ef-446551a6b5f7' = 'Cloud Application Administrator'
  '966707d0-3269-4727-9be2-8c3a10f19b9d' = 'Password Administrator'
  '7be44c8a-adaf-4e2a-84d6-ab2649e08a13' = 'Privileged Authentication Administrator'
  'e8611ab8-c189-46e8-94e1-60213ab1f814' = 'Privileged Role Administrator'
  '194ae4cb-b126-40b2-bd5b-6091b380977d' = 'Security Administrator'
  '5d6b6bb7-de71-4623-b4af-96380a352509' = 'Security Reader'
  'e3973bdf-4987-49ae-837a-ba8e231c7286' = 'Azure DevOps Administrator'
}

# Get active directory roles (currently activated roles in the tenant)
$dirRoles = Invoke-SafeGet { Get-MgDirectoryRole -All }
if ($dirRoles) {
  Write-Verbose "Successfully retrieved $($dirRoles.Count) directory roles from API"
  foreach ($r in $dirRoles) {
    # Store both the instance ID and template ID for flexible lookup
    if ($r.Id -and -not $roleLookup.ContainsKey($r.Id)) { $roleLookup[$r.Id] = $r.DisplayName }
    if ($r.RoleTemplateId -and -not $roleLookup.ContainsKey($r.RoleTemplateId)) { $roleLookup[$r.RoleTemplateId] = $r.DisplayName }
  }
}
else {
  Write-Warn 'Get-MgDirectoryRole failed; using static role template fallback for common roles'
  # Populate lookup with static template mappings as fallback
  foreach ($kv in $commonRoleTemplates.GetEnumerator()) {
    $roleLookup[$kv.Key] = $kv.Value
  }
}

# Get role definitions for comprehensive coverage (includes inactive roles)
$roleDefs = Invoke-SafeGet { Get-MgRoleManagementDirectoryRoleDefinition -All -Property Id, DisplayName, TemplateId }
if ($roleDefs) {
  Write-Verbose "Successfully retrieved $($roleDefs.Count) role definitions from API"
  foreach ($rd in $roleDefs) {
    # Store both definition ID and template ID
    if ($rd.Id -and -not $roleLookup.ContainsKey($rd.Id)) { $roleLookup[$rd.Id] = $rd.DisplayName }
    if ($rd.TemplateId -and -not $roleLookup.ContainsKey($rd.TemplateId)) { $roleLookup[$rd.TemplateId] = $rd.DisplayName }
  }
}
else {
  Write-Warn 'Get-MgRoleManagementDirectoryRoleDefinition failed; relying on directory roles and static fallback'
  # If we don't have role definitions but missed common templates, add them
  foreach ($kv in $commonRoleTemplates.GetEnumerator()) {
    if (-not $roleLookup.ContainsKey($kv.Key)) {
      $roleLookup[$kv.Key] = $kv.Value
    }
  }
}

# Process role IDs found in policies and add to unified lookup
foreach ($id in $roleIds) {
  if ($roleLookup.ContainsKey($id)) {
    $mgObjectsLookup[$id] = $roleLookup[$id]
    $RoleMap[$id] = $roleLookup[$id]
  }
  else {
    # Fallback attempt for roles that might have just become active or API lookup failed
    $obj = Invoke-SafeGet { Get-MgDirectoryRole -DirectoryRoleId $id -Property Id, DisplayName }
    if ($obj) {
      $mgObjectsLookup[$id] = $obj.DisplayName
      $RoleMap[$id] = $obj.DisplayName
    }
    else {
      # Check if this ID matches a known template ID from our static fallback
      if ($commonRoleTemplates.ContainsKey($id)) {
        $mgObjectsLookup[$id] = $commonRoleTemplates[$id]
        $RoleMap[$id] = $commonRoleTemplates[$id]
        Write-Verbose "Resolved role ID $id using static template fallback: $($commonRoleTemplates[$id])"
      }
      else {
        Write-Verbose "Unresolved role id: $id (not found in API or static fallback)"
      }
    }
  }
}

# Replace role GUIDs with display names in policy conditions
foreach ($policy in $caPolicy) {
  if (-not $policy.Conditions -or -not $policy.Conditions.Users) { continue }

  # Update included roles
  if ($policy.Conditions.Users.IncludeRoles) {
    for ($i = 0; $i -lt $policy.Conditions.Users.IncludeRoles.Count; $i++) {
      $RoleId = $policy.Conditions.Users.IncludeRoles[$i]
      if ($mgObjectsLookup.ContainsKey($RoleId)) {
        $policy.Conditions.Users.IncludeRoles[$i] = $mgObjectsLookup[$RoleId]
      }
    }
  }

  # Update excluded roles
  if ($policy.Conditions.Users.ExcludeRoles) {
    for ($i = 0; $i -lt $policy.Conditions.Users.ExcludeRoles.Count; $i++) {
      $RoleId = $policy.Conditions.Users.ExcludeRoles[$i]
      if ($mgObjectsLookup.ContainsKey($RoleId)) {
        $policy.Conditions.Users.ExcludeRoles[$i] = $mgObjectsLookup[$RoleId]
      }
    }
  }
}


# POLICY DATA EXTRACTION AND TRANSFORMATION

# Initialize array to hold processed policy data
$CAExport = @()

Write-Info 'Extracting: CA Policy Data'

# Process each policy and extract key information into structured format
foreach ( $Policy in $CAPolicy) {

  # Combine all user/group/role assignments for easier viewing
  $IncludeUG = @()
  $IncludeUG += (Convert-IdListToName $Policy.Conditions.Users.IncludeUsers $UserMap)
  $IncludeUG += (Convert-IdListToName $Policy.Conditions.Users.IncludeGroups $GroupMap)
  $IncludeUG += (Convert-IdListToName $Policy.Conditions.Users.IncludeRoles $RoleMap)

  # Extract creation and modification timestamps
  $DateCreated = $null
  $DateCreated = $Policy.CreatedDateTime
  $DateModified = $null
  $DateModified = $Policy.ModifiedDateTime

  # Combine excluded user/group/role assignments
  $ExcludeUG = @()
  $ExcludeUG += (Convert-IdListToName $Policy.Conditions.Users.ExcludeUsers $UserMap)
  $ExcludeUG += (Convert-IdListToName $Policy.Conditions.Users.ExcludeGroups $GroupMap)
  $ExcludeUG += (Convert-IdListToName $Policy.Conditions.Users.ExcludeRoles $RoleMap)

  # Collect application references (Note: $Apps variable appears unused but preserving for compatibility)
  $Apps += $Policy.Conditions.Applications.IncludeApplications
  $Apps += $Policy.Conditions.Applications.ExcludeApplications

  # Extract location conditions
  $InclLocation = $Null
  $ExclLocation = $Null
  $InclLocation = $Policy.Conditions.Locations.includelocations
  $ExclLocation = $Policy.Conditions.Locations.Excludelocations

  # Extract platform conditions
  $InclPlat = $Null
  $ExclPlat = $Null
  $InclPlat = $Policy.Conditions.Platforms.IncludePlatforms
  $ExclPlat = $Policy.Conditions.Platforms.ExcludePlatforms

  # Extract device conditions
  $InclDev = $null
  $ExclDev = $null
  $InclDev = $Policy.Conditions.Devices.IncludeDevices
  $ExclDev = $Policy.Conditions.Devices.ExcludeDevices
  $devFilters = $null
  $devFilters = $Policy.Conditions.Devices.DeviceFilter.Rule

  # Create structured object with all policy information for export
  $CAExport += [PSCustomObject][ordered]@{
    # Basic policy information
    Name = $Policy.DisplayName
    Status = Format-PolicyStatus -Status $Policy.State
    'Recommended Name' = Get-RecommendedPolicyName -Policy $Policy
    Created = $DateCreated
    Modified = $DateModified

    # User and group assignments
    'Included Users' = ($IncludeUG -join ", `r`n")
    'Excluded Users' = ($ExcludeUG -join ", `r`n")

    # Application and cloud service conditions
    'Cloud apps or actions' = ''  # Section header for visual grouping
    'Included Applications' = ($Policy.Conditions.Applications.IncludeApplications -join ", `r`n")
    'Excluded Applications' = ($Policy.Conditions.Applications.ExcludeApplications -join ", `r`n")
    'User Actions' = ($Policy.Conditions.Applications.IncludeUserActions -join ", `r`n")
    'Auth Context' = ($Policy.Conditions.Applications.IncludeAuthenticationContextClassReferences -join ", `r`n")

    # Risk and condition assessments
    Conditions = ''  # Section header for visual grouping
    'User Risk' = ($Policy.Conditions.UserRiskLevels -join ", `r`n")
    'Sign In Risk' = ($Policy.Conditions.SignInRiskLevels -join ", `r`n")

    # Platform conditions (iOS, Android, Windows, etc.)
    'Included Platforms ' = ($InclPlat -join ", `r`n")
    'Excluded Platforms ' = ($ExclPlat -join ", `r`n")

    # Network location conditions
    'Included Locations' = ($InclLocation -join ", `r`n")
    'Excluded Locations' = ($ExclLocation -join ", `r`n")

    # Client application types (modern auth, legacy auth, etc.)
    'Client Apps' = ($Policy.Conditions.ClientAppTypes -join ", `r`n")

    # Device conditions and filters
    'Included Devices' = ($InclDev -join ", `r`n")
    'Excluded Devices' = ($ExclDev -join ", `r`n")
    'Device Filters' = ($devFilters -join ", `r`n")

    # Access control section headers
    'Access Controls' = ''  # Section header for visual grouping
    'Grant Controls' = ''   # Section header for visual grouping

    # Individual grant control requirements (expanded for better visibility)
    Block = if ($Policy.GrantControls.BuiltInControls -contains 'Block') { 'True' } else { '' }
    'Require MFA' = if ($Policy.GrantControls.BuiltInControls -contains 'Mfa') { 'True' } else { '' }
    'Authentication Strength MFA' = $Policy.GrantControls.AuthenticationStrength.DisplayName
    'Compliant Device' = if ($Policy.GrantControls.BuiltInControls -contains 'CompliantDevice') { 'True' } else { '' }
    'Domain Joined Device' = if ($Policy.GrantControls.BuiltInControls -contains 'DomainJoinedDevice') { 'True' } else { '' }
    'Compliant Application' = if ($Policy.GrantControls.BuiltInControls -contains 'CompliantApplication') { 'True' } else { '' }
    'Approved Application' = if ($Policy.GrantControls.BuiltInControls -contains 'ApprovedApplication') { 'True' } else { '' }
    'Password Change' = if ($Policy.GrantControls.BuiltInControls -contains 'PasswordChange') { 'True' } else { '' }
    'Terms Of Use' = ($Policy.GrantControls.TermsOfUse -join ", `r`n")
    'Custom Controls' = ($Policy.GrantControls.CustomAuthenticationFactors -join ", `r`n")
    GrantOperator = $Policy.GrantControls.Operator

    # Session control settings
    'Session Controls' = ''  # Section header for visual grouping
    'Application Enforced Restrictions' = $Policy.SessionControls.ApplicationEnforcedRestrictions.IsEnabled
    'Cloud App Security' = $Policy.SessionControls.CloudAppSecurity.IsEnabled
    'Sign In Frequency' = "$($Policy.SessionControls.SignInFrequency.Value) $($Policy.SessionControls.SignInFrequency.Type)"
    'Persistent Browser' = $Policy.SessionControls.PersistentBrowser.Mode
    'Continuous Access Evaluation' = $Policy.SessionControls.ContinuousAccessEvaluation.Mode
    'Resilient Defaults' = $policy.SessionControls.DisableResilienceDefaults
    'Secure Sign In Session' = $policy.SessionControls.AdditionalProperties.secureSignInSession.Values
  }
}

# DATA TRANSFORMATION FOR EXPORT FORMATS

Write-Info 'Pivoting: CA to Export Format'

# Transform data into pivot table format for better visualization
# This creates a transposed view where each policy becomes a column and each property becomes a row
$pivot = @()

# Create header row to establish the column structure
$rowItem = [PSCustomObject]@{}
$rowItem | Add-Member -Type NoteProperty -Name 'CA Item' -Value 'row1'
$pcount = 1
foreach ($ca in $CAExport) {
  $rowItem | Add-Member -Type NoteProperty -Name "Policy $pcount" -Value 'row1'
  $pcount += 1
}
$pivot += $rowItem

# Determine all properties from the first policy object for consistent structure
$properties = @()
if ($CAExport -and $CAExport.Count -gt 0) {
  $properties = ($CAExport | Select-Object -First 1 | Get-Member -MemberType NoteProperty).Name
}

# Create a row for each property across all policies
foreach ($prop in $properties) {
  $rowItem = [PSCustomObject]@{}
  $rowItem | Add-Member -Type NoteProperty -Name 'CA Item' -Value $prop
  $pcount = 1
  foreach ($ca in $CAExport) {
    $value = $null
    try { $value = $ca.$prop } catch { $value = $null }
    $rowItem | Add-Member -Type NoteProperty -Name "Policy $pcount" -Value $value
    $pcount += 1
  }
  $pivot += $rowItem
}

# Define custom sort order for logical grouping of policy elements in output
$sort = 'Name', 'Recommended Name', 'Status', 'Created', 'Modified', 'Included Users', 'Excluded Users', 'Cloud apps or actions', 'Included Applications', 'Excluded Applications', 'User Actions', 'Auth Context', 'Conditions', 'User Risk', 'Sign In Risk', 'Included Platforms ', 'Excluded Platforms ', 'Client Apps', 'Included Locations', 'Excluded Locations', 'Devices', 'Included Devices', 'Excluded Devices', 'Device Filters', 'Access Controls', 'Grant Controls', 'Block', 'Require MFA', 'Authentication Strength MFA', 'Compliant Device', 'Domain Joined Device', 'Compliant Application', 'Approved Application', 'Password Change', 'Terms Of Use', 'Custom Controls', 'GrantOperator', 'Session Controls', 'Application Enforced Restrictions', 'Cloud App Security', 'Sign In Frequency', 'Persistent Browser', 'Continuous Access Evaluation', 'Resilient Defaults', 'Secure Sign In Session'

# HTML EXPORT GENERATION

if ($HTMLExport) {
  Write-Info 'Saving to File: HTML'

  # jQuery and JavaScript for interactive table features (row/column selection)
  $jquery = '  <script src="https://ajax.googleapis.com/ajax/libs/jquery/3.5.1/jquery.min.js"></script>
    <script>
    $(document).ready(function(){
        // Row selection functionality
        $("tr").click(function(){
            if(!$(this).hasClass("selected")){
                $(this).addClass("selected");
            } else {
                $(this).removeClass("selected");
            }
        });

        // Column selection functionality
        $("th").click(function(){
            // Get the index of the clicked column
            var colIndex = $(this).index();
            // Select the corresponding col element and add or remove the class
            $("colgroup col").eq(colIndex).toggleClass("colselected");
        });
    });
    </script>'

  # Complete HTML template with Bootstrap styling and custom CSS
  $htmlContent = "<html><head><base href='https://docs.microsoft.com/' target='_blank'>
    <meta charset='utf-8'>
        <meta name='viewport' content='width=device-width, initial-scale=1, shrink-to-fit=no'>

                  <!-- External CSS and JavaScript libraries -->
                  <link rel='stylesheet' href='https://cdnjs.cloudflare.com/ajax/libs/font-awesome/5.11.2/css/all.min.css' crossorigin='anonymous'>
                  <link rel='stylesheet' href='https://stackpath.bootstrapcdn.com/bootstrap/4.3.1/css/bootstrap.min.css' integrity='sha384-ggOyR0iXCbMQv3Xipma34MD+dH/1fQ784/j6cY/iJTQUOhcWr7x9JvoRxT2MZw1T' crossorigin='anonymous'>
                  <script src='https://code.jquery.com/jquery-3.3.1.slim.min.js' integrity='sha384-q8i/X+965DzO0rT7abK41JStQIAqVgRVzpbzo5smXKp4YfRvH+8abtTE1Pi6jizo' crossorigin='anonymous'></script>
                  <script src='https://cdnjs.cloudflare.com/ajax/libs/popper.js/1.14.7/umd/popper.min.js' integrity='sha384-UO2eT0CpHqdSJQ6hJty5KVphtPhzWj9WO1clHTMGa3JDZwrnQq4sF86dIHNDz0W1' crossorigin='anonymous'></script>
                  <script src='https://stackpath.bootstrapcdn.com/bootstrap/4.3.1/js/bootstrap.min.js' integrity='sha384-JjSmVgyd0p3pXB1rRibZUAYoIIy6OrQ6VrjIEaFf/nJGzIxFDsf4x0xIM+B07jRM' crossorigin='anonymous'></script>
                  <script src='https://cdnjs.cloudflare.com/ajax/libs/font-awesome/5.11.2/js/all.js'></script>
                $jquery<style>
                /* Custom CSS styling for the CA export report */
                .title {
                    font-size: 1.5em;
                    font-weight: bold;
                    font-family: Arial, sans-serif;
                    top: 0;
                    right: 0;
                    left: 0;
                }

                /* Table styling for professional appearance */
                table {
                    border-collapse: collapse;
                    margin-bottom: 30px;
                    margin-top: 55px;
                    font-size: 0.9em;
                    font-family: Arial, sans-serif;
                    min-width: 400px;

                }
                  /* Header row styling */
                  thead tr {
                      background-color: #009879;
                      color: #ffffff;
                      text-align: center;
                 }
                  th, td {
                      min-width: 250px;
                      padding: 12px 15px;
                      border: 1px solid lightgray;
                      vertical-align: top;
                      text-align: center;
                 }

                  td {
                      vertical-align: top;
                 }
                  tbody tr {
                     /* border-bottom: 1px solid #dddddd;*/
                 }
                  tbody tr:nth-of-type(even) {
                      background-color: #f3f3f3;
                 }

                  tbody tr:last-of-type {
                      border-bottom: 2px solid #009879;
                 }
                 tr:hover {
                    background-color: #d8d8d8!important;
                }

              .selected:not(th){
                  background-color:#eaf7ff!important;

                  }
                  th{
                     background-color:white ;
                  }
                  .colselected {

                      width: 10%; border: 5px solid #59c7fb;

                }
                table tr th:first-child,table tr td:first-child {
                      position: sticky;
                      inset-inline-start: 0;
                      background-color: #005494;
                      border: 0px;
                      Color: #fff;
                      font-weight: bolder;
                      text-align: center;
                 }
                 tbody tr:nth-of-type(even) td:first-child  {
                      background-color: #547c9b;
                 }
                  tbody tr:nth-of-type(5),
                  tbody tr:nth-of-type(8),
                  tbody tr:nth-of-type(13),
                  tbody tr:nth-of-type(24),
                  tbody tr:nth-of-type(36) {
                  background-color: #005494!important;
                  }
                 .navbar-custom {
                    background-color: #005494;
                    color: white;
                    padding-bottom: 10px;

                }
                /* Modify brand and text color */

                .navbar-custom .navbar-brand,
                .navbar-custom .navbar-text {
                    color: white;
                    padding-top: 70px;
                    padding-bottom: 10px;
                }
                       /* Tooltip container */
        .tooltip-container {
            position: relative;
            display: inline-block;
        }

        /* Tooltip text */
        .tooltip-text {
            visibility: hidden;
            width: 200px;
            background-color: black;
            color: #fff;
            text-align: center;
            border-radius: 6px;
            padding: 5px 0;
            position: absolute;
            z-index: 1;
            top: 115%; /* Position the tooltip below the text */
            left: 50%;
            margin-left: -100px;
            opacity: 0;
            transition: opacity 0.3s;
        }

        .tooltip-container:hover .tooltip-text {
            visibility: visible;
            opacity: 1;
        }
                </style></head><body> <nav class='navbar  fixed-top navbar-custom p-3 border-bottom'>
                <div class='container-fluid'>
                    <div class='col-sm' style='text-align:left'>
                        <div class='row'><div><i class='fa fa-server' aria-hidden='true'></i></div><div class='ml-3'><strong>CA Export</strong></div></div>
                    </div>
                    <div class='col-sm' style='text-align:center'>
                        <strong>$TenantName</strong>
                    </div>
                    <div class='col-sm' style='text-align:right'>
                    <strong>$Date</strong>
                    </div>
                </div>
            </nav> "

  # Generate the final HTML report
  Write-Info 'Launching: Web Browser'
  $Launch = Join-Path -Path $ExportLocation -ChildPath $FileName

  # Convert data to HTML table, excluding header row and applying custom sort order
  $table = $pivot | Where-Object { $_.'CA Item' -ne 'row1' } | Sort-Object { $sort.IndexOf($_.'CA Item') } | ConvertTo-Html -Fragment
  $htmlContent = $htmlContent + $table

  # Decode HTML entities and write to file
  Add-Type -AssemblyName System.Web
  [System.Web.HttpUtility]::HtmlDecode($htmlContent) | Out-File $Launch

  # Open the generated HTML report in default browser
  Start-Process $Launch
}

# CSV EXPORT GENERATION (Optional)

if ($CsvExport) {
  Write-Info 'Saving data to CSV File'
  $LaunchCsv = Join-Path -Path $ExportLocation -ChildPath $CsvFileName

  # Export flat/normalized data to CSV format for Excel analysis
  $CAExport | Export-Csv -Path $LaunchCsv -NoTypeInformation -Encoding UTF8
  Write-Info "CSV exported to: $LaunchCsv"
}

# CSV PIVOT EXPORT GENERATION (Optional)

if ($CsvPivotExport) {
  Write-Info 'Saving data to Pivot CSV File'
  $LaunchCsvPivot = Join-Path -Path $ExportLocation -ChildPath $CsvPivotFileName

  # Export pivot data to CSV format for Excel analysis (transposed format)
  $csvData = $pivot | Where-Object { $_.'CA Item' -ne 'row1' } | Sort-Object { $sort.IndexOf($_.'CA Item') }
  $csvData | Export-Csv -Path $LaunchCsvPivot -NoTypeInformation -Encoding UTF8
  Write-Info "Pivot CSV exported to: $LaunchCsvPivot"
}