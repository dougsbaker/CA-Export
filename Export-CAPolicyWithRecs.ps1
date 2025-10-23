<#
.SYNOPSIS
    Export existing Conditional Access Policies along with recommendations.

.DESCRIPTION
    This script exports Conditional Access (CA) policies from Azure AD to an HTML file.
    It includes recommendations and checks for each policy to enhance security.

.PARAMETER PolicyID
  (Optional) A specific Conditional Access policy Id (GUID). When supplied the export/report is limited
  to this single policy. When omitted all policies are processed.

.PARAMETER Html
  Switch. Generate the interactive HTML report (default if no other export switch is specified).

.PARAMETER Json
  Switch. Generate a JSON file containing the enriched policy objects.

.PARAMETER Csv
  Switch. Generate a flattened CSV export of the policy objects suitable for spreadsheet review.

.PARAMETER CsvPivot
  Switch. Generate a pivot‑friendly CSV (wide format) for ad‑hoc aggregation in Excel / BI tools.

.PARAMETER NoRecommendations
  Switch. Skip generation of the recommendations analysis and omit the Recommendations tab from the HTML output. Useful for faster exports when only raw policy data is required.

.PARAMETER CsvColumns
  Optional string array specifying a custom subset / order of columns for the Csv export. When omitted
  the default full set is used.

.OUTPUTS
  Files written to the working directory. Filenames are timestamped and prefixed with CAExportRec_<TenantName>_.

.EXAMPLE
  PS> .\Export-CAPolicyWithRecs.ps1
  Exports all Conditional Access policies and produces an HTML report (default output).

.EXAMPLE
  PS> .\Export-CAPolicyWithRecs.ps1 -PolicyID 11111111-2222-3333-4444-555555555555 -Json
  Exports only the specified policy and produces a JSON file with enriched data.

.EXAMPLE
  PS> .\Export-CAPolicyWithRecs.ps1 -Csv -CsvColumns Name,Status,'Require MFA','Block'
  Produces a CSV limited to the selected columns in the specified order.

.EXAMPLE
  PS> .\Export-CAPolicyWithRecs.ps1 -Html -Json -Csv -CsvPivot
  Produces all supported output formats in a single invocation.

.EXAMPLE
  PS> .\Export-CAPolicyWithRecs.ps1 -NoRecommendations -Csv
  Exports policies (CSV + default HTML/JSON if no other switches) while skipping recommendation analysis for faster runtime.

.EXAMPLE
  PS> .\Export-CAPolicyWithRecs.ps1 -PolicyID (Get-Clipboard) -Html
  Uses a policy Id copied to the clipboard and generates only the HTML report.

.EXAMPLE
.\Export-CAPolicyWithRecs.ps1

This example runs the script and exports all Conditional Access policies with recommendations.

.NOTES
  Author:  Douglas Baker @dougsbaker
  Version: 3.3

############################################################################
This sample script is not supported under any standard support program or service.
This sample script is provided AS IS without warranty of any kind.
This work is licensed under a Creative Commons Attribution 4.0 International License
https://creativecommons.org/licenses/by-nc-sa/4.0/
############################################################################

#>

[CmdletBinding()]
# Suppress PSAvoidLongLines for unavoidable embedded HTML/URLs
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidLongLines', '')]
param (
  [Parameter()]
  [ValidatePattern('^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$')]
  [String]$PolicyID,
  [Parameter(HelpMessage = 'Path to previously exported *_raw.json to operate offline (skips Graph calls).')]
  [ValidateScript({ Test-Path $_ -PathType Leaf })]
  [string]$RawInputFile,
  [switch]$Html,
  [switch]$NoBrowser,
  [Parameter()][ValidateScript({ Test-Path $_ -PathType Container })][string]$OutputPath = '.',
  [switch]$Json,
  [switch]$Csv,
  [switch]$CsvPivot,
  [switch]$NoRecommendations,
  [string[]]$CsvColumns
)

function Write-Info {
  <#
.SYNOPSIS
  Write an informational message to the information stream.
.DESCRIPTION
  Wrapper around Write-Information so that callers can use -InformationAction / -InformationPreference
  and optionally suppress or capture messages. Replaces earlier Write-Host usage for lint compliance.
.PARAMETER Message
  The text to emit.
#>
  [CmdletBinding()]
  param([Parameter(Mandatory)][string]$Message)
  Write-Information -MessageData "[INFO] $Message" -InformationAction Continue
}


function Write-Warn {
  <#
.SYNOPSIS
  Write a warning message.
.DESCRIPTION
  Thin wrapper kept for symmetry with Write-Info.
.EXAMPLE
  Write-Warn 'Policy collection returned zero results.'
  Emits a formatted warning to the warning stream.
#>
  [CmdletBinding()]
  param([Parameter(Mandatory)][string]$Message)
  Write-Warning $Message
}


function Write-Err {
  <#
.SYNOPSIS
  Write an error message and optionally terminate the script.
.DESCRIPTION
  Wrapper around Write-Error for consistent error handling throughout the script.
.PARAMETER Message
  The error message to display.
.EXAMPLE
  Write-Err 'Critical failure occurred. Exiting.'
  Emits a formatted error message to the error stream.
#>
  [CmdletBinding()]
  param([Parameter(Mandatory)][string]$Message)
  Write-Error -Message $Message -ErrorAction Continue
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
    'enabledForReportingButNotEnforced' {
      return 'report only' 
    }
    default {
      return $Status 
    }
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
    Returns: "CA001-AllUsers-HighRisk-RequireMFA"
  #>
  param([Parameter(Mandatory)]$Policy)

  # Prefix - Standard CA prefix
  $prefix = 'CA'

  # Scope - Determine target user scope
  $scope = 'Unknown'
  $users = $Policy.conditions.users
  if ($users.includeUsers -contains 'All') {
    if ($users.excludeUsers -and $users.excludeUsers.Count -gt 0) {
      $scope = 'AllUsers-Exceptions'
    }
    else {
      $scope = 'AllUsers'
    }
  }
  elseif ($users.includeUsers -contains 'GuestsOrExternalUsers') {
    $scope = 'Guests'
  }
  elseif ($users.includeRoles -and $users.includeRoles.Count -gt 0) {
    # Check for privileged admin roles
    $adminRoles = @('Global Administrator', 'Privileged Role Administrator', 'Security Administrator', 'Conditional Access Administrator')
    $hasAdminRole = $false
    foreach ($roleId in $users.includeRoles) {
      if ($RoleMap.ContainsKey($roleId) -and $RoleMap[$roleId] -in $adminRoles) {
        $hasAdminRole = $true
        break
      }
    }
    $scope = if ($hasAdminRole) {
      'PrivAdmins' 
    }
    else {
      'Roles' 
    }
  }
  elseif ($users.includeGroups -and $users.includeGroups.Count -gt 0) {
    $scope = 'Groups'
  }
  elseif ($users.includeUsers -and $users.includeUsers.Count -gt 0) {
    $scope = 'Users'
  }

  # Condition - Determine primary triggering condition
  $condition = @()
  $conditions = $Policy.conditions

  # Risk-based conditions (highest priority)
  if ($conditions.userRiskLevels -and $conditions.userRiskLevels.Count -gt 0) {
    if ($conditions.userRiskLevels -contains 'high') {
      $condition += 'HighUserRisk'
    }
    else {
      $condition += 'UserRisk'
    }
  }
  if ($conditions.signInRiskLevels -and $conditions.signInRiskLevels.Count -gt 0) {
    if ($conditions.signInRiskLevels -contains 'high') {
      $condition += 'HighSignInRisk'
    }
    else {
      $condition += 'SignInRisk'
    }
  }

  # Location-based conditions
  if ($conditions.locations.excludeLocations -and $conditions.locations.excludeLocations.Count -gt 0) {
    $condition += 'UntrustedLocation'
  }
  if ($conditions.locations.includeLocations -and $conditions.locations.includeLocations.Count -gt 0) {
    $condition += 'SpecificLocation'
  }

  # Client app conditions
  if ($conditions.clientAppTypes -contains 'exchangeActiveSync' -or
    $conditions.clientAppTypes -contains 'other') {
    $condition += 'LegacyAuth'
  }

  # Authentication flows conditions
  if ($conditions.authenticationFlows -and $conditions.authenticationFlows.transferMethods -and $conditions.authenticationFlows.transferMethods.Count -gt 0) {
    $condition += 'AuthFlows'
  }

  # Platform conditions
  if ($conditions.platforms.includePlatforms -and $conditions.platforms.includePlatforms.Count -gt 0) {
    $platforms = $conditions.platforms.includePlatforms
    if ($platforms -contains 'android' -and $platforms -contains 'iOS') {
      $condition += 'MobileDevices'
    }
    elseif ($platforms.Count -eq 1) {
      $condition += $platforms[0].Substring(0, 1).ToUpper() + $platforms[0].Substring(1)
    }
    else {
      $condition += 'Platforms'
    }
  }

  # Device conditions
  if ($conditions.devices.includeDevices -and $conditions.devices.includeDevices.Count -gt 0) {
    if ($conditions.devices.includeDevices -contains 'All') {
      $condition += 'AllDevices'
    }
    else {
      $condition += 'Devices'
    }
  }

  # Device filter conditions
  if ($conditions.devices.deviceFilter -and $conditions.devices.deviceFilter.rule) {
    $condition += 'DeviceFilter'
  }

  # Application conditions
  if ($conditions.applications.includeApplications -and $conditions.applications.includeApplications.Count -gt 0) {
    $apps = $conditions.applications.includeApplications
    if ($apps -contains 'All') {
      $condition += 'AllApps'
    }
    elseif ($apps -contains 'Office365') {
      $condition += 'Office365'
    }
    else {
      $condition += 'CloudApps'
    }
  }

  # Default condition if none detected
  if ($condition.Count -eq 0) {
    $condition += 'General'
  }

  # Control - Determine primary enforcement action
  $control = 'Unknown'
  $grantControls = $Policy.grantControls
  $sessionControls = $Policy.sessionControls

  if ($grantControls.builtInControls -contains 'Block') {
    $control = 'Block'
  }
  elseif ($grantControls.builtInControls -contains 'Mfa') {
    if ($grantControls.authenticationStrength -and $grantControls.authenticationStrength.displayName) {
      $control = 'RequireAuthStrength'
    }
    elseif ($grantControls.builtInControls.Count -eq 1) {
      $control = 'RequireMFA'
    }
    else {
      $control = 'RequireMFA-Plus'
    }
  }
  elseif ($grantControls.builtInControls -contains 'CompliantDevice') {
    $control = 'RequireCompliantDevice'
  }
  elseif ($grantControls.builtInControls -contains 'DomainJoinedDevice') {
    $control = 'RequireDomainJoined'
  }
  elseif ($grantControls.builtInControls -contains 'ApprovedApplication') {
    $control = 'RequireApprovedApp'
  }
  elseif ($grantControls.builtInControls -contains 'CompliantApplication') {
    $control = 'RequireCompliantApp'
  }
  elseif ($grantControls.builtInControls -contains 'PasswordChange') {
    $control = 'RequirePasswordChange'
  }
  elseif ($grantControls.termsOfUse -and $grantControls.termsOfUse.Count -gt 0) {
    $control = 'RequireToU'
  }
  else {
    $control = 'Grant'
  }

  # Add session control modifiers
  $sessionModifiers = @()
  if ($sessionControls.signInFrequency -and $sessionControls.signInFrequency.isEnabled) {
    $sessionModifiers += 'SignInFreq'
  }
  if ($sessionControls.continuousAccessEvaluation -and $sessionControls.continuousAccessEvaluation.mode -eq 'strictEnforcement') {
    $sessionModifiers += 'CAE'
  }
  if ($sessionControls.persistentBrowser -and $sessionControls.persistentBrowser.isEnabled) {
    $sessionModifiers += 'PersistentBrowser'
  }
  if ($sessionControls.applicationEnforcedRestrictions -and $sessionControls.applicationEnforcedRestrictions.isEnabled) {
    $sessionModifiers += 'AppRestrictions'
  }
  if ($sessionControls.cloudAppSecurity -and $sessionControls.cloudAppSecurity.isEnabled) {
    $sessionModifiers += 'CloudAppSec'
  }

  # Append session modifiers to control
  if ($sessionModifiers.Count -gt 0) {
    $control += '-' + ($sessionModifiers -join '-')
  }

  # Add status modifier if reporting only
  if ($Policy.state -eq 'enabledForReportingButNotEnforced') {
    $control += '-ReportOnly'
  }

  # ID - Simple incremental ID (could be enhanced with actual policy counting)
  $id = '{0:D3}' -f (Get-Random -Maximum 999)

  # Combine components - limit condition to first 3 for readability with enhanced naming
  $conditionStr = ($condition | Select-Object -First 3) -join '-'
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
      'Microsoft.Graph.Identity.SignIns',
      'Microsoft.Graph.Identity.Governance',
      'Microsoft.Graph.DeviceManagement.Enrollment'
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
      try {
        Set-PSRepository -Name 'PSGallery' -InstallationPolicy Trusted -ErrorAction Stop 
      }
      catch {
        Write-Warn 'Could not set PSGallery as Trusted. You may be prompted during install.' 
      }
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
        # Hardened install: remove -SkipPublisherCheck / -AllowClobber; prompt avoidance left to repo trust
        Install-Module -Name $m -Scope CurrentUser -Force -ErrorAction Stop
      }
      catch {
        Write-Warn ("Failed to install module '{0}': {1}" -f $m, $_.Exception.Message)
        continue
      }
    }
    try { Import-Module -Name $m -Force -ErrorAction Stop } catch { Write-Warn ("Failed to import module '{0}': {1}" -f $m, $_.Exception.Message) }
  }
}

function Connect-GraphContext {
  <#
.SYNOPSIS
  Establish a Microsoft Graph connection if one does not already exist.
.DESCRIPTION
  Checks for an existing MgGraph context and, if missing, connects with the required scopes
  (Policy.Read.All, Directory.Read.All, RoleManagement.Read.All).
.NOTES
  Replaces Ensure-GraphConnection (deprecated) to satisfy approved verb list (Connect-*).
.EXAMPLE
  Connect-GraphContext
  Ensures the current session is connected with required scopes; returns immediately if already connected.
#>
  [CmdletBinding()]
  param()
  $ctx = $null
  try {
    $ctx = Get-MgContext -ErrorAction Stop 
  }
  catch {
    Write-Verbose 'No existing Graph context found (Get-MgContext failed).' 
  }
  $requiredScopes = 'Policy.Read.All', 'Directory.Read.All', 'RoleManagement.Read.All'
  $needsConnect = $false
  if (-not $ctx -or -not $ctx.Account) {
    $needsConnect = $true 
  }
  else {
    $granted = @($ctx.Scopes)
    foreach ($r in $requiredScopes) {
      if ($granted -notcontains $r) {
        $needsConnect = $true; break 
      }
    }
  }
  if ($needsConnect) {
    Write-Info 'Connecting to Microsoft Graph (Policy.Read.All, Directory.Read.All, RoleManagement.Read.All)'
    Connect-MgGraph -Scopes $requiredScopes | Out-Null
  }
}

function Invoke-SafeGet {
  <#
.SYNOPSIS
  Safely invoke a script block, returning $null on failure.
.DESCRIPTION
  Executes the provided script block with error trapping. Exceptions are suppressed (logged to Verbose)
  and $null is returned so that callers can continue a best-effort enrichment pattern.
.PARAMETER ScriptBlock
  The code to execute.
.EXAMPLE
  Invoke-SafeGet { Get-MgUser -UserId $id -Property Id,DisplayName }
#>
  [CmdletBinding()]
  param([Parameter(Mandatory)][ScriptBlock]$ScriptBlock)
  try {
    & $ScriptBlock 
  }
  catch {
    Write-Verbose ('Invoke-SafeGet suppressed error: {0}' -f $_.Exception.Message); return $null 
  }
}


function Convert-IdListToName {
  <#
.SYNOPSIS
  Convert a list of IDs to their friendly names when present in a lookup map.
.DESCRIPTION
  For each element in List, if the Map contains that key the mapped value is output; otherwise the original value.
  Null / empty input yields an empty array.
.PARAMETER List
  Collection of IDs or values.
.PARAMETER Map
  Hashtable keyed by ID with friendly values.
.EXAMPLE
  Convert-IdListToNames -List $policy.conditions.users.includeUsers -Map $UserMap
#>
  [CmdletBinding()]
  param([string[]]$List, [hashtable]$Map)
  if (-not $List) {
    return @() 
  }
  return $List | ForEach-Object { if ($Map.ContainsKey($_)) {
      $Map[$_] 
    }
    else {
      $_ 
    } }
}

function Test-IsGuid {
  <#
.SYNOPSIS
  Test whether a string is a GUID.
.DESCRIPTION
  Wraps [guid]::TryParse for readability and reuse.
  Filters out sentinel tokens (e.g. 'All') earlier in the pipeline.
.EXAMPLE
  Test-IsGuid 'd2719d52-3f4e-4f7c-9d0d-4f5c2a8ab123'
  Returns True.
.EXAMPLE
  Test-IsGuid 'All'
  Returns False.
  .PARAMETER Value
  The string value to test.

#>
  [CmdletBinding()]
  param([string]$Value)
  if (-not $Value) {
    return $false 
  }
  return [bool]([guid]::TryParse($Value, [ref]([guid]::Empty)))
}

# Unified resolver: given a list of IDs (users/groups/roles/apps), return friendly names when available.
function Resolve-EntityNameList {
  <#
  .SYNOPSIS
  Resolve a list of entity IDs to their friendly names.
  .DESCRIPTION
  For each element in Ids, if the Map contains that key the mapped value is output; otherwise the original value.
  Null / empty input yields an empty array.
.PARAMETER Ids
  Collection of IDs to resolve.
.PARAMETER UserMap
  Hashtable mapping user IDs to friendly names.
.PARAMETER GroupMap
  Hashtable mapping group IDs to friendly names.
.PARAMETER RoleMap
  Hashtable mapping role IDs to friendly names.
.PARAMETER AppMap
  Hashtable mapping application IDs to friendly names.
.EXAMPLE
  Convert-IdListToNames -List $policy.conditions.users.includeUsers -Map $UserMap
#>
  [CmdletBinding()]
  param(
    [string[]]$Ids,
    [hashtable]$UserMap,
    [hashtable]$GroupMap,
    [hashtable]$RoleMap,
    [hashtable]$AppMap
  )
  if (-not $Ids) {
    return @() 
  }
  return $Ids | ForEach-Object {
    $id = $_
    if ([string]::IsNullOrWhiteSpace($id)) {
      return $id 
    }
    if ($UserMap -and $UserMap.ContainsKey($id)) {
      return $UserMap[$id] 
    }
    if ($GroupMap -and $GroupMap.ContainsKey($id)) {
      return $GroupMap[$id] 
    }
    if ($RoleMap -and $RoleMap.ContainsKey($id)) {
      return $RoleMap[$id] 
    }
    if ($AppMap -and $AppMap.ContainsKey($id)) {
      return $AppMap[$id] 
    }
    return $id
  }
}

function Resolve-EntityGuidsInText {
  <#
  .SYNOPSIS
  Replace GUIDs in text with friendly names from provided maps.
  .DESCRIPTION
  Scans the input text for GUIDs and replaces them with their corresponding friendly names
  .PARAMETER Text
  The input text containing potential GUIDs to resolve.
  .PARAMETER UserMap
  Hashtable mapping user IDs to friendly names.
  .PARAMETER GroupMap
  Hashtable mapping group IDs to friendly names.
  .PARAMETER RoleMap
  Hashtable mapping role IDs to friendly names.
  .PARAMETER AppMap
  Hashtable mapping application IDs to friendly names.

  #>
  [CmdletBinding()]
  param(
    [string]$Text,
    [hashtable]$UserMap,
    [hashtable]$GroupMap,
    [hashtable]$RoleMap,
    [hashtable]$AppMap
  )
  # Touch parameters to satisfy static analyzers; they are also used within the regex scriptblock below
  $null = $UserMap; $null = $GroupMap; $null = $RoleMap; $null = $AppMap
  if ([string]::IsNullOrWhiteSpace($Text)) {
    return $Text 
  }
  # Standard GUID pattern
  $pattern = '[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}'
  return ([regex]::Replace($Text, $pattern, {
        param($m)
        $g = $m.Value
        if ($UserMap -and $UserMap.ContainsKey($g)) {
          return $UserMap[$g] 
        }
        if ($GroupMap -and $GroupMap.ContainsKey($g)) {
          return $GroupMap[$g] 
        }
        if ($RoleMap -and $RoleMap.ContainsKey($g)) {
          return $RoleMap[$g] 
        }
        if ($AppMap -and $AppMap.ContainsKey($g)) {
          return $AppMap[$g] 
        }
        return $g
      }))
}

#region Helper: Admin role & authentication strength evaluation (added in 3.3)
function Test-PolicyTargetsAdminRoles {
  <#
    .SYNOPSIS
      Determines whether a Conditional Access policy targets any privileged/admin roles.
    .DESCRIPTION
      Resolves included role GUIDs to display names (when cached lookups are available) and checks
      against a curated allowlist of administrative role names (case-insensitive).
    .PARAMETER Policy
      A raw Conditional Access policy object (as returned by Microsoft Graph) with Conditions.Users.IncludeRoles.
    .OUTPUTS
      [bool] True if policy targets at least one administrative role, otherwise False.
  #>
  param($Policy)
  if (-not $Policy -or -not $Policy.Conditions.Users.IncludeRoles) { return $false }
  $rawRoleIds = @($Policy.Conditions.Users.IncludeRoles)
  $resolved = @()
  foreach ($rid in $rawRoleIds) {
    if ([string]::IsNullOrWhiteSpace($rid)) { continue }
    if ($script:RoleMap -and $RoleMap.ContainsKey($rid)) { $resolved += $RoleMap[$rid] }
    elseif ($rid -match '^[0-9a-fA-F-]{36}$' -and $script:roleLookup -and $roleLookup.ContainsKey($rid)) { $resolved += $roleLookup[$rid] }
    else { $resolved += $rid }
  }
  $resolvedLower = $resolved | ForEach-Object { ($_ | Out-String).Trim().ToLowerInvariant() }
  $adminTargets = @(
    'privileged role administrator', 'global administrator', 'privileged authentication administrator', 'security administrator',
    'sharepoint administrator', 'exchange administrator', 'conditional access administrator', 'helpdesk administrator',
    'billing administrator', 'user administrator', 'authentication administrator', 'application administrator',
    'cloud application administrator', 'password administrator'
  )
  return (($resolvedLower | Where-Object { $adminTargets -contains $_ }).Count -gt 0)
}

function Initialize-AuthStrengthCache {
  if ($script:AuthStrengthCachePopulated) { return }
  $script:AuthStrengthCache = @{}
  if ($script:OfflineMode) {
    # Offline: cannot query strengths; leave empty cache so phish-resistant checks default to False.
    $script:AuthStrengthCachePopulated = $true
    return
  }
  $strengths = Invoke-SafeGet { Get-MgPolicyAuthenticationStrengthPolicy -All -Property Id, DisplayName, AllowedCombinations }
  if ($strengths) {
    foreach ($s in $strengths) {
      if ($s.Id) { $script:AuthStrengthCache[$s.Id] = $s }
      if ($s.DisplayName) { $script:AuthStrengthCache[$s.DisplayName] = $s }
    }
  }
  $script:AuthStrengthCachePopulated = $true
}

function Get-AssignedAuthStrengthObject {
  param($Policy)
  if (-not $Policy -or -not $Policy.GrantControls.AuthenticationStrength) { return $null }
  Initialize-AuthStrengthCache
  $as = $Policy.GrantControls.AuthenticationStrength
  if ($as.Id -and $script:AuthStrengthCache.ContainsKey($as.Id)) { return $script:AuthStrengthCache[$as.Id] }
  if ($as.DisplayName -and $script:AuthStrengthCache.ContainsKey($as.DisplayName)) { return $script:AuthStrengthCache[$as.DisplayName] }
  return $null 
}

function Test-IsPhishResistantStrength {
  <#
    .SYNOPSIS
      Determines if a policy's assigned Authentication Strength is phishing resistant.
    .DESCRIPTION
      Retrieves the assigned authentication strength object (cached) and inspects AllowedCombinations.
      If any combination is in the non-phish-resistant set, policy is NOT phishing resistant.
    .PARAMETER Policy
      Raw Conditional Access policy object with GrantControls.AuthenticationStrength populated.
    .OUTPUTS
      [bool] True when strength is considered phishing resistant; otherwise False.
  #>
  param($Policy)
  $strengthObj = Get-AssignedAuthStrengthObject -Policy $Policy
  if (-not $strengthObj -or -not $strengthObj.AllowedCombinations) { return $false }
  $nonPhish = @('deviceBasedPush', 'temporaryAccessPassOneTime', 'temporaryAccessPassMultiUse', 'microsoftAuthenticatorPush', 'sms', 'voice', 'softwareOath', 'hardwareOath', 'x509CertificateSingleFactor', 'federatedSingleFactor', 'qrCodePin')
  $contains = ($strengthObj.AllowedCombinations | Where-Object { $nonPhish -contains $_ }).Count -gt 0
  return (-not $contains)
}

function Test-PolicyRequiresMfaForAdmins {
  param($Policy)
  if (-not (Test-PolicyTargetsAdminRoles -Policy $Policy)) { return $false }
  $strengthName = ''
  if ($Policy.GrantControls.AuthenticationStrength.DisplayName) { $strengthName = [string]$Policy.GrantControls.AuthenticationStrength.DisplayName }
  $l = $strengthName.ToLowerInvariant()
  $hasImpliedMfa = ($Policy.GrantControls.BuiltInControls -contains 'Mfa') -or ($l -match 'phishing') -or ($l -match 'passwordless') -or ($l -match 'multifactor') -or ($l -match '\bmfa\b')
  return $hasImpliedMfa 
}

function Test-PolicyRequiresPhishResistantMfaForAdmins {
  param($Policy)
  if (-not (Test-PolicyTargetsAdminRoles -Policy $Policy)) { return $false }
  return (Test-IsPhishResistantStrength -Policy $Policy) 
}

function Test-OverlapIncludeExclude {
  param($Include, $Exclude)
  if (-not $Include -or -not $Exclude) { return $false }
  $inc = @($Include) | Where-Object { $_ -ne $null -and $_ -ne '' }
  $exc = @($Exclude) | Where-Object { $_ -ne $null -and $_ -ne '' }
  if ($inc.Count -eq 0 -or $exc.Count -eq 0) { return $false }
  foreach ($i in $inc) { if ($exc -contains $i) { return $true } }
  return $false 
}

function New-TokenSet {
  param($Value)
  $set = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
  if ($null -ne $Value) { foreach ($tok in (@($Value) -join "`n") -split '[,\n]') { $t = ($tok).Trim(); if ($t) { $null = $set.Add($t) } } }
  return $set 
}

function Protect-RecNote {
  param([string]$Html)
  if ([string]::IsNullOrWhiteSpace($Html)) { return $Html }
  # The note markup we inject is fully script-controlled (policy detail cards); heavy encoding caused visible artifacts.
  # Strategy: strip executable/script/style content + dangerous inline handlers + javascript: URIs, otherwise pass through.
  $clean = $Html -replace '(?is)<script[^>]*>.*?</script>', '' -replace '(?is)<style[^>]*>.*?</style>', ''
  $clean = $clean -replace '(?i) on[a-z]+\s*=\s*"[^"]*"', '' -replace "(?i) on[a-z]+\s*=\s*'[^']*'", ''
  $clean = $clean -replace '(?i)href\s*=\s*"javascript:[^"]*"', 'href="#"' -replace "(?i)href\s*=\s*'javascript:[^']*'", "href='#'"
  return $clean 
}
#endregion Helper additions

#region CA Check Functions (extracted from inline hashtable for testability)
function Test-CA00 {
  param($PolicyCheck)
  $PolicyCheck.GrantControls.BuiltInControls -contains 'Block' -and
  $PolicyCheck.Conditions.ClientAppTypes -contains 'exchangeActiveSync' -and
  $PolicyCheck.Conditions.ClientAppTypes -contains 'other'
}
function Test-CA01 {
  param($PolicyCheck)
  $PolicyCheck.GrantControls.BuiltInControls -contains 'Mfa' -and
  $PolicyCheck.Conditions.Users.IncludeUsers -eq 'all' -and
  $PolicyCheck.Conditions.Applications.IncludeApplications -eq 'all'
}
function Test-CA02 {
  param($PolicyCheck)
  ($PolicyCheck.Conditions.Platforms.IncludePlatforms -contains 'android' -or
  $PolicyCheck.Conditions.Platforms.IncludePlatforms -contains 'iOS' -or
  $PolicyCheck.Conditions.Platforms.IncludePlatforms -contains 'windowsPhone') -and
  ($PolicyCheck.GrantControls.BuiltInControls -contains 'approvedApplication' -or
  $PolicyCheck.GrantControls.BuiltInControls -contains 'compliantApplication' -or
  $PolicyCheck.GrantControls.BuiltInControls -contains 'compliantDevice')
}
function Test-CA03 {
  param($PolicyCheck)
  ($PolicyCheck.Conditions.Platforms.IncludePlatforms -contains 'windows' -or
  $PolicyCheck.Conditions.Platforms.IncludePlatforms -contains 'macOS') -and
  ($PolicyCheck.GrantControls.BuiltInControls -contains 'compliantDevice' -or
  $PolicyCheck.GrantControls.BuiltInControls -contains 'domainJoinedDevice')
}
function Test-CA06 {
  param($PolicyCheck)
  (Test-OverlapIncludeExclude $PolicyCheck.Conditions.Users.IncludeUsers $PolicyCheck.Conditions.Users.ExcludeUsers) -or
  (Test-OverlapIncludeExclude $PolicyCheck.Conditions.Users.IncludeGroups $PolicyCheck.Conditions.Users.ExcludeGroups) -or
  (Test-OverlapIncludeExclude $PolicyCheck.Conditions.Users.IncludeRoles $PolicyCheck.Conditions.Users.ExcludeRoles) -or
  (Test-OverlapIncludeExclude $PolicyCheck.Conditions.Platforms.IncludePlatforms $PolicyCheck.Conditions.Platforms.ExcludePlatforms) -or
  (Test-OverlapIncludeExclude $PolicyCheck.Conditions.Locations.IncludeLocations $PolicyCheck.Conditions.Locations.ExcludeLocations) -or
  (Test-OverlapIncludeExclude $PolicyCheck.Conditions.Applications.IncludeApplications $PolicyCheck.Conditions.Applications.ExcludeApplications)
}

function Test-CA07 {
  param($PolicyCheck)
  ([string]::IsNullOrWhiteSpace($PolicyCheck.Conditions.Users.IncludeUsers) -or $PolicyCheck.Conditions.Users.IncludeUsers.Count -eq 0 -or $PolicyCheck.Conditions.Users.IncludeUsers -eq 'None') -and
  (([string]::IsNullOrWhiteSpace($PolicyCheck.Conditions.Users.IncludeGroups)) -or $PolicyCheck.Conditions.Users.IncludeGroups.Count -eq 0) -and
  (([string]::IsNullOrWhiteSpace($PolicyCheck.Conditions.Users.IncludeRoles)) -or $PolicyCheck.Conditions.Users.IncludeRoles.Count -eq 0) -and
  ($null -eq $PolicyCheck.Conditions.Users.IncludeGuestsOrExternalUsers.GuestOrExternalUserTypes)
}

function Test-CA08 {
  param($PolicyCheck)
  $PolicyCheck.Conditions.Users.IncludeUsers -ne 'None' -and
  $null -ne $PolicyCheck.Conditions.Users.IncludeUsers -and
  $PolicyCheck.Conditions.Users.IncludeUsers -ne 'All' -and
  $PolicyCheck.Conditions.Users.IncludeUsers -ne 'GuestsOrExternalUsers'
}
function Test-CA09 {
  param($PolicyCheck)
  ($null -ne $PolicyCheck.Conditions.SignInRiskLevels) -or
  ($null -ne $PolicyCheck.Conditions.UserRiskLevels)
}
function Test-CA10 {
  param($PolicyCheck)
  $PolicyCheck.Conditions.AdditionalProperties.authenticationFlows.Values -split ',' -contains 'deviceCodeFlow' -and
  $PolicyCheck.grantcontrols.BuiltInControls -contains 'Block'
}
function Test-CA11 {
  param($PolicyCheck)
  ($PolicyCheck.Conditions.Applications.IncludeUserActions -contains 'urn:user:registerdevice') -and
  ($PolicyCheck.GrantControls.BuiltInControls -contains 'Mfa')
}
function Test-CA12 {
  param($PolicyCheck)
  ($PolicyCheck.GrantControls.BuiltInControls -contains 'Block') -and
  ($PolicyCheck.Conditions.Platforms.IncludePlatforms -contains 'all') -and
  ($PolicyCheck.Conditions.Platforms.ExcludePlatforms.Count -gt 0)
}
#endregion CA Check Functions
#endregion Helper additions

# Backward compatibility aliases (deprecated names). Retained temporarily so external callers
# referencing prior function names do not break. Marked for removal in a future major version.
Set-Alias -Name Ensure-GraphConnection -Value Connect-GraphContext -ErrorAction SilentlyContinue
Set-Alias -Name Safe-Get -Value Invoke-SafeGet -ErrorAction SilentlyContinue
Set-Alias -Name Translate-List -Value Convert-IdListToName -ErrorAction SilentlyContinue

# Determine offline mode early so we can skip any Graph interaction
$script:OfflineMode = [bool]$RawInputFile
if (-not $script:OfflineMode) {
  # Ensure required modules are present before attempting to connect
  Initialize-GraphModule
  Connect-GraphContext
}
else {
  Write-Info 'Offline mode detected: skipping module initialization and Graph connection.'
}

# Script metadata / version stamp (bump when feature changes)
$Script:CAExportVersion = '3.3'

if ($OutputPath) {
  try {
    if (-not (Test-Path -LiteralPath $OutputPath)) {
      New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null 
    }
    $ExportLocation = (Resolve-Path -LiteralPath $OutputPath).Path
    Write-Info "Using custom output path: $ExportLocation"
  }
  catch {
    Write-Warn ("Failed to set custom OutputPath '{0}': {1}" -f $OutputPath, $_.Exception.Message)
  }
}

# If no -OutputPath was supplied or resolution failed, default to current working directory
if (-not $ExportLocation) {
  $ExportLocation = (Get-Location).Path
  Write-Info "No OutputPath specified – defaulting to current directory: $ExportLocation"
}

# ---------------- Retrieve Tenant & Policies ----------------
if ($RawInputFile) {
  Write-Info "Offline mode: importing raw policy JSON from $RawInputFile"
  try {
    $rawText = Get-Content -LiteralPath $RawInputFile -Raw -ErrorAction Stop
    $allPolicies = $rawText | ConvertFrom-Json -ErrorAction Stop
  }
  catch {
    Write-Err ("Failed to read or parse RawInputFile '{0}': {1}" -f $RawInputFile, $_.Exception.Message)
    exit 1
  }
  if (-not $allPolicies) {
    Write-Err 'Parsed zero policies from provided raw file. Exiting.'
    exit 1
  }
  $TenantName = 'OfflineTenant'
  $Date = (Get-Date).ToString('u')
  if ($PolicyID) {
    $CAPolicy = @($allPolicies | Where-Object { $_.id -eq $PolicyID })
  }
  else {
    $CAPolicy = @($allPolicies)
  }
  Write-Info "Loaded $($CAPolicy.Count) policies from raw file"
}
else {
  Write-Info 'Retrieving tenant information'
  $TenantName = (Get-MgOrganization -ErrorAction SilentlyContinue | Select-Object -First 1 -ExpandProperty DisplayName)
  if (-not $TenantName) { $TenantName = 'UnknownTenant' }
  $Date = (Get-Date).ToString('u')
  Write-Info 'Retrieving Conditional Access policies'
  try {
    $allPolicies = Get-MgIdentityConditionalAccessPolicy -All
  }
  catch {
    Write-Err ('Failed to retrieve policies: {0}' -f $_.Exception.Message)
    Write-Err 'Cannot continue without policies. Exiting.'
    exit 1
  }
  if ($PolicyID) {
    $CAPolicy = $allPolicies | Where-Object { $_.id -eq $PolicyID }
    if (-not $CAPolicy) {
      Write-Err "Policy with ID '$PolicyID' not found. Exiting."
      exit 1
    }
  }
  else {
    $CAPolicy = $allPolicies
  }
  if (-not $CAPolicy -or $CAPolicy.Count -eq 0) {
    Write-Err 'No Conditional Access policies found in tenant. Exiting.'
    exit 1
  }
  Write-Info "Successfully retrieved $($CAPolicy.Count) Conditional Access $(if($CAPolicy.Count -eq 1){'policy'}else{'policies'})"
}

# Raw snapshot & index
$RawPolicyIndex = @{}
foreach ($rp in $CAPolicy) {
  if ($rp.id) {
    $RawPolicyIndex[$rp.id] = $rp 
  } 
}

# ---------------- Collect IDs for enrichment ----------------
$userIds = [System.Collections.Generic.HashSet[string]]::new()
$groupIds = [System.Collections.Generic.HashSet[string]]::new()
$roleIds = [System.Collections.Generic.HashSet[string]]::new()
$appIds = [System.Collections.Generic.HashSet[string]]::new()
$locIds = [System.Collections.Generic.HashSet[string]]::new()
$touIds = [System.Collections.Generic.HashSet[string]]::new()

# Collect IDs for enrichment (consolidated for clarity)
foreach ($p in $CAPolicy) {
  $c = $p.conditions
  if ($c.users) {
    # Users (excluding sentinels)
    foreach ($i in @($c.users.includeUsers)) {
      if ($i -and $i -notin @('All', 'None', 'GuestsOrExternalUsers')) { [void]$userIds.Add($i) }
    }
    foreach ($i in @($c.users.excludeUsers)) { if ($i) { [void]$userIds.Add($i) } }
    # Groups
    foreach ($i in @($c.users.includeGroups)) { if ($i) { [void]$groupIds.Add($i) } }
    foreach ($i in @($c.users.excludeGroups)) { if ($i) { [void]$groupIds.Add($i) } }
    # Roles
    foreach ($i in @($c.users.includeRoles)) { if ($i) { [void]$roleIds.Add($i) } }
    foreach ($i in @($c.users.excludeRoles)) { if ($i) { [void]$roleIds.Add($i) } }
  }
  if ($c.applications) {
    foreach ($i in @($c.applications.includeApplications)) { if ($i -and (Test-IsGuid $i)) { [void]$appIds.Add($i) } }
    foreach ($i in @($c.applications.excludeApplications)) { if ($i -and (Test-IsGuid $i)) { [void]$appIds.Add($i) } }
  }
  if ($c.locations) {
    foreach ($i in @($c.locations.includeLocations)) { if ($i -and (Test-IsGuid $i)) { [void]$locIds.Add($i) } }
    foreach ($i in @($c.locations.excludeLocations)) { if ($i -and (Test-IsGuid $i)) { [void]$locIds.Add($i) } }
  }
  if ($p.grantControls -and $p.grantControls.termsOfUse) {
    foreach ($i in @($p.grantControls.termsOfUse)) { if ($i -and (Test-IsGuid $i)) { [void]$touIds.Add($i) } }
  }
}

# ---------------- Build lookup maps (best-effort) ----------------
$UserMap = @{}; $GroupMap = @{}; $RoleMap = @{}; $AppMap = @{}; $LocMap = @{}; $TouMap = @{}

if ($script:OfflineMode) {
  Write-Info 'Offline mode: skipping user/group/role/app/location/TOU Graph enrichment.'
}
else {
  foreach ($id in $userIds) {
    if ($id -and ($id -match '^[0-9a-fA-F\-]{36}$')) {
      $obj = Invoke-SafeGet { Get-MgUser -UserId $id -Property Id, DisplayName }
      if ($obj) { $UserMap[$id] = $obj.DisplayName }
    }
  }
  foreach ($id in $groupIds) {
    $obj = Invoke-SafeGet { Get-MgGroup -GroupId $id -Property Id, DisplayName } ; if ($obj) { $GroupMap[$id] = $obj.DisplayName }
  }

  $roleLookup = @{}
  $commonRoleTemplates = @{
    '9b895d92-2cd3-44c7-9d02-a6ac2d5ea5c3' = 'Application Administrator'
    'c4e39bd9-1100-46d3-8c65-fb160da0071f' = 'Authentication Administrator'
    'e3973bdf-4987-49ae-837a-ba8e231c7286' = 'Azure DevOps Administrator'
    'b0f54661-2d74-4c50-afa3-1ec803f12efe' = 'Billing Administrator'
    '158c047a-c907-4556-b7ef-446551a6b5f7' = 'Cloud Application Administrator'
    'b1be1c3e-b65d-4f19-8427-f6fa0d97feb9' = 'Conditional Access Administrator'
    '29232cdf-9323-42fd-ade2-1d097af3e4de' = 'Exchange Administrator'
    '62e90394-69f5-4237-9190-012177145e10' = 'Global Administrator'
    '729827e3-9c14-49f7-bb1b-9608f156bbb8' = 'Helpdesk Administrator'
    '966707d0-3269-4727-9be2-8c3a10f19b9d' = 'Password Administrator'
    '7be44c8a-adaf-4e2a-84d6-ab2649e08a13' = 'Privileged Authentication Administrator'
    'e8611ab8-c189-46e8-94e1-60213ab1f814' = 'Privileged Role Administrator'
    '194ae4cb-b126-40b2-bd5b-6091b380977d' = 'Security Administrator'
    '5d6b6bb7-de71-4623-b4af-96380a352509' = 'Security Reader'
    'f28a1f50-f6e7-4571-818b-6a12f2af6b6c' = 'SharePoint Administrator'
    'fe930be7-5e62-47db-91af-98c3a49a38b1' = 'User Administrator'
  }

  $dirRoles = Invoke-SafeGet { Get-MgDirectoryRole -All }
  if ($dirRoles) {
    foreach ($r in $dirRoles) {
      if ($r.id -and -not $roleLookup.ContainsKey($r.id)) { $roleLookup[$r.id] = $r.displayName }
      if ($r.roleTemplateId -and -not $roleLookup.ContainsKey($r.roleTemplateId)) { $roleLookup[$r.roleTemplateId] = $r.displayName }
    }
  }
  else {
    foreach ($kv in $commonRoleTemplates.GetEnumerator()) { $roleLookup[$kv.Key] = $kv.Value }
  }

  $roleDefs = Invoke-SafeGet { Get-MgRoleManagementDirectoryRoleDefinition -All -Property Id, DisplayName, TemplateId }
  if ($roleDefs) {
    foreach ($rd in $roleDefs) {
      if ($rd.Id -and -not $roleLookup.ContainsKey($rd.Id)) { $roleLookup[$rd.Id] = $rd.DisplayName }
      if ($rd.TemplateId -and -not $roleLookup.ContainsKey($rd.TemplateId)) { $roleLookup[$rd.TemplateId] = $rd.DisplayName }
    }
  }
  else {
    foreach ($kv in $commonRoleTemplates.GetEnumerator()) { if (-not $roleLookup.ContainsKey($kv.Key)) { $roleLookup[$kv.Key] = $kv.Value } }
  }

  foreach ($id in $roleIds) {
    if ($roleLookup.ContainsKey($id)) {
      $RoleMap[$id] = $roleLookup[$id]
    }
    else {
      $obj = Invoke-SafeGet { Get-MgDirectoryRole -DirectoryRoleId $id -Property Id, DisplayName }
      if ($obj) { $RoleMap[$id] = $obj.DisplayName } elseif ($commonRoleTemplates.ContainsKey($id)) { $RoleMap[$id] = $commonRoleTemplates[$id] }
    }
  }
  foreach ($id in $appIds) {
    if (Test-IsGuid $id) {
      $obj = Invoke-SafeGet { Get-MgServicePrincipal -ServicePrincipalId $id -Property Id, DisplayName, AppId }
      if ($obj) { $AppMap[$id] = $obj.DisplayName }
    }
  }
  foreach ($id in $locIds) {
    if (Test-IsGuid $id) {
      $obj = Invoke-SafeGet { Get-MgIdentityConditionalAccessNamedLocation -NamedLocationId $id -Property Id, DisplayName }
      if ($obj) { $LocMap[$id] = $obj.DisplayName }
    }
  }
  foreach ($id in $touIds) {
    if (Test-IsGuid $id) {
      $obj = Invoke-SafeGet { Get-MgAgreement -AgreementId $id -Property Id, DisplayName }
      if ($obj) { $TouMap[$id] = $obj.DisplayName }
    }
  }
}

# ---------------- Construct CAExport ----------------
$CAExport = @()
foreach ($Policy in $CAPolicy) {
  $DateModified = if ($Policy.modifiedDateTime) {
    $Policy.modifiedDateTime 
  }
  else {
    $Policy.createdDateTime 
  }
  $InclPlat = $Policy.conditions.platforms.includePlatforms
  $ExclPlat = $Policy.conditions.platforms.excludePlatforms
  $InclDev = $Policy.conditions.devices.includeDevices
  $ExclDev = $Policy.conditions.devices.excludeDevices
  $devFilters = $Policy.conditions.devices.deviceFilter.rule
  $authenticationFlowsString = ( $Policy.conditions.additionalProperties.authenticationFlows.values -join ', ' )
  $InclLocation = $Policy.conditions.locations.includeLocations | ForEach-Object { if ($_ -and (Test-IsGuid $_) -and $LocMap.ContainsKey($_)) {
      $LocMap[$_] 
    }
    else {
      $_ 
    } }
  $ExclLocation = $Policy.conditions.locations.excludeLocations | ForEach-Object { if ($_ -and (Test-IsGuid $_) -and $LocMap.ContainsKey($_)) {
      $LocMap[$_] 
    }
    else {
      $_ 
    } }
  $IncludeUG = @()
  $IncludeUG += (Convert-IdListToName $Policy.conditions.users.includeUsers $UserMap)
  $IncludeUG += (Convert-IdListToName $Policy.conditions.users.includeGroups $GroupMap)
  $IncludeUG += (Convert-IdListToName $Policy.conditions.users.includeRoles $RoleMap)
  if ($Policy.conditions.users.includeGuestsOrExternalUsers.guestOrExternalUserTypes) {
    $IncludeUG += $Policy.conditions.users.includeGuestsOrExternalUsers.guestOrExternalUserTypes 
  }
  $ExcludeUG = @()
  $ExcludeUG += (Convert-IdListToName $Policy.conditions.users.excludeUsers $UserMap)
  $ExcludeUG += (Convert-IdListToName $Policy.conditions.users.excludeGroups $GroupMap)
  $ExcludeUG += (Convert-IdListToName $Policy.conditions.users.excludeRoles $RoleMap)
  if ($Policy.conditions.users.excludeGuestsOrExternalUsers.guestOrExternalUserTypes) {
    $ExcludeUG += $Policy.conditions.users.excludeGuestsOrExternalUsers.guestOrExternalUserTypes 
  }

  $CAExport += [PSCustomObject][ordered]@{
    Name                                = $Policy.displayName
    'Recommended Name'                  = Get-RecommendedPolicyName -Policy $Policy
    PolicyId                            = $Policy.id
    Status                              = Format-PolicyStatus -Status $Policy.state
    Modified                            = $DateModified
    Created                             = $Policy.createdDateTime
    Description                         = $Policy.description
    'Included Users'                    = ($IncludeUG -join ", `r`n")
    'Excluded Users'                    = ($ExcludeUG -join ", `r`n")
    'Included Applications'             = ($Policy.conditions.applications.includeApplications -join ", `r`n")
    'Excluded Applications'             = ($Policy.conditions.applications.excludeApplications -join ", `r`n")
    'User Actions'                      = (($Policy.conditions.applications.includeUserActions | ForEach-Object { 
                                              if ($_ -eq 'urn:user:registersecurityinfo') { 'Register Security Info' }
                                              elseif ($_ -eq 'urn:user:registerdevice') { 'Register Device' }
                                              else { $_ }
                                            }) -join ", `r`n")
    'Auth Context'                      = ($Policy.conditions.applications.includeAuthenticationContextClassReferences -join ", `r`n")
    'User Risk'                         = ($Policy.conditions.userRiskLevels -join ", `r`n")
    'SignIn Risk'                       = ($Policy.conditions.signInRiskLevels -join ", `r`n")
    'Platforms Included'                = ($InclPlat -join ", `r`n")
    'Platforms Excluded'                = ($ExclPlat -join ", `r`n")
    'Included Locations'                = ($InclLocation -join ", `r`n")
    'Excluded Locations'                = ($ExclLocation -join ", `r`n")
    'Client Apps'                       = ($Policy.conditions.clientAppTypes -join ", `r`n")
    'Included Devices'                  = ($InclDev -join ", `r`n")
    'Excluded Devices'                  = ($ExclDev -join ", `r`n")
    'Device Filters'                    = ($devFilters -join ", `r`n")
    'Authentication Flows'              = $authenticationFlowsString
    'Block'                             = if ($Policy.grantControls.builtInControls -contains 'Block') {
      'True' 
    }
    else {
      '' 
    }
    'Require MFA'                       = if ($Policy.grantControls.builtInControls -contains 'Mfa') {
      'True' 
    }
    else {
      '' 
    }
    'Authentication Strength MFA'       = $Policy.grantControls.authenticationStrength.displayName
    'Compliant Device'                  = if ($Policy.grantControls.builtInControls -contains 'CompliantDevice') {
      'True' 
    }
    else {
      '' 
    }
    'Domain Joined Device'              = if ($Policy.grantControls.builtInControls -contains 'DomainJoinedDevice') {
      'True' 
    }
    else {
      '' 
    }
    'Compliant Application'             = if ($Policy.grantControls.builtInControls -contains 'CompliantApplication') {
      'True' 
    }
    else {
      '' 
    }
    'Approved Application'              = if ($Policy.grantControls.builtInControls -contains 'ApprovedApplication') {
      'True' 
    }
    else {
      '' 
    }
    'Password Change'                   = if ($Policy.grantControls.builtInControls -contains 'PasswordChange') {
      'True' 
    }
    else {
      '' 
    }
    'Terms Of Use'                      = ((Convert-IdListToName $Policy.grantControls.termsOfUse $TouMap) -join ", `r`n")
    'Custom Controls'                   = ($Policy.grantControls.customAuthenticationFactors -join ", `r`n")
    'Grant Operator'                    = $Policy.grantControls.operator
    'Application Enforced Restrictions' = $Policy.sessionControls.applicationEnforcedRestrictions.isEnabled
    'Cloud App Security'                = $Policy.sessionControls.cloudAppSecurity.isEnabled
    'Sign In Frequency'                 = if ($Policy.sessionControls.signInFrequency.value -and $Policy.sessionControls.signInFrequency.type) {
      "$( $Policy.sessionControls.signInFrequency.value ) $( $Policy.sessionControls.signInFrequency.type )" 
    }
    'Persistent Browser'                = $Policy.sessionControls.persistentBrowser.mode
    'Continuous Access Evaluation'      = $Policy.sessionControls.continuousAccessEvaluation.mode
    'Resilient Defaults'                = $Policy.sessionControls.disableResilienceDefaults
    'Secure Sign In Session'            = $Policy.sessionControls.additionalProperties.secureSignInSession.values
    RawJson                             = ''
  }
}

# Map external switch parameters to internal export control flags for backward compatibility / clarity
if ($PSBoundParameters.ContainsKey('Html')) {
  $HTMLExport = [bool]$Html 
}
if ($PSBoundParameters.ContainsKey('Json')) {
  $JsonExport = [bool]$Json 
}
if ($PSBoundParameters.ContainsKey('Csv')) {
  $CsvExport = [bool]$Csv 
}
if ($PSBoundParameters.ContainsKey('CsvPivot')) {
  $CsvPivotExport = [bool]$CsvPivot 
}
# Default behavior: if no explicit export switches supplied, emit HTML + JSON + CSV (pivot remains opt-in)
if (-not ($HTMLExport -or $JsonExport -or $CsvExport -or $CsvPivotExport)) {
  $HTMLExport = $true; $JsonExport = $true; $CsvExport = $true 
}
$LinkURL = 'https://entra.microsoft.com/#view/Microsoft_AAD_ConditionalAccess/PolicyBlade/policyId/'
$baseName = "CAExportRecs_${TenantName}_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
$FileName = "$baseName.html"
$JsonFileName = "$baseName.json"
$CsvFileName = "$baseName.csv"
$CsvPivotFileName = "$baseName-pivot.csv"
$HtmlParts = @()
if (-not $NoRecommendations) {
  Write-Info 'Analyzing: getting recommendations'
  if (-not ([System.Management.Automation.PSTypeName]'Recommendation').Type) {
    class Recommendation {
      [string]$Control
      [string]$Name
      [string]$PassText
      [string]$FailRecommendation
      [string]$Importance
      [hashtable]$Links
      [bool]$Status
      [bool]$SwapStatus
      [string]$Note
      [string[]]$Excluded
      Recommendation([string]$Control, [string]$Name, [string]$PassText, [string]$FailRecommendation, [string]$Importance, [hashtable]$Links, [bool]$Status, [bool]$SwapStatus) {
        $this.Control = $Control
        $this.Name = $Name
        $this.PassText = $PassText
        $this.FailRecommendation = $FailRecommendation
        $this.Importance = $Importance
        $this.Links = $Links
        $this.Status = $Status
        $this.SwapStatus = $SwapStatus
        $this.Note = ''
        $this.Excluded = @()
      }
    }
  }

  # Recommendations based off Microsoft security templates.
  # SCuBA: Secure Cloud Business Applications
  <#
  $recommendations is the in-memory collection that accumulates actionable findings ("recommendations")
  discovered while analyzing Conditional Access (CA) policies and related configuration. Each element is
  an instance of the [recommendation] type, which models a single issue, improvement, or best-practice
  suggestion along with its supporting context.

.VARIABLE
  $recommendations
    - Type: [recommendation[]] (array or list of [recommendation] objects)
    - Purpose: Collects one [recommendation] per detected condition that warrants attention or change.
    - Lifecycle: Initialized before analysis; appended to as the script evaluates policies; consumed by
      exporters or written to output at the end of the run.

.TYPE
  [recommendation]
    A custom type representing a single actionable recommendation. Typical fields include:
      - Id            : string      // Stable identifier for correlation/deduplication
      - Title         : string      // Short, human-readable summary
      - Description   : string      // Detailed rationale and context/evidence
      - Severity      : string      // Info | Low | Medium | High | Critical
      - Category      : string      // e.g., Security, Reliability, Hygiene, Compliance
      - AppliesTo     : string[]    // Names/IDs of affected policies, assignments, or scopes
      - Remediation   : string[]    // Concrete steps or guidance to address the finding
      - References    : string[]    // URLs or document identifiers for further reading
      - Tags          : string[]    // Optional labels used for filtering/reporting
      - Metadata      : hashtable   // Optional free-form data for exporters or auditing

.USAGE
  - Creating and adding:
      # $recommendations += [recommendation]::new(<id>, <title>, <description>, <severity>, ...)
      # or instantiate and set properties, then append to $recommendations.

  - Consuming:
      # Filter, group, and export $recommendations (e.g., JSON/CSV/Markdown) as needed by downstream steps.

#>
  # Inline recommendations (moved back from external PSD1)
  $recommendations = @(
    [recommendation]::new('CA-00', 'Legacy Authentication', 'Legacy Authentication is blocked or minimized, targeting Legacy Authentication protocols.', 'Review and update policies to restrict or block Legacy Authentication protocols to ensure security.', 'Legacy Authentication protocols are outdated and less secure. It is recommended to block or minimize their usage to enhance the security of your environment.', @{ 'Legacy Authentication Overview' = 'https://learn.microsoft.com/en-us/entra/identity/conditional-access/policy-block-legacy-authentication' }, $false, $true),
    [recommendation]::new('CA-01', 'MFA Policy targets All users Group and All Cloud Apps', 'There is at least one policy that targets all users and cloud apps.', 'Review and update MFA policies to ensure they target all users and cloud apps, including any necessary exclusions.', 'Multi-factor Authentication (MFA) should apply to all users and cloud apps as a baseline for security. Policies should include the necessary exclusions if required but should primarily target all users and apps for maximum security.', @{ 'The Challenge with Targeted Architecture' = 'https://learn.microsoft.com/en-us/azure/architecture/guide/security/conditional-access-architecture#:~:text=The%20challenge%20with%20the,that%20number%20isn%27t%20supported.' }, $false, $true),
    [recommendation]::new('CA-02', 'Mobile Device Policy requires MDM or MAM', 'There is at least one policy that requires MDM or MAM for mobile devices.', 'Consider adding policies to check for device management, either through MDM or MAM, to ensure secure mobile access.', 'Mobile Device Management (MDM) or Mobile Application Management (MAM) should be enforced to ensure that mobile devices accessing organizational data are properly managed and secure. Policies should include requirements for MDM or MAM to increase security for mobile devices.', @{ 'MAM Overview' = 'https://learn.microsoft.com/en-us/mem/intune/apps/app-management#mobile-application-management-mam-basics'; 'Protect Data on personally owned devices' = 'https://smbtothecloud.com/protecting-company-data-on-personally-owned-devices/' }, $false, $true),
    [recommendation]::new('CA-03', 'Require Hybrid Join or Intune Compliance on Windows or Mac', 'There is at least one policy that requires Hybrid Join or Intune Compliance for Windows or Mac devices.', 'Consider adding policies to ensure that Windows or Mac devices are either Hybrid Joined or compliant with Intune to enhance security.', 'Hybrid Join or Intune Compliance should be enforced to ensure that Windows or Mac devices accessing organizational data are properly managed and secure. Policies should include requirements for Hybrid Join or Intune Compliance to increase security for these devices.', @{ 'Hybrid Join Overview' = 'https://learn.microsoft.com/en-us/azure/active-directory/devices/hybrid-azuread-join-plan'; 'Intune Compliance Overview' = 'https://learn.microsoft.com/en-us/mem/intune/protect/compliance-policy-create-windows' }, $false, $true),
    [recommendation]::new('CA-04', 'Require MFA for Admins', 'There is at least one policy that requires Multi-Factor Authentication (MFA) for administrators.', 'Consider adding policies to ensure that administrators are required to use Multi-Factor Authentication (MFA) to enhance security.', 'Multi-Factor Authentication (MFA) should be enforced for administrators to ensure that access to critical systems and data is secure. Policies should include requirements for MFA to increase security for administrative accounts.', @{ 'MFA Overview' = 'https://learn.microsoft.com/en-us/azure/active-directory/authentication/concept-mfa-howitworks'; 'MFA for Admins' = 'https://learn.microsoft.com/en-us/entra/identity/conditional-access/policy-old-require-mfa-admin' }, $false, $true),
    [recommendation]::new('CA-05', 'Require Phish-Resistant MFA for Admins', 'There is at least one policy that requires phish-resistant Multi-Factor Authentication (MFA) for administrators.', 'Consider adding policies to ensure that administrators are required to use phish-resistant Multi-Factor Authentication (MFA) to enhance security.', 'Phish-resistant Multi-Factor Authentication (MFA) should be enforced for administrators to ensure secure access.', @{ 'MSFT Authentication Strengths' = 'https://learn.microsoft.com/en-us/entra/identity/authentication/concept-authentication-strengths'; 'Phish-Resistant MFA for Admins' = 'https://learn.microsoft.com/en-us/entra/identity/conditional-access/policy-admin-phish-resistant-mfa' }, $false, $true),
    [recommendation]::new('CA-06', 'Policy Excludes Entities That Are Also Included', 'There is at least one policy that excludes the same entities it includes, resulting in no effective condition being checked.', 'Review and update policies to ensure that they do not exclude the same entities they include.', 'Policies should include and exclude distinct sets of entities to ensure conditions are effectively checked.', @{ 'Policy Configuration Best Practices' = 'https://learn.microsoft.com/en-us/azure/active-directory/conditional-access/best-practices' }, $true, $false),
    [recommendation]::new('CA-07', 'No Users Targeted in Policy', 'All policies are scoped to users, groups, or roles.', 'There is at least one policy that does not target any users. Review and update policies to ensure that they target specific users, groups, or roles to be effective.', 'Policies should target specific users, groups, or roles to ensure they apply correctly.', @{ 'Policy Configuration Best Practices' = 'https://learn.microsoft.com/en-us/azure/active-directory/conditional-access/best-practices' }, $true, $false),
    [recommendation]::new('CA-08', 'Direct User Assignment', 'There are no direct user assignments in the policy.', 'Review and update policies to avoid direct user assignments; prefer groups.', 'Direct user assignments reduce scalability; use exclusion/target groups instead.', @{}, $true, $false),
    [recommendation]::new('CA-09', 'Implement Risk-Based Policy', 'There is at least 1 policy that addresses risk-based conditional access.', 'Consider implementing risk-based conditional access policies for dynamic controls.', 'Risk-based policies assess risk of sign-ins/users and apply appropriate controls.', @{ 'Risk-Based Conditional Access Overview' = 'https://learn.microsoft.com/en-us/entra/id-protection/howto-identity-protection-configure-risk-policies'; 'Require MFA for Risky Sign-in' = 'https://learn.microsoft.com/en-us/entra/identity/conditional-access/policy-risk-based-sign-in#enable-with-conditional-access-policy'; 'Require Passsword Change for Risky User' = 'https://learn.microsoft.com/en-us/entra/identity/conditional-access/policy-risk-based-user#enable-with-conditional-access-policy' }, $false, $true),
    [recommendation]::new('CA-10', 'Block Device Code Flow', 'There is at least 1 policy that blocks device code flow.', 'Consider implementing a policy to block device code flow.', 'Blocking device code flow prevents potential abuse of device code auth.', @{ 'Block Device Code Flow Overview' = 'https://learn.microsoft.com/en-us/entra/identity/conditional-access/concept-authentication-flows#device-code-flow' }, $false, $true),
    [recommendation]::new('CA-11', 'Require MFA to Enroll a Device in Intune', 'There is at least 1 policy that requires Multi-Factor Authentication (MFA) to enroll a device in Intune.', 'Consider implementing a policy to require MFA for Intune enrollment.', 'MFA for enrollment ensures only authorized users can enroll devices.', @{ 'MFA for Intune Enrollment Overview' = 'https://learn.microsoft.com/en-us/mem/intune/enrollment/multi-factor-authentication' }, $false, $true),
    [recommendation]::new('CA-12', 'Block Unknown/Unsupported Devices', 'There is no policy that blocks unknown or unsupported devices.', 'Implement a policy to block unknown or unsupported devices.', 'Blocking unknown or unsupported devices prevents access from non-compliant endpoints.', @{ 'Block Unknown/Unsupported Devices Overview' = 'https://learn.microsoft.com/en-us/entra/identity/conditional-access/policy-all-users-device-unknown-unsupported' }, $false, $true)
  )

  function Test-PolicyStatus {
    param (
      [ref]$Recommendation,
      $PolicyCheck,
      $StatusCheck
    )

    if (&$StatusCheck $PolicyCheck) {
      $Recommendation.Value.Status = $Recommendation.Value.SwapStatus
    }

    if ($Recommendation.Value.Status) {
      # Recommendation passed. Differentiate enabled vs reporting-only vs disabled.
      if ($PolicyCheck.state -eq 'enabledForReportingButNotEnforced') {
        $Status1 = 'policy-item success success-report'
      }
      elseif ($PolicyCheck.state -eq 'enabled') {
        $Status1 = 'policy-item success success-enabled'
      }
      elseif ($PolicyCheck.state -eq 'disabled') {
        $Status1 = 'policy-item success success-disabled'
      }
      else {
        # Other states considered success but generic (retain base success styling)
        $Status1 = 'policy-item success'
      }
      $Status2 = 'status-icon-large success'
      $Status3 = '✔'
    }
    else {
      # Recommendation failed (policy needs attention). Differentiate enabled vs reporting-only vs disabled policies.
      if ($PolicyCheck.state -eq 'enabledForReportingButNotEnforced') {
        # Reporting-only: distinct background per requirements
        $Status1 = 'policy-item warning warning-report'
      }
      elseif ($PolicyCheck.state -eq 'enabled') {
        # Enabled & failing: keep existing warning style (optional class for clarity)
        $Status1 = 'policy-item warning warning-enabled'
      }
      elseif ($PolicyCheck.state -eq 'disabled') {
        # Disabled & failing: distinct warning-disabled variant
        $Status1 = 'policy-item warning warning-disabled'
      }
      else {
        # Generic fallback (other states)
        $Status1 = 'policy-item warning'
      }
      $Status2 = 'status-icon-large warning'
      $Status3 = '⚠'
    }

    $CheckExcUG = $PolicyCheck.Conditions.Users.ExcludeUsers + $PolicyCheck.Conditions.Users.ExcludeGroups + $PolicyCheck.Conditions.Users.ExcludeRoles + $PolicyCheck.conditions.users.ExcludeGuestsOrExternalUsers.GuestOrExternalUserTypes -replace ',', ', '
    $CheckIncUG = $PolicyCheck.Conditions.Users.IncludeUsers + $PolicyCheck.Conditions.Users.IncludeGroups + $PolicyCheck.Conditions.Users.IncludeRoles + $PolicyCheck.conditions.users.IncludeGuestsOrExternalUsers.GuestOrExternalUserTypes -replace ',', ', '
    $CheckIncCond = $PolicyCheck.Conditions.Locations.includelocations + $PolicyCheck.Conditions.Platforms.IncludePlatforms
    $CheckExcCond = $PolicyCheck.Conditions.Locations.Excludelocations + $PolicyCheck.Conditions.Platforms.ExcludePlatforms
    $CheckGrant = $PolicyCheck.GrantControls.BuiltInControls + $PolicyCheck.GrantControls.AuthenticationStrength.DisplayName + $PolicyCheck.GrantControls.CustomAuthenticationFactors + $PolicyCheck.GrantControls.TermsOfUse
    $checkSession = ''

    if ($PolicyCheck.SessionControls.ApplicationEnforcedRestrictions.IsEnabled) {
      $checkSession += "    ApplicationEnforcedRestrictions: $($PolicyCheck.SessionControls.ApplicationEnforcedRestrictions.IsEnabled)`n"
    }
    if ($PolicyCheck.SessionControls.CloudAppSecurity.IsEnabled) {
      $checkSession += "    CloudAppSecurity: $($PolicyCheck.SessionControls.CloudAppSecurity.IsEnabled)`n"
    }
    if ($PolicyCheck.SessionControls.SignInFrequency.Value -and $PolicyCheck.SessionControls.SignInFrequency.Type) {
      $checkSession += "    SignInFrequency: $($PolicyCheck.SessionControls.SignInFrequency.Value) $($PolicyCheck.SessionControls.SignInFrequency.Type)`n"
    }
    if ($PolicyCheck.SessionControls.PersistentBrowser.Mode) {
      $checkSession += "    PersistentBrowser: $($PolicyCheck.SessionControls.PersistentBrowser.Mode)`n"
    }
    if ($PolicyCheck.SessionControls.ContinuousAccessEvaluation.Mode) {
      $checkSession += "    ContinuousAccessEvaluation: $($PolicyCheck.SessionControls.ContinuousAccessEvaluation.Mode)`n"
    }
    if ($PolicyCheck.SessionControls.DisableResilienceDefaults) {
      $checkSession += "    ResiliantDefaults: $($PolicyCheck.SessionControls.DisableResilienceDefaults)`n"
    }
    if ($PolicyCheck.SessionControls.AdditionalProperties.secureSignInSession.Values) {
      $checkSession += "    secureSignInSession: $($PolicyCheck.SessionControls.AdditionalProperties.secureSignInSession.Values)`n"
    }

    if (&$StatusCheck $PolicyCheck) {
      $Recommendation.Value.Note += "
    <div class='policy'>
        <div class='$($Status1)'>
            <div class='policy-header'>
                <strong>$($PolicyCheck.DisplayName) <a href='$($LinkURL)$($PolicyCheck.Id)' target='_blank'><span class='icon-ext'></span></a></strong>
                <div class='recommendation-status'>Status: $($PolicyCheck.state)</div>
                <div class='$($Status2)'>$Status3</div>
            </div>
            <div class='policy-content'>
                <div class='policy-include'>
                    <div class='label-container'>
                        <span class='include-label'>Include</span>
                    </div>
                    <div class='include-content'>
                        <b>Users:</b> $($CheckIncUG -join ', ')
                        <br>
                        <b>Application/Actions:</b> $($PolicyCheck.Conditions.Applications.IncludeApplications -join ', ') $($PolicyCheck.Conditions.Applications.IncludeUserActions -join ', ')
                        <br>
                        <b>Conditions:</b> $($CheckIncCond -join ', ')
                    </div>
                </div>
                <div class='policy-exclude'>
                    <div class='label-container'>
                        <span class='exclude-label'>Exclude</span>
                    </div>
                    <div class='exclude-content'>
                        <b> Users:</b> $($CheckExcUG -join ', ')
                        <br>
                        <b>Applications:</b> $($PolicyCheck.Conditions.Applications.ExcludeApplications -join ', ')
                        <br>
                        <b>Conditions:</b> $($CheckExcCond -join ', ')
                    </div>
                </div>
                 <div class='policy-grant'>
                    <div class='label-container'>
                        <span class='grant-label'>Access</span>
                    </div>
                    <div class='grant-content'>
                        <b> Grant:</b> $($CheckGrant  -join ', ') :$($PolicyCheck.GrantControls.Operator)
                        <br>
                        <b> Session:</b> $($CheckSession  -join ', ')
                    </div>
                </div>
            </div>
        </div>
    </div>"
    }
  }


  $CheckFunctions = @{
    'CA-00' = { param($p) Test-CA00 $p }
    'CA-01' = { param($p) Test-CA01 $p }
    'CA-02' = { param($p) Test-CA02 $p }
    'CA-03' = { param($p) Test-CA03 $p }
    'CA-04' = { param($p) Test-PolicyRequiresMfaForAdmins -Policy $p }
    'CA-05' = { param($p) Test-PolicyRequiresPhishResistantMfaForAdmins -Policy $p }
    'CA-06' = { param($p) Test-CA06 $p }
    'CA-07' = { param($p) Test-CA07 $p }
    'CA-08' = { param($p) Test-CA08 $p }
    'CA-09' = { param($p) Test-CA09 $p }
    'CA-10' = { param($p) Test-CA10 $p }
    'CA-11' = { param($p) Test-CA11 $p }
    'CA-12' = { param($p) Test-CA12 $p }
  }


  foreach ($policy in $CAPolicy) {
    foreach ($recommendation in $recommendations) {
      Test-PolicyStatus -Recommendation ([ref]$recommendation) -PolicyCheck $policy -StatusCheck $CheckFunctions[$recommendation.Control]
    }
  }
}

## (Removed obsolete Set Row Order block)

# ---------------- Build Pivot Dataset (wide) ----------------
# The pivot format rotates policy-centric data so that each policy becomes a column and each attribute a row.
# This allows quick scanning in Excel / BI of where controls/conditions differ across policies.
# Implementation notes:
#  - Use a stable ordered list of pivot fields ($pivotFields)
#  - Each row object has a 'CA Item' property (row label) plus one property per policy (policy display name)
#  - Boolean-like values preserved (✔ for True style fields) while multi-line fields condensed to single line for readability

$pivot = @()
$pivotFields = @(
  #  'Recommended Name', 'Modified', 'Created', 'Description',
  #   'Included Users', 'Excluded Users',
  #   'Included Applications', 'Excluded Applications', 'User Actions', 'Auth Context', 'Included Locations', 'Excluded Locations',
  #   'User Risk', 'SignIn Risk', 'Platforms Included', 'Platforms Excluded', 'Included Devices', 'Excluded Devices', 'Client Apps','Device Filters', 'Authentication Flows',
  #   'Block', 'Require MFA', 'Authentication Strength MFA', 'Compliant Device', 'Domain Joined Device', 'Compliant Application', 'Approved Application', 'Password Change', 'Terms Of Use', 'Custom Controls', 'Grant Operator',
  #   'Application Enforced Restrictions', 'Cloud App Security', 'Sign In Frequency', 'Persistent Browser', 'Continuous Access Evaluation', 'Resilient Defaults', 'Secure Sign In Session'
  'Recommended Name', 'Modified', 'Created', 'Description',
  'Included Users', 'Excluded Users',
  'Included Applications', 'Excluded Applications', 'User Actions', 'Auth Context', 'Included Locations', 'Excluded Locations', 'User Risk', 'SignIn Risk', <#'Insider Risk',#> 'Platforms Included', 'Platforms Excluded', 'Included Devices', 'Excluded Devices', 'Client Apps', 'Device Filters', 'Authentication Flows', 'Block', 'Grant Operator', 'Require MFA', 'Authentication Strength MFA', 'Compliant Device', 'Domain Joined Device', 'Compliant Application', 'Approved Application', 'Password Change', 'Terms Of Use', 'Custom Controls', 'Application Enforced Restrictions', 'Cloud App Security', 'Sign In Frequency', 'Persistent Browser', 'Continuous Access Evaluation', 'Resilient Defaults', 'Secure Sign In Session', 'RawJson'
)

if ($CAExport.Count -gt 0) {
  foreach ($field in $pivotFields) {
    $row = [ordered]@{ 'CA Item' = $field }
    foreach ($pol in $CAExport) {
      $val = $null
      if ($pol.PSObject.Properties.Match($field)) {
        $val = $pol.$field 
      }
      # Normalize multiline -> semi-colon separated single line for sheet friendliness
      if ($val -is [string]) {
        $val = ($val -split "`r?`n") -join '; ' 
      }
      $row[$pol.Name] = $val
    }
    $pivot += [pscustomobject]$row
  }
}
function Get-RecommendationsHtmlFragment {
  param (
    [Parameter(Mandatory = $true)]
    [object[]]$Recommendations
  )
  # Redesigned compact recommendation cards using <details> for collapsible body.
  $htmlFragment = @'
<div class='recommendations' id='ca-security-checks' style=''>
'@
  foreach ($rec in $Recommendations) {
    # Replace any GUIDs in the note with friendly names then sanitize curated HTML
    $rec.Note = Protect-RecNote -Html (Resolve-EntityGuidsInText -Text $rec.Note -UserMap $UserMap -GroupMap $GroupMap -RoleMap $RoleMap -AppMap $AppMap)
    $links = ''
    foreach ($key in $rec.Links.Keys) {
      $links += "<a class='rec-link' href='$($rec.Links[$key])' target='_blank' rel='noopener'>$key<span class='icon-ext'></span></a>" 
    }
    if ($rec.Status) {
      $RecStatus = 'pass'
      $RecStatusNote = [System.Web.HttpUtility]::HtmlEncode($rec.PassText)
      $Icon = '✔'
      $StateLabel = 'Pass'
      $detailOpen = ''
    }
    else {
      $RecStatus = 'fail'
      $RecStatusNote = [System.Web.HttpUtility]::HtmlEncode($rec.FailRecommendation)
      $Icon = '⚠'
      $StateLabel = 'Attention'
      $detailOpen = ' open'
    }
    $encName = [System.Web.HttpUtility]::HtmlEncode($rec.Name)
    $encCtrl = [System.Web.HttpUtility]::HtmlEncode($rec.Control)
    $importance = [System.Web.HttpUtility]::HtmlEncode($rec.Importance)
    # Policy note already sanitized via Sanitize-RecNote (retains limited safe markup)
    $rawNote = $rec.Note
    # If this is CA-06 (overlapping include/exclude) attempt to enrich note with explicit per-policy overlap summary
    if ($rec.Control -eq 'CA-06') {
      # We'll look for pattern of policy lines already in $rawNote and append Overlaps section
      # Attempt to reconstruct overlaps from exported data if available ($CAExport)
      $overlapDetails = @()
      foreach ($pol in $CAExport) {
        # Recreate raw include/exclude lists similar to table columns
        $incUsers = @(); $excUsers = @(); $incRoles = @(); $excRoles = @()
        if ($pol.PSObject.Properties.Match('Included Users')) { $incUsers = @($pol.'Included Users' -split '[,\n]') }
        if ($pol.PSObject.Properties.Match('Excluded Users')) { $excUsers = @($pol.'Excluded Users' -split '[,\n]') }

        # Groups currently surfaced via Included Users column when resolved; skip unless future columns added
        function Get-OverlapTokens { param($A, $B) ($A | ForEach-Object { $_.Trim() } | Where-Object { $_ -and ($B -contains $_.Trim()) }) | Sort-Object -Unique }
        $userOverlap = Get-OverlapTokens $incUsers $excUsers
        $roleOverlap = Get-OverlapTokens $incRoles $excRoles
        if ($userOverlap.Count -gt 0 -or $roleOverlap.Count -gt 0) {
          $pn = [System.Web.HttpUtility]::HtmlEncode($pol.Name)
          $parts = @()
          if ($userOverlap.Count -gt 0) { $parts += ('Users: ' + ($userOverlap | ForEach-Object { [System.Web.HttpUtility]::HtmlEncode($_) }) -join ', ') }
          if ($roleOverlap.Count -gt 0) { $parts += ('Roles: ' + ($roleOverlap | ForEach-Object { [System.Web.HttpUtility]::HtmlEncode($_) }) -join ', ') }
          $overlapDetails += "<div class='overlap-line'><span class='overlap-policy-name'>$pn</span> — " + ($parts -join ' | ') + '</div>'
        }
      }
      if ($overlapDetails.Count -gt 0) {
        $overlapBlock = "<div class='overlap-summary'><div class='overlap-heading'>Overlapping Targets Detected:</div>" + ($overlapDetails -join '') + '</div>'
        $rawNote = $rawNote + $overlapBlock
      }
    }
    $encNote = $rawNote
    $htmlFragment += @"
<details class='recommendation-card $RecStatus'$detailOpen>
  <summary><span class='rec-status-icon' aria-label='$StateLabel'>$Icon</span><span class='rec-code'>$encCtrl</span><span class='rec-title'>$encName</span></summary>
  <div class='rec-body'>
    <div class='rec-importance'>$importance</div>
    <div class='rec-status-text'>$RecStatusNote</div>
    <div class='rec-links'>$links</div>
  <div class='rec-matched-policies'>$encNote</div>
  </div>
</details>
"@
  }
  $htmlFragment += "<div class='timestamp-note'>Report generated: $([System.Web.HttpUtility]::HtmlEncode((Get-Date).ToString('u')))</div></div>"
  return $htmlFragment
}

# If HTML export requested, generate full HTML report with embedded CSS/JS and write to file
if ($HTMLExport) {
  if (-not $NoRecommendations) {
    # Replace the legacy recommendations root div id while preserving original class and avoid overriding base CSS padding.
    # Inline padding previously here caused mismatch with stylesheet; rely on unified .recommendations styling instead.
    $SecurityCheck = (Get-RecommendationsHtmlFragment -Recommendations $recommendations) -replace "id='ca-security-checks' style=''", "id='panel-recommendations' role='tabpanel' aria-labelledby='tab-recommendations' aria-hidden='true' style='display:none'"
  }
  else {
    $SecurityCheck = ''
  }
  $recTabs = if ($NoRecommendations) {
    "<span id='tab-summary' class='btn-toggle' role='tab' aria-selected='false' tabindex='-1' aria-controls='panel-summary'>Summary</span><span id='tab-policies' class='btn-toggle active' role='tab' aria-selected='true' tabindex='0' aria-controls='panel-policies'>Policy Details</span>"
  }
  else {
    "<span id='tab-summary' class='btn-toggle' role='tab' aria-selected='false' tabindex='-1' aria-controls='panel-summary'>Summary</span><span id='tab-policies' class='btn-toggle active' role='tab' aria-selected='true' tabindex='0' aria-controls='panel-policies'>Policy Details</span><span id='tab-recommendations' class='btn-toggle' role='tab' aria-selected='false' tabindex='-1' aria-controls='panel-recommendations'>Recommendations</span>"
  }
  $OmissionBannerHtml = if ($NoRecommendations) {
    "<div class='no-recs-banner' role='note'>Recommendations omitted (-NoRecommendations)</div>" 
  }
  else {
    '' 
  }
  Write-Info 'Saving to File: HTML'
  # Self-contained CSS (no external dependencies)
  $style = @'
  /* General Styles */
  html, body { font-family: Arial, sans-serif; margin:0; padding:0; }
  .title { font-size: 1.2em; font-weight: bold; }
  /* Navigation */
  .navbar-custom { position:fixed; top:0; left:0; right:0; display:flex; align-items:center; justify-content:space-between; background:#005494; color:#fff; padding:14px 18px; box-shadow:0 2px 4px rgba(0,0,0,.25); z-index:999; font-size:14px; }
  .no-recs-banner { margin:70px 16px 10px 18px; background:#fff4cc; border:1px solid #e0c766; padding:10px 14px; border-radius:5px; font-size:0.75rem; color:#5a4700; box-shadow:0 1px 2px rgba(0,0,0,.05); }
  /* Offset main tab panels so content isn't hidden under fixed navbar */
  /* Panel offsets: recommendations needs more vertical offset due to heading density; others can be tighter */
  #panel-recommendations { padding-top:68px; margin-top:0; scroll-margin-top:80px; }
  #panel-summary, #panel-policies { padding:46px 5px; margin-top:0; scroll-margin-top:60px; }
  /* Fine-tune summary since its internal section already has top margin */
  #panel-summary .summary-wrapper { margin-top:10px !important; padding-left:5px; }
  .nav-left, .nav-center, .nav-right { display:flex; align-items:center; }
  .nav-center { flex:1; justify-content:center; font-weight:600; }
  .nav-left .brand { font-weight:700; margin-left:8px; }
  .nav-right { gap:12px; font-size:10px; }
  .icon-server { font-size:18px; line-height:1; }
  .view-toggle-group .btn-toggle { color:#fff; background:#0d6efd33; border:1px solid rgba(255,255,255,0.4); padding:4px 10px; margin-right:6px; border-radius:4px; cursor:pointer; font-size:0.7rem; user-select:none; }
  .view-toggle-group .btn-toggle.active { background:#ffffff; color:#005494; font-weight:600; box-shadow:0 0 0 2px #ffffff55; }
  .view-toggle-group .btn-toggle:focus { outline:none; }
  .search-box { position:relative; margin-left:12px; }
  .search-box input { padding:4px 26px 4px 8px; border-radius:4px; border:1px solid #fff; background:#ffffff; color:#003553; font-size:0.6rem; min-width:190px; }
  .search-box input:focus { outline:2px solid #91d2ff; }
  .search-box .search-clear { position:absolute; right:6px; top:50%; transform:translateY(-50%); cursor:pointer; color:#005494; font-weight:bold; display:none; }
  .search-box.has-value .search-clear { display:inline; }
  .status-filter { margin-left:12px; }
  .status-filter select { padding:4px 8px; border-radius:4px; border:1px solid #fff; background:#ffffff; color:#003553; font-size:0.6rem; cursor:pointer; }
  .status-filter select:focus { outline:2px solid #91d2ff; }
  /* Table */
  table { border-collapse: collapse; margin-bottom:30px; margin-top:55px; font-size:0.8em; min-width:400px; width:100%; table-layout:auto; }
  thead tr { background:linear-gradient(90deg,#005494,#0a79c5); color:#ffffff; text-align:center; }
  th, td { padding:6px 6px; border:1px solid #d2d2d2; vertical-align:top; text-align:center; }
  /* Dynamic width adjustments: allow natural content sizing, but keep some guidance */
  th.name-col, td.name-col { max-width:300px; }
  th.bool-col, td.bool-col { width:46px; min-width:46px; }
  th.group-assign-users, td.group-assign-users { max-width:150px !important; white-space:normal; word-wrap:break-word; overflow-wrap:break-word; }
  tbody tr:nth-of-type(even) { background-color:#f3f3f3; }
  tbody tr:last-of-type { border-bottom:2px solid #005494; }
  tr:hover { background-color:#d8d8d8 !important; }
  .selected:not(th) { background-color:#afe1ff !important; }
  /* Improved header readability + distinct sticky first column header */
  th { background-color:#005494; color:#ffffff; font-weight:600; font-size:0.9rem; letter-spacing:.3px; border-bottom:1px solid #00416d; border-top:0px; }
  th.sticky-name { background:#004d7f; box-shadow:4px 0 4px -4px rgba(0,0,0,.35); }
  .sticky-name { position:sticky; left:0; inset-inline-start:0; background:#005494; color:#fff; z-index:5; font-weight:700; box-shadow:4px 0 6px -4px rgba(0,0,0,.35); }
  th.sticky-name { text-align:center; }
  td.sticky-name { padding:0; text-align:left; }
  td.sticky-name .status-label { position:absolute; left:0; top:0; bottom:0; width:24px; }
  td.sticky-name .name-content { padding:8px 8px 8px 32px; display:block; }
  /* Sticky header (below navbar ~55px high) - only column headers, not group headers */
  .sticky-header thead tr.group-header th { position:static; }
  .sticky-header thead tr.column-header-row th { position:sticky; top:55px; z-index:6; }
  .sticky-header thead tr.column-header-row th.sticky-name { z-index:7; }
  /* Grouped header color palette */
  .sticky-header thead tr.group-header th.group-span { font-size:0.65rem; letter-spacing:.4px; text-transform:uppercase; color:#fff; border-bottom:0px; }
  .sticky-header thead tr.column-header-row th { font-size:0.60rem; }
  .group-general { background:#204b73; }
  .group-assign-users { background:#0d5c63; }
  .group-assign-target { background:#5a3d7a; }
  .group-network { background:#7a4b2d; }
  .group-conditions { background:#2d6e5a; }
  .group-grant { background:#735e0d; }
  .group-session { background:#5a1f3d; }
  .group-output { background:#444c57; }
  /* Match second row colors by inheriting group classes */
  .column-header-row th.group-general { background:#295e91; }
  .column-header-row th.group-assign-users { background:#117279; }
  .column-header-row th.group-assign-target { background:#6c4993; }
  .column-header-row th.group-network { background:#945c39; }
  .column-header-row th.group-conditions { background:#378a70; }
  .column-header-row th.group-grant { background:#8e7412; }
  .column-header-row th.group-session { background:#722648; }
  .column-header-row th.group-output { background:#556170; }
  .group-header th { position:sticky; top:55px; z-index:8; }
  .column-header-row th { position:sticky; top:86px; z-index:7; }
  /* Adjust sticky first column under new two-row header stack */
  .sticky-header thead th.sticky-name { top:86px; }
  .sticky-name a { color:#fff; }
  .sticky-name .policy-name { color:#fff; }
  .sticky-name .policy-link { display:inline-flex; align-items:center; justify-content:center; margin-left:4px; text-decoration:none; width:20px; height:20px; border-radius:4px; background:rgba(255,255,255,0.12); transition:background .18s ease, transform .18s ease; }
  .sticky-name .policy-link:focus { outline:2px solid #ffcc33; outline-offset:2px; }
  .sticky-name .policy-link:hover, .sticky-name .policy-link:focus-visible { background:#ffcc33; }
  .sticky-name .policy-link:hover svg.icon-ext, .sticky-name .policy-link:focus-visible svg.icon-ext { color:#003553; transform:scale(1.12); }
  .sticky-name .policy-link svg.icon-ext { width:14px; height:14px; stroke:currentColor; stroke-width:1.9; fill:none; color:#fff; transition:color .18s ease, transform .18s ease; }
  .recommendation-card .overlap-summary { margin-top:8px; padding:6px 8px; background:#fff8f0; border:1px solid #f2c08f; border-radius:4px; font-size:0.62rem; line-height:1.25; }
  .recommendation-card .overlap-heading { font-weight:700; margin-bottom:4px; color:#8c4f00; }
  .recommendation-card .overlap-line { margin:2px 0; }
  .recommendation-card .overlap-policy-name { font-weight:600; }
  tbody tr:nth-of-type(even) .sticky-name { background:#547c9b; }
/*  tbody tr:nth-of-type(5), tbody tr:nth-of-type(8), tbody tr:nth-of-type(13), tbody tr:nth-of-type(25), tbody tr:nth-of-type(37) { background-color:#005494 !important; }*/
  .tooltip-container { position:relative; display:inline-block; }
  .tooltip-text { visibility:hidden; width:200px; background:#000; color:#fff; text-align:center; border-radius:6px; padding:5px 0; position:absolute; z-index:1; top:115%; left:50%; margin-left:-100px; opacity:0; transition:opacity .3s; }
  .tooltip-container:hover .tooltip-text { visibility:visible; opacity:1; }
  /* Recommendations - compact cards */
  /* Base recommendation panel styling applied via class; id-specific rules no longer required */
  .recommendations { padding:20px 18px 40px 18px; background:#f5f7fa; border:1px solid #d0d7de; border-radius:6px; margin-top:55px; }
  details.recommendation-card { border:1px solid #d8dee4; border-left:5px solid #888; border-radius:4px; padding:4px 10px 6px 10px; margin:0 0 8px 0; background:#ffffff; box-shadow:0 1px 2px rgba(0,0,0,.04); }
  details.recommendation-card[open] { box-shadow:0 2px 4px rgba(0,0,0,.07); }
  details.recommendation-card.pass { border-left-color:#218739; }
  details.recommendation-card.fail { border-left-color:#d37d00; }
  details.recommendation-card summary { list-style:none; cursor:pointer; display:flex; align-items:center; gap:8px; font-weight:600; font-size:0.82rem; }
  details.recommendation-card summary::-webkit-details-marker { display:none; }
  .rec-status-icon { font-size:0.85rem; width:18px; text-align:center; }
  details.recommendation-card.pass .rec-status-icon { color:#218739; }
  details.recommendation-card.fail .rec-status-icon { color:#d37d00; }
  .rec-code { font-family:ui-monospace,SFMono-Regular,Menlo,Monaco,Consolas,'Liberation Mono','Courier New',monospace; background:#eef2f5; padding:2px 5px; border-radius:3px; font-size:0.68rem; color:#303b44; }
  details.recommendation-card.fail .rec-code { background:#ffe8cc; }
  .rec-title { flex:1; }
  .rec-body { margin:6px 2px 2px 2px; font-size:0.72rem; line-height:1.25; }
  .rec-importance { color:#334; margin-bottom:4px; }
  .rec-status-text { margin:4px 0 4px 0; font-weight:500; }
  .rec-links { margin:4px 0 4px 0; display:flex; flex-wrap:wrap; gap:6px; }
  .rec-link { font-size:0.66rem; background:#e6f2fb; padding:3px 6px; border-radius:3px; text-decoration:none; color:#005494; border:1px solid #c6e0f2; }
  .rec-link:hover { background:#d2e8f8; }
  .rec-matched-policies .policy-item { font-size:0.66rem; }
  .rec-matched-policies .policy-header { font-size:0.65rem; }
  .policy { margin-top:10px; }
  .policy-item { border:2px solid; padding:10px; border-radius:5px; margin-bottom:10px; }
  .policy-item.success { border-color:green; background:#e6ffe6; }
  .policy-item.success-report { border-color:green; background:#C9DFC9; }
  .policy-item.success-enabled { border-color:green; background:#e6ffe6; }
  .policy-item.success-disabled { border-color:green; background:#909F90; }
  .policy-item.warning { border-color:orange; background:#fff8e6; }
  /* Failing policies that are reporting-only (enabledForReportingButNotEnforced) */
  .policy-item.warning-report { border-color:orange; background:#e6dfcf; }
  /* Optional semantic alias for enabled failing policies; inherits base warning background */
  .policy-item.warning-enabled { border-color:orange; background:#fff8e6; }
  .policy-item.warning-disabled { border-color:orange; background:#b3aea1; }
  .policy-item.error { border-color:red; background:#ffe6e6; }
  .policy-content { display:flex; flex-direction:column; padding-left:20px; margin-top:5px; }
  .policy-include, .policy-exclude, .policy-grant { display:flex; align-items:flex-start; margin-top:5px; }
  .label-container { display:flex; align-items:center; margin-right:10px; }
  .include-label, .exclude-label, .grant-label { writing-mode:vertical-rl; transform:rotate(180deg); border-left:3px solid darkgrey; color:darkgray; }
  .status-label { writing-mode:vertical-rl; transform:rotate(180deg); font-size:0.65rem; font-weight:600; padding:0; margin:0; height:100%; width:100%; display:flex; align-items:center; justify-content:center; text-transform:uppercase; letter-spacing:0.5px; }
  .status-label.status-enabled { background:#d4edda; color:#155724; border-left:3px solid #28a745; }
  .status-label.status-report { background:#ffe0b3; color:#856404; border-left:3px solid #ff9800; }
  .status-label.status-disabled { background:#f8d7da; color:#721c24; border-left:3px solid #dc3545; }
  .status-icon-large { position:absolute; top:0; right:0; font-size:2em; }
  .status-icon-large.success { color:green; }
  .status-icon-large.warning { color:orange; }
  .status-icon-large.error { color:red; }
  .icon-ext { font-size:0.75em; margin-left:4px; color:#000; }
  .selected td { background:#afe1ff; }
  #back-to-top { position:fixed; right:18px; bottom:18px; background:#005494; color:#fff; border:none; padding:10px 14px; border-radius:50%; font-size:14px; cursor:pointer; box-shadow:0 2px 6px rgba(0,0,0,.3); display:none; z-index:998; }
  #back-to-top:hover { background:#0073c7; }
  .timestamp-note { font-size:0.7rem; color:#555; margin:6px 0 14px 0; }
  /* Raw JSON details */
  details.raw-json { max-width:400px; }
  /* Summary table styling (separate class so not counted as a policy data table) */
  /* Compact summary table styling */
  .summary-wrapper { max-width:760px; margin:0 0 25px 0; }
  table.summary-table { border-collapse:separate; border-spacing:0; width:100%; font-size:0.70rem; line-height:1.05; margin-bottom:18px; box-shadow:0 0 0 1px #d0d7de; }
  table.summary-table th, table.summary-table td { border:0; padding:3px 6px; background:#fff; white-space:nowrap; }
  table.summary-table thead th { position:sticky; top:0; background:#0d3855; color:#fff; font-weight:600; font-size:0.69rem; letter-spacing:.5px; text-transform:uppercase; }
  /* Striped rows for summary table */
  table.summary-table tbody tr:nth-child(odd)  { background:#ffffff; }
  table.summary-table tbody tr:nth-child(even) { background:#f1f5f8; }
  table.summary-table tbody tr:hover { background:#e2edf3; }
  table.summary-table tbody td:first-child { font-weight:500; }
  table.summary-table td:nth-child(2), table.summary-table td:nth-child(3) { text-align:right; }
  /* Create a subtle column separation */
  table.summary-table td + td, table.summary-table th + th { border-left:1px solid #e2e8ec; }
  /* Allow wrapping on longer metric labels */
  table.summary-table td:first-child { white-space:normal; }
  @media (min-width:900px){
    /* Present metrics in two side-by-side columns (two tables) if desired later; placeholder for responsive enhancements */
  }
  details.raw-json summary { cursor:pointer; font-weight:600; color:#005494; }
  details.raw-json pre { max-height:300px; overflow:auto; background:#0f1f2a; color:#d7f1ff; padding:8px; border-radius:4px; font-size:0.7rem; }
  .json-toggle-bar { display:flex; gap:6px; margin-bottom:6px; }
  .json-toggle-bar button { background:#194e73; color:#fff; border:1px solid #0d3956; padding:3px 8px; font-size:0.65rem; cursor:pointer; border-radius:4px; }
  .json-toggle-bar button.active { background:#ffcc33; color:#003553; font-weight:600; }

  /* Boolean / placeholder & status styling */
  .bool-yes { color:#1f7a33; font-weight:600; }
  .bool-no { color:#c4c4c4; font-weight:400; }
  .placeholder { color:#b9b9b9; font-style:italic; }
  .pill-block { display:inline-block; padding:4px 10px; border-radius:12px; font-size:0.7rem; font-weight:700; letter-spacing:0.5px; background:#dc3545; color:#fff; border:1px solid #bd2130; }
  .pill-grant { display:inline-block; padding:4px 10px; border-radius:12px; font-size:0.7rem; font-weight:700; letter-spacing:0.5px; background:#28a745; color:#fff; border:1px solid #1e7e34; }
  .pill-recommended { display:inline-block; padding:2px 6px; margin-left:6px; border-radius:8px; font-size:0.6rem; font-weight:600; letter-spacing:0.3px; background:#e3f2fd; color:#1565c0; border:1px solid #90caf9; }
  th.bool-col, td.bool-col { min-width:46px; width:46px; padding:6px 4px; font-size:0.65rem; }
  td.bool-col span { display:inline-block; min-width:16px; }
  .status-badge { display:inline-block; padding:2px 6px; border-radius:10px; font-size:0.62rem; font-weight:600; letter-spacing:.3px; }
  .status-enabled { background:#daf5d9; color:#1f7a33; border:1px solid #b2e3b0; }
  .status-disabled { background:#f2dede; color:#b94a48; border:1px solid #e0b4b3; }
  .status-report { background:#fff4cc; color:#8c6d00; border:1px solid #f2dd8f; }
  td:has(details.raw-json) { min-width:240px; }
  /* Legend & utility bars */
  .legend-bar { margin:25px 0 12px 0; background:#ffffff; border:1px solid #d8dee4; border-left:4px solid #005494; padding:8px 12px; border-radius:4px; font-size:0.68rem; display:flex; flex-wrap:wrap; gap:10px; align-items:center; }
  .layout-toggle { background:#ffffff; color:#005494; border:1px solid #c3d1dc; padding:4px 8px; font-size:0.65rem; border-radius:4px; cursor:pointer; }
  .layout-toggle.active { background:#005494; color:#fff; }
  .legend-title { font-weight:700; margin-right:4px; }
  .legend-item { display:flex; align-items:center; gap:4px; }
  .legend-swatch { display:inline-flex; align-items:center; justify-content:center; min-width:18px; height:18px; font-size:0.65rem; border-radius:4px; border:1px solid #c9d1d9; background:#f3f4f6; color:#333; }
  .legend-swatch.enabled { background:#daf5d9; border-color:#b2e3b0; }
  .legend-swatch.report { background:#fff4cc; border-color:#f2dd8f; }
  .legend-swatch.disabled { background:#f2dede; border-color:#e0b4b3; }
  .legend-swatch.pass { background:#d9f2e3; border-color:#b4e2c4; color:#1f7a33; }
  .legend-swatch.fail { background:#ffe8cc; border-color:#f2c08f; color:#8c5a00; }
  .legend-divider { width:1px; height:18px; background:#d0d7de; }
  /* Legend chip buttons (interactive filters) */
  .legend-chip { background:#edf3fc; color:#194b7d; border:1px solid #b8d0ef; padding:2px 10px; font-size:0.62rem; border-radius:999px; cursor:pointer; letter-spacing:.3px; font-weight:600; box-shadow:0 0 0 1px rgba(255,255,255,0.4) inset; transition:background .15s ease, color .15s ease; }
  .legend-chip:hover, .legend-chip:focus-visible { background:#2f81f7; color:#fff; outline:none; }
  .legend-chip.active { background:#1f6feb; color:#fff; border-color:#1b5dbf; }
  .util-button { background:#e6eef4; border:1px solid #c3d1dc; padding:4px 8px; font-size:0.65rem; border-radius:4px; cursor:pointer; color:#003553; }
  .util-button.active { background:#005494; color:#fff; border-color:#00416d; }
  .util-button:focus { outline:2px solid #91d2ff; }
  .rec-filter-group { display:flex; gap:4px; }
  .bool-mode-toggle { margin-left:auto; }
  /* Description clamp */
  .desc-cell { max-width:280px; position:relative; }
  .desc-text { display:-webkit-box; -webkit-line-clamp:3; -webkit-box-orient:vertical; overflow:hidden; text-overflow:ellipsis; white-space:normal; }
  .desc-cell.expanded .desc-text { -webkit-line-clamp:unset; max-height:none; }
  .desc-expand { position:absolute; bottom:2px; right:4px; background:#ffffffcc; border:1px solid #c3d1dc; font-size:0.55rem; padding:2px 4px; cursor:pointer; border-radius:3px; }
  .desc-expand:hover { background:#f1f5f8; }
  /* Assignment cell truncation */
  .assignment-cell { position:relative; }
  .assignment-text { display:-webkit-box; -webkit-line-clamp:3; -webkit-box-orient:vertical; overflow:hidden; text-overflow:ellipsis; white-space:normal; }
  .assignment-cell.expanded .assignment-text { -webkit-line-clamp:unset; max-height:none; }
  .assignment-expand { position:absolute; bottom:2px; right:4px; background:#ffffffcc; border:1px solid #c3d1dc; font-size:0.55rem; padding:2px 4px; cursor:pointer; border-radius:3px; }
  .assignment-expand:hover { background:#f1f5f8; }
  @media (max-width:1200px){ th,td{font-size:0.6em;} .search-box input{min-width:150px;} th.bool-col, td.bool-col { min-width:40px; } }
  @media (max-width:800px){ th,td{font-size:0.5em;} .view-toggle-group{display:flex;flex-wrap:wrap;} .search-box{margin-top:6px;} }
'@

  # Add visually hidden utility class and live region container for accessibility announcements
  # The live region enables non-visual users to receive feedback for actions (e.g., JSON copied, filter changes)
  $style = $style + '\n  .visually-hidden { position:absolute !important; width:1px !important; height:1px !important; padding:0 !important; margin:-1px !important; overflow:hidden !important; clip:rect(0 0 0 0) !important; white-space:nowrap !important; border:0 !important; }'
  $htmlDoc = "<!DOCTYPE html><html><head><meta charset='utf-8'><meta name='viewport' content='width=device-width, initial-scale=1'><style>$style</style><title>CA Export v$CAExportVersion - $TenantName</title></head><body><nav class='navbar-custom'><div class='nav-left'><div class='icon-server' aria-hidden='true'>🖥️</div><div class='brand'>CA Export $CAExportVersion</div><div class='view-toggle-group' role='tablist' aria-label='Report view' style='margin-left:20px;'>$recTabs</div><div class='search-box'><input id='policy-search' type='text' placeholder='Search policies...' aria-label='Search policies'/><span class='search-clear' id='policy-search-clear' title='Clear' role='button' aria-label='Clear search'>&times;</span></div><div class='status-filter'><select id='status-filter' aria-label='Filter by status'><option value='all'>All Policies</option><option value='enabled'>Enabled Only</option><option value='report'>Report-Only</option><option value='disabled'>Disabled Only</option></select></div></div><div class='nav-center'><strong>$TenantName</strong></div><div class='nav-right'><strong>$Date</strong></div></nav>$OmissionBannerHtml<div id='live-region' class='visually-hidden' aria-live='polite' aria-atomic='true'></div><button id='back-to-top' aria-label='Back to top' title='Back to top'>↑</button>"

  Write-Info 'Launching: Web Browser'
  # Ensure export path ends with separator
  if ($ExportLocation -and $ExportLocation[-1] -notin @('\', '/')) {
    $ExportLocation = $ExportLocation + [IO.Path]::DirectorySeparatorChar 
  }
  $Launch = Join-Path -Path $ExportLocation -ChildPath $FileName
  # Build the summary fragment first, then add the real summary panel (no placeholder/replacement needed)

  # ----- Summary Table -----
  $policyTotal = $CAPolicy.Count
  $enabledCount = ($CAPolicy | Where-Object { $_.State -eq 'enabled' }).Count
  $reportOnly = ($CAPolicy | Where-Object { $_.State -eq 'enabledForReportingButNotEnforced' }).Count
  $disabledCount = ($CAPolicy | Where-Object { $_.State -eq 'disabled' }).Count
  $withMfa = ($CAExport | Where-Object { $_.'Require MFA' -eq 'True' -or $_.'Authentication Strength MFA' }).Count
  $withStrength = ($CAExport | Where-Object { $null -ne $_.'Authentication Strength MFA' -and $_.'Authentication Strength MFA' -ne '' }).Count
  $withBlock = ($CAExport | Where-Object { $_.Block -eq 'True' }).Count
  $riskPolicies = ($CAPolicy | Where-Object { $_.Conditions.SignInRiskLevels -or $_.Conditions.UserRiskLevels }).Count
  $devicePolicies = ($CAPolicy | Where-Object { $_.Conditions.Devices.IncludeDevices -or $_.Conditions.Devices.ExcludeDevices -or $_.GrantControls.BuiltInControls -contains 'compliantDevice' -or $_.GrantControls.BuiltInControls -contains 'domainJoinedDevice' }).Count
  $termsPolicies = ($CAPolicy | Where-Object { $_.GrantControls.TermsOfUse }).Count
  $phishResistant = ($CAExport | Where-Object { $_.'Authentication Strength MFA' -match 'Phishing-resistant' }).Count
  $avgModifiedDays = 0
  $now = Get-Date
  $dateVals = @()
  foreach ( $p in $CAExport ) {
    if ( $p.DateModified ) {
      $d = [datetime] $p.DateModified
      $dateVals += ( ( $now - $d ).TotalDays )
    }
  }
  if ( $dateVals.Count -gt 0 ) {
    $avgModifiedDays = [math]::Round( ( $dateVals | Measure-Object -Average | Select-Object -ExpandProperty Average ), 1 )
  }

  $summaryTable = @()
  $summaryTable += '<div class="summary-section summary-wrapper" style="margin:20px 25px;">'
  $summaryTable += '<h2 style="margin:10px 6px;font-size:1rem;color:#003553;letter-spacing:.5px;">Policy Summary</h2>'
  $summaryTable += '<table class="summary-table">'
  $summaryTable += '<thead><tr><th>Metric</th><th>Value</th><th>Percent</th></tr></thead><tbody>'
  function Add-SummaryRow {
    param(
      [string]$Name,
      $Value,
      [double]$TotalRef
    )
    $pct = ''
    $numericVal = 0.0
    $canParse = $false
    if ($null -ne $Value) {
      if (
        $Value -is [int] -or $Value -is [long] -or $Value -is [double] -or $Value -is [decimal]) {
        $numericVal = [double]$Value; $canParse = $true
      }
      elseif ($Value -isnot [string]) {
        try {
          $numericVal = [double]$Value; $canParse = $true 
        }
        catch {
          Write-Verbose 'Add-SummaryRow: numeric conversion failed.' 
        }
      }
      elseif ([double]::TryParse([string]$Value, [ref]$numericVal)) {
        $canParse = $true
      }
    }
    if ($canParse -and $TotalRef -gt 0) {
      try {
        $pct = ('{0:P1}' -f ($numericVal / $TotalRef)) 
      }
      catch {
        $pct = '' 
      }
    }
    $encName = [System.Web.HttpUtility]::HtmlEncode($Name)
    $displayValue = if ($Value -is [string]) {
      [System.Web.HttpUtility]::HtmlEncode($Value) 
    }
    else {
      $Value 
    }
    return "<tr><td>$encName</td><td>$displayValue</td><td>$pct</td></tr>"
  }
  # Non-numeric row (percent intentionally blank)
  $summaryTable += (Add-SummaryRow -Name 'Total Policies' -Value $policyTotal -TotalRef '')
  $summaryTable += (Add-SummaryRow -Name 'Enabled' -Value $enabledCount -TotalRef $policyTotal)
  $summaryTable += (Add-SummaryRow -Name 'Report-only' -Value $reportOnly -TotalRef $policyTotal)
  $summaryTable += (Add-SummaryRow -Name 'Disabled' -Value $disabledCount -TotalRef $policyTotal)
  $summaryTable += (Add-SummaryRow -Name 'Policies Requiring MFA (any)' -Value $withMfa -TotalRef $policyTotal)
  $summaryTable += (Add-SummaryRow -Name 'Policies With Auth Strength' -Value $withStrength -TotalRef $policyTotal)
  $summaryTable += (Add-SummaryRow -Name 'Phishing-resistant Strength' -Value $phishResistant -TotalRef $policyTotal)
  $summaryTable += (Add-SummaryRow -Name 'Policies With Block Control' -Value $withBlock -TotalRef $policyTotal)
  $summaryTable += (Add-SummaryRow -Name 'Risk-Based Policies' -Value $riskPolicies -TotalRef $policyTotal)
  $summaryTable += (Add-SummaryRow -Name 'Device Condition / Control Policies' -Value $devicePolicies -TotalRef $policyTotal)
  $summaryTable += (Add-SummaryRow -Name 'Terms of Use Policies' -Value $termsPolicies -TotalRef $policyTotal)
  # Admin-focused metrics using raw policy objects with helper functions
  if ($RawPolicyIndex.Count -gt 0) {
    $rawPolicies = $RawPolicyIndex.Values
    $adminTargets = 0
    $adminMfa = 0
    $adminPhish = 0
    foreach ($rp in $rawPolicies) {
      if (Test-PolicyTargetsAdminRoles -Policy $rp) {
        $adminTargets++
        if (Test-PolicyRequiresMfaForAdmins -Policy $rp) { $adminMfa++ }
        if (Test-PolicyRequiresPhishResistantMfaForAdmins -Policy $rp) { $adminPhish++ }
      }
    }
    $summaryTable += (Add-SummaryRow -Name 'Policies Targeting Admin Roles' -Value $adminTargets -TotalRef $policyTotal)
    $summaryTable += (Add-SummaryRow -Name 'Admin Policies Requiring MFA' -Value $adminMfa -TotalRef $adminTargets)
    $summaryTable += (Add-SummaryRow -Name 'Admin Policies Phish-Resistant' -Value $adminPhish -TotalRef $adminTargets)
  }
  $summaryTable += ("<tr><td>Average Age Since Modified (days)</td><td>$avgModifiedDays</td><td></td></tr>")
  $summaryTable += '</tbody></table>'
  $summaryTable += '</div>'
  # Prepare summary HTML fragment and assign stable id for later detection
  $SummaryHtmlFragment = ($summaryTable -join "`n") -replace '<div class="summary-section"', '<div class="summary-section" id="summary-root"'
  # Add populated summary panel, then open policies panel
  $HtmlParts += "<div class='policy-summary' id='panel-summary' role='tabpanel' aria-labelledby='tab-summary' aria-hidden='true' style='display:none;'>$SummaryHtmlFragment</div>"
  $HtmlParts += "<div class='policy-export' id='panel-policies' role='tabpanel' aria-labelledby='tab-policies' aria-hidden='false'>"
  # Legend / utilities bar
  $legendHtml = @'
  <div class='legend-bar' id='legend-bar'>
    <span class='legend-title'>Legend:</span>
    <span class='legend-item'><span class='legend-swatch'>—</span> Empty / None</span>
    <span class='legend-item'><span class='legend-swatch enabled'>EN</span> Enabled</span>
    <span class='legend-item'><span class='legend-swatch report'>RP</span> Report-only</span>
    <span class='legend-item'><span class='legend-swatch disabled'>DIS</span> Disabled</span>
    <span class='legend-item'><span class='legend-swatch pass'>✔</span> Rec Pass</span>
    <span class='legend-item'><span class='legend-swatch fail'>⚠</span> Rec Attention</span>
    <span class='legend-item'><span class='legend-swatch'>✔</span> Boolean True</span>
    <span class='legend-item'><span class='legend-swatch'>—</span> Boolean False</span>
    <span class='legend-divider' aria-hidden='true'></span>
  <button type='button' id='bool-mode-toggle' class='util-button bool-mode-toggle' data-mode='icon' aria-pressed='false' title='Toggle boolean display mode'>Boolean: Icons</button>
  </div>
'@
  $HtmlParts += $legendHtml
  # Define columns (reuse CSV default ordering for consistency, but can trim for HTML readability)
  $htmlColumns = @(
    'Name', 'Modified', 'Created',
    'Included Users', 'Excluded Users',
    'Included Applications', 'Excluded Applications', 'User Actions', 'Auth Context', 'Included Locations', 'Excluded Locations', 'User Risk', 'SignIn Risk', 'Platforms Included', 'Platforms Excluded', 'Included Devices', 'Excluded Devices', 'Client Apps', 'Device Filters', 'Authentication Flows', 'Block', 'Grant Operator', 'Require MFA', 'Authentication Strength MFA',
    'Compliant Device', 'Domain Joined Device', 'Compliant Application', 'Approved Application', 'Password Change', 'Terms Of Use', 'Custom Controls',
    'Application Enforced Restrictions', 'Cloud App Security', 'Sign In Frequency', 'Persistent Browser', 'Continuous Access Evaluation', 'Resilient Defaults', 'Secure Sign In Session', 'RawJson'
  )
  # Build table rows per policy
  $table = @()
  $table += '<table class="sticky-header">'
  # Specify boolean columns for icon rendering
  $boolColumns = @('Block', 'Require MFA', 'Compliant Device', 'Domain Joined Device', 'Compliant Application', 'Approved Application', 'Password Change', 'Application Enforced Restrictions', 'Cloud App Security', 'Resilient Defaults')
  # Build multi-row header: first row (spanning categories) + second row (individual columns)
  $groupDefinitions = @(
    @{ Name = 'General'; Class = 'group-general;'; Columns = @('Name', 'Modified', 'Created') },
    @{ Name = 'Assignments - Users'; Class = 'group-assign-users;'; Columns = @('Included Users', 'Excluded Users') },
    @{ Name = 'Assignments - Target Resources'; Class = 'group-assign-target'; Columns = @('Included Applications', 'Excluded Applications', 'User Actions', 'Auth Context') },
    @{ Name = 'Assignments - Network'; Class = 'group-network'; Columns = @('Included Locations', 'Excluded Locations') },
    @{ Name = 'Assignments - Conditions'; Class = 'group-conditions'; Columns = @('User Risk', 'SignIn Risk', 'Platforms Included', 'Platforms Excluded', 'Included Devices', 'Excluded Devices', 'Client Apps', 'Device Filters', 'Authentication Flows') },
    @{ Name = 'Access controls - Grant'; Class = 'group-grant'; Columns = @('Block', 'Grant Operator', 'Require MFA', 'Authentication Strength MFA', 'Compliant Device', 'Domain Joined Device', 'Compliant Application', 'Approved Application', 'Password Change', 'Terms Of Use', 'Custom Controls') },
    @{ Name = 'Access controls - Session'; Class = 'group-session'; Columns = @('Application Enforced Restrictions', 'Cloud App Security', 'Sign In Frequency', 'Persistent Browser', 'Continuous Access Evaluation', 'Resilient Defaults', 'Secure Sign In Session') },
    @{ Name = 'Output'; Class = 'group-output'; Columns = @('RawJson') }
  )
  # Sanity: ensure ordering in htmlColumns matches group concatenation; if mismatch, groups may mis-align
  # Compose group header row
  $groupRow = '<tr class="group-header">' + (
    $groupDefinitions | ForEach-Object {
      $totalCols = $_.Columns.Count
      # Count how many ID columns are in this group (these will be hidden)
      $hiddenCount = ($_.Columns | Where-Object { $idColumns -contains $_ }).Count
      $visibleColspan = $totalCols - $hiddenCount
      $gName = [System.Web.HttpUtility]::HtmlEncode($_.Name)
      $gClass = $_.Class
      "<th class='group-span $gClass' colspan='$visibleColspan' data-total='$totalCols' data-visible='$visibleColspan'>$gName</th>"
    }
  ) -join '' + '</tr>'
  # Map column -> group class
  $columnGroupClassMap = @{}
  foreach ($gd in $groupDefinitions) { foreach ($col in $gd.Columns) { $columnGroupClassMap[$col] = $gd.Class } }
  # Compose second (field) header row with group color classes
  $fieldHeaderRow = '<tr class="column-header-row">' + ($htmlColumns | ForEach-Object {
      $gCls = $columnGroupClassMap[$_]
      if ($_ -eq 'Name') {
        "<th id='th-name' class='sticky-name name-col $gCls'>Name</th>" 
      }
      elseif ($boolColumns -contains $_) {
        "<th class='bool-col $gCls'>$_</th>" 
      }
      else {
        "<th class='$gCls'>$_</th>" 
      }
    }) -join '' + '</tr>'
  $header = '<thead>' + $groupRow + $fieldHeaderRow + '</thead>'
  $table += $header
  $table += '<tbody>'
  foreach ($p in $CAExport) {
    $rowTds = @()
    $colIndex = 0
    foreach ($col in $htmlColumns) {
      $colIndex++
      $raw = $null
      # Use direct property access instead of Match() which uses wildcards
      if ($p.PSObject.Properties.Name -contains $col) {
        $raw = $p.$col 
      }
      
      if ($col -eq 'RawJson') {
        # Lazy JSON: store minimal placeholder with data attributes; JSON filled client-side on first expansion
        $hasOrig = ($RawPolicyIndex -and $p.PolicyId -and $RawPolicyIndex.ContainsKey($p.PolicyId))
        $origFlag = if ($hasOrig) { 'true' } else { 'false' }
        $rowTds += "<td class='$idHiddenClass'><details class='raw-json lazy-json' data-policyid='$($p.PolicyId)' data-has-orig='$origFlag'><summary>View</summary><div class='json-toggle-bar'><button type='button' class='json-btn active' data-mode='mut' aria-pressed='true'>Mutated</button><button type='button' class='json-btn' data-mode='orig' aria-pressed='false'$(if($hasOrig){''}else{' disabled title=\"No original snapshot\"'})>Original</button><button type='button' class='json-copy util-button' data-copy='mut' title='Copy displayed JSON' aria-label='Copy JSON'>Copy</button></div><pre class='json-block' data-mode='mut' data-loaded='false'></pre><pre class='json-block' data-mode='orig' style='display:none;' data-loaded='false'></pre></details></td>"
      }
      elseif ($col -eq 'Name' -and $p.PolicyId) {
        $plink = "$LinkURL$($p.PolicyId)"
        $rawName = [string]$raw
        $safe = [System.Web.HttpUtility]::HtmlEncode($rawName)
        # Get status for label
        $statusVal = [string]$p.Status
        # Generate status label
        $statusLabelClass = ''
        $statusLabelText = ''
        switch -Regex ($statusVal) {
          '^enabled$' {
            $statusLabelClass = 'status-label status-enabled'
            $statusLabelText = 'Enabled'
            break
          }
          'enabledForReportingButNotEnforced' {
            $statusLabelClass = 'status-label status-report'
            $statusLabelText = 'Report'
            break
          }
          'report only' {
            $statusLabelClass = 'status-label status-report'
            $statusLabelText = 'Report'
            break
          }
          '^disabled$' {
            $statusLabelClass = 'status-label status-disabled'
            $statusLabelText = 'Disabled'
            break
          }
          default {
            $statusLabelClass = 'status-label'
            $statusLabelText = $statusVal
          }
        }
        $statusLabel = "<span class='$statusLabelClass' title='$statusVal'>$statusLabelText</span>"
        
        # Add recommended name pill if available
        $recNamePill = ''
        if ($p.'Recommended Name' -and $p.'Recommended Name' -ne $p.Name) {
          $recNameSafe = [System.Web.HttpUtility]::HtmlEncode([string]$p.'Recommended Name')
          $recNamePill = "<span class='pill-recommended' title='Recommended Name'>$recNameSafe</span>"
        }
        
        $rowTds += "<td class='sticky-name name-col'>$statusLabel<div class='name-content'><span class='policy-name'>$safe</span><!--$recNamePill--><a class='policy-link' href='$plink' target='_blank' aria-label='Open policy in new tab' title='Open policy in new tab'><svg class='icon-ext' viewBox='0 0 24 24' focusable='false' aria-hidden='true'><path d='M14 3h7v7m0-7L10 14' stroke-linecap='round' stroke-linejoin='round'/><path d='M21 21H3V3h7' stroke-linecap='round' stroke-linejoin='round'/></svg><span class='visually-hidden'></span></a></div></td>"
      }
      elseif ($col -eq 'Description') {
        if ($null -eq $raw -or [string]::IsNullOrWhiteSpace([string]$raw)) {
          $rowTds += "<td class='desc-cell'><span class='placeholder' aria-label='None'>—</span></td>"
        }
        else {
          $safe = [System.Web.HttpUtility]::HtmlEncode([string]$raw)
          $descId = 'desc-' + [guid]::NewGuid().ToString('N')
          $rowTds += "<td class='desc-cell'><div class='desc-text' id='$descId'>$safe</div><button type='button' class='desc-expand' aria-label='Expand description' aria-expanded='false' aria-controls='$descId'>More</button></td>"
        }
      }
      elseif ($col -eq 'Block') {
        if ($raw -eq 'True' -or $raw -eq $true) {
          $rowTds += "<td class='$idHiddenClass'><span class='pill-block' aria-label='Block'>BLOCK</span></td>" 
        }
        else {
          $rowTds += "<td class='$idHiddenClass'><span class='pill-grant' aria-label='Grant'>ALLOW</span></td>" 
        }
      }
      elseif ($boolColumns -contains $col) {
        if ($raw -eq 'True' -or $raw -eq $true) {
          $rowTds += "<td class='bool-col$idHiddenClass'><span class='bool-yes' aria-label='True'>✔</span></td>" 
        }
        else {
          $rowTds += "<td class='bool-col$idHiddenClass'><span class='bool-no' aria-label='False'>—</span></td>" 
        }
      }
      else {
        if ($null -eq $raw -or [string]::IsNullOrWhiteSpace([string]$raw)) {
          $rowTds += "<td class='$idHiddenClass'><span class='placeholder' aria-label='None'>—</span></td>" 
        }
        else {
          $rawString = [string]$raw
          $safe = [System.Web.HttpUtility]::HtmlEncode($rawString)
          
          # Check if this is an assignment column that should be truncatable
          $truncatableColumns = @('Included Users', 'Excluded Users', 'Included Applications', 'Excluded Applications', 'Included Locations', 'Excluded Locations')
          if ($truncatableColumns -contains $col -and $rawString.Length -gt 100) {
            $assignId = 'assign-' + [guid]::NewGuid().ToString('N')
            $rowTds += "<td class='assignment-cell $idHiddenClass'><div class='assignment-text' id='$assignId'>$safe</div><button type='button' class='assignment-expand' aria-label='Expand' aria-expanded='false' aria-controls='$assignId'>More</button></td>"
          }
          else {
            $rowTds += "<td class='$idHiddenClass'>$safe</td>"
          }
        }
      }
    }
    $table += '<tr>' + ($rowTds -join '') + '</tr>'
  }
  $table += '</tbody></table>'
  # (Summary panel already populated above)
  # Append policies table (inside policies panel) exactly once, then close policies panel
  $HtmlParts += ($table -join "`n")
  $HtmlParts += '</div>'  # close policy-export / panel-policies
  # Append recommendations panel as a sibling (was previously nested causing it to be hidden and table omitted)
  $HtmlParts += $SecurityCheck

  # Embed lightweight policy index for lazy JSON (excluding heavy raw JSON fields)
  $policyDataJson = ($CAExport | Select-Object PolicyId, Name, Status, 'Grant Controls', 'Included Users', 'Excluded Users', 'Applications', 'Locations', 'Platforms', 'Client Apps', 'State' | ConvertTo-Json -Depth 4 -Compress)
  $rawIndexJson = if ($RawPolicyIndex.Count -gt 0) { ($CAPolicy | ConvertTo-Json -Depth 6 -Compress) } else { '[]' }
  $fallbackToggle = @"
<script>
window.CAExportData = $policyDataJson;
// Build RawPolicyDataIndex keyed by id for O(1) lookup during lazy load
window.RawPolicyDataIndex = (function(){
  try { var arr = $rawIndexJson; var idx = {}; if(Array.isArray(arr)){ for(var i=0;i<arr.length;i++){ var it=arr[i]; if(it && it.id){ idx[it.id]=it; } } } return idx; } catch(e){ return {}; }
})();
// Enhanced vanilla JS interactions (3-tab aware, accessible)
document.addEventListener('DOMContentLoaded', function(){
  var tabs = Array.prototype.slice.call(document.querySelectorAll('.view-toggle-group [role=tab]'));
  var hasRecommendations = !!document.getElementById('tab-recommendations');
  var panels = {
    'tab-summary': document.getElementById('panel-summary'),
    'tab-policies': document.getElementById('panel-policies')
  };
  if(hasRecommendations){ panels['tab-recommendations'] = document.getElementById('panel-recommendations'); }
  // Explicit panel element references for later logic (table lookup, etc.)
  var policies = panels['tab-policies'];
  var summary  = panels['tab-summary'];
  var recs     = hasRecommendations ? panels['tab-recommendations'] : null;
  var searchInput = document.getElementById('policy-search');
  var searchClear = document.getElementById('policy-search-clear');
  var backToTop = document.getElementById('back-to-top');
  var recToggle = document.getElementById('toggle-recs');
  var boolModeToggle = document.getElementById('bool-mode-toggle');

  function activate(tabId){ // Persist and switch between main report panels (Summary / Policy Details / Recommendations)
    if(!panels[tabId]){ if(window.console) console.warn('Activate called for unknown tabId', tabId); }
    tabs.forEach(function(t){
      var on = (t.id === tabId);
      t.classList.toggle('active', on);
      t.setAttribute('aria-selected', on ? 'true' : 'false');
      t.tabIndex = on ? 0 : -1;
      var panel = panels[t.id];
      if(panel){
        panel.style.display = on ? 'block' : 'none';
        panel.setAttribute('aria-hidden', on ? 'false':'true');
      }
    });
    if(window.console){ console.log('Tab activated:', tabId, 'summary visible?', summary && summary.style.display, 'policies visible?', policies && policies.style.display); }
    if(searchInput){
      var isPolicies = (tabId === 'tab-policies');
      searchInput.disabled = !isPolicies;
      searchInput.title = isPolicies ? 'Search policies' : 'Search only available in Policy Details view';
    }
  var hash = '#policies';
  if(hasRecommendations && tabId === 'tab-recommendations') hash = '#recommendations';
    if(tabId === 'tab-summary') hash = '#summary';
  try { history.replaceState(null, '', hash); } catch(e){}
  try { localStorage.setItem('caexport.activeTab', tabId); } catch(e){}
    window.scrollTo({top:0});
  }

  // Click handling
  tabs.forEach(function(t){ t.addEventListener('click', function(){ activate(t.id); }); });

  // Keyboard navigation (Left/Right/Home/End)
  function tabKeyHandler(e){
    var key = e.key;
    var currentIndex = tabs.indexOf(document.activeElement);
    if(currentIndex === -1) return;
    if(key === 'ArrowRight'){ e.preventDefault(); var next = (currentIndex+1) % tabs.length; tabs[next].focus(); activate(tabs[next].id); }
    else if(key === 'ArrowLeft'){ e.preventDefault(); var prev = (currentIndex-1+tabs.length) % tabs.length; tabs[prev].focus(); activate(tabs[prev].id); }
    else if(key === 'Home'){ e.preventDefault(); tabs[0].focus(); activate(tabs[0].id); }
    else if(key === 'End'){ e.preventDefault(); tabs[tabs.length-1].focus(); activate(tabs[tabs.length-1].id); }
  }
  tabs.forEach(function(t){ t.addEventListener('keydown', tabKeyHandler); });

  // Hash routing
  var storedTab = null; try { storedTab = localStorage.getItem('caexport.activeTab'); } catch(e){}
  if(window.location.hash){
    var h = window.location.hash.toLowerCase();
    if(hasRecommendations && h.includes('recommend')) activate('tab-recommendations');
    else if(h.includes('summary')) activate('tab-summary');
    else activate('tab-policies');
  } else if(storedTab && panels[storedTab]) { activate(storedTab); } else { activate('tab-policies'); }

  // Table interactions (guarded)
  var table = policies ? policies.querySelector('table') : null;
  if(table){
    table.addEventListener('click', function(e){
      // Lazy JSON population on first expand
      var summary = e.target.closest('summary');
      if(summary){
        var details = summary.parentElement;
        if(details && details.classList.contains('lazy-json')){
          var opened = details.hasAttribute('open');
          // Delay execution until the element is actually opened (after default toggle)
          setTimeout(function(){
            if(details.hasAttribute('open')){
              var mutPre = details.querySelector("pre.json-block[data-mode='mut']");
              if(mutPre && mutPre.getAttribute('data-loaded')==='false'){
                var pid = details.getAttribute('data-policyid');
                try {
                  var policyRow = window.CAExportData ? window.CAExportData.find(function(x){ return x.PolicyId===pid; }) : null;
                  if(policyRow){
                    var shallow = Object.assign({}, policyRow);
                    delete shallow.RawJson;
                    mutPre.textContent = JSON.stringify(shallow, null, 2);
                    mutPre.setAttribute('data-loaded','true');
                  }
                  var hasOrig = details.getAttribute('data-has-orig')==='true';
                  if(hasOrig){
                    var origPre = details.querySelector("pre.json-block[data-mode='orig']");
                    if(origPre && origPre.getAttribute('data-loaded')==='false' && window.RawPolicyDataIndex){
                      var origObj = window.RawPolicyDataIndex[pid];
                      if(origObj){ origPre.textContent = JSON.stringify(origObj, null, 2); origPre.setAttribute('data-loaded','true'); }
                    }
                  }
                } catch(ex){ console.warn('Lazy JSON load failed', ex); }
              }
            }
          }, 0);
        }
      }
      // JSON toggle buttons
      var jsonBtn = e.target.closest('button.json-btn');
      if(jsonBtn){
        var td = jsonBtn.closest('td');
        var btns = td.querySelectorAll('button.json-btn');
        for(var i=0;i<btns.length;i++){ btns[i].classList.remove('active'); btns[i].setAttribute('aria-pressed','false'); }
        jsonBtn.classList.add('active'); jsonBtn.setAttribute('aria-pressed','true');
        var mode = jsonBtn.getAttribute('data-mode');
        var blocks = td.querySelectorAll('pre.json-block');
        for(var b=0;b<blocks.length;b++){ blocks[b].style.display = (blocks[b].getAttribute('data-mode') === mode) ? 'block' : 'none'; }
        // Update copy button data attribute to reflect current mode
        var copyBtn = td.querySelector('button.json-copy');
        if(copyBtn){ copyBtn.setAttribute('data-copy', mode); }
        return; // don't fall through to row select logic for the toggle click
      }
      // JSON copy button
      var jsonCopy = e.target.closest('button.json-copy');
      if(jsonCopy){
        var td = jsonCopy.closest('td');
        var mode = jsonCopy.getAttribute('data-copy') || 'mut';
        var block = td.querySelector("pre.json-block[data-mode='"+mode+"']");
        if(block){
          var text = block.textContent || '';
          navigator.clipboard.writeText(text).then(function(){
            var lr = document.getElementById('live-region');
            if(lr){ lr.textContent = 'JSON copied to clipboard ('+mode+').'; }
            jsonCopy.textContent = 'Copied';
            setTimeout(function(){ jsonCopy.textContent='Copy'; }, 1500);
          }).catch(function(){
            var lr = document.getElementById('live-region');
            if(lr){ lr.textContent = 'Copy failed'; }
          });
        }
        return;
      }
      var tr = e.target.closest('tr');
      if(tr && tr.parentElement.tagName !== 'THEAD'){
        tr.classList.toggle('selected');
      }
    });

  }

  // Status filter functionality
  var statusFilter = document.getElementById('status-filter');
  if(statusFilter && table){
    function applyStatusFilter(){
      var filterValue = statusFilter.value;
      var rows = table.querySelectorAll('tbody tr');
      var visibleCount = 0;
      rows.forEach(function(row){
        var nameCell = row.querySelector('td.name-col');
        if(!nameCell){ return; }
        var statusLabel = nameCell.querySelector('.status-label');
        if(!statusLabel){
          row.style.display = '';
          visibleCount++;
          return;
        }
        var shouldShow = false;
        if(filterValue === 'all'){
          shouldShow = true;
        } else if(filterValue === 'enabled' && statusLabel.classList.contains('status-enabled')){
          shouldShow = true;
        } else if(filterValue === 'report' && statusLabel.classList.contains('status-report')){
          shouldShow = true;
        } else if(filterValue === 'disabled' && statusLabel.classList.contains('status-disabled')){
          shouldShow = true;
        }
        row.style.display = shouldShow ? '' : 'none';
        if(shouldShow){ visibleCount++; }
      });
      var liveRegion = document.getElementById('live-region');
      if(liveRegion){
        var filterText = filterValue === 'all' ? 'all policies' : filterValue + ' policies';
        liveRegion.textContent = 'Showing ' + visibleCount + ' ' + filterText;
      }
    }
    statusFilter.addEventListener('change', applyStatusFilter);
    // Restore filter state
    try {
      var storedFilter = localStorage.getItem('caexport.statusFilter');
      if(storedFilter){ statusFilter.value = storedFilter; applyStatusFilter(); }
    } catch(e){}
    // Save filter state on change
    statusFilter.addEventListener('change', function(){
      try { localStorage.setItem('caexport.statusFilter', statusFilter.value); } catch(e){}
    });
  }

  // Recommendation expand/collapse (simple show/hide of recommendation cards)
  if(recToggle){
    recToggle.addEventListener('click', function(){
      var recPanel = document.getElementById('panel-recommendations');
      if(!recPanel) return;
      var items = recPanel.querySelectorAll('.recommendation');
      var collapsing = recToggle.getAttribute('data-collapsed') !== 'true';
      items.forEach(function(el){ el.style.display = collapsing ? 'none' : ''; });
      recToggle.setAttribute('data-collapsed', collapsing ? 'true' : 'false');
      recToggle.textContent = collapsing ? 'Show Recs' : 'Hide Recs';
      recToggle.classList.toggle('active', !collapsing);
    });
  }

  // Recommendation pass/fail filter buttons injected dynamically (with persistence)
  // Stored key: caexport.recFilter -> 'all' | 'fail' | 'pass'
  (function(){
    var recPanel = document.getElementById('panel-recommendations');
    if(!recPanel) return;
    if(!recPanel.querySelector('.rec-filter-group')){
      var bar = document.createElement('div');
      bar.className='rec-filter-group';
      bar.setAttribute('role','group');
      bar.style.margin='0 0 10px 0';
      var buttons=[
        {id:'rec-filter-all', label:'All', filter:'all'},
        {id:'rec-filter-fail', label:'Attention', filter:'fail'},
        {id:'rec-filter-pass', label:'Pass', filter:'pass'}
      ];
      var storedRecFilter = null; try { storedRecFilter = localStorage.getItem('caexport.recFilter'); } catch(e){}
      buttons.forEach(function(cfg,idx){
        var b=document.createElement('button');
        var makeActive = storedRecFilter ? (storedRecFilter===cfg.filter) : (idx===0);
        b.type='button'; b.textContent=cfg.label; b.className='util-button'+(makeActive?' active':'');
        b.dataset.filter=cfg.filter; b.id=cfg.id; b.addEventListener('click', function(){
          var allBtns=bar.querySelectorAll('button'); allBtns.forEach(function(bb){bb.classList.remove('active');});
          b.classList.add('active');
          var cards = recPanel.querySelectorAll('details.recommendation-card');
          cards.forEach(function(card){
            if(cfg.filter==='all'){ card.style.display=''; }
            else {
              var isPass = card.classList.contains('pass');
              var show = (cfg.filter==='pass' && isPass) || (cfg.filter==='fail' && !isPass);
              card.style.display = show ? '' : 'none';
            }
          });
          try { localStorage.setItem('caexport.recFilter', cfg.filter); } catch(e){}
        });
        bar.appendChild(b);
      });
      if(storedRecFilter && storedRecFilter!=='all'){
        var trigger = bar.querySelector('button[data-filter="'+storedRecFilter+'"]');
        if(trigger){ trigger.click(); }
      }
      recPanel.insertBefore(bar, recPanel.firstChild);
    }
  })();

  // Boolean icon/text toggle with persistence
  // Stored key: caexport.boolMode -> 'icon' (default) or 'text'
  if(boolModeToggle && table){
    var storedBoolMode = null; try { storedBoolMode = localStorage.getItem('caexport.boolMode'); } catch(e){}
    function applyBoolMode(mode){
      var toIcons = (mode === 'icon');
      boolModeToggle.setAttribute('data-mode', mode);
      boolModeToggle.textContent = 'Boolean: ' + (toIcons ? 'Icons' : 'Text');
      var boolCols = ['Block','Require MFA','CompliantDevice','DomainJoinedDevice','CompliantApplication','ApprovedApplication','PasswordChange','ApplicationEnforcedRestrictions','CloudAppSecurity','ResilientDefaults'];
      var headerCells = table.querySelectorAll('thead th');
      var indexes=[]; headerCells.forEach(function(h,i){ if(boolCols.indexOf(h.textContent.trim())>-1){ indexes.push(i); }});
      var rows = table.querySelectorAll('tbody tr');
      rows.forEach(function(r){ indexes.forEach(function(idx){ var cell=r.children[idx]; if(!cell)return; var span=cell.querySelector('span.bool-yes, span.bool-no'); if(!span)return; span.textContent = span.classList.contains('bool-yes') ? (toIcons?'✔':'True') : (toIcons?'—':'False'); }); });
    }
    applyBoolMode(storedBoolMode && ['icon','text'].indexOf(storedBoolMode)>-1 ? storedBoolMode : 'icon');
    boolModeToggle.addEventListener('click', function(){
      var current = boolModeToggle.getAttribute('data-mode');
      var next = current === 'icon' ? 'text' : 'icon';
      applyBoolMode(next);
      try { localStorage.setItem('caexport.boolMode', next); } catch(e){}
    });
  }

  // Description expand/collapse (aria-expanded + aria-controls for accessibility)
  if(table){
    table.addEventListener('click', function(e){
      var btn = e.target.closest('.desc-expand');
      if(!btn) return;
      var cell = btn.closest('.desc-cell');
      var expanded = cell.classList.toggle('expanded');
      btn.textContent = expanded ? 'Less' : 'More';
  btn.setAttribute('aria-label', expanded ? 'Collapse description' : 'Expand description');
  btn.setAttribute('aria-expanded', expanded ? 'true' : 'false');
    });
  }

  // Assignment cell expand/collapse
  if(table){
    table.addEventListener('click', function(e){
      var btn = e.target.closest('.assignment-expand');
      if(!btn) return;
      var cell = btn.closest('.assignment-cell');
      var expanded = cell.classList.toggle('expanded');
      btn.textContent = expanded ? 'Less' : 'More';
      btn.setAttribute('aria-label', expanded ? 'Collapse' : 'Expand');
      btn.setAttribute('aria-expanded', expanded ? 'true' : 'false');
    });
  }

  // Policy search filter
  function filterPolicies(){
    if(!table) return; // nothing to do
    var q = (searchInput.value || '').toLowerCase();
    if(searchInput.parentElement){
      if(q.length>0){ searchInput.parentElement.classList.add('has-value'); }
      else { searchInput.parentElement.classList.remove('has-value'); }
    }
    var rows = table.querySelectorAll('tbody tr');
    for(var i=0;i<rows.length;i++){
      var row = rows[i];
      var show = true;
      if(show && q.length>0){
        var cells = row.children;
        var match = false;
        for(var c=0;c<cells.length;c++){
          if(cells[c].textContent && cells[c].textContent.toLowerCase().indexOf(q) > -1){ match = true; break; }
        }
        show = match;
      }
      row.style.display = show ? '' : 'none';
    }
  }
  if(searchInput){
    searchInput.addEventListener('input', filterPolicies);
    searchInput.addEventListener('keydown', function(e){ if(e.key==='Escape'){ searchInput.value=''; filterPolicies(); searchInput.blur(); } });
  }
  if(searchClear){
    searchClear.addEventListener('click', function(){ searchInput.value=''; filterPolicies(); searchInput.focus(); });
  }

  // Back to top visibility
  function updateBackToTop(){
    if(!backToTop) return;
    if(window.scrollY > 300){ backToTop.style.display='block'; } else { backToTop.style.display='none'; }
  }
  window.addEventListener('scroll', updateBackToTop);
  if(backToTop){
    backToTop.addEventListener('click', function(){ window.scrollTo({top:0, behavior:'smooth'}); });
  }
  updateBackToTop();

  // Ensure search state matches initial view (handled in activate but safeguard)
  if(searchInput && searchInput.disabled && window.location.hash.toLowerCase().indexOf('policies')>-1){ searchInput.disabled=false; }
});
</script>
"@
  $fullPage = $htmlDoc + $HtmlParts + $fallbackToggle + '</body></html>'
  $fullPage | Out-File $Launch -Encoding UTF8
  if (-not $NoBrowser) {
    if ($IsWindows -or $PSVersionTable.PSVersion.Major -le 5) {
      Start-Process $Launch
    }
    elseif ($IsMacOS) {
      & open $Launch
    }
    elseif ($IsLinux) {
      & xdg-open $Launch
    }
    else {
      Write-Info "HTML report saved to: $Launch (auto-open not supported on this platform)"
    }
  }
}
if ($JsonExport) {
  Write-Info 'Saving to File: JSON (enriched policies) & RAW JSON'
  $LaunchJson = Join-Path -Path $ExportLocation -ChildPath $JsonFileName
  try {
    $CAExport | ConvertTo-Json -Depth 12 | Out-File $LaunchJson
    Write-Info "JSON (enriched) saved: $LaunchJson"
  }
  catch {
    Write-Warn "Failed to save enriched JSON: $_" 
  }
  if ($RawPolicyObjects) {
    $rawFile = Join-Path -Path $ExportLocation -ChildPath "${baseName}_raw.json"
    try {
      $RawPolicyObjects | ConvertTo-Json -Depth 12 | Out-File $rawFile; Write-Info "Raw JSON saved: $rawFile" 
    }
    catch {
      Write-Warn "Failed to save raw JSON: $_" 
    }
  }
}
if ($CsvExport) {
  Write-Info 'Saving to File: CSV (one row per policy)'
  $LaunchCsv = Join-Path -Path $ExportLocation -ChildPath $CsvFileName
  $defaultColumns = @(
    'Name', 'PolicyId', 'Status', 'Modified', 'Created',
    'Included Users', 'Excluded Users',
    'Included Applications', 'Excluded Applications',
    'User Actions', 'Auth Context', 'Included Locations', 'Excluded Locations', 'User Risk', 'SignIn Risk', 'Platforms Included', 'Platforms Excluded', 'Included Devices', 'Excluded Devices', 'Client Apps', 'Device Filters', 'Authentication Flows', 'Block', 'Grant Operator', 'Require MFA', 'Authentication Strength MFA',
    'Compliant Device', 'Domain Joined Device', 'Compliant Application', 'Approved Application', 'Password Change', 'Terms Of Use', 'Custom Controls', 'Application Enforced Restrictions', 'Cloud App Security', 'Sign In Frequency', 'Persistent Browser', 'Continuous Access Evaluation', 'Resilient Defaults', 'Secure Sign In Session'
  )
  $chosenColumns = if ($CsvColumns) {
    $CsvColumns 
  }
  else {
    $defaultColumns 
  }
  $available = if ($CAExport.Count -gt 0) {
    $CAExport[0].PSObject.Properties.Name 
  }
  else {
    @() 
  }
  $missingCols = @(); $finalCols = @()
  foreach ($c in $chosenColumns) {
    if ($available -contains $c) {
      $finalCols += $c 
    }
    else {
      $missingCols += $c 
    } 
  }
  if ($missingCols) {
    Write-Warn "Ignoring unknown CSV columns: $($missingCols -join ', ')" 
  }
  $exportSet = if ($finalCols) {
    $CAExport | Select-Object -Property $finalCols 
  }
  else {
    $null 
  }
  if (-not $exportSet) {
    Write-Warn 'No CAExport data found; exporting raw policies.'; $CAPolicy | Select-Object * | Export-Csv -NoTypeInformation -Path $LaunchCsv 
  }
  else {
    $exportSet | Export-Csv -NoTypeInformation -Path $LaunchCsv 
  }
  Write-Info "CSV saved: $LaunchCsv"
}
if ($CsvPivotExport) {
  Write-Info 'Saving to File: Pivot CSV'
  $LaunchCsvPivot = Join-Path -Path $ExportLocation -ChildPath $CsvPivotFileName
  $pivotToExport = $pivot
  if (-not $pivotToExport -or $pivotToExport.Count -eq 0) {
    Write-Warn 'Pivot empty; skipping pivot CSV.' 
  }
  else {
    try {
      $pivotToExport | Export-Csv -NoTypeInformation -Path $LaunchCsvPivot
      Write-Info "Pivot CSV saved: $LaunchCsvPivot"
    }
    catch {
      Write-Warn "Failed to write pivot CSV: $($_.Exception.Message)"
    }
  }
}