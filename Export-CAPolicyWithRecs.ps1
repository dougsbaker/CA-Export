<#
.SYNOPSIS
    Exports Entra ID Conditional Access policies to an HTML report with security recommendations.

.DESCRIPTION
    Connects to Microsoft Graph and exports Conditional Access (CA) policies to a formatted HTML
    report. GUIDs are resolved to display names for users, groups, roles, applications, named
    locations, and Terms of Use. The report includes per-policy security checks and actionable
    recommendations based on Microsoft best practices.

    Required Microsoft Graph scopes:
        Policy.Read.All, Directory.Read.All, Application.Read.All,
        Agreement.Read.All, GroupMember.Read.All

    Required module:
        Microsoft.Graph (Install-Module Microsoft.Graph)

.PARAMETER PolicyID
    The GUID of a single Conditional Access policy to export. If omitted, all policies
    in the tenant are exported.

.EXAMPLE
    .\Export-CAPolicyWithRecs.ps1

    Exports all Conditional Access policies to CAPolicy.html in the script directory.

.EXAMPLE
    .\Export-CAPolicyWithRecs.ps1 -PolicyID 'xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx'

    Exports a single policy by its object ID.

.NOTES
    Author:   Douglas Baker (@dougsbaker)
    Version:  3.2
    License:  Creative Commons Attribution-NonCommercial-ShareAlike 4.0 International
              https://creativecommons.org/licenses/by-nc-sa/4.0/

    This script is provided AS IS without warranty of any kind and is not supported
    under any standard support program or service.

    Open-source components used in the HTML report:
        Bootstrap 4  - MIT License        - https://getbootstrap.com/docs/4.0/about/license/
        Font Awesome - CC BY 4.0 License  - https://fontawesome.com/license/free

.LINK
    https://learn.microsoft.com/en-us/entra/identity/conditional-access/overview
#>

[CmdletBinding()]
param (
    [Parameter()]
    [String]$PolicyID
)

$ExportLocation = $PSScriptRoot
if (!$ExportLocation) { $ExportLocation = $PWD }
$FileName = "\CAPolicy.html"
$JsonFileName = "\CAPolicy.json"
$HTMLExport = $true
$JsonExport = $false

function Invoke-MgCall {
    param([string]$Label, [scriptblock]$Call)
    try { return & $Call }
    catch {
        Write-Host "Error: Failed during '$Label'." -ForegroundColor Red
        if ($_.Exception.Message -match 'Assembly with same name|could not be loaded') {
            Write-Host "Cause: Microsoft.Graph module version conflict." -ForegroundColor Yellow
            Write-Host "Fix: Run 'Update-Module Microsoft.Graph -Force' in a new terminal, then retry." -ForegroundColor Yellow
        } else {
            Write-Host $_.Exception.Message -ForegroundColor Yellow
        }
        Exit
    }
}

$installedGraphModules = Get-InstalledModule -Name 'Microsoft.Graph*' -ErrorAction SilentlyContinue
if ($installedGraphModules) {
    $versionGroups = $installedGraphModules | Group-Object Version
    if ($versionGroups.Count -gt 1) {
        $versionList = ($versionGroups | ForEach-Object { $_.Name }) -join ', '
        Write-Host "Error: Multiple Microsoft.Graph module versions detected ($versionList)." -ForegroundColor Red
        Write-Host "This causes assembly-load conflicts. Fix by running in a new PowerShell window:" -ForegroundColor Yellow
        Write-Host "  Update-Module Microsoft.Graph -Force" -ForegroundColor Cyan
        Write-Host "Then close and reopen your terminal before re-running this script." -ForegroundColor Yellow
        Exit
    }
}


$RequiredScopes = @(
    'Policy.Read.All',
    'Directory.Read.All',
    'Application.Read.All',
    'Agreement.Read.All',
    'GroupMember.Read.All'
)

$context = Get-MgContext
$missingScopes = $RequiredScopes | Where-Object { $context.Scopes -notcontains $_ }

if ($null -eq $context -or $missingScopes.Count -gt 0) {
    if ($null -ne $context -and $missingScopes.Count -gt 0) {
        Write-Host "Connected but missing required scopes: $($missingScopes -join ', ')" -ForegroundColor Yellow
        Write-Host "Reconnecting with required scopes..." -ForegroundColor Yellow
    } else {
        Write-Host "Connecting: MgGraph"
    }
    try {
        Connect-MgGraph -Scopes $RequiredScopes -NoWelcome -ErrorAction Stop
    } catch {
        Write-Host "Error: Failed to connect. Ensure the Microsoft.Graph module is installed." -ForegroundColor Red
        Write-Host "Run: Install-Module Microsoft.Graph" -ForegroundColor Yellow
        Exit
    }

    $context = Get-MgContext
    $stillMissing = $RequiredScopes | Where-Object { $context.Scopes -notcontains $_ }
    if ($stillMissing.Count -gt 0) {
        Write-Host "Error: The following required scopes were not granted: $($stillMissing -join ', ')" -ForegroundColor Red
        Write-Host "Ensure your account or app registration has the necessary permissions." -ForegroundColor Yellow
        Exit
    }
}

Write-Host "Connected: MgGraph (Tenant: $($context.TenantId))"


$TenantData = Invoke-MgCall 'Get-MgOrganization' { Get-MgOrganization }
$TenantName = $TenantData.DisplayName
$date = Get-Date
Write-Host "Connected: $TenantName tenant"
$LinkURL = "https://portal.azure.com/#view/Microsoft_AAD_ConditionalAccess/PolicyBlade/policyId/"

#Collect CA Policy
Write-host "Exporting: CA Policy"
if ($PolicyID) {
    $CAPolicy = Invoke-MgCall 'Get-MgIdentityConditionalAccessPolicy' { Get-MgIdentityConditionalAccessPolicy -ConditionalAccessPolicyId $PolicyID }
}
else {
    $CAPolicy = Invoke-MgCall 'Get-MgIdentityConditionalAccessPolicy' { Get-MgIdentityConditionalAccessPolicy -all -ExpandProperty * }
}


Write-host "Extracting: Names from Guid's"
#Swap User Guid With Names
#Get Name
$ADUsers = $CAPolicy.Conditions.Users.IncludeUsers
$ADUsers += $CAPolicy.Conditions.Users.IncludeGroups
$ADUsers += $CAPolicy.Conditions.Users.IncludeRoles
$ADUsers += $CAPolicy.Conditions.Users.ExcludeUsers
$ADUsers += $CAPolicy.Conditions.Users.ExcludeGroups
$ADUsers += $CAPolicy.Conditions.Users.ExcludeRoles



# Filter the $AdUsers array to include only valid GUIDs
$ADsearch = $AdUsers | Where-Object {
    ([Guid]::TryParse($_, [ref] [Guid]::Empty))
}

#users Hashtable
$mgobjects = Invoke-MgCall 'Get-MgDirectoryObjectById' { Get-MgDirectoryObjectById -ids $ADsearch }
$mgObjectsLookup = @{}
foreach ($obj in $mgobjects) {
    $mgObjectsLookup[$obj.Id] = $obj.AdditionalProperties.displayName
}


#Change Guid's
foreach ($policy in $caPolicy) {
    # Check if the policy has Conditions and Users and ExcludeUsers properties
    if ($policy.Conditions -and $policy.Conditions.Users <#-and $policy.Conditions.Users.ExcludeUsers#>) {
        # Loop through each user in ExcludeUsers and replace with displayName if found in $mgObjectsLookup
        for ($i = 0; $i -lt $policy.Conditions.Users.ExcludeUsers.Count; $i++) {
            $userId = $policy.Conditions.Users.ExcludeUsers[$i]
            if ($mgObjectsLookup.ContainsKey($userId)) {
                $policy.Conditions.Users.ExcludeUsers[$i] = $mgObjectsLookup[$userId]
            }
        }

        for ($i = 0; $i -lt $policy.Conditions.Users.IncludeUsers.Count; $i++) {
            $userId = $policy.Conditions.Users.IncludeUsers[$i]
            if ($mgObjectsLookup.ContainsKey($userId)) {
                $policy.Conditions.Users.IncludeUsers[$i] = $mgObjectsLookup[$userId]
            }
        }
        for ($i = 0; $i -lt $policy.Conditions.Users.IncludeGroups.Count; $i++) {
            $userId = $policy.Conditions.Users.IncludeGroups[$i]
            if ($mgObjectsLookup.ContainsKey($userId)) {
                $memberCount = 0
                $groupMembers = Get-MgGroupMember -GroupId $userId
                $memberCount = $groupMembers | Measure-Object | Select-Object -ExpandProperty Count
                $policy.Conditions.Users.IncludeGroups[$i] = $mgObjectsLookup[$userId] + " ($memberCount)"
            }
        }
        for ($i = 0; $i -lt $policy.Conditions.Users.ExcludeGroups.Count; $i++) {
            $userId = $policy.Conditions.Users.ExcludeGroups[$i]
            if ($mgObjectsLookup.ContainsKey($userId)) {
                $memberCount = 0
                $groupMembers = Get-MgGroupMember -GroupId $userId
                $memberCount = $groupMembers | Measure-Object | Select-Object -ExpandProperty Count
                $policy.Conditions.Users.ExcludeGroups[$i] = $mgObjectsLookup[$userId] + " ($memberCount)"
            }
        }
        for ($i = 0; $i -lt $policy.Conditions.Users.IncludeRoles.Count; $i++) {
            $userId = $policy.Conditions.Users.IncludeRoles[$i]
            if ($mgObjectsLookup.ContainsKey($userId)) {
                $policy.Conditions.Users.IncludeRoles[$i] = $mgObjectsLookup[$userId]
            }
        }
        for ($i = 0; $i -lt $policy.Conditions.Users.ExcludeRoles.Count; $i++) {
            $userId = $policy.Conditions.Users.ExcludeRoles[$i]
            if ($mgObjectsLookup.ContainsKey($userId)) {
                $policy.Conditions.Users.ExcludeRoles[$i] = $mgObjectsLookup[$userId]
            }
        }
    }
   
}

#Swap App ID with name
$MGApps = Invoke-MgCall 'Get-MgServicePrincipal' { Get-MgServicePrincipal -All }
#Hash Table
$MGAppsLookup = @{}
foreach ($obj in $MGApps) {
    $MGAppsLookup[$obj.AppId] = "<div class='tooltip-container'>" + $obj.DisplayName + "<span class='tooltip-text'>App Id:" + $obj.AppId + "</span></div>"
}
#"<div class='tooltip-container'>" + $obj.DisplayName +"<span class='tooltip-text'>App Id:"+ $obj.AppId +"</span></div>"

foreach ($policy in $caPolicy) {
    if ($policy.Conditions -and $policy.Conditions.Applications) {

        for ($i = 0; $i -lt $policy.Conditions.Applications.ExcludeApplications.Count; $i++) {
            $AppId = $policy.Conditions.Applications.ExcludeApplications[$i]
            if ($MGAppsLookup.ContainsKey($AppId)) {
                $policy.Conditions.Applications.ExcludeApplications[$i] = $MGAppsLookup[$AppId]
            }
        }
       
        for ($i = 0; $i -lt $policy.Conditions.Applications.IncludeApplications.Count; $i++) {
            $AppId = $policy.Conditions.Applications.IncludeApplications[$i]
            if ($MGAppsLookup.ContainsKey($AppId)) {
                $policy.Conditions.Applications.IncludeApplications[$i] = $MGAppsLookup[$AppId]
            }
        }
    }
   
}

#Swap Location with Names
$mgLoc = Invoke-MgCall 'Get-MgIdentityConditionalAccessNamedLocation' { Get-MgIdentityConditionalAccessNamedLocation }
$MGLocLookup = @{}
foreach ($obj in $mgLoc) {
    $MGLocLookup[$obj.Id] = $obj.DisplayName
}
foreach ($policy in $caPolicy) {
    #Set Locations
    if ($policy.Conditions -and $policy.Conditions.Locations) {
        for ($i = 0; $i -lt $policy.Conditions.Locations.IncludeLocations.Count; $i++) {
            $LocId = $policy.Conditions.Locations.IncludeLocations[$i]
            if ($MGLocLookup.ContainsKey($LocId)) {
                $policy.Conditions.Locations.IncludeLocations[$i] = $MGLocLookup[$LocId]
            }
        }
        for ($i = 0; $i -lt $policy.Conditions.Locations.ExcludeLocations.Count; $i++) {
            $LocId = $policy.Conditions.Locations.ExcludeLocations[$i]
            if ($MGLocLookup.ContainsKey($LocId)) {
                $policy.Conditions.Locations.ExcludeLocations[$i] = $MGLocLookup[$LocId]
            }
        }
    }
 
}

#Switch TOU Id for Name
$mgTou = Invoke-MgCall 'Get-MgAgreement' { Get-MgAgreement }
$MGTouLookup = @{}
foreach ($obj in $mgTou) {
    $MGTouLookup[$obj.Id] = $obj.DisplayName
}
foreach ($policy in $caPolicy) {
    if ($policy.GrantControls -and $policy.GrantControls.TermsOfUse) {
              
        for ($i = 0; $i -lt $policy.GrantControls.TermsOfUse.Count; $i++) {
            $TouId = $policy.GrantControls.TermsOfUse[$i]
            if ($MGTouLookup.ContainsKey($TouId)) {
                $policy.GrantControls.TermsOfUse[$i] = $MGTouLookup[$TouId]
            }
        }
    }
}
#swap Admin Roles
$mgRole = Invoke-MgCall 'Get-MgDirectoryRoleTemplate' { Get-MgDirectoryRoleTemplate }
$mgRoleLookup = @{}
foreach ($obj in $mgRole) {
    $mgRoleLookup[$obj.Id] = $obj.DisplayName
}
foreach ($policy in $caPolicy) {
    if ($policy.Conditions.Users -and $policy.Conditions.Users.IncludeRoles) {
              
        for ($i = 0; $i -lt $policy.Conditions.Users.IncludeRoles.Count; $i++) {
            $RoleId = $policy.Conditions.Users.IncludeRoles[$i]
            if ($mgRoleLookup.ContainsKey($RoleId)) {
                $policy.Conditions.Users.IncludeRoles[$i] = $mgRoleLookup[$RoleId]
            }
        }
    }
    if ($policy.Conditions.Users -and $policy.Conditions.Users.ExcludeRoles) {
              
        for ($i = 0; $i -lt $policy.Conditions.Users.ExcludeRoles.Count; $i++) {
            $RoleId = $policy.Conditions.Users.ExcludeRoles[$i]
            if ($mgRoleLookup.ContainsKey($RoleId)) {
                $policy.Conditions.Users.ExcludeRoles[$i] = $mgRoleLookup[$RoleId]
            }
        }
    }
}

# exit
$CAExport = [PSCustomObject]@()

#Extract Values
Write-host "Extracting: CA Policy Data"
foreach ( $Policy in $CAPolicy) {

    $IncludeUG = $null
    $IncludeUG = $Policy.Conditions.Users.IncludeUsers
    $IncludeUG += $Policy.Conditions.Users.IncludeGroups
    $IncludeUG += $Policy.Conditions.Users.IncludeRoles
    $IncludeUG += $policy.conditions.users.IncludeGuestsOrExternalUsers.GuestOrExternalUserTypes -replace ',', ', '
    $DateModified = $null
    $DateModified = $Policy.ModifiedDateTime
    

    $ExcludeUG = $null
    $ExcludeUG = $Policy.Conditions.Users.ExcludeUsers
    $ExcludeUG += $Policy.Conditions.Users.ExcludeGroups
    $ExcludeUG += $Policy.Conditions.Users.ExcludeRoles
    $ExcludeUG += $policy.conditions.users.ExcludeGuestsOrExternalUsers.GuestOrExternalUserTypes -replace ',', ', '
    

    $InclLocation = $Null
    $ExclLocation = $Null 
    $InclLocation = $Policy.Conditions.Locations.includelocations
    $ExclLocation = $Policy.Conditions.Locations.Excludelocations

    $InclPlat = $Null
    $ExclPlat = $Null 
    $InclPlat = $Policy.Conditions.Platforms.IncludePlatforms
    $ExclPlat = $Policy.Conditions.Platforms.ExcludePlatforms
    $InclDev = $null
    $ExclDev = $null
    $InclDev = $Policy.Conditions.Devices.IncludeDevices
    $ExclDev = $Policy.Conditions.Devices.ExcludeDevices
    $devFilters = $null
    $devFilters = $Policy.Conditions.Devices.DeviceFilter.Rule

    $authenticationFlowsString = $Policy.Conditions.AdditionalProperties.authenticationFlows.Values -join ', '
    
 
    $CAExport += New-Object PSObject -Property @{ 
        Name                            = $Policy.DisplayName;
        Status                          = $Policy.State;
        DateModified                    = $DateModified;
        Users                           = "";
        UsersInclude                    = ($IncludeUG -join ", `r`n");
        UsersExclude                    = ($ExcludeUG -join ", `r`n");
        'Cloud apps or actions'         = "";
        ApplicationsIncluded            = ($Policy.Conditions.Applications.IncludeApplications -join ", `r`n");
        ApplicationsExcluded            = ($Policy.Conditions.Applications.ExcludeApplications -join ", `r`n");
        userActions                     = ($Policy.Conditions.Applications.IncludeUserActions -join ", `r`n");
        AuthContext                     = ($Policy.Conditions.Applications.IncludeAuthenticationContextClassReferences -join ", `r`n");
        Conditions                      = "";
        UserRisk                        = ($Policy.Conditions.UserRiskLevels -join ", `r`n");
        SignInRisk                      = ($Policy.Conditions.SignInRiskLevels -join ", `r`n");
        # Platforms = $Policy.Conditions.Platforms;
        PlatformsInclude                = ($InclPlat -join ", `r`n");
        PlatformsExclude                = ($ExclPlat -join ", `r`n");
        # Locations = $Policy.Conditions.Locations;
        LocationsIncluded               = ($InclLocation -join ", `r`n");
        LocationsExcluded               = ($ExclLocation -join ", `r`n");
        ClientApps                      = ($Policy.Conditions.ClientAppTypes -join ", `r`n");
        # Devices = $Policy.Conditions.Devices;
        DevicesIncluded                 = ($InclDev -join ", `r`n");
        DevicesExcluded                 = ($ExclDev -join ", `r`n");
        DeviceFilters                   = ($devFilters -join ", `r`n");
        AuthenticationFlows             = $authenticationFlowsString ;
        'Grant Controls'                = "";
        # Grant = ($Policy.GrantControls.BuiltInControls -join ", `r`n");
        Block                           = if ($Policy.GrantControls.BuiltInControls -contains "Block") { "True" } else { "" }
        'Require MFA'                   = if ($Policy.GrantControls.BuiltInControls -contains "Mfa") { "True" } else { "" }
        'Authentication Strength MFA'   = $Policy.GrantControls.AuthenticationStrength.DisplayName
        'CompliantDevice'               = if ($Policy.GrantControls.BuiltInControls -contains "CompliantDevice") { "True" } else { "" }
        'DomainJoinedDevice'            = if ($Policy.GrantControls.BuiltInControls -contains "DomainJoinedDevice") { "True" } else { "" }
        'CompliantApplication'          = if ($Policy.GrantControls.BuiltInControls -contains "CompliantApplication") { "True" } else { "" }
        'ApprovedApplication'           = if ($Policy.GrantControls.BuiltInControls -contains "ApprovedApplication") { "True" } else { "" }
        'PasswordChange'                = if ($Policy.GrantControls.BuiltInControls -contains "PasswordChange") { "True" } else { "" }
        TermsOfUse                      = ($Policy.GrantControls.TermsOfUse -join ", `r`n");
        CustomControls                  = ($Policy.GrantControls.CustomAuthenticationFactors -join ", `r`n");
        GrantOperator                   = $Policy.GrantControls.Operator
        # Session = $Policy.SessionControls
        'Session Controls'              = "";
        ApplicationEnforcedRestrictions = $Policy.SessionControls.ApplicationEnforcedRestrictions.IsEnabled
        CloudAppSecurity                = $Policy.SessionControls.CloudAppSecurity.IsEnabled
        SignInFrequency                 = "$($Policy.SessionControls.SignInFrequency.Value) $($Policy.SessionControls.SignInFrequency.Type)"
        PersistentBrowser               = $Policy.SessionControls.PersistentBrowser.Mode
        ContinuousAccessEvaluation      = $Policy.SessionControls.ContinuousAccessEvaluation.Mode
        ResiliantDefaults               = $policy.SessionControls.DisableResilienceDefaults
        secureSignInSession             = $policy.SessionControls.AdditionalProperties.secureSignInSession.Values
    }
    
}


#Export Setup
Write-host "Pivoting: CA to Export Format"
$pivot = @()
$rowItem = New-Object PSObject
$rowitem | Add-Member -type NoteProperty -Name 'CA Item' -Value "row1"
$Pcount = 1
foreach ($CA in $CAExport) {
    $rowitem | Add-Member -type NoteProperty -Name "Policy $pcount" -Value "row1"
    #$ca.Name
    $pcount += 1
}
$pivot += $rowItem

#Add Data to Report
$Rows = $CAExport | Get-Member | Where-Object { $_.MemberType -eq "NoteProperty" }
$Rows | ForEach-Object {
    $rowItem = New-Object PSObject
    $rowname = $_.Name
    $rowitem | Add-Member -type NoteProperty -Name 'CA Item' -Value $_.Name
    $Pcount = 1
    foreach ($CA in $CAExport) {
        $ca | Get-Member | Where-Object { $_.MemberType -eq "NoteProperty" } | ForEach-Object {
            $a = $_.name
            $b = $ca.$a
            if ($a -eq $rowname) {
                $rowitem | Add-Member -type NoteProperty -Name "Policy $pcount" -Value $b  
            }
            
        }
        # $ca.UsersInclude
        $pcount += 1
    }
    $pivot += $rowItem
}
#Export Setup
Write-host "Analyzing: getting recommendations"

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
        $this.Note = ""
        $this.Excluded = @()
        
    }
}


$recommendations = @(
    [Recommendation]::new(
        "CA-00", 
        "Legacy Authentication", 
        "Legacy Authentication is blocked or minimized, targeting Legacy Authentication protocols.", 
        "Review and update policies to restrict or block Legacy Authentication protocols to ensure security.", 
        "Legacy Authentication protocols are outdated and less secure. It is recommended to block or minimize their usage to enhance the security of your environment.", 
        @{"Legacy Authentication Overview" = "https://learn.microsoft.com/en-us/entra/identity/conditional-access/policy-block-legacy-authentication" }, 
        $false, 
        $true
    ),
    [Recommendation]::new(
        "CA-01", 
        "MFA Policy targets All users Group and All Cloud Apps", 
        "There is at least one policy that targets all users and cloud apps.", 
        "Review and update MFA policies to ensure they target all users and cloud apps, including any necessary exclusions.", 
        "Multi-factor Authentication (MFA) should apply to all users and cloud apps as a baseline for security. Policies should include the necessary exclusions if required but should primarily target all users and apps for maximum security.", 
        @{"The Challenge with Targeted Architecture" = "https://learn.microsoft.com/en-us/azure/architecture/guide/security/conditional-access-architecture#:~:text=The%20challenge%20with%20the,that%20number%20isn%27t%20supported." }, 
        $false, 
        $true
    ),
    [Recommendation]::new(
        "CA-02", 
        "Mobile Device Policy requires MDM or MAM", 
        "There is at least one policy that requires MDM or MAM for mobile devices.", 
        "Consider adding policies to check for device management, either through MDM or MAM, to ensure secure mobile access.", 
        "Mobile Device Management (MDM) or Mobile Application Management (MAM) should be enforced to ensure that mobile devices accessing organizational data are properly managed and secure. Policies should include requirements for MDM or MAM to increase security for mobile devices.", 
        @{"MAM Overview"                               = "https://learn.microsoft.com/en-us/mem/intune/apps/app-management#mobile-application-management-mam-basics"
            "Protect Data on personally owned devices" = "https://smbtothecloud.com/protecting-company-data-on-personally-owned-devices/" 
        }, 
        $false, 
        $true
    ),
    [Recommendation]::new(
        "CA-03", 
        "Require Hybrid Join or Intune Compliance on Windows or Mac", 
        "There is at least one policy that requires Hybrid Join or Intune Compliance for Windows or Mac devices.", 
        "Consider adding policies to ensure that Windows or Mac devices are either Hybrid Joined or compliant with Intune to enhance security.", 
        "Hybrid Join or Intune Compliance should be enforced to ensure that Windows or Mac devices accessing organizational data are properly managed and secure. Policies should include requirements for Hybrid Join or Intune Compliance to increase security for these devices.", 
        @{
            "Hybrid Join Overview"       = "https://learn.microsoft.com/en-us/azure/active-directory/devices/hybrid-azuread-join-plan"
            "Intune Compliance Overview" = "https://learn.microsoft.com/en-us/mem/intune/protect/compliance-policy-create-windows"
        }, 
        $false, 
        $true
    ),
    [Recommendation]::new(
        "CA-04", 
        "Require MFA for Admins", 
        "There is at least one policy that requires Multi-Factor Authentication (MFA) for administrators.", 
        "Consider adding policies to ensure that administrators are required to use Multi-Factor Authentication (MFA) to enhance security.", 
        "Multi-Factor Authentication (MFA) should be enforced for administrators to ensure that access to critical systems and data is secure. Policies should include requirements for MFA to increase security for administrative accounts. Policies should target the folowing roles Global Administrator, Security Administrator, SharePoint Administrator, Exchange Administrator, Conditional Access Administrator, Helpdesk Administrator, Billing Administrator, User Administrator, Authentication Administrator, Application Administrator, Cloud Application Administrator, Password Administrator, Privileged Authentication Administrator, Privileged Role Administrator", 
        @{
            "MFA Overview"   = "https://learn.microsoft.com/en-us/azure/active-directory/authentication/concept-mfa-howitworks"
            "MFA for Admins" = "https://learn.microsoft.com/en-us/entra/identity/conditional-access/policy-old-require-mfa-admin"
        }, 
        $false, 
        $true
    ),
    [Recommendation]::new(
        "CA-05", 
        "Require Phish-Resistant MFA for Admins", 
        "There is at least one policy that requires phish-resistant Multi-Factor Authentication (MFA) for administrators.", 
        "Consider adding policies to ensure that administrators are required to use phish-resistant Multi-Factor Authentication (MFA) to enhance security.", 
        "Phish-resistant Multi-Factor Authentication (MFA) should be enforced for administrators to ensure that access to critical systems and data is secure. Policies should include requirements for phish-resistant MFA to increase security for administrative accounts. Policies should target the following roles: Global Administrator, Security Administrator, SharePoint Administrator, Exchange Administrator, Conditional Access Administrator, Helpdesk Administrator, Billing Administrator, User Administrator, Authentication Administrator, Application Administrator, Cloud Application Administrator, Password Administrator, Privileged Authentication Administrator, Privileged Role Administrator.", 
        @{
            "MSFT Authentication Strengths"  = "https://learn.microsoft.com/en-us/entra/identity/authentication/concept-authentication-strengths"
            "Phish-Resistant MFA for Admins" = "https://learn.microsoft.com/en-us/entra/identity/conditional-access/policy-admin-phish-resistant-mfa"
        }, 
        $false, 
        $true
    ),
    [Recommendation]::new(
        "CA-06", 
        "Policy Excludes Same Entities It Includes", 
        "There is at least one policy that excludes the same entities it includes, resulting in no effective condition being checked.", 
        "Review and update policies to ensure that they do not exclude the same entities they include, as this results in no effective condition being checked.", 
        "Policies should be configured to include and exclude distinct sets of entities to ensure that conditions are effectively checked. This helps in maintaining the integrity and effectiveness of the policy.", 
        @{
            "Policy Configuration Best Practices" = "https://learn.microsoft.com/en-us/azure/active-directory/conditional-access/best-practices"
        }, 
        $true, 
        $false
    )
    [Recommendation]::new(
        "CA-07", 
        "No Users Targeted in Policy", 
        "There is at least one policy that does not target any users.", 
        "Review and update policies to ensure that they target specific users, groups, or roles to be effective.", 
        "Policies should be configured to target specific users, groups, or roles to ensure that they are applied correctly and provide the intended security controls.", 
        @{
            "Policy Configuration Best Practices" = "https://learn.microsoft.com/en-us/azure/active-directory/conditional-access/best-practices"
        }, 
        $true, 
        $false
    ),
    [Recommendation]::new(
        "CA-08", 
        "Direct User Assignment", 
        "There are no direct user assignments in the policy.", 
        "Review and update policies to avoid direct user assignments and instead use exclusion groups to manage user access more efficiently.", 
        "Direct user assignments in policies are not ideal for maintaining flexibility and scalability. Exclusion groups should be used instead to manage policies efficiently without manually adding users to each policy.", 
        @{}, 
        $true, 
        $false
    ),
    [Recommendation]::new(
        "CA-09", 
        "Implement Risk-Based Policy", 
        "There is at least 1 policy that addresses risk-based conditional access.", 
        "Consider implementing risk-based conditional access policies to enhance security by dynamically applying access controls based on the risk level of the sign-in or user.", 
        "Risk-based policies help in dynamically assessing the risk level of sign-ins and users, and applying appropriate access controls to mitigate potential threats. This ensures that high-risk activities are subject to stricter controls, thereby enhancing the overall security posture.", 
        @{
            "Risk-Based Conditional Access Overview"  = "https://learn.microsoft.com/en-us/entra/id-protection/howto-identity-protection-configure-risk-policies"
            "Require MFA for Risky Sign-in"           = "https://learn.microsoft.com/en-us/entra/identity/conditional-access/policy-risk-based-sign-in#enable-with-conditional-access-policy"
            "Require Passsword Change for Risky USer" = "https://learn.microsoft.com/en-us/entra/identity/conditional-access/policy-risk-based-user#enable-with-conditional-access-policy"
        }, 
        $false, 
        $true
    ),
    [Recommendation]::new(
        "CA-10", 
        "Block Device Code Flow", 
        "There is at least 1 policy that blocks device code flow.", 
        "Consider implementing a policy to block device code flow to enhance security by preventing unauthorized access through device code authentication.", 
        "Blocking device code flow helps in preventing unauthorized access through device code authentication, which can be exploited by attackers. Implementing this policy ensures that only secure authentication methods are used.", 
        @{
            "Block Device Code Flow Overview" = "https://learn.microsoft.com/en-us/entra/identity/conditional-access/concept-authentication-flows#device-code-flow"
        }, 
        $false, 
        $true
    ),
    [Recommendation]::new(
        "CA-11", 
        "Require MFA to Enroll a Device in Intune", 
        "There is at least 1 policy that requires Multi-Factor Authentication (MFA) to enroll a device in Intune.", 
        "Consider implementing a policy to require Multi-Factor Authentication (MFA) for enrolling devices in Intune to enhance security.", 
        "Requiring MFA for device enrollment in Intune ensures that only authorized users can enroll devices, thereby enhancing the security of your organization's mobile device management.", 
        @{
            "MFA for Intune Enrollment Overview" = "https://learn.microsoft.com/en-us/mem/intune/enrollment/multi-factor-authentication"
        }, 
        $false, 
        $true
    ),
    [Recommendation]::new(
        "CA-12", 
        "Block Unknown/Unsupported Devices", 
        "There is no policy that blocks unknown or unsupported devices.", 
        "Consider implementing a policy to block unknown or unsupported devices to enhance security by preventing unauthorized access from devices that do not meet your organization's security standards.", 
        "Blocking unknown or unsupported devices helps in preventing unauthorized access from devices that may not comply with your organization's security policies. Implementing this policy ensures that only secure and compliant devices can access organizational resources.", 
        @{
            "Block Unknown/Unsupported Devices Overview" = "https://learn.microsoft.com/en-us/entra/identity/conditional-access/policy-all-users-device-unknown-unsupported"
        }, 
        $false, 
        $true
    )
)

function Update-PolicyStatus {
    param (
        [ref]$Recommendation,
        $PolicyCheck,
        $StatusCheck
    )

    if (&$StatusCheck $PolicyCheck) {
        $Recommendation.Value.Status = $Recommendation.Value.SwapStatus
    }

    if ($Recommendation.Value.Status -and $PolicyCheck.state -eq "enabled") {
        $Status1 = "policy-item success"
        $Status2 = "status-icon-large success"
        $Status3 = "fas fa-check-circle"
    }
    else {
        $Status1 = "policy-item warning"
        $Status2 = "status-icon-large warning"
        $Status3 = "fas fa-exclamation-triangle"
    }
    
    $CheckExcUG = $PolicyCheck.Conditions.Users.ExcludeUsers + $PolicyCheck.Conditions.Users.ExcludeGroups + $PolicyCheck.Conditions.Users.ExcludeRoles + $PolicyCheck.conditions.users.ExcludeGuestsOrExternalUsers.GuestOrExternalUserTypes -replace ',', ', '
    $CheckIncUG = $PolicyCheck.Conditions.Users.IncludeUsers + $PolicyCheck.Conditions.Users.IncludeGroups + $PolicyCheck.Conditions.Users.IncludeRoles + $PolicyCheck.conditions.users.IncludeGuestsOrExternalUsers.GuestOrExternalUserTypes -replace ',', ', '
    $CheckIncCond = $PolicyCheck.Conditions.Locations.includelocations + $PolicyCheck.Conditions.Platforms.IncludePlatforms
    $CheckExcCond = $PolicyCheck.Conditions.Locations.Excludelocations + $PolicyCheck.Conditions.Platforms.ExcludePlatforms
    $CheckGrant = $PolicyCheck.GrantControls.BuiltInControls + $PolicyCheck.GrantControls.AuthenticationStrength.DisplayName + $PolicyCheck.GrantControls.CustomAuthenticationFactors + $PolicyCheck.GrantControls.TermsOfUse
    $checkSession = ""
    
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
                <strong>$($PolicyCheck.DisplayName) <a href='$($LinkURL)$($PolicyCheck.Id)' target='_blank'><i class='fas fa-external-link-alt'></i></a></strong>
                <div class='recommendation-status'>Status: $($PolicyCheck.state)</div>
                <div class='$($Status2)'><i class='$($Status3)'></i></div>
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
    "CA-00" = {
        param($PolicyCheck)
        $PolicyCheck.GrantControls.BuiltInControls -contains "Block" -and 
        $PolicyCheck.Conditions.ClientAppTypes -contains "exchangeActiveSync" -and 
        $PolicyCheck.Conditions.ClientAppTypes -contains "other"
    }
    "CA-01" = {
        param($PolicyCheck)
        $PolicyCheck.GrantControls.BuiltInControls -contains "Mfa" -and 
        $PolicyCheck.Conditions.Users.IncludeUsers -eq "all" -and 
        $PolicyCheck.Conditions.Applications.IncludeApplications -eq "all"
    }
    "CA-02" = {
        param($PolicyCheck)
        ($PolicyCheck.Conditions.Platforms.IncludePlatforms -contains "android" -or 
        $PolicyCheck.Conditions.Platforms.IncludePlatforms -contains "iOS" -or 
        $PolicyCheck.Conditions.Platforms.IncludePlatforms -contains "windowsPhone") -and
        ($PolicyCheck.GrantControls.BuiltInControls -contains "approvedApplication" -or
        $PolicyCheck.GrantControls.BuiltInControls -contains "compliantApplication" -or
        $PolicyCheck.GrantControls.BuiltInControls -contains "compliantDevice")
    }
    "CA-03" = {
        param($PolicyCheck)
        ($PolicyCheck.Conditions.Platforms.IncludePlatforms -contains "windows" -or 
        $PolicyCheck.Conditions.Platforms.IncludePlatforms -contains "macOS") -and
        ($PolicyCheck.GrantControls.BuiltInControls -contains "compliantDevice" -or
        $PolicyCheck.GrantControls.BuiltInControls -contains "domainJoinedDevice")
    }
    
    "CA-04" = {
        param($PolicyCheck)
        ($PolicyCheck.Conditions.Users.IncludeRoles -contains "Privileged Role Administrator" -or
        $PolicyCheck.Conditions.Users.IncludeRoles -contains "Global Administrator" -or
        $PolicyCheck.Conditions.Users.IncludeRoles -contains "Privileged Authentication Administrator" -or
        $PolicyCheck.Conditions.Users.IncludeRoles -contains "Security Administrator" -or
        $PolicyCheck.Conditions.Users.IncludeRoles -contains "SharePoint Administrator" -or
        $PolicyCheck.Conditions.Users.IncludeRoles -contains "Exchange Administrator" -or
        $PolicyCheck.Conditions.Users.IncludeRoles -contains "Conditional Access Administrator" -or
        $PolicyCheck.Conditions.Users.IncludeRoles -contains "Helpdesk Administrator" -or
        $PolicyCheck.Conditions.Users.IncludeRoles -contains "Billing Administrator" -or
        $PolicyCheck.Conditions.Users.IncludeRoles -contains "User Administrator" -or
        $PolicyCheck.Conditions.Users.IncludeRoles -contains "Authentication Administrator" -or
        $PolicyCheck.Conditions.Users.IncludeRoles -contains "Application Administrator" -or
        $PolicyCheck.Conditions.Users.IncludeRoles -contains "Cloud Application Administrator" -or
        $PolicyCheck.Conditions.Users.IncludeRoles -contains "Password Administrator") -and
        ($PolicyCheck.GrantControls.BuiltInControls -contains "Mfa" -or
        $PolicyCheck.GrantControls.AuthenticationStrength.DisplayName -contains "Phishing-resistant MFA" -or
        $PolicyCheck.GrantControls.AuthenticationStrength.DisplayName -contains "Passwordless MFA" -or
        $PolicyCheck.GrantControls.AuthenticationStrength.DisplayName -contains "Multifactor authentication")
    }
    "CA-05" = {
        param($PolicyCheck)
        ($PolicyCheck.Conditions.Users.IncludeRoles -contains "Privileged Role Administrator" -or
        $PolicyCheck.Conditions.Users.IncludeRoles -contains "Global Administrator" -or
        $PolicyCheck.Conditions.Users.IncludeRoles -contains "Privileged Authentication Administrator" -or
        $PolicyCheck.Conditions.Users.IncludeRoles -contains "Security Administrator" -or
        $PolicyCheck.Conditions.Users.IncludeRoles -contains "SharePoint Administrator" -or
        $PolicyCheck.Conditions.Users.IncludeRoles -contains "Exchange Administrator" -or
        $PolicyCheck.Conditions.Users.IncludeRoles -contains "Conditional Access Administrator" -or
        $PolicyCheck.Conditions.Users.IncludeRoles -contains "Helpdesk Administrator" -or
        $PolicyCheck.Conditions.Users.IncludeRoles -contains "Billing Administrator" -or
        $PolicyCheck.Conditions.Users.IncludeRoles -contains "User Administrator" -or
        $PolicyCheck.Conditions.Users.IncludeRoles -contains "Authentication Administrator" -or
        $PolicyCheck.Conditions.Users.IncludeRoles -contains "Application Administrator" -or
        $PolicyCheck.Conditions.Users.IncludeRoles -contains "Cloud Application Administrator" -or
        $PolicyCheck.Conditions.Users.IncludeRoles -contains "Password Administrator") -and
        ($PolicyCheck.GrantControls.AuthenticationStrength.DisplayName -contains "Phishing-resistant MFA")
    }
    "CA-06" = {
        param($PolicyCheck)
        ($PolicyCheck.Conditions.Users.IncludeUsers -ne $null -and $PolicyCheck.Conditions.Users.ExcludeUsers -ne $null -and 
        ($PolicyCheck.Conditions.Users.IncludeUsers | ForEach-Object { $PolicyCheck.Conditions.Users.ExcludeUsers -contains $_ })) -or
        ($PolicyCheck.Conditions.Users.IncludeGroups -ne $null -and $PolicyCheck.Conditions.Users.ExcludeGroups -ne $null -and 
        ($PolicyCheck.Conditions.Users.IncludeGroups | ForEach-Object { $PolicyCheck.Conditions.Users.ExcludeGroups -contains $_ })) -or
        ($PolicyCheck.Conditions.Users.IncludeRoles -ne $null -and $PolicyCheck.Conditions.Users.ExcludeRoles -ne $null -and 
        ($PolicyCheck.Conditions.Users.IncludeRoles | ForEach-Object { $PolicyCheck.Conditions.Users.ExcludeRoles -contains $_ })) -or
        ($PolicyCheck.Conditions.Platforms.IncludePlatforms -ne $null -and $PolicyCheck.Conditions.Platforms.ExcludePlatforms -ne $null -and 
        ($PolicyCheck.Conditions.Platforms.IncludePlatforms | ForEach-Object { $PolicyCheck.Conditions.Platforms.ExcludePlatforms -contains $_ })) -or
              ($PolicyCheck.Conditions.Locations.IncludeLocations -ne $null -and $PolicyCheck.Conditions.Locations.ExcludeLocations -ne $null -and 
        ($PolicyCheck.Conditions.Locations.IncludeLocations | ForEach-Object { $PolicyCheck.Conditions.Locations.ExcludeLocations -contains $_ })) -or
        ($PolicyCheck.Conditions.Applications.IncludeApplications -ne $null -and $PolicyCheck.Conditions.Applications.ExcludeApplications -ne $null -and 
            ($PolicyCheck.Conditions.Applications.IncludeApplications | ForEach-Object { $PolicyCheck.Conditions.Applications.ExcludeApplications -contains $_ }))
    }
    "CA-07" = {
        param($PolicyCheck)
        ($PolicyCheck.Conditions.Users.IncludeUsers -eq $null -or $PolicyCheck.Conditions.Users.IncludeUsers.Count -eq 0 -or $PolicyCheck.Conditions.Users.IncludeUsers -eq "None") -and
        ($PolicyCheck.Conditions.Users.IncludeGroups -eq $null -or $PolicyCheck.Conditions.Users.IncludeGroups.Count -eq 0 -or
        ($PolicyCheck.Conditions.Users.IncludeGroups | ForEach-Object { $_ -match '\((\d+)\)' -and [int]$matches[1] -eq 0 })) -and
        ($PolicyCheck.Conditions.Users.IncludeRoles -eq $null -or $PolicyCheck.Conditions.Users.IncludeRoles.Count -eq 0) -and
        ($PolicyCheck.conditions.users.IncludeGuestsOrExternalUsers.GuestOrExternalUserTypes -eq $null)
    }
    "CA-08" = {
        param($PolicyCheck)
        $PolicyCheck.Conditions.Users.IncludeUsers -ne "None" -and 
        $PolicyCheck.Conditions.Users.IncludeUsers -ne $null -and 
        $PolicyCheck.Conditions.Users.IncludeUsers -ne "All" -and 
        $PolicyCheck.Conditions.Users.IncludeUsers -ne "GuestsOrExternalUsers"
    }
    "CA-09" = {
        param($PolicyCheck)
        ($PolicyCheck.Conditions.SignInRiskLevels -ne $null) -or
        ($PolicyCheck.Conditions.UserRiskLevels -ne $null)
    }
    "CA-10" = {
        param($PolicyCheck)
        $PolicyCheck.Conditions.AdditionalProperties.authenticationFlows.Values -split ',' -contains "deviceCodeFlow" -and
        $PolicyCheck.grantcontrols.BuiltInControls -contains "Block"
    }
    "CA-11" = {
        param($PolicyCheck)
        ($PolicyCheck.Conditions.Applications.IncludeUserActions -contains "urn:user:registerdevice") -and
        ($PolicyCheck.GrantControls.BuiltInControls -contains "Mfa")
    }
    "CA-12" = {
        param($PolicyCheck)
        ($PolicyCheck.GrantControls.BuiltInControls -contains "Block") -and
        ($PolicyCheck.Conditions.Platforms.IncludePlatforms -contains "all") -and
        ($PolicyCheck.Conditions.Platforms.ExcludePlatforms.Count -gt 0)
    }
}


foreach ($policy in $CAPolicy) {
    foreach ($recommendation in $recommendations) {
        Update-PolicyStatus -Recommendation ([ref]$recommendation) -PolicyCheck $policy -StatusCheck $CheckFunctions[$recommendation.Control]
    }
}





#Set Row Order
$sort = "Name", "Status", "DateModified", "Users", "UsersInclude", "UsersExclude", "Cloud apps or actions", "ApplicationsIncluded", "ApplicationsExcluded", `
    "userActions", "AuthContext", "Conditions", "UserRisk", "SignInRisk", "PlatformsInclude", "PlatformsExclude", "ClientApps", "LocationsIncluded", `
    "LocationsExcluded", "Devices", "DevicesIncluded", "DevicesExcluded", "DeviceFilters", "AuthenticationFlows", "Grant Controls", "Block", "Require MFA", "Authentication Strength MFA", "CompliantDevice", `
    "DomainJoinedDevice", "CompliantApplication", "ApprovedApplication", "PasswordChange", "TermsOfUse", "CustomControls", "GrantOperator", `
    "Session Controls", "ApplicationEnforcedRestrictions", "CloudAppSecurity", "SignInFrequency", "PersistentBrowser", "ContinuousAccessEvaluation", "ResiliantDefaults", "secureSignInSession"

#Debug
#$pivot | Sort-Object $sort | Out-GridView           
function Display-RecommendationsAsHTMLFragment {
    param (
        [Parameter(Mandatory = $true)]
        [Recommendation[]]$Recommendations
    )

    $htmlFragment = @"
<div class='recommendations' id='ca-security-checks' style=''>
"@

    foreach ($rec in $Recommendations) {
        $links = ""
        foreach ($key in $rec.Links.Keys) {
            $links += "<div><a href='$($rec.Links[$key])'>$key</a></div>"
        }

        $excluded = $rec.Excluded 
        if ($($rec.Status)) { $RecStatus = "success"; $RecStatusNote = $($rec.PassText) }else { $RecStatus = "warning" ; $RecStatusNote = $($rec.FailRecommendation) }
        $htmlFragment += @"
    <div class='recommendation $RecStatus'>
        <div class='header'>
            <div class='title'>$($rec.Name)</div>
            <div class='control'>$($rec.Control)</div>
            
           
        </div>
        <div class='recommendation-description'>$($rec.Importance)</div>
        <div class='recommendation-links'><strong>Links:</strong>
            <div class='links'>
                $links
            </div>
        </div>    
        
          
        <div class='recommendation-comment'>$RecStatusNote</div>
        
        <div>$($rec.Note)</div>
        
    </div>
    <hr>
"@
    }

    $htmlFragment += @"
</div>
"@

    return $htmlFragment
}

if ($HTMLExport) {
    Write-host "Saving to File: HTML"
    $jquery = '  <script src="https://ajax.googleapis.com/ajax/libs/jquery/3.5.1/jquery.min.js"></script>
        <script>
        $(document).ready(function(){
            $("tr").click(function(){
                if (!$(this).hasClass("selected")) {
                    $(this).addClass("selected");
                } else {
                    $(this).removeClass("selected");
                }
            });
            $("th").click(function(){
                var colIndex = $(this).index();
                $("colgroup col").eq(colIndex).toggleClass("colselected");
            });
            function showView(view) {
                if (view === "table") {
                    $("#ca-export").show();
                    $("#ca-security-checks").hide();
                    $("#btn-table").addClass("active");
                    $("#btn-recs").removeClass("active");
                } else {
                    $("#ca-export").hide();
                    $("#ca-security-checks").show();
                    $("#btn-recs").addClass("active");
                    $("#btn-table").removeClass("active");
                }
                $("html, body").animate({ scrollTop: 0 }, 200);
            }
            window.showView = showView;
            $("#btn-recs").click(function() { showView("recs"); });
            $("#btn-table").click(function() { showView("table"); });
            $("#panel-toggle").click(function() {
                $("#side-panel").addClass("open");
                $("#panel-overlay").addClass("open");
            });
            function closePanel() {
                $("#side-panel").removeClass("open");
                $("#panel-overlay").removeClass("open");
            }
            $("#panel-close").click(closePanel);
            $("#panel-overlay").click(closePanel);
            $(document).keydown(function(e) { if (e.key === "Escape") closePanel(); });
        });
        </script>'

    $style = @"
*, *::before, *::after { box-sizing: border-box; }
body {
    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
    font-size: 15px;
    background: #f1f5f9;
    color: #1e293b;
    margin: 0;
    padding-top: 56px;
}

/* Navbar */
.navbar-custom {
    background: linear-gradient(135deg, #003f7a 0%, #005494 100%);
    position: fixed;
    top: 0; left: 0; right: 0;
    height: 56px;
    display: flex;
    align-items: center;
    justify-content: space-between;
    padding: 0 20px;
    box-shadow: 0 2px 10px rgba(0,0,0,0.3);
    z-index: 1030;
}
.nav-left { display: flex; align-items: center; gap: 12px; min-width: 220px; }
.nav-panel-btn {
    background: rgba(255,255,255,0.12);
    border: none;
    color: #fff;
    width: 34px; height: 34px;
    border-radius: 6px;
    cursor: pointer;
    font-size: 1em;
    display: flex;
    align-items: center;
    justify-content: center;
    transition: background 0.2s;
    flex-shrink: 0;
}
.nav-panel-btn:hover { background: rgba(255,255,255,0.22); }
.nav-brand-icon { color: rgba(255,255,255,0.6); font-size: 1.1em; }
.nav-title  { font-size: 0.95em; font-weight: 600; color: #fff; white-space: nowrap; }
.nav-subtitle { font-size: 0.7em; color: rgba(255,255,255,0.55); white-space: nowrap; }
.nav-center { display: flex; align-items: center; flex: 1; justify-content: center; }
.view-toggle {
    display: flex;
    background: rgba(0,0,0,0.22);
    border-radius: 8px;
    padding: 3px;
    gap: 3px;
}
.view-btn {
    background: transparent;
    border: none;
    color: rgba(255,255,255,0.65);
    padding: 6px 18px;
    border-radius: 6px;
    cursor: pointer;
    font-size: 0.82em;
    font-weight: 500;
    transition: all 0.2s;
    white-space: nowrap;
    display: flex;
    align-items: center;
    gap: 6px;
}
.view-btn.active { background: #fff; color: #005494; font-weight: 600; }
.view-btn:not(.active):hover { background: rgba(255,255,255,0.15); color: #fff; }
.nav-right { display: flex; flex-direction: column; align-items: flex-end; min-width: 220px; }
.nav-tenant { font-size: 0.88em; font-weight: 600; color: #fff; }
.nav-date   { font-size: 0.7em; color: rgba(255,255,255,0.55); }

/* Side Panel */
.side-panel {
    position: fixed;
    top: 0; left: -280px;
    width: 265px; height: 100%;
    background: #0f172a;
    z-index: 2100;
    transition: left 0.28s cubic-bezier(0.4,0,0.2,1);
    display: flex;
    flex-direction: column;
    box-shadow: 4px 0 24px rgba(0,0,0,0.45);
}
.side-panel.open { left: 0; }
.panel-header {
    display: flex;
    align-items: center;
    justify-content: space-between;
    padding: 0 16px;
    height: 56px;
    background: #005494;
    flex-shrink: 0;
}
.panel-header-title {
    color: #fff;
    font-weight: 600;
    font-size: 0.88em;
    letter-spacing: 0.04em;
    display: flex;
    align-items: center;
    gap: 8px;
}
.panel-close-btn {
    background: rgba(255,255,255,0.1);
    border: none;
    color: rgba(255,255,255,0.8);
    width: 28px; height: 28px;
    border-radius: 4px;
    cursor: pointer;
    display: flex;
    align-items: center;
    justify-content: center;
    font-size: 0.85em;
    transition: background 0.15s;
}
.panel-close-btn:hover { background: rgba(255,255,255,0.2); color: #fff; }
.panel-body { flex: 1; overflow-y: auto; padding: 8px 0; }
.panel-section { padding: 12px 16px 6px; }
.panel-section-title {
    font-size: 0.65em;
    font-weight: 700;
    color: rgba(255,255,255,0.3);
    letter-spacing: 0.12em;
    text-transform: uppercase;
    margin-bottom: 6px;
    padding-left: 8px;
}
.panel-link {
    display: flex;
    align-items: center;
    gap: 10px;
    padding: 8px 12px;
    color: rgba(255,255,255,0.7);
    text-decoration: none;
    border-radius: 6px;
    font-size: 0.83em;
    transition: all 0.15s;
    margin-bottom: 1px;
}
.panel-link:hover { background: rgba(255,255,255,0.08); color: #fff; text-decoration: none; }
.panel-link i { width: 14px; text-align: center; opacity: 0.5; font-size: 0.9em; }
.panel-link:hover i { opacity: 0.85; }
.panel-divider { height: 1px; background: rgba(255,255,255,0.06); margin: 6px 16px; }
.panel-footer {
    padding: 14px 18px;
    border-top: 1px solid rgba(255,255,255,0.06);
    font-size: 0.7em;
    color: rgba(255,255,255,0.25);
    flex-shrink: 0;
}
.panel-overlay {
    display: none;
    position: fixed;
    inset: 0;
    background: rgba(15,23,42,0.55);
    z-index: 2090;
    backdrop-filter: blur(2px);
}
.panel-overlay.open { display: block; }

/* Policy Matrix Table */
.policy-export { padding: 8px 24px 20px; box-sizing: border-box; }
table {
    border-collapse: collapse;
    margin-bottom: 20px;
    font-size: 0.87em;
    min-width: 400px;
    background: #fff;
    border-radius: 8px;
    overflow: hidden;
    box-shadow: 0 1px 4px rgba(0,0,0,0.07);
}
thead tr { background: linear-gradient(90deg, #005494, #0066b3); color: #fff; text-align: center; }
th, td { min-width: 150px; padding: 8px 10px; border: 1px solid #e2e8f0; vertical-align: top; text-align: center; }
tbody tr:nth-of-type(even) { background: #f8fafc; }
tbody tr:last-of-type { border-bottom: 2px solid #009879; }
tr:hover { background: #e0f2fe !important; }
.selected:not(th) { background: #dbeafe !important; }
th { background: #fff; }
.colselected { border: 3px solid #0ea5e9; }
table tr th:first-child, table tr td:first-child {
    position: sticky;
    inset-inline-start: 0;
    background: #005494;
    border: 0;
    color: #fff;
    font-weight: 600;
    text-align: center;
}
tbody tr:nth-of-type(even) td:first-child { background: #0066b3; }
tbody tr:nth-of-type(5),
tbody tr:nth-of-type(8),
tbody tr:nth-of-type(13),
tbody tr:nth-of-type(25),
tbody tr:nth-of-type(37) { background: #1e3a5f !important; color: #fff; }

/* Tooltip */
.tooltip-container { position: relative; display: inline-block; }
.tooltip-text {
    visibility: hidden;
    width: 200px;
    background: #1e293b;
    color: #fff;
    text-align: center;
    border-radius: 6px;
    padding: 6px 8px;
    position: absolute;
    z-index: 1;
    top: 115%; left: 50%;
    margin-left: -100px;
    opacity: 0;
    transition: opacity 0.2s;
    font-size: 0.8em;
    box-shadow: 0 4px 12px rgba(0,0,0,0.2);
}
.tooltip-container:hover .tooltip-text { visibility: visible; opacity: 1; }

/* Recommendations */
#ca-security-checks { padding: 12px 24px 24px; max-width: 1000px; margin: 0 auto; }
#ca-security-checks > h2 { font-size: 1.25em; color: #0f172a; font-weight: 700; margin: 0 0 4px; }
#ca-security-checks > p  { color: #64748b; margin: 0 0 20px; font-size: 0.85em; }
.recommendation {
    background: #fff;
    border: 1px solid #e2e8f0;
    border-left: 4px solid #cbd5e1;
    border-radius: 8px;
    padding: 14px 16px 12px;
    margin-bottom: 10px;
    box-shadow: 0 1px 3px rgba(0,0,0,0.04);
    position: relative;
    transition: box-shadow 0.2s, transform 0.1s;
}
.recommendation:hover { box-shadow: 0 4px 14px rgba(0,0,0,0.08); transform: translateY(-1px); }
.recommendation.success { border-left-color: #10b981; }
.recommendation.warning { border-left-color: #f59e0b; }
.recommendation div { margin-bottom: 3px; }
.header { display: flex; align-items: flex-start; gap: 10px; margin-bottom: 8px; }
.title  { font-size: 0.97em; font-weight: 600; color: #0f172a; flex: 1; }
.control {
    font-size: 0.72em;
    background: #f1f5f9;
    color: #64748b;
    padding: 2px 8px;
    border-radius: 20px;
    white-space: nowrap;
    font-weight: 600;
    border: 1px solid #e2e8f0;
}
.recommendation-description { font-size: 0.88em; color: #475569; line-height: 1.55; margin-bottom: 8px; }
.recommendation-links  { font-size: 0.84em; color: #64748b; margin-bottom: 6px; }
.recommendation-links strong { color: #334155; }
.links div { margin-top: 3px; margin-left: 0; }
.recommendation a       { color: #0369a1; text-decoration: none; }
.recommendation a:hover { text-decoration: underline; color: #0284c7; }
.recommendation-comment {
    font-size: 0.86em;
    color: #475569;
    padding: 7px 10px;
    background: #f8fafc;
    border-radius: 5px;
    border-left: 3px solid #cbd5e1;
    margin-top: 8px;
}
.recommendation.success .recommendation-comment { border-left-color: #10b981; background: #f0fdf4; color: #065f46; }
.recommendation.warning .recommendation-comment { border-left-color: #f59e0b; background: #fffbeb; color: #92400e; }
.recommendation strong { display: inline-block; }
.status-icon { display: inline-block; padding-left: 6px; }
.status-icon.success { color: #10b981; }
.status-icon.warning { color: #f59e0b; }
.status-icon.error   { color: #ef4444; }
.status-icon-large { position: absolute; top: 12px; right: 14px; font-size: 1.3em; }
.status-icon-large.success { color: #10b981; }
.status-icon-large.warning { color: #f59e0b; }
.status-icon-large.error   { color: #ef4444; }

/* Policy items */
.policy { margin-top: 10px; }
.policy-item { border: 1px solid #e2e8f0; border-left: 3px solid #cbd5e1; padding: 10px 12px; border-radius: 6px; margin-bottom: 8px; background: #fafafa; }
.policy-item.success { border-left-color: #10b981; background: #f0fdf4; }
.policy-item.warning { border-left-color: #f59e0b; background: #fffbeb; }
.policy-item.error   { border-left-color: #ef4444; background: #fef2f2; }
.policy-item strong { display: block; margin-bottom: 4px; font-size: 0.88em; }
.policy-header { position: relative; padding-right: 40px; }
.policy-content { display: flex; flex-direction: column; padding-left: 14px; margin-top: 6px; gap: 2px; }
.policy-include, .policy-exclude, .policy-grant { display: flex; align-items: flex-start; margin-top: 4px; font-size: 0.85em; }
.label-container { display: flex; align-items: center; margin-right: 8px; min-width: 40px; }
.include-label, .exclude-label, .grant-label {
    writing-mode: vertical-rl;
    transform: rotate(180deg);
    border-left: 2px solid #cbd5e1;
    color: #94a3b8;
    font-size: 0.72em;
    padding: 2px 3px;
}
.include-content, .exclude-content, .grant-content { margin-left: 6px; color: #334155; line-height: 1.5; }
.fa-external-link-alt { color: #94a3b8; font-size: 0.75em; margin-left: 4px; }
"@

    $html = "<html><head><base href='https://docs.microsoft.com/' target='_blank'>
    <meta charset='utf-8'>
        <meta name='viewport' content='width=device-width, initial-scale=1, shrink-to-fit=no'>
    
                  <link rel='stylesheet' href='https://cdnjs.cloudflare.com/ajax/libs/font-awesome/5.11.2/css/all.min.css' crossorigin='anonymous'>
                  <link rel='stylesheet' href='https://stackpath.bootstrapcdn.com/bootstrap/4.3.1/css/bootstrap.min.css' integrity='sha384-ggOyR0iXCbMQv3Xipma34MD+dH/1fQ784/j6cY/iJTQUOhcWr7x9JvoRxT2MZw1T' crossorigin='anonymous'>
                  <script src='https://cdnjs.cloudflare.com/ajax/libs/popper.js/1.14.7/umd/popper.min.js' integrity='sha384-UO2eT0CpHqdSJQ6hJty5KVphtPhzWj9WO1clHTMGa3JDZwrnQq4sF86dIHNDz0W1' crossorigin='anonymous'></script>
                  <script src='https://stackpath.bootstrapcdn.com/bootstrap/4.3.1/js/bootstrap.min.js' integrity='sha384-JjSmVgyd0p3pXB1rRibZUAYoIIy6OrQ6VrjIEaFf/nJGzIxFDsf4x0xIM+B07jRM' crossorigin='anonymous'></script>
                $jquery<style>
                $style
                </style>
                </head><body>
                <nav class='navbar-custom'>
                    <div class='nav-left'>
                        <button id='panel-toggle' class='nav-panel-btn' title='Open navigation'><i class='fas fa-bars'></i></button>
                        <i class='fas fa-shield-alt nav-brand-icon'></i>
                        <div>
                            <div class='nav-title'>CA Policy Report</div>
                            <div class='nav-subtitle'>Conditional Access Analysis</div>
                        </div>
                    </div>
                    <div class='nav-center'>
                        <div class='view-toggle'>
                            <button id='btn-recs' class='view-btn active'><i class='fas fa-clipboard-check'></i> Recommendations</button>
                            <button id='btn-table' class='view-btn'><i class='fas fa-table'></i> Policy Matrix</button>
                        </div>
                    </div>
                    <div class='nav-right'>
                        <span class='nav-tenant'><i class='fas fa-building' style='opacity:0.5;margin-right:5px;font-size:0.85em'></i>$Tenantname</span>
                        <span class='nav-date'>$Date</span>
                    </div>
                </nav>
                <div id='side-panel' class='side-panel'>
                    <div class='panel-header'>
                        <div class='panel-header-title'><i class='fas fa-shield-alt'></i> CA Export</div>
                        <button id='panel-close' class='panel-close-btn' title='Close panel'><i class='fas fa-times'></i></button>
                    </div>
                    <div class='panel-body'>
                        <div class='panel-section'>
                            <div class='panel-section-title'>Azure Administration</div>
                            <a href='https://portal.azure.com' class='panel-link' target='_blank'><i class='fas fa-cloud'></i> Azure Portal</a>
                            <a href='https://entra.microsoft.com' class='panel-link' target='_blank'><i class='fas fa-id-card'></i> Entra Admin Center</a>
                            <a href='https://entra.microsoft.com/#view/Microsoft_AAD_ConditionalAccess/ConditionalAccessBlade' class='panel-link' target='_blank'><i class='fas fa-lock'></i> Conditional Access Policies</a>
                            <a href='https://security.microsoft.com' class='panel-link' target='_blank'><i class='fas fa-shield-alt'></i> Microsoft Defender</a>
                        </div>
                        <div class='panel-divider'></div>
                        <div class='panel-section'>
                            <div class='panel-section-title'>References</div>
                            <a href='https://learn.microsoft.com/en-us/entra/identity/conditional-access/overview' class='panel-link' target='_blank'><i class='fas fa-book'></i> CA Documentation</a>
                            <a href='https://aka.ms/CATemplates' class='panel-link' target='_blank'><i class='fas fa-layer-group'></i> CA Policy Templates</a>
                            <a href='https://learn.microsoft.com/en-us/entra/identity/conditional-access/plan-conditional-access' class='panel-link' target='_blank'><i class='fas fa-sitemap'></i> CA Planning Guide</a>
                        </div>
                        <div class='panel-divider'></div>
                        <div class='panel-section'>
                            <div class='panel-section-title'>My Links</div>
                            <a href='https://dougsbaker.com' class='panel-link' target='_blank'><i class='fas fa-globe'></i> My Website</a>
                            <a href='https://github.com/dougsbaker' class='panel-link' target='_blank'><i class='fab fa-github'></i> GitHub</a>
                        </div>
                    </div>
                    <div class='panel-footer'>CA Policy Export &bull; DougSBaker &bull; 2026</div>
                </div>
                <div id='panel-overlay' class='panel-overlay'></div>"
       
    $SecurityCheck = Display-RecommendationsAsHTMLFragment -Recommendations $recommendations

    Write-host "Launching: Web Browser"           
    $Launch = $ExportLocation + $FileName
    $HTML += "<div class='policy-export' id='ca-export' style='display: none;'>"
    $HTML += $pivot  | Where-Object { $_."CA Item" -ne 'row1' } | Sort-object { $sort.IndexOf($_."CA Item") } | convertto-html -Fragment
    $html += "</div>"
    $html += $SecurityCheck
    Add-Type -AssemblyName System.Web
    [System.Web.HttpUtility]::HtmlDecode($HTML) | Out-File $Launch
   
    start-process $Launch
}
if ($JsonExport) {
    Write-host "Saving to File: JSON" 
    $LaunchJson = $ExportLocation + $JsonFileName
    $CAPolicy | ConvertTo-Json -Depth 8 | Out-File $LaunchJson
    start-process $LaunchJson
}