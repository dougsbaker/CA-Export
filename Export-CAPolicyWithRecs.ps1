
[CmdletBinding()]
param (
    [Parameter()]
    [String]$TenantID,
    [Parameter()]
    [String]$PolicyID
)

$ExportLocation = $PSScriptRoot
if (!$ExportLocation) { $ExportLocation = $PWD }
$FileName = "\CAPolicyV2.html"
$JsonFileName = "\CAPolicyV2.json"
$HTMLExport = $true




try {
    Get-MgIdentityConditionalAccessPolicy -ErrorAction Stop > $null
    Write-host "Connected: MgGraph"
}
catch {
    Write-host "Connecting: MgGraph"  
    Try {
        #Connect-AzureAD
        #Select-MgProfile -Name "beta"
        Connect-MgGraph -Scopes 'Policy.Read.All', 'Directory.Read.All', 'Application.Read.All', 'Agreement.Read.All' -nowelcome
    }
    Catch {
        Write-host "Error: Please Install MgGraph Module" -ForegroundColor Yellow
        Write-Host "Run: Install-Module Microsoft.Graph" -ForegroundColor Yellow
        Exit
    }
}


$TenantData = Get-MgOrganization
$TenantName = $TenantData.DisplayName
$date = Get-Date
Write-Host "Connected: $TenantName tenant"


#Collect CA Policy
Write-host "Exporting: CA Policy"
if ($PolicyID) {
    $CAPolicy = Get-MgIdentityConditionalAccessPolicy -ConditionalAccessPolicyId $PolicyID
}
else {
    $CAPolicy = Get-MgIdentityConditionalAccessPolicy -all -ExpandProperty *

}

$TenantData = Get-MgOrganization
$TenantName = $TenantData.DisplayName
$date = Get-Date


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
$mgobjects = Get-MgDirectoryObjectById -ids $ADsearch 
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
$MGApps = Get-MgServicePrincipal -All 
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
$mgLoc = Get-MgIdentityConditionalAccessNamedLocation
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
$mgTou = Get-MgAgreement 
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
$mgRole = Get-MgDirectoryRoleTemplate 
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

$AdUsers = @()
$Apps = @()
#Extract Values
Write-host "Extracting: CA Policy Data"
foreach ( $Policy in $CAPolicy) {

    $IncludeUG = $null
    $IncludeUG = $Policy.Conditions.Users.IncludeUsers
    $IncludeUG += $Policy.Conditions.Users.IncludeGroups
    $IncludeUG += $Policy.Conditions.Users.IncludeRoles
    $IncludeUG += $policy.conditions.users.IncludeGuestsOrExternalUsers.GuestOrExternalUserTypes -replace ',', ', '
    $DateCreated = $null
    $DateCreated = $policy.CreatedDateTime
    $DateModified = $null
    $DateModified = $Policy.ModifiedDateTime
    

    $ExcludeUG = $null
    $ExcludeUG = $Policy.Conditions.Users.ExcludeUsers
    $ExcludeUG += $Policy.Conditions.Users.ExcludeGroups
    $ExcludeUG += $Policy.Conditions.Users.ExcludeRoles
    $ExcludeUG += $policy.conditions.users.ExcludeGuestsOrExternalUsers.GuestOrExternalUserTypes -replace ',', ', '
    
    
    $Apps += $Policy.Conditions.Applications.IncludeApplications
    $Apps += $Policy.Conditions.Applications.ExcludeApplications

    
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
        AuthenticationFlows             = $Policy.Conditions.AdditionalProperties.authenticationFlows.Values;
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
        SignInFrequency                 = "$($Policy.SessionControls.SignInFrequency.Value) $($conditionalAccessPolicy.SessionControls.SignInFrequency.Type)"
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
    
    $CheckExcUG = $PolicyCheck.Conditions.Users.ExcludeUsers + $PolicyCheck.Conditions.Users.ExcludeGroups + $PolicyCheck.Conditions.Users.ExcludeRoles
    $CheckIncUG = $PolicyCheck.Conditions.Users.IncludeUsers + $PolicyCheck.Conditions.Users.IncludeGroups + $PolicyCheck.Conditions.Users.IncludeRoles
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
    if ($PolicyCheck.SessionControls.SignInFrequency.Value -and $conditionalAccessPolicy.SessionControls.SignInFrequency.Type) {
        $checkSession += "    SignInFrequency: $($PolicyCheck.SessionControls.SignInFrequency.Value) $($conditionalAccessPolicy.SessionControls.SignInFrequency.Type)`n"
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
                <strong>$($PolicyCheck.DisplayName)</strong>
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
                        <b>Applications:</b> $($PolicyCheck.Conditions.Applications.IncludeApplications -join ', ')
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
        ($PolicyCheck.Conditions.Users.IncludeRoles -eq $null -or $PolicyCheck.Conditions.Users.IncludeRoles.Count -eq 0)
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
<div class='recommendations' id='ca-security-checks' style='display: none;'>
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
            if(!$(this).hasClass("selected")){
                $(this).addClass("selected");
            } else {
                $(this).removeClass("selected");
            }
        });
        
        $("th").click(function(){
            // Get the index of the clicked column
            var colIndex = $(this).index();
            // Select the corresponding col element and add or remove the class
            $("colgroup col").eq(colIndex).toggleClass("colselected");
        });
        $(document).ready(function() {
        $("#toggle-icon").click(function() {
            $("table").toggle();
            $("#ca-security-checks").toggle();
        });
        }); 
    });
    </script>'
    $style = @"
    /* General Styles */
                    html, body {
                        font-family: Arial, sans-serif;
                    }

                    .title {
                        font-size: 1.5em;
                        font-weight: bold;
                        top: 0;
                        right: 0;
                        left: 0;
                    }

                    .navbar-custom { 
                        background-color: #005494;
                        color: white; 
                        padding-bottom: 10px;
                    }

                    .navbar-custom .navbar-brand, 
                    .navbar-custom .navbar-text { 
                        color: white; 
                        padding-top: 70px;
                        padding-bottom: 10px;
                    }

                    .sr-only {
                        border: 0;
                        clip: rect(0, 0, 0, 0);
                        height: 1px;
                        margin: -1px;
                        overflow: hidden;
                        padding: 0;
                        position: absolute;
                        width: 1px;
                    }

                    .sr-only-focusable:active, 
                    .sr-only-focusable:focus {
                        clip: auto;
                        height: auto;
                        margin: 0;
                        overflow: visible;
                        position: static;
                        width: auto;
                    }

                    /* Export Policies Styles */
                    table {
                        border-collapse: collapse;
                        margin-bottom: 30px;
                        margin-top: 55px;
                        font-size: 0.9em;
                        min-width: 400px;
                    }

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

                    tbody tr:nth-of-type(even) {
                        background-color: #f3f3f3;
                    }

                    tbody tr:last-of-type {
                        border-bottom: 2px solid #009879;
                    }

                    tr:hover {
                        background-color: #d8d8d8 !important;
                    }

                    .selected:not(th) {
                        background-color: #eaf7ff !important;
                    }

                    th {
                        background-color: white;
                    }

                    .colselected {
                        width: 10%;
                        border: 5px solid #59c7fb;
                    }

                    table tr th:first-child, table tr td:first-child {
                        position: sticky;
                        inset-inline-start: 0; 
                        background-color: #005494;
                        border: 0px;
                        color: #fff;
                        font-weight: bolder;
                        text-align: center;
                    }

                    tbody tr:nth-of-type(even) td:first-child {
                        background-color: #547c9b;
                    }

                    tbody tr:nth-of-type(5),
                    tbody tr:nth-of-type(8),
                    tbody tr:nth-of-type(13),
                    tbody tr:nth-of-type(25),
                    tbody tr:nth-of-type(37) {
                        background-color: #005494 !important;
                    }
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
                    /* Recommendations Styles */
                    #ca-security-checks {
                        padding: 20px;
                        background-color: #f8f9fa; 
                        border: 1px solid #ddd; 
                        border-radius: 5px;
                        margin-top: 55px;
                    }

                    #ca-security-checks h2 {
                        margin-top: 0;
                        color: #343a40;
                    }

                    #ca-security-checks p {
                        color: #6c757d;
                    }
                    .header {
                        display: flex;
                        align-items: center;
                        }
                    .recommendations {
                        font-family: Arial, sans-serif;
                    }
                    .recommendation.success {
                        border-left-color: green;
                        border-left-width: 7px;
                    }
                    .recommendation.warning {
                        border-left-color: orange;
                        border-left-width: 7px;
                    }
                    .recommendation {
                        padding: 10px;
                        margin-bottom: 10px;
                        border: 1px solid #ddd;
                        border-radius: 5px;
                        background-color: #f9f9f9;
                    }

                    .recommendation div {
                        margin-bottom: 5px;
                    }
                    .control {
                        color: gray;
                    }

                    .links div {
                        margin-left: 20px;
                    }

                    .recommendation a {
                        color: #0073e6;
                        text-decoration: none;
                    }

                    .recommendation a:hover {
                        text-decoration: underline;
                    }

                    .recommendation strong {
                        display: inline-block;
                    }

                    .status-icon {
                        display: inline-block;
                        padding-left: 10px;
                    }

                    .status-icon.success {
                        color: green;
                    }

                    .status-icon.warning {
                        color: orange;
                    }

                    .status-icon.error {
                        color: red;
                    }

                    .policy {
    margin-top: 10px;
}

.policy-item {
    border: 2px solid;
    padding: 10px;
    border-radius: 5px;
    margin-bottom: 10px;
}

.policy-item.success {
    border-color: green;
    background-color: #e6ffe6;
}

.policy-item.warning {
    border-color: orange;
    background-color: #fff8e6;
}

.policy-item.error {
    border-color: red;
    background-color: #ffe6e6;
}

.policy-item strong {
    display: block;
    margin-bottom: 5px;
}

.policy-content {
    display: flex;
    flex-direction: column;
    padding-left: 20px;
    margin-top: 5px;
}

.policy-include, .policy-exclude, .policy-grant {
    display: flex;
    align-items: flex-start;
    margin-top: 5px;
}

.label-container {
    display: flex;
    align-items: center;
    margin-right: 10px; /* Space between the label and content */
}

.include-label, .exclude-label, .grant-label {
    writing-mode: vertical-rl;
    transform: rotate(180deg); /* Optional: Rotate the text to read from bottom to top */
    border-left: 3px solid darkgrey;
    color: darkgray;
}

.include-content, .exclude-content, .grant-content {
    margin-left: 10px; /* Space between the label and content */
}

.policy-header {
    position: relative;
    padding-right: 40px;
}


                    .status-icon-large {
                        position: absolute;
                        top: 0;
                        right: 0;
                        font-size: 2em;
                    }

                    .status-icon-large.success {
                        color: green;
                    }

                    .status-icon-large.warning {
                        color: orange;
                    }

                    .status-icon-large.error {
                        color: red;
                    }


"@

    $html = "<html><head><base href='https://docs.microsoft.com/' target='_blank'>
    <meta charset='utf-8'>
        <meta name='viewport' content='width=device-width, initial-scale=1, shrink-to-fit=no'>
    
                  <link rel='stylesheet' href='https://cdnjs.cloudflare.com/ajax/libs/font-awesome/5.11.2/css/all.min.css' crossorigin='anonymous'>
                  <link rel='stylesheet' href='https://stackpath.bootstrapcdn.com/bootstrap/4.3.1/css/bootstrap.min.css' integrity='sha384-ggOyR0iXCbMQv3Xipma34MD+dH/1fQ784/j6cY/iJTQUOhcWr7x9JvoRxT2MZw1T' crossorigin='anonymous'>
                  <script src='https://code.jquery.com/jquery-3.3.1.slim.min.js' integrity='sha384-q8i/X+965DzO0rT7abK41JStQIAqVgRVzpbzo5smXKp4YfRvH+8abtTE1Pi6jizo' crossorigin='anonymous'></script>
                  <script src='https://cdnjs.cloudflare.com/ajax/libs/popper.js/1.14.7/umd/popper.min.js' integrity='sha384-UO2eT0CpHqdSJQ6hJty5KVphtPhzWj9WO1clHTMGa3JDZwrnQq4sF86dIHNDz0W1' crossorigin='anonymous'></script>
                  <script src='https://stackpath.bootstrapcdn.com/bootstrap/4.3.1/js/bootstrap.min.js' integrity='sha384-JjSmVgyd0p3pXB1rRibZUAYoIIy6OrQ6VrjIEaFf/nJGzIxFDsf4x0xIM+B07jRM' crossorigin='anonymous'></script>
                  <script src='https://cdnjs.cloudflare.com/ajax/libs/font-awesome/5.11.2/js/all.js'></script>
                $jquery<style>
                $style
                </style>
                </head><body> <nav class='navbar  fixed-top navbar-custom p-3 border-bottom'>
                <div class='container-fluid'>
                    <div class='col-sm' style='text-align:left'>
                        <div class='row'><div><i class='fa fa-server' aria-hidden='true'></i></div><div class='ml-3'><strong>CA Export</strong></div><div class='ml-3' id='toggle-icon' style='cursor: pointer;'><i class='fas fa-exchange-alt'></i></div>
</div>
                    </div>
                    <div class='col-sm' style='text-align:center'>
                        <strong>$Tenantname</strong>
                    </div>
                    <div class='col-sm' style='text-align:right'>
                    <strong>$Date</strong>
                    </div>
                </div>
            </nav> "
       
    $SecurityCheck = Display-RecommendationsAsHTMLFragment -Recommendations $recommendations

    Write-host "Launching: Web Browser"           
    $Launch = $ExportLocation + $FileName
    $LaunchJson = $ExportLocation + $JsonFileName
    $HTML += $pivot  | Where-Object { $_."CA Item" -ne 'row1' } | Sort-object { $sort.IndexOf($_."CA Item") } | convertto-html -Fragment
    $html += $SecurityCheck
    Add-Type -AssemblyName System.Web
    [System.Web.HttpUtility]::HtmlDecode($HTML) | Out-File $Launch
    $CAPolicy | ConvertTo-Json -Depth 8 | Out-File $LaunchJson
    $LaunchJson
    start-process $Launch
}