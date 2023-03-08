#Conditional Access Export Utility

<#
	.SYNOPSIS
		Conditional Access Export Utility
	.DESCRIPTION
       Exports CA Policy to HTML Format for auditing/historical purposes. 

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
param (
    [Parameter()]
    [String]$TenantID,
    [Parameter()]
    [String]$PolicyID
)

$ExportLocation = $PSScriptRoot
if (!$ExportLocation) {$ExportLocation = $PWD }
$FileName = "\CAPolicy.html"
$HTMLExport = $true

try {
    Get-MgIdentityConditionalAccessPolicy -ErrorAction Stop > $null
    Write-host "Connected: MgGraph"
}
catch {
    Write-host "Connecting: MgGraph"  
    Try {
        #Connect-AzureAD
        Select-MgProfile -Name "beta"
        Connect-MgGraph -Scopes 'Policy.Read.All', 'Directory.Read.All', 'Application.Read.All', 'Agreement.Read.All'
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
    $CAPolicy = Get-MgIdentityConditionalAccessPolicy -all

}

$TenantData = Get-MgOrganization
$TenantName = $TenantData.DisplayName
$date = Get-Date

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


    $ExcludeUG = $null
    $ExcludeUG = $Policy.Conditions.Users.ExcludeUsers
    $ExcludeUG += $Policy.Conditions.Users.ExcludeGroups
    $ExcludeUG += $Policy.Conditions.Users.ExcludeRoles
    
    
    $Apps += $Policy.Conditions.Applications.IncludeApplications
    $Apps += $Policy.Conditions.Applications.ExcludeApplications

    
    $AdUsers += $ExcludeUG
    $AdUsers += $IncludeUG
    
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
        'Grant Controls'               = "";
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
        'Session Controls'               = "";
        ApplicationEnforcedRestrictions = $Policy.SessionControls.ApplicationEnforcedRestrictions.IsEnabled
        CloudAppSecurity                = $Policy.SessionControls.CloudAppSecurity.IsEnabled
        SignInFrequency                 = "$($Policy.SessionControls.SignInFrequency.Value) $($conditionalAccessPolicy.SessionControls.SignInFrequency.Type)"
        PersistentBrowser               = $Policy.SessionControls.PersistentBrowser.Mode
        ContinuousAccessEvaluation      = $Policy.SessionControls.ContinuousAccessEvaluation.Mode
        ResiliantDefaults               = $policy.SessionControls.DisableResilienceDefaults
    }
  
    
}

#Swith user/group Guid to display names
Write-host "Converting: AzureAD Guid"
#Filter out Objects
$ADsearch = $AdUsers | Where-Object { $_ -ne 'All' -and $_ -ne 'GuestsOrExternalUsers' -and $_ -ne 'None' }
$cajson = $CAExport | ConvertTo-Json -Depth 4
$AdNames = @{}
Get-MgDirectoryObjectById -ids $ADsearch | ForEach-Object { 
    $obj = $_.Id
    $disp = $_.AdditionalProperties.displayName
    $AdNames.$obj = $disp
    $cajson = $cajson -replace "$obj", "$disp"
}
$CAExport = $cajson | ConvertFrom-Json
#Switch Apps Guid with Display names
$allApps = Get-MgServicePrincipal -All
$allApps | Where-Object { $_.AppId -in $Apps } | ForEach-Object {
    $obj = $_.AppId
    $disp = $_.DisplayName
    $cajson = $cajson -replace "$obj", "$disp"
}
#switch named location Guid for Display Names
Get-MgIdentityConditionalAccessNamedLocation | ForEach-Object {
    $obj = $_.Id
    $disp = $_.DisplayName
    $cajson = $cajson -replace "$obj", "$disp"
}
#Switch TOU Id for display name
Get-MgAgreement | ForEach-Object {
    $obj = $_.Id
    $disp = $_.DisplayName
    $cajson = $cajson -replace "$obj", "$disp"
}

#Switch Roles Guid to Names
#Get-MgDirectoryRole | ForEach-Object{
Get-MgDirectoryRoleTemplate | ForEach-Object {
    $obj = $_.Id
    $disp = $_.DisplayName
    $cajson = $cajson -replace "$obj", "$disp"
}
$CAExport = $cajson | ConvertFrom-Json

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



#Set Row Order
$sort = "Name", "Status", "Users", "UsersInclude", "UsersExclude", "Cloud apps or actions", "ApplicationsIncluded", "ApplicationsExcluded", `
    "userActions", "AuthContext", "Conditions", "UserRisk", "SignInRisk", "PlatformsInclude", "PlatformsExclude", "ClientApps", "LocationsIncluded", `
    "LocationsExcluded", "Devices", "DevicesIncluded", "DevicesExcluded", "DeviceFilters", "Grant Controls", "Block", "Require MFA", "Authentication Strength MFA", "CompliantDevice", `
    "DomainJoinedDevice", "CompliantApplication", "ApprovedApplication", "PasswordChange", "TermsOfUse", "CustomControls", "GrantOperator", `
    "Session Controls",  "ApplicationEnforcedRestrictions", "CloudAppSecurity", "SignInFrequency", "PersistentBrowser",  "ContinuousAccessEvaluation", "ResiliantDefaults"

#Debug
#$pivot | Sort-Object $sort | Out-GridView           


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
    });
    </script>'
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
                .title {
                    font-size: 1.5em;
                    font-weight: bold;
                    font-family: Arial, sans-serif;
                    top: 0;
                    right: 0;
                    left: 0;
                }
                
                table {
                    border-collapse: collapse;
                    margin-bottom: 30px;
                    margin-top: 55px;
                    font-size: 0.9em;
                    font-family: Arial, sans-serif;
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
                  tbody tr:nth-of-type(4),
                  tbody tr:nth-of-type(7),
                  tbody tr:nth-of-type(12),
                  tbody tr:nth-of-type(23),
                  tbody tr:nth-of-type(35) {
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
                </style></head><body> <nav class='navbar  fixed-top navbar-custom p-3 border-bottom'>
                <div class='container-fluid'>
                    <div class='col-sm' style='text-align:left'>
                        <div class='row'><div><i class='fa fa-server' aria-hidden='true'></i></div><div class='ml-3'><strong>CA Export</strong></div></div>
                    </div>
                    <div class='col-sm' style='text-align:center'>
                        <strong>$Tenantname</strong>
                    </div>
                    <div class='col-sm' style='text-align:right'>
                    <strong>$Date</strong>
                    </div>
                </div>
            </nav> "
                

    Write-host "Launching: Web Browser"           
    $Launch = $ExportLocation + $FileName
    $HTML += $pivot  | Where-Object { $_."CA Item" -ne 'row1' } | Sort-object { $sort.IndexOf($_."CA Item") } | convertto-html -Fragment
    $HTML | Out-File $Launch
    start-process $Launch
}
