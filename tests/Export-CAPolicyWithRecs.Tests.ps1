# Pester tests for core helper logic in Export-CAPolicyWithRecs.ps1
# Run with: Invoke-Pester -Path tests -Output Detailed

<# Each Describe defines needed helper functions in its own BeforeAll to work with Pester V5 discovery (which does not execute top-level code). #>

Describe 'Get-NormalizedPolicyHash' {
    BeforeAll {
        function Get-NormalizedPolicyHash { param([Parameter(Mandatory)]$Policy)
            $norm = $Policy | Select-Object * -ExcludeProperty PolicyId, DateModified, CreatedDateTime, Description, 'Duplicate Matches', IsDuplicate, RawJson, ContentHash
            $json = ($norm | ConvertTo-Json -Depth 8 -Compress)
            $bytes = [Text.Encoding]::UTF8.GetBytes($json)
            $sha = [System.Security.Cryptography.SHA256]::Create()
            return ([BitConverter]::ToString($sha.ComputeHash($bytes))).Replace('-', '')
        }
    }
    It 'Produces identical hash when only description changes' {
        $base = [pscustomobject]@{ Name='TestPolicy'; Description='A'; Status='enabled'; Created=(Get-Date); Modified=(Get-Date); 'Duplicate Matches'=''; IsDuplicate=$false; RawJson=''; ContentHash='' }
        $variant = $base | Select-Object *
        $variant.Description = 'Different description text'
        $h1 = Get-NormalizedPolicyHash -Policy $base
        $h2 = Get-NormalizedPolicyHash -Policy $variant
        $h1 | Should -Be $h2
    }
}

Describe 'Test-IsPhishResistantStrength' {
    BeforeAll {
        function Test-IsPhishResistantStrength { param($Policy)
            $strengthObj = $Policy.GrantControls.AuthenticationStrength
            if (-not $strengthObj -or -not $strengthObj.AllowedCombinations) { return $false }
            $nonPhish = @('deviceBasedPush','temporaryAccessPassOneTime','temporaryAccessPassMultiUse','microsoftAuthenticatorPush','sms','voice','softwareOath','hardwareOath','x509CertificateSingleFactor','federatedSingleFactor','qrCodePin')
            $contains = ($strengthObj.AllowedCombinations | Where-Object { $nonPhish -contains $_ }).Count -gt 0
            return (-not $contains)
        }
        # Prepare fake auth strength cache
        $script:AuthStrengthCache = @{}
        $script:AuthStrengthCachePopulated = $true
        $script:AuthStrengthCache['strength-phish'] = [pscustomobject]@{ Id='strength-phish'; DisplayName='PhishResistant'; AllowedCombinations=@('fido2','windowsHelloForBusiness') }
        $script:AuthStrengthCache['strength-nonphish'] = [pscustomobject]@{ Id='strength-nonphish'; DisplayName='NonPhish'; AllowedCombinations=@('microsoftAuthenticatorPush','fido2') }
    }
    It 'Returns True for strength with only phish-resistant combos' {
        $policy = [pscustomobject]@{ GrantControls = [pscustomobject]@{ AuthenticationStrength = [pscustomobject]@{ Id='strength-phish'; DisplayName='PhishResistant'; AllowedCombinations=@('fido2','windowsHelloForBusiness') } }; Conditions = [pscustomobject]@{ Users = [pscustomobject]@{ IncludeRoles=@() } } }
        (Test-IsPhishResistantStrength -Policy $policy) | Should -BeTrue
    }
    It 'Returns False when any non-phish combo present' {
        $policy = [pscustomobject]@{ GrantControls = [pscustomobject]@{ AuthenticationStrength = [pscustomobject]@{ Id='strength-nonphish'; DisplayName='NonPhish'; AllowedCombinations=@('microsoftAuthenticatorPush','fido2') } }; Conditions = [pscustomobject]@{ Users = [pscustomobject]@{ IncludeRoles=@() } } }
        (Test-IsPhishResistantStrength -Policy $policy) | Should -BeFalse
    }
}

Describe 'CA Check Functions Representative Tests' {
    Context 'CA-04 Require MFA for Admins' {
        BeforeAll {
            function Test-PolicyTargetsAdminRoles { param($Policy) if (-not $Policy -or -not $Policy.Conditions.Users.IncludeRoles) { return $false }; return ($Policy.Conditions.Users.IncludeRoles.Count -gt 0) }
            function Test-PolicyRequiresMfaForAdmins { param($Policy)
                if (-not (Test-PolicyTargetsAdminRoles -Policy $Policy)) { return $false }
                $strengthName = ''
                if ($Policy.GrantControls.AuthenticationStrength.DisplayName) { $strengthName = [string]$Policy.GrantControls.AuthenticationStrength.DisplayName }
                $l = $strengthName.ToLowerInvariant()
                $hasImpliedMfa = ($Policy.GrantControls.BuiltInControls -contains 'Mfa') -or ($l -match 'phishing') -or ($l -match 'passwordless') -or ($l -match 'multifactor') -or ($l -match '\bmfa\b')
                return $hasImpliedMfa
            }
        }
        It 'Returns False when policy does not target admin roles' {
            $policy = [pscustomobject]@{
                GrantControls = [pscustomobject]@{ BuiltInControls = @('Mfa'); AuthenticationStrength = [pscustomobject]@{ DisplayName = 'MFA' } }
                Conditions = [pscustomobject]@{ Users = [pscustomobject]@{ IncludeRoles = @() } }
            }
            (Test-PolicyRequiresMfaForAdmins -Policy $policy) | Should -BeFalse
        }
        It 'Returns True when policy targets admin roles and implies MFA' {
            $policy = [pscustomobject]@{
                GrantControls = [pscustomobject]@{ BuiltInControls = @('Mfa'); AuthenticationStrength = [pscustomobject]@{ DisplayName = 'MultiFactor' } }
                Conditions = [pscustomobject]@{ Users = [pscustomobject]@{ IncludeRoles = @('62e90394-69f5-4237-9190-012177145e10') } } # Global Admin role id example
            }
            (Test-PolicyRequiresMfaForAdmins -Policy $policy) | Should -BeTrue
        }
    }

    Context 'CA-05 Require Phish-Resistant MFA for Admins' {
        BeforeAll {
            function Test-IsPhishResistantStrength { param($Policy)
                $strengthObj = $Policy.GrantControls.AuthenticationStrength
                if (-not $strengthObj -or -not $strengthObj.AllowedCombinations) { return $false }
                $nonPhish = @('deviceBasedPush','temporaryAccessPassOneTime','temporaryAccessPassMultiUse','microsoftAuthenticatorPush','sms','voice','softwareOath','hardwareOath','x509CertificateSingleFactor','federatedSingleFactor','qrCodePin')
                $contains = ($strengthObj.AllowedCombinations | Where-Object { $nonPhish -contains $_ }).Count -gt 0
                return (-not $contains)
            }
            function Test-PolicyTargetsAdminRoles { param($Policy) if (-not $Policy -or -not $Policy.Conditions.Users.IncludeRoles) { return $false }; return ($Policy.Conditions.Users.IncludeRoles.Count -gt 0) }
            function Test-PolicyRequiresPhishResistantMfaForAdmins { param($Policy) if (-not (Test-PolicyTargetsAdminRoles -Policy $Policy)) { return $false }; return (Test-IsPhishResistantStrength -Policy $Policy) }
            # Ensure cache reflects phish-resistant combos only
            $script:AuthStrengthCache = @{ }
            $script:AuthStrengthCachePopulated = $true
            $script:AuthStrengthCache['strength-pr'] = [pscustomobject]@{ Id='strength-pr'; DisplayName='PhishResistant'; AllowedCombinations=@('fido2') }
        }
        It 'Returns False when not targeting admin roles' {
            $policy = [pscustomobject]@{
                GrantControls = [pscustomobject]@{ AuthenticationStrength = [pscustomobject]@{ Id='strength-pr'; DisplayName='PhishResistant'; AllowedCombinations=@('fido2') } }
                Conditions = [pscustomobject]@{ Users = [pscustomobject]@{ IncludeRoles = @() } }
            }
            (Test-PolicyRequiresPhishResistantMfaForAdmins -Policy $policy) | Should -BeFalse
        }
        It 'Returns True when targeting admin roles AND strength is phish-resistant' {
            $policy = [pscustomobject]@{
                GrantControls = [pscustomobject]@{ AuthenticationStrength = [pscustomobject]@{ Id='strength-pr'; DisplayName='PhishResistant'; AllowedCombinations=@('fido2') } }
                Conditions = [pscustomobject]@{ Users = [pscustomobject]@{ IncludeRoles = @('62e90394-69f5-4237-9190-012177145e10') } }
            }
            (Test-PolicyRequiresPhishResistantMfaForAdmins -Policy $policy) | Should -BeTrue
        }
    }

    Context 'CA-08 Direct User Assignment' {
        BeforeAll {
            function Test-CA08 { param($PolicyCheck)
                $PolicyCheck.Conditions.Users.IncludeUsers -ne 'None' -and
                $null -ne $PolicyCheck.Conditions.Users.IncludeUsers -and
                $PolicyCheck.Conditions.Users.IncludeUsers -ne 'All' -and
                $PolicyCheck.Conditions.Users.IncludeUsers -ne 'GuestsOrExternalUsers'
            }
        }
        It 'Returns False when IncludeUsers is All' {
            $policy = [pscustomobject]@{ Conditions = [pscustomobject]@{ Users = [pscustomobject]@{ IncludeUsers = 'All' } } }
            (Test-CA08 -PolicyCheck $policy) | Should -BeFalse
        }
        It 'Returns True when IncludeUsers is a direct user id' {
            $policy = [pscustomobject]@{ Conditions = [pscustomobject]@{ Users = [pscustomobject]@{ IncludeUsers = '12345678-aaaa-bbbb-cccc-1234567890ab' } } }
            (Test-CA08 -PolicyCheck $policy) | Should -BeTrue
        }
        It 'Returns False when IncludeUsers is GuestsOrExternalUsers' {
            $policy = [pscustomobject]@{ Conditions = [pscustomobject]@{ Users = [pscustomobject]@{ IncludeUsers = 'GuestsOrExternalUsers' } } }
            (Test-CA08 -PolicyCheck $policy) | Should -BeFalse
        }
    }
}
