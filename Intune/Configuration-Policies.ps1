
#Requires -Version 7.0

<#
.SYNOPSIS
    Creates Intune configuration policies from Settings Catalog export
.DESCRIPTION
    Creates device configuration policies with dynamic tenant value substitution
    and assigns them to appropriate device groups
.AUTHOR
    CB & Claude Partnership
.VERSION
    1.0
#>

# Required Modules
$RequiredModules = @(
    'Microsoft.Graph.Authentication',
    'Microsoft.Graph.DeviceManagement',
    'Microsoft.Graph.Groups',
    'Microsoft.Graph.Identity.DirectoryManagement'
)

# Auto-install and import required modules
function Initialize-Modules {
    Write-Host "🔧 Checking required modules..." -ForegroundColor Yellow
    
    foreach ($Module in $RequiredModules) {
        if (!(Get-Module -ListAvailable -Name $Module)) {
            Write-Host "Installing $Module..." -ForegroundColor Yellow
            Install-Module $Module -Force -Scope CurrentUser -AllowClobber
        }
        if (!(Get-Module -Name $Module)) {
            Write-Host "Importing $Module..." -ForegroundColor Yellow
            Import-Module $Module -Force
        }
    }
    Write-Host "✅ Modules ready!" -ForegroundColor Green
}

# Get tenant information for dynamic values
function Get-TenantDynamicValues {
    try {
        $org = Get-MgOrganization | Select-Object -First 1
        $domain = $org.VerifiedDomains | Where-Object { $_.IsDefault -eq $true } | Select-Object -ExpandProperty Name
        $companyInitials = ($domain -split '\.')[0].ToUpper()
        
        # Extract SharePoint domain (remove .onmicrosoft.com if present)
        $sharePointDomain = $domain -replace '\.onmicrosoft\.com$', ''
        
        return @{
            TenantDomain = $domain
            CompanyInitials = $companyInitials
            SharePointDomain = $sharePointDomain
            SharePointUrl = "https://$sharePointDomain.sharepoint.com"
            TenantId = $org.Id
            OrganizationName = $org.DisplayName
        }
    }
    catch {
        Write-Error "Failed to get tenant info: $($_.Exception.Message)"
        return $null
    }
}

# Resolve group name to ID for assignments
function Get-GroupId {
    param([string]$GroupName)
    
    try {
        $group = Get-MgGroup -Filter "displayName eq '$GroupName'" -ErrorAction Stop
        if ($group) {
            return $group.Id
        } else {
            Write-Warning "Group '$GroupName' not found"
            return $null
        }
    }
    catch {
        Write-Error "Failed to resolve group '$GroupName': $($_.Exception.Message)"
        return $null
    }
}

# Substitute dynamic values in policy JSON
function Set-DynamicValues {
    param(
        [string]$PolicyJson,
        [hashtable]$TenantValues
    )
    
    # Common substitutions based on export analysis
    $substitutions = @{
        # SharePoint URLs
        'https://contoso.sharepoint.com' = $TenantValues.SharePointUrl
        'contoso.sharepoint.com' = "$($TenantValues.SharePointDomain).sharepoint.com"
        
        # Company branding
        'CONTOSO' = $TenantValues.CompanyInitials
        'Contoso' = $TenantValues.CompanyInitials
        
        # LAPS admin naming - always use BITS-Admin pattern
        'Local Admin' = 'BITS-Admin-Local'
        'Administrator' = 'BITS-Admin-System'
        
        # OneDrive tenant ID (if needed)
        'b2b2b2b2-c3c3-d4d4-e5e5-f6f6f6f6f6f6' = $TenantValues.TenantId
    }
    
    $updatedJson = $PolicyJson
    foreach ($find in $substitutions.Keys) {
        $replace = $substitutions[$find]
        $updatedJson = $updatedJson -replace [regex]::Escape($find), $replace
    }
    
    return $updatedJson
}

# Create configuration policy
function New-ConfigurationPolicy {
    param(
        [hashtable]$PolicyConfig,
        [hashtable]$TenantValues
    )
    
    try {
        # Check if policy already exists
        $existingPolicy = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies" -Method GET | 
                         Where-Object { $_.value.name -eq $PolicyConfig.name }
        
        if ($existingPolicy.value) {
            Write-Host "⚠️  Policy '$($PolicyConfig.name)' already exists" -ForegroundColor Yellow
            return $existingPolicy.value[0]
        }
        
        # Apply dynamic value substitutions
        $policyJson = $PolicyConfig | ConvertTo-Json -Depth 20
        $updatedJson = Set-DynamicValues -PolicyJson $policyJson -TenantValues $TenantValues
        $updatedPolicy = $updatedJson | ConvertFrom-Json
        
        # Create the policy using Graph API
        $newPolicy = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies" -Method POST -Body $updatedJson
        
        Write-Host "✅ Created: $($PolicyConfig.name)" -ForegroundColor Green
        Write-Host "   Policy ID: $($newPolicy.id)" -ForegroundColor Gray
        Write-Host "   Platform: $($PolicyConfig.platforms)" -ForegroundColor Gray
        
        return $newPolicy
    }
    catch {
        Write-Error "❌ Failed to create policy '$($PolicyConfig.name)': $($_.Exception.Message)"
        if ($_.Exception.Response) {
            $errorDetails = $_.Exception.Response | ConvertFrom-Json
            Write-Error "Details: $($errorDetails.error.message)"
        }
        return $null
    }
}

# Assign policy to device group
function Set-PolicyAssignment {
    param(
        [string]$PolicyId,
        [string]$GroupId,
        [string]$PolicyName
    )
    
    try {
        $assignmentBody = @{
            assignments = @(
                @{
                    target = @{
                        '@odata.type' = '#microsoft.graph.groupAssignmentTarget'
                        groupId = $GroupId
                    }
                }
            )
        } | ConvertTo-Json -Depth 10
        
        $uri = "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies('$PolicyId')/assign"
        Invoke-MgGraphRequest -Uri $uri -Method POST -Body $assignmentBody
        
        Write-Host "   ✅ Assigned to device group" -ForegroundColor Green
    }
    catch {
        Write-Warning "Failed to assign policy '$PolicyName': $($_.Exception.Message)"
    }
}

# Determine appropriate device group for policy assignment
function Get-PolicyAssignmentGroup {
    param(
        [string]$PolicyName,
        [string]$Platform,
        [hashtable]$DeviceGroups
    )
    
    # Assignment logic based on policy name and platform
    switch -Wildcard ($PolicyName) {
        "*Windows*" { return $DeviceGroups['Windows'] }
        "*macOS*" { return $DeviceGroups['macOS'] }
        "*iOS*" { return $DeviceGroups['iOS'] }
        "*Android*" { return $DeviceGroups['Android'] }
        "*Pilot*" { return $DeviceGroups['Pilot'] }
        "*Corporate*" { return $DeviceGroups['Corporate'] }
        default {
            # Default assignment based on platform
            switch ($Platform) {
                "windows10" { return $DeviceGroups['Windows'] }
                "macOS" { return $DeviceGroups['macOS'] }
                "iOS" { return $DeviceGroups['iOS'] }
                "android" { return $DeviceGroups['Android'] }
                default { return $DeviceGroups['Windows'] } # Default to Windows
            }
        }
    }
}

# Enable LAPS in Entra ID tenant (required for LAPS policies)
function Enable-TenantLAPS {
    try {
        Write-Host "🔍 Checking LAPS tenant enablement..." -ForegroundColor Yellow
        
        # Check current LAPS setting
        $deviceRegistrationPolicy = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/policies/deviceRegistrationPolicy" -Method GET
        
        if ($deviceRegistrationPolicy.localAdminPassword.isEnabled -eq $false) {
            Write-Host "LAPS is disabled in tenant. Enabling now..." -ForegroundColor Yellow
            
            $updateBody = @{
                localAdminPassword = @{
                    isEnabled = $true
                }
            } | ConvertTo-Json -Depth 10
            
            Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/policies/deviceRegistrationPolicy" -Method PATCH -Body $updateBody
            Write-Host "✅ LAPS enabled in tenant" -ForegroundColor Green
        } else {
            Write-Host "✅ LAPS already enabled in tenant" -ForegroundColor Green
        }
    }
    catch {
        Write-Warning "Failed to enable LAPS in tenant: $($_.Exception.Message)"
        Write-Host "💡 Manually enable: Entra admin center > Identity > Devices > Device settings > Enable LAPS" -ForegroundColor Yellow
    }
}

# Load policy definitions from SettingsCatalog.json export
function Get-PolicyDefinitions {
    # Filter out auto-generated and Autopatch policies
    $allPolicies = @(
        @{
            name = "Default Web Pages"
            description = ""
            platforms = "windows10"
            technologies = "mdm"
            templateReference = @{
                templateFamily = "none"
                templateId = ""
            }
            settingCount = 3
        },
        @{
            name = "Defender Configuration"
            description = "Microsoft Defender comprehensive security configuration"
            platforms = "windows10"
            technologies = "mdm"
            templateReference = @{
                templateFamily = "none"
                templateId = ""
            }
            settingCount = 1
        },
        @{
            name = "Disable UAC for Quickassist"
            description = "Disable UAC secure desktop prompt for QuickAssist"
            platforms = "windows10"
            technologies = "mdm"
            templateReference = @{
                templateFamily = "none"
                templateId = ""
            }
            settingCount = 1
        },
        @{
            name = "Edge Update Policy"
            description = "Microsoft Edge update configuration"
            platforms = "windows10"
            technologies = "mdm"
            templateReference = @{
                templateFamily = "none"
                templateId = ""
            }
            settingCount = 4
        },
        @{
            name = "EDR Policy"
            description = ""
            platforms = "windows10"
            technologies = "mdm,microsoftSense"
            templateReference = @{
                templateFamily = "endpointSecurityEndpointDetectionAndResponse"
                templateId = "0385b795-0f2f-44ac-8602-9f65bf6adede_1"
                templateDisplayName = "Endpoint detection and response"
                templateDisplayVersion = "Version 1"
            }
            settingCount = 1
        },
        @{
            name = "Enable Bitlocker"
            description = "Comprehensive BitLocker drive encryption configuration"
            platforms = "windows10"
            technologies = "mdm"
            templateReference = @{
                templateFamily = "none"
                templateId = ""
            }
            settingCount = 13
        },
        @{
            name = "Enable Built-in Administrator Account"
            description = "Enable and configure built-in administrator account for LAPS"
            platforms = "windows10"
            technologies = "mdm"
            templateReference = @{
                templateFamily = "none"
                templateId = ""
            }
            settingCount = 2
        },
        @{
            name = "LAPS"
            description = ""
            platforms = "windows10"
            technologies = "mdm"
            templateReference = @{
                templateFamily = "endpointSecurityAccountProtection"
                templateId = "adc46e5a-f4aa-4ff6-aeff-4f27bc525796_1"
                templateDisplayName = "Local admin password solution (Windows LAPS)"
                templateDisplayVersion = "Version 1"
            }
            settingCount = 4
        },
        @{
            name = "Office Updates Configuration"
            description = "Microsoft Office update settings"
            platforms = "windows10"
            technologies = "mdm"
            templateReference = @{
                templateFamily = "none"
                templateId = ""
            }
            settingCount = 1
        },
        @{
            name = "OneDrive Configuration"
            description = "OneDrive for Business configuration with Known Folder Move"
            platforms = "windows10"
            technologies = "mdm"
            templateReference = @{
                templateFamily = "none"
                templateId = ""
            }
            settingCount = 7
        },
        @{
            name = "Outlook Configuration"
            description = ""
            platforms = "windows10"
            technologies = "mdm"
            templateReference = @{
                templateFamily = "none"
                templateId = ""
            }
            settingCount = 3
        },
        @{
            name = "Power Options"
            description = "Comprehensive power management settings for devices"
            platforms = "windows10"
            technologies = "mdm"
            templateReference = @{
                templateFamily = "none"
                templateId = ""
            }
            settingCount = 6
        },
        @{
            name = "Prevent Users From Unenrolling Devices"
            description = "Prevent users from manually unenrolling devices from Intune"
            platforms = "windows10"
            technologies = "mdm"
            templateReference = @{
                templateFamily = "none"
                templateId = ""
            }
            settingCount = 1
        },
        @{
            name = "Sharepoint File Sync"
            description = ""
            platforms = "windows10"
            technologies = "mdm"
            templateReference = @{
                templateFamily = "none"
                templateId = ""
            }
            settingCount = 1
        },
        @{
            name = "System Services"
            description = ""
            platforms = "windows10"
            technologies = "mdm"
            templateReference = @{
                templateFamily = "none"
                templateId = ""
            }
            settingCount = 4
        },
        @{
            name = "Tamper Protection"
            description = "Windows Security tamper protection configuration"
            platforms = "windows10"
            technologies = "mdm,microsoftSense"
            templateReference = @{
                templateFamily = "endpointSecurityAntivirus"
                templateId = "d948ff9b-99cb-4ee0-8012-1fbc09685377_1"
                templateDisplayName = "Windows Security Experience"
                templateDisplayVersion = "Version 1"
            }
            settingCount = 1
        },
        @{
            name = "Web Sign-in Policy"
            description = "Allows for Sign in with Temp Pass"
            platforms = "windows10"
            technologies = "mdm"
            templateReference = @{
                templateFamily = "none"
                templateId = ""
            }
            settingCount = 1
        }
    )
    
    # Skip auto-generated policies (already filtered out from above list)
    # Auto-generated ones excluded: Firewall/NGP Windows default policy, Autopatch policies
    
    return $allPolicies
}

# Main execution function
function Start-ConfigurationPolicyCreation {
    Write-Host "`n🚀 Creating Intune Configuration Policies..." -ForegroundColor Cyan
    Write-Host "=" * 60 -ForegroundColor Cyan
    
    # Verify connection
    $context = Get-MgContext
    if (!$context) {
        Write-Error "❌ Not connected to Microsoft Graph. Please connect first."
        return
    }
    
    # Get tenant dynamic values
    $tenantValues = Get-TenantDynamicValues
    if (!$tenantValues) {
        Write-Error "❌ Failed to get tenant information"
        return
    }
    
    Write-Host "✅ Connected to: $($tenantValues.OrganizationName)" -ForegroundColor Green
    Write-Host "   SharePoint URL: $($tenantValues.SharePointUrl)" -ForegroundColor Gray
    Write-Host "   Company Initials: $($tenantValues.CompanyInitials)" -ForegroundColor Gray
    
    # Resolve device groups for assignments
    Write-Host "`n🔍 Resolving device groups..." -ForegroundColor Yellow
    $deviceGroups = @{
        'Windows' = Get-GroupId -GroupName "Windows Devices (Autopilot)"
        'macOS' = Get-GroupId -GroupName "macOS Devices"
        'iOS' = Get-GroupId -GroupName "iOS Devices"
        'Android' = Get-GroupId -GroupName "Android Devices"
        'Corporate' = Get-GroupId -GroupName "Corporate Owned Devices"
        'Personal' = Get-GroupId -GroupName "Personal Devices"
        'Pilot' = Get-GroupId -GroupName "Pilot Device Group"
    }
    
    $missingGroups = $deviceGroups.Keys | Where-Object { !$deviceGroups[$_] }
    if ($missingGroups) {
        Write-Warning "Missing device groups: $($missingGroups -join ', ')"
        Write-Host "💡 Run Device-Groups.ps1 first to create required groups" -ForegroundColor Yellow
    }
    
    # Enable LAPS in tenant if needed
    Enable-TenantLAPS
    
    # Load policy definitions
    $policies = Get-PolicyDefinitions
    if ($policies.Count -eq 0) {
        Write-Error "❌ No policy definitions found. Please add SettingsCatalog.json content to script."
        return
    }
    
    Write-Host "📋 Found $($policies.Count) policies to create" -ForegroundColor Yellow
    Write-Host "   (Auto-generated Autopatch/MDE policies excluded)" -ForegroundColor Gray
    
    # Create policies
    $createdPolicies = @()
    $failedPolicies = @()
    
    foreach ($policy in $policies) {
        # Skip Windows Autopatch policies (auto-generated)
        if ($policy.name -like "*Autopatch*" -or $policy.createdDateTime -like "*2024-*") {
            Write-Host "⏭️  Skipping auto-generated policy: $($policy.name)" -ForegroundColor Gray
            continue
        }
        
        Write-Host "`n📱 Creating: $($policy.name)" -ForegroundColor White
        
        $result = New-ConfigurationPolicy -PolicyConfig $policy -TenantValues $tenantValues
        
        if ($result) {
            $createdPolicies += $result
            
            # Assign to appropriate device group
            $assignmentGroup = Get-PolicyAssignmentGroup -PolicyName $policy.name -Platform $policy.platforms -DeviceGroups $deviceGroups
            if ($assignmentGroup) {
                Set-PolicyAssignment -PolicyId $result.id -GroupId $assignmentGroup -PolicyName $policy.name
            }
        } else {
            $failedPolicies += $policy.name
        }
        
        # Small delay to avoid throttling
        Start-Sleep -Milliseconds 500
    }
    
    # Summary
    Write-Host "`n" + "=" * 60 -ForegroundColor Cyan
    Write-Host "📊 SUMMARY" -ForegroundColor Cyan
    Write-Host "=" * 60 -ForegroundColor Cyan
    Write-Host "✅ Successfully created: $($createdPolicies.Count) policies" -ForegroundColor Green
    
    if ($failedPolicies.Count -gt 0) {
        Write-Host "❌ Failed to create: $($failedPolicies.Count) policies" -ForegroundColor Red
        foreach ($failed in $failedPolicies) {
            Write-Host "   - $failed" -ForegroundColor Red
        }
    }
    
    Write-Host "`n💡 Next Steps:" -ForegroundColor Yellow
    Write-Host "   1. Verify policies in Intune admin center" -ForegroundColor Gray
    Write-Host "   2. Check device group assignments" -ForegroundColor Gray
    Write-Host "   3. Create compliance policies to complement configuration" -ForegroundColor Gray
    Write-Host "   4. Monitor policy deployment status" -ForegroundColor Gray
    
    return $createdPolicies
}

# Initialize and run
try {
    Initialize-Modules
    $results = Start-ConfigurationPolicyCreation
    
    if ($results) {
        Write-Host "`n🎉 Configuration policy creation completed!" -ForegroundColor Green
    }
}
catch {
    Write-Error "❌ Script execution failed: $($_.Exception.Message)"
}

# ▼ CB & Claude | BITS 365 Automation | v1.0 | "Smarter not Harder"
