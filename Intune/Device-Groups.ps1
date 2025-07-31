#Requires -Version 7.0

<#
.SYNOPSIS
    Creates dynamic device groups for Intune management
.DESCRIPTION
    Creates OS-specific dynamic security groups for device management and policy assignment
.AUTHOR
    CB & Claude Partnership
.VERSION
    1.0
#>

# Required Modules
$RequiredModules = @(
    'Microsoft.Graph.Authentication',
    'Microsoft.Graph.Groups'
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

# Device group definitions
$DeviceGroups = @(
    @{
        Name = "Windows Devices (Autopilot)"
        Description = "All Windows (Autopilot) devices managed by Intune"
        MembershipRule = '(device.devicePhysicalIds -any _ -eq "[OrderID]:WIN-AP-Corp")'
        GroupType = "DynamicMembership"
    },
    @{
        Name = "macOS Devices"
        Description = "All macOS devices managed by Intune"
        MembershipRule = '(device.deviceOSType -eq "macOS")'
        GroupType = "DynamicMembership"
    },
    @{
        Name = "iOS Devices"
        Description = "All iOS devices managed by Intune"
        MembershipRule = '(device.deviceOSType -eq "iOS")'
        GroupType = "DynamicMembership"
    },
    @{
        Name = "Android Devices"
        Description = "All Android devices managed by Intune"
        MembershipRule = '(device.deviceOSType -eq "Android")'
        GroupType = "DynamicMembership"
    },
    @{
        Name = "Corporate Owned Devices"
        Description = "All corporate owned devices"
        MembershipRule = '(device.deviceOwnership -eq "Company")'
        GroupType = "DynamicMembership"
    },
    @{
        Name = "Personal Devices"
        Description = "All personal owned devices"
        MembershipRule = '(device.deviceOwnership -eq "Personal")'
        GroupType = "DynamicMembership"
    },
    @{
        Name = "Pilot Device Group"
        Description = "Test group for pilot deployments"
        MembershipRule = '(device.displayName -startsWith "PILOT-")'
        GroupType = "DynamicMembership"
    }
)

# Create device group function
function New-DeviceGroup {
    param(
        [hashtable]$GroupConfig
    )
    
    try {
        # Check if group already exists
        $existingGroup = Get-MgGroup -Filter "displayName eq '$($GroupConfig.Name)'" -ErrorAction SilentlyContinue
        
        if ($existingGroup) {
            Write-Host "⚠️  Group '$($GroupConfig.Name)' already exists" -ForegroundColor Yellow
            return $existingGroup
        }
        
        # Create mail nickname (required for all groups)
        $mailNickname = $GroupConfig.Name -replace '[^a-zA-Z0-9]', '' -replace '\s', ''
        
        # Create group parameters
        $groupParams = @{
            DisplayName = $GroupConfig.Name
            Description = $GroupConfig.Description
            GroupTypes = @("DynamicMembership")
            SecurityEnabled = $true
            MailEnabled = $false
            MailNickname = $mailNickname
            MembershipRule = $GroupConfig.MembershipRule
            MembershipRuleProcessingState = "On"
        }
        
        # Create the group
        $newGroup = New-MgGroup @groupParams
        
        Write-Host "✅ Created: $($GroupConfig.Name)" -ForegroundColor Green
        Write-Host "   Group ID: $($newGroup.Id)" -ForegroundColor Gray
        Write-Host "   Rule: $($GroupConfig.MembershipRule)" -ForegroundColor Gray
        
        return $newGroup
    }
    catch {
        Write-Error "❌ Failed to create group '$($GroupConfig.Name)': $($_.Exception.Message)"
        return $null
    }
}

# Main execution function
function Start-DeviceGroupCreation {
    Write-Host "`n🚀 Creating Intune Device Groups..." -ForegroundColor Cyan
    Write-Host "=" * 60 -ForegroundColor Cyan
    
    # Verify connection
    $context = Get-MgContext
    if (!$context) {
        Write-Error "❌ Not connected to Microsoft Graph. Please connect first."
        return
    }
    
    Write-Host "✅ Connected to tenant: $($context.TenantId)" -ForegroundColor Green
    
    $createdGroups = @()
    $failedGroups = @()
    
    foreach ($group in $DeviceGroups) {
        Write-Host "`n📱 Creating: $($group.Name)" -ForegroundColor White
        
        $result = New-DeviceGroup -GroupConfig $group
        
        if ($result) {
            $createdGroups += $result
        } else {
            $failedGroups += $group.Name
        }
        
        # Small delay to avoid throttling
        Start-Sleep -Milliseconds 500
    }
    
    # Summary
    Write-Host "`n" + "=" * 60 -ForegroundColor Cyan
    Write-Host "📊 SUMMARY" -ForegroundColor Cyan
    Write-Host "=" * 60 -ForegroundColor Cyan
    Write-Host "✅ Successfully created: $($createdGroups.Count) groups" -ForegroundColor Green
    
    if ($failedGroups.Count -gt 0) {
        Write-Host "❌ Failed to create: $($failedGroups.Count) groups" -ForegroundColor Red
        foreach ($failed in $failedGroups) {
            Write-Host "   - $failed" -ForegroundColor Red
        }
    }
    
    Write-Host "`n💡 Next Steps:" -ForegroundColor Yellow
    Write-Host "   1. Wait 5-10 minutes for group membership to populate" -ForegroundColor Gray
    Write-Host "   2. Create configuration and compliance policies" -ForegroundColor Gray
    Write-Host "   3. Assign policies to these device groups" -ForegroundColor Gray
    
    return $createdGroups
}

# Initialize and run
try {
    Initialize-Modules
    $results = Start-DeviceGroupCreation
    
    Write-Host "`n🎉 Device group creation completed!" -ForegroundColor Green
}
catch {
    Write-Error "❌ Script execution failed: $($_.Exception.Message)"
}

# ▼ CB & Claude | BITS 365 Automation | v1.0 | "Smarter not Harder"