#Requires -Version 7.0

<#
.SYNOPSIS
    Creates security groups for Entra ID management
.DESCRIPTION
    Creates user security groups for MFA exclusions, admin identification, SSPR, and license-based grouping
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

# Security group definitions
$SecurityGroups = @(
    @{
        Name = "NoMFA Exclusion Group"
        Description = "Users excluded from all Conditional Access policies and MFA requirements"
        GroupType = "Assigned"
        SecurityEnabled = $true
        MailEnabled = $false
    },
    @{
        Name = "BITS Admin Users"
        Description = "Dynamic group containing all BITS admin users"
        MembershipRule = '(user.userPrincipalName -contains "BITS-Admin") or (user.displayName -contains "BITS-Admin")'
        GroupType = "DynamicMembership"
        SecurityEnabled = $true
        MailEnabled = $false
    },
    @{
        Name = "SSPR Eligible Users"
        Description = "All users eligible for Self-Service Password Reset (excludes BITS admins)"
        MembershipRule = '(user.userType -eq "Member") and not ((user.userPrincipalName -contains "BITS-Admin") or (user.displayName -contains "BITS-Admin"))'
        GroupType = "DynamicMembership"
        SecurityEnabled = $true
        MailEnabled = $false
    },
    @{
        Name = "License - Business Basic"
        Description = "Users with Business Basic licensing assignment"
        MembershipRule = '(user.extensionAttribute1 -eq "BusinessBasic")'
        GroupType = "DynamicMembership"
        SecurityEnabled = $true
        MailEnabled = $false
    },
    @{
        Name = "License - Business Standard"
        Description = "Users with Business Standard licensing assignment"
        MembershipRule = '(user.extensionAttribute1 -eq "BusinessStandard")'
        GroupType = "DynamicMembership"
        SecurityEnabled = $true
        MailEnabled = $false
    },
    @{
        Name = "License - Business Premium"
        Description = "Users with Business Premium licensing assignment"
        MembershipRule = '(user.extensionAttribute1 -eq "BusinessPremium")'
        GroupType = "DynamicMembership"
        SecurityEnabled = $true
        MailEnabled = $false
    },
    @{
        Name = "License - Exchange Online Plan 1"
        Description = "Users with Exchange Online Plan 1 licensing assignment"
        MembershipRule = '(user.extensionAttribute1 -eq "ExchangeOnline1")'
        GroupType = "DynamicMembership"
        SecurityEnabled = $true
        MailEnabled = $false
    },
    @{
        Name = "License - Exchange Online Plan 2"
        Description = "Users with Exchange Online Plan 2 licensing assignment"
        MembershipRule = '(user.extensionAttribute1 -eq "ExchangeOnline2")'
        GroupType = "DynamicMembership"
        SecurityEnabled = $true
        MailEnabled = $false
    }
)

# Create security group function
function New-SecurityGroup {
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
        
        # Base group parameters
        $groupParams = @{
            DisplayName = $GroupConfig.Name
            Description = $GroupConfig.Description
            SecurityEnabled = $GroupConfig.SecurityEnabled
            MailEnabled = $GroupConfig.MailEnabled
            MailNickname = $mailNickname
        }
        
        # Add dynamic membership settings if specified
        if ($GroupConfig.GroupType -eq "DynamicMembership") {
            $groupParams.GroupTypes = @("DynamicMembership")
            $groupParams.MembershipRule = $GroupConfig.MembershipRule
            $groupParams.MembershipRuleProcessingState = "On"
        } else {
            # Assigned group
            $groupParams.GroupTypes = @()
        }
        
        # Create the group
        $newGroup = New-MgGroup @groupParams
        
        Write-Host "✅ Created: $($GroupConfig.Name)" -ForegroundColor Green
        Write-Host "   Group ID: $($newGroup.Id)" -ForegroundColor Gray
        Write-Host "   Type: $($GroupConfig.GroupType)" -ForegroundColor Gray
        if ($GroupConfig.MembershipRule) {
            Write-Host "   Rule: $($GroupConfig.MembershipRule)" -ForegroundColor Gray
        }
        
        return $newGroup
    }
    catch {
        Write-Error "❌ Failed to create group '$($GroupConfig.Name)': $($_.Exception.Message)"
        return $null
    }
}

# Main execution function
function Start-SecurityGroupCreation {
    Write-Host "`n🚀 Creating Entra ID Security Groups..." -ForegroundColor Cyan
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
    
    foreach ($group in $SecurityGroups) {
        Write-Host "`n👥 Creating: $($group.Name)" -ForegroundColor White
        
        $result = New-SecurityGroup -GroupConfig $group
        
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
    Write-Host "   1. Wait 5-10 minutes for dynamic group membership to populate" -ForegroundColor Gray
    Write-Host "   2. Manually add break-glass accounts to 'NoMFA Exclusion Group'" -ForegroundColor Gray
    Write-Host "   3. Configure ExtensionAttribute1 for users to populate license groups" -ForegroundColor Gray
    Write-Host "   4. Create Conditional Access policies using these groups" -ForegroundColor Gray
    
    # Show important group IDs for reference
    Write-Host "`n🔑 Important Group IDs:" -ForegroundColor Yellow
    foreach ($group in $createdGroups) {
        if ($group.DisplayName -like "*NoMFA*" -or $group.DisplayName -like "*BITS Admin*") {
            Write-Host "   $($group.DisplayName): $($group.Id)" -ForegroundColor Gray
        }
    }
    
    return $createdGroups
}

# Initialize and run
try {
    Initialize-Modules
    $results = Start-SecurityGroupCreation
    
    Write-Host "`n🎉 Security group creation completed!" -ForegroundColor Green
}
catch {
    Write-Error "❌ Script execution failed: $($_.Exception.Message)"
}

# ▼ CB & Claude | BITS 365 Automation | v1.0 | "Smarter not Harder"
