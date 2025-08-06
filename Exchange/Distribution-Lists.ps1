#Requires -Version 7.0

<#
.SYNOPSIS
    Exchange Online Distribution List Manual Creation Script
.DESCRIPTION
    Interactive script for manually creating distribution lists in Exchange Online.
    Integrates seamlessly with the Complete-365Tenant-Creation main menu system.
.AUTHOR
    CB & Claude Partnership
.VERSION
    1.0
.NOTES
    - Requires Exchange Online PowerShell v3.5.1+
    - Must be called through main menu authentication system
    - Uses modern REST API (no Basic Authentication)
.EXAMPLE
    Invoke-GitHubScript -ScriptPath "Exchange/Distribution-Lists.ps1"
#>

# Required Modules and Roles
$RequiredModules = @(
    'ExchangeOnlineManagement'
)

$RequiredRoles = @(
    "Exchange Administrator",
    "Global Administrator", 
    "Mail Recipients"
)

# Initialize modules function
function Initialize-Modules {
    <#
    .SYNOPSIS
        Checks and imports required PowerShell modules
    .DESCRIPTION
        Verifies ExchangeOnlineManagement module is available and properly imports it
    #>
    
    Write-Host "🔧 Checking Exchange Online PowerShell module..." -ForegroundColor Cyan
    
    foreach ($Module in $RequiredModules) {
        try {
            # Check if module is installed
            $installedModule = Get-Module -ListAvailable -Name $Module | Sort-Object Version -Descending | Select-Object -First 1
            
            if (!$installedModule) {
                Write-Host "❌ Module $Module not found. Installing..." -ForegroundColor Yellow
                Install-Module $Module -Force -Scope CurrentUser -AllowClobber
                Write-Host "✅ Module $Module installed successfully" -ForegroundColor Green
                $installedModule = Get-Module -ListAvailable -Name $Module | Sort-Object Version -Descending | Select-Object -First 1
            } else {
                Write-Host "✅ Module $Module found (Version: $($installedModule.Version))" -ForegroundColor Green
            }
            
            # Remove any existing module to ensure clean import
            Remove-Module $Module -Force -ErrorAction SilentlyContinue
            
            # Import module with explicit parameters
            Write-Host "📦 Importing $Module module..." -ForegroundColor Cyan
            Import-Module $Module -Force -Global -Scope Global
            
            # Verify key cmdlets are available after import
            $keyCmdlets = @('Connect-ExchangeOnline', 'Get-AcceptedDomain', 'New-DistributionGroup')
            $missingAfterImport = @()
            
            foreach ($cmdlet in $keyCmdlets) {
                if (!(Get-Command $cmdlet -ErrorAction SilentlyContinue)) {
                    $missingAfterImport += $cmdlet
                }
            }
            
            if ($missingAfterImport.Count -gt 0) {
                Write-Host "⚠️  Some cmdlets not available after import: $($missingAfterImport -join ', ')" -ForegroundColor Yellow
                Write-Host "   This may indicate a module version or installation issue" -ForegroundColor Yellow
            } else {
                Write-Host "✅ Module $Module imported successfully with all required cmdlets" -ForegroundColor Green
            }
        }
        catch {
            Write-Error "❌ Failed to initialize module ${Module}: $($_.Exception.Message)"
            return $false
        }
    }
    return $true
}

# Test Exchange Online connection
function Test-ExchangeConnection {
    <#
    .SYNOPSIS
        Tests if Exchange Online connection is active
    .DESCRIPTION
        Verifies connection or establishes new Exchange Online PowerShell connection
    #>
    
    Write-Host "🔗 Testing Exchange Online connection..." -ForegroundColor Cyan
    
    # Try multiple connection test methods
    $connectionMethods = @(
        { Get-AcceptedDomain -ErrorAction Stop | Select-Object -First 1 },
        { Get-OrganizationConfig -ErrorAction Stop | Select-Object -First 1 },
        { Get-ConnectionInformation -ErrorAction Stop | Select-Object -First 1 }
    )
    
    foreach ($method in $connectionMethods) {
        try {
            $null = & $method
            Write-Host "✅ Exchange Online connection active" -ForegroundColor Green
            return $true
        }
        catch {
            # Continue to next method
        }
    }
    
    # If no existing connection, try to establish one
    Write-Host "⚠️  No active Exchange Online PowerShell connection found" -ForegroundColor Yellow
    Write-Host "🔄 Attempting to establish Exchange Online connection..." -ForegroundColor Cyan
    
    try {
        # Try to connect using modern auth
        Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
        
        # Test the new connection
        $null = Get-AcceptedDomain -ErrorAction Stop | Select-Object -First 1
        Write-Host "✅ Exchange Online connection established successfully!" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Host "❌ Failed to establish Exchange Online connection" -ForegroundColor Red
        Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host ""
        Write-Host "💡 Troubleshooting steps:" -ForegroundColor Yellow
        Write-Host "   1. Ensure you have the required permissions: $($RequiredRoles -join ', ')" -ForegroundColor White
        Write-Host "   2. Try connecting manually: Connect-ExchangeOnline" -ForegroundColor White
        Write-Host "   3. Verify your account has Exchange Online license" -ForegroundColor White
        Write-Host "   4. Check if MFA is properly configured" -ForegroundColor White
        return $false
    }
}

# Validate email address format
function Test-EmailFormat {
    param(
        [Parameter(Mandatory)]
        [string]$EmailAddress
    )
    
    $emailRegex = "^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$"
    return $EmailAddress -match $emailRegex
}

# Check if distribution group already exists
function Test-DistributionGroupExists {
    param(
        [Parameter(Mandatory)]
        [string]$Identity
    )
    
    try {
        # Check if cmdlet is available first
        if (!(Get-Command 'Get-DistributionGroup' -ErrorAction SilentlyContinue)) {
            # If cmdlet not available, assume it doesn't exist (will be caught during creation)
            return $false
        }
        
        $null = Get-DistributionGroup -Identity $Identity -ErrorAction Stop
        return $true
    }
    catch {
        return $false
    }
}

# Get accepted domain for validation
function Get-AcceptedDomains {
    try {
        # First check if the cmdlet is available
        if (!(Get-Command 'Get-AcceptedDomain' -ErrorAction SilentlyContinue)) {
            Write-Host "⚠️  Get-AcceptedDomain cmdlet not available - Exchange Online connection may not be established" -ForegroundColor Yellow
            return @()
        }
        
        Write-Host "🔍 Retrieving accepted domains..." -ForegroundColor Cyan
        $domains = Get-AcceptedDomain -ErrorAction Stop | Select-Object -ExpandProperty DomainName
        Write-Host "✅ Retrieved $($domains.Count) accepted domain(s)" -ForegroundColor Green
        return $domains
    }
    catch {
        Write-Host "⚠️  Could not retrieve accepted domains: $($_.Exception.Message)" -ForegroundColor Yellow
        Write-Host "   Email validation will be limited to format checking only" -ForegroundColor Yellow
        return @()
    }
}

# Validate email domain against tenant domains
function Test-EmailDomain {
    param(
        [Parameter(Mandatory)]
        [string]$EmailAddress,
        
        [Parameter(Mandatory)]
        [string[]]$AcceptedDomains
    )
    
    $domain = ($EmailAddress -split '@')[1]
    return $domain -in $AcceptedDomains
}

# Get user input with validation
function Get-ValidatedInput {
    param(
        [Parameter(Mandatory)]
        [string]$Prompt,
        
        [Parameter(Mandatory)]
        [string]$ValidationScript,
        
        [string]$ErrorMessage = "Invalid input. Please try again.",
        
        [string]$Example = "",
        
        [switch]$AllowCancel
    )
    
    do {
        if ($Example -and $AllowCancel) {
            $userInput = Read-Host "$Prompt (Example: $Example) [Enter 'exit' to cancel]"
        } elseif ($Example) {
            $userInput = Read-Host "$Prompt (Example: $Example)"
        } elseif ($AllowCancel) {
            $userInput = Read-Host "$Prompt [Enter 'exit' to cancel]"
        } else {
            $userInput = Read-Host $Prompt
        }
        
        # Check for exit command
        if ($AllowCancel -and ($userInput -eq 'exit' -or $userInput -eq 'quit' -or $userInput -eq 'cancel')) {
            return $null
        }
        
        # Replace $input with $userInput in validation script to avoid PowerShell reserved variable conflict
        $validationScriptFixed = $ValidationScript -replace '\$input', '$userInput'
        $isValid = Invoke-Expression $validationScriptFixed
        
        if (!$isValid) {
            Write-Host $ErrorMessage -ForegroundColor Red
        }
    } while (!$isValid)
    
    return $userInput
}

# Interactive distribution group creation
function New-InteractiveDistributionGroup {
    <#
    .SYNOPSIS
        Interactive function to create a new distribution group
    .DESCRIPTION
        Collects user input and creates a distribution group with comprehensive validation
    #>
    
    Write-Host ""
    Write-Host "=" * 60 -ForegroundColor Blue
    Write-Host "📧 DISTRIBUTION GROUP CREATION WIZARD" -ForegroundColor Blue
    Write-Host "=" * 60 -ForegroundColor Blue
    Write-Host ""
    
    # Early exit option
    Write-Host "🚀 Ready to create a new distribution group!" -ForegroundColor Green
    $proceed = Read-Host "Continue with distribution group creation? (Y/n)"
    if ($proceed -like "n*") {
        Write-Host "❌ Distribution group creation cancelled" -ForegroundColor Yellow
        return
    }
    Write-Host ""
    
    # Get accepted domains for validation
    $acceptedDomains = Get-AcceptedDomains
    if ($acceptedDomains.Count -eq 0) {
        Write-Host "⚠️  Could not retrieve accepted domains - email validation will be basic format only" -ForegroundColor Yellow
        $acceptedDomains = @()  # Empty array for safety
    } else {
        Write-Host "📋 Available domains: $($acceptedDomains -join ', ')" -ForegroundColor Cyan
    }
    Write-Host ""
    
    # Collect Distribution Group Information
    Write-Host "📝 Please provide the following information:" -ForegroundColor Yellow
    Write-Host "💡 Type 'exit' at any prompt to cancel and return to menu" -ForegroundColor Cyan
    Write-Host ""
    
    # Group Name
    $groupName = Get-ValidatedInput -Prompt "Distribution Group Display Name" -ValidationScript "`$userInput.Length -gt 0 -and `$userInput.Length -le 256" -ErrorMessage "Group name cannot be empty and must be 256 characters or less" -Example "Marketing Team" -AllowCancel
    if ($null -eq $groupName) {
        Write-Host "❌ Distribution group creation cancelled" -ForegroundColor Yellow
        return
    }
    
    # Email Address
    if ($acceptedDomains.Count -gt 0) {
        $primaryEmail = Get-ValidatedInput -Prompt "Primary Email Address" -ValidationScript "(Test-EmailFormat `$userInput) -and (Test-EmailDomain `$userInput @('$($acceptedDomains -join "','")')) -and -not (Test-DistributionGroupExists `$userInput)" -ErrorMessage "Invalid email format, domain not accepted by tenant, or email already exists" -Example "marketing@$($acceptedDomains[0])" -AllowCancel
    } else {
        $primaryEmail = Get-ValidatedInput -Prompt "Primary Email Address" -ValidationScript "(Test-EmailFormat `$userInput) -and -not (Test-DistributionGroupExists `$userInput)" -ErrorMessage "Invalid email format or email already exists" -Example "marketing@yourdomain.com" -AllowCancel
    }
    if ($null -eq $primaryEmail) {
        Write-Host "❌ Distribution group creation cancelled" -ForegroundColor Yellow
        return
    }
    
    # Alias (derived from email if not specified)
    $suggestedAlias = ($primaryEmail -split '@')[0] -replace '[^a-zA-Z0-9]', ''
    $alias = Read-Host "Alias (press Enter to use '$suggestedAlias')"
    if ([string]::IsNullOrWhiteSpace($alias)) {
        $alias = $suggestedAlias
    }
    
    # Description
    $description = Read-Host "Description (optional)"
    if ([string]::IsNullOrWhiteSpace($description)) {
        $description = "Distribution group: $groupName"
    }
    
    # Owner (Manager)
    Write-Host ""
    Write-Host "👤 Group Owner Configuration:" -ForegroundColor Yellow
    $owner = Read-Host "Group Owner Email (press Enter to use current user)"
    
    # Join/Leave Restrictions
    Write-Host ""
    Write-Host "🔐 Group Membership Restrictions:" -ForegroundColor Yellow
    Write-Host "1. Open - Anyone can join/leave"
    Write-Host "2. Closed - Only owners can add/remove members"  
    Write-Host "3. ApprovalRequired - Owner approval required to join"
    
    do {
        $restrictionChoice = Read-Host "Select membership restriction (1-3)"
        switch ($restrictionChoice) {
            "1" { 
                $joinRestriction = "Open"
                $departRestriction = "Open"
                $valid = $true
            }
            "2" { 
                $joinRestriction = "Closed"
                $departRestriction = "Closed" 
                $valid = $true
            }
            "3" { 
                $joinRestriction = "ApprovalRequired"
                $departRestriction = "Closed"
                $valid = $true
            }
            default { 
                Write-Host "Invalid selection. Please choose 1, 2, or 3." -ForegroundColor Red
                $valid = $false
            }
        }
    } while (!$valid)
    
    # External Email Configuration
    Write-Host ""
    Write-Host "🌐 External Email Configuration:" -ForegroundColor Yellow
    $allowExternal = Read-Host "Allow external senders to email this group? (y/N)"
    $requireAuth = $allowExternal -notlike "y*"
    
    # Initial Members
    Write-Host ""
    Write-Host "👥 Initial Members (optional):" -ForegroundColor Yellow
    Write-Host "Enter email addresses separated by commas, or press Enter to skip"
    $membersInput = Read-Host "Initial members"
    $members = @()
    
    if (![string]::IsNullOrWhiteSpace($membersInput)) {
        $memberEmails = $membersInput -split ',' | ForEach-Object { $_.Trim() }
        foreach ($email in $memberEmails) {
            if (Test-EmailFormat $email) {
                $members += $email
            } else {
                Write-Host "⚠️  Invalid email format: $email (skipped)" -ForegroundColor Yellow
            }
        }
    }
    
    # Summary
    Write-Host ""
    Write-Host "📋 DISTRIBUTION GROUP SUMMARY:" -ForegroundColor Green
    Write-Host "=" * 40 -ForegroundColor Green
    Write-Host "Name: $groupName" -ForegroundColor White
    Write-Host "Email: $primaryEmail" -ForegroundColor White  
    Write-Host "Alias: $alias" -ForegroundColor White
    Write-Host "Description: $description" -ForegroundColor White
    Write-Host "Owner: $(if($owner) { $owner } else { 'Current user' })" -ForegroundColor White
    Write-Host "Join Restriction: $joinRestriction" -ForegroundColor White
    Write-Host "Leave Restriction: $departRestriction" -ForegroundColor White
    Write-Host "External Senders: $(if($requireAuth) { 'Blocked' } else { 'Allowed' })" -ForegroundColor White
    Write-Host "Initial Members: $(if($members.Count -gt 0) { $members.Count } else { 'None' })" -ForegroundColor White
    Write-Host ""
    
    $confirm = Read-Host "Create this distribution group? (Y/n)"
    if ($confirm -like "n*") {
        Write-Host "❌ Distribution group creation cancelled" -ForegroundColor Yellow
        return
    }
    
    # Create Distribution Group
    Write-Host ""
    Write-Host "🚀 Creating distribution group..." -ForegroundColor Cyan
    
    # Final check that required cmdlets are available
    if (!(Get-Command 'New-DistributionGroup' -ErrorAction SilentlyContinue)) {
        Write-Host "❌ New-DistributionGroup cmdlet not available" -ForegroundColor Red
        Write-Host "   Exchange Online PowerShell connection is not properly established" -ForegroundColor Red
        Write-Host "💡 Please run the script again to re-establish the connection" -ForegroundColor Yellow
        return
    }
    
    try {
        $dgParams = @{
            Name = $groupName
            Alias = $alias
            PrimarySmtpAddress = $primaryEmail
            Type = "Distribution"
            MemberJoinRestriction = $joinRestriction
            MemberDepartRestriction = $departRestriction
            RequireSenderAuthenticationEnabled = $requireAuth
        }
        
        if ($description) {
            $dgParams.Notes = $description
        }
        
        if ($owner) {
            $dgParams.ManagedBy = $owner
        }
        
        if ($members.Count -gt 0) {
            $dgParams.Members = $members
        }
        
        $newGroup = New-DistributionGroup @dgParams
        
        Write-Host "✅ Distribution group '$groupName' created successfully!" -ForegroundColor Green
        Write-Host "📧 Email address: $primaryEmail" -ForegroundColor Green
        Write-Host "🆔 Group ID: $($newGroup.ExternalDirectoryObjectId)" -ForegroundColor Green
        
        # Additional member management if needed
        if ($members.Count -eq 0) {
            Write-Host ""
            Write-Host "💡 Next Steps:" -ForegroundColor Yellow
            Write-Host "   • Add members via Exchange Admin Center" -ForegroundColor White
            Write-Host "   • Or use: Add-DistributionGroupMember -Identity '$primaryEmail' -Member 'user@domain.com'" -ForegroundColor White
        }
        
        Write-Host ""
        Write-Host "🔗 Manage this group at:" -ForegroundColor Cyan
        Write-Host "   https://admin.exchange.microsoft.com/#/recipients/groups/distribution" -ForegroundColor Blue
        
    }
    catch {
        Write-Host "❌ Failed to create distribution group" -ForegroundColor Red
        Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
        
        # Specific error guidance
        if ($_.Exception.Message -like "*already exists*") {
            Write-Host "💡 A group with this name or email address already exists" -ForegroundColor Yellow
        } elseif ($_.Exception.Message -like "*permission*") {
            Write-Host "💡 Insufficient permissions. Required roles: $($RequiredRoles -join ', ')" -ForegroundColor Yellow
        } elseif ($_.Exception.Message -like "*domain*") {
            Write-Host "💡 Email domain not accepted by your tenant" -ForegroundColor Yellow
        }
    }
}

# Main execution function
function Start-DistributionListCreation {
    <#
    .SYNOPSIS
        Main execution function for distribution list creation
    .DESCRIPTION
        Orchestrates the entire distribution list creation process
    #>
    
    Write-Host "🔧 Loading Distribution List Creation Script..." -ForegroundColor Cyan
    
    # Step 1: Initialize modules
    Write-Host "📦 Step 1: Initialize Modules" -ForegroundColor Cyan
    if (!(Initialize-Modules)) {
        Write-Host "❌ Module initialization failed. Cannot continue." -ForegroundColor Red
        return
    }
    
    # Step 2: Test Exchange connection  
    Write-Host "🔗 Step 2: Test Exchange Connection" -ForegroundColor Cyan
    if (!(Test-ExchangeConnection)) {
        Write-Host "❌ Exchange connection test failed. Cannot continue." -ForegroundColor Red
        Write-Host ""
        Write-Host "🛠️  Possible solutions:" -ForegroundColor Yellow
        Write-Host "   1. Run this script again (connection may have timed out)" -ForegroundColor White
        Write-Host "   2. Manually connect: Connect-ExchangeOnline" -ForegroundColor White
        Write-Host "   3. Verify you have Exchange Administrator permissions" -ForegroundColor White
        Write-Host "   4. Check ExchangeOnlineManagement module version (needs v3.0+)" -ForegroundColor White
        Write-Host ""
        Read-Host "Press Enter to return to Exchange menu"
        return
    }
    
    # Step 3: Start interactive creation
    Write-Host "🎯 Step 3: Start Distribution Group Creation" -ForegroundColor Cyan
    
    do {
        New-InteractiveDistributionGroup
        
        Write-Host ""
        $another = Read-Host "Create another distribution group? (y/N)"
        
    } while ($another -like "y*")
    
    Write-Host ""
    Write-Host "✅ Distribution list creation completed!" -ForegroundColor Green
    Write-Host "🔙 Returning to Exchange menu..." -ForegroundColor Cyan
    
    # Clean return to menu
    Start-Sleep -Seconds 2
}

# Execute main function
Start-DistributionListCreation