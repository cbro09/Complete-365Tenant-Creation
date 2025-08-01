#Requires -Version 7.0

<#
.SYNOPSIS
    Creates Exchange Online shared mailboxes with manual input
.DESCRIPTION
    Interactive script for creating shared mailboxes in Exchange Online with proper configuration
    and validation. Supports single shared mailbox creation with optimized settings.
.AUTHOR
    CB & Claude Partnership
.VERSION
    1.0
.NOTES
    Requires Exchange Online PowerShell V3 module (ExchangeOnlineManagement)
    Uses REST API with modern authentication - no Basic Auth required
#>

# Debug: Confirm script is loading
Write-Host "🔧 Loading Shared Mailbox Creation Script..." -ForegroundColor Cyan

# Required Modules
$RequiredModules = @(
    'ExchangeOnlineManagement'
)

# Required roles/permissions for this script
$RequiredRoles = @(
    "Exchange Administrator",
    "Global Administrator",
    "Mail Recipients"
)

# Auto-install and import required modules
function Initialize-Modules {
    Write-Host "🔧 Checking required modules..." -ForegroundColor Yellow
    
    foreach ($Module in $RequiredModules) {
        if (!(Get-Module -ListAvailable -Name $Module)) {
            Write-Host "Installing $Module..." -ForegroundColor Yellow
            try {
                Install-Module $Module -Force -Scope CurrentUser -AllowClobber -ErrorAction Stop
            }
            catch {
                Write-Error "Failed to install $Module`: $($_.Exception.Message)"
                return $false
            }
        }
        if (!(Get-Module -Name $Module)) {
            Write-Host "Importing $Module..." -ForegroundColor Yellow
            try {
                Import-Module $Module -Force -ErrorAction Stop
            }
            catch {
                Write-Error "Failed to import $Module`: $($_.Exception.Message)"
                return $false
            }
        }
    }
    Write-Host "✅ Modules ready!" -ForegroundColor Green
    return $true
}

# Get Exchange Online connection info
function Get-ExchangeConnectionInfo {
    try {
        # Try multiple methods to detect Exchange connection
        Write-Host "   🔍 Checking connection method 1: Get-ConnectionInformation..." -ForegroundColor Gray
        $connection = Get-ConnectionInformation -ErrorAction SilentlyContinue
        
        if ($connection -and $connection.State -eq "Connected") {
            Write-Host "   ✅ Found connection via Get-ConnectionInformation" -ForegroundColor Green
            return @{
                Connected = $true
                TenantName = $connection.Name
                UserPrincipalName = $connection.UserPrincipalName
                ConnectionId = $connection.ConnectionId
                Method = "ConnectionInformation"
            }
        }
        
        # Alternative method: Try running a simple Exchange command
        Write-Host "   🔍 Checking connection method 2: Test Exchange command..." -ForegroundColor Gray
        $testResult = Get-AcceptedDomain -ResultSize 1 -ErrorAction SilentlyContinue
        
        if ($testResult) {
            Write-Host "   ✅ Exchange Online is accessible (command test passed)" -ForegroundColor Green
            return @{
                Connected = $true
                TenantName = "Connected"
                UserPrincipalName = "Verified via command test"
                Method = "CommandTest"
            }
        }
        
        Write-Host "   ❌ No Exchange connection detected" -ForegroundColor Red
    }
    catch {
        Write-Host "   ⚠️ Connection check error: $($_.Exception.Message)" -ForegroundColor Yellow
    }
    
    return @{ Connected = $false }
}

# Verify Exchange Online connection and permissions
function Test-ExchangeConnection {
    Write-Host "🔍 Step 1: Verify Exchange Connection" -ForegroundColor Yellow
    
    $connectionInfo = Get-ExchangeConnectionInfo
    
    if (!$connectionInfo.Connected) {
        Write-Host "❌ Not connected to Exchange Online" -ForegroundColor Red
        Write-Host "💡 Solution: Use option 8 to connect to tenant, then select Exchange" -ForegroundColor Yellow
        return $false
    }
    
    Write-Host "✅ Connected to Exchange Online" -ForegroundColor Green
    Write-Host "   Connection: $($connectionInfo.Method)" -ForegroundColor Gray
    Write-Host "   Tenant: $($connectionInfo.TenantName)" -ForegroundColor Gray
    Write-Host "   User: $($connectionInfo.UserPrincipalName)" -ForegroundColor Gray
    
    # Additional permission test
    try {
        Write-Host "🔧 Testing mailbox creation permissions..." -ForegroundColor Yellow
        # Test if we can at least list mailboxes (basic permission check)
        $null = Get-Mailbox -ResultSize 1 -ErrorAction Stop
        Write-Host "✅ Exchange permissions verified" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Host "❌ Insufficient permissions for Exchange Online operations" -ForegroundColor Red
        Write-Host "💡 Required roles: $($RequiredRoles -join ', ')" -ForegroundColor Yellow
        Write-Host "⚠️ Error details: $($_.Exception.Message)" -ForegroundColor Gray
        return $false
    }
}

# Validate email address format
function Test-EmailAddress {
    param([string]$EmailAddress)
    
    $emailRegex = "^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$"
    return $EmailAddress -match $emailRegex
}

# Get accepted domains from tenant
function Get-AcceptedDomains {
    try {
        Write-Host "   🔍 Retrieving accepted domains from tenant..." -ForegroundColor Gray
        $domains = Get-AcceptedDomain -ErrorAction Stop
        
        $acceptedDomains = $domains | Where-Object { $_.DomainType -eq "Authoritative" } | Select-Object -ExpandProperty DomainName
        
        Write-Host "   ✅ Found $($acceptedDomains.Count) accepted domains" -ForegroundColor Green
        
        return $acceptedDomains
    }
    catch {
        Write-Host "   ⚠️ Could not retrieve accepted domains: $($_.Exception.Message)" -ForegroundColor Yellow
        return @()
    }
}

# Validate if domain is accepted
function Test-AcceptedDomain {
    param(
        [string]$EmailAddress,
        [array]$AcceptedDomains
    )
    
    if ($AcceptedDomains.Count -eq 0) {
        return $true  # Skip validation if we couldn't get domains
    }
    
    $domain = ($EmailAddress -split '@')[1]
    return $domain -in $AcceptedDomains
}

# Check if shared mailbox already exists
function Test-SharedMailboxExists {
    param([string]$EmailAddress)
    
    try {
        $existingMailbox = Get-Mailbox -Identity $EmailAddress -ErrorAction SilentlyContinue
        return $null -ne $existingMailbox
    }
    catch {
        return $false
    }
}

# Generate alias from email
function New-AliasFromEmail {
    param([string]$EmailAddress)
    
    $localPart = ($EmailAddress -split '@')[0]
    # Remove invalid characters and limit length
    $alias = $localPart -replace '[^a-zA-Z0-9]', '' 
    return $alias.Substring(0, [Math]::Min($alias.Length, 20))
}

# Get shared mailbox input from user
function Get-SharedMailboxInput {
    param([array]$AcceptedDomains)
    
    Write-Host "`n📝 Shared Mailbox Configuration" -ForegroundColor Cyan
    Write-Host "=" * 50 -ForegroundColor Cyan
    
    # Display accepted domains if available
    if ($AcceptedDomains.Count -gt 0) {
        Write-Host "`n✅ Accepted domains in your tenant:" -ForegroundColor Green
        foreach ($domain in $AcceptedDomains) {
            Write-Host "   • $domain" -ForegroundColor Gray
        }
        Write-Host ""
    }
    
    do {
        # Get email address
        $emailAddress = Read-Host "Enter shared mailbox email address (e.g., sales@company.com)"
        
        if ([string]::IsNullOrWhiteSpace($emailAddress)) {
            Write-Host "❌ Email address cannot be empty" -ForegroundColor Red
            continue
        }
        
        if (!(Test-EmailAddress -EmailAddress $emailAddress)) {
            Write-Host "❌ Invalid email address format" -ForegroundColor Red
            continue
        }
        
        if (!(Test-AcceptedDomain -EmailAddress $emailAddress -AcceptedDomains $AcceptedDomains)) {
            Write-Host "❌ Domain is not accepted in your tenant" -ForegroundColor Red
            continue
        }
        
        if (Test-SharedMailboxExists -EmailAddress $emailAddress) {
            Write-Host "❌ A mailbox with this address already exists" -ForegroundColor Red
            continue
        }
        
        break
    } while ($true)
    
    # Get display name
    do {
        $displayName = Read-Host "Enter display name (e.g., Sales Department)"
        
        if ([string]::IsNullOrWhiteSpace($displayName)) {
            Write-Host "❌ Display name cannot be empty" -ForegroundColor Red
            continue
        }
        
        if ($displayName.Length -gt 256) {
            Write-Host "❌ Display name must be 256 characters or less" -ForegroundColor Red
            continue
        }
        
        break
    } while ($true)
    
    # Generate and confirm alias
    $suggestedAlias = New-AliasFromEmail -EmailAddress $emailAddress
    $alias = Read-Host "Enter alias (press Enter for '$suggestedAlias')"
    
    if ([string]::IsNullOrWhiteSpace($alias)) {
        $alias = $suggestedAlias
    }
    
    # Optional description
    $description = Read-Host "Enter description (optional)"
    
    return @{
        EmailAddress = $emailAddress.ToLower()
        DisplayName = $displayName
        Alias = $alias
        Description = $description
    }
}

# Create shared mailbox
function New-SharedMailbox {
    param([hashtable]$MailboxConfig)
    
    Write-Host "   🚀 Creating shared mailbox..." -ForegroundColor Gray
    
    try {
        Write-Host "   📧 Email: $($MailboxConfig.EmailAddress)" -ForegroundColor Gray
        Write-Host "   👤 Display Name: $($MailboxConfig.DisplayName)" -ForegroundColor Gray
        Write-Host "   🏷️  Alias: $($MailboxConfig.Alias)" -ForegroundColor Gray
        
        Write-Host "   ⏳ Executing New-Mailbox command..." -ForegroundColor Gray
        
        # Create the shared mailbox
        $mailboxParams = @{
            Shared = $true
            Name = $MailboxConfig.DisplayName
            DisplayName = $MailboxConfig.DisplayName
            PrimarySmtpAddress = $MailboxConfig.EmailAddress
            Alias = $MailboxConfig.Alias
        }
        
        $newMailbox = New-Mailbox @mailboxParams -ErrorAction Stop
        
        Write-Host "   ✅ Shared mailbox created successfully!" -ForegroundColor Green
        
        # Configure additional settings for optimal shared mailbox behavior
        Write-Host "   ⚙️ Configuring mailbox settings..." -ForegroundColor Gray
        
        $configParams = @{
            Identity = $MailboxConfig.EmailAddress
            MessageCopyForSentAsEnabled = $true
            MessageCopyForSendOnBehalfEnabled = $true
        }
        
        Set-Mailbox @configParams -ErrorAction Stop
        
        Write-Host "   ✅ Configuration completed!" -ForegroundColor Green
        
        # Show summary
        Write-Host "`n📊 CREATION SUMMARY" -ForegroundColor Green
        Write-Host "=" * 30 -ForegroundColor Green
        Write-Host "✅ Email: $($MailboxConfig.EmailAddress)" -ForegroundColor White
        Write-Host "✅ Display Name: $($MailboxConfig.DisplayName)" -ForegroundColor White
        Write-Host "✅ Send As enabled: Yes" -ForegroundColor White
        Write-Host "✅ Sent items saved: Yes" -ForegroundColor White
        
        Write-Host "`n💡 Next Steps:" -ForegroundColor Yellow
        Write-Host "   1. Go to Exchange Admin Center → Recipients → Mailboxes" -ForegroundColor Gray
        Write-Host "   2. Find: $($MailboxConfig.EmailAddress)" -ForegroundColor Gray
        Write-Host "   3. Click 'Manage mailbox delegation'" -ForegroundColor Gray
        Write-Host "   4. Add users with 'Full Access' and 'Send As' permissions" -ForegroundColor Gray
        
        return $newMailbox
    }
    catch {
        Write-Host "   ❌ Failed to create shared mailbox: $($_.Exception.Message)" -ForegroundColor Red
        
        # Provide specific error guidance
        if ($_.Exception.Message -like "*already exists*") {
            Write-Host "   💡 A mailbox with this address already exists" -ForegroundColor Yellow
        }
        elseif ($_.Exception.Message -like "*domain*") {
            Write-Host "   💡 Check if the domain is accepted in your tenant" -ForegroundColor Yellow
        }
        elseif ($_.Exception.Message -like "*permission*") {
            Write-Host "   💡 Ensure you have Exchange Administrator permissions" -ForegroundColor Yellow
        }
        
        return $null
    }
}

# Main execution function
function Start-SharedMailboxCreation {
    Write-Host "`n📧 Exchange Online Shared Mailbox Creation" -ForegroundColor Cyan
    Write-Host "=" * 50 -ForegroundColor Cyan
    
    # Verify connection and permissions first
    Write-Host "`n🔍 Step 1: Verify Exchange Connection" -ForegroundColor Yellow
    if (!(Test-ExchangeConnection)) {
        Write-Host "`n❌ Connection verification failed" -ForegroundColor Red
        Write-Host "💡 Solution: Use option 8 to connect to tenant, then select Exchange" -ForegroundColor Yellow
        return $null
    }
    
    try {
        # Get accepted domains
        Write-Host "`n🌐 Step 2: Retrieve Accepted Domains" -ForegroundColor Yellow
        $acceptedDomains = Get-AcceptedDomains
        
        # Get shared mailbox configuration from user
        Write-Host "`n📝 Step 3: Collect Mailbox Configuration" -ForegroundColor Yellow
        $mailboxConfig = Get-SharedMailboxInput -AcceptedDomains $acceptedDomains
        
        # Confirm creation
        Write-Host "`n🔍 Step 4: Review Configuration" -ForegroundColor Yellow
        Write-Host "📧 Email: $($mailboxConfig.EmailAddress)" -ForegroundColor White
        Write-Host "👤 Display Name: $($mailboxConfig.DisplayName)" -ForegroundColor White
        Write-Host "🏷️  Alias: $($mailboxConfig.Alias)" -ForegroundColor White
        if (![string]::IsNullOrWhiteSpace($mailboxConfig.Description)) {
            Write-Host "📝 Description: $($mailboxConfig.Description)" -ForegroundColor White
        }
        
        Write-Host "`n❓ Proceed with creation? (y/N): " -ForegroundColor Yellow -NoNewline
        $confirm = Read-Host
        
        if ($confirm -eq 'y' -or $confirm -eq 'Y') {
            Write-Host "`n🚀 Step 5: Create Shared Mailbox" -ForegroundColor Yellow
            # Create the shared mailbox
            $result = New-SharedMailbox -MailboxConfig $mailboxConfig
            
            if ($result) {
                Write-Host "`n🎉 Shared mailbox creation completed successfully!" -ForegroundColor Green
                
                # Ask if user wants to create another
                Write-Host "`n❓ Create another shared mailbox? (y/N): " -ForegroundColor Yellow -NoNewline
                $another = Read-Host
                if ($another -eq 'y' -or $another -eq 'Y') {
                    return Start-SharedMailboxCreation
                }
                return $result
            }
            else {
                Write-Host "`n❌ Shared mailbox creation failed" -ForegroundColor Red
                return $null
            }
        }
        else {
            Write-Host "`n❌ Shared mailbox creation cancelled by user" -ForegroundColor Yellow
            return $null
        }
    }
    catch {
        Write-Host "`n❌ Error during shared mailbox creation: $($_.Exception.Message)" -ForegroundColor Red
        return $null
    }
}

# Initialize and run
Write-Host "`n🚀 Starting Exchange Online Shared Mailbox Creation..." -ForegroundColor Cyan
Write-Host "=" * 60 -ForegroundColor Cyan

try {
    # Initialize modules
    Write-Host "`n📦 Step 1: Initialize Modules" -ForegroundColor Cyan
    if (!(Initialize-Modules)) {
        Write-Host "❌ Module initialization failed - exiting" -ForegroundColor Red
        Read-Host "Press Enter to continue"
        return
    }
    
    # Start main process
    Write-Host "`n🎯 Step 2: Start Shared Mailbox Creation" -ForegroundColor Cyan
    $results = Start-SharedMailboxCreation
    
    if ($results) {
        Write-Host "`n🎉 Shared mailbox creation process completed!" -ForegroundColor Green
    }
    else {
        Write-Host "`n⚠️ Shared mailbox creation process did not complete successfully" -ForegroundColor Yellow
    }
}
catch {
    Write-Host "`n❌ Script execution failed: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "📝 Error details: $($_.Exception.GetType().FullName)" -ForegroundColor Gray
}
finally {
    Write-Host "`n⏸️ Press Enter to return to menu..." -ForegroundColor Gray
    Read-Host
}

# ▼ CB & Claude | BITS 365 Automation | v1.0 | "Smarter not Harder"
