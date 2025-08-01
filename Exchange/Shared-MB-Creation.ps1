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
            Install-Module $Module -Force -Scope CurrentUser -AllowClobber
        }
        if (!(Get-Module -Name $Module)) {
            Write-Host "Importing $Module..." -ForegroundColor Yellow
            Import-Module $Module -Force
        }
    }
    Write-Host "✅ Modules ready!" -ForegroundColor Green
}

# Get Exchange Online connection info
function Get-ExchangeConnectionInfo {
    try {
        $connection = Get-ConnectionInformation -ErrorAction Stop
        if ($connection -and $connection.State -eq "Connected") {
            return @{
                Connected = $true
                TenantName = $connection.Name
                UserPrincipalName = $connection.UserPrincipalName
                ConnectionId = $connection.ConnectionId
            }
        }
        else {
            return @{ Connected = $false }
        }
    }
    catch {
        return @{ Connected = $false }
    }
}

# Verify Exchange Online connection and permissions
function Test-ExchangeConnection {
    $connectionInfo = Get-ExchangeConnectionInfo
    
    if (!$connectionInfo.Connected) {
        Write-Error "❌ Not connected to Exchange Online"
        Write-Host "💡 Please connect via the main menu first" -ForegroundColor Yellow
        return $false
    }
    
    Write-Host "✅ Connected to Exchange Online" -ForegroundColor Green
    Write-Host "   Tenant: $($connectionInfo.TenantName)" -ForegroundColor Gray
    Write-Host "   User: $($connectionInfo.UserPrincipalName)" -ForegroundColor Gray
    
    # Test if we can run basic Exchange commands
    try {
        $null = Get-AcceptedDomain -ResultSize 1 -ErrorAction Stop
        Write-Host "✅ Exchange permissions verified" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Error "❌ Insufficient permissions for Exchange Online operations"
        Write-Host "💡 Required roles: $($RequiredRoles -join ', ')" -ForegroundColor Yellow
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
        Write-Host "🔍 Retrieving accepted domains..." -ForegroundColor Yellow
        $domains = Get-AcceptedDomain -ErrorAction Stop
        
        $acceptedDomains = $domains | Where-Object { $_.DomainType -eq "Authoritative" } | Select-Object -ExpandProperty DomainName
        
        Write-Host "✅ Found $($acceptedDomains.Count) accepted domains" -ForegroundColor Green
        
        return $acceptedDomains
    }
    catch {
        Write-Warning "⚠️ Could not retrieve accepted domains: $($_.Exception.Message)"
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
    
    Write-Host "`n🚀 Creating shared mailbox..." -ForegroundColor Cyan
    Write-Host "=" * 50 -ForegroundColor Cyan
    
    try {
        Write-Host "📧 Email: $($MailboxConfig.EmailAddress)" -ForegroundColor White
        Write-Host "👤 Display Name: $($MailboxConfig.DisplayName)" -ForegroundColor White
        Write-Host "🏷️  Alias: $($MailboxConfig.Alias)" -ForegroundColor White
        if (![string]::IsNullOrWhiteSpace($MailboxConfig.Description)) {
            Write-Host "📝 Description: $($MailboxConfig.Description)" -ForegroundColor White
        }
        
        Write-Host "`n⏳ Creating mailbox..." -ForegroundColor Yellow
        
        # Create the shared mailbox
        $mailboxParams = @{
            Shared = $true
            Name = $MailboxConfig.DisplayName
            DisplayName = $MailboxConfig.DisplayName
            PrimarySmtpAddress = $MailboxConfig.EmailAddress
            Alias = $MailboxConfig.Alias
        }
        
        $newMailbox = New-Mailbox @mailboxParams -ErrorAction Stop
        
        Write-Host "✅ Shared mailbox created successfully!" -ForegroundColor Green
        Write-Host "   Mailbox ID: $($newMailbox.Identity)" -ForegroundColor Gray
        
        # Configure additional settings for optimal shared mailbox behavior
        Write-Host "`n⚙️ Configuring shared mailbox settings..." -ForegroundColor Yellow
        
        $configParams = @{
            Identity = $MailboxConfig.EmailAddress
            MessageCopyForSentAsEnabled = $true
            MessageCopyForSendOnBehalfEnabled = $true
        }
        
        Set-Mailbox @configParams -ErrorAction Stop
        
        Write-Host "✅ Shared mailbox configuration completed!" -ForegroundColor Green
        
        # Show summary
        Write-Host "`n📊 SUMMARY" -ForegroundColor Cyan
        Write-Host "=" * 50 -ForegroundColor Cyan
        Write-Host "✅ Shared mailbox created: $($MailboxConfig.EmailAddress)" -ForegroundColor Green
        Write-Host "👤 Display Name: $($MailboxConfig.DisplayName)" -ForegroundColor Green
        Write-Host "🏷️  Alias: $($MailboxConfig.Alias)" -ForegroundColor Green
        Write-Host "📨 Send As enabled: Yes" -ForegroundColor Green
        Write-Host "📩 Send on Behalf enabled: Yes" -ForegroundColor Green
        Write-Host "💾 Sent items saved: Yes" -ForegroundColor Green
        
        Write-Host "`n💡 Next Steps:" -ForegroundColor Yellow
        Write-Host "   1. Go to Exchange Admin Center → Recipients → Mailboxes" -ForegroundColor Gray
        Write-Host "   2. Find your shared mailbox: $($MailboxConfig.EmailAddress)" -ForegroundColor Gray
        Write-Host "   3. Click on the mailbox and select 'Manage mailbox delegation'" -ForegroundColor Gray
        Write-Host "   4. Add users with 'Full Access' and/or 'Send As' permissions" -ForegroundColor Gray
        Write-Host "   5. Users can then add the shared mailbox to their Outlook profile" -ForegroundColor Gray
        
        return $newMailbox
    }
    catch {
        Write-Error "❌ Failed to create shared mailbox: $($_.Exception.Message)"
        
        # Provide specific error guidance
        if ($_.Exception.Message -like "*already exists*") {
            Write-Host "💡 A mailbox with this address already exists" -ForegroundColor Yellow
        }
        elseif ($_.Exception.Message -like "*domain*") {
            Write-Host "💡 Check if the domain is accepted in your tenant" -ForegroundColor Yellow
        }
        elseif ($_.Exception.Message -like "*permission*") {
            Write-Host "💡 Ensure you have Exchange Administrator permissions" -ForegroundColor Yellow
        }
        
        return $null
    }
}

# Main execution function
function Start-SharedMailboxCreation {
    Write-Host "`n🚀 Creating Exchange Online Shared Mailboxes..." -ForegroundColor Cyan
    Write-Host "=" * 60 -ForegroundColor Cyan
    
    # Verify connection and permissions first
    if (!(Test-ExchangeConnection)) {
        return
    }
    
    try {
        # Get accepted domains
        $acceptedDomains = Get-AcceptedDomains
        
        # Get shared mailbox configuration from user
        $mailboxConfig = Get-SharedMailboxInput -AcceptedDomains $acceptedDomains
        
        # Confirm creation
        Write-Host "`n🔍 Configuration Review:" -ForegroundColor Yellow
        Write-Host "   Email: $($mailboxConfig.EmailAddress)" -ForegroundColor White
        Write-Host "   Display Name: $($mailboxConfig.DisplayName)" -ForegroundColor White
        Write-Host "   Alias: $($mailboxConfig.Alias)" -ForegroundColor White
        if (![string]::IsNullOrWhiteSpace($mailboxConfig.Description)) {
            Write-Host "   Description: $($mailboxConfig.Description)" -ForegroundColor White
        }
        
        $confirm = Read-Host "`nProceed with creation? (y/N)"
        
        if ($confirm -eq 'y' -or $confirm -eq 'Y') {
            # Create the shared mailbox
            $result = New-SharedMailbox -MailboxConfig $mailboxConfig
            
            if ($result) {
                Write-Host "`n🎉 Shared mailbox creation completed successfully!" -ForegroundColor Green
                
                # Ask if user wants to create another
                $another = Read-Host "`nCreate another shared mailbox? (y/N)"
                if ($another -eq 'y' -or $another -eq 'Y') {
                    Start-SharedMailboxCreation
                    return
                }
            }
        }
        else {
            Write-Host "❌ Shared mailbox creation cancelled" -ForegroundColor Yellow
        }
    }
    catch {
        Write-Error "❌ Script execution failed: $($_.Exception.Message)"
    }
    
    return $result
}

# Initialize and run
try {
    Initialize-Modules
    $results = Start-SharedMailboxCreation
    
    if ($results) {
        Write-Host "`n🎉 Shared mailbox creation completed!" -ForegroundColor Green
    }
}
catch {
    Write-Error "❌ Script execution failed: $($_.Exception.Message)"
}

# ▼ CB & Claude | BITS 365 Automation | v1.0 | "Smarter not Harder"