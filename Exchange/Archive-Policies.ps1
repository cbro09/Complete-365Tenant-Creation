#Requires -Version 7.0

<#
.SYNOPSIS
    Complete Mailbox Configuration for Exchange Online
.DESCRIPTION
    Comprehensive script that configures Exchange Online mailboxes with:
    - Archive mailbox enablement and verification
    - Storage quotas (Warning: 40GB, Prohibit Send: 45GB, Prohibit Send/Receive: 49GB)
    - Progress tracking and detailed error handling
    Integrated with Complete-365Tenant-Creation automation hub.
.AUTHOR
    CB & Claude Partnership - BITS 365 Automation
.VERSION
    1.0
.NOTES
    Requires ExchangeOnlineManagement module v3.2.0+
    Compatible with PowerShell 7.0+
    Uses REST-backed cmdlets (no WinRM Basic Auth required)
#>

# ═══════════════════════════════════════════════════════════════════════════════
# CONFIGURATION & INITIALIZATION
# ═══════════════════════════════════════════════════════════════════════════════

# Required modules and roles for this script
$RequiredModules = @('ExchangeOnlineManagement')
$RequiredRoles = @("Exchange Administrator", "Global Administrator", "Mail Recipients")

# Mailbox quota configuration (easily adjustable)
$QuotaConfig = @{
    WarningQuota = 40GB
    ProhibitSendQuota = 45GB  
    ProhibitSendReceiveQuota = 49GB
}

# Global variables for tracking
$Global:ProcessedMailboxes = 0
$Global:TotalMailboxes = 0
$Global:ProcessingErrors = @()

# ═══════════════════════════════════════════════════════════════════════════════
# UTILITY FUNCTIONS
# ═══════════════════════════════════════════════════════════════════════════════

function Write-Progress-Custom {
    <#
    .SYNOPSIS
        Enhanced progress display with detailed information
    #>
    param(
        [string]$Activity,
        [string]$Status,
        [int]$PercentComplete,
        [string]$CurrentOperation = ""
    )
    
    Write-Progress -Activity $Activity -Status $Status -PercentComplete $PercentComplete -CurrentOperation $CurrentOperation
    Write-Host "🔄 $Status" -ForegroundColor Cyan
    if ($CurrentOperation) {
        Write-Host "   └─ $CurrentOperation" -ForegroundColor Gray
    }
}

function Write-Status-Message {
    <#
    .SYNOPSIS
        Standardized status messaging with colors and emojis
    #>
    param(
        [string]$Message,
        [ValidateSet("Info", "Success", "Warning", "Error", "Step")]
        [string]$Type = "Info"
    )
    
    $emoji = switch ($Type) {
        "Info" { "ℹ️" }
        "Success" { "✅" }
        "Warning" { "⚠️" }
        "Error" { "❌" }
        "Step" { "🎯" }
    }
    
    $color = switch ($Type) {
        "Info" { "Cyan" }
        "Success" { "Green" }
        "Warning" { "Yellow" }
        "Error" { "Red" }
        "Step" { "Magenta" }
    }
    
    Write-Host "$emoji $Message" -ForegroundColor $color
}

function Test-QuotaValues {
    <#
    .SYNOPSIS
        Validates quota configuration values
    #>
    param($Config)
    
    if ($Config.WarningQuota -ge $Config.ProhibitSendQuota) {
        throw "Warning quota ($($Config.WarningQuota)) must be less than Prohibit Send quota ($($Config.ProhibitSendQuota))"
    }
    
    if ($Config.ProhibitSendQuota -ge $Config.ProhibitSendReceiveQuota) {
        throw "Prohibit Send quota ($($Config.ProhibitSendQuota)) must be less than Prohibit Send/Receive quota ($($Config.ProhibitSendReceiveQuota))"
    }
    
    Write-Status-Message "Quota configuration validated successfully" -Type "Success"
}

# ═══════════════════════════════════════════════════════════════════════════════
# MODULE & CONNECTION MANAGEMENT
# ═══════════════════════════════════════════════════════════════════════════════

function Initialize-Modules {
    <#
    .SYNOPSIS
        Initialize required PowerShell modules and verify cmdlet availability
    #>
    Write-Status-Message "Initializing required modules..." -Type "Step"
    
    foreach ($Module in $RequiredModules) {
        try {
            Write-Host "  📦 Checking $Module..." -ForegroundColor Gray
            
            # Check if module is installed
            $installedModule = Get-InstalledModule -Name $Module -ErrorAction SilentlyContinue
            if (-not $installedModule) {
                Write-Status-Message "Installing $Module..." -Type "Info"
                Install-Module -Name $Module -Scope CurrentUser -Force -AllowClobber
            }
            
            # Check if module is imported
            $importedModule = Get-Module -Name $Module
            if (-not $importedModule) {
                Write-Host "  🔄 Importing $Module..." -ForegroundColor Gray
                Import-Module -Name $Module -Force
            }
            
            # Verify specific version for ExchangeOnlineManagement
            if ($Module -eq "ExchangeOnlineManagement") {
                $moduleVersion = (Get-Module -Name $Module).Version
                if ($moduleVersion -lt [Version]"3.2.0") {
                    Write-Status-Message "Updating $Module to latest version..." -Type "Info"
                    Update-Module -Name $Module -Force
                    Remove-Module -Name $Module -Force
                    Import-Module -Name $Module -Force
                }
                Write-Host "  ✅ $Module version: $((Get-Module -Name $Module).Version)" -ForegroundColor Green
                
                # Verify Exchange cmdlets are available (they load dynamically after connection)
                Write-Host "  🔍 Checking Exchange cmdlets availability..." -ForegroundColor Gray
                $exchangeCmdlets = @('Get-AcceptedDomain', 'Get-Mailbox', 'Set-Mailbox', 'Enable-Mailbox')
                $missingCmdlets = @()
                
                foreach ($cmdlet in $exchangeCmdlets) {
                    if (-not (Get-Command $cmdlet -ErrorAction SilentlyContinue)) {
                        $missingCmdlets += $cmdlet
                    }
                }
                
                if ($missingCmdlets.Count -gt 0) {
                    Write-Host "  ⚠️ Exchange cmdlets not yet loaded (this is normal before connection)" -ForegroundColor Yellow
                    Write-Host "    Missing: $($missingCmdlets -join ', ')" -ForegroundColor Gray
                }
                else {
                    Write-Host "  ✅ Exchange cmdlets are available" -ForegroundColor Green
                }
            }
            
        }
        catch {
            Write-Status-Message "Failed to initialize $Module : $($_.Exception.Message)" -Type "Error"
            return $false
        }
    }
    
    Write-Status-Message "All required modules initialized successfully" -Type "Success"
    return $true
}

function Test-ExchangeConnection {
    <#
    .SYNOPSIS
        Verify Exchange Online connection and cmdlet availability
    #>
    Write-Status-Message "Verifying Exchange Online connection..." -Type "Step"
    
    try {
        # Step 1: Check if Exchange cmdlets are available
        Write-Host "  🔍 Checking Exchange cmdlets availability..." -ForegroundColor Gray
        $exchangeCmdlets = @('Get-AcceptedDomain', 'Get-Mailbox', 'Set-Mailbox', 'Enable-Mailbox')
        $missingCmdlets = @()
        
        foreach ($cmdlet in $exchangeCmdlets) {
            if (-not (Get-Command $cmdlet -ErrorAction SilentlyContinue)) {
                $missingCmdlets += $cmdlet
            }
        }
        
        if ($missingCmdlets.Count -gt 0) {
            Write-Host "  ⚠️ Exchange cmdlets not available: $($missingCmdlets -join ', ')" -ForegroundColor Yellow
            Write-Host "  🔄 Attempting to establish Exchange Online connection..." -ForegroundColor Cyan
            
            try {
                # Try to connect to Exchange Online (this should load the cmdlets)
                Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
                Write-Host "  ✅ Connected to Exchange Online successfully" -ForegroundColor Green
                
                # Re-check cmdlets after connection
                $stillMissing = @()
                foreach ($cmdlet in $exchangeCmdlets) {
                    if (-not (Get-Command $cmdlet -ErrorAction SilentlyContinue)) {
                        $stillMissing += $cmdlet
                    }
                }
                
                if ($stillMissing.Count -gt 0) {
                    Write-Status-Message "Exchange cmdlets still not available after connection attempt" -Type "Error"
                    return $false
                }
            }
            catch {
                Write-Status-Message "Failed to connect to Exchange Online: $($_.Exception.Message)" -Type "Error"
                Write-Host "  💡 Please ensure you have Exchange Administrator permissions" -ForegroundColor Yellow
                Write-Host "  🔄 Try using option 8 to reconnect to the tenant" -ForegroundColor Yellow
                return $false
            }
        }
        
        # Step 2: Test connection using multiple methods
        Write-Host "  🔍 Testing connection method 1: Get-ConnectionInformation..." -ForegroundColor Gray
        $connectionInfo = Get-ConnectionInformation -ErrorAction SilentlyContinue
        
        if ($connectionInfo -and $connectionInfo.State -eq "Connected") {
            Write-Status-Message "Exchange Online connection verified successfully" -Type "Success"
            Write-Host "  🌐 Connected to: $($connectionInfo.TenantId)" -ForegroundColor Gray
            Write-Host "  👤 Connected as: $($connectionInfo.UserPrincipalName)" -ForegroundColor Gray
            return $true
        }
        
        # Step 3: Test with Exchange cmdlets directly
        Write-Host "  🔍 Testing connection method 2: Exchange cmdlet test..." -ForegroundColor Gray
        
        $testDomains = Get-AcceptedDomain -ResultSize 1 -ErrorAction SilentlyContinue
        if ($testDomains) {
            Write-Status-Message "Exchange Online connection verified via cmdlet test" -Type "Success"
            Write-Host "  🌐 Exchange Online is accessible and responding" -ForegroundColor Gray
            return $true
        }
        
        # Step 4: Test mailbox access
        Write-Host "  🔍 Testing connection method 3: Mailbox access test..." -ForegroundColor Gray
        $testMailbox = Get-Mailbox -ResultSize 1 -ErrorAction SilentlyContinue
        if ($testMailbox) {
            Write-Status-Message "Exchange Online connection verified via mailbox access" -Type "Success"
            Write-Host "  📧 Mailbox data accessible" -ForegroundColor Gray
            return $true
        }
        
        # No connection detected
        Write-Status-Message "Unable to verify Exchange Online connection" -Type "Error"
        Write-Host "  💡 Please check your Exchange permissions and connection status" -ForegroundColor Yellow
        return $false
        
    }
    catch {
        Write-Status-Message "Failed to verify Exchange Online connection: $($_.Exception.Message)" -Type "Error"
        Write-Host "  📝 This might indicate missing Exchange permissions or authentication issues" -ForegroundColor Gray
        return $false
    }
}

# ═══════════════════════════════════════════════════════════════════════════════
# CORE MAILBOX CONFIGURATION FUNCTIONS
# ═══════════════════════════════════════════════════════════════════════════════

function Enable-MailboxArchiving {
    <#
    .SYNOPSIS
        Enable and verify archive mailboxes for all user mailboxes
    #>
    param([array]$Mailboxes)
    
    Write-Status-Message "Processing mailbox archiving configuration..." -Type "Step"
    
    $archiveStats = @{
        AlreadyEnabled = 0
        NewlyEnabled = 0
        Failed = 0
    }
    
    foreach ($mailbox in $Mailboxes) {
        $Global:ProcessedMailboxes++
        $percent = [math]::Round(($Global:ProcessedMailboxes / $Global:TotalMailboxes) * 100, 1)
        
        try {
            Write-Progress-Custom -Activity "Configuring Mailbox Archives" -Status "Processing $($mailbox.UserPrincipalName)" -PercentComplete $percent -CurrentOperation "Checking archive status..."
            
            # Check current archive status
            $currentMailbox = Get-Mailbox -Identity $mailbox.UserPrincipalName -ErrorAction Stop
            
            if ($currentMailbox.ArchiveStatus -eq 'Active') {
                Write-Host "  ✅ Archive already enabled for: $($mailbox.UserPrincipalName)" -ForegroundColor Green
                $archiveStats.AlreadyEnabled++
            }
            else {
                Write-Host "  🔄 Enabling archive for: $($mailbox.UserPrincipalName)" -ForegroundColor Yellow
                
                # Enable archive mailbox
                Enable-Mailbox -Identity $mailbox.UserPrincipalName -Archive -ErrorAction Stop
                
                # Verify enablement
                Start-Sleep -Seconds 2  # Brief pause for propagation
                $verifyMailbox = Get-Mailbox -Identity $mailbox.UserPrincipalName -ErrorAction Stop
                
                if ($verifyMailbox.ArchiveStatus -eq 'Active') {
                    Write-Host "  ✅ Archive successfully enabled for: $($mailbox.UserPrincipalName)" -ForegroundColor Green
                    $archiveStats.NewlyEnabled++
                }
                else {
                    throw "Archive enablement verification failed"
                }
            }
        }
        catch {
            $errorMsg = "Failed to process archive for $($mailbox.UserPrincipalName): $($_.Exception.Message)"
            Write-Status-Message $errorMsg -Type "Error"
            $Global:ProcessingErrors += $errorMsg
            $archiveStats.Failed++
        }
    }
    
    # Summary
    Write-Host "`n📊 Archive Processing Summary:" -ForegroundColor Cyan
    Write-Host "  ✅ Already enabled: $($archiveStats.AlreadyEnabled)" -ForegroundColor Green
    Write-Host "  🆕 Newly enabled: $($archiveStats.NewlyEnabled)" -ForegroundColor Green  
    Write-Host "  ❌ Failed: $($archiveStats.Failed)" -ForegroundColor Red
}

function Set-MailboxQuotas {
    <#
    .SYNOPSIS
        Configure storage quotas for all user mailboxes
    #>
    param(
        [array]$Mailboxes,
        [hashtable]$QuotaConfiguration
    )
    
    Write-Status-Message "Processing mailbox quota configuration..." -Type "Step"
    
    $quotaStats = @{
        Updated = 0
        AlreadyConfigured = 0
        Failed = 0
    }
    
    # Reset progress counter for quota processing
    $Global:ProcessedMailboxes = 0
    
    foreach ($mailbox in $Mailboxes) {
        $Global:ProcessedMailboxes++
        $percent = [math]::Round(($Global:ProcessedMailboxes / $Global:TotalMailboxes) * 100, 1)
        
        try {
            Write-Progress-Custom -Activity "Configuring Mailbox Quotas" -Status "Processing $($mailbox.UserPrincipalName)" -PercentComplete $percent -CurrentOperation "Setting quota limits..."
            
            # Get current quota settings
            $currentMailbox = Get-Mailbox -Identity $mailbox.UserPrincipalName -ErrorAction Stop
            
            # Check if quotas are already set correctly
            $needsUpdate = $false
            if ($currentMailbox.IssueWarningQuota -ne $QuotaConfiguration.WarningQuota) { $needsUpdate = $true }
            if ($currentMailbox.ProhibitSendQuota -ne $QuotaConfiguration.ProhibitSendQuota) { $needsUpdate = $true }
            if ($currentMailbox.ProhibitSendReceiveQuota -ne $QuotaConfiguration.ProhibitSendReceiveQuota) { $needsUpdate = $true }
            
            if (-not $needsUpdate) {
                Write-Host "  ✅ Quotas already configured for: $($mailbox.UserPrincipalName)" -ForegroundColor Green
                $quotaStats.AlreadyConfigured++
            }
            else {
                Write-Host "  🔄 Updating quotas for: $($mailbox.UserPrincipalName)" -ForegroundColor Yellow
                
                # Apply quota settings
                Set-Mailbox -Identity $mailbox.UserPrincipalName `
                    -IssueWarningQuota $QuotaConfiguration.WarningQuota `
                    -ProhibitSendQuota $QuotaConfiguration.ProhibitSendQuota `
                    -ProhibitSendReceiveQuota $QuotaConfiguration.ProhibitSendReceiveQuota `
                    -UseDatabaseQuotaDefaults $false `
                    -ErrorAction Stop
                
                Write-Host "  ✅ Quotas updated for: $($mailbox.UserPrincipalName)" -ForegroundColor Green
                Write-Host "    📏 Warning: $($QuotaConfiguration.WarningQuota), Send: $($QuotaConfiguration.ProhibitSendQuota), Send/Receive: $($QuotaConfiguration.ProhibitSendReceiveQuota)" -ForegroundColor Gray
                $quotaStats.Updated++
            }
        }
        catch {
            $errorMsg = "Failed to set quotas for $($mailbox.UserPrincipalName): $($_.Exception.Message)"
            Write-Status-Message $errorMsg -Type "Error"
            $Global:ProcessingErrors += $errorMsg
            $quotaStats.Failed++
        }
    }
    
    # Summary
    Write-Host "`n📊 Quota Processing Summary:" -ForegroundColor Cyan
    Write-Host "  🆕 Updated: $($quotaStats.Updated)" -ForegroundColor Green
    Write-Host "  ✅ Already configured: $($quotaStats.AlreadyConfigured)" -ForegroundColor Green
    Write-Host "  ❌ Failed: $($quotaStats.Failed)" -ForegroundColor Red
}

# ═══════════════════════════════════════════════════════════════════════════════
# MAIN EXECUTION FUNCTION
# ═══════════════════════════════════════════════════════════════════════════════

function Start-MailboxConfiguration {
    <#
    .SYNOPSIS
        Main orchestration function for complete mailbox configuration
    #>
    
    Write-Host "`n🚀 Exchange Online Mailbox Configuration Script" -ForegroundColor Magenta
    Write-Host "═══════════════════════════════════════════════════════════════" -ForegroundColor Magenta
    Write-Host "📧 Archive Enablement & Storage Quota Configuration" -ForegroundColor Cyan
    Write-Host "⚙️  Warning: 40GB | Prohibit Send: 45GB | Prohibit Send/Receive: 49GB" -ForegroundColor Gray
    Write-Host "═══════════════════════════════════════════════════════════════`n" -ForegroundColor Magenta
    
    try {
        # Step 1: Validate quota configuration
        Write-Status-Message "Validating quota configuration..." -Type "Step"
        Test-QuotaValues -Config $QuotaConfig
        
        # Step 2: Get all user mailboxes
        Write-Status-Message "Retrieving user mailboxes from Exchange Online..." -Type "Step"
        $allMailboxes = Get-Mailbox -RecipientTypeDetails UserMailbox -ResultSize Unlimited -ErrorAction Stop
        $Global:TotalMailboxes = $allMailboxes.Count
        
        if ($Global:TotalMailboxes -eq 0) {
            Write-Status-Message "No user mailboxes found in the tenant" -Type "Warning"
            return $false
        }
        
        Write-Status-Message "Found $Global:TotalMailboxes user mailboxes to process" -Type "Success"
        
        # User confirmation
        Write-Host "`n📋 Configuration Summary:" -ForegroundColor Yellow
        Write-Host "  📧 Mailboxes to process: $Global:TotalMailboxes" -ForegroundColor White
        Write-Host "  📦 Archive enablement: Yes" -ForegroundColor White
        Write-Host "  📏 Warning quota: $($QuotaConfig.WarningQuota)" -ForegroundColor White
        Write-Host "  🚫 Prohibit send quota: $($QuotaConfig.ProhibitSendQuota)" -ForegroundColor White
        Write-Host "  ⛔ Prohibit send/receive quota: $($QuotaConfig.ProhibitSendReceiveQuota)" -ForegroundColor White
        
        $confirmation = Read-Host "`n🤔 Do you want to proceed with mailbox configuration? (Y/N)"
        if ($confirmation -notin @('Y', 'y', 'Yes', 'yes')) {
            Write-Status-Message "Operation cancelled by user" -Type "Warning"
            return $false
        }
        
        # Step 3: Enable mailbox archiving
        Write-Host "`n" + "═" * 70 -ForegroundColor Cyan
        Write-Status-Message "PHASE 1: MAILBOX ARCHIVE ENABLEMENT" -Type "Step"
        Write-Host "═" * 70 -ForegroundColor Cyan
        
        Enable-MailboxArchiving -Mailboxes $allMailboxes
        
        # Step 4: Configure mailbox quotas
        Write-Host "`n" + "═" * 70 -ForegroundColor Cyan
        Write-Status-Message "PHASE 2: MAILBOX QUOTA CONFIGURATION" -Type "Step" 
        Write-Host "═" * 70 -ForegroundColor Cyan
        
        Set-MailboxQuotas -Mailboxes $allMailboxes -QuotaConfiguration $QuotaConfig
        
        # Step 5: Final summary and recommendations
        Write-Host "`n" + "═" * 70 -ForegroundColor Green
        Write-Status-Message "MAILBOX CONFIGURATION COMPLETED" -Type "Success"
        Write-Host "═" * 70 -ForegroundColor Green
        
        Write-Host "`n📊 Overall Processing Summary:" -ForegroundColor Cyan
        Write-Host "  📧 Total mailboxes processed: $Global:TotalMailboxes" -ForegroundColor White
        Write-Host "  ❌ Total errors encountered: $($Global:ProcessingErrors.Count)" -ForegroundColor Red
        
        if ($Global:ProcessingErrors.Count -gt 0) {
            Write-Host "`n📋 Error Details:" -ForegroundColor Red
            foreach ($errorMessage in $Global:ProcessingErrors) {
                Write-Host "  • $errorMessage" -ForegroundColor Red
            }
        }
        
        Write-Host "`n💡 Post-Configuration Recommendations:" -ForegroundColor Yellow
        Write-Host "  • Monitor mailbox usage with Get-MailboxStatistics" -ForegroundColor Gray
        Write-Host "  • Consider auto-expanding archives for high-usage mailboxes" -ForegroundColor Gray
        Write-Host "  • Review retention policies for optimal storage management" -ForegroundColor Gray
        Write-Host "  • Test archive accessibility in Outlook clients" -ForegroundColor Gray
        
        return $true
        
    }
    catch {
        Write-Status-Message "Critical error in mailbox configuration: $($_.Exception.Message)" -Type "Error"
        Write-Host "📝 Stack trace: $($_.ScriptStackTrace)" -ForegroundColor Gray
        return $false
    }
    finally {
        # Clean up progress bars
        Write-Progress -Activity "Mailbox Configuration" -Completed
    }
}

# ═══════════════════════════════════════════════════════════════════════════════
# SCRIPT EXECUTION ENTRY POINT
# ═══════════════════════════════════════════════════════════════════════════════

try {
    Write-Status-Message "Loading Exchange Mailbox Configuration Script..." -Type "Info"
    
    # Step 1: Initialize modules
    Write-Status-Message "Step 1: Initialize Modules" -Type "Step"
    if (-not (Initialize-Modules)) {
        Write-Status-Message "Module initialization failed - exiting" -Type "Error"
        Read-Host "Press Enter to continue"
        return
    }
    
    # Step 2: Test Exchange connection
    Write-Status-Message "Step 2: Verify Exchange Online Connection" -Type "Step"
    if (-not (Test-ExchangeConnection)) {
        Write-Status-Message "Exchange Online connection verification failed - exiting" -Type "Error"
        Read-Host "Press Enter to continue"
        return
    }
    
    # Step 3: Execute main configuration
    Write-Status-Message "Step 3: Start Mailbox Configuration Process" -Type "Step"
    $results = Start-MailboxConfiguration
    
    if ($results) {
        Write-Status-Message "Mailbox configuration process completed successfully!" -Type "Success"
    }
    else {
        Write-Status-Message "Mailbox configuration process did not complete successfully" -Type "Warning"
    }
}
catch {
    Write-Status-Message "Script execution failed: $($_.Exception.Message)" -Type "Error"
    Write-Host "📝 Error details: $($_.Exception.GetType().FullName)" -ForegroundColor Gray
    Write-Host "📍 Line number: $($_.InvocationInfo.ScriptLineNumber)" -ForegroundColor Gray
}
finally {
    Write-Host "`n⏸️ Press Enter to return to menu..." -ForegroundColor Gray
    Read-Host
}

# ▼ CB & Claude | BITS 365 Automation | v1.0 | "Smarter not Harder"