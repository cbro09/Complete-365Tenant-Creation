#Requires -Version 7.0

<#
.SYNOPSIS
    M365 Tenant Automation Hub - Universal Main Menu
.DESCRIPTION
    Universal PowerShell 7 automation hub for Microsoft 365 tenant configuration.
    Downloads latest scripts from GitHub and provides centralized authentication.
.AUTHOR
    CB & Claude Partnership
.VERSION
    1.0
#>

# Global Variables
$Global:TenantConnection = $null
$Global:CurrentScopes = @()
$Global:GitHubRepo = "cbro09/Complete-365Tenant-Creation"
$Global:GitHubBranch = "main" # Change to "dev" for testing
$Global:ScriptCache = @{}

# Required Modules for Main Menu
$RequiredModules = @(
    'Microsoft.Graph.Authentication',
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

# Download script from GitHub
function Get-GitHubScript {
    param(
        [string]$ScriptPath,
        [string]$Branch = $Global:GitHubBranch
    )
    
    $url = "https://raw.githubusercontent.com/$Global:GitHubRepo/$Branch/$ScriptPath"
    
    try {
        $response = Invoke-RestMethod -Uri $url -ErrorAction Stop
        return $response
    }
    catch {
        Write-Error "Failed to download $ScriptPath from GitHub: $($_.Exception.Message)"
        return $null
    }
}

# Execute downloaded script
function Invoke-GitHubScript {
    param(
        [string]$ScriptPath,
        [hashtable]$Parameters = @{}
    )
    
    if ($Global:ScriptCache.ContainsKey($ScriptPath)) {
        $scriptContent = $Global:ScriptCache[$ScriptPath]
    } else {
        Write-Host "📥 Downloading $ScriptPath..." -ForegroundColor Yellow
        $scriptContent = Get-GitHubScript -ScriptPath $ScriptPath
        if ($scriptContent) {
            $Global:ScriptCache[$ScriptPath] = $scriptContent
        } else {
            return
        }
    }
    
    try {
        # Create script block and execute with parameters
        $scriptBlock = [ScriptBlock]::Create($scriptContent)
        & $scriptBlock @Parameters
    }
    catch {
        Write-Error "Error executing ${ScriptPath}: $($_.Exception.Message)"
    }
}

# Connect to Microsoft 365 Tenant
function Connect-M365Tenant {
    Write-Host "`n🔐 Connecting to Microsoft 365 Tenant..." -ForegroundColor Cyan
    
    try {
        # Basic connection for tenant info
        Connect-MgGraph -Scopes "Organization.Read.All" -NoWelcome
        
        $context = Get-MgContext
        $org = Get-MgOrganization | Select-Object -First 1
        
        $Global:TenantConnection = @{
            TenantId = $context.TenantId
            Account = $context.Account
            OrgName = $org.DisplayName
            ConnectedTime = Get-Date
        }
        
        Write-Host "✅ Connected to: $($org.DisplayName)" -ForegroundColor Green
        Write-Host "   Tenant ID: $($context.TenantId)" -ForegroundColor Gray
        Write-Host "   Account: $($context.Account)" -ForegroundColor Gray
        
        return $true
    }
    catch {
        Write-Error "Failed to connect: $($_.Exception.Message)"
        return $false
    }
}

# Set service-specific scopes and connect
function Set-ServiceScopes {
    param([string]$Service)
    
    $ServiceScopes = @{
        'Entra' = @(
            "User.ReadWrite.All",
            "Group.ReadWrite.All", 
            "Policy.ReadWrite.ConditionalAccess",
            "Directory.ReadWrite.All",
            "RoleManagement.ReadWrite.Directory"
        )
        'Intune' = @(
            "DeviceManagementConfiguration.ReadWrite.All",
            "DeviceManagementManagedDevices.ReadWrite.All",
            "Group.ReadWrite.All",
            "DeviceManagementApps.ReadWrite.All"
        )
        'Exchange' = @(
            "Group.ReadWrite.All",
            "Directory.ReadWrite.All"
        )
        'SharePoint' = @(
            "Sites.FullControl.All",
            "Group.ReadWrite.All"
        )
        'Security' = @(
            "SecurityActions.ReadWrite.All"
        )
        'Purview' = @(
            "InformationProtectionPolicy.Read.All",
            "RecordsManagement.ReadWrite.All"
        )
    }
    
    if ($ServiceScopes.ContainsKey($Service)) {
        $newScopes = $ServiceScopes[$Service]
        $currentContext = Get-MgContext
        
        # Check if we need to reconnect with different scopes
        if ($currentContext -and ($Global:CurrentScopes -ne $newScopes)) {
            Write-Host "🔄 Updating $Service permissions..." -ForegroundColor Yellow
            
            try {
                # Reconnect with new scopes using same account
                Disconnect-MgGraph -ErrorAction SilentlyContinue
                Connect-MgGraph -Scopes $newScopes -NoWelcome
                $Global:CurrentScopes = $newScopes
                Write-Host "✅ $Service permissions updated!" -ForegroundColor Green
                return $true
            }
            catch {
                Write-Error "Failed to set $Service scopes: $($_.Exception.Message)"
                return $false
            }
        }
        elseif (!$currentContext) {
            Write-Host "❌ Not connected to tenant. Please connect first." -ForegroundColor Red
            return $false
        }
        else {
            Write-Host "✅ $Service permissions already active!" -ForegroundColor Green
            return $true
        }
    }
    return $false
}

# Service menus
function Show-EntraMenu {
    if (!(Set-ServiceScopes -Service "Entra")) { return }
    
    do {
        Write-Host "`n" + "=" * 60 -ForegroundColor Cyan
        Write-Host "🏢 ENTRA ID AUTOMATION" -ForegroundColor Cyan
        Write-Host "=" * 60 -ForegroundColor Cyan
        Write-Host "1. Conditional Access Policies"
        Write-Host "2. Admin Account Creation"
        Write-Host "3. User Creation & Management"
        Write-Host "4. Security Groups (Dynamic)"
        Write-Host "5. Password Policies"
        Write-Host "0. Back to Main Menu"
        Write-Host ""
        
        $choice = Read-Host "Select option"
        
        switch ($choice) {
            "1" { Invoke-GitHubScript -ScriptPath "entra/CA-Policies.ps1" }
            "2" { Invoke-GitHubScript -ScriptPath "entra/Admin-Creation.ps1" }
            "3" { Invoke-GitHubScript -ScriptPath "entra/User-Creation.ps1" }
            "4" { Invoke-GitHubScript -ScriptPath "entra/Security-Groups.ps1" }
            "5" { Invoke-GitHubScript -ScriptPath "entra/Password-Policies.ps1" }
            "0" { break }
            default { Write-Host "Invalid option!" -ForegroundColor Red }
        }
    } while ($choice -ne "0")
}

function Show-IntuneMenu {
    if (!(Set-ServiceScopes -Service "Intune")) { return }
    
    do {
        Write-Host "`n" + "=" * 60 -ForegroundColor Magenta
        Write-Host "📱 INTUNE AUTOMATION" -ForegroundColor Magenta
        Write-Host "=" * 60 -ForegroundColor Magenta
        Write-Host "1. Configuration Policies"
        Write-Host "2. Compliance Policies"
        Write-Host "3. Device Groups (OS-based)"
        Write-Host "4. Application Deployment"
        Write-Host "5. Autopilot Configuration"
        Write-Host "0. Back to Main Menu"
        Write-Host ""
        
        $choice = Read-Host "Select option"
        
        switch ($choice) {
            "1" { Invoke-GitHubScript -ScriptPath "Intune/Configuration-Policies.ps1" }
            "2" { Invoke-GitHubScript -ScriptPath "Intune/Compliance-Policies.ps1" }
            "3" { Invoke-GitHubScript -ScriptPath "Intune/Device-Groups.ps1" }
            "4" { Invoke-GitHubScript -ScriptPath "Intune/App-Deployment.ps1" }
            "5" { Invoke-GitHubScript -ScriptPath "Intune/Autopilot-Config.ps1" }
            "0" { break }
            default { Write-Host "Invalid option!" -ForegroundColor Red }
        }
    } while ($choice -ne "0")
}

function Show-ExchangeMenu {
    if (!(Set-ServiceScopes -Service "Exchange")) { return }
    
    do {
        Write-Host "`n" + "=" * 60 -ForegroundColor Blue
        Write-Host "📧 EXCHANGE ONLINE AUTOMATION" -ForegroundColor Blue
        Write-Host "=" * 60 -ForegroundColor Blue
        Write-Host "1. Shared Mailbox Creation"
        Write-Host "2. Distribution Lists"
        Write-Host "3. Archive Policies"
        Write-Host "4. Mail Flow Rules"
        Write-Host "0. Back to Main Menu"
        Write-Host ""
        
        $choice = Read-Host "Select option"
        
        switch ($choice) {
            "1" { Invoke-GitHubScript -ScriptPath "Exchange/Shared-MB-Creation.ps1" }
            "2" { Invoke-GitHubScript -ScriptPath "Exchange/Distribution-Lists.ps1" }
            "3" { Invoke-GitHubScript -ScriptPath "Exchange/Archive-Policies.ps1" }
            "4" { Invoke-GitHubScript -ScriptPath "Exchange/Mail-Flow-Rules.ps1" }
            "0" { break }
            default { Write-Host "Invalid option!" -ForegroundColor Red }
        }
    } while ($choice -ne "0")
}

function Show-SharePointMenu {
    if (!(Set-ServiceScopes -Service "SharePoint")) { return }
    
    do {
        Write-Host "`n" + "=" * 60 -ForegroundColor Green
        Write-Host "🌐 SHAREPOINT ONLINE AUTOMATION" -ForegroundColor Green
        Write-Host "=" * 60 -ForegroundColor Green
        Write-Host "1. Site Collection Creation"
        Write-Host "2. Permission Groups"
        Write-Host "3. External Sharing Policies"
        Write-Host "0. Back to Main Menu"
        Write-Host ""
        
        $choice = Read-Host "Select option"
        
        switch ($choice) {
            "1" { Invoke-GitHubScript -ScriptPath "SharePoint/Site-Creation.ps1" }
            "2" { Invoke-GitHubScript -ScriptPath "SharePoint/Permission-Groups.ps1" }
            "3" { Invoke-GitHubScript -ScriptPath "SharePoint/External-Sharing.ps1" }
            "0" { break }
            default { Write-Host "Invalid option!" -ForegroundColor Red }
        }
    } while ($choice -ne "0")
}

function Show-SecurityMenu {
    if (!(Set-ServiceScopes -Service "Security")) { return }
    
    do {
        Write-Host "`n" + "=" * 60 -ForegroundColor Red
        Write-Host "🛡️ SECURITY & DEFENDER AUTOMATION" -ForegroundColor Red
        Write-Host "=" * 60 -ForegroundColor Red
        Write-Host "1. Web Content Filtering"
        Write-Host "2. Safe Attachments/Links"
        Write-Host "3. Anti-phishing Policies"
        Write-Host "0. Back to Main Menu"
        Write-Host ""
        
        $choice = Read-Host "Select option"
        
        switch ($choice) {
            "1" { Invoke-GitHubScript -ScriptPath "Security/Web-Filtering.ps1" }
            "2" { Invoke-GitHubScript -ScriptPath "Security/Safe-Attachments.ps1" }
            "3" { Invoke-GitHubScript -ScriptPath "Security/Anti-Phishing.ps1" }
            "0" { break }
            default { Write-Host "Invalid option!" -ForegroundColor Red }
        }
    } while ($choice -ne "0")
}

function Show-PurviewMenu {
    if (!(Set-ServiceScopes -Service "Purview")) { return }
    
    do {
        Write-Host "`n" + "=" * 60 -ForegroundColor DarkCyan
        Write-Host "🔒 PURVIEW COMPLIANCE AUTOMATION" -ForegroundColor DarkCyan
        Write-Host "=" * 60 -ForegroundColor DarkCyan
        Write-Host "1. Retention Policies"
        Write-Host "2. Data Loss Prevention"
        Write-Host "3. Sensitivity Labels"
        Write-Host "0. Back to Main Menu"
        Write-Host ""
        
        $choice = Read-Host "Select option"
        
        switch ($choice) {
            "1" { Invoke-GitHubScript -ScriptPath "Purview/Retention-Policies.ps1" }
            "2" { Invoke-GitHubScript -ScriptPath "Purview/DLP-Policies.ps1" }
            "3" { Invoke-GitHubScript -ScriptPath "Purview/Sensitivity-Labels.ps1" }
            "0" { break }
            default { Write-Host "Invalid option!" -ForegroundColor Red }
        }
    } while ($choice -ne "0")
}

# Refresh script cache
function Clear-ScriptCache {
    $Global:ScriptCache.Clear()
    Write-Host "🔄 Script cache cleared!" -ForegroundColor Green
}

# Main Menu
function Show-MainMenu {
    Clear-Host
    Write-Host "=" * 80 -ForegroundColor Cyan
    Write-Host "🚀 M365 TENANT AUTOMATION HUB" -ForegroundColor Cyan
    Write-Host "=" * 80 -ForegroundColor Cyan
    
    if ($Global:TenantConnection) {
        Write-Host "✅ Connected to: $($Global:TenantConnection.OrgName)" -ForegroundColor Green
        Write-Host "   Account: $($Global:TenantConnection.Account)" -ForegroundColor Gray
    } else {
        Write-Host "❌ Not connected to tenant" -ForegroundColor Red
    }
    
    Write-Host "`nPortal Automation:" -ForegroundColor Yellow
    Write-Host "1. 🏢 Entra ID (Identity & Access)"
    Write-Host "2. 📱 Intune (Device Management)"
    Write-Host "3. 📧 Exchange Online (Email)"
    Write-Host "4. 🌐 SharePoint Online (Collaboration)"
    Write-Host "5. 🛡️ Security & Defender"
    Write-Host "6. 🔒 Purview (Compliance)"
    Write-Host ""
    Write-Host "Options:" -ForegroundColor Yellow
    Write-Host "8. 🔐 Connect to Tenant"
    Write-Host "9. 🔄 Refresh Scripts"
    Write-Host "0. ❌ Exit"
    Write-Host ""
}

# Main execution loop
function Start-AutomationHub {
    Initialize-Modules
    
    do {
        Show-MainMenu
        $choice = Read-Host "Select option"
        
        switch ($choice) {
            "1" { 
                if ($Global:TenantConnection) { Show-EntraMenu } 
                else { Write-Host "Please connect to tenant first!" -ForegroundColor Red; Start-Sleep 2 }
            }
            "2" { 
                if ($Global:TenantConnection) { Show-IntuneMenu } 
                else { Write-Host "Please connect to tenant first!" -ForegroundColor Red; Start-Sleep 2 }
            }
            "3" { 
                if ($Global:TenantConnection) { Show-ExchangeMenu } 
                else { Write-Host "Please connect to tenant first!" -ForegroundColor Red; Start-Sleep 2 }
            }
            "4" { 
                if ($Global:TenantConnection) { Show-SharePointMenu } 
                else { Write-Host "Please connect to tenant first!" -ForegroundColor Red; Start-Sleep 2 }
            }
            "5" { 
                if ($Global:TenantConnection) { Show-SecurityMenu } 
                else { Write-Host "Please connect to tenant first!" -ForegroundColor Red; Start-Sleep 2 }
            }
            "6" { 
                if ($Global:TenantConnection) { Show-PurviewMenu } 
                else { Write-Host "Please connect to tenant first!" -ForegroundColor Red; Start-Sleep 2 }
            }
            "8" { Connect-M365Tenant }
            "9" { Clear-ScriptCache }
            "0" { 
                Write-Host "Goodbye! 👋" -ForegroundColor Cyan
                if ($Global:TenantConnection) { Disconnect-MgGraph -ErrorAction SilentlyContinue }
                break 
            }
            default { Write-Host "Invalid option!" -ForegroundColor Red; Start-Sleep 1 }
        }
    } while ($choice -ne "0")
}

# Start the automation hub
Start-AutomationHub

# ▼ CB & Claude | BITS 365 Automation | v1.0 | "Smarter not Harder"