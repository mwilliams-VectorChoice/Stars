# Script: Citrix Workspace Automated Installer
# Purpose: Handles installation and updates of Citrix Workspace App
# Version: 2.3

#Requires -RunAsAdministrator

# Enable strict mode
Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# Global variables
$TempDir = "C:\Temp"
$InstallerPath = Join-Path $TempDir "CitrixWorkspaceApp.exe"
$LogPath = "C:\Temp\CitrixInstall_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
$CitrixVersions = @{
    "Current" = "https://downloadplugins.citrix.com/Windows/CitrixWorkspaceApp.exe"
    "LTSR" = "https://downloadplugins.citrix.com/Windows/CitrixWorkspaceAppLTSR.exe"
}

# Ensure Temp directory exists
if (-not (Test-Path $TempDir)) {
    New-Item -ItemType Directory -Path $TempDir -Force | Out-Null
}

function Write-LogMessage {
    param(
        [string]$Message,
        [ValidateSet('Info', 'Warning', 'Error')]
        [string]$Level = 'Info'
    )
    
    $TimeStamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $LogMessage = "[$TimeStamp] [$Level] $Message"
    Write-Host $LogMessage
    
    try {
        $LogMessage | Out-File -FilePath $LogPath -Append
    }
    catch {
        Write-Warning "Could not write to log file: $($_.Exception.Message)"
    }
}

function Wait-ProcessEnd {
    param(
        [string]$ProcessName,
        [int]$TimeoutSeconds = 60
    )
    
    Write-LogMessage "Waiting for process to end: $ProcessName"
    $timer = 0
    while ($timer -lt $TimeoutSeconds) {
        if (-not (Get-Process -Name $ProcessName -ErrorAction SilentlyContinue)) {
            Write-LogMessage "Process $ProcessName has ended"
            return $true
        }
        Start-Sleep -Seconds 1
        $timer++
    }
    Write-LogMessage "Timeout waiting for $ProcessName to end" -Level 'Warning'
    return $false
}

function Test-CitrixRunning {
    $citrixProcesses = @(
        "CDViewer",
        "Receiver",
        "AuthManSvr",
        "concentr",
        "wfcrun32",
        "CitrixWorkspaceApp",
        "redirector",
        "WebHelper",
        "SelfService",
        "HdxRtcEngine",
        "CitrixCGP"
    )
    
    foreach ($process in $citrixProcesses) {
        if (Get-Process -Name $process -ErrorAction SilentlyContinue) {
            return $true
        }
    }
    return $false
}

function Wait-CitrixProcessesEnd {
    param([int]$TimeoutSeconds = 120)
    
    Write-LogMessage "Waiting for all Citrix processes to end..."
    $timer = 0
    while ($timer -lt $TimeoutSeconds) {
        if (-not (Test-CitrixRunning)) {
            Write-LogMessage "All Citrix processes have ended"
            return $true
        }
        Start-Sleep -Seconds 2
        $timer += 2
    }
    Write-LogMessage "Timeout waiting for Citrix processes to end" -Level 'Warning'
    return $false
}

function Stop-CitrixProcesses {
    try {
        Write-LogMessage "Stopping Citrix processes..."
        $citrixProcesses = @(
            "CDViewer", "Receiver", "AuthManSvr", "concentr", "wfcrun32",
            "CitrixWorkspaceApp", "redirector", "WebHelper", "SelfService",
            "HdxRtcEngine", "CitrixCGP"
        )
        
        foreach ($process in $citrixProcesses) {
            $runningProcess = Get-Process -Name $process -ErrorAction SilentlyContinue
            if ($runningProcess) {
                Write-LogMessage "Stopping process: $process"
                $runningProcess | Stop-Process -Force
                Wait-ProcessEnd -ProcessName $process -TimeoutSeconds 30
            }
        }
        
        return Wait-CitrixProcessesEnd
    }
    catch {
        Write-LogMessage "Error stopping Citrix processes: $($_.Exception.Message)" -Level 'Error'
        return $false
    }
}

function Remove-CitrixRegistry {
    try {
        Write-LogMessage "Cleaning up Citrix registry entries..."
        $registryPaths = @(
            "HKLM:\SOFTWARE\Citrix",
            "HKLM:\SOFTWARE\WOW6432Node\Citrix",
            "HKLM:\SOFTWARE\Policies\Citrix",
            "HKCU:\Software\Citrix"
        )
        
        foreach ($path in $registryPaths) {
            if (Test-Path $path) {
                Write-LogMessage "Removing registry path: $path"
                Remove-Item -Path $path -Recurse -Force -ErrorAction SilentlyContinue
            }
        }
        Write-LogMessage "Registry cleanup completed"
        return $true
    }
    catch {
        Write-LogMessage "Error cleaning registry: $($_.Exception.Message)" -Level 'Warning'
        return $false
    }
}

function Get-CitrixInstalledVersion {
    try {
        # Primary check through file system first
        $citrixPath = "C:\Program Files (x86)\Citrix\ICA Client"
        if (Test-Path $citrixPath) {
            $wfcrun32 = Join-Path $citrixPath "wfcrun32.exe"
            if (Test-Path $wfcrun32) {
                $fileVersion = (Get-Item $wfcrun32).VersionInfo.FileVersion
                if ($fileVersion) {
                    return $fileVersion
                }
            }
        }

        # Only try WMI as fallback
        try {
            $wmiVersion = Get-WmiObject -Class Win32_Product | 
                Where-Object { $_.Name -like "*Citrix Workspace*" } |
                Select-Object -ExpandProperty Version
            if ($wmiVersion) { return $wmiVersion }
        }
        catch {
            Write-LogMessage "WMI check failed (this is normal in some environments): $($_.Exception.Message)" -Level 'Info'
        }
    }
    catch {
        Write-LogMessage "Error getting installed version: $($_.Exception.Message)" -Level 'Error'
    }
    return $null
}

function Wait-CitrixInstallation {
    param([int]$TimeoutSeconds = 300)
    
    Write-LogMessage "Waiting for Citrix installation to complete..."
    $timer = 0
    while ($timer -lt $TimeoutSeconds) {
        $version = Get-CitrixInstalledVersion
        if ($version) {
            Write-LogMessage "Citrix installation completed. Version: $version"
            return $true
        }
        Start-Sleep -Seconds 5
        $timer += 5
        Write-LogMessage "Still waiting for installation... ($timer seconds elapsed)"
    }
    Write-LogMessage "Timeout waiting for installation to complete" -Level 'Error'
    return $false
}

function Test-CitrixNeedsUpdate {
    param (
        [string]$DownloadedVersion,
        [string]$InstalledVersion
    )
    
    if (-not $InstalledVersion) { return $true }
    if (-not $DownloadedVersion) { return $false }
    
    try {
        return [version]$DownloadedVersion -gt [version]$InstalledVersion
    }
    catch {
        Write-LogMessage "Error comparing versions: $($_.Exception.Message)" -Level 'Error'
        return $false
    }
}

function Remove-CitrixWorkspace {
    try {
        Write-LogMessage "Beginning Citrix Workspace removal process..."
        
        # Stop all Citrix processes first
        if (-not (Stop-CitrixProcesses)) {
            Write-LogMessage "Warning: Could not stop all Citrix processes" -Level 'Warning'
        }
        
        # Uninstall using native uninstaller
        $CitrixPath = Get-ChildItem "C:\Program Files (x86)\Citrix\" -Filter "Citrix Workspace *" -ErrorAction SilentlyContinue |
            Select-Object -ExpandProperty FullName
        
        if ($CitrixPath) {
            Write-LogMessage "Uninstalling Citrix Workspace..."
            $UninstallerPath = Join-Path $CitrixPath "CWAInstaller.exe"
            if (Test-Path $UninstallerPath) {
                $process = Start-Process -FilePath $UninstallerPath -ArgumentList "/uninstall /silent" -Wait -PassThru
                
                if ($process.ExitCode -eq 0) {
                    Write-LogMessage "Uninstallation process completed"
                    
                    # Wait for uninstallation to complete
                    $timer = 0
                    while ($timer -lt 60 -and (Get-CitrixInstalledVersion)) {
                        Start-Sleep -Seconds 2
                        $timer += 2
                    }
                }
                else {
                    Write-LogMessage "Uninstallation process exited with code: $($process.ExitCode)" -Level 'Warning'
                }
            }
        }
        
        # Final cleanup
        Stop-CitrixProcesses  # Stop any remaining processes
        Remove-CitrixRegistry  # Clean registry
        
        # Verify removal
        if (-not (Get-CitrixInstalledVersion)) {
            Write-LogMessage "Citrix Workspace removal verified"
            return $true
        }
        else {
            Write-LogMessage "Warning: Citrix Workspace might still be installed" -Level 'Warning'
            return $false
        }
    }
    catch {
        Write-LogMessage "Error during uninstallation: $($_.Exception.Message)" -Level 'Error'
        return $false
    }
}

function Download-CitrixWorkspace {
    try {
        Write-LogMessage "Downloading latest Citrix Workspace App..."
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        
        $webClient = New-Object System.Net.WebClient
        $webClient.Headers.Add("User-Agent", "PowerShell Script")
        $webClient.DownloadFile($CitrixVersions["Current"], $InstallerPath)

        if (Test-Path $InstallerPath) {
            $fileVersion = (Get-Item $InstallerPath).VersionInfo.FileVersion
            Write-LogMessage "Download completed. Version: $fileVersion"
            return $fileVersion
        }
    }
    catch {
        Write-LogMessage "Error downloading Citrix Workspace: $($_.Exception.Message)" -Level 'Error'
    }
    return $null
}

function Install-CitrixWorkspace {
    try {
        Write-LogMessage "Installing Citrix Workspace..."
        $process = Start-Process -FilePath $InstallerPath -ArgumentList "/silent /noreboot" -Wait -PassThru
        
        if ($process.ExitCode -eq 0) {
            Write-LogMessage "Installation process started successfully"
            if (Wait-CitrixInstallation -TimeoutSeconds 300) {
                Write-LogMessage "Installation completed successfully"
                return $true
            }
        }
        else {
            Write-LogMessage "Installation failed with exit code: $($process.ExitCode)" -Level 'Error'
        }
    }
    catch {
        Write-LogMessage "Error during installation: $($_.Exception.Message)" -Level 'Error'
    }
    return $false
}

function Clean-TempFiles {
    try {
        if (Test-Path $InstallerPath) {
            Remove-Item $InstallerPath -Force
            Write-LogMessage "Temporary files cleaned up"
        }
    }
    catch {
        Write-LogMessage "Error cleaning temporary files: $($_.Exception.Message)" -Level 'Warning'
    }
}

# Main execution block
try {
    Write-LogMessage "Starting Citrix Workspace management script..."
    
    # Get current installed version
    $installedVersion = Get-CitrixInstalledVersion
    if ($installedVersion) {
        Write-LogMessage "Current installed version: $installedVersion"
    }
    else {
        Write-LogMessage "No Citrix Workspace installation found"
    }
    
    # Download latest version
    $downloadedVersion = Download-CitrixWorkspace
    if (-not $downloadedVersion) {
        throw "Failed to download new version"
    }
    
    # Check if update is needed
    if (Test-CitrixNeedsUpdate -DownloadedVersion $downloadedVersion -InstalledVersion $installedVersion) {
        Write-LogMessage "Update needed. Downloaded version ($downloadedVersion) is newer than installed version ($installedVersion)"
        
        # Remove existing installation if present
        if ($installedVersion) {
            if (-not (Remove-CitrixWorkspace)) {
                Write-LogMessage "Warning: Complete removal could not be verified" -Level 'Warning'
            }
        }
        
        # Install new version
        if (-not (Install-CitrixWorkspace)) {
            throw "Failed to install new version"
        }
        
        # Verify installation
        $newVersion = Get-CitrixInstalledVersion
        if ($newVersion -eq $downloadedVersion) {
            Write-LogMessage "Successfully updated to version $newVersion"
        }
        else {
            Write-LogMessage "Version mismatch after installation. Expected: $downloadedVersion, Got: $newVersion" -Level 'Warning'
        }
    }
    else {
        Write-LogMessage "No update needed. Current version is up to date."
    }
}
catch {
    Write-LogMessage "Critical error: $($_.Exception.Message)" -Level 'Error'
    exit 1
}
finally {
    # Cleanup
    Clean-TempFiles
}