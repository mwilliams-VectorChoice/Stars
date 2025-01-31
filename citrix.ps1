# Script: Citrix Workspace Automated Installer
# Purpose: Handles installation and updates of Citrix Workspace App
# Version: 3.1

#Requires -RunAsAdministrator

# Enable strict mode
Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# Global variables
$TempDir = "C:\Temp"
$InstallerPath = Join-Path $TempDir "CitrixWorkspaceApp.exe"
$LogPath = "C:\Temp\CitrixInstall_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
$InstallStartTime = $null
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

function Get-CitrixInstallationStatus {
    param (
        [Parameter(Mandatory=$true)]
        [DateTime]$InstallStartTime
    )
    try {
        Write-LogMessage "Checking installation status in Event Logs since $($InstallStartTime.ToString('HH:mm:ss'))..."
        
        # Look for MSI events only after our install started
        $events = Get-WinEvent -FilterHashtable @{
            LogName = 'Application'
            ProviderName = 'MsiInstaller'
            ID = @(1033, 1042)
            StartTime = $InstallStartTime
        } -ErrorAction SilentlyContinue

        if ($events) {
            # Sort events by time and check for sequence
            $events = $events | Sort-Object TimeCreated
            
            # Look specifically for our Citrix install events
            $installComplete = $events | Where-Object { 
                $_.Id -eq 1033 -and 
                $_.TimeCreated -gt $InstallStartTime -and 
                $_.Message -like "*Citrix*"
            } | Select-Object -Last 1

            $installEnd = $events | Where-Object { 
                $_.Id -eq 1042 -and 
                $_.TimeCreated -gt $InstallStartTime -and 
                $_.TimeCreated -gt $installComplete.TimeCreated -and 
                $_.Message -like "*Citrix*"
            } | Select-Object -Last 1

            if ($installComplete -and $installEnd) {
                Write-LogMessage "Found Citrix installation events after start time:"
                Write-LogMessage "- Install Complete (1033) at: $($installComplete.TimeCreated.ToString('HH:mm:ss'))"
                Write-LogMessage "- Install End (1042) at: $($installEnd.TimeCreated.ToString('HH:mm:ss'))"
                return $true
            }
        }

        Write-LogMessage "No matching Citrix installation events found since $($InstallStartTime.ToString('HH:mm:ss'))"
        return $false
    }
    catch {
        Write-LogMessage "Error checking installation events: $($_.Exception.Message)" -Level 'Warning'
        return $false
    }
}

function Stop-CitrixProcesses {
    try {
        Write-LogMessage "Stopping Citrix processes..."
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
            $runningProcess = Get-Process -Name $process -ErrorAction SilentlyContinue
            if ($runningProcess) {
                Write-LogMessage "Stopping process: $process"
                $runningProcess | Stop-Process -Force
            }
        }
        
        # Verify all processes are stopped
        Start-Sleep -Seconds 2  # Brief pause to let processes stop
        $remainingProcesses = Get-Process | Where-Object { $_.Name -in $citrixProcesses }
        if (-not $remainingProcesses) {
            Write-LogMessage "All Citrix processes stopped successfully"
            return $true
        }
        else {
            Write-LogMessage "Some Citrix processes could not be stopped" -Level 'Warning'
            return $false
        }
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
        # Check file system first
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

        # Try WMI as fallback
        try {
            $wmiVersion = Get-WmiObject -Class Win32_Product -ErrorAction SilentlyContinue | 
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
                }
                else {
                    Write-LogMessage "Uninstallation process exited with code: $($process.ExitCode)" -Level 'Warning'
                }
            }
        }
        
        # Clean registry
        Remove-CitrixRegistry

        # Final verification
        $remainingVersion = Get-CitrixInstalledVersion
        if (-not $remainingVersion) {
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
        $InstallStartTime = Get-Date
        Write-LogMessage "Starting installation at $($InstallStartTime.ToString('HH:mm:ss'))"

        # Start the installation
        $process = Start-Process -FilePath $InstallerPath -ArgumentList "/silent /noreboot" -PassThru

        $installSuccess = $false

        # Monitor the event log for up to 120 seconds
        $timer = [Diagnostics.Stopwatch]::StartNew()

        while ($timer.Elapsed.TotalSeconds -lt 120) {
            Start-Sleep -Seconds 10
            $installSuccess = Get-CitrixInstallationStatus -InstallStartTime $InstallStartTime

            if ($installSuccess) {
                Write-LogMessage "Installation completed successfully based on event logs."
                break
            }
        }

        $timer.Stop()

        if (-not $installSuccess) {
            Write-LogMessage "Installation did not complete within 120 seconds based on event logs." -Level 'Warning'
        }

        return $installSuccess
    }
    catch {
        Write-LogMessage "Error occurred: $($_.Exception.Message)" -Level 'Error'
        return $false
    }
}

function Clean-TempFiles {
    param (
        [string]$TempDir = "C:\Temp"
    )
    try {
        Write-LogMessage "Cleaning up temporary files in $TempDir..."
        # Ensure the file is not in use before attempting to delete
        Start-Sleep -Seconds 20
        $retryCount = 3
        $retryDelay = 10
        for ($i = 0; $i -lt $retryCount; $i++) {
            try {
                if (Test-Path $InstallerPath) {
                    Write-LogMessage "Attempting to remove installer file $InstallerPath (attempt $($i + 1) of $retryCount)..."
                    Remove-Item -Path $InstallerPath -Force
                    Write-LogMessage "Installer file $InstallerPath removed successfully."
                    break
                }
            }
            catch {
                Write-LogMessage "Failed to remove installer file ${InstallerPath}: $($_.Exception.Message)" -Level 'Warning'
                Start-Sleep -Seconds $retryDelay
            }
        }
        Remove-Item -Path "$TempDir\*" -Recurse -Force
        Write-LogMessage "Temporary files cleaned up successfully."
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