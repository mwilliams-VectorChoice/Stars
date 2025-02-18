﻿# Script: User Management, Bloat Remover, Basic Setup, YellowDog, TeamViewer, Citrix Install

param (
    [switch]$Debug,
    [string]$Config,
    [switch]$Run
)

# Define the script version
$scriptVersion = "11.0.0"

# Function to compare version numbers properly
function Compare-Versions {
    param (
        [string]$Version1,
        [string]$Version2
    )
    
    try {
        $v1 = [version]$Version1
        $v2 = [version]$Version2
        return $v1.CompareTo($v2)
    }
    catch {
        Write-Warning "Error comparing versions: $_"
        return 0
    }
}

# Function to get latest version from GitHub
function Get-LatestGitHubVersion {
    try {
        $response = Invoke-WebRequest -Uri "https://raw.githubusercontent.com/mwilliams-VectorChoice/Stars/main/version.txt" -ErrorAction Stop
        return $response.Content.Trim()
    }
    catch {
        Write-Warning "Failed to get latest version from GitHub: $_"
        return $null
    }
}

# Version check logic
function Check-ScriptVersion {
    param (
        [string]$CurrentVersion,
        [string]$ScriptVersion
    )
    
    try {
        $latestVersion = Get-LatestGitHubVersion
        
        if ($null -eq $latestVersion) {
            Write-Host "Unable to fetch latest version. Using local version."
            return $false
        }
        
        Write-Host "Current Version: $CurrentVersion"
        Write-Host "Script Version: $ScriptVersion"
        Write-Host "Latest GitHub Version: $latestVersion"
        
        # First compare local script version with GitHub version
        $compareGitHub = Compare-Versions -Version1 $latestVersion -Version2 $ScriptVersion
        
        if ($compareGitHub -gt 0) {
            Write-Host "A newer version ($latestVersion) is available. Updating..."
            return $true
        }
        
        Write-Host "Using version: $ScriptVersion"
        return $false
    }
    catch {
        Write-Warning "Error during version check: $_"
        return $false
    }
}

##################################################################################################################
#                                             TABLE OF CONTENTS                                                  #
#                                                                                                                #
#                                        AUTO UPDATE CODE - Lines 33 - 101                                       #
#                                        PROGRAM MENU - Lines 154 - 166                                          #
#                                        USER PROFILES AND FOLDERS - Lines 175 - 327                             #
#                                        INSTALL TEAMVIWER - Lines 336 - 392                                     #
#                                        INSTALL CHROME - Lines 401 - 437                                        
#                                        INSTALL ADOBE READER - Lines 446 - 524                                  #
#                                        INSTALL MICROSOFT TEAMS - Lines 532 - 588                               #
#                                        INSTALL YELLOWDOG FOR STARS AND STRIKES - Lines 597 - 757               #
#                                        PIN ICONS TO TASKBAR - Lines 767 - 825                                  #
#                                        BAISC COMPUTER SETUP - Lines 833 - 999                                  #
#                                        INSTALL CITRIX - WELLCARE - Lines 1008 - 1309                           #
#                                        REMOVE BLOAT  - Lines 1319 - 3084                                       #
#                                        REMOVE WINDOWS.OLD  - Lines 3092 - 3110                                 #
#                                        MaAFEE - Lines 3120 - 3350                                              #
#                                        MAIN PROGRAM LOOP - Lines 3360 - 3467                                   # 
#                                                                                                                #
#                                                                                                                #
#                                                                                                                #
#                                                                                                                #
#                                                                                                                #
##################################################################################################################

#################################################################################################################
####                                                                                                          ###
#### WARNING: This file is automatically generated DO NOT modify this file directly as it will be overwritten ###
####                                                                                                          ###
#################################################################################################################


# Add Windows API support for window manipulation
Add-Type @"
    using System;
    using System.Runtime.InteropServices;
    public class Win32 {
        [DllImport("user32.dll")]
        public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
        
        [DllImport("user32.dll")]
        public static extern bool SetForegroundWindow(IntPtr hWnd);
        
        [DllImport("user32.dll")]
        public static extern IntPtr GetForegroundWindow();
    }
"@

# Function to maximize window
function Maximize-Window {
    try {
        $processes = Get-Process | Where-Object { $_.MainWindowTitle -ne "" }
        $window = $processes | Where-Object { $_.Id -eq $pid } | Select-Object -First 1
        
        if ($window) {
            # 3 is maximize
            [void][Win32]::ShowWindow($window.MainWindowHandle, 3)
            [void][Win32]::SetForegroundWindow($window.MainWindowHandle)
        }
        
        # Also set buffer and window size
        $maxWidth = $host.UI.RawUI.MaxWindowSize.Width
        $maxHeight = $host.UI.RawUI.MaxWindowSize.Height
        
        $bufferSize = New-Object System.Management.Automation.Host.Size($maxWidth, 32766)
        $windowSize = New-Object System.Management.Automation.Host.Size($maxWidth, $maxHeight)
        
        $host.UI.RawUI.BufferSize = $bufferSize
        $host.UI.RawUI.WindowSize = $windowSize
    } catch {
        Write-Host "Note: Unable to maximize window automatically. Please maximize manually if needed."
    }
}

# Call the maximize function
Maximize-Window

# Initialize the warning color variable
$script:warningColor = 'Red'
$script:warningColorIndex = 0
$script:warningColors = @('Red', 'Yellow')


# Set DebugPreference based on the -Debug switch
if ($Debug) {
    $DebugPreference = "Continue"
}

if ($Config) {
    $PARAM_CONFIG = $Config
}

$PARAM_RUN = $false
# Handle the -Run switch
if ($Run) {
    Write-Host "Running config file tasks..."
    $PARAM_RUN = $true
}

# Function to clean up temporary files
function Cleanup-TempFiles {
    param ($localArchivePath, $localScriptPath)
    try {
        if (Test-Path $localArchivePath) { Remove-Item $localArchivePath -Force }
        if (Test-Path $localScriptPath) { Remove-Item $localScriptPath -Force }
        Write-Host "Temporary files cleaned up successfully."
    } catch {
        Write-Error "Failed to clean up temporary files: $_"
    }
}

# Ensure cleanup happens on script exit
Register-EngineEvent -SourceIdentifier PowerShell.Exiting -Action {
    Cleanup-TempFiles "$env:TEMP\devicemanagement.zip" "$env:TEMP\DeviceManagement.ps1"
}

# Check for and terminate existing jobs that are not in a final state
Get-Job | ForEach-Object {
    if ($_.State -eq 'Running' -or $_.State -eq 'NotStarted') {
        Write-Host "Terminating existing job: $($_.Id)"
        Stop-Job -Job $_ -ErrorAction SilentlyContinue
        Wait-Job -Job $_ -ErrorAction SilentlyContinue
    }
    Remove-Job -Job $_ -Force -ErrorAction SilentlyContinue

}

# Check for administrative privileges
if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Output "DeviceManagement needs to be run as Administrator. Attempting to relaunch."
    $argList = @()

    $PSBoundParameters.GetEnumerator() | ForEach-Object {
        $argList += if ($_.Value -is [switch] -and $_.Value) {
            "-$($_.Key)"
        } elseif ($_.Value) {
            "-$($_.Key) `"$($_.Value)`""
        }
    }

    $script = if ($MyInvocation.MyCommand.Path) {
        "& { & '$($MyInvocation.MyCommand.Path)' $argList }"
    } else {
        "iex '& { $(irm https://github.com/mwilliams-VectorChoice/Stars/releases/latest/download/DeviceManagement.ps1) } $argList'"
    }

    $powershellcmd = if (Get-Command pwsh -ErrorAction SilentlyContinue) { "pwsh" } else { "powershell" }
    $processCmd = if (Get-Command wt.exe -ErrorAction SilentlyContinue) { "wt.exe" } else { $powershellcmd }

    Start-Process $processCmd -ArgumentList "$powershellcmd -ExecutionPolicy Bypass -NoProfile -WindowStyle Maximized -Command $script" -Verb RunAs

    break
}

if (Check-ScriptVersion -CurrentVersion $global:CurrentScriptVersion -ScriptVersion $scriptVersion) {
    # Download and extract the latest version
    try {
        $latestArchiveUrl = "https://github.com/mwilliams-VectorChoice/Stars/releases/latest/download/devicemanagement.zip"
        $localArchivePath = "$env:TEMP\devicemanagement.zip"
        $localScriptPath = "$env:TEMP\DeviceManagement.ps1"
        
        Write-Host "Downloading latest version..."
        Invoke-WebRequest -Uri $latestArchiveUrl -OutFile $localArchivePath
        
        Write-Host "Extracting script..."
        Expand-Archive -Path $localArchivePath -DestinationPath $env:TEMP -Force
        
        if (Test-Path $localScriptPath) {
            Write-Host "Running updated version..."
            . $localScriptPath
            Cleanup-TempFiles $localArchivePath $localScriptPath
            exit
        }
    }
    catch {
        Write-Warning "Update failed: $_"
        Write-Host "Continuing with current version..."
    }
}

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

#region Error Handling and Post Actions
function PostActions {
    # Refresh environment variables
    $env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path","User")
    
    # Restart Explorer to apply changes
    Stop-Process -Name explorer -Force
    Start-Process explorer
    
    Write-Host "Post-actions completed. Some changes may require a system restart." -ForegroundColor Green
}

function Errors {
    if ($Error.Count -gt 0) {
        Write-Host "`nErrors encountered during script execution:" -ForegroundColor Red
        foreach ($err in $Error) {
            Write-Host "- $($err.Exception.Message)" -ForegroundColor Red
        }
    }
}


############################################################################################################
#                                                   PROGRAM MENU                                           #
#                                                                                                          #
############################################################################################################

# Function to Show Flashing Warning Without New Line
function Show-FlashingWarning {
    param (
        [string]$Text,
        [int]$Delay = 500
    )

    while ($true) {
        Write-Host "`r$Text" -ForegroundColor Red -NoNewline
        Start-Sleep -Milliseconds $Delay
        Write-Host "`r$Text" -ForegroundColor Yellow -NoNewline
        Start-Sleep -Milliseconds $Delay
    }
}

# Function to Display Main Menu
function Show-MainMenu {
    Maximize-Window 
    Clear-Host  # Clear screen before displaying menu
    Write-Host ""
    Write-Host "          _______  _______ _________ _______  _______   " -ForegroundColor DarkBlue 
    Write-Host "|\     /|(  ____ \(  ____ \\__   __/(  ___  )(  ____ )  " -ForegroundColor Blue
    Write-Host "| )   ( || (    \/| (    \/   ) (   | (   ) || (    )|  " -ForegroundColor Cyan
    Write-Host "| |   | || (__    | |         | |   | |   | || (____)|  " -ForegroundColor Yellow
    Write-Host "( (   ) )|  __)   | |         | |   | |   | ||     __)  " -ForegroundColor DarkCyan
    Write-Host " \ \_/ / | (      | |         | |   | |   | || (\ (     " -ForegroundColor DarkGreen
    Write-Host "  \   /  | (____/\| (____/\   | |   | (___) || ) \ \__  " -ForegroundColor Red
    Write-Host "   \_/   (_______/(_______/   )_(   (_______)|/   \__/  " -ForegroundColor Magenta
    Write-Host ""                                                        
    Write-Host " _______           _______ _________ _______  _______   " -ForegroundColor DarkBlue
    Write-Host "(  ____ \|\     /|(  ___  )\__   __/(  ____ \(  ____ \  " -ForegroundColor Blue
    Write-Host "| (    \/| )   ( || (   ) |   ) (   | (    \/| (    \/  " -ForegroundColor Cyan
    Write-Host "| |      | (___) || |   | |   | |   | |      | (__      " -ForegroundColor Yellow
    Write-Host "| |      |  ___  || |   | |   | |   | |      |  __)     " -ForegroundColor DarkCyan
    Write-Host "| |      | (   ) || |   | |   | |   | |      | (        " -ForegroundColor DarkGreen
    Write-Host "| (____/\| )   ( || (___) |___) (___| (____/\| (____/\  " -ForegroundColor Red
    Write-Host "(_______/|/     \|(_______)\_______/(_______/(_______/  " -ForegroundColor Magenta                                                  
    Write-Host ""
    Write-Host "╔══════════════════════════════════════════╗" -ForegroundColor Magenta
    Write-Host "║   🚀 Running Version: $scriptVersion 🚀 ║" -ForegroundColor Green
    Write-Host "╚══════════════════════════════════════════╝" -ForegroundColor Magenta
    Write-Host "================ Main Menu ================" -ForegroundColor Cyan
    Write-Host "1. System Information" -ForegroundColor Cyan
    Write-Host "2. Full Remove Bloat" -ForegroundColor Cyan
    Write-Host "3. Basic Computer Setup (Including Printer Options)" -ForegroundColor Cyan
    Write-Host "4. Stars and Strikes (YellowDog and TeamViewer)" -ForegroundColor Cyan
    Write-Host "5. Citrix Installation" -ForegroundColor Cyan
    Write-Host "6. Local User Profile Management" -ForegroundColor Cyan
    Write-Host "Q. Quit Program" -ForegroundColor Red
    Write-Host "=========================================" -ForegroundColor Cyan
    Write-Host "IMPORTANT: Press Q when finished to properly exit and cleanup downloaded update files" -ForegroundColor Red
}

# Display the menu FIRST
Show-MainMenu

# Start Flashing Warning in Background Job
$flashingJob = Start-Job -ScriptBlock {
    param ($message)
    function Show-FlashingWarning {
        param (
            [string]$Text,
            [int]$Delay = 500
        )
        while ($true) {
            Write-Host "`r$Text" -ForegroundColor Red -NoNewline
            Start-Sleep -Milliseconds $Delay
            Write-Host "`r$Text" -ForegroundColor Yellow -NoNewline
            Start-Sleep -Milliseconds $Delay
        }
    }
    Show-FlashingWarning -Text $message
} -ArgumentList "IMPORTANT: Press 'Q' when finished to properly exit and cleanup downloaded update files"


############################################################################################################
#                                        USER PROFILES AND FOLDERS                                         #
#                                                                                                          #
############################################################################################################


function Show-ProfileMenu {
    Write-Host "`n------ Profile Management Options ------" -ForegroundColor Yellow
    Write-Host "1. Remove user completely (account and profile)" -ForegroundColor Cyan
    Write-Host "2. Remove only profile" -ForegroundColor Cyan
    Write-Host "3. Remove only the C:\Users\ folder" -ForegroundColor Cyan
    Write-Host "Q. Return to Main Menu" -ForegroundColor Red
    Write-Host "-------------------------------------" -ForegroundColor Yellow
}


# Device Management Functions
function Get-UserList {
    try {
        $userFolders = Get-ChildItem "C:\Users" -Directory | Where-Object { $_.Name -notin @('Public', 'Default', 'Default User', 'All Users') }
        $userList = @()
        $index = 1
        
        foreach ($folder in $userFolders) {
            $isLocalUser = $false
            $isEnabled = $false
            try {
                $localUser = Get-LocalUser -Name $folder.Name -ErrorAction SilentlyContinue
                if ($localUser) {
                    $isLocalUser = $true
                    $isEnabled = $localUser.Enabled
                }
            } catch {
                # Not a local user, probably domain/Azure AD user
            }

            $userInfo = [PSCustomObject]@{
                Index = $index
                Name = $folder.Name
                ProfileExists = $true
                IsLocalUser = $isLocalUser
                Enabled = $isEnabled
                Path = $folder.FullName
            }
            $userList += $userInfo
            
            Write-Host "$index. $($folder.Name)" -NoNewline -ForegroundColor Green
            Write-Host " - Path: $($folder.FullName)" -NoNewline -ForegroundColor Yellow
            if ($isLocalUser) {
                Write-Host " (Local User, Enabled: $isEnabled)" -ForegroundColor Cyan
            } else {
                Write-Host " (Domain/Azure AD User)" -ForegroundColor Magenta
            }
            $index++
        }
        return $userList
    }
    catch {
        Write-Host "Error getting user list: $_" -ForegroundColor Red
        Read-Host "Press any key to continue..."
        return $null
    }
}

function Remove-UserProfile {
    param (
        [string]$Username,
        [string]$RemovalType
    )
    
    try {
        # Create backup
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $backupPath = "C:\UserBackups\$Username`_$timestamp"
        
        if (!(Test-Path "C:\UserBackups")) {
            New-Item -ItemType Directory -Path "C:\UserBackups" -Force | Out-Null
        }

        Write-Host "Creating backup at $backupPath..." -ForegroundColor Yellow
        if (Test-Path "C:\Users\$Username") {
            Copy-Item "C:\Users\$Username" -Destination $backupPath -Recurse -Force
        }

        switch ($RemovalType) {
            '1' {
                Write-Host "Removing user account $Username..." -ForegroundColor Yellow
                Remove-LocalUser -Name $Username -Confirm:$false -ErrorAction Stop
                Write-Host "Removing user profile..." -ForegroundColor Yellow
                if (Test-Path "C:\Users\$Username") {
                    Remove-Item -Path "C:\Users\$Username" -Recurse -Force
                }
            }
            '2' {
                if (Test-Path "C:\Users\$Username\AppData\Local") {
                    Write-Host "Removing profile for $Username..." -ForegroundColor Yellow
                    Remove-Item -Path "C:\Users\$Username\AppData\Local" -Recurse -Force
                }
            }
            '3' {
                if (Test-Path "C:\Users\$Username") {
                    Write-Host "Removing C:\Users\$Username folder..." -ForegroundColor Yellow
                    Remove-Item -Path "C:\Users\$Username" -Recurse -Force
                }
            }
        }
        Write-Host "Operation completed successfully for $Username" -ForegroundColor Green
        Write-Host "Backup available at: $backupPath" -ForegroundColor Green
    }
    catch {
        Write-Host "Error during removal process for $Username`: $_" -ForegroundColor Red
    }
    Read-Host "Press any key to continue..."
}

function Handle-ProfileManagement {
    do {
        Show-ProfileMenu
        $profileChoice = Read-Host "Enter your choice"
        
        if ($profileChoice -match '^[123]$') {
            $userList = Get-UserList
            if ($null -eq $userList) { continue }
            
            Write-Host "`nEnter indices of users (comma-separated) or 'Q' to return" -ForegroundColor Cyan
            $userIndices = Read-Host "Selection"
            
            if ($userIndices -eq 'Q') { 
                return
            }
            
            $indices = $userIndices -split ',' | ForEach-Object { $_.Trim() }
            
            foreach ($index in $indices) {
                if ($index -match '^\d+$' -and [int]$index -ge 1 -and [int]$index -le $userList.Count) {
                    $user = $userList[[int]$index - 1]
                    
                    $confirm = Read-Host "Are you sure you want to process user $($user.Name)? (Y/N)"
                    if ($confirm -eq 'Y') {
                        Remove-UserProfile -Username $user.Name -RemovalType $profileChoice
                    }
                }
                else {
                    Write-Host "Invalid index: $index" -ForegroundColor Red
                }
            }
        }
        elseif ($profileChoice -eq 'Q') {
            return
        }
        else {
            Write-Host "Invalid choice. Please select a valid option." -ForegroundColor Red
        }
        
        Write-Host "`nPress any key to continue..."
        $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
        
    } while ($true)
}


############################################################################################################
#                                        INSTALL TEAMVIWER                                                 #
#                                                                                                          #
############################################################################################################


function Download-TeamViewerHost {
    try {
        $teamViewerPath = @(
            "${env:ProgramFiles}\TeamViewer\TeamViewer.exe",
            "${env:ProgramFiles(x86)}\TeamViewer\TeamViewer.exe"
        ) | Where-Object { Test-Path $_ } | Select-Object -First 1

        if ($teamViewerPath) {
            Write-Host "TeamViewer is already installed at: $teamViewerPath" -ForegroundColor Green
            Write-Host "Press any key to continue..."
            $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
            return $true
        }

        Write-Host "TeamViewer Host is not installed. Would you like to install it? (Y/N)" -ForegroundColor Yellow
        $installChoice = Read-Host
        
        if ($installChoice.ToUpper() -ne 'Y') {
            Write-Host "TeamViewer Host installation cancelled." -ForegroundColor Yellow
            return $false
        }

        $tempDir = "C:\temp"
        if (-not (Test-Path $tempDir)) {
            New-Item -ItemType Directory -Path $tempDir -Force | Out-Null
        }

        $installerPath = "C:\temp\TeamViewer_Host.exe"
        $url = "https://dl.teamviewer.com/download/TeamViewer_Host_Setup.exe"

        Write-Host "Downloading TeamViewer Host installer..." -ForegroundColor Yellow
        
        $webClient = New-Object System.Net.WebClient
        $webClient.DownloadFile($url, $installerPath)
        
        Write-Host "Installing TeamViewer Host..." -ForegroundColor Yellow
        
        $process = Start-Process -FilePath $installerPath -ArgumentList "/S" -Wait -PassThru
        
        if ($process.ExitCode -eq 0) {
            Write-Host "TeamViewer Host installed successfully!" -ForegroundColor Green
            return $true
        } else {
            Write-Host "TeamViewer Host installation failed with exit code: $($process.ExitCode)" -ForegroundColor Red
            return $false
        }
    }
    catch {
        Write-Host "Error during TeamViewer Host installation: $_" -ForegroundColor Red
        return $false
    }
    finally {
        if (($installerPath) -and (Test-Path $installerPath)) {
            Remove-Item $installerPath -Force
        }
    }
}


############################################################################################################
#                                        INSTALL CHROME                                                    #
#                                                                                                          #
############################################################################################################


function Download-Chrome {
    try {
        $tempDir = "C:\temp"
        if (-not (Test-Path $tempDir)) {
            New-Item -ItemType Directory -Path $tempDir -Force | Out-Null
        }

        $installerPath = "C:\temp\ChromeSetup.exe"
        $url = "https://dl.google.com/chrome/install/latest/chrome_installer.exe"

        Write-Host "Downloading Chrome installer..." -ForegroundColor Yellow
        
        $webClient = New-Object System.Net.WebClient
        $webClient.DownloadFile($url, $installerPath)
        
        Write-Host "Installing Chrome..." -ForegroundColor Yellow
        
        $process = Start-Process -FilePath $installerPath -ArgumentList "/silent", "/install" -Wait -PassThru
        
        if ($process.ExitCode -eq 0) {
            Write-Host "Chrome installed successfully!" -ForegroundColor Green
            return $true
        } else {
            Write-Host "Chrome installation failed with exit code: $($process.ExitCode)" -ForegroundColor Red
            return $false
        }
    }
    catch {
        Write-Host "Error during Chrome installation: An unexpected error occurred" -ForegroundColor Red
        return $false
    }
    finally {
        if (Test-Path $installerPath) {
            Remove-Item $installerPath -Force
        }
    }
}


############################################################################################################
#                                        INSTALL ADOBE READER                                              #
#                                                                                                          #
############################################################################################################


function Download-AdobeReader {
    try {
        $tempDir = "C:\temp"
        if (-not (Test-Path $tempDir)) {
            New-Item -ItemType Directory -Path $tempDir -Force | Out-Null
        }

        $installerPath = "C:\temp\AdobeReaderDC.exe"
        $url = "https://ardownload2.adobe.com/pub/adobe/reader/win/AcrobatDC/2300320269/AcroRdrDC2300320269_en_US.exe"

        Write-Host "Downloading Adobe Reader DC installer..." -ForegroundColor Yellow
        
        $webClient = New-Object System.Net.WebClient
        $webClient.DownloadFile($url, $installerPath)
        
        Write-Host "Installing Adobe Reader DC..." -ForegroundColor Yellow
        
        $process = Start-Process -FilePath $installerPath -ArgumentList "/sAll /rs /msi /norestart /quiet" -Wait -PassThru
        
        if ($process.ExitCode -eq 0) {
            Write-Host "Adobe Reader installed successfully!" -ForegroundColor Green
            
            Start-Sleep -Seconds 5
            
            Write-Host "Setting Adobe Reader as default PDF viewer..." -ForegroundColor Yellow
            
            try {
                $assocPath = "HKLM:\SOFTWARE\Classes\.pdf"
                if (!(Test-Path $assocPath)) {
                    New-Item -Path $assocPath -Force | Out-Null
                }
                Set-ItemProperty -Path $assocPath -Name "(Default)" -Value "AcroExch.Document.DC" -Force

                $contentPath = "HKLM:\SOFTWARE\Classes\AcroExch.Document.DC"
                if (!(Test-Path $contentPath)) {
                    New-Item -Path $contentPath -Force | Out-Null
                }
                Set-ItemProperty -Path $contentPath -Name "(Default)" -Value "Adobe Acrobat Document DC" -Force

                $tempXml = "C:\temp\AdobeReaderDefault.xml"
                @"
<?xml version="1.0" encoding="UTF-8"?>
<DefaultAssociations>
  <Association Identifier=".pdf" ProgId="AcroExch.Document.DC" ApplicationName="Adobe Acrobat Reader DC" />
</DefaultAssociations>
"@ | Out-File -FilePath $tempXml -Encoding UTF8

                $dismResult = Start-Process "DISM.exe" -ArgumentList "/Online /Import-DefaultAppAssociations:`"$tempXml`"" -Wait -PassThru -WindowStyle Hidden

                if ($dismResult.ExitCode -eq 0) {
                    Write-Host "Adobe Reader set as default PDF viewer successfully!" -ForegroundColor Green
                } else {
                    Write-Host "Warning: Could not set Adobe Reader as default PDF viewer" -ForegroundColor Yellow
                }

                if (Test-Path $tempXml) {
                    Remove-Item $tempXml -Force
                }

            } catch {
                Write-Host "Warning: Could not set Adobe Reader as default PDF viewer" -ForegroundColor Yellow
            }

            return $true
        } else {
            Write-Host "Adobe Reader installation failed with exit code: $($process.ExitCode)" -ForegroundColor Red
            return $false
        }
    }
    catch {
        Write-Host "Error during Adobe Reader installation: An unexpected error occurred" -ForegroundColor Red
        return $false
    }
    finally {
        if (Test-Path $installerPath) {
            Remove-Item $installerPath -Force
        }
    }
}

############################################################################################################
#                                        INSTALL MICROSOFT TEAMS                                           #
#                                                                                                          #
############################################################################################################

function Download-MicrosoftTeams {
    try {
        Write-Host "Checking for Microsoft Teams installation..." -ForegroundColor Yellow
        
        # Check if Teams is installed for current user
        $teamsPath = Get-AppxPackage -Name "MSTeams" -ErrorAction SilentlyContinue
        
        if ($teamsPath) {
            Write-Host "Microsoft Teams is already installed!" -ForegroundColor Green
            return $true
        }
        
        Write-Host "Microsoft Teams is not installed." -ForegroundColor Yellow
        Write-Host "Would you like to install Microsoft Teams? (Y/N)" -ForegroundColor Cyan
        $installChoice = Read-Host
        
        if ($installChoice.ToUpper() -eq 'Y') {
            Write-Host "Installing Microsoft Teams..." -ForegroundColor Yellow
            
            try {
                # Set the variable before using it
                $currentAppName = "Microsoft.Teams"
                
                # Install Teams using winget
                $process = Start-Process "winget" -ArgumentList "install", "--id", $currentAppName -Wait -PassThru -NoNewWindow
                
                if ($process.ExitCode -eq 0) {
                    Write-Host "Microsoft Teams installed successfully!" -ForegroundColor Green
                    return $true
                } else {
                    Write-Host "Failed to install Microsoft Teams using winget. Attempting Microsoft Store..." -ForegroundColor Yellow
                    
                    # Fallback to Microsoft Store
                    Start-Process "ms-windows-store://pdp/?productid=XP8BT4HCVPZQ"
                    
                    Write-Host "Microsoft Store has been opened. Please complete the installation there." -ForegroundColor Yellow
                    Write-Host "Press any key after installation is complete..." -ForegroundColor Cyan
                    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
                    
                    # Verify installation
                    $teamsPath = Get-AppxPackage -Name "MSTeams" -ErrorAction SilentlyContinue
                    if ($teamsPath) {
                        Write-Host "Microsoft Teams installation verified!" -ForegroundColor Green
                        return $true
                    } else {
                        Write-Host "Could not verify Microsoft Teams installation. Please try again." -ForegroundColor Red
                        return $false
                    }
                }
            }
            catch {
                Write-Host "Error during Teams installation: $_" -ForegroundColor Red
                return $false
            }
        } else {
            Write-Host "Skipping Microsoft Teams installation" -ForegroundColor Yellow
            return $false
        }
    }
    catch {
        Write-Host "Error checking/installing Teams: $_" -ForegroundColor Red
        return $false
    }
}

############################################################################################################
#                                        INSTALL YELLOWDOG FOR STARS AND STRIKES                           #
#                                                                                                          #
############################################################################################################


function Deploy-StarsFiles {
    $isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    
    if (-not $isAdmin) {
        Write-Host "This function requires administrative privileges. Restarting with elevated permissions..." -ForegroundColor Yellow
        try {
            Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
            return
        }
        catch {
            Write-Host "Failed to restart with administrative privileges. Please run the script as administrator." -ForegroundColor Red
            return
        }
    }

    $arcadePath = "C:\Program Files (x86)\YDI_Arcade"
    $mainPath = "C:\Program Files (x86)\YDI_StarsnStrikes"
    
    if ((Test-Path $arcadePath) -and (Test-Path $mainPath)) {
        Write-Host "YellowDog software is already installed:" -ForegroundColor Green
        Write-Host "Arcade found at: $arcadePath" -ForegroundColor Green
        Write-Host "Main found at: $mainPath" -ForegroundColor Green
        Write-Host "Press any key to continue..."
        $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
        return
    }

    Write-Host "YellowDog software not found or incomplete. Would you like to install/repair it? (Y/N)" -ForegroundColor Yellow
    $installChoice = Read-Host
    
    if ($installChoice.ToUpper() -ne 'Y') {
        Write-Host "YellowDog installation cancelled." -ForegroundColor Yellow
        return
    }

    try {
        $tempDir = "C:\temp"
        if (-not (Test-Path $tempDir)) {
            New-Item -ItemType Directory -Path $tempDir -Force | Out-Null
        }

        Write-Host "Downloading files from GitHub..." -ForegroundColor Yellow
        
        $arcadeUrl = "https://raw.githubusercontent.com/mwilliams-VectorChoice/Stars/main/Arcade.zip"
        $mainUrl = "https://raw.githubusercontent.com/mwilliams-VectorChoice/Stars/main/Main.zip"
        $arcadePath = "C:\temp\Arcade.zip"
        $mainPath = "C:\temp\Main.zip"

        try {
            $webClient = New-Object System.Net.WebClient
            [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor [System.Net.SecurityProtocolType]::Tls12
            
            Write-Host "Downloading Arcade.zip..." -ForegroundColor Yellow
            $webClient.DownloadFile($arcadeUrl, $arcadePath)
            
            Write-Host "Downloading Main.zip..." -ForegroundColor Yellow
            $webClient.DownloadFile($mainUrl, $mainPath)
        }
        catch {
            Write-Host "Error downloading files: $_" -ForegroundColor Red
            Write-Host "Attempting alternative download method..." -ForegroundColor Yellow
            
            try {
                Invoke-WebRequest -Uri $arcadeUrl -OutFile $arcadePath -UseBasicParsing
                Invoke-WebRequest -Uri $mainUrl -OutFile $mainPath -UseBasicParsing
            }
            catch {
                Write-Host "Both download methods failed. Please check your internet connection and try again." -ForegroundColor Red
                return
            }
        }

        if (-not (Test-Path $arcadePath) -or -not (Test-Path $mainPath)) {
            Write-Host "Failed to download one or both zip files." -ForegroundColor Red
            return
        }

        Write-Host "Extracting files..." -ForegroundColor Yellow
        try {
            $extractPath = "C:\temp\Stars and Strikes"
            if (-not (Test-Path $extractPath)) {
                New-Item -ItemType Directory -Path $extractPath -Force | Out-Null
            }
            
            Expand-Archive -Path $arcadePath -DestinationPath "$extractPath\Arcade" -Force
            Expand-Archive -Path $mainPath -DestinationPath "$extractPath\Main" -Force
        }
        catch {
            Write-Host "Error extracting zip files: $_" -ForegroundColor Red
            return
        }

        Write-Host "Deploying Arcade files..." -ForegroundColor Yellow
        try {
            $arcadeShortcutPath = "$extractPath\Arcade\Arcade\Public Desktop\YDInv_Arcade Inventory.lnk"
            $arcadeProgramPath = "$extractPath\Arcade\Arcade\YDI_Arcade"

            if (Test-Path $arcadeShortcutPath) {
                $acl = Get-Acl "C:\Users\Public\Desktop"
                $identity = "BUILTIN\Administrators"
                $fileSystemRights = "FullControl"
                $type = "Allow"
                $fileSystemAccessRule = New-Object System.Security.AccessControl.FileSystemAccessRule($identity, $fileSystemRights, $type)
                $acl.AddAccessRule($fileSystemAccessRule)
                Set-Acl "C:\Users\Public\Desktop" $acl

                Copy-Item $arcadeShortcutPath -Destination "C:\Users\Public\Desktop" -Force
                Write-Host "Arcade shortcut deployed successfully" -ForegroundColor Green
            } else {
                Write-Host "Arcade shortcut file not found at: $arcadeShortcutPath" -ForegroundColor Red
            }

            if (Test-Path $arcadeProgramPath) {
                Copy-Item $arcadeProgramPath -Destination "C:\Program Files (x86)" -Recurse -Force
                Write-Host "Arcade program files deployed successfully" -ForegroundColor Green
            } else {
                Write-Host "Arcade program files not found at: $arcadeProgramPath" -ForegroundColor Red
            }
        }
        catch {
            Write-Host "Error deploying Arcade files: $_" -ForegroundColor Red
        }

        Write-Host "Deploying Main files..." -ForegroundColor Yellow
        try {
            $mainShortcutPath = "$extractPath\Main\Main\Public Desktop\YDInv_S&S Inventory.lnk"
            $mainProgramPath = "$extractPath\Main\Main\YDI_StarsnStrikes"

            if (Test-Path $mainShortcutPath) {
                Copy-Item $mainShortcutPath -Destination "C:\Users\Public\Desktop" -Force
                Write-Host "Main shortcut deployed successfully" -ForegroundColor Green
            } else {
                Write-Host "Main shortcut file not found at: $mainShortcutPath" -ForegroundColor Red
            }

            if (Test-Path $mainProgramPath) {
                Copy-Item $mainProgramPath -Destination "C:\Program Files (x86)" -Recurse -Force
                Write-Host "Main program files deployed successfully" -ForegroundColor Green
            } else {
                Write-Host "Main program files not found at: $mainProgramPath" -ForegroundColor Red
            }
        }
        catch {
            Write-Host "Error deploying Main files: $_" -ForegroundColor Red
        }

        Write-Host "Deployment process completed!" -ForegroundColor Green
    }
    catch {
        Write-Host "Error in deployment process: $_" -ForegroundColor Red
    }
    finally {
        Write-Host "Cleaning up temporary files..." -ForegroundColor Yellow
        if (Test-Path $arcadePath) { Remove-Item $arcadePath -Force }
        if (Test-Path $mainPath) { Remove-Item $mainPath -Force }
        if (Test-Path $extractPath) { Remove-Item $extractPath -Recurse -Force }
    }

    Write-Host "Press any key to continue..."
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
}


############################################################################################################
#                                        PIN ICONS TO TASKBAR                                              #
#                                                                                                          #
############################################################################################################



function Pin-ToTaskbar {
    param (
        [Parameter(Mandatory=$true)]
        [string]$ApplicationPath
    )

    try {
        # Create the XML template with the application
        $taskbarLayout = @"
<?xml version="1.0" encoding="utf-8"?>
<LayoutModificationTemplate
    xmlns="http://schemas.microsoft.com/Start/2014/LayoutModification"
    xmlns:defaultlayout="http://schemas.microsoft.com/Start/2014/FullDefaultLayout"
    xmlns:start="http://schemas.microsoft.com/Start/2014/StartLayout"
    xmlns:taskbar="http://schemas.microsoft.com/Start/2014/TaskbarLayout"
    Version="1">
  <CustomTaskbarLayoutCollection PinListPlacement="Replace">
    <defaultlayout:TaskbarLayout>
      <taskbar:TaskbarPinList>
        <taskbar:DesktopApp DesktopApplicationLinkPath="$ApplicationPath" />
      </taskbar:TaskbarPinList>
    </defaultlayout:TaskbarLayout>
 </CustomTaskbarLayoutCollection>
</LayoutModificationTemplate>
"@

        # Create provisioning directory if it doesn't exist
        $provisioningPath = "$env:ProgramData\provisioning"
        if (!(Test-Path $provisioningPath)) {
            New-Item -Path $provisioningPath -ItemType Directory -Force | Out-Null
        }

        # Save the layout file
        $layoutFile = Join-Path $provisioningPath "taskbar_layout.xml"
        $taskbarLayout | Out-File $layoutFile -Encoding utf8 -Force

        # Configure registry settings
        $explorerKeyPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Explorer"
        if (!(Test-Path $explorerKeyPath)) {
            New-Item -Path $explorerKeyPath -Force | Out-Null
        }

        # Set registry values
        Set-ItemProperty -Path $explorerKeyPath -Name "StartLayoutFile" -Value $layoutFile -Type ExpandString
        Set-ItemProperty -Path $explorerKeyPath -Name "LockedStartLayout" -Value 1 -Type DWord

        # Force a taskbar refresh
        Stop-Process -Name explorer -Force -ErrorAction SilentlyContinue
        Start-Process explorer

        Write-Host "Successfully pinned $ApplicationPath to taskbar" -ForegroundColor Green
        return $true
    }
    catch {
        $errorMessage = $_.Exception.Message
        Write-Host "Failed to pin $ApplicationPath`: $errorMessage" -ForegroundColor Red
        return $false
    }
}

############################################################################################################
#                                        BAISC COMPUTER SETUP                                              #
#                                                                                                          #
############################################################################################################

function Install-PrinterDrivers {
    param (
        [string]$Client
    )
    
    $driversPath = "C:\Temp\PrinterDrivers"
    
    # Create drivers directory if it doesn't exist
    if (-not (Test-Path $driversPath)) {
        try {
            New-Item -ItemType Directory -Path $driversPath -Force -ErrorAction Stop
            Write-Host "Created printer drivers directory at $driversPath" -ForegroundColor Green
        }
        catch {
            Write-Host "Failed to create drivers directory: $($_.Exception.Message)" -ForegroundColor Red
            return $false
        }
    }

    function Download-AndInstall {
        param (
            [string]$Url,
            [string]$OutputFile,
            [string]$Arguments = ""
        )
        try {
            Write-Host "Downloading printer driver..." -ForegroundColor Yellow
            Invoke-WebRequest -Uri $Url -OutFile $OutputFile
            
            if (Test-Path $OutputFile) {
                Write-Host "Installing printer driver..." -ForegroundColor Yellow
                if ($OutputFile -like "*.zip") {
                    # Create a unique extraction directory for TMT
                    $extractPath = "$driversPath\TMT_extracted"
                    if (Test-Path $extractPath) {
                        Remove-Item -Path $extractPath -Recurse -Force
                    }
                    
                    Write-Host "Extracting TMT driver files..." -ForegroundColor Yellow
                    Expand-Archive -Path $OutputFile -DestinationPath $extractPath -Force
                    
                    # Wait for extraction and verify
                    $maxAttempts = 6  # Maximum number of attempts (60 seconds total)
                    $attempts = 0
                    $setupFound = $false
                    
                    do {
                        Start-Sleep -Seconds 10
                        Write-Host "Verifying extracted files... Attempt $($attempts + 1) of $maxAttempts" -ForegroundColor Yellow
                        
                        if (Test-Path $extractPath) {
                            # Look specifically for the setup executable
                            $installer = Get-ChildItem -Path $extractPath -Recurse -Filter "TMT_Generic_Plus_PS3_v3.11_Set-up.exe" | Select-Object -First 1
                            if ($installer) {
                                $setupFound = $true
                                break
                            }
                        }
                        
                        $attempts++
                    } while ($attempts -lt $maxAttempts -and -not $setupFound)
                    
                    if ($setupFound) {
                        Write-Host "Found installer at: $($installer.FullName)" -ForegroundColor Green
                        Write-Host "Launching installer..." -ForegroundColor Yellow
                        $process = Start-Process -FilePath $installer.FullName -ArgumentList "/install", "/quiet" -Wait -PassThru
                        
                        if ($process.ExitCode -eq 0) {
                            Write-Host "Installation completed successfully" -ForegroundColor Green
                        } else {
                            Write-Host "Installation completed with exit code: $($process.ExitCode)" -ForegroundColor Yellow
                        }
                    } else {
                        Write-Host "Could not find setup.exe in extracted files" -ForegroundColor Red
                        Write-Host "Available files in extraction directory:" -ForegroundColor Yellow
                        Get-ChildItem -Path $extractPath -Recurse | ForEach-Object {
                            Write-Host $_.FullName
                        }
                    }
                } else {
                    Start-Process -FilePath $OutputFile -ArgumentList $Arguments -Wait
                }
                return $true
            }
            return $false
        }
        catch {
            Write-Host "Error during download/installation: $($_.Exception.Message)" -ForegroundColor Red
            return $false
        }
    }

    function Show-ManualDownloadInstructions {
        param (
            [string]$ManufacturerUrl,
            [string]$ModelName,
            [string]$Instructions
        )
        Write-Host "`nPlease follow these steps for manual download:" -ForegroundColor Yellow
        Write-Host "1. Opening manufacturer's download page in your default browser..." -ForegroundColor Cyan
        Write-Host "2. $Instructions" -ForegroundColor Cyan
        Write-Host "`nPress any key to open the download page..."
        $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
        Start-Process $ManufacturerUrl
    }

    switch ($Client) {
        "TMT" {
            $url = "https://github.com/mwilliams-VectorChoice/Stars/raw/main/TMT_Generic_Plus_PS3_v3.11_Set-up.zip"
            $outputFile = "$driversPath\TMT_driver.zip"
            Download-AndInstall -Url $url -OutputFile $outputFile
        }
        
        "Dogwood" {
            Show-ManualDownloadInstructions `
                -ManufacturerUrl "https://www.kyoceradocumentsolutions.com.au/supportcentre/Pages/downloads.aspx?model=TASKalfa+5054ci" `
                -ModelName "Kyocera TASKalfa 5054ci" `
                -Instructions "Select your Windows version and download the PCL or PostScript driver."
        }
        
        "ICM" {
            Show-ManualDownloadInstructions `
                -ManufacturerUrl "https://www.kyoceradocumentsolutions.com.au/supportcentre/Pages/downloads.aspx?model=TASKalfa+5054ci" `
                -ModelName "Kyocera Document Solutions Printer" `
                -Instructions "Select your Windows version and download the PCL or PostScript driver."
        }
        
        "QuestGroup" {
            $url = "https://github.com/mwilliams-VectorChoice/Stars/raw/main/Quest_SH_D33_PCL6_PS_2412a_EnglishUS_64bit.exe"
            $outputFile = "$driversPath\Quest_driver.exe"
            Download-AndInstall -Url $url -OutputFile $outputFile -Arguments "/silent"
        }
        
        "S&S" {
            $url = "https://github.com/mwilliams-VectorChoice/Stars/raw/main/S%26S_XeroxSmartStart_2.1.22.0.exe"
            $outputFile = "$driversPath\SS_driver.exe"
            Download-AndInstall -Url $url -OutputFile $outputFile -Arguments "/quiet"
        }
        
        default {
            Write-Host "Invalid client selection" -ForegroundColor Red
            return $false
        }
    }

    Write-Host "`nWould you like to:" -ForegroundColor Yellow
    Write-Host "1. Set up another printer" -ForegroundColor Cyan
    Write-Host "2. Return to main menu" -ForegroundColor Cyan
    
    $choice = Read-Host "Enter your choice (1-2)"
    return $choice
}

function Show-PrinterMenu {
    Write-Host "`n------ Printer Setup ------" -ForegroundColor Yellow
    Write-Host "1. TMT (Canon ImageRunner Advance C5535i)" -ForegroundColor Cyan
    Write-Host "2. Dogwood (Kyocera TASKalfa 5054ci)" -ForegroundColor Cyan
    Write-Host "3. ICM (Kyocera)" -ForegroundColor Cyan
    Write-Host "4. QuestGroup (Sharp BP-70C45)" -ForegroundColor Cyan
    Write-Host "5. S&S (Xerox AltaLink C8030 MFP)" -ForegroundColor Cyan
    Write-Host "Q. Return to previous menu" -ForegroundColor Red
    Write-Host "--------------------------" -ForegroundColor Yellow
}

function Handle-PrinterSetup {
    do {
        Show-PrinterMenu
        $printerChoice = Read-Host "Select client for printer setup"
        
        switch ($printerChoice) {
            "1" { $result = Install-PrinterDrivers -Client "TMT" }
            "2" { $result = Install-PrinterDrivers -Client "Dogwood" }
            "3" { $result = Install-PrinterDrivers -Client "ICM" }
            "4" { $result = Install-PrinterDrivers -Client "QuestGroup" }
            "5" { $result = Install-PrinterDrivers -Client "S&S" }
            "Q" { return }
            default {
                Write-Host "Invalid choice. Please select a valid option." -ForegroundColor Red
                Start-Sleep -Seconds 2
                continue
            }
        }
        
        if ($printerChoice -ne "Q" -and $result -eq "2") {
            return
        }
        
        Write-Host "Press any key to continue..."
        $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
    } while ($true)
}


function Basic-Setup {
       
    $setupDone = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Setup" -Name "BasicSetupComplete" -ErrorAction SilentlyContinue

    if ($null -eq $setupDone) {
        # Enable Num Lock at startup
        Write-Host "Enabling Num Lock at startup..." -ForegroundColor Yellow
        Set-ItemProperty -Path "HKCU:\Control Panel\Keyboard" -Name "InitialKeyboardIndicators" -Value 2

        # Enable Network Discovery and File Sharing for workgroup networks
        Write-Host "Enabling Network Discovery and File Sharing..." -ForegroundColor Yellow
        netsh advfirewall firewall set rule group="Network Discovery" new enable=Yes
        netsh advfirewall firewall set rule group="File and Printer Sharing" new enable=Yes

        # Show This PC icon on desktop
        Write-Host "Adding This PC icon to desktop..." -ForegroundColor Yellow
        $ThisPCRegPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons\NewStartPanel"
        If (!(Test-Path $ThisPCRegPath)) {
            New-Item -Path $ThisPCRegPath -Force | Out-Null
        }
        Set-ItemProperty -Path $ThisPCRegPath -Name "{20D04FE0-3AEA-1069-A2D8-08002B30309D}" -Value 0

        # Disable Windows 260 character path limit
        Write-Host "Removing Windows 260 character path limit..." -ForegroundColor Yellow
        Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\FileSystem" -Name "LongPathsEnabled" -Value 1

        # Install Visual C++ Redistributables
        Write-Host "Installing latest Visual C++ Redistributables..." -ForegroundColor Yellow
        # Download and install both x86 and x64 versions
        $VCRedist2015_2022_x64 = "https://aka.ms/vs/17/release/vc_redist.x64.exe"
        $VCRedist2015_2022_x86 = "https://aka.ms/vs/17/release/vc_redist.x86.exe"
    
        Invoke-WebRequest -Uri $VCRedist2015_2022_x64 -OutFile "$env:TEMP\vc_redist.x64.exe"
        Invoke-WebRequest -Uri $VCRedist2015_2022_x86 -OutFile "$env:TEMP\vc_redist.x86.exe"
    
        Start-Process -FilePath "$env:TEMP\vc_redist.x64.exe" -Args "/quiet /norestart" -Wait
        Start-Process -FilePath "$env:TEMP\vc_redist.x86.exe" -Args "/quiet /norestart" -Wait

        # Install .NET Runtime
        Write-Host "Installing .NET Runtime..." -ForegroundColor Yellow
        winget install Microsoft.DotNet.DesktopRuntime.8 --silent
        winget install Microsoft.DotNet.DesktopRuntime.9 --silent

        # Enable latest .NET runtime for all apps
        Write-Host "Enabling latest .NET runtime..." -ForegroundColor Yellow
        Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\.NETFramework" -Name "OnlyUseLatestCLR" -Value 1
        Set-ItemProperty -Path "HKLM:\SOFTWARE\Wow6432Node\Microsoft\.NETFramework" -Name "OnlyUseLatestCLR" -Value 1
        
        # Set flag that setup has been run
        New-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Setup" -Name "BasicSetupComplete" -Value 1 -PropertyType DWord
    }

# Define the apps in order with their corresponding executable paths
    $appConfig = @{
        "Chrome" = @{
            Paths = @(
                "${env:ProgramFiles}\Google\Chrome\Application\chrome.exe",
                "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe"
            )
            InstallFunction = "Download-Chrome"
        }
        "Edge" = @{
            Paths = @(
                "${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe",
                "${env:ProgramFiles}\Microsoft\Edge\Application\msedge.exe"
            )
        }
        "Adobe Reader" = @{
            Paths = @(
                "${env:ProgramFiles(x86)}\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe",
                "${env:ProgramFiles}\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe"
            )
            InstallFunction = "Download-AdobeReader"
        }
        "Word" = @{
            Paths = @(
                "${env:ProgramFiles}\Microsoft Office\root\Office16\WINWORD.EXE",
                "${env:ProgramFiles(x86)}\Microsoft Office\root\Office16\WINWORD.EXE"
            )
        }
        "Excel" = @{
            Paths = @(
                "${env:ProgramFiles}\Microsoft Office\root\Office16\EXCEL.EXE",
                "${env:ProgramFiles(x86)}\Microsoft Office\root\Office16\EXCEL.EXE"
            )
        }
        "PowerPoint" = @{
            Paths = @(
                "${env:ProgramFiles}\Microsoft Office\root\Office16\POWERPNT.EXE",
                "${env:ProgramFiles(x86)}\Microsoft Office\root\Office16\POWERPNT.EXE"
            )
        }
        "Outlook" = @{
            Paths = @(
                "${env:ProgramFiles}\Microsoft Office\root\Office16\OUTLOOK.EXE",
                "${env:ProgramFiles(x86)}\Microsoft Office\root\Office16\OUTLOOK.EXE"
            )
        }
         "Teams" = @{
            Paths = @(
                "shell:AppsFolder\MSTeams_8wekyb3d8bbwe!MSTeams"
            )
            InstallFunction = "Download-MicrosoftTeams"
        }
    }

    # Create a combined XML for all applications
    $xmlTemplate = @"
<?xml version="1.0" encoding="utf-8"?>
<LayoutModificationTemplate
    xmlns="http://schemas.microsoft.com/Start/2014/LayoutModification"
    xmlns:defaultlayout="http://schemas.microsoft.com/Start/2014/FullDefaultLayout"
    xmlns:start="http://schemas.microsoft.com/Start/2014/StartLayout"
    xmlns:taskbar="http://schemas.microsoft.com/Start/2014/TaskbarLayout"
    Version="1">
  <CustomTaskbarLayoutCollection PinListPlacement="Replace">
    <defaultlayout:TaskbarLayout>
      <taskbar:TaskbarPinList>
        <taskbar:DesktopApp DesktopApplicationID="Microsoft.Windows.Explorer" />
"@

# Modify the processing loop to skip installation check if the app is found
# Process each application
    $appsToProcess = @("Chrome", "Edge", "Word", "Excel", "PowerPoint", "Outlook", "Teams")
    foreach ($currentAppName in $appsToProcess) {
        try {
            Write-Host "Processing $currentAppName..." -ForegroundColor Yellow
            
            # Check if the app exists in the configuration
            if (-not $appConfig.ContainsKey($currentAppName)) {
                Write-Host "Configuration not found for $currentAppName" -ForegroundColor Red
                continue
            }
            
            $appInfo = $appConfig[$currentAppName]
            
            # Find first existing path for the app
            $exePath = $appInfo.Paths | Where-Object { Test-Path $_ } | Select-Object -First 1
            
            if ($exePath) {
                Write-Host "Found $currentAppName at: $exePath" -ForegroundColor Green
                $xmlTemplate += "        <taskbar:DesktopApp DesktopApplicationLinkPath=`"$exePath`" />`n"
            }
            else {
                # Special handling for apps with install functions
                if ($appInfo.ContainsKey('InstallFunction')) {
                    Write-Host "$currentAppName is not installed." -ForegroundColor Yellow
                    Write-Host "Would you like to install $currentAppName`? (Y/N)" -ForegroundColor Cyan
                    $installChoice = Read-Host
                    
                    if ($installChoice.ToUpper() -eq 'Y') {
                        $installSuccess = & $appInfo.InstallFunction
                        if ($installSuccess) {
                            Start-Sleep -Seconds 5  # Wait for installation to complete
                            $exePath = $appInfo.Paths | Where-Object { Test-Path $_ } | Select-Object -First 1
                            if ($exePath) {
                                $xmlTemplate += "        <taskbar:DesktopApp DesktopApplicationLinkPath=`"$exePath`" />`n"
                            }
                        }
                    }
                    else {
                        Write-Host "Skipping $currentAppName installation" -ForegroundColor Yellow
                    }
                }
                else {
                    Write-Host "Could not find $currentAppName installation" -ForegroundColor Red
                }
            }
        }
        catch {
            Write-Host "An error occurred while processing $currentAppName`: $($_.Exception.Message)" -ForegroundColor Red
        }
        
        Write-Host "----------------------------------------"
    }

    # Close the XML template
    $xmlTemplate += @"
      </taskbar:TaskbarPinList>
    </defaultlayout:TaskbarLayout>
 </CustomTaskbarLayoutCollection>
</LayoutModificationTemplate>
"@

   
    # Create provisioning directory if it doesn't exist
    $provisioningPath = "$env:ProgramData\provisioning"
    if (!(Test-Path $provisioningPath)) {
        New-Item -Path $provisioningPath -ItemType Directory -Force | Out-Null
    }

    # Save the layout file
    $layoutFile = Join-Path $provisioningPath "taskbar_layout.xml"
    $xmlTemplate | Out-File $layoutFile -Encoding utf8 -Force

    # Configure registry settings
    $explorerKeyPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Explorer"
    if (!(Test-Path $explorerKeyPath)) {
        New-Item -Path $explorerKeyPath -Force | Out-Null
    }

    # Set registry values
    Set-ItemProperty -Path $explorerKeyPath -Name "StartLayoutFile" -Value $layoutFile -Type ExpandString
    Set-ItemProperty -Path $explorerKeyPath -Name "LockedStartLayout" -Value 1 -Type DWord

    # Force a taskbar refresh
    Stop-Process -Name explorer -Force -ErrorAction SilentlyContinue
    Start-Process explorer

    # Install .NET Framework 3.5 and 4.8.1 Cumulative Update KB5050577
    Write-Host "Installing .NET Framework 3.5 and 4.8.1 Cumulative Update (KB5050577)..." -ForegroundColor Yellow
    $kb5050577_url = "https://catalog.s.download.windowsupdate.com/d/msdownload/update/software/secu/2023/04/windows11.0-kb5050577-x64_103e7ede983a848c75a63c2bbd8f2bc5e4498771.cab"
    $kb5050577_file = "windows11.0-kb5050577-x64_103e7ede983a848c75a63c2bbd8f2bc5e4498771.cab" 
    $kb5050577_temp = "$env:TEMP\$kb5050577_file"

    Try 
    {
        Invoke-WebRequest -Uri $kb5050577_url -OutFile $kb5050577_temp
        DISM.exe /Online /Add-Package /PackagePath:$kb5050577_temp /NoRestart
        Write-Host ".NET Framework 3.5 and 4.8.1 Cumulative Update installed successfully" -ForegroundColor Green
    }
    Catch
    {
        Write-Host "Failed to install .NET Framework 3.5 and 4.8.1 Cumulative Update" -ForegroundColor Red
    }
   
    Write-Host "`nBasic setup completed. The taskbar has been configured with your applications." -ForegroundColor Green
    Write-Host "Note: You may need to log out and log back in to see all changes." -ForegroundColor Yellow
    
    Write-Host "Press any key to continue..."
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')  

    #Add printer setup prompt
    Write-Host "`nWould you like to set up any printers? (Y/N)" -ForegroundColor Yellow
    $printerSetup = Read-Host
    if ($printerSetup.ToUpper() -eq 'Y') {
        Handle-PrinterSetup
    }
    

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

function Configure-WindowsUpdate {
    # Enable getting updates as soon as they're available
    Write-Host "Configuring Windows Update for immediate updates..." -ForegroundColor Yellow
    If (!(Test-Path "HKLM:\SOFTWARE\Microsoft\WindowsUpdate\UX\Settings")) {
        New-Item -Path "HKLM:\SOFTWARE\Microsoft\WindowsUpdate\UX\Settings" -Force | Out-Null
    }
    Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\WindowsUpdate\UX\Settings" -Name "IsContinuouslyAvailable" -Value 1
}

     
    Write-Host "`nPress any key to continue..."
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
}

############################################################################################################
#                                        INSTALL CITRIX - WELLCARE                                         #
#                                                                                                          #
############################################################################################################



function Get-CitrixInstallationStatus {
    param (
        [Parameter(Mandatory=$true)]
        [DateTime]$InstallStartTime
    )
    try {
        Write-LogMessage "Checking installation status in Event Logs since $($InstallStartTime.ToString('HH:mm:ss'))..."
        
        $events = Get-WinEvent -FilterHashtable @{
            LogName = 'Application'
            ProviderName = 'MsiInstaller'
            ID = @(1033, 1042)
            StartTime = $InstallStartTime
        } -ErrorAction SilentlyContinue

        if ($events) {
            $events = $events | Sort-Object TimeCreated
            
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
        
        Start-Sleep -Seconds 2
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
        
        if (-not (Stop-CitrixProcesses)) {
            Write-LogMessage "Warning: Could not stop all Citrix processes" -Level 'Warning'
        }
        
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
        
        Remove-CitrixRegistry

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

        $process = Start-Process -FilePath $InstallerPath -ArgumentList "/silent /noreboot" -PassThru

        $installSuccess = $false
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
        
        $citrixFiles = Get-ChildItem -Path $TempDir -Filter "Citrix*"
        foreach ($file in $citrixFiles) {
            try {
                Remove-Item $file.FullName -Force
                Write-LogMessage "Deleted: $($file.FullName)" -ForegroundColor Green
            } catch {
                Write-LogMessage "Failed to delete: $($file.FullName) - $($_.Exception.Message)" -Level 'Warning'
            }
        }
        
        Write-LogMessage "Citrix-related temporary files cleaned up successfully."
    }
    catch {
        Write-LogMessage "Error cleaning temporary files: $($_.Exception.Message)" -Level 'Warning'
    }
}

############################################################################################################
#                                        SYSTEM INFORMATION                                                #
############################################################################################################

function Show-SystemInfo {
    try {
        Clear-Host
        Write-Host "`n=== System Information ===" -ForegroundColor Cyan
        
        # Computer System Information
        $computerSystem = Get-CimInstance Win32_ComputerSystem
        Write-Host "`nComputer System" -ForegroundColor Green
        Write-Host "Manufacturer: $($computerSystem.Manufacturer)"
        Write-Host "Model: $($computerSystem.Model)"
        Write-Host "System Type: $($computerSystem.SystemType)"
        Write-Host "Number of Processors: $($computerSystem.NumberOfProcessors)"
        Write-Host "Total Physical Memory: $([math]::Round($computerSystem.TotalPhysicalMemory / 1GB, 2)) GB"

        # Operating System Information
        $operatingSystem = Get-CimInstance Win32_OperatingSystem
        Write-Host "`nOperating System" -ForegroundColor Green
        Write-Host "Caption: $($operatingSystem.Caption)"
        Write-Host "Version: $($operatingSystem.Version)"
        Write-Host "Build Number: $($operatingSystem.BuildNumber)"
        Write-Host "Install Date: $($operatingSystem.InstallDate)"

        # Processor Information
        $processor = Get-CimInstance Win32_Processor
        Write-Host "`nProcessor" -ForegroundColor Green
        Write-Host "Name: $($processor.Name)"
        Write-Host "Description: $($processor.Description)"
        Write-Host "Maximum Clock Speed: $($processor.MaxClockSpeed) MHz"

        # Disk Information
        Write-Host "`nDisk Drives" -ForegroundColor Green
        $drives = Get-CimInstance Win32_LogicalDisk | Where-Object {$_.DriveType -eq 3}
        foreach ($drive in $drives) {
            $sizeGB = [math]::Round($drive.Size / 1GB, 2)
            $freeGB = [math]::Round($drive.FreeSpace / 1GB, 2)
            Write-Host "Drive $($drive.DeviceID) - Total Size: $sizeGB GB, Free Space: $freeGB GB"
        }

        # Network Adapters
        Write-Host "`nNetwork Adapters" -ForegroundColor Green
        $networkAdapters = Get-NetAdapter | Where-Object { $_.Status -eq "Up" }
        foreach ($adapter in $networkAdapters) {
            $ipAddresses = (Get-NetIPAddress -InterfaceIndex $adapter.ifIndex -AddressFamily IPv4).IPAddress
            Write-Host "Name: $($adapter.Name)"
            Write-Host "  Status: $($adapter.Status)"
            Write-Host "  Speed: $($adapter.LinkSpeed)"
            Write-Host "  IP Address(es): $($ipAddresses -join ', ')"
        }

        # Memory Information
        Write-Host "`nMemory Modules" -ForegroundColor Green
        $memoryModules = Get-CimInstance Win32_PhysicalMemory
        foreach ($module in $memoryModules) {
            Write-Host "Capacity: $([math]::Round($module.Capacity / 1GB, 2)) GB"
            Write-Host "Manufacturer: $($module.Manufacturer)"
            Write-Host "Part Number: $($module.PartNumber)"
        }

        Write-Host "`nPress any key to continue..." -ForegroundColor Yellow
        $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
    }
    catch {
        Write-Host "An error occurred while retrieving system information:" -ForegroundColor Red
        Write-Host $_.Exception.Message -ForegroundColor Red
        Write-Host "`nPress any key to continue..." -ForegroundColor Yellow
        $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
    }
}

function Show-StarsStrikesMenu {
    do {
        Clear-Host
        Write-Host "`n------ Stars and Strikes Options ------" -ForegroundColor Yellow
        Write-Host "1. YellowDog Setup" -ForegroundColor Cyan
        Write-Host "2. Install TeamViewer Host" -ForegroundColor Cyan
        Write-Host "Q. Return to Main Menu" -ForegroundColor Red
        Write-Host "-------------------------------------" -ForegroundColor Yellow
        
        $starsChoice = Read-Host "Enter your choice"
        
        switch ($starsChoice) {
            '1' {
                Deploy-StarsFiles
            }
            '2' {
                Write-Host "Installing TeamViewer Host..." -ForegroundColor Yellow
                $result = Download-TeamViewerHost
                if ($result) {
                    Write-Host "TeamViewer Host installation completed." -ForegroundColor Green
                } else {
                    Write-Host "TeamViewer Host installation failed." -ForegroundColor Red
                }
                Write-Host "Press any key to continue..."
                $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
            }
            'Q' {
                return
            }
            default {
                Write-Host "Invalid choice. Please select a valid option." -ForegroundColor Red
                Write-Host "Press any key to continue..."
                $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
            }
        }
    } while ($true)
}

############################################################################################################
#                                        Remove Bloat                                                      #
#                                                                                                          #
############################################################################################################



function Remove-Bloat {
    param (
        [string[]]$customwhitelist = @()  # Set default value to an empty array
    )
    ##Elevate if needed

If (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]'Administrator')) {
    write-output "You didn't run this script as an Administrator. This script will self elevate to run as an Administrator and continue."
    Start-Sleep 1
    write-output "                                               3"
    Start-Sleep 1
    write-output "                                               2"
    Start-Sleep 1
    write-output "                                               1"
    Start-Sleep 1
    Start-Process powershell.exe -ArgumentList ("-NoProfile -ExecutionPolicy Bypass -File `"{0}`" -WhitelistApps {1}" -f $PSCommandPath, ($WhitelistApps -join ',')) -Verb RunAs
    Exit
}

#no errors throughout
$ErrorActionPreference = 'silentlycontinue'


#Create Folder
$DebloatFolder = "C:\ProgramData\Debloat"
If (Test-Path $DebloatFolder) {
    Write-Output "$DebloatFolder exists. Skipping."
}
Else {
    Write-Output "The folder '$DebloatFolder' doesn't exist. This folder will be used for storing logs created after the script runs. Creating now."
    Start-Sleep 1
    New-Item -Path "$DebloatFolder" -ItemType Directory
    Write-Output "The folder $DebloatFolder was successfully created."
}

Start-Transcript -Path "C:\ProgramData\Debloat\Debloat.log"

############################################################################################################
#                                        Remove AppX Packages                                              #
#                                                                                                          #
############################################################################################################

#Removes AppxPackages
$WhitelistedApps = @(
    'Microsoft.WindowsNotepad',
    'Microsoft.CompanyPortal',
    'Microsoft.ScreenSketch',
    'Microsoft.Paint3D',
    'Microsoft.WindowsCalculator',
    'Microsoft.WindowsStore',
    'Microsoft.Windows.Photos',
    'CanonicalGroupLimited.UbuntuonWindows',
    'Microsoft.MicrosoftStickyNotes',
    'Microsoft.MSPaint',
    'Microsoft.WindowsCamera',
    '.NET Framework',
    'Microsoft.HEIFImageExtension',
    'Microsoft.ScreenSketch',
    'Microsoft.StorePurchaseApp',
    'Microsoft.VP9VideoExtensions',
    'Microsoft.WebMediaExtensions',
    'Microsoft.WebpImageExtension',
    'Microsoft.DesktopAppInstaller',
    'WindSynthBerry',
    'MIDIBerry',
    'Slack',
    'Microsoft.SecHealthUI',
    'WavesAudio.MaxxAudioProforDell2019',
    'Dell Optimizer Core',
    'Dell SupportAssist Remediation',
    'Dell SupportAssist OS Recovery Plugin for Dell Update',
    'Dell Pair',
    'Dell Display Manager 2.0',
    'Dell Display Manager 2.1',
    'Dell Display Manager 2.2',
    'Dell Peripheral Manager',
    'MSTeams',
    'Microsoft.Paint',
    'Microsoft.OutlookForWindows',
    'Microsoft.WindowsTerminal',
    'Microsoft.MicrosoftEdge.Stable'
)
##If $customwhitelist is set, split on the comma and add to whitelist
if ($customwhitelist) {
    $customWhitelistApps = $customwhitelist -split ","
    foreach ($whitelistapp in $customwhitelistapps) {
        ##Add to the array
        $WhitelistedApps += $whitelistapp
    }
}

#NonRemovable Apps that where getting attempted and the system would reject the uninstall, speeds up debloat and prevents 'initalizing' overlay when removing apps
$NonRemovable = @(
    '1527c705-839a-4832-9118-54d4Bd6a0c89',
    'c5e2524a-ea46-4f67-841f-6a9465d9d515',
    'E2A4F912-2574-4A75-9BB0-0D023378592B',
    'F46D4000-FD22-4DB4-AC8E-4E1DDDE828FE',
    'InputApp',
    'Microsoft.AAD.BrokerPlugin',
    'Microsoft.AccountsControl',
    'Microsoft.BioEnrollment',
    'Microsoft.CredDialogHost',
    'Microsoft.ECApp',
    'Microsoft.LockApp',
    'Microsoft.MicrosoftEdgeDevToolsClient',
    'Microsoft.MicrosoftEdge',
    'Microsoft.PPIProjection',
    'Microsoft.Win32WebViewHost',
    'Microsoft.Windows.Apprep.ChxApp',
    'Microsoft.Windows.AssignedAccessLockApp',
    'Microsoft.Windows.CapturePicker',
    'Microsoft.Windows.CloudExperienceHost',
    'Microsoft.Windows.ContentDeliveryManager',
    'Microsoft.Windows.Cortana',
    'Microsoft.Windows.NarratorQuickStart',
    'Microsoft.Windows.ParentalControls',
    'Microsoft.Windows.PeopleExperienceHost',
    'Microsoft.Windows.PinningConfirmationDialog',
    'Microsoft.Windows.SecHealthUI',
    'Microsoft.Windows.SecureAssessmentBrowser',
    'Microsoft.Windows.ShellExperienceHost',
    'Microsoft.Windows.XGpuEjectDialog',
    'Microsoft.XboxGameCallableUI',
    'Windows.CBSPreview',
    'windows.immersivecontrolpanel',
    'Windows.PrintDialog',
    'Microsoft.VCLibs.140.00',
    'Microsoft.Services.Store.Engagement',
    'Microsoft.UI.Xaml.2.0',
    'Microsoft.AsyncTextService',
    'Microsoft.UI.Xaml.CBS',
    'Microsoft.Windows.CallingShellApp',
    'Microsoft.Windows.OOBENetworkConnectionFlow',
    'Microsoft.Windows.PrintQueueActionCenter',
    'Microsoft.Windows.StartMenuExperienceHost',
    'MicrosoftWindows.Client.CBS',
    'MicrosoftWindows.Client.Core',
    'MicrosoftWindows.UndockedDevKit',
    'NcsiUwpApp',
    'Microsoft.NET.Native.Runtime.2.2',
    'Microsoft.NET.Native.Framework.2.2',
    'Microsoft.UI.Xaml.2.8',
    'Microsoft.UI.Xaml.2.7',
    'Microsoft.UI.Xaml.2.3',
    'Microsoft.UI.Xaml.2.4',
    'Microsoft.UI.Xaml.2.1',
    'Microsoft.UI.Xaml.2.2',
    'Microsoft.UI.Xaml.2.5',
    'Microsoft.UI.Xaml.2.6',
    'Microsoft.VCLibs.140.00.UWPDesktop',
    'MicrosoftWindows.Client.LKG',
    'MicrosoftWindows.Client.FileExp',
    'Microsoft.WindowsAppRuntime.1.5',
    'Microsoft.WindowsAppRuntime.1.3',
    'Microsoft.WindowsAppRuntime.1.1',
    'Microsoft.WindowsAppRuntime.1.2',
    'Microsoft.WindowsAppRuntime.1.4',
    'Microsoft.Windows.OOBENetworkCaptivePortal',
    'Microsoft.Windows.Search'
)

##Combine the two arrays
$appstoignore = $WhitelistedApps += $NonRemovable

##Bloat list for future reference
$Bloatware = @(
#Unnecessary Windows 10/11 AppX Apps
"*ActiproSoftwareLLC*"
"*AdobeSystemsIncorporated.AdobePhotoshopExpress*"
"*BubbleWitch3Saga*"
"*CandyCrush*"
"*DevHome*"
"*Disney*"
"*Dolby*"
"*Duolingo-LearnLanguagesforFree*"
"*EclipseManager*"
"*Facebook*"
"*Flipboard*"
"*gaming*"
"*Minecraft*"
#"*Office*"
"*PandoraMediaInc*"
"*Royal Revolt*"
"*Speed Test*"
"*Spotify*"
"*Sway*"
"*Twitter*"
"*Wunderlist*"
"AD2F1837.HPPrinterControl"
"AppUp.IntelGraphicsExperience"
"C27EB4BA.DropboxOEM*"
"Disney.37853FC22B2CE"
"DolbyLaboratories.DolbyAccess"
"DolbyLaboratories.DolbyAudio"
"E0469640.SmartAppearance"
"Microsoft.549981C3F5F10"
"Microsoft.AV1VideoExtension"
"Microsoft.BingNews"
"Microsoft.BingSearch"
"Microsoft.BingWeather"
"Microsoft.GetHelp"
"Microsoft.Getstarted"
"Microsoft.GamingApp"
"Microsoft.HEVCVideoExtension"
"Microsoft.Messaging"
"Microsoft.Microsoft3DViewer"
"Microsoft.MicrosoftEdge.Stable"
"Microsoft.MicrosoftJournal"
"Microsoft.MicrosoftOfficeHub"
"Microsoft.MicrosoftSolitaireCollection"
"Microsoft.MixedReality.Portal"
"Microsoft.MPEG2VideoExtension"
"Microsoft.News"
"Microsoft.Office.Lens"
#"Microsoft.Office.OneNote"
"Microsoft.Office.Sway"
"Microsoft.OneConnect"
#"Microsoft.OneDriveSync"
"Microsoft.People"
"Microsoft.PowerAutomateDesktop"
"Microsoft.PowerAutomateDesktopCopilotPlugin"
"Microsoft.Print3D"
#"Microsoft.RemoteDesktop"
"Microsoft.SkypeApp"
"Microsoft.StorePurchaseApp"
"Microsoft.SysinternalsSuite"
#"Microsoft.Teams"
"Microsoft.Todos"
"Microsoft.Whiteboard"
"Microsoft.Windows.DevHome"
"Microsoft.WindowsAlarms"
#"Microsoft.WindowsCamera"
"Microsoft.windowscommunicationsapps"
"Microsoft.WindowsFeedbackHub"
"Microsoft.WindowsMaps"
"Microsoft.WindowsSoundRecorder"
#"Microsoft.WindowsStore"
"Microsoft.Xbox.TCUI"
"Microsoft.XboxApp"
"Microsoft.XboxGameOverlay"
"Microsoft.XboxGamingOverlay"
"Microsoft.XboxGamingOverlay_5.721.10202.0_neutral_~_8wekyb3d8bbwe"
"Microsoft.XboxIdentityProvider"
"Microsoft.XboxSpeechToTextOverlay"
"Microsoft.YourPhone"
"Microsoft.ZuneMusic"
"Microsoft.ZuneVideo"
"MicrosoftCorporationII.MicrosoftFamily"
"MicrosoftCorporationII.QuickAssist"
"MicrosoftWindows.Client.WebExperience"
"MicrosoftWindows.CrossDevice"
"MirametrixInc.GlancebyMirametrix"
"MSTeams"
"RealtimeboardInc.RealtimeBoard"
"SpotifyAB.SpotifyMusic"
#Optional: Typically not removed but you can if you need to for some reason
#"*Microsoft.Advertising.Xaml_10.1712.5.0_x64__8wekyb3d8bbwe*"
#"*Microsoft.Advertising.Xaml_10.1712.5.0_x86__8wekyb3d8bbwe*"
#"*Microsoft.BingWeather*"
#"*Microsoft.MSPaint*"
#"*Microsoft.MicrosoftStickyNotes*"
#"*Microsoft.Windows.Photos*"
#"*Microsoft.WindowsCalculator*"
#"*Microsoft.WindowsStore*"
#"Microsoft.Office.Todo.List"
#"Microsoft.Whiteboard"
#"Microsoft.WindowsCamera"
#"Microsoft.WindowsSoundRecorder"
#"Microsoft.YourPhone"
#"Microsoft.Todos"
#"Microsoft.PowerAutomateDesktop"
)


$provisioned = Get-AppxProvisionedPackage -Online | Where-Object { $_.DisplayName -in $Bloatware -and $_.DisplayName -notin $appstoignore -and $_.DisplayName -notlike 'MicrosoftWindows.Voice*' -and $_.DisplayName -notlike 'Microsoft.LanguageExperiencePack*' -and $_.DisplayName -notlike 'MicrosoftWindows.Speech*' }
foreach ($appxprov in $provisioned) {
    $packagename = $appxprov.PackageName
    $displayname = $appxprov.DisplayName
    write-output "Removing $displayname AppX Provisioning Package"
    try {
        Remove-AppxProvisionedPackage -PackageName $packagename -Online -ErrorAction SilentlyContinue
        write-output "Removed $displayname AppX Provisioning Package"
    }
    catch {
        write-output "Unable to remove $displayname AppX Provisioning Package"
    }

}


$appxinstalled = Get-AppxPackage -AllUsers | Where-Object { $_.Name -in $Bloatware -and $_.Name -notin $appstoignore  -and $_.Name -notlike 'MicrosoftWindows.Voice*' -and $_.Name -notlike 'Microsoft.LanguageExperiencePack*' -and $_.Name -notlike 'MicrosoftWindows.Speech*'}
foreach ($appxapp in $appxinstalled) {
    $packagename = $appxapp.PackageFullName
    $displayname = $appxapp.Name
    write-output "$displayname AppX Package exists"
    write-output "Removing $displayname AppX Package"
    try {
        Remove-AppxPackage -Package $packagename -AllUsers -ErrorAction SilentlyContinue
        write-output "Removed $displayname AppX Package"
    }
    catch {
        write-output "$displayname AppX Package does not exist"
    }



}


############################################################################################################
#                                        Remove Registry Keys                                              #
#                                                                                                          #
############################################################################################################

##We need to grab all SIDs to remove at user level
$UserSIDs = Get-ChildItem "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList" | Select-Object -ExpandProperty PSChildName


#These are the registry keys that it will delete.

$Keys = @(

    #Remove Background Tasks
    "HKCR:\Extensions\ContractId\Windows.BackgroundTasks\PackageId\46928bounde.EclipseManager_2.2.4.51_neutral__a5h4egax66k6y"
    "HKCR:\Extensions\ContractId\Windows.BackgroundTasks\PackageId\ActiproSoftwareLLC.562882FEEB491_2.6.18.18_neutral__24pqs290vpjk0"
    "HKCR:\Extensions\ContractId\Windows.BackgroundTasks\PackageId\Microsoft.MicrosoftOfficeHub_17.7909.7600.0_x64__8wekyb3d8bbwe"
    "HKCR:\Extensions\ContractId\Windows.BackgroundTasks\PackageId\Microsoft.PPIProjection_10.0.15063.0_neutral_neutral_cw5n1h2txyewy"
    "HKCR:\Extensions\ContractId\Windows.BackgroundTasks\PackageId\Microsoft.XboxGameCallableUI_1000.15063.0.0_neutral_neutral_cw5n1h2txyewy"
    "HKCR:\Extensions\ContractId\Windows.BackgroundTasks\PackageId\Microsoft.XboxGameCallableUI_1000.16299.15.0_neutral_neutral_cw5n1h2txyewy"

    #Windows File
    "HKCR:\Extensions\ContractId\Windows.File\PackageId\ActiproSoftwareLLC.562882FEEB491_2.6.18.18_neutral__24pqs290vpjk0"

    #Registry keys to delete if they aren't uninstalled by RemoveAppXPackage/RemoveAppXProvisionedPackage
    "HKCR:\Extensions\ContractId\Windows.Launch\PackageId\46928bounde.EclipseManager_2.2.4.51_neutral__a5h4egax66k6y"
    "HKCR:\Extensions\ContractId\Windows.Launch\PackageId\ActiproSoftwareLLC.562882FEEB491_2.6.18.18_neutral__24pqs290vpjk0"
    "HKCR:\Extensions\ContractId\Windows.Launch\PackageId\Microsoft.PPIProjection_10.0.15063.0_neutral_neutral_cw5n1h2txyewy"
    "HKCR:\Extensions\ContractId\Windows.Launch\PackageId\Microsoft.XboxGameCallableUI_1000.15063.0.0_neutral_neutral_cw5n1h2txyewy"
    "HKCR:\Extensions\ContractId\Windows.Launch\PackageId\Microsoft.XboxGameCallableUI_1000.16299.15.0_neutral_neutral_cw5n1h2txyewy"

    #Scheduled Tasks to delete
    "HKCR:\Extensions\ContractId\Windows.PreInstalledConfigTask\PackageId\Microsoft.MicrosoftOfficeHub_17.7909.7600.0_x64__8wekyb3d8bbwe"

    #Windows Protocol Keys
    "HKCR:\Extensions\ContractId\Windows.Protocol\PackageId\ActiproSoftwareLLC.562882FEEB491_2.6.18.18_neutral__24pqs290vpjk0"
    "HKCR:\Extensions\ContractId\Windows.Protocol\PackageId\Microsoft.PPIProjection_10.0.15063.0_neutral_neutral_cw5n1h2txyewy"
    "HKCR:\Extensions\ContractId\Windows.Protocol\PackageId\Microsoft.XboxGameCallableUI_1000.15063.0.0_neutral_neutral_cw5n1h2txyewy"
    "HKCR:\Extensions\ContractId\Windows.Protocol\PackageId\Microsoft.XboxGameCallableUI_1000.16299.15.0_neutral_neutral_cw5n1h2txyewy"

    #Windows Share Target
    "HKCR:\Extensions\ContractId\Windows.ShareTarget\PackageId\ActiproSoftwareLLC.562882FEEB491_2.6.18.18_neutral__24pqs290vpjk0"
)

#This writes the output of each key it is removing and also removes the keys listed above.
ForEach ($Key in $Keys) {
    write-output "Removing $Key from registry"
    Remove-Item $Key -Recurse
}


#Disables Windows Feedback Experience
write-output "Disabling Windows Feedback Experience program"
$Advertising = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\AdvertisingInfo"
If (!(Test-Path $Advertising)) {
    New-Item $Advertising
}
If (Test-Path $Advertising) {
    Set-ItemProperty $Advertising Enabled -Value 0
}

#Stops Cortana from being used as part of your Windows Search Function
write-output "Stopping Cortana from being used as part of your Windows Search Function"
$Search = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Windows Search"
If (!(Test-Path $Search)) {
    New-Item $Search
}
If (Test-Path $Search) {
    Set-ItemProperty $Search AllowCortana -Value 0
}

#Disables Web Search in Start Menu
write-output "Disabling Bing Search in Start Menu"
$WebSearch = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Windows Search"
If (!(Test-Path $WebSearch)) {
    New-Item $WebSearch
}
Set-ItemProperty $WebSearch DisableWebSearch -Value 1
##Loop through all user SIDs in the registry and disable Bing Search
foreach ($sid in $UserSIDs) {
    $WebSearch = "Registry::HKU\$sid\SOFTWARE\Microsoft\Windows\CurrentVersion\Search"
    If (!(Test-Path $WebSearch)) {
        New-Item $WebSearch
    }
    Set-ItemProperty $WebSearch BingSearchEnabled -Value 0
}

Set-ItemProperty "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Search" BingSearchEnabled -Value 0


#Stops the Windows Feedback Experience from sending anonymous data
write-output "Stopping the Windows Feedback Experience program"
$Period = "HKCU:\Software\Microsoft\Siuf\Rules"
If (!(Test-Path $Period)) {
    New-Item $Period
}
Set-ItemProperty $Period PeriodInNanoSeconds -Value 0

##Loop and do the same
foreach ($sid in $UserSIDs) {
    $Period = "Registry::HKU\$sid\Software\Microsoft\Siuf\Rules"
    If (!(Test-Path $Period)) {
        New-Item $Period
    }
    Set-ItemProperty $Period PeriodInNanoSeconds -Value 0
}

##Disables games from showing in Search bar
write-output "Adding Registry key to stop games from search bar"
$registryPath = "HKLM:\	SOFTWARE\Policies\Microsoft\Windows\Windows Search"
If (!(Test-Path $registryPath)) {
    New-Item $registryPath
}
Set-ItemProperty $registryPath EnableDynamicContentInWSB -Value 0

#Prevents bloatware applications from returning and removes Start Menu suggestions
write-output "Adding Registry key to prevent bloatware apps from returning"
$registryPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\CloudContent"
$registryOEM = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\ContentDeliveryManager"
If (!(Test-Path $registryPath)) {
    New-Item $registryPath
}
Set-ItemProperty $registryPath DisableWindowsConsumerFeatures -Value 1

If (!(Test-Path $registryOEM)) {
    New-Item $registryOEM
}
Set-ItemProperty $registryOEM  ContentDeliveryAllowed -Value 0
Set-ItemProperty $registryOEM  OemPreInstalledAppsEnabled -Value 0
Set-ItemProperty $registryOEM  PreInstalledAppsEnabled -Value 0
Set-ItemProperty $registryOEM  PreInstalledAppsEverEnabled -Value 0
Set-ItemProperty $registryOEM  SilentInstalledAppsEnabled -Value 0
Set-ItemProperty $registryOEM  SystemPaneSuggestionsEnabled -Value 0

##Loop through users and do the same
foreach ($sid in $UserSIDs) {
    $registryOEM = "Registry::HKU\$sid\SOFTWARE\Microsoft\Windows\CurrentVersion\ContentDeliveryManager"
    If (!(Test-Path $registryOEM)) {
        New-Item $registryOEM
    }
    Set-ItemProperty $registryOEM  ContentDeliveryAllowed -Value 0
    Set-ItemProperty $registryOEM  OemPreInstalledAppsEnabled -Value 0
    Set-ItemProperty $registryOEM  PreInstalledAppsEnabled -Value 0
    Set-ItemProperty $registryOEM  PreInstalledAppsEverEnabled -Value 0
    Set-ItemProperty $registryOEM  SilentInstalledAppsEnabled -Value 0
    Set-ItemProperty $registryOEM  SystemPaneSuggestionsEnabled -Value 0
}

#Preping mixed Reality Portal for removal
write-output "Setting Mixed Reality Portal value to 0 so that you can uninstall it in Settings"
$Holo = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Holographic"
If (Test-Path $Holo) {
    Set-ItemProperty $Holo  FirstRunSucceeded -Value 0
}

##Loop through users and do the same
foreach ($sid in $UserSIDs) {
    $Holo = "Registry::HKU\$sid\Software\Microsoft\Windows\CurrentVersion\Holographic"
    If (Test-Path $Holo) {
        Set-ItemProperty $Holo  FirstRunSucceeded -Value 0
    }
}

#Disables Wi-fi Sense
write-output "Disabling Wi-Fi Sense"
$WifiSense1 = "HKLM:\SOFTWARE\Microsoft\PolicyManager\default\WiFi\AllowWiFiHotSpotReporting"
$WifiSense2 = "HKLM:\SOFTWARE\Microsoft\PolicyManager\default\WiFi\AllowAutoConnectToWiFiSenseHotspots"
$WifiSense3 = "HKLM:\SOFTWARE\Microsoft\WcmSvc\wifinetworkmanager\config"
If (!(Test-Path $WifiSense1)) {
    New-Item $WifiSense1
}
Set-ItemProperty $WifiSense1  Value -Value 0
If (!(Test-Path $WifiSense2)) {
    New-Item $WifiSense2
}
Set-ItemProperty $WifiSense2  Value -Value 0
Set-ItemProperty $WifiSense3  AutoConnectAllowedOEM -Value 0

#Disables live tiles
write-output "Disabling live tiles"
$Live = "HKCU:\SOFTWARE\Policies\Microsoft\Windows\CurrentVersion\PushNotifications"
If (!(Test-Path $Live)) {
    New-Item $Live
}
Set-ItemProperty $Live  NoTileApplicationNotification -Value 1

##Loop through users and do the same
foreach ($sid in $UserSIDs) {
    $Live = "Registry::HKU\$sid\SOFTWARE\Policies\Microsoft\Windows\CurrentVersion\PushNotifications"
    If (!(Test-Path $Live)) {
        New-Item $Live
    }
    Set-ItemProperty $Live  NoTileApplicationNotification -Value 1
}

#Turns off Data Collection via the AllowTelemtry key by changing it to 0
# This is needed for Intune reporting to work, uncomment if using via other method
#write-output "Turning off Data Collection"
#$DataCollection1 = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\DataCollection"
#$DataCollection2 = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\DataCollection"
#$DataCollection3 = "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Policies\DataCollection"
#If (Test-Path $DataCollection1) {
#    Set-ItemProperty $DataCollection1  AllowTelemetry -Value 0
#}
#If (Test-Path $DataCollection2) {
#    Set-ItemProperty $DataCollection2  AllowTelemetry -Value 0
#}
#If (Test-Path $DataCollection3) {
#    Set-ItemProperty $DataCollection3  AllowTelemetry -Value 0
#}


###Enable location tracking for "find my device", uncomment if you don't need it

#Disabling Location Tracking
#write-output "Disabling Location Tracking"
#$SensorState = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Sensor\Overrides\{BFA794E4-F964-4FDB-90F6-51056BFE4B44}"
#$LocationConfig = "HKLM:\SYSTEM\CurrentControlSet\Services\lfsvc\Service\Configuration"
#If (!(Test-Path $SensorState)) {
#    New-Item $SensorState
#}
#Set-ItemProperty $SensorState SensorPermissionState -Value 0
#If (!(Test-Path $LocationConfig)) {
#    New-Item $LocationConfig
#}
#Set-ItemProperty $LocationConfig Status -Value 0

#Disables People icon on Taskbar
write-output "Disabling People icon on Taskbar"
$People = 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced\People'
If (Test-Path $People) {
    Set-ItemProperty $People -Name PeopleBand -Value 0
}

##Loop through users and do the same
foreach ($sid in $UserSIDs) {
    $People = "Registry::HKU\$sid\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced\People"
    If (Test-Path $People) {
        Set-ItemProperty $People -Name PeopleBand -Value 0
    }
}

write-output "Disabling Cortana"
$Cortana1 = "HKCU:\SOFTWARE\Microsoft\Personalization\Settings"
$Cortana2 = "HKCU:\SOFTWARE\Microsoft\InputPersonalization"
$Cortana3 = "HKCU:\SOFTWARE\Microsoft\InputPersonalization\TrainedDataStore"
If (!(Test-Path $Cortana1)) {
    New-Item $Cortana1
}
Set-ItemProperty $Cortana1 AcceptedPrivacyPolicy -Value 0
If (!(Test-Path $Cortana2)) {
    New-Item $Cortana2
}
Set-ItemProperty $Cortana2 RestrictImplicitTextCollection -Value 1
Set-ItemProperty $Cortana2 RestrictImplicitInkCollection -Value 1
If (!(Test-Path $Cortana3)) {
    New-Item $Cortana3
}
Set-ItemProperty $Cortana3 HarvestContacts -Value 0

##Loop through users and do the same
foreach ($sid in $UserSIDs) {
    $Cortana1 = "Registry::HKU\$sid\SOFTWARE\Microsoft\Personalization\Settings"
    $Cortana2 = "Registry::HKU\$sid\SOFTWARE\Microsoft\InputPersonalization"
    $Cortana3 = "Registry::HKU\$sid\SOFTWARE\Microsoft\InputPersonalization\TrainedDataStore"
    If (!(Test-Path $Cortana1)) {
        New-Item $Cortana1
    }
    Set-ItemProperty $Cortana1 AcceptedPrivacyPolicy -Value 0
    If (!(Test-Path $Cortana2)) {
        New-Item $Cortana2
    }
    Set-ItemProperty $Cortana2 RestrictImplicitTextCollection -Value 1
    Set-ItemProperty $Cortana2 RestrictImplicitInkCollection -Value 1
    If (!(Test-Path $Cortana3)) {
        New-Item $Cortana3
    }
    Set-ItemProperty $Cortana3 HarvestContacts -Value 0
}


#Removes 3D Objects from the 'My Computer' submenu in explorer
write-output "Removing 3D Objects from explorer 'My Computer' submenu"
$Objects32 = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace\{0DB7E03F-FC29-4DC6-9020-FF41B59E513A}"
$Objects64 = "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace\{0DB7E03F-FC29-4DC6-9020-FF41B59E513A}"
If (Test-Path $Objects32) {
    Remove-Item $Objects32 -Recurse
}
If (Test-Path $Objects64) {
    Remove-Item $Objects64 -Recurse
}

##Removes the Microsoft Feeds from displaying
$registryPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Windows Feeds"
$Name = "EnableFeeds"
$value = "0"

if (!(Test-Path $registryPath)) {
    New-Item -Path $registryPath -Force | Out-Null
    New-ItemProperty -Path $registryPath -Name $name -Value $value -PropertyType DWORD -Force | Out-Null
}

else {
    New-ItemProperty -Path $registryPath -Name $name -Value $value -PropertyType DWORD -Force | Out-Null
}

##Kill Cortana again
Get-AppxPackage Microsoft.549981C3F5F10 -allusers | Remove-AppxPackage



############################################################################################################
#                                   Disable unwanted OOBE screens for Device Prep                          #
#                                                                                                          #
############################################################################################################

$registryPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\OOBE"
$registryPath2 = "HKLM:\Software\Microsoft\Windows\CurrentVersion\Policies\System"
$Name1 = "DisablePrivacyExperience"
$Name2 = "DisableVoice"
$Name3 = "PrivacyConsentStatus"
$Name4 = "Protectyourpc"
$Name5 = "HideEULAPage"
$Name6 = "EnableFirstLogonAnimation"
New-ItemProperty -Path $registryPath -Name $name1 -Value 1 -PropertyType DWord -Force
New-ItemProperty -Path $registryPath -Name $name2 -Value 1 -PropertyType DWord -Force
New-ItemProperty -Path $registryPath -Name $name3 -Value 1 -PropertyType DWord -Force
New-ItemProperty -Path $registryPath -Name $name4 -Value 3 -PropertyType DWord -Force
New-ItemProperty -Path $registryPath -Name $name5 -Value 1 -PropertyType DWord -Force
New-ItemProperty -Path $registryPath2 -Name $name6 -Value 1 -PropertyType DWord -Force



############################################################################################################
#                                        Remove Learn about this picture                                   #
#                                                                                                          #
############################################################################################################

# Turn off Learn about this picture
write-output "Disabling Learn about this picture"

# Main Registry Paths
$paths = @(
    "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\ContentDeliveryManager",
    "HKLM:\SOFTWARE\Policies\Microsoft\Windows\CloudContent",
    "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Lock Screen"
)

# Create and set registry values
foreach ($path in $paths) {
    if (!(Test-Path $path)) {
        New-Item -Path $path -Force | Out-Null
    }
    
    Switch ($path) {
        {$_ -like "*ContentDeliveryManager"} {
            Set-ItemProperty -Path $path -Name "RotatingLockScreenEnabled" -Value 0 -Type DWord -Force
            Set-ItemProperty -Path $path -Name "RotatingLockScreenOverlayEnabled" -Value 0 -Type DWord -Force
            Set-ItemProperty -Path $path -Name "LockScreenImageUrl" -Value "" -Type String -Force
        }
        {$_ -like "*CloudContent"} {
            Set-ItemProperty -Path $path -Name "DisableWindowsSpotlightFeatures" -Value 1 -Type DWord -Force
            Set-ItemProperty -Path $path -Name "DisableWindowsSpotlightOnActionCenter" -Value 1 -Type DWord -Force
            Set-ItemProperty -Path $path -Name "DisableWindowsSpotlightWindowsWelcomeExperience" -Value 1 -Type DWord -Force
        }
        {$_ -like "*Lock Screen"} {
            Set-ItemProperty -Path $path -Name "LockScreenOptions" -Value 0 -Type DWord -Force
        }
    }
}

# Apply the same settings for all users
foreach ($sid in $UserSIDs) {
    $userPaths = @(
        "Registry::HKU\$sid\SOFTWARE\Microsoft\Windows\CurrentVersion\ContentDeliveryManager",
        "Registry::HKU\$sid\SOFTWARE\Microsoft\Windows\CurrentVersion\Lock Screen"
    )
    
    foreach ($path in $userPaths) {
        if (!(Test-Path $path)) {
            New-Item -Path $path -Force | Out-Null
        }
        
        Switch ($path) {
            {$_ -like "*ContentDeliveryManager"} {
                Set-ItemProperty -Path $path -Name "RotatingLockScreenEnabled" -Value 0 -Type DWord -Force
                Set-ItemProperty -Path $path -Name "RotatingLockScreenOverlayEnabled" -Value 0 -Type DWord -Force
                Set-ItemProperty -Path $path -Name "LockScreenImageUrl" -Value "" -Type String -Force
            }
            {$_ -like "*Lock Screen"} {
                Set-ItemProperty -Path $path -Name "LockScreenOptions" -Value 0 -Type DWord -Force
            }
        }
    }
}

# Disable the Windows Spotlight Service
$spotlightService = "RetailDemo"
Set-Service -Name $spotlightService -StartupType Disabled -ErrorAction SilentlyContinue
Stop-Service -Name $spotlightService -Force -ErrorAction SilentlyContinue


############################################################################################################
#                                     Disable Consumer Experiences                                         #
#                                                                                                          #
############################################################################################################

#Disabling consumer experience
write-output "Disabling consumer experience"
$consumer = 'HKLM:\\SOFTWARE\Policies\Microsoft\Windows\CloudContent'
If (Test-Path $consumer) {
    Set-ItemProperty $consumer -Name "DisableWindowsConsumerFeatures" -Value 1
}


############################################################################################################
#                                                   Disable Spotlight                                      #
#                                                                                                          #
############################################################################################################

write-output "Disabling Windows Spotlight on lockscreen"
$spotlight = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager'
If (Test-Path $spotlight) {
    Set-ItemProperty $spotlight -Name "RotatingLockScreenOverlayEnabled" -Value 0
    Set-ItemProperty $spotlight -Name "RotatingLockScreenEnabled" -Value 0
}

##Loop through users and do the same
foreach ($sid in $UserSIDs) {
    $spotlight = "Registry::HKU\$sid\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager"
    If (Test-Path $spotlight) {
        Set-ItemProperty $spotlight -Name "RotatingLockScreenOverlayEnabled" -Value 0
        Set-ItemProperty $spotlight -Name "RotatingLockScreenEnabled" -Value 0
    }
}

write-output "Disabling Windows Spotlight on background"
$spotlight = 'HKCU:\Software\Policies\Microsoft\Windows\CloudContent'
If (Test-Path $spotlight) {
    Set-ItemProperty $spotlight -Name "DisableSpotlightCollectionOnDesktop" -Value 1
    Set-ItemProperty $spotlight -Name "DisableWindowsSpotlightFeatures" -Value 1
}

##Loop through users and do the same
foreach ($sid in $UserSIDs) {
    $spotlight = "Registry::HKU\$sid\Software\Policies\Microsoft\Windows\CloudContent"
    If (Test-Path $spotlight) {
        Set-ItemProperty $spotlight -Name "DisableSpotlightCollectionOnDesktop" -Value 1
        Set-ItemProperty $spotlight -Name "DisableWindowsSpotlightFeatures" -Value 1
    }
}

############################################################################################################
#                                        Remove Scheduled Tasks                                            #
#                                                                                                          #
############################################################################################################

#Disables scheduled tasks that are considered unnecessary
write-output "Disabling scheduled tasks"
$task1 = Get-ScheduledTask -TaskName XblGameSaveTaskLogon -ErrorAction SilentlyContinue
if ($null -ne $task1) {
    Get-ScheduledTask  XblGameSaveTaskLogon | Disable-ScheduledTask -ErrorAction SilentlyContinue
}
$task2 = Get-ScheduledTask -TaskName XblGameSaveTask -ErrorAction SilentlyContinue
if ($null -ne $task2) {
    Get-ScheduledTask  XblGameSaveTask | Disable-ScheduledTask -ErrorAction SilentlyContinue
}
$task3 = Get-ScheduledTask -TaskName Consolidator -ErrorAction SilentlyContinue
if ($null -ne $task3) {
    Get-ScheduledTask  Consolidator | Disable-ScheduledTask -ErrorAction SilentlyContinue
}
$task4 = Get-ScheduledTask -TaskName UsbCeip -ErrorAction SilentlyContinue
if ($null -ne $task4) {
    Get-ScheduledTask  UsbCeip | Disable-ScheduledTask -ErrorAction SilentlyContinue
}
$task5 = Get-ScheduledTask -TaskName DmClient -ErrorAction SilentlyContinue
if ($null -ne $task5) {
    Get-ScheduledTask  DmClient | Disable-ScheduledTask -ErrorAction SilentlyContinue
}
$task6 = Get-ScheduledTask -TaskName DmClientOnScenarioDownload -ErrorAction SilentlyContinue
if ($null -ne $task6) {
    Get-ScheduledTask  DmClientOnScenarioDownload | Disable-ScheduledTask -ErrorAction SilentlyContinue
}


############################################################################################################
#                                             Disable Services                                             #
#                                                                                                          #
############################################################################################################
##write-output "Stopping and disabling Diagnostics Tracking Service"
#Disabling the Diagnostics Tracking Service
##Stop-Service "DiagTrack"
##Set-Service "DiagTrack" -StartupType Disabled


############################################################################################################
#                                        Windows 11 Specific                                               #
#                                                                                                          #
############################################################################################################
#Windows 11 Customisations
write-output "Removing Windows 11 Customisations"


##Disable Feeds
$registryPath = "HKLM:\SOFTWARE\Policies\Microsoft\Dsh"
If (!(Test-Path $registryPath)) {
    New-Item $registryPath
}
Set-ItemProperty $registryPath "AllowNewsAndInterests" -Value 0
write-output "Disabled Feeds"

############################################################################################################
#                                           Windows Backup App                                             #
#                                                                                                          #
############################################################################################################
$version = Get-CimInstance Win32_OperatingSystem | Select-Object -ExpandProperty Caption
if ($version -like "*Windows 10*") {
    write-output "Removing Windows Backup"
    $filepath = "C:\Windows\SystemApps\MicrosoftWindows.Client.CBS_cw5n1h2txyewy\WindowsBackup\Assets"
    if (Test-Path $filepath) {
        Remove-WindowsPackage -Online -PackageName "Microsoft-Windows-UserExperience-Desktop-Package~31bf3856ad364e35~amd64~~10.0.19041.3393"

        ##Add back snipping tool functionality
        write-output "Adding Windows Shell Components"
        DISM /Online /Add-Capability /CapabilityName:Windows.Client.ShellComponents~~~~0.0.1.0
        write-output "Components Added"
    }
    write-output "Removed"
}

############################################################################################################
#                                           Windows CoPilot                                                #
#                                                                                                          #
############################################################################################################
$version = Get-CimInstance Win32_OperatingSystem | Select-Object -ExpandProperty Caption
if ($version -like "*Windows 11*") {
    write-output "Removing Windows Copilot"
    # Define the registry key and value
    $registryPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsCopilot"
    $propertyName = "TurnOffWindowsCopilot"
    $propertyValue = 1

    # Check if the registry key exists
    if (!(Test-Path $registryPath)) {
        # If the registry key doesn't exist, create it
        New-Item -Path $registryPath -Force | Out-Null
    }

    # Get the property value
    $currentValue = Get-ItemProperty -Path $registryPath -Name $propertyName -ErrorAction SilentlyContinue

    # Check if the property exists and if its value is different from the desired value
    if ($null -eq $currentValue -or $currentValue.$propertyName -ne $propertyValue) {
        # If the property doesn't exist or its value is different, set the property value
        Set-ItemProperty -Path $registryPath -Name $propertyName -Value $propertyValue
    }


    ##Grab the default user as well
    $registryPath = "HKEY_USERS\.DEFAULT\Software\Policies\Microsoft\Windows\WindowsCopilot"
    $propertyName = "TurnOffWindowsCopilot"
    $propertyValue = 1

    # Check if the registry key exists
    if (!(Test-Path $registryPath)) {
        # If the registry key doesn't exist, create it
        New-Item -Path $registryPath -Force | Out-Null
    }

    # Get the property value
    $currentValue = Get-ItemProperty -Path $registryPath -Name $propertyName -ErrorAction SilentlyContinue

    # Check if the property exists and if its value is different from the desired value
    if ($null -eq $currentValue -or $currentValue.$propertyName -ne $propertyValue) {
        # If the property doesn't exist or its value is different, set the property value
        Set-ItemProperty -Path $registryPath -Name $propertyName -Value $propertyValue
    }


    ##Load the default hive from c:\users\Default\NTUSER.dat
    reg load HKU\temphive "c:\users\default\ntuser.dat"
    $registryPath = "registry::hku\temphive\Software\Policies\Microsoft\Windows\WindowsCopilot"
    $propertyName = "TurnOffWindowsCopilot"
    $propertyValue = 1

    # Check if the registry key exists
    if (!(Test-Path $registryPath)) {
        # If the registry key doesn't exist, create it
        [Microsoft.Win32.RegistryKey]$HKUCoPilot = [Microsoft.Win32.Registry]::Users.CreateSubKey("temphive\Software\Policies\Microsoft\Windows\WindowsCopilot", [Microsoft.Win32.RegistryKeyPermissionCheck]::ReadWriteSubTree)
        $HKUCoPilot.SetValue("TurnOffWindowsCopilot", 0x1, [Microsoft.Win32.RegistryValueKind]::DWord)
    }

    $HKUCoPilot.Flush()
    $HKUCoPilot.Close()
    [gc]::Collect()
    [gc]::WaitForPendingFinalizers()
    reg unload HKU\temphive


    write-output "Removed"


    foreach ($sid in $UserSIDs) {
        $registryPath = "Registry::HKU\$sid\SOFTWARE\Policies\Microsoft\Windows\WindowsCopilot"
        $propertyName = "TurnOffWindowsCopilot"
        $propertyValue = 1

        # Check if the registry key exists
        if (!(Test-Path $registryPath)) {
            # If the registry key doesn't exist, create it
            New-Item -Path $registryPath -Force | Out-Null
        }

        # Get the property value
        $currentValue = Get-ItemProperty -Path $registryPath -Name $propertyName -ErrorAction SilentlyContinue

        # Check if the property exists and if its value is different from the desired value
        if ($null -eq $currentValue -or $currentValue.$propertyName -ne $propertyValue) {
            # If the property doesn't exist or its value is different, set the property value
            Set-ItemProperty -Path $registryPath -Name $propertyName -Value $propertyValue
        }
    }
}
############################################################################################################
#                                              Remove Recall                                               #
#                                                                                                          #
############################################################################################################

#Turn off Recall
write-output "Disabling Recall"
$recall = "HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Windows\WindowsAI"
If (!(Test-Path $recall)) {
    New-Item $recall
}
Set-ItemProperty $recall DisableAIDataAnalysis -Value 1


$recalluser = 'HKCU:\SOFTWARE\Policies\Microsoft\Windows\WindowsAI'
If (!(Test-Path $recalluser)) {
    New-Item $recalluser
}
Set-ItemProperty $recalluser DisableAIDataAnalysis -Value 1

##Loop through users and do the same
foreach ($sid in $UserSIDs) {
    $recallusers = "Registry::HKU\$sid\SOFTWARE\Policies\Microsoft\Windows\WindowsAI"
    If (!(Test-Path $recallusers)) {
        New-Item $recallusers
    }
    Set-ItemProperty $recallusers DisableAIDataAnalysis -Value 1
}


############################################################################################################
#                                             Clear Start Menu                                             #
#                                                                                                          #
############################################################################################################
write-output "Clearing Start Menu"
#Delete layout file if it already exists

##Check windows version
$version = Get-CimInstance Win32_OperatingSystem | Select-Object -ExpandProperty Caption
if ($version -like "*Windows 10*") {
    write-output "Windows 10 Detected"
    write-output "Removing Current Layout"
    If (Test-Path C:\Windows\StartLayout.xml)
    {

        Remove-Item C:\Windows\StartLayout.xml

    }
    write-output "Creating Default Layout"
    #Creates the blank layout file

    Write-Output "<LayoutModificationTemplate xmlns:defaultlayout=""http://schemas.microsoft.com/Start/2014/FullDefaultLayout"" xmlns:start=""http://schemas.microsoft.com/Start/2014/StartLayout"" Version=""1"" xmlns=""http://schemas.microsoft.com/Start/2014/LayoutModification"">" >> C:\Windows\StartLayout.xml

    Write-Output " <LayoutOptions StartTileGroupCellWidth=""6"" />" >> C:\Windows\StartLayout.xml

    Write-Output " <DefaultLayoutOverride>" >> C:\Windows\StartLayout.xml

    Write-Output " <StartLayoutCollection>" >> C:\Windows\StartLayout.xml

    Write-Output " <defaultlayout:StartLayout GroupCellWidth=""6"" />" >> C:\Windows\StartLayout.xml

    Write-Output " </StartLayoutCollection>" >> C:\Windows\StartLayout.xml

    Write-Output " </DefaultLayoutOverride>" >> C:\Windows\StartLayout.xml

    Write-Output "</LayoutModificationTemplate>" >> C:\Windows\StartLayout.xml
}
if ($version -like "*Windows 11*") {
    write-output "Windows 11 Detected"
    write-output "Removing Current Layout"
    If (Test-Path "C:\Users\Default\AppData\Local\Microsoft\Windows\Shell\LayoutModification.xml")
    {

        Remove-Item "C:\Users\Default\AppData\Local\Microsoft\Windows\Shell\LayoutModification.xml"

    }

    $blankjson = @'
{
    "pinnedList": [
{ "desktopAppId": "MSEdge" },
{ "packagedAppId": "Microsoft.WindowsStore_8wekyb3d8bbwe!App" },
{ "packagedAppId": "desktopAppId":"Microsoft.Windows.Explorer" }
    ]
}
'@

    $blankjson | Out-File "C:\Users\Default\AppData\Local\Microsoft\Windows\Shell\LayoutModification.xml" -Encoding utf8 -Force
    $intunepath = "HKLM:\SOFTWARE\Microsoft\IntuneManagementExtension\Win32Apps"
    $intunecomplete = @(Get-ChildItem $intunepath).count
    $userpath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList"
    $userprofiles = Get-ChildItem $userpath | ForEach-Object { Get-ItemProperty $_.PSPath }

    $nonAdminLoggedOn = $false
    foreach ($user in $userprofiles) {
        if ($user.PSChildName -ne '.DEFAULT' -and $user.PSChildName -ne 'S-1-5-18' -and $user.PSChildName -ne 'S-1-5-19' -and $user.PSChildName -ne 'S-1-5-20' -and $user.PSChildName -notmatch 'S-1-5-21-\d+-\d+-\d+-500') {
            $nonAdminLoggedOn = $true
            break
        }
    }

    if ($nonAdminLoggedOn -eq $false) {
        MkDir -Path "C:\Users\Default\AppData\Local\Packages\Microsoft.Windows.StartMenuExperienceHost_cw5n1h2txyewy\LocalState" -Force -ErrorAction SilentlyContinue | Out-Null
        $starturl = "https://github.com/andrew-s-taylor/public/raw/main/De-Bloat/start2.bin"
        invoke-webrequest -uri $starturl -outfile "C:\Users\Default\AppData\Local\Packages\Microsoft.Windows.StartMenuExperienceHost_cw5n1h2txyewy\LocalState\Start2.bin"
    }
}


############################################################################################################
#                                              Remove Xbox Gaming                                          #
#                                                                                                          #
############################################################################################################

New-ItemProperty -Path "HKLM:\System\CurrentControlSet\Services\xbgm" -Name "Start" -PropertyType DWORD -Value 4 -Force
Set-Service -Name XblAuthManager -StartupType Disabled
Set-Service -Name XblGameSave -StartupType Disabled
Set-Service -Name XboxGipSvc -StartupType Disabled
Set-Service -Name XboxNetApiSvc -StartupType Disabled
$task = Get-ScheduledTask -TaskName "Microsoft\XblGameSave\XblGameSaveTask" -ErrorAction SilentlyContinue
if ($null -ne $task) {
    Set-ScheduledTask -TaskPath $task.TaskPath -Enabled $false
}

##Check if GamePresenceWriter.exe exists
if (Test-Path "$env:WinDir\System32\GameBarPresenceWriter.exe") {
    write-output "GamePresenceWriter.exe exists"
    #Take-Ownership -Path "$env:WinDir\System32\GameBarPresenceWriter.exe"
    $NewAcl = Get-Acl -Path "$env:WinDir\System32\GameBarPresenceWriter.exe"
    # Set properties
    $identity = "$builtin\Administrators"
    $fileSystemRights = "FullControl"
    $type = "Allow"
    # Create new rule
    $fileSystemAccessRuleArgumentList = $identity, $fileSystemRights, $type
    $fileSystemAccessRule = New-Object -TypeName System.Security.AccessControl.FileSystemAccessRule -ArgumentList $fileSystemAccessRuleArgumentList
    # Apply new rule
    $NewAcl.SetAccessRule($fileSystemAccessRule)
    Set-Acl -Path "$env:WinDir\System32\GameBarPresenceWriter.exe" -AclObject $NewAcl
    Stop-Process -Name "GameBarPresenceWriter.exe" -Force
    Remove-Item "$env:WinDir\System32\GameBarPresenceWriter.exe" -Force -Confirm:$false

}
else {
    write-output "GamePresenceWriter.exe does not exist"
}

New-ItemProperty -Path "HKLM:\Software\Policies\Microsoft\Windows\GameDVR" -Name "AllowgameDVR" -PropertyType DWORD -Value 0 -Force
New-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer" -Name "SettingsPageVisibility" -PropertyType String -Value "hide:gaming-gamebar;gaming-gamedvr;gaming-broadcasting;gaming-gamemode;gaming-xboxnetworking" -Force
Remove-Item C:\Windows\Temp\SetACL.exe -recurse

############################################################################################################
#                                        Disable Edge Surf Game                                            #
#                                                                                                          #
############################################################################################################
$surf = "HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Edge"
If (!(Test-Path $surf)) {
    New-Item $surf
}
New-ItemProperty -Path $surf -Name 'AllowSurfGame' -Value 0 -PropertyType DWord

############################################################################################################
#                                       Grab all Uninstall Strings                                         #
#                                                                                                          #
############################################################################################################


write-output "Checking 32-bit System Registry"
##Search for 32-bit versions and list them
$allstring = @()
$path1 = "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
#Loop Through the apps if name has Adobe and NOT reader
$32apps = Get-ChildItem -Path $path1 | Get-ItemProperty | Select-Object -Property DisplayName, UninstallString

foreach ($32app in $32apps) {
    #Get uninstall string
    $string1 = $32app.uninstallstring
    #Check if it's an MSI install
    if ($string1 -match "^msiexec*") {
        #MSI install, replace the I with an X and make it quiet
        $string2 = $string1 + " /quiet /norestart"
        $string2 = $string2 -replace "/I", "/X "
        #Create custom object with name and string
        $allstring += New-Object -TypeName PSObject -Property @{
            Name   = $32app.DisplayName
            String = $string2
        }
    }
    else {
        #Exe installer, run straight path
        $string2 = $string1
        $allstring += New-Object -TypeName PSObject -Property @{
            Name   = $32app.DisplayName
            String = $string2
        }
    }

}
write-output "32-bit check complete"
write-output "Checking 64-bit System registry"
##Search for 64-bit versions and list them

$path2 = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
#Loop Through the apps if name has Adobe and NOT reader
$64apps = Get-ChildItem -Path $path2 | Get-ItemProperty | Select-Object -Property DisplayName, UninstallString

foreach ($64app in $64apps) {
    #Get uninstall string
    $string1 = $64app.uninstallstring
    #Check if it's an MSI install
    if ($string1 -match "^msiexec*") {
        #MSI install, replace the I with an X and make it quiet
        $string2 = $string1 + " /quiet /norestart"
        $string2 = $string2 -replace "/I", "/X "
        #Uninstall with string2 params
        $allstring += New-Object -TypeName PSObject -Property @{
            Name   = $64app.DisplayName
            String = $string2
        }
    }
    else {
        #Exe installer, run straight path
        $string2 = $string1
        $allstring += New-Object -TypeName PSObject -Property @{
            Name   = $64app.DisplayName
            String = $string2
        }
    }

}

write-output "64-bit checks complete"

##USER
write-output "Checking 32-bit User Registry"
##Search for 32-bit versions and list them
$path1 = "HKCU:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
##Check if path exists
if (Test-Path $path1) {
    #Loop Through the apps if name has Adobe and NOT reader
    $32apps = Get-ChildItem -Path $path1 | Get-ItemProperty | Select-Object -Property DisplayName, UninstallString

    foreach ($32app in $32apps) {
        #Get uninstall string
        $string1 = $32app.uninstallstring
        #Check if it's an MSI install
        if ($string1 -match "^msiexec*") {
            #MSI install, replace the I with an X and make it quiet
            $string2 = $string1 + " /quiet /norestart"
            $string2 = $string2 -replace "/I", "/X "
            #Create custom object with name and string
            $allstring += New-Object -TypeName PSObject -Property @{
                Name   = $32app.DisplayName
                String = $string2
            }
        }
        else {
            #Exe installer, run straight path
            $string2 = $string1
            $allstring += New-Object -TypeName PSObject -Property @{
                Name   = $32app.DisplayName
                String = $string2
            }
        }
    }
}

write-output "32-bit check complete"
write-output "Checking 64-bit Use registry"
##Search for 64-bit versions and list them

$path2 = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
#Loop Through the apps if name has Adobe and NOT reader
$64apps = Get-ChildItem -Path $path2 | Get-ItemProperty | Select-Object -Property DisplayName, UninstallString

foreach ($64app in $64apps) {
    #Get uninstall string
    $string1 = $64app.uninstallstring
    #Check if it's an MSI install
    if ($string1 -match "^msiexec*") {
        #MSI install, replace the I with an X and make it quiet
        $string2 = $string1 + " /quiet /norestart"
        $string2 = $string2 -replace "/I", "/X "
        #Uninstall with string2 params
        $allstring += New-Object -TypeName PSObject -Property @{
            Name   = $64app.DisplayName
            String = $string2
        }
    }
    else {
        #Exe installer, run straight path
        $string2 = $string1
        $allstring += New-Object -TypeName PSObject -Property @{
            Name   = $64app.DisplayName
            String = $string2
        }
    }

}


function UninstallAppFull {

    param (
        [string]$appName
    )

    # Get a list of installed applications from Programs and Features
    $installedApps = Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*,
    HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* |
    Where-Object { $null -ne $_.DisplayName } |
    Select-Object DisplayName, UninstallString

    $userInstalledApps = Get-ItemProperty HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* |
    Where-Object { $null -ne $_.DisplayName } |
    Select-Object DisplayName, UninstallString

    $allInstalledApps = $installedApps + $userInstalledApps | Where-Object { $_.DisplayName -eq "$appName" }

    # Loop through the list of installed applications and uninstall them

    foreach ($app in $allInstalledApps) {
        $uninstallString = $app.UninstallString
        $displayName = $app.DisplayName
        if ($uninstallString -match "^msiexec*") {
            #MSI install, replace the I with an X and make it quiet
            $string2 = $uninstallString + " /quiet /norestart"
            $string2 = $string2 -replace "/I", "/X "
        }
        else {
            #Exe installer, run straight path
            $string2 = $uninstallString
        }
        write-output "Uninstalling: $displayName"
        Start-Process $string2
        write-output "Uninstalled: $displayName" -ForegroundColor Green
    }
}


############################################################################################################
#                                        Remove Manufacturer Bloat                                         #
#                                                                                                          #
############################################################################################################
##Check Manufacturer
write-output "Detecting Manufacturer"
$details = Get-CimInstance -ClassName Win32_ComputerSystem
$manufacturer = $details.Manufacturer

if ($manufacturer -like "*HP*") {
    write-output "HP detected"
    #Remove HP bloat

    ##HP Specific
    $UninstallPrograms = @(
        "Poly Lens"
        "HP Client Security Manager"
        "HP Notifications"
        "HP Security Update Service"
        "HP System Default Settings"
        "HP Wolf Security"
        "HP Wolf Security Application Support for Sure Sense"
        "HP Wolf Security Application Support for Windows"
        "AD2F1837.HPPCHardwareDiagnosticsWindows"
        "AD2F1837.HPPowerManager"
        "AD2F1837.HPPrivacySettings"
        "AD2F1837.HPQuickDrop"
        "AD2F1837.HPSupportAssistant"
        "AD2F1837.HPSystemInformation"
        "AD2F1837.myHP"
        "RealtekSemiconductorCorp.HPAudioControl",
        "HP Sure Recover",
        "HP Sure Run Module"
        "RealtekSemiconductorCorp.HPAudioControl_2.39.280.0_x64__dt26b99r8h8gj"
        "HP Wolf Security - Console"
        "HP Wolf Security Application Support for Chrome 122.0.6261.139"
        "Windows Driver Package - HP Inc. sselam_4_4_2_453 AntiVirus  (11/01/2022 4.4.2.453)"
    )

    $UninstallPrograms = $UninstallPrograms | Where-Object { $appstoignore -notcontains $_ }

    #$HPidentifier = "AD2F1837"

    #$ProvisionedPackages = Get-AppxProvisionedPackage -Online | Where-Object {(($UninstallPrograms -contains $_.DisplayName) -or (($_.DisplayName -like "*$HPidentifier"))-and ($_.DisplayName -notin $WhitelistedApps))}

    #$InstalledPackages = Get-AppxPackage -AllUsers | Where-Object {(($UninstallPrograms -contains $_.Name) -or (($_.Name -like "^$HPidentifier"))-and ($_.Name -notin $WhitelistedApps))}

    $InstalledPrograms = $allstring | Where-Object { $UninstallPrograms -contains $_.Name }
    foreach ($app in $UninstallPrograms) {

        if (Get-AppxProvisionedPackage -Online | Where-Object DisplayName -like $app -ErrorAction SilentlyContinue) {
            Get-AppxProvisionedPackage -Online | Where-Object DisplayName -like $app | Remove-AppxProvisionedPackage -Online
            write-output "Removed provisioned package for $app."
        }
        else {
            write-output "Provisioned package for $app not found."
        }

        if (Get-AppxPackage -allusers -Name $app -ErrorAction SilentlyContinue) {
            Get-AppxPackage -allusers -Name $app | Remove-AppxPackage -AllUsers
            write-output "Removed $app."
        }
        else {
            write-output "$app not found."
        }

        UninstallAppFull -appName $app


    }

    ##Belt and braces, remove via CIM too
    foreach ($program in $UninstallPrograms) {
        Get-CimInstance -Classname Win32_Product | Where-Object Name -Match $program | Invoke-CimMethod -MethodName UnInstall
    }


    #Remove HP Documentation if it exists
    if (test-path -Path "C:\Program Files\HP\Documentation\Doc_uninstall.cmd") {
        Start-Process -FilePath "C:\Program Files\HP\Documentation\Doc_uninstall.cmd" -Wait -passthru -NoNewWindow
    }

    ##Remove HP Connect Optimizer if setup.exe exists
    if (test-path -Path 'C:\Program Files (x86)\InstallShield Installation Information\{6468C4A5-E47E-405F-B675-A70A70983EA6}\setup.exe') {
        invoke-webrequest -uri "https://raw.githubusercontent.com/andrew-s-taylor/public/main/De-Bloat/HPConnOpt.iss" -outfile "C:\Windows\Temp\HPConnOpt.iss"

        &'C:\Program Files (x86)\InstallShield Installation Information\{6468C4A5-E47E-405F-B675-A70A70983EA6}\setup.exe' @('-s', '-f1C:\Windows\Temp\HPConnOpt.iss')
    }
##Remove HP Data Science Stack Manager
if (test-path -Path 'C:\Program Files\HP\Z By HP Data Science Stack Manager\Uninstall Z by HP Data Science Stack Manager.exe') {
    &'C:\Program Files\HP\Z By HP Data Science Stack Manager\Uninstall Z by HP Data Science Stack Manager.exe' @('/allusers', '/S')
}


    ##Remove other crap
    if (Test-Path -Path "C:\Program Files (x86)\HP\Shared" -PathType Container) { Remove-Item -Path "C:\Program Files (x86)\HP\Shared" -Recurse -Force }
    if (Test-Path -Path "C:\Program Files (x86)\Online Services" -PathType Container) { Remove-Item -Path "C:\Program Files (x86)\Online Services" -Recurse -Force }
    if (Test-Path -Path "C:\ProgramData\HP\TCO" -PathType Container) { Remove-Item -Path "C:\ProgramData\HP\TCO" -Recurse -Force }
    if (Test-Path -Path "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Amazon.com.lnk" -PathType Leaf) { Remove-Item -Path "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Amazon.com.lnk" -Force }
    if (Test-Path -Path "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Angebote.lnk" -PathType Leaf) { Remove-Item -Path "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Angebote.lnk" -Force }
    if (Test-Path -Path "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\TCO Certified.lnk" -PathType Leaf) { Remove-Item -Path "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\TCO Certified.lnk" -Force }
    if (Test-Path -Path "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Booking.com.lnk" -PathType Leaf) { Remove-Item -Path "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Booking.com.lnk" -Force }
    if (Test-Path -Path "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Adobe offers.lnk" -PathType Leaf) { Remove-Item -Path "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Adobe offers.lnk" -Force }


    ##Remove Wolf Security
    wmic product where "name='HP Wolf Security'" call uninstall
    wmic product where "name='HP Wolf Security - Console'" call uninstall
    wmic product where "name='HP Security Update Service'" call uninstall

    write-output "Removed HP bloat"
}



if ($manufacturer -like "*Dell*") {
    write-output "Dell detected"
    #Remove Dell bloat

    ##Dell

    $UninstallPrograms = @(
        "Dell Optimizer"
        "Dell Power Manager"
        "DellOptimizerUI"
        "Dell SupportAssist OS Recovery"
        "Dell SupportAssist"
        "Dell Optimizer Service"
        "Dell Optimizer Core"
        "DellInc.PartnerPromo"
        "DellInc.DellOptimizer"
        "DellInc.DellCommandUpdate"
        "DellInc.DellPowerManager"
        "DellInc.DellDigitalDelivery"
        "DellInc.DellSupportAssistforPCs"
        "DellInc.PartnerPromo"
        "Dell Command | Update"
        "Dell Command | Update for Windows Universal"
        "Dell Command | Update for Windows 10"
        "Dell Command | Power Manager"
        "Dell Digital Delivery Service"
        "Dell Digital Delivery"
        "Dell Peripheral Manager"
        "Dell Power Manager Service"
        "Dell SupportAssist Remediation"
        "SupportAssist Recovery Assistant"
        "Dell SupportAssist OS Recovery Plugin for Dell Update"
        "Dell SupportAssistAgent"
        "Dell Update - SupportAssist Update Plugin"
        "Dell Core Services"
        "Dell Pair"
        "Dell Display Manager 2.0"
        "Dell Display Manager 2.1"
        "Dell Display Manager 2.2"
        "Dell SupportAssist Remediation"
        "Dell Update - SupportAssist Update Plugin"
        "DellInc.PartnerPromo"
    )



    $UninstallPrograms = $UninstallPrograms | Where-Object { $appstoignore -notcontains $_ }


    foreach ($app in $UninstallPrograms) {

        if (Get-AppxProvisionedPackage -Online | Where-Object DisplayName -like $app -ErrorAction SilentlyContinue) {
            Get-AppxProvisionedPackage -Online | Where-Object DisplayName -like $app | Remove-AppxProvisionedPackage -Online
            write-output "Removed provisioned package for $app."
        }
        else {
            write-output "Provisioned package for $app not found."
        }

        if (Get-AppxPackage -allusers -Name $app -ErrorAction SilentlyContinue) {
            Get-AppxPackage -allusers -Name $app | Remove-AppxPackage -AllUsers
            write-output "Removed $app."
        }
        else {
            write-output "$app not found."
        }

        UninstallAppFull -appName $app



    }

    ##Belt and braces, remove via CIM too
    foreach ($program in $UninstallPrograms) {
        write-output "Removing $program"
        Get-CimInstance -Query "SELECT * FROM Win32_Product WHERE name = '$program'" | Invoke-CimMethod -MethodName Uninstall
    }

    ##Manual Removals

    ##Dell Optimizer
    $dellSA = Get-ChildItem -Path HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall, HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall | Get-ItemProperty | Where-Object { $_.DisplayName -like "Dell*Optimizer*Core" } | Select-Object -Property UninstallString

    ForEach ($sa in $dellSA) {
        If ($sa.UninstallString) {
            try {
                cmd.exe /c $sa.UninstallString -silent
            }
            catch {
                Write-Warning "Failed to uninstall Dell Optimizer"
            }
        }
    }


    ##Dell Dell SupportAssist Remediation
    $dellSA = Get-ChildItem -Path HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall, HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall | Get-ItemProperty | Where-Object { $_.DisplayName -match "Dell SupportAssist Remediation" } | Select-Object -Property QuietUninstallString

    ForEach ($sa in $dellSA) {
        If ($sa.QuietUninstallString) {
            try {
                cmd.exe /c $sa.QuietUninstallString
            }
            catch {
                Write-Warning "Failed to uninstall Dell Support Assist Remediation"
            }
        }
    }

    ##Dell Dell SupportAssist OS Recovery Plugin for Dell Update
    $dellSA = Get-ChildItem -Path HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall, HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall | Get-ItemProperty | Where-Object { $_.DisplayName -match "Dell SupportAssist OS Recovery Plugin for Dell Update" } | Select-Object -Property QuietUninstallString

    ForEach ($sa in $dellSA) {
        If ($sa.QuietUninstallString) {
            try {
                cmd.exe /c $sa.QuietUninstallString
            }
            catch {
                Write-Warning "Failed to uninstall Dell Support Assist Remediation"
            }
        }
    }



    ##Dell Display Manager
    $dellSA = Get-ChildItem -Path HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall, HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall | Get-ItemProperty | Where-Object { $_.DisplayName -like "Dell*Display*Manager*" } | Select-Object -Property UninstallString

    ForEach ($sa in $dellSA) {
        If ($sa.UninstallString) {
            try {
                cmd.exe /c $sa.UninstallString /S
            }
            catch {
                Write-Warning "Failed to uninstall Dell Optimizer"
            }
        }
    }

    ##Dell Peripheral Manager

    try {
        start-process c:\windows\system32\cmd.exe '/c "C:\Program Files\Dell\Dell Peripheral Manager\Uninstall.exe" /S'
    }
    catch {
        Write-Warning "Failed to uninstall Dell Optimizer"
    }


    ##Dell Pair

    try {
        start-process c:\windows\system32\cmd.exe '/c "C:\Program Files\Dell\Dell Pair\Uninstall.exe" /S'
    }
    catch {
        Write-Warning "Failed to uninstall Dell Optimizer"
    }

}


if ($manufacturer -like "Lenovo") {
    write-output "Lenovo detected"

    #Remove HP bloat

    ##Lenovo Specific
    # Function to uninstall applications with .exe uninstall strings

    function UninstallApp {

        param (
            [string]$appName
        )

        # Get a list of installed applications from Programs and Features
        $installedApps = Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*,
        HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* |
        Where-Object { $_.DisplayName -like "*$appName*" }

        # Loop through the list of installed applications and uninstall them

        foreach ($app in $installedApps) {
            $uninstallString = $app.UninstallString
            $displayName = $app.DisplayName
            write-output "Uninstalling: $displayName"
            Start-Process $uninstallString -ArgumentList "/VERYSILENT" -Wait
            write-output "Uninstalled: $displayName" -ForegroundColor Green
        }
    }

    ##Stop Running Processes

    $processnames = @(
        "SmartAppearanceSVC.exe"
        "UDClientService.exe"
        "ModuleCoreService.exe"
        "ProtectedModuleHost.exe"
        "*lenovo*"
        "FaceBeautify.exe"
        "McCSPServiceHost.exe"
        "mcapexe.exe"
        "MfeAVSvc.exe"
        "mcshield.exe"
        "Ammbkproc.exe"
        "AIMeetingManager.exe"
        "DADUpdater.exe"
        "CommercialVantage.exe"
    )

    foreach ($process in $processnames) {
        write-output "Stopping Process $process"
        Get-Process -Name $process | Stop-Process -Force
        write-output "Process $process Stopped"
    }

    $UninstallPrograms = @(
        "E046963F.AIMeetingManager"
        "E0469640.SmartAppearance"
        "MirametrixInc.GlancebyMirametrix"
        "E046963F.LenovoCompanion"
        "E0469640.LenovoUtility"
        "E0469640.LenovoSmartCommunication"
        "E046963F.LenovoSettingsforEnterprise"
        "E046963F.cameraSettings"
        "4505Fortemedia.FMAPOControl2_2.1.37.0_x64__4pejv7q2gmsnr"
        "ElevocTechnologyCo.Ltd.SmartMicrophoneSettings_1.1.49.0_x64__ttaqwwhyt5s6t"
        "Lenovo User Guide"
    )


    $UninstallPrograms = $UninstallPrograms | Where-Object { $appstoignore -notcontains $_ }



    $InstalledPrograms = $allstring | Where-Object { (($_.Name -in $UninstallPrograms)) }


    foreach ($app in $UninstallPrograms) {

        if (Get-AppxProvisionedPackage -Online | Where-Object DisplayName -like $app -ErrorAction SilentlyContinue) {
            Get-AppxProvisionedPackage -Online | Where-Object DisplayName -like $app | Remove-AppxProvisionedPackage -Online
            write-output "Removed provisioned package for $app."
        }
        else {
            write-output "Provisioned package for $app not found."
        }

        if (Get-AppxPackage -allusers -Name $app -ErrorAction SilentlyContinue) {
            Get-AppxPackage -allusers -Name $app | Remove-AppxPackage -AllUsers
            write-output "Removed $app."
        }
        else {
            write-output "$app not found."
        }

        UninstallAppFull -appName $app


    }


    ##Belt and braces, remove via CIM too
    foreach ($program in $UninstallPrograms) {
        Get-CimInstance -Classname Win32_Product | Where-Object Name -Match $program | Invoke-CimMethod -MethodName UnInstall
    }

    # Get Lenovo Vantage service uninstall string to uninstall service
    $lvs = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*", "HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*" | Where-Object DisplayName -eq "Lenovo Vantage Service"
    if (!([string]::IsNullOrEmpty($lvs.QuietUninstallString))) {
        $uninstall = "cmd /c " + $lvs.QuietUninstallString
        write-output $uninstall
        Invoke-Expression $uninstall
    }

    # Uninstall Lenovo Smart
    UninstallApp -appName "Lenovo Smart"

    # Uninstall Ai Meeting Manager Service
    UninstallApp -appName "Ai Meeting Manager"

    # Uninstall ImController service
    ##Check if exists
    $path = "c:\windows\system32\ImController.InfInstaller.exe"
    if (Test-Path $path) {
        write-output "ImController.InfInstaller.exe exists"
        $uninstall = "cmd /c " + $path + " -uninstall"
        write-output $uninstall
        Invoke-Expression $uninstall
    }
    else {
        write-output "ImController.InfInstaller.exe does not exist"
    }
    ##Invoke-Expression -Command 'cmd.exe /c "c:\windows\system32\ImController.InfInstaller.exe" -uninstall'

    # Remove vantage associated registry keys
    Remove-Item 'HKLM:\SOFTWARE\Policies\Lenovo\E046963F.LenovoCompanion_k1h2ywk1493x8' -Recurse -ErrorAction SilentlyContinue
    Remove-Item 'HKLM:\SOFTWARE\Policies\Lenovo\ImController' -Recurse -ErrorAction SilentlyContinue
    Remove-Item 'HKLM:\SOFTWARE\Policies\Lenovo\Lenovo Vantage' -Recurse -ErrorAction SilentlyContinue
    #Remove-Item 'HKLM:\SOFTWARE\Policies\Lenovo\Commercial Vantage' -Recurse -ErrorAction SilentlyContinue

    # Uninstall AI Meeting Manager Service
    $path = 'C:\Program Files\Lenovo\Ai Meeting Manager Service\unins000.exe'
    $params = "/SILENT"
    if (test-path -Path $path) {
        Start-Process -FilePath $path -ArgumentList $params -Wait
    }
    # Uninstall Lenovo Vantage
    $pathname = (Get-ChildItem -Path "C:\Program Files (x86)\Lenovo\VantageService").name
    $path = "C:\Program Files (x86)\Lenovo\VantageService\$pathname\Uninstall.exe"
    $params = '/SILENT'
    if (test-path -Path $path) {
        Start-Process -FilePath $path -ArgumentList $params -Wait
    }

    ##Uninstall Smart Appearance
    $path = 'C:\Program Files\Lenovo\Lenovo Smart Appearance Components\unins000.exe'
    $params = '/SILENT'
    if (test-path -Path $path) {
        try {
            Start-Process -FilePath $path -ArgumentList $params -Wait
        }
        catch {
            Write-Warning "Failed to start the process"
        }
    }
    $lenovowelcome = "c:\program files (x86)\lenovo\lenovowelcome\x86"
    if (Test-Path $lenovowelcome) {
        # Remove Lenovo Now
        Set-Location "c:\program files (x86)\lenovo\lenovowelcome\x86"

        # Update $PSScriptRoot with the new working directory
        $PSScriptRoot = (Get-Item -Path ".\").FullName
        try {
            invoke-expression -command .\uninstall.ps1 -ErrorAction SilentlyContinue
        }
        catch {
            write-output "Failed to execute uninstall.ps1"
        }

        write-output "All applications and associated Lenovo components have been uninstalled." -ForegroundColor Green
    }

    $lenovonow = "c:\program files (x86)\lenovo\LenovoNow\x86"
    if (Test-Path $lenovonow) {
        # Remove Lenovo Now
        Set-Location "c:\program files (x86)\lenovo\LenovoNow\x86"

        # Update $PSScriptRoot with the new working directory
        $PSScriptRoot = (Get-Item -Path ".\").FullName
        try {
            invoke-expression -command .\uninstall.ps1 -ErrorAction SilentlyContinue
        }
        catch {
            write-output "Failed to execute uninstall.ps1"
        }

        write-output "All applications and associated Lenovo components have been uninstalled." -ForegroundColor Green
    }


    $filename = "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\User Guide.lnk"

    if (Test-Path $filename) {
        Remove-Item -Path $filename -Force
    }

    ##Camera fix for Lenovo E14
    $model = Get-CimInstance -ClassName Win32_ComputerSystem | Select-Object -ExpandProperty Model
    if ($model -eq "21E30001MY") {
        $keypath = "HKLM:\SOFTWARE\\Microsoft\Windows Media Foundation\Platform"
        $keyname = "EnableFrameServerMode"
        $value = 0
        if (!(Test-Path $keypath)) {
            New-Item -Path $keypath -Force
        }
        Set-ItemProperty -Path $keypath -Name $keyname -Value $value -Type DWord -Force

        $keypath2 = "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows Media Foundation\Platform"
        if (!(Test-Path $keypath2)) {
            New-Item -Path $keypath2 -Force
        }
        Set-ItemProperty -Path $keypath2 -Name $keyname -Value $value -Type DWord -Force
    }


        ##Remove Lenovo theme and background image
        $registryPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes"

        # Check and remove ThemeName if it exists
        if (Get-ItemProperty -Path $registryPath -Name "ThemeName" -ErrorAction SilentlyContinue) {
            Remove-ItemProperty -Path $registryPath -Name "ThemeName"
        }
    
        # Check and remove DesktopBackground if it exists
        if (Get-ItemProperty -Path $registryPath -Name "DesktopBackground" -ErrorAction SilentlyContinue) {
            Remove-ItemProperty -Path $registryPath -Name "DesktopBackground"
        }

}


##Remove bookmarks

##Enumerate all users
$users = Get-ChildItem -Path "C:\Users" -Directory
foreach ($user in $users) {
    $userpath = $user.FullName
    $bookmarks = "C:\Users\$userpath\AppData\Local\Microsoft\Edge\User Data\Default\Bookmarks"
    ##Remove any files if they exist
    foreach ($bookmark in $bookmarks) {
        if (Test-Path -Path $bookmark) {
            Remove-Item -Path $bookmark -Force
        }
    }
}

############################################################################################################
#                                           REMOVE WINDOWS.OLD                                             #
#                                                                                                          #
############################################################################################################
#Windows.old

write-output "Checking for Windows.old folder"
$windowsOldPath = "C:\Windows.old"

if (Test-Path $windowsOldPath) {
    write-output "Windows.old folder found. Attempting to remove..."
    try {
        Remove-Item -Path $windowsOldPath -Recurse -Force
        write-output "Windows.old folder successfully removed"
    }
    catch {
        write-output "Could not remove Windows.old folder using standard method. Attempting cleanup with Disk Cleanup utility..."
        # Alternative method using cleanmgr
        Start-Process -Wait cleanmgr.exe -ArgumentList "/sagerun:1" -NoNewWindow
        write-output "Disk Cleanup completed"
    }
}
else {
    write-output "Windows.old folder not found. No action needed."
}


############################################################################################################
#                                        Remove Any other installed crap                                   #
#                                                                                                          #
############################################################################################################

#McAfee

write-output "Detecting McAfee"
$mcafeeinstalled = "false"
$InstalledSoftware = Get-ChildItem "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall"
foreach ($obj in $InstalledSoftware) {
    $name = $obj.GetValue('DisplayName')
    if ($name -like "*McAfee*") {
        $mcafeeinstalled = "true"
    }
}

$InstalledSoftware32 = Get-ChildItem "HKLM:\Software\WOW6432NODE\Microsoft\Windows\CurrentVersion\Uninstall"
foreach ($obj32 in $InstalledSoftware32) {
    $name32 = $obj32.GetValue('DisplayName')
    if ($name32 -like "*McAfee*") {
        $mcafeeinstalled = "true"
    }
}

if ($mcafeeinstalled -eq "true") {
    write-output "McAfee detected"
    #Remove McAfee bloat
    ##McAfee
    ### Download McAfee Consumer Product Removal Tool ###
    write-output "Downloading McAfee Removal Tool"
    # Download Source
    $URL = 'https://github.com/andrew-s-taylor/public/raw/main/De-Bloat/mcafeeclean.zip'

    # Set Save Directory
    $destination = 'C:\ProgramData\Debloat\mcafee.zip'

    #Download the file
    Invoke-WebRequest -Uri $URL -OutFile $destination -Method Get

    Expand-Archive $destination -DestinationPath "C:\ProgramData\Debloat" -Force

    write-output "Removing McAfee"
    # Automate Removal and kill services
    start-process "C:\ProgramData\Debloat\Mccleanup.exe" -ArgumentList "-p StopServices,MFSY,PEF,MXD,CSP,Sustainability,MOCP,MFP,APPSTATS,Auth,EMproxy,FWdiver,HW,MAS,MAT,MBK,MCPR,McProxy,McSvcHost,VUL,MHN,MNA,MOBK,MPFP,MPFPCU,MPS,SHRED,MPSCU,MQC,MQCCU,MSAD,MSHR,MSK,MSKCU,MWL,NMC,RedirSvc,VS,REMEDIATION,MSC,YAP,TRUEKEY,LAM,PCB,Symlink,SafeConnect,MGS,WMIRemover,RESIDUE -v -s"
    write-output "McAfee Removal Tool has been run"

    ###New MCCleanup
    ### Download McAfee Consumer Product Removal Tool ###
    write-output "Downloading McAfee Removal Tool"
    # Download Source
    $URL = 'https://github.com/andrew-s-taylor/public/raw/main/De-Bloat/mccleanup.zip'

    # Set Save Directory
    $destination = 'C:\ProgramData\Debloat\mcafeenew.zip'

    #Download the file
    Invoke-WebRequest -Uri $URL -OutFile $destination -Method Get

    New-Item -Path "C:\ProgramData\Debloat\mcnew" -ItemType Directory
    Expand-Archive $destination -DestinationPath "C:\ProgramData\Debloat\mcnew" -Force

    write-output "Removing McAfee"
    # Automate Removal and kill services
    start-process "C:\ProgramData\Debloat\mcnew\Mccleanup.exe" -ArgumentList "-p StopServices,MFSY,PEF,MXD,CSP,Sustainability,MOCP,MFP,APPSTATS,Auth,EMproxy,FWdiver,HW,MAS,MAT,MBK,MCPR,McProxy,McSvcHost,VUL,MHN,MNA,MOBK,MPFP,MPFPCU,MPS,SHRED,MPSCU,MQC,MQCCU,MSAD,MSHR,MSK,MSKCU,MWL,NMC,RedirSvc,VS,REMEDIATION,MSC,YAP,TRUEKEY,LAM,PCB,Symlink,SafeConnect,MGS,WMIRemover,RESIDUE -v -s"
    write-output "McAfee Removal Tool has been run"

    $InstalledPrograms = $allstring | Where-Object { ($_.Name -like "*McAfee*") }
    $InstalledPrograms | ForEach-Object {

        write-output "Attempting to uninstall: [$($_.Name)]..."
        $uninstallcommand = $_.String

        Try {
            if ($uninstallcommand -match "^msiexec*") {
                #Remove msiexec as we need to split for the uninstall
                $uninstallcommand = $uninstallcommand -replace "msiexec.exe", ""
                $uninstallcommand = $uninstallcommand + " /quiet /norestart"
                $uninstallcommand = $uninstallcommand -replace "/I", "/X "
                #Uninstall with string2 params
                Start-Process 'msiexec.exe' -ArgumentList $uninstallcommand -NoNewWindow -Wait
            }
            else {
                #Exe installer, run straight path
                $string2 = $uninstallcommand
                start-process $string2
            }
            #$A = Start-Process -FilePath $uninstallcommand -Wait -passthru -NoNewWindow;$a.ExitCode
            #$Null = $_ | Uninstall-Package -AllVersions -Force -ErrorAction Stop
            write-output "Successfully uninstalled: [$($_.Name)]"
        }
        Catch { Write-Warning -Message "Failed to uninstall: [$($_.Name)]" }
    }

    ##Remove Safeconnect
    $safeconnects = Get-ChildItem -Path HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall, HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall | Get-ItemProperty | Where-Object { $_.DisplayName -match "McAfee Safe Connect" } | Select-Object -Property UninstallString

    ForEach ($sc in $safeconnects) {
        If ($sc.UninstallString) {
            cmd.exe /c $sc.UninstallString /quiet /norestart
        }
    }

    ##
    ##remove some extra leftover Mcafee items from StartMenu-AllApps and uninstall registry keys
    ##
    if (Test-Path -Path "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\McAfee") {
        Remove-Item -Path "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\McAfee" -Recurse -Force
    }
    if (Test-Path -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\McAfee.WPS") {
        Remove-Item -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\McAfee.WPS" -Recurse -Force
    }
    #Interesting emough, this producese an error, but still deletes the package anyway
    get-appxprovisionedpackage -online | sort-object displayname | format-table displayname, packagename
    get-appxpackage -allusers | sort-object name | format-table name, packagefullname
    Get-AppxProvisionedPackage -Online | Where-Object DisplayName -eq "McAfeeWPSSparsePackage" | Remove-AppxProvisionedPackage -Online -AllUsers
}
    
    # Disable advertising ID
    Write-Host "Disabling advertising ID..." -ForegroundColor Yellow
    If (!(Test-Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\AdvertisingInfo")) {
        New-Item -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\AdvertisingInfo" -Force | Out-Null
    }
    Set-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\AdvertisingInfo" -Name "Enabled" -Value 0

    # Disable automatic installation of suggested apps
    Write-Host "Disabling automatic installation of suggested apps..." -ForegroundColor Yellow
    If (!(Test-Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\ContentDeliveryManager")) {
        New-Item -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\ContentDeliveryManager" -Force | Out-Null
    }
    Set-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\ContentDeliveryManager" -Name "SilentInstalledAppsEnabled" -Value 0

    # Disable tailored experiences
    Write-Host "Disabling tailored experiences..." -ForegroundColor Yellow
    If (!(Test-Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Privacy")) {
        New-Item -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Privacy" -Force | Out-Null
    }
    Set-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Privacy" -Name "TailoredExperiencesWithDiagnosticDataEnabled" -Value 0

    # Disable Xbox Game Bar
    Write-Host "Disabling Xbox Game Bar..." -ForegroundColor Yellow
    Set-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\GameDVR" -Name "AppCaptureEnabled" -Value 0
    Set-ItemProperty -Path "HKCU:\System\GameConfigStore" -Name "GameDVR_Enabled" -Value 0

    # Disable Xbox Game Bar tips
    Write-Host "Disabling Xbox Game Bar tips..." -ForegroundColor Yellow
    If (!(Test-Path "HKCU:\SOFTWARE\Microsoft\GameBar")) {
        New-Item -Path "HKCU:\SOFTWARE\Microsoft\GameBar" -Force | Out-Null
    }
    Set-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\GameBar" -Name "ShowStartupPanel" -Value 0

##Make sure Intune hasn't installed anything so we don't remove installed apps

$intunepath = "HKLM:\SOFTWARE\Microsoft\IntuneManagementExtension\Win32Apps"
$intunecomplete = @(Get-ChildItem $intunepath).count
$userpath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList"
$userprofiles = Get-ChildItem $userpath | Get-ItemProperty

$nonAdminLoggedOn = $false
foreach ($user in $userprofiles) {
    # Exclude default, system, and network service profiles, and the Administrator profile
    if ($user.PSChildName -notin '.DEFAULT', 'S-1-5-18', 'S-1-5-19', 'S-1-5-20' -and $user.PSChildName -notmatch 'S-1-5-21-\d+-\d+-\d+-500') {
        $nonAdminLoggedOn = $true
        break
    }
}
$TypeDef = @"
 
using System;
using System.Text;
using System.Collections.Generic;
using System.Runtime.InteropServices;
 
namespace Api
{
 public class Kernel32
 {
   [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
   public static extern int OOBEComplete(ref int bIsOOBEComplete);
 }
}
"@
 
Add-Type -TypeDefinition $TypeDef -Language CSharp
 
$IsOOBEComplete = $false
$hr = [Api.Kernel32]::OOBEComplete([ref] $IsOOBEComplete)
 

if ($IsOOBEComplete -eq 0) {

    write-output "Still in OOBE, continue"
    ##Apps to remove - NOTE: Chrome has an unusual uninstall so sort on it's own
    $blacklistapps = @(

    )


    foreach ($blacklist in $blacklistapps) {

        UninstallAppFull -appName $blacklist

    }


    ##Remove Chrome
   # $chrome32path = "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\Google Chrome"

   # if ($null -ne $chrome32path) {

    #    $versions = (Get-ItemProperty -path 'HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\Google Chrome').version
     #   ForEach ($version in $versions) {
      #      write-output "Found Chrome version $version"
       #     $directory = ${env:ProgramFiles(x86)}
        #    write-output "Removing Chrome"
         #   Start-Process "$directory\Google\Chrome\Application\$version\Installer\setup.exe" -argumentlist  "--uninstall --multi-install --chrome --system-level --force-uninstall"
        #}

    #}

  #  $chromepath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Google Chrome"

    #if ($null -ne $chromepath) {

     #   $versions = (Get-ItemProperty -path 'HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\Google Chrome').version
      #  ForEach ($version in $versions) {
       #     write-output "Found Chrome version $version"
        #    $directory = ${env:ProgramFiles}
         #   write-output "Removing Chrome"
          #  Start-Process "$directory\Google\Chrome\Application\$version\Installer\setup.exe" -argumentlist  "--uninstall --multi-install --chrome --system-level --force-uninstall"
       # }


    #}
$xml = @"
<Configuration>
  <Display Level="None" AcceptEULA="True" />
  <Property Name="FORCEAPPSHUTDOWN" Value="True" />
  <Remove>
    <Product ID="O365HomePremRetail"/>
    <Product ID="OneNoteFreeRetail"/>
  </Remove>
</Configuration>
"@

##write XML to the debloat folder
$xml | Out-File -FilePath "C:\ProgramData\Debloat\o365.xml"

##Download the ODT
$odturl = "https://github.com/andrew-s-taylor/public/raw/main/De-Bloat/odt.exe"
$odtdestination = "C:\ProgramData\Debloat\odt.exe"
Invoke-WebRequest -Uri $odturl -OutFile $odtdestination -Method Get -UseBasicParsing

##Run it
Start-Process -FilePath "C:\ProgramData\Debloat\odt.exe" -ArgumentList "/configure C:\ProgramData\Debloat\o365.xml" -Wait

}
else {
    write-output "Intune detected, skipping removal of apps"
    write-output "$intunecomplete number of apps detected"

}

write-output "Completed"

#Stop-Transcript

}


############################################################################################################
#                                        MAIN PROGRAM LOOP                                                 #
#                                                                                                          #
############################################################################################################

do {
    Clear-Host  # Clear screen once at the start of each loop
    Show-MainMenu   # Then show the menu with the flashing warning
    $mainChoice = Read-Host "Enter your choice"
    
    switch ($mainChoice) {
        '1' {
            Show-SystemInfo
        }
        
        '2' {
            Write-Host "Starting full bloat removal..." -ForegroundColor Yellow
            Remove-Bloat
            Write-Host "Full bloat removal completed." -ForegroundColor Green
            Write-Host "Press any key to continue..."
            $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
        }

        '3' {
            Basic-Setup
        }
        '4' {
            Show-StarsStrikesMenu
        }
      
        '5' {
            Write-Host "Starting Citrix Installation..." -ForegroundColor Yellow
            try {
                Write-LogMessage "Starting Citrix Workspace management script..."
                # ... rest of your Citrix installation code ...
            }
            catch {
                Write-LogMessage "Critical error: $($_.Exception.Message)" -Level 'Error'
            }
            finally {
                Clean-TempFiles
            }
            Write-Host "Press any key to continue..."
            $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
        }
        
        '6' {
            Handle-ProfileManagement
        }
        
         'Q' {
            if (Get-Job -Id $flashingJob.Id -ErrorAction SilentlyContinue) {
                Stop-Job $flashingJob
                Remove-Job $flashingJob
            }
            Write-Host "Exiting program..." -ForegroundColor Yellow
            Cleanup-TempFiles "$env:TEMP\devicemanagement.zip" "$env:TEMP\DeviceManagement.ps1"
            exit
        }

        default {
            Write-Host "Invalid choice. Please select a valid option." -ForegroundColor Red
            Write-Host "Press any key to continue..."
            $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
        }
    }
} while ($true)
