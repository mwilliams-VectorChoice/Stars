   Write-Host ""
   Write-Host "@@@@@     @@@@@   @@@@@@@@@@@@@ @@@@          @@@ " -ForegroundColor DarkBlue   
   Write-Host "@@@@@     @@@@@   @@@@@@@@@@@@@ @@@@          @@@ " -ForegroundColor Blue
   Write-Host "@@@@@     @@@@@   @@@@          @@@@         @@@@ " -ForegroundColor Cyan
   Write-Host "@@@@@@   @@@@@@   @@@@          @@@@         @@@@ " -ForegroundColor Yellow
   Write-Host "@@@@@@   @@@@@@   @@@@           @@@         @@@@ " -ForegroundColor DarkCyan
   Write-Host "@@@@@@   @@@@@@   @@@@           @@@   @@@@  @@@@ " -ForegroundColor DarkGreen   
   Write-Host "@@@@@@@  @@@@@@   @@@@           @@@@ @@@@@  @@@  " -ForegroundColor Green  
   Write-Host "@@@@@@@ @@@@@@@   @@@@           @@@@ @@@@@  @@@  " -ForegroundColor DarkYellow 
   Write-Host "@@@ @@@ @@@ @@@   @@@@@@@@@@@@   @@@@ @@@@@  @@@  " -ForegroundColor Yellow
   Write-Host "@@@ @@@@@@@ @@@   @@@@@@@@@@@@   @@@@ @@@@@@@@@@  " -ForegroundColor DarkRed
   Write-Host "@@@ @@@@@@@ @@@   @@@@           @@@@ @@@@@@@@@@  " -ForegroundColor Red  
   Write-Host "@@@  @@@@@  @@@   @@@@            @@@ @@@@@@@@@@  " -ForegroundColor DarkMagenta
   Write-Host "@@@  @@@@@  @@@   @@@@            @@@@@@@@@@@@@@  " -ForegroundColor Magenta
   Write-Host "@@@  @@@@@  @@@   @@@@            @@@@@@ @@@@@@   " -ForegroundColor Gray
   Write-Host "@@@  @@@@   @@@   @@@@            @@@@@@ @@@@@@   " -ForegroundColor DarkGray
   Write-Host "@@@         @@@   @@@@            @@@@@@ @@@@@@   " -ForegroundColor White
   Write-Host "@@@         @@@   @@@@            @@@@@@ @@@@@@   " -ForegroundColor Blue
   Write-Host "@@@         @@@   @@@@            @@@@@@  @@@@@   " -ForegroundColor Gray
   Write-Host "@@@         @@@   @@@@@@@@@@@@@@  @@@@@   @@@@@   " -ForegroundColor Yellow
   Write-Host "@@@         @@@   @@@@@@@@@@@@@@   @@@@   @@@@@   " -ForegroundColor Magenta




   Write-Host ""
   Write-Host "====Marshall Ellis Williams JR====="  -ForegroundColor Blue
   Write-Host "=====The Marshall Kit=====" -ForegroundColor Blue

  Write-Host ""
  Write-Host ""
  Write-Host ""
  

# Function to check if the bloatware removal script exists
function Check-RemoveBloatScript {
    $bloatScriptPath = "C:\temp\removebloat.ps1"
    if (-Not (Test-Path $bloatScriptPath)) {
        Write-Host "Remove bloatware script not found at $bloatScriptPath." -ForegroundColor Red
        Write-Host "Press 'Q' to quit and find the script."
        $continueChoice = Read-Host "Enter your choice"
        if ($continueChoice -eq 'Q') {
            Write-Host "Exiting program..." -ForegroundColor Yellow
            exit
        }
    }
}

function Show-MainMenu {
    # Clear-Host
    Write-Host "================ Main Menu ================" -ForegroundColor Cyan
    Write-Host "1. User Profile Management" -ForegroundColor Cyan
    Write-Host "2. Remove Bloatware" -ForegroundColor Cyan
    Write-Host "3. Basic Setup" -ForegroundColor Cyan
    Write-Host "4. Stars and Stikes - YellowDog Setup" -ForegroundColor Cyan
    Write-Host "5. Install TeamViewer Host" -ForegroundColor Cyan
    Write-Host "Q. Quit Program" -ForegroundColor Red
    Write-Host "=========================================" -ForegroundColor Cyan
}

function Show-ProfileMenu {
    # Clear-Host
    Write-Host "`n------ Profile Management Options ------" -ForegroundColor Yellow
    Write-Host "1. Remove user completely (account and profile)" -ForegroundColor Cyan
    Write-Host "2. Remove only profile" -ForegroundColor Cyan
    Write-Host "3. Remove only the C:\Users\ folder" -ForegroundColor Cyan
    Write-Host "Q. Return to Main Menu" -ForegroundColor Red
    Write-Host "-------------------------------------" -ForegroundColor Yellow
}

function Get-UserList {
    try {
        $userFolders = Get-ChildItem "C:\Users" -Directory | Where-Object { $_.Name -notin @('Public', 'Default', 'Default User', 'All Users') }
        $userList = @()
        $index = 1
        
        foreach ($folder in $userFolders) {
            # Check if it's a local user
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
                ProfileExists = $true  # Since we're getting this from the actual directory
                IsLocalUser = $isLocalUser
                Enabled = $isEnabled
                Path = $folder.FullName
            }
            $userList += $userInfo
            
            # Display user information with additional details
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
        
        if ($profileChoice -match '^[123]$') {  # Changed from [APF] to [123]
            $userList = Get-UserList
            if ($null -eq $userList) { continue }
            
            Write-Host "`nEnter indices of users (comma-separated) or 'Q' to return" -ForegroundColor Cyan
            $userIndices = Read-Host "Selection"
            
            if ($userIndices -eq 'Q') { 
                # Clear-Host
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
            # Clear-Host
            return
        }
        else {
            Write-Host "Invalid choice. Please select a valid option." -ForegroundColor Red
        }
        
        Write-Host "`nPress any key to continue..."
        $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
        
    } while ($true)
}
function Download-TeamViewerHost {
    try {
        # Check if TeamViewer Host is already installed
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

        # If not installed, proceed with installation
        Write-Host "TeamViewer Host is not installed. Would you like to install it? (Y/N)" -ForegroundColor Yellow
        $installChoice = Read-Host
        
        if ($installChoice.ToUpper() -ne 'Y') {
            Write-Host "TeamViewer Host installation cancelled." -ForegroundColor Yellow
            return $false
        }

        # Create temp directory if it doesn't exist
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
        # Only attempt cleanup if installer was downloaded
        if (($installerPath) -and (Test-Path $installerPath)) {
            Remove-Item $installerPath -Force
        }
    }
}


function Download-Chrome {
    try {
        # Create temp directory if it doesn't exist
        $tempDir = "C:\temp"
        if (-not (Test-Path $tempDir)) {
            New-Item -ItemType Directory -Path $tempDir -Force | Out-Null
        }

        $installerPath = "C:\temp\ChromeSetup.exe"
        $url = "https://dl.google.com/chrome/install/latest/chrome_installer.exe"

        Write-Host "Downloading Chrome installer..." -ForegroundColor Yellow
        
        # Download Chrome
        $webClient = New-Object System.Net.WebClient
        $webClient.DownloadFile($url, $installerPath)
        
        Write-Host "Installing Chrome..." -ForegroundColor Yellow
        
        # Install Chrome silently
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
        # Cleanup
        if (Test-Path $installerPath) {
            Remove-Item $installerPath -Force
        }
    }
}

function Download-AdobeReader {
    try {
        # Create temp directory if it doesn't exist
        $tempDir = "C:\temp"
        if (-not (Test-Path $tempDir)) {
            New-Item -ItemType Directory -Path $tempDir -Force | Out-Null
        }

        $installerPath = "C:\temp\AdobeReaderDC.exe"
        # Using the enterprise version which doesn't include McAfee
        $url = "https://ardownload2.adobe.com/pub/adobe/reader/win/AcrobatDC/2300320269/AcroRdrDC2300320269_en_US.exe"

        Write-Host "Downloading Adobe Reader DC installer..." -ForegroundColor Yellow
        
        # Download Adobe Reader
        $webClient = New-Object System.Net.WebClient
        $webClient.DownloadFile($url, $installerPath)
        
        Write-Host "Installing Adobe Reader DC..." -ForegroundColor Yellow
        
        # Install Adobe Reader silently
        $process = Start-Process -FilePath $installerPath -ArgumentList "/sAll /rs /msi /norestart /quiet" -Wait -PassThru
        
        if ($process.ExitCode -eq 0) {
            Write-Host "Adobe Reader installed successfully!" -ForegroundColor Green
            
            # Wait a moment for installation to complete
            Start-Sleep -Seconds 5
            
            Write-Host "Setting Adobe Reader as default PDF viewer..." -ForegroundColor Yellow
            
            # Set Adobe Reader as default PDF viewer using Group Policy
            try {
                # Set default association for .pdf files
                $assocPath = "HKLM:\SOFTWARE\Classes\.pdf"
                if (!(Test-Path $assocPath)) {
                    New-Item -Path $assocPath -Force | Out-Null
                }
                Set-ItemProperty -Path $assocPath -Name "(Default)" -Value "AcroExch.Document.DC" -Force

                # Set default association for PDF content type
                $contentPath = "HKLM:\SOFTWARE\Classes\AcroExch.Document.DC"
                if (!(Test-Path $contentPath)) {
                    New-Item -Path $contentPath -Force | Out-Null
                }
                Set-ItemProperty -Path $contentPath -Name "(Default)" -Value "Adobe Acrobat Document DC" -Force

                # Use DISM to set the default app association
                $tempXml = "C:\temp\AdobeReaderDefault.xml"
                @"
<?xml version="1.0" encoding="UTF-8"?>
<DefaultAssociations>
  <Association Identifier=".pdf" ProgId="AcroExch.Document.DC" ApplicationName="Adobe Acrobat Reader DC" />
</DefaultAssociations>
"@ | Out-File -FilePath $tempXml -Encoding UTF8

                # Apply the default association
                $dismResult = Start-Process "DISM.exe" -ArgumentList "/Online /Import-DefaultAppAssociations:`"$tempXml`"" -Wait -PassThru -WindowStyle Hidden

                if ($dismResult.ExitCode -eq 0) {
                    Write-Host "Adobe Reader set as default PDF viewer successfully!" -ForegroundColor Green
                } else {
                    Write-Host "Warning: Could not set Adobe Reader as default PDF viewer" -ForegroundColor Yellow
                }

                # Clean up the temporary XML file
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
        # Cleanup
        if (Test-Path $installerPath) {
            Remove-Item $installerPath -Force
        }
    }
}

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
                # Install Teams using winget
                $process = Start-Process "winget" -ArgumentList "install", "Microsoft.Teams" -Wait -PassThru -NoNewWindow
                
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


function Deploy-StarsFiles {
    # Check if running with admin privileges
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

    # Check if YellowDog software is already installed
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
        # Create temp directory if it doesn't exist
        $tempDir = "C:\temp"
        if (-not (Test-Path $tempDir)) {
            New-Item -ItemType Directory -Path $tempDir -Force | Out-Null
        }

        # Download zip files from GitHub
        Write-Host "Downloading files from GitHub..." -ForegroundColor Yellow
        
        # Using raw GitHub URLs
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

        # Verify files were downloaded
        if (-not (Test-Path $arcadePath) -or -not (Test-Path $mainPath)) {
            Write-Host "Failed to download one or both zip files." -ForegroundColor Red
            return
        }

        # Extract zip files
        Write-Host "Extracting files..." -ForegroundColor Yellow
        try {
            # Extract to Stars and Strikes folder
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

        # Deploy Arcade files
        Write-Host "Deploying Arcade files..." -ForegroundColor Yellow
        try {
            $arcadeShortcutPath = "$extractPath\Arcade\Arcade\Public Desktop\YDInv_Arcade Inventory.lnk"
            $arcadeProgramPath = "$extractPath\Arcade\Arcade\YDI_Arcade"

            if (Test-Path $arcadeShortcutPath) {
                # Take ownership and grant full control to Administrators
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

        # Deploy Main files
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
        # Cleanup
        Write-Host "Cleaning up temporary files..." -ForegroundColor Yellow
        if (Test-Path $arcadePath) { Remove-Item $arcadePath -Force }
        if (Test-Path $mainPath) { Remove-Item $mainPath -Force }
        if (Test-Path $extractPath) { Remove-Item $extractPath -Recurse -Force }
    }

    Write-Host "Press any key to continue..."
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
}

function Basic-Setup {
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

    # Process each application
    foreach ($appName in @("Chrome", "Edge", "Word", "Excel", "PowerPoint", "Outlook", "Teams")) {
        Write-Host "Processing $appName..." -ForegroundColor Yellow
        $appInfo = $appConfig[$appName]
        
        # Find first existing path for the app
        $exePath = $appInfo.Paths | Where-Object { Test-Path $_ } | Select-Object -First 1
        
        if ($exePath) {
            Write-Host "Found $appName at: $exePath" -ForegroundColor Green
            $xmlTemplate += "        <taskbar:DesktopApp DesktopApplicationLinkPath=`"$exePath`" />`n"
        }
        else {
            # Special handling for apps with install functions
            if ($appInfo.InstallFunction) {
                Write-Host "$appName is not installed." -ForegroundColor Yellow
                Write-Host "Would you like to install $appName? (Y/N)" -ForegroundColor Cyan
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
                    Write-Host "Skipping $appName installation" -ForegroundColor Yellow
                }
            }
            else {
                Write-Host "Could not find $appName installation" -ForegroundColor Red
            }
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

    Write-Host "`nBasic setup completed. The taskbar has been configured with your applications." -ForegroundColor Green
    Write-Host "Note: You may need to log out and log back in to see all changes." -ForegroundColor Yellow
    
    Write-Host "Press any key to continue..."
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
}

# Main Program Loop
Check-RemoveBloatScript  # Check if the remove bloat script exists

do {
    Show-MainMenu
    $mainChoice = Read-Host "Enter your choice"
    
    switch ($mainChoice) {
        '1' {
            # Clear-Host
            Handle-ProfileManagement
        }
        '2' {
            Write-Host "Removing bloatware..." -ForegroundColor Yellow
            Write-Host "Press any key to continue..."
            $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
            Start-Process powershell.exe -ArgumentList "-File C:\temp\removebloat.ps1" -Verb RunAs
        }
        '3' {
            # Clear-Host
            Basic-Setup
        }
        '4' {
            # Clear-Host
            Deploy-StarsFiles
        }
        '5' {
            # Clear-Host
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
            Write-Host "Exiting program..." -ForegroundColor Yellow
            exit
        }
        default {
            Write-Host "Invalid choice. Please select a valid option." -ForegroundColor Red
            Write-Host "Press any key to continue..."
            $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
        }
    }
} while ($true)