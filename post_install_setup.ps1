# post_install_setup.ps1 - Install for actual user when running as admin
param(
    [string]$TargetUser = $null
)

Write-Host "Setting up Primus Implant Report Generator..." -ForegroundColor Green
Write-Host ""

$USER_FOUND = $false
$TARGET_USER = $null

# Function to test if a user directory exists and is valid
function Test-UserProfile {
    param([string]$Username)
    
    if ([string]::IsNullOrEmpty($Username)) { return $false }
    
    $userPath = "C:\Users\$Username"
    if (Test-Path $userPath) {
        $ntUserPath = Join-Path $userPath "NTUSER.DAT"
        return (Test-Path $ntUserPath)
    }
    return $false
}

# Function to get available user profiles
function Get-ValidUserProfiles {
    $users = @()
    $userDirs = Get-ChildItem "C:\Users" -Directory | Where-Object { 
        $_.Name -notin @("All Users", "Default", "Default User", "Public") 
    }
    
    foreach ($userDir in $userDirs) {
        if (Test-UserProfile $userDir.Name) {
            $users += $userDir.Name
        }
    }
    return $users
}

Write-Host "Detecting actual logged-in user..." -ForegroundColor Yellow

# Method 1: Check if user was passed as parameter or environment variable
if ($TargetUser) {
    $TARGET_USER = $TargetUser
    $USER_FOUND = $true
    Write-Host "Found user from parameter: $TARGET_USER" -ForegroundColor Cyan
}
elseif ($env:ORIGINAL_USER) {
    $TARGET_USER = $env:ORIGINAL_USER
    $USER_FOUND = $true
    Write-Host "Found original user from environment: $TARGET_USER" -ForegroundColor Cyan
}

# Method 2: Use WMI to get current interactive user (most reliable)
if (-not $USER_FOUND) {
    Write-Host "Trying WMI method..." -ForegroundColor Gray
    try {
        $computerSystem = Get-WmiObject -Class Win32_ComputerSystem -ErrorAction Stop
        if ($computerSystem.UserName) {
            $fullUserName = $computerSystem.UserName
            if ($fullUserName.Contains('\')) {
                $TARGET_USER = $fullUserName.Split('\')[1]
            } else {
                $TARGET_USER = $fullUserName
            }
            $USER_FOUND = $true
            Write-Host "Found user from WMI: $TARGET_USER" -ForegroundColor Cyan
        }
    }
    catch {
        Write-Host "WMI method failed: $($_.Exception.Message)" -ForegroundColor Yellow
    }
}

# Method 3: Check for users with explorer.exe running
if (-not $USER_FOUND) {
    Write-Host "Checking for active explorer processes..." -ForegroundColor Gray
    try {
        $explorerProcesses = Get-WmiObject -Class Win32_Process -Filter "Name='explorer.exe'" -ErrorAction Stop
        foreach ($process in $explorerProcesses) {
            $owner = $process.GetOwner()
            if ($owner.User -and $owner.User -ne "SYSTEM") {
                if (Test-UserProfile $owner.User) {
                    $TARGET_USER = $owner.User
                    $USER_FOUND = $true
                    Write-Host "Found user from explorer process: $TARGET_USER" -ForegroundColor Cyan
                    break
                }
            }
        }
    }
    catch {
        Write-Host "Explorer process check failed: $($_.Exception.Message)" -ForegroundColor Yellow
    }
}

# Method 4: Find the most recently modified user profile
if (-not $USER_FOUND) {
    Write-Host "Checking for most recently used profile..." -ForegroundColor Gray
    try {
        $validUsers = Get-ValidUserProfiles
        $newestUser = $null
        $newestTime = [DateTime]::MinValue
        
        foreach ($username in $validUsers) {
            $localAppData = "C:\Users\$username\AppData\Local"
            if (Test-Path $localAppData) {
                $lastWrite = (Get-Item $localAppData).LastWriteTime
                if ($lastWrite -gt $newestTime) {
                    $newestTime = $lastWrite
                    $newestUser = $username
                }
            }
        }
        
        if ($newestUser) {
            $TARGET_USER = $newestUser
            $USER_FOUND = $true
            Write-Host "Found most recently used profile: $TARGET_USER" -ForegroundColor Cyan
        }
    }
    catch {
        Write-Host "Profile check failed: $($_.Exception.Message)" -ForegroundColor Yellow
    }
}

# Method 5: Interactive selection if needed
if (-not $USER_FOUND) {
    Write-Host ""
    Write-Host "Could not automatically detect the target user." -ForegroundColor Yellow
    
    $validUsers = Get-ValidUserProfiles
    if ($validUsers.Count -eq 0) {
        Write-Host "No valid user profiles found!" -ForegroundColor Red
        exit 1
    }
    
    if ($validUsers.Count -eq 1) {
        $TARGET_USER = $validUsers[0]
        $USER_FOUND = $true
        Write-Host "Only one user found, selecting: $TARGET_USER" -ForegroundColor Cyan
    }
    else {
        Write-Host "Available users on this system:" -ForegroundColor White
        for ($i = 0; $i -lt $validUsers.Count; $i++) {
            Write-Host "  $($i + 1). $($validUsers[$i])" -ForegroundColor White
        }
        
        Write-Host ""
        do {
            $userChoice = Read-Host "Please select the user (1-$($validUsers.Count)) or type the username directly"
            
            # Check if it's a number
            if ($userChoice -match '^\d+$') {
                $index = [int]$userChoice - 1
                if ($index -ge 0 -and $index -lt $validUsers.Count) {
                    $TARGET_USER = $validUsers[$index]
                    $USER_FOUND = $true
                    Write-Host "Selected user by number: $TARGET_USER" -ForegroundColor Cyan
                }
                else {
                    Write-Host "Invalid selection. Please try again." -ForegroundColor Red
                }
            }
            else {
                # Treat as direct username input
                if (Test-UserProfile $userChoice) {
                    $TARGET_USER = $userChoice
                    $USER_FOUND = $true
                    Write-Host "Entered username: $TARGET_USER" -ForegroundColor Cyan
                }
                else {
                    Write-Host "User '$userChoice' not found or invalid. Please try again." -ForegroundColor Red
                }
            }
        } while (-not $USER_FOUND)
    }
}

if (-not $USER_FOUND) {
    Write-Host "Error: Could not determine target user for installation." -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "Target user: $TARGET_USER" -ForegroundColor Green

# Validate user directory
$USER_HOME_DIR = "C:\Users\$TARGET_USER"
if (-not (Test-Path $USER_HOME_DIR)) {
    Write-Host "Error: User directory $USER_HOME_DIR does not exist." -ForegroundColor Red
    exit 1
}

Write-Host "Validated user directory: $USER_HOME_DIR" -ForegroundColor Green

# Confirm with user before proceeding
Write-Host ""
Write-Host "About to install Primus Implant Report Generator for user: $TARGET_USER" -ForegroundColor White
Write-Host "Installation directory: $USER_HOME_DIR\AppData\Local\CreoDent\PrimusImplant" -ForegroundColor White
Write-Host ""

# $confirm = Read-Host "Is this correct? (Y/N)"
# if ($confirm -notmatch '^[Yy]') {
#     Write-Host "Installation cancelled by user." -ForegroundColor Yellow
#     exit 0
# }

# Set up paths for the target user
$USER_DATA_DIR = "$USER_HOME_DIR\AppData\Local\CreoDent\PrimusImplant"
$USER_DESKTOP = "$USER_HOME_DIR\Desktop"
$USER_START_MENU = "$USER_HOME_DIR\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\CreoDent"

Write-Host ""
Write-Host "Target directory: $USER_DATA_DIR" -ForegroundColor White
Write-Host "Desktop path: $USER_DESKTOP" -ForegroundColor White
Write-Host "Start menu path: $USER_START_MENU" -ForegroundColor White

# Create directory structure
Write-Host ""
Write-Host "Creating directory structure..." -ForegroundColor Yellow

try {
    New-Item -Path $USER_DATA_DIR -ItemType Directory -Force | Out-Null
    New-Item -Path "$USER_DATA_DIR\updates" -ItemType Directory -Force | Out-Null
    New-Item -Path "$USER_DATA_DIR\logs" -ItemType Directory -Force | Out-Null
    Write-Host "Successfully created directories" -ForegroundColor Green
}
catch {
    Write-Host "Error: Failed to create user data directory: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

# Copy application files
Write-Host ""
Write-Host "Copying application files..." -ForegroundColor Yellow
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$filescopied = 0

# Copy main executable
$mainExe = Join-Path $scriptDir "Primus Implant Report Generator 1.0.5.exe"
if (Test-Path $mainExe) {
    try {
        Copy-Item $mainExe $USER_DATA_DIR -Force
        Write-Host "Copied main executable" -ForegroundColor Green
        $filescopied++
    }
    catch {
        Write-Host "Warning: Failed to copy main executable: $($_.Exception.Message)" -ForegroundColor Yellow
    }
}

# Copy data files
@("*.csv", "*.png", "*.ico") | ForEach-Object {
    $files = Get-ChildItem (Join-Path $scriptDir $_) -ErrorAction SilentlyContinue
    if ($files) {
        try {
            Copy-Item $files.FullName $USER_DATA_DIR -Force
            Write-Host "Copied $($_.TrimStart('*')) files" -ForegroundColor Green
            $filescopied++
        }
        catch {
            Write-Host "Warning: Failed to copy $($_.TrimStart('*')) files: $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }
}

if ($filescopied -eq 0) {
    Write-Host "No application files found to copy" -ForegroundColor Yellow
}

# Set directory ownership
Write-Host ""
Write-Host "Setting directory ownership..." -ForegroundColor Yellow
try {
    $acl = Get-Acl $USER_DATA_DIR
    $accessRule = New-Object System.Security.AccessControl.FileSystemAccessRule($TARGET_USER, "FullControl", "ContainerInherit,ObjectInherit", "None", "Allow")
    $acl.SetAccessRule($accessRule)
    Set-Acl $USER_DATA_DIR $acl
    Write-Host "Directory ownership set successfully" -ForegroundColor Green
}
catch {
    Write-Host "Warning: Failed to set directory ownership: $($_.Exception.Message)" -ForegroundColor Yellow
}

# Create registry entries
Write-Host ""
Write-Host "Setting registry entries..." -ForegroundColor Yellow
try {
    $regPath = "HKLM:\SOFTWARE\CreoDent\PrimusImplant"
    if (-not (Test-Path $regPath)) {
        New-Item -Path $regPath -Force | Out-Null
    }
    Set-ItemProperty -Path $regPath -Name "UserDataPath" -Value "%LOCALAPPDATA%\CreoDent\PrimusImplant" -Force
    Set-ItemProperty -Path $regPath -Name "InstalledForUser" -Value $TARGET_USER -Force
    Write-Host "Registry entries created successfully" -ForegroundColor Green
}
catch {
    Write-Host "Warning: Failed to create registry entries: $($_.Exception.Message)" -ForegroundColor Yellow
}

# Create shortcuts
function New-Shortcut {
    param(
        [string]$ShortcutPath,
        [string]$TargetPath,
        [string]$WorkingDirectory,
        [string]$Description
    )
    
    try {
        $WshShell = New-Object -ComObject WScript.Shell
        $Shortcut = $WshShell.CreateShortcut($ShortcutPath)
        $Shortcut.TargetPath = $TargetPath
        $Shortcut.WorkingDirectory = $WorkingDirectory
        $Shortcut.IconLocation = $TargetPath
        $Shortcut.Description = $Description
        $Shortcut.Save()
        
        # Set ownership
        $acl = Get-Acl $ShortcutPath
        $owner = New-Object System.Security.Principal.NTAccount($TARGET_USER)
        $acl.SetOwner($owner)
        Set-Acl $ShortcutPath $acl
        
        return $true
    }
    catch {
        Write-Host "Shortcut creation failed: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

# Create desktop shortcut
Write-Host ""
Write-Host "Creating desktop shortcut..." -ForegroundColor Yellow
$userExePath = Join-Path $USER_DATA_DIR "Primus Implant Report Generator 1.0.5.exe"
$desktopShortcutPath = Join-Path $USER_DESKTOP "Primus Implant Report Generator.lnk"

if (-not (Test-Path $userExePath)) {
    Write-Host "Warning: Executable not found in user directory, using original location" -ForegroundColor Yellow
    $userExePath = $mainExe
}

Write-Host "Creating desktop shortcut pointing to: $userExePath" -ForegroundColor Gray

if (-not (Test-Path $USER_DESKTOP)) {
    New-Item -Path $USER_DESKTOP -ItemType Directory -Force | Out-Null
}

if (New-Shortcut -ShortcutPath $desktopShortcutPath -TargetPath $userExePath -WorkingDirectory $USER_DATA_DIR -Description "Primus Implant Report Generator") {
    Write-Host "Desktop shortcut created successfully" -ForegroundColor Green
} else {
    Write-Host "Warning: Desktop shortcut creation failed" -ForegroundColor Yellow
}

# Create Start Menu shortcut
Write-Host ""
Write-Host "Creating Start Menu shortcut..." -ForegroundColor Yellow
$startMenuShortcutPath = Join-Path $USER_START_MENU "Primus Implant Report Generator.lnk"

if (-not (Test-Path $USER_START_MENU)) {
    New-Item -Path $USER_START_MENU -ItemType Directory -Force | Out-Null
}

Write-Host "Creating Start Menu shortcut pointing to: $userExePath" -ForegroundColor Gray

if (New-Shortcut -ShortcutPath $startMenuShortcutPath -TargetPath $userExePath -WorkingDirectory $USER_DATA_DIR -Description "Primus Implant Report Generator") {
    Write-Host "Start Menu shortcut created successfully" -ForegroundColor Green
} else {
    Write-Host "Warning: Start Menu shortcut creation failed" -ForegroundColor Yellow
}

# Create uninstaller
Write-Host ""
Write-Host "Creating uninstaller..." -ForegroundColor Yellow
$uninstallerPath = Join-Path $USER_DATA_DIR "uninstall.ps1"

$uninstallerContent = @"
# Uninstaller for Primus Implant Report Generator
Write-Host "Uninstalling Primus Implant Report Generator..." -ForegroundColor Yellow
Write-Host "Installation was for user: $TARGET_USER" -ForegroundColor White
Write-Host ""

# Stop any running instances
try {
    Get-Process "Primus Implant Report Generator 1.0.5" -ErrorAction Stop | Stop-Process -Force
    Start-Sleep -Seconds 2
    Write-Host "Stopped running application instances" -ForegroundColor Green
}
catch {
    Write-Host "No running instances found" -ForegroundColor Gray
}

# Remove desktop shortcut
`$desktopShortcut = "$USER_DESKTOP\Primus Implant Report Generator.lnk"
if (Test-Path `$desktopShortcut) {
    Remove-Item `$desktopShortcut -Force
    Write-Host "Removed desktop shortcut" -ForegroundColor Green
}

# Remove Start Menu shortcuts
`$startMenuDir = "$USER_START_MENU"
if (Test-Path `$startMenuDir) {
    Remove-Item `$startMenuDir -Recurse -Force
    Write-Host "Removed Start Menu shortcuts" -ForegroundColor Green
}

# Optional: Remove user data
Write-Host ""
`$removeData = Read-Host "Do you want to remove user data and settings? (Y/N)"
if (`$removeData -match '^[Yy]') {
    Write-Host "Removing user data..." -ForegroundColor Yellow
    Remove-Item "$USER_DATA_DIR" -Recurse -Force
    Write-Host "User data removed" -ForegroundColor Green
} else {
    Write-Host "User data preserved at: $USER_DATA_DIR" -ForegroundColor White
}

# Remove registry entries
try {
    Remove-Item "HKLM:\SOFTWARE\CreoDent\PrimusImplant" -Recurse -Force -ErrorAction Stop
    Write-Host "Registry entries removed" -ForegroundColor Green
}
catch {
    Write-Host "Warning: Failed to remove registry entries: `$(`$_.Exception.Message)" -ForegroundColor Yellow
}

Write-Host ""
Write-Host "Uninstallation complete." -ForegroundColor Green
Write-Host ""
Read-Host "Press Enter to exit"
"@

try {
    Set-Content -Path $uninstallerPath -Value $uninstallerContent -Force
    Write-Host "Uninstaller created successfully" -ForegroundColor Green
}
catch {
    Write-Host "Warning: Uninstaller creation failed: $($_.Exception.Message)" -ForegroundColor Yellow
}

# Final summary
Write-Host ""
Write-Host "================================================" -ForegroundColor Green
Write-Host "Setup completed successfully!" -ForegroundColor Green
Write-Host "================================================" -ForegroundColor Green
Write-Host "Application installed for user: $TARGET_USER" -ForegroundColor White
Write-Host "Application data directory: $USER_DATA_DIR" -ForegroundColor White
Write-Host "Desktop shortcut: $USER_DESKTOP\Primus Implant Report Generator.lnk" -ForegroundColor White
Write-Host "Start menu shortcut: $USER_START_MENU\Primus Implant Report Generator.lnk" -ForegroundColor White
Write-Host "================================================" -ForegroundColor Green
Write-Host ""
Write-Host "Installation script completed." -ForegroundColor Green