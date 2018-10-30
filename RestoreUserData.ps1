###############################
#  User data restore script   #
#                             #
#  15/08/18 - First version   #
#  27/09/18 - Latest revision #
#                             #
#  Sam Castillo               #
###############################
#region setup
$winver = [Environment]::OSVersion
[string]$username = WMIC /node:$env:computername ComputerSystem Get username 
$username = $username  -replace ".*\\", ""
$username = $username.Trim()

Write-Host "====================="
Write-Host "User Restore script.."
Write-Host "====================="

#Checking defaults.
$contents = Get-Content ".\defaults.txt"
$i=0
Write-Host "`nDefault values from defaults.txt`n"
ForEach ($line in $contents) {
    New-Variable -Name $contents[$i].split(":")[0] -Value $contents[$i].split(":")[1]
    Write-Host -NoNewline "$($contents[$i].split(":")[0]) = "
    if ($($contents[$i].split(":")[1]) -eq 'true') {
        Write-Host -ForegroundColor Green "$($contents[$i].split(":")[1])"
    } elseif ($($contents[$i].split(":")[1]) -eq 'false') {
        Write-Host -ForegroundColor Red "$($contents[$i].split(":")[1])"
    } else {
        Write-Host "$($contents[$i].split(":")[1])"
    }
    $i++
}

Write-Host "`n===`n"

$oldname = Read-Host("What was the old username for this user? ")
$UserProfile = [Environment]::GetFolderPath("UserProfile")
$onedrive = "$userprofile\OneDrive - IMPERIAL TOBACCO LTD"
if (!(Test-Path $onedrive)) {
    Write-Host "Can't find the users Onedrive folder! - Make sure you do the gpupdate in a minute and reboot after."
    Read-Host
    $documents = [Environment]::GetFolderPath("MyDocuments")
    $logfile = "$documents\$username.restore.log"
} else {
    $logfile = "$onedrive\$username.restore.log"
}
function logit([string]$Entry) {
    $Entry | Out-File -Append $logfile
    Write-Host $entry
}

logit("Log file generated at $(Get-Date)`n")
logit("`nUsername : $username")
logit("`nMachine  : $env:computername")
logit("Restoring to $($winver.VersionString)")
#endregion

#region applications
#Check all applications are installed
function Installed( $program ) {    
    $x86 = ((Get-ChildItem "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall") |
        Where-Object { $_.GetValue( "DisplayName" ) -like "*$program*" } ).Length -gt 0;
    $x64 = ((Get-ChildItem "HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall") |
        Where-Object { $_.GetValue( "DisplayName" ) -like "*$program*" } ).Length -gt 0;
    return $x86 -or $x64;
}

$response = Read-Host "Do you want to do a GPUdate?"
if ($response -eq 'y') {
    Write-Host "Gonna do a GPUpdate to make sure the onedrive folder is present"
    Write-Host "You will now probably be logged off...."
    gpupdate.exe /force /Logoff
}

$applications =
"Trend Micro Full Disk Encryption",
"Forcepoint Endpoint",
"Office 16",
"Pulse Secure",
"Snow Inventory Agent",
"Microsoft Office 365 ProPlus - en-us",
"Trend Micro OfficeScan Agent",
"EDD Client",
"Google Chrome"

ForEach ($app in $applications) {
    $installed = Installed($app)
    Write-Host -NoNewLine "Checking $app : " 
    if (!($installed)) {
        Write-Host -ForegroundColor Red $installed
        logit("$app is not installed!!!!")
        Read-Host
    } else {
        Write-Host -ForegroundColor Green $installed
    }
}

#Check if essentials are running. 

Write-Host "Continue if all the applications are installed correctly...`n"
Write-Host "Now going to check if the essential processes are running...`n`n"

$process = 
"OneDrive",
"PulseSecureService",
"snowagent",
"TMCCSF",
"PccNTMon",
"TmListen",
"NTRtScan",
"TMBMSRV",
"wepsvc"

ForEach ($proc in $process) {
    if ((Get-Process $proc -ErrorAction SilentlyContinue) -eq $Null) {
        Write-Host -ForegroundColor Red "$proc is not running!"
        Write-Host "Please check if this is something that can be installed manually before continuing."
        Read-Host
    } else {
        Write-Host -ForegroundColor Green "$proc is working..."
    }
}

Write-Host "If any of the above processes are not running, please fix before continuing...`n`n"
#endregion

#region startrestore
if ($UseRestoreDefaults -eq 'false') {
    Write-Host "Are you going to be restoring from the USB Key?"
    $response = Read-Host "Restoring from USB? "
}

if ($response -eq 'y' -or $AlwaysRestoreFromUSB -eq 'true') {
    $driveletter = Read-Host("Please enter drive letter for the usb key")
    $userfolder = "$($driveletter):\$oldname"
    $usb = $true
} else {
    Write-Host "Default source is : $nas"
    if (($userfolder = Read-Host "Please enter the source or enter for default:") -eq '') {
        $userfolder = New-PSDrive -Name X -PSProvider FileSystem -Root $nas -Credential $nascreds
    } else {
        $userfolder
    }   
    $userfolder = "$($userfolder):\$oldname"
}

if (Test-Path $logfile) {
    Remove-Item -Path $logfile -Force
}
Write-Host "Looking for $userfolder"
If (Test-Path $userfolder) {
    logit("Found the users folder in $userfolder")
} else {
    Write-Host "Unable to find the users folder! Will restore only from Users$"
}

if ($UseRestoreDefaults -eq 'false'){
    $response = Read-Host("Do you want to restore the users folders? ")
}

if ($response -eq 'y' -or $AlwaysRestoreUsersFolders -eq 'true') {
    #Rename Sticky notes ready for import
    if ($winver.Version.Major -eq '10') {
        if (Test-Path "$userfolder\Sticky Notes\StickyNotes.snt") {
            Rename-Item -Path "$userfolder\Sticky Notes\StickyNotes.snt" -NewName "ThresholdNotes.snt"
            logit("Converted sticky notes")
        }
        # Set destination array
        $destfolder = "$userprofile\",
        "$userprofile\",
        "$userprofile\AppData\Roaming\Microsoft\",
        "$userprofile\AppData\Local\Packages\Microsoft.MicrosoftStickyNotes_8wekyb3d8bbwe\LocalState\Legacy",
        "$userprofile\AppData\Local\Microsoft\",
        "$userprofile\Videos\"
    } else {
        $destfolder = "$userprofile\",
        "$userprofile\",
        "$userprofile\AppData\Roaming\Microsoft\",
        "$userprofile\AppData\Roaming\Microsoft\",
        "$userprofile\AppData\Local\Microsoft\",
        "$userprofile\Videos\"
    }
    $sourcefolder = "Favorites","Links","Signatures","Sticky Notes","OneNote","Lync Recordings"

    # Copy the files back
    $i = 0
    foreach ($folder in $sourcefolder) {
        If (Test-Path $userfolder\$folder) {
            Copy-Item -Path $userfolder\$folder -Destination $($destfolder[$i]) -Recurse -Force
            logit("Copied $userfolder\$folder to $($destfolder[$i])")
        }
        $i++
    }
    if (Test-Path "$userprofile\AppData\Local\Packages\Microsoft.MicrosoftStickyNotes_8wekyb3d8bbwe\LocalState\Sticky Notes") {
        Rename-Item -Path "$userprofile\AppData\Local\Packages\Microsoft.MicrosoftStickyNotes_8wekyb3d8bbwe\LocalState\Sticky Notes" -NewName "Legacy"
    }
}
# Restore Chrome
if (Test-Path $userfolder\bookmarks) {
    logit("Found Chrome bookmarks, restoring...")
    Write-Host "Starting google chrome.."
    Start-Process -FilePath "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
    Start-Sleep -Seconds 10
    $chrome = Get-Process chrome -ErrorAction SilentlyContinue
    if ($chrome) {
        $chrome.CloseMainWindow()
        Start-Sleep -Seconds 5
        if (!$chrome.HasExited){
            $chrome | Stop-Process -Force
        }
    }
    Copy-Item -Path "$userfolder\bookmarks" -Destination "$userprofile\appdata\Local\Google\Chrome\User Data\default\" -Force
}

if ($UseRestoreDefaults -eq 'false') {
    $response = Read-Host("Do you want to restore the Docs/desktop/pic/vids folders to Onedrive? ") 
}

if ($response -eq 'y' -or $AlwaysRestoreToOneDrive -eq 'true') {
    if ($usb) {
        $sauce = $userfolder

    } else {
        $sauce = Read-Host "Please enter full path to restore from "
    }
    # Now it's time to copy everything to onedrive
    # Find the files in the users user folder
 
    $sourcedocs = "$sauce\Documents",
    "$sauce\Desktop",
    "$sauce\Favorites",
    "$sauce\Pictures",
    "$sauce\Videos"

    $destdocs = "$onedrive\Documents",
    "$onedrive\Desktop",
    "$onedrive\Favorites",
    "$onedrive\Pictures",
    "$onedrive\Videos"
    # Restore My docs and desktop from 014 for onedrive, Don't copy any psts

    $i=0
    ForEach ($folder in $sourcedocs) {
        logit("Copying $folder please wait...")
        if (Test-Path $folder) {
            $cmdargs = @("$folder","$($destdocs[$i])","/xf","*.pst","/MIR","/NFL","/R:0","/W:2")
            Invoke-Expression "robocopy @cmdargs"
            $i++
        } else {
            Write-Host "Arrrrgggghhh!!! I can't find the folder!!! "
            Invoke-Item $sauce
        }
    }
}
#endregion

#region finishup

if ($UseRestoreDefaults -eq 'false') {
    $response = Read-Host ("Do you want to copy the psts to downloads? ")
}

if ($response -eq 'y' -or $AlwaysCopyPSTSToDLs -eq 'true') {
    # Copy pst files to Downloads folder. 
    logit("Copying the psts to the Downloads folder")
    Copy-Item -Path $userfolder\*.pst -Include "*.pst" -Destination "$userprofile\Downloads\" -Recurse
    Invoke-Item $userprofile\Downloads\
}

logit("All tasks finished at $(Get-Date) ")
Write-Host "Fin"
Read-Host
#endregion