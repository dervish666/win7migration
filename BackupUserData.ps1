###############################
#  User data backup script    #
#                             #
#  15/08/18 - First version   #
#  15/10/18 - Latest revision #
#                             #
#  Sam Castillo               #
###############################

Write-Host "====================="
Write-Host "User Backup script..."
Write-Host "====================="

# This is a simple function to create a log file 
function logit([string]$Entry) {
    $Entry | Out-File -Append $logfile
    Write-Host $Entry
}
#region setup
# Get username 
[string]$username = WMIC /node:$env:computername ComputerSystem Get username 
$username = $username  -replace ".*\\", ""
$username = $username.Trim()

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

# Find the main folder we are going to work with
if (!($UseBackupDefaults)) {
    $response = Read-Host("Do you want to backup to USB key? ")
} 
if ($response -eq 'y' -or $AlwaysBackupToUSB) {
    Get-WmiObject Win32_Volume -Filter ("DriveType={0}" -f [int][System.io.Drivetype]::removable) | Select-Object Name, capacity, filesystem
    $driveletter = Read-Host("Please enter the drive letter for the usb key")
    $userfolder = "$($driveletter):\$username"
    $usb = $true
} else {
    Write-Host -Nonewline "Default destination : "
    Write-Host $defdest
   if(($userfolder = Read-Host "Please enter destination or press enter for default") -eq ''){
       $userfolder = $defdest
    } 
   else {
       $userfolder
    }
}

# Set the logfile
$logfile = "$userfolder\$username.log"

# Check if the folder exists already, if it doesn't create a new one
If (Test-Path $userfolder) {
    Write-Host "Found the users folder in $userfolder !"
    Invoke-Item $userfolder
} else {
    # Create folder if it doesn't exist
    New-Item $userfolder -ItemType Directory
}

# We only really need one logfile so check that we are starting from fresh
if (Test-Path $logfile) {
    Remove-Item -Path $logfile -Force
}

logit("Log file generated at $(Get-Date)`n`n")
logit("`nUsername : $username")

#Find the user folder
logit("`nUser folder : $userfolder `n")
logit("Powershell version : $($PSVersionTable.PSVersion.Major)")

#Show the current network drives
$networkdrives = Get-PSDrive -PSProvider FileSystem | Select-Object Name, DisplayRoot, CurrentLocation, Description
logit($networkdrives)
#endregion

# Check if sourcefolder is accessible.
if ($($PSVersionTable.PSVersion.Major) -gt '4') {
    $UserProfile = [Environment]::GetFolderPath("UserProfile")
} else {
    $UserProfile = "C:\Users\$username"
}

#region copystuff
# Copy all files to destination folder
# Create array of folders to back up from
$sourcefolder = "$Userprofile\Favorites",
"$Userprofile\Links",
"$Userprofile\AppData\Roaming\Microsoft\Signatures",
"$Userprofile\AppData\Roaming\Microsoft\Sticky Notes",
"$Userprofile\AppData\Local\Microsoft\OneNote",
"$Userprofile\Videos\Lync Recordings",
"$Userprofile\appdata\Local\Google\Chrome\User Data\default\bookmarks"

# Copy each folder
ForEach ($folder in $sourcefolder){
    if (Test-Path $folder) {
        Copy-Item -Path $folder -Destination $userfolder -Recurse -Force
        logit("`nCopied $folder")
    }
}

# Check all files have been copied correctly
ForEach ($folder in $sourcefolder) {
    if (Test-Path $folder) {
        $Sourcecheck += Get-ChildItem -Recurse $folder
    }
}

Write-Host "Source Files"
$Sourcecheck

# Create destination array
$destfolder = "Favorites","Links","Signatures","Sticky Notes","OneNote","Lync Recordings"

# Work through the array, only copying if there is data
$i = 0
foreach ($folder in $destfolder) {
    Write-Host "Checking $userfolder\$folder"
    If (Test-Path $userfolder\$folder) {
        $Destcheck += Get-ChildItem -Recurse "$userfolder\$folder"
    }
    $i++
}
Write-Host "Destination Files"
$Destcheck
logit($(Compare-Object -ReferenceObject $Sourcecheck -DifferenceObject $Destcheck))
Write-Host "All the above files have been copied"

if (!($UseBackupDefaults)) {
    $response = Read-Host "Do you want to back up the docs/desktop/pics/videos folders? "
}

if ($response -eq 'y' -or $AlwaysBackupDocs){
    $documents = [Environment]::GetFolderPath("MyDocuments")
    $desktop = [Environment]::GetFolderPath("Desktop")
    $pictures = [Environment]::GetFolderPath("MyPictures")
    $videos = [Environment]::GetFolderPath("MyVideos")

    $sourcedocs = $documents,$desktop,$pictures,$videos
    $destdocs = "$userfolder\Documents","$userfolder\Desktop","$userfolder\Pictures","$userfolder\Videos"
    $i=0
    ForEach ($folder in $sourcedocs) {
        if (!($null -eq $folder)) {
            logit("Copying $folder please wait...")
            if (Test-Path $folder) {
                $cmdargs = @("$folder","$($destdocs[$i])","/xf","*.pst","/MIR","/NFL","/R:0","/W:2")
                Invoke-Expression "robocopy @cmdargs"
            } else {
                Write-Host "Arrrrgggghhh!!! I can't find the folder!!! "
                Invoke-Item $sauce
            }
        }
        $i++
    }
}
#endregion

#region psts
if (!($UseBackupDefaults)) {
    $response = Read-Host "Do you want to back up the psts? "
}

if ($response -eq 'y' -or $AlwaysBackupPST) {
    # Find out if outlook is running
    if (Get-Process -EA SilentlyContinue outlook | Where-Object {$_.ProcessName -eq "Outlook"}) {
        Write-Host "Outlook running..."
        $outlook = [Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application") 
    } else {
        Write-Host "Starting Outlook"
        $outlook = New-Object -ComObject Outlook.Application
        Write-Host "Waiting for outlook to start"
        Start-Sleep -Seconds 10
    }
    # Check all mapped psts from outlook
    $Namespace = $outlook.getNamespace("MAPI")
    $all_psts = $Namespace.Stores | Where-Object {($_.ExchangeStoreType -eq '3') -and ($_.FilePath -like '*.pst') -and ($_.IsDataFileStore -eq $true)}
    $pstfilepaths = @()
    ForEach ($pst in $all_psts) {
        $pstfilepaths += $pst.FilePath.ToString()
        $pstrootfolders += $pst.GetRootFolder.ToString()
    }
    Write-Host "Found the below psts, closing outlook"
    $pstfilepaths
    if ($pstfilepaths -ne $null) {
        Stop-Process -Name Outlook
        Write-Host "This can take some time...."
        $outlookrunning = Get-Process outlook -ErrorAction SilentlyContinue
        do {
            Start-Sleep 1
            Write-Host "Outlook still running, please wait...."
        } Until ($outlookrunning.HasExited)
    # Copy all files
        $pstfilepaths | ForEach-Object { 
            $size = Get-ChildItem $_ | Select-Object Length
            Write-Host "Copying $_ ... - Size : $($size.Length)"
            try {
                $name = Get-ChildItem $_
                $name.name
                Write-Host "$userfolder\$($name.name)"

                if (!(Test-Path $userfolder\$($name.name))) {
                    Copy-Item -Path $_ -Destination $userfolder -Force
                    Write-Host "Copied $_ to $userfolder"
                    logit("Copied $_ to $userfolder - $($size.Length)")
                } else {
                    Write-Host "Duplicate file found!"
                    Copy-Item -Path $_ -Destination "$userfolder\copy_$($name.name)"
                    Write-Host "Copied copy_$_ to $userfolder"
                    logit("Copied copy_$_ to $userfolder - $($size.Length)")
                }
            }
            catch {
                Write-Host "Unable to copy file."
            }
        }
        $response = Read-Host "Do you want to disconnect the PSTS from outlook?"
        if ($response -eq 'y'){
            $outlook = New-Object -ComObject Outlook.Application
            # Disconnect all mapped psts and close outlook
            Write-Host "Waiting for outlook to start"
            Start-Sleep -Seconds 10
            $Namespace = $outlook.getNamespace("MAPI")
            $all_psts = $Namespace.Stores | Where-Object {($_.ExchangeStoreType -eq '3') -and ($_.FilePath -like '*.pst') -and ($_.IsDataFileStore -eq $true)}
            ForEach ($pst in $all_psts){
                Write-Host "Going to try and disconnect $($pst.FilePath)"
                try {
                    $Outlook.Session.RemoveStore($pst.GetRootFolder())
                    logit("Disconnected $($pst.FilePath)")
                }
                catch [Exception] {
                    Write-Host "Something went wrong!! `n Details below `n`n $($_.Exception.ToString)"
                    logit("Unable to disconnect: Exception: $($_.Exception.ToString)")
                }   
            }
            Stop-Process -Name Outlook
        }
    } else {
        Write-Host "Did not find any active PSTs on the computer..."
    }
}
#endregion

Invoke-Item $userfolder
"$username copied at $(Get-Date) to $userfolder `n" | Out-File -FilePath "$def\MigrationLog.log" -Append
"$username,$userfolder,$(Get-Date)" | Out-File -FilePath "$def\Current.csv" -Append

if (!($UseBackupDefaults)) {
    $response = Read-Host "Do you want to copy the restored folder to the NAS? "
}

if ($response -eq 'y' -or $AlwaysBackupToNAS) {
    if ($usb) {
        New-PSDrive -Name X -PSProvider FileSystem -Root $nas -Credential $nascreds
        Copy-Item -Path $userfolder -Destination X:\$oldname -Recurse -Force
    }
}

if (!($UseBackupDefaults)) {
    $response = Read-Host "Do you want to check the hard drive for any more psts?"
}

if ($response -eq 'y' -or $AlwaysCheckForPSTs) {
    Write-Host "Looking on the C drive for all psts... "
    $drivepsts = Get-ChildItem -Path "C:\" -Recurse -Include '*.pst' -ErrorAction SilentlyContinue | Select-Object fullname,lastwritetime
    if ($drivepsts -ne $null) {
        Write-Host "I have found the following psts on the computer, some may have been imported already"
        logit($drivepsts)
    } else {
        Write-Host "No Pst's found."
    }
}

logit("All tasks finished at $(Get-Date) ")
Write-Host "Fin"
Write-Host "Press enter to exit"
Read-Host