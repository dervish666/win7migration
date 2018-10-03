###############################
#  Application checker        #
#                             #
#  29/09/18 - First version   #
#  02/10/18 - Latest revision #
#                             #
#  Sam Castillo               #
###############################


$outputfile = "$($env:COMPUTERNAME)_Applications.txt"

Write-Host "This script will check if the applications have installed correctly. "
Write-Host "outputting to $outputfile"
Write-Host "Checking $($env:COMPUTERNAME)"
#Find all the install applications
function Is-Installed( $program ) {    
    $x86 = ((Get-ChildItem "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall") |
        Where-Object { $_.GetValue( "DisplayName" ) -like "*$program*" } ).Length -gt 0;
    $x64 = ((Get-ChildItem "HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall") |
        Where-Object { $_.GetValue( "DisplayName" ) -like "*$program*" } ).Length -gt 0;
    return $x86 -or $x64;
}

# The list of applications to be searched for. The names don't have to be exact, but close enough
# to not be mistaken for a similar app. 
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
    $installed = Is-Installed($app)
    Write-Host -NoNewLine "Checking $app : " 
    if (!($installed)) {
        Write-Host -ForegroundColor Red $installed
        "$app is not installed" | Out-File -Append $outputfile
    } else {
        Write-Host -ForegroundColor Green $installed
    }
}

Write-Host "`n`nNow going to check all the essential processes are running correctly"
# Now to check the running processes. 
# List of processes to check, these must be exactly the same as the process name. 
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
        Write-Host -ForegroundColor Red -NoNewLine "$proc is not running! - "
        "$proc is not running" | Out-File -Append $outputfile
        Write-Host "Please check if this is something that can be installed manually before continuing."
    } else {
        Write-Host -ForegroundColor Green "$proc is running..."
    }
}

