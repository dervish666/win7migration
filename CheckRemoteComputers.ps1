###############################
#  Remote computer checker    #
#                             #
#  29/09/18 - First version   #
#  02/10/18 - Latest revision #
#                             #
#  Sam Castillo               #
###############################

$pclist = Get-Content "$psscriptroot\adpclist.txt"
$creds = Get-Credential
Enable-PSRemoting -Force

ForEach ($pc in $pclist) {
    if (Test-Connection -ComputerName $pc -Count 1 -Quiet) {
        try {
            "$pc,Found" | Out-File -Append ".\Machines.csv"
            Invoke-Command -ComputerName $pc -FilePath .\CheckInstalled.ps1 -Credential $creds
        }
        catch {
            Write-Host "Unable to find $pc"
            "$pc,Lost" | Out-File -Append ".\Machines.csv"
        }
    } else {
        Write-Host "Can't find $pc"
        "$pc,Lost" | Out-File -Append ".\Machines.csv"
    }
}
Read-Host