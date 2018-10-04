@Echo OFF
pushd "%~dp0"
cd
C:\Windows\system32\WindowsPowerShell\v1.0\powershell.exe -ExecutionPolicy Bypass %~dp0CheckInstalled.ps1
popd
