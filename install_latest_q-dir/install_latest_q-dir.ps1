Write-Host "Downloading latest Q-Dir ..."
$url = "https://www.softwareok.com/Download/Q-Dir_Installer_x64.zip"
$output = "$PSScriptRoot\Q-Dir_Installer_x64.zip"
$start_time = Get-Date
(New-Object System.Net.WebClient).DownloadFile($url, $output)
Write-Host "Time taken: $((Get-Date).Subtract($start_time).Seconds) second(s)"


Write-Host "Extracting $output ..."
Expand-Archive "$output" -DestinationPath "$PSScriptRoot" -Force


Write-Host "Updating version number ..."
$version = (Get-Command "$PSScriptRoot\Q-Dir_Installer_x64.exe").FileVersionInfo.FileVersion
$version = $version.replace(",",".")
$newName = "Q-Dir_Installer_x64_v$version.exe"
Move-Item -Path "$PSScriptRoot\Q-Dir_Installer_x64.exe" -Destination "$PSScriptRoot\$newName" -Force


Write-Host "Cleanning downloaded files ..."
Remove-Item $output


Write-Host "Stopping running Q-Dir process if exists ..."
Stop-Process -Name "Q-Dir" -Force -ErrorAction SilentlyContinue


Write-Host "Installing Q-Dir ..."
& "$PSScriptRoot\$newName" -install /silent noautostart nodesktop admin noquicklaunch
