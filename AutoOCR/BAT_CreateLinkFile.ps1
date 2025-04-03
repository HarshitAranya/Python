Set-ExecutionPolicy RemoteSigned -Scope CurrentUser

$a = Get-Location
$b = Split-Path $a

$shortcutPath = Join-Path $b "AutoOCR.lnk"
$targetPath = Join-Path $a "start.bat"
$workingDirectory = "$a"
$iconLocation = Join-Path $a "HARSHIT.ico"  # Optional, if you want to set a custom icon

# Define window style (1=Normal, 3=Maximized, 7=Minimized)
#$windowStyle = 3 

# Create the Shell COM Object to create the shortcut
$WshShell = New-Object -ComObject WScript.Shell
$shortcut = $WshShell.CreateShortcut($shortcutPath)

# Set the properties of the shortcut
$shortcut.TargetPath = $targetPath
$shortcut.WorkingDirectory = $workingDirectory
$shortcut.IconLocation = $iconLocation  # Optional
#$shortcut.WindowStyle = $windowStyle  # Set window style
$shortcut.Save()

Write-Host "Shortcut created successfully at $shortcutPath"
timeout 3
