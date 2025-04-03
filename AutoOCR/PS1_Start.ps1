Set-ExecutionPolicy RemoteSigned -Scope CurrentUser
$ScriptPath = $MyInvocation.MyCommand.Path
$ScriptDir = Split-Path -Parent $ScriptPath
Write-Host "$ScriptPath"
Write-Host "$ScriptDir"
CD "$ScriptDir"
.\Scripts\activate
python .\readDocFiles.py