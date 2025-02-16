[System.Net.WebRequest]::DefaultWebProxy.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials

$currentDir = Get-Location
$patFilePath = Join-Path $currentDir "PAT.txt"
$PAT = Get-Content -Path $patFilePath
# Write-Host "This is pat - $PAT"
# CdD7vTeUJtaurHKbTiYYEG6Z0dIVSHh22zC1XDnV1IXQK2UUT8nCJQQJ99BAACAAAAANjyhyAAASAZDOOLeO
# if(!$PAT){
#     Write-Host "Please update PAT in PAT.txt file"
#     exit
# }

# Determine if the script is running as a standalone executable or as a script from a file
if ($PSScriptRoot -eq "") {
    # If running as an executable, $PSScriptRoot will be empty, so we get the directory of the running executable
    $currentExeDir = [System.IO.Path]::GetDirectoryName($MyInvocation.MyCommand.Definition)
} else {
    # If running as a script, $PSScriptRoot will be the folder where the script is located
    $currentExeDir = $PSScriptRoot
}

$henvFilePath = Join-Path $currentExeDir "henv.json"

if (Test-Path $henvFilePath) {
    $henvContent = Get-Content -Path $henvFilePath | ConvertFrom-Json
} else {
    Write-Host "The file 'henv.json' does not exist in the current folder."
}

# $PAT = $henvContent.PAT
$Query_existingItems = $henvContent.Query_existingItems

function boardsdata(){
    param(
        [string]$PAT,
        [string]$Query
    )
    #$user = ""
    #$base64AuthInfo = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(("{0}:{1}" -f $user,$PAT)))
    $Base64AuthInfo = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(":$PAT"))
    $orgUrl = "https://dev.azure.com/civica-cp/CE"
    $targetQueryFolder = "Shared%20Queries/CJS_ChangeManagement/$Query"
    $queryUrl = "$orgUrl/_apis/wit/queries/${targetQueryFolder}?api-version=5.0"
    $queryId = (Invoke-RestMethod -Uri $queryUrl -Headers @{Authorization=("Basic {0}" -f $base64AuthInfo)} -Method Get).id
    $executeQueryUrl = "$orgUrl/_apis/wit/wiql/${queryId}?api-version=5.0"
    $workItems = (Invoke-RestMethod -Uri $executeQueryUrl -Headers @{Authorization=("Basic {0}" -f $base64AuthInfo)} -Method Get).workItems.id
    $finalOutput = @()
    foreach($workItem in $workItems){
    
        $ApiUrl = "$orgUrl/_apis/wit/workitems/${workItem}?api-version=7.1-preview.3"
    
        try {
            $response = Invoke-RestMethod -Uri $ApiUrl -Method GET -Headers @{
            Authorization = "Basic $Base64AuthInfo" }
        } catch {
            Write-Host "Error:" -ForegroundColor Red
            Write-Host $_.Exception.Message
        }
    
        $finalOutput += $response.fields
    }
    
    return $finalOutput #| Select-Object @{Name='Deployment'; Expression={$_.('WEF_C73ED0D35FEC4159BB0D807A6B360EE6_Kanban.Column')}}, System.Tags, @{Name='Description'; Expression={$_.('System.Description') -replace '<[^>]*>', ''}} | Format-Table
    
}

# $existingItems = boardsdata -PAT $PAT -Query $Query_existingItems
# $existingItems
$existingItemsList = @()
$existingItems = boardsdata -PAT $PAT -Query $Query_existingItems
# $existingItemsList | Get-Member 

# $temp1 = $existingItems | Where-Object { $_.'System.Tags' -eq "Awaiting Hypercare completion 26/02; OCR25046; RPA" }
# $temp1 | Format-List *

$existingItemsList = $existingItems | Select-Object System.Tags, @{Name='Description'; Expression={$_.('System.Description') -replace '<[^>]*>', ''}}

$availableData = @{}

foreach ($x in $existingItemsList) {
    $tag = $x.'System.Tags'
    $description = $x.Description
     # Check if the tag is not null or empty
     if (-not [string]::IsNullOrEmpty($tag)) {
        $tag = $tag.Trim()
        $availableData[$tag] = $description
    } else {
        Write-Host "Skipping empty tag while checking exists records | $description"
    }
}
$currentDir = Get-Location
# Define the path to available.json (same directory as data.json)
$availableJsonFilePath = Join-Path $currentDir "available.json"
# Convert to JSON and save to file
$availableData | ConvertTo-Json -Depth 2 | Set-Content -Path $availableJsonFilePath -Force
# Write-Host "available.json has been created successfully."
