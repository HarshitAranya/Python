[System.Net.WebRequest]::DefaultWebProxy.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials

function Create-WorkItem {
    param (
        [string]$PAT
        ,[string]$Type # e.g., Task, Bug, User Story
        # ,[string]$Title
        # ,[string]$AssignedTo
        # ,[string]$State
        # ,[string]$Tags
		# ,[string]$AreaPath
    )

    # Authorization Header
    $Base64AuthInfo = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(":$PAT"))
    $Headers = @{
        Authorization = "Basic $Base64AuthInfo"
        "Content-Type" = "application/json-patch+json"
    }

    # API Endpoint
    $Organization = "civica-cp"
    $Project = "CE"
    $url = "https://dev.azure.com/$Organization/$Project/_apis/wit/workitems/$"+"${Type}?api-version=7.1"

    # API Endpoint
    #$url = "https://dev.azure.com/civica-cp/CE/_apis/wit/workitems/${Type}?api-version=7.1"

    # Request Body for creating the Work Item
    $Body = @(
        @{ "op" = "add"; "path" = "/fields/System.Title"; "value" = $OCRTitle }
        ,@{ "op" = "add"; "path" = "/fields/System.AssignedTo"; "value" = $AssignedTo }
        ,@{ "op" = "add"; "path" = "/fields/System.State"; "value" = $State }
        ,@{ "op" = "add"; "path" = "/fields/System.Tags"; "value" = $Tags }
		,@{ "op" = "add"; "path" = "/fields/System.AreaPath"; "value" = $AreaPath }
		,@{ "op" = "add"; "path" = "/fields/System.IterationPath"; "value" = $iterationPath }        
		,@{ "op" = "add"; "path" = "/fields/System.Description"; "value" = $Desc }
		,@{ "op" = "add"; "path" = "/fields/Microsoft.VSTS.Common.AcceptanceCriteria"; "value" = $OCRType }
		,@{ "op" = "add"; "path" = "/fields/CivicaAgile.FunctionalArea"; "value" = $OCRDocType }
		,@{ "op" = "add"; "path" = "/fields/Microsoft.VSTS.CMMI.ImpactAssessmentHtml"; "value" = "NA" }
		,@{ "op" = "add"; "path" = "/fields/Microsoft.VSTS.Common.Priority"; "value" = $Priority }
    ) | ConvertTo-Json -Depth 10

    # Execute the POST Request
    try {
        $response = Invoke-RestMethod -Uri $url -Headers $Headers -Method POST -Body $Body
        Write-Host "$($Tags) | Work Item Created Successfully. ID: $($response.id)" -ForegroundColor Green
        # Write-Output $response
    }
    catch {
        Write-Host "Error Creating Work Item: $($_.Exception.Message)" -ForegroundColor Red
    }
}

$currentDir = Get-Location

# Define the path to the data.json file
$jsonFilePath = Join-Path $currentDir "data.json"
# Check if the file exists
if (Test-Path $jsonFilePath) {
    $jsonContent = Get-Content -Path $jsonFilePath | ConvertFrom-Json
} else {
    Write-Host "The file 'data.json' does not exist in the current folder."
}

# Determine if the script is running as a standalone executable or as a script from a file
if ($PSScriptRoot -eq "") {
    # If running as an executable, $PSScriptRoot will be empty, so we get the directory of the running executable
    $currentDir = [System.IO.Path]::GetDirectoryName($MyInvocation.MyCommand.Definition)
} else {
    # If running as a script, $PSScriptRoot will be the folder where the script is located
    $currentDir = $PSScriptRoot
}

$henvFilePath = Join-Path $currentDir "henv.json"

if (Test-Path $henvFilePath) {
    $henvContent = Get-Content -Path $henvFilePath | ConvertFrom-Json
} else {
    Write-Host "The file 'henv.json' does not exist in the current folder."
}

# $henvFilePath = Join-Path $currentDir "henv.json"
# if (Test-Path $henvFilePath) {
#     $henvContent = Get-Content -Path $henvFilePath | ConvertFrom-Json
# } else {
#     Write-Host "The file 'henv.json' does not exist in the current folder."
# }
# <#
# Inputs for the Function
$OCRTitle = $jsonContent.OCRTitle
# $AssignedTo = ""  # User to whom the task will be assigned
$AssignedTo = $henvContent.AssignedTo  # User to whom the task will be assigned
# $State = "New"  # State of the work item
$State = $henvContent.State  # State of the work item
$Tags = $jsonContent.OCRNo
# $AreaPath = "CE\CJS Change Management"
$AreaPath = $henvContent.AreaPath
# $iterationPath = "CE\CJS Change Management\CJS Change Management 2024"
$iterationPath = $henvContent.IterationPath
$Desc = $jsonContent.Desc
$OCRType = $jsonContent.OCRType
$OCRDocType = $jsonContent.OCRDocType #Data for Functional Area
$Priority = $jsonContent.Priority
# $PAT = "CdD7vTeUJtaurHKbTiYYEG6Z0dIVSHh22zC1XDnV1IXQK2UUT8nCJQQJ99BAACAAAAANjyhyAAASAZDOOLeO"
$PAT = $henvContent.PAT
# $Type = "User Story"
$Type = $henvContent.Type
# >

# $jsonContent
# $henvContent

if (-not [string]::IsNullOrWhiteSpace($OCRTitle) -and
    -not [string]::IsNullOrWhiteSpace($Tags) -and
    -not [string]::IsNullOrWhiteSpace($AreaPath) -and
    -not [string]::IsNullOrWhiteSpace($IterationPath) -and    
    -not [string]::IsNullOrWhiteSpace($Desc) -and
    -not [string]::IsNullOrWhiteSpace($OCRType) -and
    -not [string]::IsNullOrWhiteSpace($OCRDocType) -and
    -not [string]::IsNullOrWhiteSpace($Priority) -and
    -not [string]::IsNullOrWhiteSpace($PAT) -and
    -not [string]::IsNullOrWhiteSpace($Type)) {
    
    # Call the function
    Create-WorkItem -PAT $PAT -Type $Type
} else {
    Write-Host "Error: One or more required variables are not set or invalid."
}