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
    $Organization = "organization"
    $Project = "project"
    $url = "https://dev.azure.com/$Organization/$Project/_apis/wit/workitems/$"+"${Type}?api-version=7.1"


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

# Inputs for the Function
$OCRTitle = $jsonContent.OCRTitle
$AssignedTo = ""  # User to whom the task will be assigned
$State = "New"  # State of the work item
$Tags = $jsonContent.OCRNo
$AreaPath = "Dummy"
$iterationPath = "CE\Dummy\Dummy 2024"
$Desc = $jsonContent.Desc
$OCRType = $jsonContent.OCRType
$OCRDocType = $jsonContent.OCRDocType #Data for Functional Area
$Priority = $jsonContent.Priority
$PAT = "UPDATEIT"
$Type = "User Story"

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