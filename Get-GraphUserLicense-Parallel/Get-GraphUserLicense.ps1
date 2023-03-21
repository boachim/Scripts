<#
.SYNOPSIS
    Generates tenant-wide user license assignment report by creating one Excel-sheet (CSV) per user license type.
#>

[CmdletBinding()]param()

function GetProductInfo() {

    # Get the latest product names. If you wish to avoid downloading the csv, download the csv manually and comment out the Invoke-WebRequest line.
    # CSV Source: https://learn.microsoft.com/en-us/azure/active-directory/enterprise-users/licensing-service-plan-reference
    Write-Verbose 'Getting latest Microsoft product names and service plan identifiers'
    $m365licenseInfoCsvPath = Join-Path $licenseReportPath 'M365LicenseInfo.csv' -Verbose
    $ProgressPreference = 'SilentlyContinue' 
    Invoke-WebRequest -Uri 'https://raw.githubusercontent.com/merill/license/main/license.csv' -OutFile $m365licenseInfoCsvPath

    $m365licenseInfoCsvPath = Import-Csv $m365licenseInfoCsvPath -Verbose

    foreach ($line in $m365licenseInfoCsvPath) {
        if (!$productNames.ContainsKey($line.GUID)) {
            $productNames[$line.GUID] = $line.Product_Display_Name
        }
        $planKey = $line.GUID + $line.Service_Plan_Id
        if (!$planNames.ContainsKey($planKey)) {
            $planNames[$planKey] = $line.Service_Plans_Included_Friendly_Names + '(' + $line.Service_Plan_Name + ')'
        }
    }
}

function GetLicenseTypesInTenant {
    # Get a list of all licences that exist within the tenant
    Write-Verbose 'Getting license types.'
    $licenseTypes = Get-MgSubscribedSku | `
        Select-Object -Property SkuId, SkuPartNumber, ConsumedUnits, ServicePlans -ExpandProperty PrepaidUnits | `
        Where-Object { $_.ConsumedUnits -ge 1 }
    $licenseTypeSummaryReportPath = Join-Path $licenseReportPath ($date + '__LicenseSummary.csv') -Verbose

    $licenseTypeSummary = @()
    foreach ($item in $licenseTypes) {
        $itemInfo = [ordered]@{
            SkuId         = $item.SkuId
            SkuPartNumber = $item.SkuPartNumber
            Name          = GetProductDisplayName -skuId $item.SkuId -productNames $productNames
            Total         = $item.Enabled
            Assigned      = $item.ConsumedUnits
            Available     = ($item.Enabled - $item.ConsumedUnits)
        } 
        $licenseTypeSummary += $itemInfo
    }

    $licenseTypeSummary | Export-Csv $licenseTypeSummaryReportPath

    return $licenseTypes
}

Import-Module .\GraphUserLicense.psm1

#Import-Module Microsoft.Graph.Users
#Import-Module Microsoft.Graph.Identity.DirectoryManagement
Connect-MgGraph -Scopes Directory.Read.All

$ScriptStart = Get-Date
Write-Verbose ('START TIME: ' + $ScriptStart)

$graphRequest = 'https://graph.microsoft.com/v1.0/users/'

$pageCount = 999
$cacheFilePath = './script.cache'
$currentCount = 0
$resume = $false
$percentComplete = 0

$licenseReportPath = '.\'
#$licenseReportPath = Get-Location
$date = Get-Date -Format 'yyyyMMdd'
$productNames = @{}
$planNames = @{}

$uri = "$($graphRequest)?`$top=$pageCount&`$filter=assignedLicenses/`$count ne 0&`$count=true&`$Select=UserPrincipalName,DisplayName,JobTitle,Office,AssignedLicenses,AssignedPlans"

$m365licenseInfoCsv = GetProductInfo

$licenseTypes = GetLicenseTypesInTenant

if((Test-Path -Path $cacheFilePath -PathType Leaf)){ ## Cache exists from previous run
    $resumePrompt = Read-Host "Do you want to resume from where the script was interrupted? [y/n]"
    $resume = $resumePrompt -match "[yY]"
    if($resume) {
        $cacheInfo = Get-Content -Path $cacheFilePath | ConvertFrom-Json
        $uri = $cacheInfo.nextLink
        $currentCount = $cacheInfo.currentCount
        $totalCount = $cacheInfo.totalCount
    }
}

if(!$resume){  
    Write-Progress -Activity "Getting items" -Status "Looking up total count"
    $totalCount = (Invoke-GraphRequest -Uri "$($graphRequest)?`$count=true&$`$top=1" -Headers @{ConsistencyLevel='eventual'}).'@odata.count'
}

Write-Progress -Activity "Getting items"

do {

    if($null -ne $licensedUsers.'@odata.nextLink') {
        $currentCount += $pageCount
        $cacheInfo = @{
            nextLink = $licensedUsers.'@odata.nextLink'
            currentCount = $currentCount
            totalCount = $totalCount
        }
        $cacheInfo | ConvertTo-Json | Out-File -FilePath $cacheFilePath
        $licensedUsers = Invoke-GraphRequest -Uri $licensedUsers.'@odata.nextLink'
    }
    else 
    {
        $licensedUsers = Invoke-GraphRequest -Uri $uri -Headers @{ConsistencyLevel='eventual'}
    }

    if($totalCount -gt $pageCount) { 
        $percentComplete =  (($currentCount / $totalCount) * 100)
        $statusMessage = "[$percentComplete% $currentCount / $totalCount]"
    }
    Write-Progress -Activity "Getting items" -Status "$statusMessage $($item.displayName)" -PercentComplete $percentComplete

    $job = $licenseTypes | ForEach-Object -Parallel {

        Import-Module .\GraphUserLicense.psm1

        $licenseType = $_
        $licensedUsers = $using:licensedUsers
        $licenseReportPath = $using:licenseReportPath
        $productNames = $using:productNames
        $planNames = $using:planNames
        $date = $using:date

        $userLicenseSummaryReportPath = Join-Path $licenseReportPath ($date + '_' + $licenseType.SkuPartNumber + '.csv') 
        Write-Verbose "Generating $userLicenseSummaryReportPath"

        $users = $licensedUsers.value | Where-Object { $_.AssignedLicenses.SkuId -contains $licenseType.SkuId }

        $productDisplayName = GetProductDisplayName $licenseType.SkuId -productNames $productNames

        $userLicenseSummary = @()
        foreach ($user in $users) {

            $itemInfo = [ordered]@{
                Date              = $date
                UserPrincipalName = $user.UserPrincipalName
                DisplayName       = $user.DisplayName + ''
                JobTitle          = $user.JobTitle + ''
                Office            = $user.Office + ''
                License           = $licenseType.SkuPartNumber
                LicenseName       = $productDisplayName
            }
            $userLicense = $user.AssignedLicenses | Where-Object { $_.SkuId -eq $licenseType.SkuId }
            foreach ($plan in $licenseType.ServicePlans) {
                # Set status of disabled plans in report
                $status = $true
                if ($userLicense.DisabledPlans -contains $plan.ServicePlanId) {
                    $status = $false
                }
                $spName = GetServicePlanDisplayName $licenseType.SkuId $plan.ServicePlanId $plan.ServicePlanName
                $itemInfo[$spName] = $status
            }
            $userLicenseSummary += $itemInfo
        }    
        $userLicenseSummary | Export-Csv -Path $userLicenseSummaryReportPath -Append
    } -ThrottleLimit 3 -AsJob

    $job | Receive-Job -Wait

} while ($null -ne $licensedUsers.'@odata.nextLink') 

$ScriptEnd = (Get-Date)
$RunTime = New-Timespan -Start $ScriptStart -End $ScriptEnd
$Total = "Elapsed Time: {0}:{1}:{2}:{3}" -f $RunTime.Hours, $Runtime.Minutes, $RunTime.Seconds, $RunTime.Milliseconds
Write-Verbose "END TIME: $ScriptEnd"
Write-Verbose $Total
Write-Verbose 'Script Completed.'