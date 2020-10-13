function Export-ATPCustomDetections {

    Param(
        [string] $OutputFolder = $( if ($IsLinux -or $IsMacOS) { $Env:HOME + "/Desktop/ATPCustomDetections" } else { $env:USERPROFILE + "\Desktop\ATPCustomDetections" }),
        [string] $BearerTokenFile = "$PSScriptRoot\bearerToken.txt",
        [string] $Region = "eus",
        [switch] $Overwrite = $false
    )

    $bearerToken = Get-Content $bearerTokenFile
    $headers = @{}
    $headers.Add("User-Agent", "Mozilla/5.0 (Macintosh; Intel Mac OS X 10.15; rv:72.0) Gecko/20100101 Firefox/72.0")
    $headers.Add("Authorization", $bearerToken)
    $headers.Add("Content-Type", "application/json")
    $queriesUrl = "https://ineportalapi-$Region-prd.securitycenter.windows.com/hunting/queries/"
    $rulesUrl = "https://ineportalapi-$Region-prd.securitycenter.windows.com/hunting/rules/byquery/"

    $urls = [ordered]@{
        "scheduled" = "https://ineportalapi-$Region-prd.securitycenter.windows.com/hunting/queries/scheduled"
        "shared"    = "https://ineportalapi-$Region-prd.securitycenter.windows.com/hunting/queries/shared"
        "user"      = "https://ineportalapi-$Region-prd.securitycenter.windows.com/hunting/queries/user"
    }

    # Create needed directory structure for output files
    New-Item -Path "$OutputFolder\meta" -ItemType Directory -Force 
    New-Item -Path "$OutputFolder\scheduled\rules" -ItemType Directory -Force 
    New-Item -Path "$OutputFolder\scheduled\alerts" -ItemType Directory -Force 
    New-Item -Path "$OutputFolder\shared" -ItemType Directory -Force 
    New-Item -Path "$OutputFolder\user" -ItemType Directory -Force 

    function Get-Query ($url, $queryId, $outpath) {
        $outFile = Join-Path $outpath "$queryId.json"
        try {
            $response = Invoke-WebRequest -Uri $url -Headers $headers -OutFile $outFile
        }
        catch [System.Exception] {
            throw $_
        }
    }

    function Get-Queries {
        foreach ($type in $urls.Keys) {
            Write-Host -fore green "Fetching $type queries"
            $url = $urls[$type]
            $outFile = Join-Path $OutputFolder "/meta/$type.json"
            $response = Invoke-WebRequest -Uri $url -Headers $headers -OutFile $outFile
            $json = ConvertFrom-Json (Get-Content $outFile)
            foreach ($id in $json.Id) {
                $outPath = Join-Path $OutputFolder $type
                $queryFile = Join-Path $outPath "$id.json"
                if ($overwrite -or (-not (Test-Path $queryFile)) ) { 
                    $url = $queriesUrl + $id
                    Get-Query $url $id $outPath
                    Write-Host -Fore cyan "Downloading $type query details for query ID $id"
                }
                else { Write-Host -fore yellow "Skipping getting $type query details for query ID $id, file already exits. Use -overwrite flag to force download" }
            }
        }
    }

    function Get-Rules {
        Write-Host -fore green "Fetching rules for scheduled queries"
        $type = "scheduled"
        $outPath = Join-Path $OutputFolder $type
        Get-ChildItem -Path $outPath -Filter *.json |
        Foreach-Object {
            $id = $_.Name.split(".")[0]
            $outPath = Join-Path $OutputFolder "$type/rules"
            $ruleFile = Join-Path $outPath "$id.json"
            if ($overwrite -or (-not (Test-Path $ruleFile)) ) {
                $url = $rulesUrl + $id + "?includeLastRun=true"
                Get-Query $url $id $outPath
                Write-Host -Fore cyan "Downloading rules for scheduled query ID $id"
            }
            else { Write-Host -fore yellow "Skipping fetching rules for query ID $id, file already exits. Use -overwrite flag to force download" }
        }
    }

    function Get-AssociatedAlerts {
        Write-Host -fore green "Fetching alerts for scheduled queries"
        $type = "scheduled"
        $baseOutPath = Join-Path $OutputFolder $type
        $outPath = Join-Path $baseOutPath "alerts"
        Get-ChildItem -Path "$baseOutPath/rules" -Filter *.json |
        Foreach-Object {
            $id = $_.Name.split(".")[0]
            $alertFile = Join-Path $outPath "$id.json"
            $json = Get-Content $_.FullName | ConvertFrom-Json
            $IoaDefinitionId = $json.IoaDefinitionId
            if ($overwrite -or (-not (Test-Path $alertFile)) ) {
                $alertsUrl = "https://ineportalapi-$Region-prd.securitycenter.windows.com/AssociatedAlerts?page_size=100&ordering=-lasteventtime&alertType=Alert&lookBackInDays=30&IoaDefinitionIds=$IoaDefinitionId&pageSize=100&sortByField=lasteventtime&sortOrder=Descending"
                Get-Query $alertsUrl $id $outPath
                Write-Host -Fore cyan "Downloading alerts associated with query ID $id"
            }
            else { Write-Host -fore yellow "Skipping getting assoicated alerts for query ID $id, file already exits. Use -overwrite flag to force download" }
        }
    }

    function Get-AlertInfo ($json) {
        $numAlerts = 0
        $latestAlert = ""
        foreach ($alert in $json.Items) {
            $numAlerts = $numAlerts + 1
            if ($latestAlert -eq "" -or ($alert.LastEventTime -gt $latestAlert)) { $latestAlert = $alert.LastEventTime }
        }
        $numAlerts, $latestAlert
    }

    function Invoke-WriteToCSV {
        foreach ($type in $urls.Keys) {
            $outPath = Join-Path $OutputFolder $type
            $reportName = "Custom-Detections-Report.csv"
            if ($type -ne "scheduled") { $reportName = "$((Get-Culture).TextInfo.ToTitleCase($type))-Detections-Report.csv" }
            $LogPath = Join-Path $OutputFolder $reportName
            Remove-Item $LogPath -Force -ErrorAction ignore
            Get-ChildItem -Path $outPath -Filter *.json |
            Foreach-Object {
                $id = $_.Name.split(".")[0]
                $jq = Get-Content $_.FullName | ConvertFrom-Json
                if ($type -eq "scheduled") {
                    $outPath = Join-Path $OutputFolder "$type/rules"
                    $ruleFile = Join-Path $outPath "$id.json"
                    $alertFile = Join-Path $OutputFolder "$type/alerts/$id.json"
                    $jr = Get-Content $ruleFile | ConvertFrom-Json
                    $ja = Get-Content $alertFile | ConvertFrom-Json
                    $numAlerts, $latestAlert = Get-AlertInfo $ja        
                    [PSCustomObject][ordered]@{ 
                        "Id"                   = $jq.Id;
                        "Name"                 = $jq.Name; 
                        "QueryText"            = $jq.QueryText;
                        "IsEnabled"            = $jr.IsEnabled; 
                        "IntervalHours"        = $jr.IntervalHours; 
                        "LastRunTime"          = $jr.LastRunTime; 
                        "CreatedBy"            = $jr.CreatedBy; 
                        "Severity"             = $jr.Severity;
                        "# of Alerts"       = $numAlerts; 
                        "Latest Alert" = $latestAlert; 
                    } | Export-Csv -Path $LogPath -NoTypeInformation -Append
                }
                else {
                    [PSCustomObject][ordered]@{ 
                        "Id"        = $jq.Id;
                        "Name"      = $jq.Name; 
                        "QueryText" = $jq.QueryText;

                    } | Export-Csv -Path $LogPath -NoTypeInformation -Append
                }
            }
        }
    }

    try {
        Get-Queries
        Get-Rules
        Get-AssociatedAlerts
        Invoke-WriteToCSV
    }
    catch [System.Exception] {
        $responseCode = $_.Exception.Response.StatusCode.Value__
        if ($responseCode -eq "401") {
            Write-Host -ForegroundColor red "Your bearer token has expired, try updating it and starting the script again. Exiting now."
        }
        else {
            Write-Host -ForegroundColor red $_.Exception
        }
    }

    Write-Host -fore green "DONE"
}