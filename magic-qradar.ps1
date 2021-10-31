<#
        03/06/2021 (dd/MM/YYYY)
        IBM QRADAR API version 14.0

        Requirement : 
            - importExcel module (powershell)
            - rights granted to the qradar SEC token (https://www.ibm.com/docs/en/qsip/7.3.3?topic=api-restful-overview)

#>

param(
    [Parameter(Mandatory)] $qradar_api_key="",
    $start_date = (Get-Date "01/01/1970"),
    $end_date = (get-date),
    $ReferenceSet_json = '',
    $ReferenceSet_csv = '',
    $delimiter = ";",
    $listReferenceSet =$null,
    [switch] $defaultTTL=$false,
    $ttl_year,
    $ttl_month,
    $ttl_day,
    $ttl_hour,
    $ttl_min,
    $ttl_sec,
    $searchID,
    [switch]$checkSearchStatus,
    [switch] $cancelSearch,
    [switch] $deleteSearch,
    $runSearch,
    [switch] $saveSearch,
    [switch] $listSearch,
    [switch] $test=$false
)



$qradar_api_url = '' #https://example_url_qradar.com/api/
$headers = @{}
$headers.Add('Accept','application/json')
$headers.Add('SEC',$qradar_api_key)

$start_epoch_time = Get-Date "01/01/1970"

$tlp_red_ttl = 12
$tlp_amber_ttl = 9
$tlp_green_ttl = 6
$tlp_white_ttl = 3

function Get-epochTime{
    param(
        $end_date = (Get-Date)
    )

    $current_epoch_time = (New-TimeSpan -Start $start_epoch_time -End ($end_date))
    return $current_epoch_time
}

function Get-epochTime_seconds{
    param(
        $end_date = (Get-Date)
    )

    $current_epoch_time = Get-epochTime -end_date $end_date
    $current_epoch_time_seconds = [int64] $current_epoch_time.TotalSeconds
    return $current_epoch_time_seconds
}

function Get-epochTime_milliseconds{
    param(
        $end_date = (Get-Date)
    )

    $current_epoch_time = Get-epochTime -end_date $end_date
    $current_epoch_time_milliseconds = [int64] $current_epoch_time.TotalMilliseconds
    return $current_epoch_time_milliseconds
}

function convert_epochtime_seconds{
    param([parameter (mandatory)]$epoch_time_to_convert)
    $converted_epochtime = $start_epoch_time.Addseconds($epoch_time_to_convert)
    return ($converted_epochtime)
}

function convert_epochtime_milliseconds{
    param([parameter (mandatory)]$epoch_time_to_convert)
    $converted_epochtime = $start_epoch_time.AddMilliseconds($epoch_time_to_convert)
    return ($converted_epochtime)
}

Function object_toString{
param(
    [parameter(mandatory)]$array,
    $delimiter ="-"
    )
    $out=""
    foreach ($value in $array){
        $out += [string]$value + "-"
    }

    $out=$out.Substring(0,$out.Length-1)

    return $out
}



Function Csv_creation{
    param(
        $path ="kpi_qradar_out.csv",
        #[Parameter(ValueFromPipeline=$true)] $value,
        [Parameter(mandatory)] $value,
        $delimiter=";",
        $encoding = "UTF8"
    )
    $value | Export-Csv -Path $path -Delimiter $delimiter -NoTypeInformation -Encoding $encoding
}

Function ImportExcel_initialisation{
    try{Install-Module -name ImportExcel -Force}
    catch{write-host "ImportExcel module is not installed and must be installed as an administrator"}
}

Function Csv_to_xlsx{
    Param(
        [parameter(mandatory)]$csv_path,
        $output = "output.xlsx",
        $delimiter = ";",
        $worksheetName = $csv_path,
        $encoding = "UTF8"
    )

    import-csv -Delimiter $delimiter -Encoding $encoding -Path $csv_path | export-excel 
}

Function Create_chart_excel{
    Param(
        [parameter(mandatory)]$value,
        [parameter(mandatory)]$output,
        [parameter(mandatory)]$chartType,
        [parameter(mandatory)]$worksheetName,
        $seriesHeader,
        $title ="no title",
        [parameter(mandatory)]$Yrange,
        [parameter(mandatory)]$Xrange,
        $chart_col=0,
        $chart_row=8,
        [switch] $noLegend = $false,
        $height=225,
        [switch] $show = $false,
        [switch] $autosize=$false,
        [switch] $ShowCategory=$false,
        [switch] $ShowPercent=$false
    )

    $params = @{
    # Spreadsheet Properties
    Path                 = $output
    AutoSize             = $autosize
    AutoFilter           = $true
    BoldTopRow           = $true
    FreezeTopRow         = $true
    WorksheetName        = $worksheetName
    PassThru             = $true
    }  
    
    $ExcelPackage = $value | Export-Excel @params
    $WorkSheet = $ExcelPackage."$worksheetName"

    $chartDefinition = @{
        YRange   = $Yrange
        XRange   = $Xrange
        seriesHeader = $seriesHeader
        Title    = $title
        Column   = $chart_col
        Row      = $chart_row
        NoLegend = $noLegend
        Height   = $height
        ChartType= $chartType
        ShowCategory= $ShowCategory
        ShowPercent=$ShowPercent
    }

    $Chart = New-ExcelChartDefinition  @chartDefinition 

    $params = @{
    ExcelPackage         = $ExcelPackage
    ExcelChartDefinition = $chart
    WorksheetName        = $WorkSheet
    AutoNameRange        = $true
    Show                 = $show
    }

    Export-Excel @params
}

Function PivotChartTable{
    param(
        [parameter(mandatory)] $output,
        [parameter(mandatory)] $data,
        [parameter(mandatory)] $WorkSheetName,
        [parameter(mandatory)] $TableName,
        [parameter(mandatory)] $PivotRows,
        [parameter(mandatory)] $PivotData,
        [parameter(mandatory)] $PivotColumns,
        $charType = "ColumnStacked",
        $PivotDataAction="Count"
    )

    $excelSplat = @{
        Path               = $output
        WorkSheetName      = $WorkSheetName
        TableName          = $TableName
        Autosize           = $true
        IncludePivotTable  = $true
        PivotRows          = $PivotRows
        PivotData          = @{$PivotData=$PivotDataAction}#$PivotData
        PivotColumns       = $PivotColumns
        IncludePivotChart  = $true
        ChartType          = $charType
    }
 
    $data |Export-Excel @excelSplat
}


Function Create-ReferenceSet{
    Param ($RSName, $RSType, 
        $ttl_year,
        $ttl_month,
        $ttl_day,
        $ttl_hour,
        $ttl_min,
        $ttl_sec,
        $timeout_type = "UNKNOWN"
        )

        $ttl="&time_to_live="
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    if($null -ne $ttl_year) {$ttl +=$ttl_year.toString() + " year "}
    if($null -ne $ttl_month) {$ttl +=$ttl_month.toString() + " month "}
    if($null -ne $ttl_day) {$ttl +=$ttl_day.toString() + " day "}
    if($null -ne $ttl_hour) {$ttl +=$ttl_hour.toString() + " hour "}
    if($null -ne $ttl_min) {$ttl +=$ttl_min.toString() + " min "}
    if($null -ne $ttl_sec) {$ttl +=$ttl_sec.toString() + " sec "}
    $ttl = [uri]::EscapeUriString($ttl)
    if($ttl -eq "&time_to_live="){$ttl=""}

    $url = $qradar_api_url + 'reference_data/sets?element_type='+ $RSType +'&name='+ $RSName + $ttl + "&timeout_type="+$timeout_type

    $response = Invoke-RestMethod -Method Post $url -Headers $headers
    Write-host "$RSName has been created"
    return $response
}

Function Set-ReferenceSet{
Param ($RSName, $RSvalue)

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12


$url = $qradar_api_url + 'reference_data/sets/'+$RSName+'?value='+$RSvalue

$response = Invoke-RestMethod -Method Post $url -Headers $headers

Write-verbose "$RSName has been filled"
return $response
}

Function Update-ReferenceSet{
    Param ($RSName, $RSvalue, $RSAvailable)
    if($null -ne $RSAvailable.data)
    {
        if($RSAvailable.data.value.contains($RSvalue)) 
            {Write-Verbose "$RSvalue already exists in $RSName"}
        else {
            Set-ReferenceSet -RSName $RSName -RSvalue $RSvalue
            Write-verbose "$RSName has been updated with new data"}
    }
    else {
        write-verbose "$RSName is empty"
        Set-ReferenceSet -RSName $RSName -RSvalue $RSvalue
        Write-verbose "$RSName has been updated with new data"}
}

Function Get-ReferenceSet{
Param ($RSName)

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

if($RSName -eq $null)
{
    $sub_url = ''
}
else 
{
    $sub_url ='/'+ $RSName
}

$url = $qradar_api_url + 'reference_data/sets'+$sub_url

$Error.Clear()
try{$response = Invoke-RestMethod -Method Get $url -Headers $headers}
catch { 
    Write-verbose "Error occured in Get-ReferenceSet function"
    }
return $response
}

Function Check-ReferenceSet{
Param ($RSName)

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12


$url = $qradar_api_url + 'reference_data/sets/'+ $RSName

$Error.Clear()
try{$response = Invoke-RestMethod -Method Get $url -Headers $headers}
catch { 
    Write-verbose "Error occured in Check-ReferenceSet function"
    }
if (!$Error) { Write-host "$RSName already exists"
return $true}
else {    Write-host "$RSName does not exist"
    return $false}

}

Function Get-buildingBlockList{

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12


$url = $qradar_api_url + 'analytics/building_blocks'

$Error.Clear()
try{$response = Invoke-RestMethod -Method Get $url -Headers $headers}
catch { 
    Write-verbose "Error occured in Get-buildingBlockList function"
    }
return $response
}

Function Get-buildingBlock{
Param ([parameter(mandatory)] $BuldingBlockid)

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12


$url = $qradar_api_url + 'analytics/building_blocks/' + $BuldingBlockid

$Error.Clear()
try{$response = Invoke-RestMethod -Method Get $url -Headers $headers}
catch { 
    Write-verbose "Error occured in Get-buildingBlock function"
    }
return $response
}

Function Get-rulelist{

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12


$url = $qradar_api_url + 'analytics/rules'

$Error.Clear()
try{$response = Invoke-RestMethod -Method Get $url -Headers $headers}
catch { 
    Write-verbose "Error occured in Get-RuleList function"
    }
return $response
}

Function Get-rule{
Param ([parameter(mandatory)] $Ruleid)

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12


$url = $qradar_api_url + 'analytics/rules/'+ $Ruleid

$Error.Clear()
try{$response = Invoke-RestMethod -Method Get $url -Headers $headers}
catch { 
    Write-verbose "Error occured in Get-ReferenceSet function"
    }
return $response
}

Function Get-closingReasonList{

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12


$url = $qradar_api_url + 'siem/offense_closing_reasons'

$Error.Clear()
try{$response = Invoke-RestMethod -Method Get $url -Headers $headers}
catch { 
    Write-verbose "Error occured in Get-closingReasonList function"
    }
return $response
}

Function Get-closingReason{
param([parameter(mandatory)] $ClosingReasonID)

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12


$url = $qradar_api_url + 'siem/offense_closing_reasons/' + $ClosingReasonID

$Error.Clear()
try{$response = Invoke-RestMethod -Method Get $url -Headers $headers}
catch { 
    Write-verbose "Error occured in Get-closingReasonList function"
    }
return $response
}

Function Get-closingReasonTextCache{
param(
    [parameter(mandatory)] $ClosingReasonID,
    [parameter(mandatory)]$closingReasonList=""
)

$Error.Clear()
try{$response = ($closingReasonList | ?{$_.id -eq $ClosingReasonID}).text}
catch { 
    Write-verbose "Error occured in Get-closingReasonTextCache function"
    }
return $response
}

Function Get-OffenseList{

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12


$url = $qradar_api_url + 'siem/offenses'

$Error.Clear()
try{$response = Invoke-RestMethod -Method Get $url -Headers $headers}
catch { 
    Write-verbose "Error occured in Get-OffenseList function"
    }
return $response
}

Function Get-Offense{
param([parameter(mandatory)] $offenseID)

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12


$url = $qradar_api_url + 'siem/offenses//' + $offenseID

$Error.Clear()
try{$response = Invoke-RestMethod -Method Get $url -Headers $headers}
catch { 
    Write-verbose "Error occured in Get-Offense function"
    }
return $response
}

Function Get-OffenseNoteList{
param([parameter(mandatory)] $offenseID)
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12


$url = $qradar_api_url + 'siem/offenses/' + $offenseID + '/notes'

$Error.Clear()
try{$response = Invoke-RestMethod -Method Get $url -Headers $headers}
catch { 
    Write-verbose "Error occured in Get-OffenseNoteList function"
    }
return $response
}

Function Get-OffenseNote{
param(
    [parameter(mandatory)] $offenseID,
    [parameter(mandatory)] $offenseNoteID
)

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12


$url = $qradar_api_url + 'siem/offenses/' + $offenseID + '/notes/' + $offenseNoteID

$Error.Clear()
try{$response = Invoke-RestMethod -Method Get $url -Headers $headers}
catch { 
    Write-verbose "Error occured in Get-OffenseNote function"
    }
return $response
}

Function Get-rulelist_and_BBlist{
    param(
    $start_date = $start_epoch_time,
    $end_date = (get-date)
    )
    
    $start_date = (Get-epochTime_milliseconds -end_date $start_date)
    $end_date = (Get-epochTime_milliseconds -end_date $end_date)

    $rulelist = Get-rulelist
    $out =  $rulelist | ?{($_.modification_date -ge $start_date) -or ($_.creation_date -ge $start_date) -and ($_.modification_date -le $end_date)} | select *,@{name="creation_date_classic";Expression={(convert_epochtime_milliseconds $_.creation_date)}}, @{name="modification_date_classic";Expression={(convert_epochtime_milliseconds $_.modification_date)}},@{name="Rule_type";Expression={"rule"}}

    $BBlist = Get-buildingblocklist
    $out += $BBlist | ?{($_.modification_date -ge $start_date) -or ($_.creation_date -ge $start_date) -and ($_.modification_date -le $end_date)} | select *,@{name="creation_date_classic";Expression={(convert_epochtime_milliseconds $_.creation_date)}},  @{name="modification_date_classic";Expression={(convert_epochtime_milliseconds $_.modification_date)}},@{name="Rule_type";Expression={"BB"}}

    return $out
}

Function Get-SavedSearchList{

    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    
    
    $url = $qradar_api_url + 'ariel/saved_searches'
    
    $Error.Clear()
    try{$response = Invoke-RestMethod -Method Get $url -Headers $headers}
    catch { 
        Write-verbose "Error occured in Get-SavedSearchList function"
    }
    return $response
}

Function Get-SavedSearch{
    param(
        [parameter(mandatory)] $savedSearchID
    )
    
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    
    
    $url = $qradar_api_url + 'ariel/saved_searches/' + $savedSearchID
    
    $Error.Clear()
    try{$response = Invoke-RestMethod -Method Get $url -Headers $headers}
    catch { 
        Write-verbose "Error occured in Get-SavedSearch function"
    }
    return $response
}

Function Update-SavedSearch{
    param(
        [parameter(mandatory)] $savedSearchID,
        [parameter(mandatory)] $savedSearchBody #JSON
    )
    
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    
    
    $url = " --data-binary "+$savedSearchBody+$qradar_api_url + 'ariel/saved_searches/' + $savedSearchID
    
    $Error.Clear()
    try{$response = Invoke-RestMethod -Method Post $url -Headers $headers}
    catch { 
        Write-verbose "Error occured in Update-SavedSearch function"
    }
    return $response
}

Function Delete-SavedSearch{
    param(
        [parameter(mandatory)] $savedSearchID
    )
    
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    
    
    $url = $qradar_api_url + 'ariel/saved_searches/' + $savedSearchID
    
    $Error.Clear()
    try{$response = Invoke-RestMethod -Method DELETE $url -Headers $headers}
    catch { 
        Write-verbose "Error occured in Delete-SavedSearch function"
    }
    return $response
}

Function Get-SearchList{

    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    
    
    $url = $qradar_api_url + 'ariel/searches'
    
    $Error.Clear()
    try{$response = Invoke-RestMethod -Method Get $url -Headers $headers}
    catch { 
        Write-verbose "Error occured in Get-SearchList function"
    }
    return $response
}

Function Create-Search{
    param(
        [parameter(mandatory)] $AQL
    )
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    
    $AQL = [uri]::EscapeUriString($AQL)

    $url = $qradar_api_url + 'ariel/searches?query_expression=' + $AQL
    
    $Error.Clear()
    try{$response = Invoke-RestMethod -Method Post $url -Headers $headers}
    catch { 
        Write-verbose "Error occured in Create-Search function"
    }
    return $response
}

Function Get-Search{
    param(
        [parameter(mandatory)] $SearchID
    )
    
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    
    
    $url = $qradar_api_url + 'ariel/searches/' + $SearchID

    $Error.Clear()
    try{$response = Invoke-RestMethod -Method Get $url -Headers $headers}
    catch { 
        Write-verbose "Error occured in Get-Search function"
    }
    return $response
}

Function Update-Search{
    param(
        [parameter(mandatory)] $SearchID,
        [switch] $saveResults=$false,
        $status
    )
    
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    

    
    $url = $qradar_api_url + 'ariel/searches/' + $SearchID
    
    if($saveResults){
        $url += "?save_results=true"
        if($status -ieq "CANCELED"){
            $url += "&status=CANCELED"
        }
        else {
            Write-verbose "Status accepts only 'CANCELED' value"
        }
    }
    else{
        if($status -ieq "CANCELED"){
            $url += "?status=CANCELED"
        }
        else {
            Write-verbose "Status accepts only 'CANCELED' value"
        }
    }

    $Error.Clear()
    try{$response = Invoke-RestMethod -Method Post $url -Headers $headers}
    catch { 
        Write-verbose "Error occured in Update-Search function"
    }
    return $response
}

Function Delete-Search{
    param(
        [parameter(mandatory)] $SearchID
    )
    
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    
    
    $url = $qradar_api_url + 'ariel/searches/' + $SearchID
    
    $Error.Clear()
    try{
        $response = Invoke-RestMethod -Method Delete $url -Headers $headers
        Write-Host "The search $SearchID has been deleted"
    }
    catch { 
        Write-verbose "Error occured in Delete-Search function"
    }
    return $response
}

Function Get-SearchMetadata{
    param(
        [parameter(mandatory)] $SearchID
    )
    
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    
    
    $url = $qradar_api_url + 'ariel/searches/' + $savedSearchID + "/metadada"
    
    $Error.Clear()
    try{$response = Invoke-RestMethod -Method Get $url -Headers $headers}
    catch { 
        Write-verbose "Error occured in Get-SearchMetadata function"
    }
    return $response
}

Function Get-SearchResults{
    param(
        [parameter(mandatory)] $SearchID
    )
    
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    
    
    $url = $qradar_api_url + 'ariel/searches/' + $SearchID +"/results"
    
    $Error.Clear()
    try{$response = Invoke-RestMethod -Method Get $url -Headers $headers}
    catch { 
        Write-verbose "Error occured in Get-SearchResults function"
    }
    return $response
}



Function Get-SearchListName{
    param(
        $SearchList = (Get-SearchList)
    )

    $collectionWithItems = @()

    if ($null -eq $SearchList)
    {
        Write-Host ("No search has been found")
    }
    else{
        $SearchList | %{
            $searchinfo = Get-Search -SearchID $_
            $temp = New-Object System.Object
            $temp | Add-Member -MemberType NoteProperty -Name "searchId" -Value $searchinfo.search_id
            $temp | Add-Member -MemberType NoteProperty -Name "AQL query" -Value $searchinfo.query_string
            $temp | Add-Member -MemberType NoteProperty -Name "Search status" -Value $searchinfo.status
            $temp | Add-Member -MemberType NoteProperty -Name "Record count" -Value $searchinfo.record_count
            $temp | Add-Member -MemberType NoteProperty -Name "Progress" -Value $searchinfo.progress
            $temp | Add-Member -MemberType NoteProperty -Name "Query execution time" -Value $searchinfo.query_execution_time
            $collectionWithItems +=$temp
        }
    }
    
    
    return $collectionWithItems
}

Function Data_to_Qradar_ReferenceSet {
Param (
    [parameter(mandatory)] $RSName,
    [parameter(mandatory)] $RStype,
    [parameter(mandatory)] $RSdata,
        $RS_ttl_year,
        $RS_ttl_month,
        $RS_ttl_day,
        $RS_ttl_hour,
        $RS_ttl_min,
        $RS_ttl_sec,
        $RS_timeout_type
)
    switch ( $RStype )
            {
                sha256 {$RS_type = 'ALNIC';break}
                sha1 {$RS_type = 'ALNIC';break}
                md5 {$RS_type = 'ALNIC';break}
                filename {$RS_type = 'ALNIC';break}
                domain {$RS_type = 'ALNIC';break}
                fqdn {$RS_type = 'ALNIC';break}
                email {$RS_type = 'ALNIC';break}
                url {$RS_type = 'ALNIC';break}
                hostname {$RS_type = 'ALNIC';break}
                ip-dst {$RS_type = 'IP';break}
                ip {$RS_type = 'IP';break}
                default {$RS_type = 'ALNIC';break}
            }
            $return_status = Check-ReferenceSet -RSName $RSName
            if($return_status -eq $true)
            {
                $RS_available = Get-ReferenceSet -RSName $RSName
                $count=$RSdata.value.count
                $ptr=1
                ForEach ($element1 in $RSdata.value) {
                    write-host "[$ptr/$count] job"
                    Update-ReferenceSet -RSName $RSName -RSvalue $element1 -RSAvailable $RS_available
                    $ptr++
                }
                Write-Verbose "Check-RefereceSet = true => Reference Set $RSname already exists"
                #write-host "$RS_name need to be updated. Check each value."
            }
            elseif ($return_status -eq $false)
            {
                Create-ReferenceSet -RSName $RSName -RSType $RS_type -ttl_year $RS_ttl_year -ttl_month $RS_ttl_month -ttl_day $RS_ttl_day -ttl_hour $RS_ttl_hour -ttl_min $RS_ttl_min -ttl_sec $RS_ttl_sec -timeout_type $RS_timeout_type
                $count=$RSdata.value.count
                $ptr=1
                ForEach ($element2 in $RSdata.value) {
                    write-host "[$ptr/$count] job"
                    Set-ReferenceSet -RSName $RSName -RSvalue $element2
                    $ptr++
                }
            }
            else { Write-host "Error in if condition : return value = $return_status"}
}

Function jsonMISP_to_Qradar_ReferenceSet {
Param (
    [parameter(mandatory)] $json
)
    
    $info = $json.event.info 
    $published =  $json.event.date
    $dataset = $json.Event.Attribute
    $file_extension = ".csv"

    $RS_name_source = "MISP_event_"+$json.event.id

    $types = $dataset | select type -Unique

    $types | %{
        $current = $_
        $RS_data = $dataset | where {$_.type -match $current.type}
        $RS_type = $current.type
        $RS_name = $RS_name_source +"_"+$current.type

        Data_to_Qradar_ReferenceSet -RSName $RS_name -RStype $RS_type -RSdata $RS_data
    }
}

Function csv_to_Qradar_ReferenceSet {
    Param (
        [parameter(mandatory)] $csv
    )
    $RS_timeout_type = "LAST_SEEN"
    $RS_name_source = $ReferenceSet_csv.split('.')[0]

    $collectionWithItems =@()

    $response=$null

    $APT_list = $csv.threats | select -unique

    $APT_list | %{
        $current_apt = $_
        $RS_data_current_apt = $csv | where {$_.threats -eq $current_apt}
        $IOC_APT = $current_apt.replace(" ","")

        $current_TLP_list = $RS_data_current_apt.tlp | select -unique
        $current_TLP_list | %{
            $current_apt_tlp = $_
            $RS_data_current_apt_tlp = $RS_data_current_apt | where {$_.tlp -eq $current_apt_tlp}

            $current_type_list = $RS_data_current_apt_tlp.type | select -unique

            $current_type_list | %{
                $current_apt_tlp_type = $_
                $RS_data_current_apt_tlp_type = $RS_data_current_apt_tlp | where {$_.type -eq $current_apt_tlp_type}
                $RS_name = "IOC"+"_"+$current_apt_tlp_type+"_"+$RS_name_source +"_"+$IOC_APT+"_"+"TLP-"+$current_apt_tlp
                $RS_name = $RS_name.ToUpper()
                $RS_type = $current_apt_tlp_type
                $RS_data = $RS_data_current_apt_tlp_type


                if($defaultTTL) {
                    $effectiveTTL = ""
                    switch($current_apt_tlp)
                    {
                        "red" {$effectiveTTL = $tlp_red_ttl;break}
                        "amber" {$effectiveTTL = $tlp_amber_ttl;break}
                        "green" {$effectiveTTL = $tlp_green_ttl;break}
                        "white" {$effectiveTTL = $tlp_white_ttl;break}
                    }
                    $response = Data_to_Qradar_ReferenceSet -RSName $RS_name -RStype $RS_type -RSdata $RS_data -RS_ttl_month $effectiveTTL
                }
                else {$response = Data_to_Qradar_ReferenceSet -RSName $RS_name -RStype $RS_type -RSdata $RS_data -RS_ttl_year $ttl_year -RS_ttl_month $ttl_month -RS_ttl_day $ttl_day -RS_ttl_hour $ttl_hour -RS_ttl_min $ttl_min -RS_ttl_sec $ttl_sec -RS_timeout_type $RS_timeout_type}
                
                $temp = New-Object System.Object
                $temp | Add-Member -MemberType NoteProperty -Name "Reference_Set_Name" -Value $RS_name
                $temp | Add-Member -MemberType NoteProperty -Name "Reference_Set_Type" -Value $current_apt_tlp_type
                $temp | Add-Member -MemberType NoteProperty -Name "Reference_Set_Threat" -Value $IOC_APT
                $temp | Add-Member -MemberType NoteProperty -Name "Reference_Set_TLP" -Value $current_apt_tlp
                $collectionWithItems +=$temp
            }

        }
    }
    return $collectionWithItems
}

Function Search-ReferenceSet {
    Param (
        [parameter(mandatory)] $RS_set,
        $start_date 
    )
    if($null -eq $start_date){$start_date = " LAST 10 DAYS "}
    else {$start_date = " START " +$start_date}

    Write-Host ("inside of Search-ReferenceSet")

    $collectionWithItems = @()

    $threats_list = $RS_set | select -Unique Reference_Set_Threat
    foreach ($threats in $threats_list.Reference_Set_Threat){
        $current_RS = $RS_set | ?{$_.Reference_Set_Threat -eq $threats}
        $type_list = $current_RS| select -Unique Reference_Set_Type
        #write-host ($type_list)
        $condition="" 
        $count_type = $type_list.Reference_Set_Type.count
        $ptr_type=1
        foreach ($type in $type_list.Reference_Set_Type){
            $ptr_rs = 1
            $current_RS_by_type = $current_RS | ?{$_.Reference_Set_Type -eq $type}
            $count_rs = $current_RS_by_type.count
            foreach ($rs in $current_RS_by_type){
                $condition += Get-AQLConditions -type $type -RS_name $rs.Reference_Set_Name
                if($ptr_rs -lt $count_rs){
                    $condition += " OR "
                }
                $ptr_rs++
            }
            
           
            if($ptr_type -lt $count_type){
                $condition += " OR "                
            }
            $ptr_type++
            
        }
        $condition += $start_date
        $AQL = "SELECT * FROM events WHERE " + $condition
        $search = Create-Search -AQL $AQL
        #write-host($AQL)

        $temp = New-Object System.Object
        $temp | Add-Member -MemberType NoteProperty -Name "Search_id" -Value $search.search_id
        $temp | Add-Member -MemberType NoteProperty -Name "AQL_query" -Value $search.query_string
        $temp | Add-Member -MemberType NoteProperty -Name "Threat" -Value $threats
        $temp | Add-Member -MemberType NoteProperty -Name "Status" -Value $search.status
        $collectionWithItems +=$temp
    }
    return $collectionWithItems
}


Function Get-AQLConditions {
    Param (
        [parameter(mandatory)] $type,
        [parameter(mandatory)] $RS_name
    )
        $aql_condition =""
    
        switch ( $type )
            {
                sha256 {$aql_condition = "REFERENCESETCONTAINS('$RS_name',`"hash`")";break}#to adapt
                sha1 {$aql_condition = "REFERENCESETCONTAINS('$RS_name',`"hash`")";break}#to adapt
                md5 {$aql_condition = "REFERENCESETCONTAINS('$RS_name',`"hash`")";break}#to adapt
                filename {$aql_condition = "REFERENCESETCONTAINS('$RS_name',`"filename`")";break}#to adapt
                domain {$aql_condition = "REFERENCESETCONTAINS('$RS_name',`"URL`")";break}#to adapt
                fqdn {$aql_condition = "REFERENCESETCONTAINS('$RS_name',`"URL`")";break}#to adapt
                email {$aql_condition = "REFERENCESETCONTAINS('$RS_name',`"sender`")";break} #to adapt
                url {$aql_condition = "REFERENCESETCONTAINS('$RS_name',`"URL`")";break}#to adapt
                hostname {$aql_condition = "REFERENCESETCONTAINS('$RS_name',`"URL`")";break}#to adapt
                ip-dst {$aql_condition = "REFERENCESETCONTAINS('$RS_name',sourceip) OR REFERENCESETCONTAINS('$RS_name',destinationip)";break}#to adapt
                ip {$aql_condition = "REFERENCESETCONTAINS('$RS_name',sourceip) OR REFERENCESETCONTAINS('$RS_name',destinationip)";break}#to adapt
                cve {$aql_condition = "REFERENCESETCONTAINS('$RS_name',`"URL`")";break} #to adapt
                default {$aql_condition = "REFERENCESETCONTAINS('$RS_name',UTF8(payload))";break} #to adapt
            }
    return $aql_condition
}

Function Format_Offense{
    param(
    [parameter(mandatory)]$offense_list_cache=""
    )

        $out= $offense_list_cache | 
        select @{Name="last_persisted_time";Expression={convert_epochtime_milliseconds -epoch_time_to_convert $_.last_persisted_time}},
            username_count,
            description,
            rules,
            rules_id,
            event_count,
            flow_count,
            assigned_to,
            security_category_count,
            follow_up,
            source_address_ids,
            source_count,
            inactive,
            protected,
            closing_user,
            destination_networks,
            source_network,
            category_count,
            @{Name="close_time";Expression={convert_epochtime_milliseconds -epoch_time_to_convert $_.close_time}},
            remote_destination_count,
             @{Name="start_time";Expression={convert_epochtime_milliseconds -epoch_time_to_convert $_.start_time}},
            magnitude,
            @{Name="last_updated_time";Expression={convert_epochtime_milliseconds -epoch_time_to_convert $_.last_updated_time}},
            credibility,
            id,
            categories,
            severity,
            policy_category_count,
            log_sources,
            closing_reason_id,
            closing_reason,
            device_count,
            @{Name="first_persisted_time";Expression={convert_epochtime_milliseconds -epoch_time_to_convert $_.first_persisted_time}},
            offense_type,
            relevance,
            domain_id,
            offense_source,
            local_destination_address_ids,
            local_destination_count,
            status

        return $out
}

Function Format_Offense_epoch_time{
    param(
    [parameter(mandatory)]$offense_list_cache=""
    )

        $out= $offense_list_cache | 
        select last_persisted_time,
            username_count,
            description,
            rules,
            event_count,
            flow_count,
            assigned_to,
            security_category_count,
            follow_up,
            source_address_ids,
            source_count,
            inactive,
            protected,
            closing_user,
            destination_networks,
            source_network,
            category_count,
            close_time,
            remote_destination_count,
            start_time,
            magnitude,
            last_updated_time,
            credibility,
            id,
            categories,
            severity,
            policy_category_count,
            log_sources,
            closing_reason_id,
            closing_reason,
            device_count,
            first_persisted_time,
            offense_type,
            relevance,
            domain_id,
            offense_source,
            local_destination_address_ids,
            local_destination_count,
            status

        return $out
}

Function Format_Offense_closingReason{
    param(
    [parameter(mandatory)]$offense_list_cache="",
    [parameter(mandatory)]$closingReasonList ="",
    [parameter(mandatory)]$rule_list_cache=""
    )

        $out= $offense_list_cache | 
        select *,
            @{Name="closing_reason";Expression={(Get-closingReasonTextCache -closingReasonList $closing_reason_list_cache -ClosingReasonID $_.closing_reason_id)}}

        return $out
}

Function Format_Offense_rules{
    param(
    [parameter(mandatory)]$offense_list_cache="",
    [parameter(mandatory)]$rule_list_cache=""
    )

        $out= $offense_list_cache | 
        select *,
            @{Name="rules_id";Expression={(object_toString -array $_.rules.id)}}

        return $out
}

Function Format_Rule{
    param(
    [parameter(mandatory)]$rule_list_cache=""
    )

        $out= $rule_list_cache | 
        select owner,
            identifier,
            base_host_id,
            capacity_timestamp,
            origin,
            @{Name="creation_date";Expression={convert_epochtime_milliseconds -epoch_time_to_convert $_.creation_date}},
            type,
            enabled,
            @{Name="modification_date";Expression={convert_epochtime_milliseconds -epoch_time_to_convert $_.modification_date}},
            linked_rule_identifier,
            name,
            average_capacity,
            id,
            base_capacity,
            Rule_type

        return $out
}


Function KPI_rules_status{
    param(
        [parameter(mandatory)] $rule_list=""
    )

    $collectionWithItems = @()
    $count=@()
    $count += $rule_list | ?{($_.enabled -eq $true) -and ($_.origin -eq "USER")}
    $count = $count.count
    
    if($count -eq $null) {$count=0}

    $count2=@()
    $count2 += $rule_list | ?{($_.enabled -eq $false) -and ($_.origin -eq "USER")}
    $count2 = $count2.count
    if($count2 -eq $null) {$count2=0}

    $total = $count + $count2

    $temp = New-Object System.Object
    $temp | Add-Member -MemberType NoteProperty -Name "Rules_Owner" -Value "SOC"
    $temp | Add-Member -MemberType NoteProperty -Name "Active_rules_count" -Value $count
    $temp | Add-Member -MemberType NoteProperty -Name "Disabled_rules_count" -Value $count2
    $temp | Add-Member -MemberType NoteProperty -Name "Total" -Value $total

    $collectionWithItems +=$temp

    $count=@()
    $count += $rule_list | ?{($_.enabled -eq $true) -and ($_.origin -eq "SYSTEM")}
    $count = $count.count
    if($count -eq $null) {$count=0}

    $count2=@()
    $count2 += $rule_list | ?{($_.enabled -eq $false) -and ($_.origin -eq "SYSTEM")}
    $count2 = $count2.count
    if($count2 -eq $null) {$count2=0}

    $total = $count + $count2

    $temp = New-Object System.Object
    $temp | Add-Member -MemberType NoteProperty -Name "Rules_Owner" -Value "IBM"
    $temp | Add-Member -MemberType NoteProperty -Name "Active_rules_count" -Value $count
    $temp | Add-Member -MemberType NoteProperty -Name "Disabled_rules_count" -Value $count2
    $temp | Add-Member -MemberType NoteProperty -Name "Total" -Value $total

    $collectionWithItems +=$temp

    $count=@()
    $count += $rule_list | ?{($_.enabled -eq $true) -and ($_.origin -eq "OVERRIDE")}
    $count = $count.count
    if($count -eq $null) {$count=0}

    $count2=@()
    $count2 += $rule_list | ?{($_.enabled -eq $false) -and ($_.origin -eq "OVERRIDE")}
    $count2 = $count2.count
    if($count2 -eq $null) {$count2=0}

    $total = $count + $count2

    $temp = New-Object System.Object
    $temp | Add-Member -MemberType NoteProperty -Name "Rules_Owner" -Value "OVERRIDE"
    $temp | Add-Member -MemberType NoteProperty -Name "Active_rules_count" -Value $count
    $temp | Add-Member -MemberType NoteProperty -Name "Disabled_rules_count" -Value $count2
    $temp | Add-Member -MemberType NoteProperty -Name "Total" -Value $total

    $collectionWithItems +=$temp

    $count=@()
    $count += $rule_list | ?{($_.enabled -eq $true)}
    $count = $count.count
    if($count -eq $null) {$count=0}

    $count2=@()
    $count2 += $rule_list | ?{($_.enabled -eq $false)}
    $count2 = $count2.count
    if($count2 -eq $null) {$count2=0}

    $total = $count + $count2

    $temp = New-Object System.Object
    $temp | Add-Member -MemberType NoteProperty -Name "Rules_Owner" -Value "Total"
    $temp | Add-Member -MemberType NoteProperty -Name "Active_rules_count" -Value $count
    $temp | Add-Member -MemberType NoteProperty -Name "Disabled_rules_count" -Value $count2
    $temp | Add-Member -MemberType NoteProperty -Name "Total" -Value $total

    $collectionWithItems +=$temp

    return $collectionWithItems

}

Function KPI_rule_and_BB_modified{
    param(
        $start_date = $start_epoch_time,
        $end_date = (get-date)
    )
    

    $out= Get-rulelist_and_BBlist -start_date $start_date -end_date $end_date| select id,
    name,
    creation_date_classic,
    modification_date_classic,
    @{Name="Status";Expression={
        if ($_.creation_date_classic -ge (get-date $start_date))
            {"CREATED"}
        else {"MODIFIED"}
    }},
    rule_type

    return $out
}

Function KPI_rule_and_BB_modified_status{
    param(
        [parameter(mandatory)] $rule_list=""
    )

    $collectionWithItems = @()
    $count_rule=@()
    $count_rule += ($rule_list|?{($_.status -eq "CREATED") -and ($_.Rule_type -eq "rule")})
    $count_rule = $count_rule.count

    $count_BB=@()
    $count_BB += ($rule_list|?{($_.status -eq "CREATED") -and ($_.Rule_type -eq "BB")})
    $count_BB = $count_BB.count

    $total_created = $count_rule +$count_BB

    if($count -eq $null) {$count=0}

    $temp = New-Object System.Object
    $temp | Add-Member -MemberType NoteProperty -Name "Status" -Value "created"
    $temp | Add-Member -MemberType NoteProperty -Name "Rule" -Value $count_rule
    $temp | Add-Member -MemberType NoteProperty -Name "BuldingBlock" -Value $count_BB
    $temp | Add-Member -MemberType NoteProperty -Name "count" -Value $total_created

    $collectionWithItems +=$temp

    $count_rule=@()
    $count_rule += ($rule_list|?{($_.status -eq "MODIFIED") -and ($_.Rule_type -eq "rule")})
    $count_rule = $count_rule.count

    $count_BB=@()
    $count_BB += ($rule_list|?{($_.status -eq "MODIFIED") -and ($_.Rule_type -eq "BB")})
    $count_BB = $count_BB.count

    $total_modified = $count_rule +$count_BB

    $temp = New-Object System.Object
    $temp | Add-Member -MemberType NoteProperty -Name "Status" -Value "modified"
    $temp | Add-Member -MemberType NoteProperty -Name "Rule" -Value $count_rule
    $temp | Add-Member -MemberType NoteProperty -Name "BuldingBlock" -Value $count_BB
    $temp | Add-Member -MemberType NoteProperty -Name "count" -Value $total_modified

    $collectionWithItems +=$temp

    $count_rule=@()
    $count_rule += ($rule_list|?{($_.Rule_type -eq "rule")})
    $count_rule = $count_rule.count

    $count_BB=@()
    $count_BB += ($rule_list|?{($_.Rule_type -eq "BB")})
    $count_BB = $count_BB.count

    $total = $count_rule +$count_BB

    $temp = New-Object System.Object
    $temp | Add-Member -MemberType NoteProperty -Name "Status" -Value "Total"
    $temp | Add-Member -MemberType NoteProperty -Name "Rule" -Value $count_rule
    $temp | Add-Member -MemberType NoteProperty -Name "BuldingBlock" -Value $count_BB
    $temp | Add-Member -MemberType NoteProperty -Name "count" -Value $total

    $collectionWithItems +=$temp

    return $collectionWithItems
}

Function KPI_select_partial_offense_dataset {
    Param(
            $start_date = $start_epoch_time,
            $end_date = (get-date),
            [parameter(mandatory)] $offense_list_cache
        )
        $start_date = (Get-epochTime_milliseconds -end_date $start_date)
        $end_date = (Get-epochTime_milliseconds -end_date $end_date)

        $out = $offense_list_cache | ?{($_.start_time -ge $start_date) -and ($_.start_time -le $end_date)}

        return $out
}

Function KPI_offense_by_rule{
    Param(
        $start_date = $start_epoch_time,
        $end_date = (get-date),
        [parameter(mandatory)] $offense_list_cache,
        [parameter(mandatory)] $rule_list_cache
    )
    
    $start_date = (Get-epochTime_milliseconds -end_date $start_date) 
    $end_date = (Get-epochTime_milliseconds -end_date $end_date)

    $offense_list_cache=$offense_list_cache |?{(
        <#($_.close_time -ge $start_date) -or 
        ($_.last_updated_time -ge $start_date) -or#>
        ($_.start_time -ge $start_date) -and
        ($_.start_time -le $end_date)
        )}
    
    $rule_triggered_list =$offense_list_cache.rules
    $rule_id_triggered_list = $rule_triggered_list.id
    $collectionWithItems = @()
    foreach ($rule in $rule_list_cache){
    
        $count=@()
        $count += ($rule_id_triggered_list|?{$_ -eq $rule.id})
        $count = $count.count

        if($count -eq $null) {$count=0}

        $temp = New-Object System.Object
        $temp | Add-Member -MemberType NoteProperty -Name "id" -Value $rule.id
        $temp | Add-Member -MemberType NoteProperty -Name "rule_name" -Value $rule.name
        $temp | Add-Member -MemberType NoteProperty -Name "rule_origin" -Value $rule.origin
        $temp | Add-Member -MemberType NoteProperty -Name "count" -Value $count

        $collectionWithItems +=$temp
    }

    return $collectionWithItems

}

Function KPI_offense_new{
    Param(
        $start_date = $start_epoch_time,
        $end_date = (get-date),
        [parameter(mandatory)]$offense_list_cache=""
    )
    
    $start_date = (Get-epochTime_milliseconds -end_date $start_date)
    $end_date = (Get-epochTime_milliseconds -end_date $end_date)
        
    $pre_out=$offense_list_cache |?{($_.start_time -ge $start_date) -and ($_.start_time -le $end_date)}
    $out= $pre_out| select *
    return $out
}

Function KPI_offense_closed{
    Param(
        $start_date = $start_epoch_time,
        $end_date = (get-date),
        [parameter(mandatory)]$offense_list_cache=""
    )
    
    $start_date = (Get-epochTime_milliseconds -end_date $start_date)
    $end_date = (Get-epochTime_milliseconds -end_date $end_date)

    $pre_out=$offense_list_cache |?{($_.status -eq "CLOSED") -and ($_.close_time -ge $start_date) -and ($_.close_time -le $end_date)} 
    #Write-Verbose $offense_list_cache
    #$pre_out= Get-OffenseList | ?{($_.status -eq "CLOSED") -and ($_.close_time -ge (Get-epochTime_milliseconds -end_date $start_date))} 
    $out= $pre_out #| select *,@{Name="reason";Expression={(Get-closingReasonTextCache -closingReasonList $closingReasonList -ClosingReasonID $_.closing_reason_id)}} -ExcludeProperty closing_reason_id 

    return $out
}

Function KPI_offense_opened{
    Param(
        $start_date = $start_epoch_time,
        $end_date = (get-date),
        [parameter(mandatory)]$offense_list_cache=""
    )
    
    $start_date = (Get-epochTime_milliseconds -end_date $start_date)
    $end_date = (Get-epochTime_milliseconds -end_date $end_date)

    $pre_out=$offense_list_cache |?{($_.status -eq "OPEN") -and ($_.start_time -ge $start_date) -and ($_.start_time -le $end_date)}
    $out= $pre_out| select *

    return $out
}

Function KPI_offense_open_by_rule{
    Param(
        [parameter(mandatory)]$offense_list_cache="",
        [parameter(mandatory)]$rule_list_cache=""
    )
    $collectionWithItems=@()
    $rules_list = $offense_list_cache.rules
    Foreach($rule in $rule_list_cache){
        $offense = ($offense_list_cache|?{$_.rules.id -eq $rule.id})
        if ($offense -eq $null){$offense_id_string="-"}
        else {
        $offense_id_string = (object_toString -array $offense.id)}

        $rule_name = $rule.name
        $count=@()
        $count += $offense
        $count = $count.count

        $temp = New-Object System.Object
        $temp | Add-Member -MemberType NoteProperty -Name "ruleid" -Value $rule.id
        $temp | Add-Member -MemberType NoteProperty -Name "rule_name" -Value $rule_name
        $temp | Add-Member -MemberType NoteProperty -Name "rule_origin" -Value $rule.origin
        $temp | Add-Member -MemberType NoteProperty -Name "Offense_id" -Value $offense_id_string
        $temp | Add-Member -MemberType NoteProperty -Name "count" -Value $count

        $collectionWithItems +=$temp
    }

    return $collectionWithItems
}

Function KPI_offense_status{
        Param(
        $start_date = $start_epoch_time,
        $end_date = (get-date),
        [parameter(mandatory)]$offense_list_cache=""
    )

    $collectionWithItems = @()
    $count=@()
    $count += (KPI_offense_new -start_date $start_date -end_date $end_date -offense_list_cache $offense_list_cache)
    $count = $count.count
    if($count -eq $null) {$count=0}

    $temp = New-Object System.Object
    $temp | Add-Member -MemberType NoteProperty -Name "Offense_status" -Value "new"
    $temp | Add-Member -MemberType NoteProperty -Name "count" -Value $count

    $collectionWithItems +=$temp

    $count=@()
    $count += (KPI_offense_closed -start_date $start_date -end_date $end_date -offense_list_cache $offense_list_cache)
    $count = $count.count
    if($count -eq $null) {$count=0}

    $closed_handled=$count

    $temp = New-Object System.Object
    $temp | Add-Member -MemberType NoteProperty -Name "Offense_status" -Value "close"
    $temp | Add-Member -MemberType NoteProperty -Name "count" -Value $count

    $collectionWithItems +=$temp

    $count=@()
    $count += (KPI_offense_opened -start_date $start_date -end_date $end_date -offense_list_cache $offense_list_cache)
    $count = $count.count
    if($count -eq $null) {$count=0}


    $temp = New-Object System.Object
    $temp | Add-Member -MemberType NoteProperty -Name "Offense_status" -Value "open"
    $temp | Add-Member -MemberType NoteProperty -Name "count" -Value $count



    $collectionWithItems +=$temp

    $count=@()
    $count += (KPI_offense_opened -offense_list_cache $offense_list_cache)
    $count = $count.count
    if($count -eq $null) {$count=0}
    $backlog_total = $count

    $temp = New-Object System.Object
    $temp | Add-Member -MemberType NoteProperty -Name "Offense_status" -Value "Backlog open (total)"
    $temp | Add-Member -MemberType NoteProperty -Name "count" -Value $count

    $collectionWithItems +=$temp

   

    $count=@()
    $count += $offense_list_cache
    $count = $count.count
    if($count -eq $null) {$count=0}

    $temp = New-Object System.Object
    $temp | Add-Member -MemberType NoteProperty -Name "Offense_status" -Value "Total offense in Qradar"
    $temp | Add-Member -MemberType NoteProperty -Name "count" -Value $count

    $collectionWithItems +=$temp


    $count=@()
    $count += (KPI_offense_opened -offense_list_cache $offense_list_cache) | ?{($_.assigned_to).length -gt 1}
    $count = $count.count
    if($count -eq $null) {$count=0}
    $backlog_handled=$count

    $temp = New-Object System.Object
    $temp | Add-Member -MemberType NoteProperty -Name "Offense_status" -Value "backlog_assigned"
    $temp | Add-Member -MemberType NoteProperty -Name "count" -Value $count

    $collectionWithItems +=$temp


    $count=@()
    $count += (KPI_offense_opened -start_date $start_date -end_date $end_date -offense_list_cache $offense_list_cache) | ?{(($_.assigned_to).length -gt 1)}
    $count = $count.count
    if($count -eq $null) {$count=0}
    $new_backlog_assigned=$count

    $temp = New-Object System.Object
    $temp | Add-Member -MemberType NoteProperty -Name "Offense_status" -Value "new_backlog_assigned"
    $temp | Add-Member -MemberType NoteProperty -Name "count" -Value $count

    $collectionWithItems +=$temp

    
    $backlog_not_assigned= $backlog_total - $backlog_handled

    $temp = New-Object System.Object
    $temp | Add-Member -MemberType NoteProperty -Name "Offense_status" -Value "backlog_not_assigned"
    $temp | Add-Member -MemberType NoteProperty -Name "count" -Value $backlog_not_assigned

    $collectionWithItems +=$temp



    #$count = $backlog_handled + $closed_handled
    $count = $new_backlog_assigned + $closed_handled

    $temp = New-Object System.Object
    $temp | Add-Member -MemberType NoteProperty -Name "Offense_status" -Value "total_handled"
    $temp | Add-Member -MemberType NoteProperty -Name "count" -Value $count



    $collectionWithItems +=$temp



    return $collectionWithItems

}

Function KPI_closingReason{
    Param(
        $start_date = $start_epoch_time,
        $end_date = (get-date),
        [parameter(mandatory)]$closingReasonList ="",
        [parameter(mandatory)]$offense_list_cache=""
    )
    
    $start_date = (Get-epochTime_milliseconds -end_date $start_date)
    $end_date = (Get-epochTime_milliseconds -end_date $end_date)

    $offense_list_cache=$offense_list_cache |?{($_.status -eq "CLOSED") -and ($_.close_time -ge $start_date) -and ($_.close_time -le $end_date)}
    $collectionWithItems = @()
    foreach ($closing_reason in $closingReasonList){
    
    $count=@()
    $count += ($offense_list_cache |?{$_.closing_reason_id -eq $closing_reason.id})
    $count = $count.count

    if($count -eq $null) {$count=0}
    
    $temp = New-Object System.Object
    $temp | Add-Member -MemberType NoteProperty -Name "id" -Value $closing_reason.id
    $temp | Add-Member -MemberType NoteProperty -Name "type" -Value $closing_reason.text
    $temp | Add-Member -MemberType NoteProperty -Name "count" -Value $count

    $collectionWithItems +=$temp
    }

    return $collectionWithItems
}

Function KPI_closingReason_by_rule{
    Param(
        $start_date = $start_epoch_time,
        $end_date = (get-date),
        [parameter(mandatory)]$closingReasonList ="",
        [parameter(mandatory)]$offense_list_cache="",
        [parameter(mandatory)]$rule_list_cache=""
    )
    
    $start_date = (Get-epochTime_milliseconds -end_date $start_date)
    $end_date = (Get-epochTime_milliseconds -end_date $end_date)
    
    $collectionWithItems = @()
    foreach ($closing_reason in $closingReasonList){
        $sub_offense_list = ($offense_list_cache |?{$_.closing_reason_id -eq $closing_reason.id})
        $rules_list = $sub_offense_list.rules

        foreach ($rule in $rule_list_cache){
            $rule_name = $rule.name
            $count=@()
            $count += ($rules_list|?{$_.id -eq $rule.id})
            $count = $count.count

            $temp = New-Object System.Object
            $temp | Add-Member -MemberType NoteProperty -Name "ruleid" -Value $rule.id
            $temp | Add-Member -MemberType NoteProperty -Name "rule_name" -Value $rule_name
            $temp | Add-Member -MemberType NoteProperty -Name "rule_origin" -Value $rule.origin
            $temp | Add-Member -MemberType NoteProperty -Name "closingReason_id" -Value $closing_reason.id
            $temp | Add-Member -MemberType NoteProperty -Name "closingReason_name" -Value $closing_reason.text
            $temp | Add-Member -MemberType NoteProperty -Name "count" -Value $count

            $collectionWithItems +=$temp
        }
    }
    return $collectionWithItems
}

Function KPI_closingReason_by_rule_transported{
    Param(
        [parameter(mandatory)]$closingReasonList ="",
        [parameter(mandatory)]$KPI_closingReason_by_rule=""
    )
    
    $KPI_closingReason_by_rule = $KPI_closingReason_by_rule | select *,@{Name='counter';Expression={$_.count}} 

    $dataset_temp = $KPI_closingReason_by_rule | sort -Unique ruleid

    $collectionWithItems = @()

    foreach($unique_rule in $dataset_temp ){
        $rule= $KPI_closingReason_by_rule | ?{$_.ruleid -eq $unique_rule.ruleid}

        $temp = New-Object System.Object
        $temp | Add-Member -MemberType NoteProperty -Name "ruleid" -Value $unique_rule.ruleid
        $temp | Add-Member -MemberType NoteProperty -Name "rule_name" -Value $unique_rule.rule_name
        $temp | Add-Member -MemberType NoteProperty -Name "rule_origin" -Value $unique_rule.rule_origin

        foreach ($closing_reason in $closingReasonList){
            $closing_reason_name = $closing_reason.text
            
            $value = $rule | ?{$_.closingReason_id -eq $closing_reason.id}
            $count = $value.counter
            $temp | Add-Member -MemberType NoteProperty -Name $closing_reason_name -Value $count
        }
        $collectionWithItems +=$temp
    }
    return $collectionWithItems
}


Function KPI_generation{
    Param(
        $start_date = $start_epoch_time,
        $end_date = (get-date),
        [switch]$output_csv=$false,
        $output="kpi_qradar.xlsx"
    )

    # Variable initialization
    $end_date = get-date $end_date
    $end_date_formated = Get-Date $end_date -Format "dd/MM/yyyy"
    $rule_list_cache = Get-rulelist
    $closing_reason_list_cache= Get-closingReasonList
    $offense_list_cache_raw = Get-OffenseList
    $rule_list_BB_list_cache = Get-rulelist_and_BBlist


    $offense_list_cache_raw_formated_epochtime = Format_Offense_closingReason -closingReasonList $closing_reason_list_cache -offense_list_cache $offense_list_cache_raw  -rule_list_cache $rule_list_cache
    $offense_list_cache_raw_formated_epochtime = Format_Offense_rules -offense_list_cache $offense_list_cache_raw_formated_epochtime -rule_list_cache $rule_list_cache
    $offense_list_cache = $offense_list_cache_raw_formated_epochtime | ?{$_.offense_source -ne "Offense Monitoring Event"}
    <#
    $params =@(
        start_date = $start_date
        end_date = $end_date
        offense_list_cache = $offense_list_cache
    )
    #>

    $KPI_rule_and_BB_modified=KPI_rule_and_BB_modified -start_date $start_date -end_date $end_date
    $KPI_rule_and_BB_modified_status = KPI_rule_and_BB_modified_status -rule_list $KPI_rule_and_BB_modified
    $KPI_rules_status = KPI_rules_status -rule_list $rule_list_cache
    $KPI_closingReason_by_rule=(KPI_closingReason_by_rule -start_date $start_date -end_date $end_date -closingReasonList $closing_reason_list_cache -offense_list_cache (KPI_offense_closed -start_date $start_date -end_date $end_date -offense_list_cache $offense_list_cache) -rule_list_cache $rule_list_cache) |?{$_.count -ne 0}| sort -Descending count
    $KPI_offense_by_rule=(KPI_offense_by_rule -offense_list_cache $offense_list_cache -rule_list_cache $rule_list_cache -start_date $start_date -end_date $end_date)  <#|?{$_.count -ne 0}#>| sort  -Descending count
    $KPI_closingReason=(KPI_closingReason -start_date $start_date -end_date $end_date -closingReasonList $closing_reason_list_cache -offense_list_cache $offense_list_cache) | sort -Descending count
    $KPI_offense_status=(KPI_offense_status -start_date $start_date -end_date $end_date -offense_list_cache $offense_list_cache)
    $KPI_offense_full_dataset=(Format_Offense -offense_list_cache $offense_list_cache)
    $KPI_offense_partial_dataset= (Format_Offense -offense_list_cache (KPI_select_partial_offense_dataset -offense_list_cache $offense_list_cache -start_date $start_date -end_date $end_date))
    $KPI_rule_dataset = Format_Rule -rule_list_cache $rule_list_cache
    $KPI_rule_and_BB_dataset = Format_Rule -rule_list_cache $rule_list_BB_list_cache 
    $KPI_offense_open_by_rule = KPI_offense_open_by_rule -offense_list_cache (KPI_offense_opened -offense_list_cache $offense_list_cache) -rule_list_cache $rule_list_cache |?{$_.count -ne 0}|sort -Descending count
    $KPI_closingReason_by_rule_transported = (KPI_closingReason_by_rule_transported  -closingReasonList $closing_reason_list_cache -KPI_closingReason_by_rule $KPI_closingReason_by_rule)
    

    #Output

    if($output_csv){
        Csv_creation -value $KPI_rule_and_BB_modified -path KPI_rule_and_BB_modified.csv
        Csv_creation -value $KPI_rule_and_BB_modified_status -path KPI_rule_and_BB_modified_status.csv
        Csv_creation -value $KPI_closingReason_by_rule -path KPI_closingReason_by_rule.csv
        Csv_creation -value $KPI_offense_by_rule -path KPI_offense_by_rule.csv
        Csv_creation -value $KPI_closingReason -path KPI_closingReason.csv 
        Csv_creation -value $KPI_offense_status -path KPI_offense_status.csv
        Csv_creation -value $KPI_offense_dataset -path KPI_offense_dataset.csv
        Csv_creation -value $KPI_rule_dataset -path KPI_rule_dataset.csv
        Csv_creation -value $KPI_offense_open_by_rule -path KPI_offense_open_by_rule.csv
    }

    
    Remove-Item -Path $output -ErrorAction SilentlyContinue

    $KPI_rule_and_BB_modified | export-excel -path $output -WorksheetName "rule_and_BB_modified" -TableName "rule_and_BB_modified"
    Create_chart_excel -value $KPI_rule_and_BB_modified_status -output $output -worksheetName "KPI_rule_and_BB_modified_status" -title "Tuning SIEM" -Xrange "status" -Yrange "Rule" -seriesHeader "Rule" -chartType "columnstacked" -noLegend
    #$KPI_rule_and_BB_modified_status | Export-Excel -path $output -worksheetName "KPI_rule_and_BB_modified_status"

    $KPI_closingReason_by_rule | Export-Excel -path $output -worksheetName "closingReason_by_rule"
    
    $KPI_rules_status | Export-Excel -path $output -worksheetName "Rules_status" -TableName "Rules_status"
    
    $NonAlphaEscaped = ($closing_reason_list_cache.text -replace '\W','_')
    Create_chart_excel -value $KPI_closingReason_by_rule_transported -output $output -worksheetName "closingReason_by_rule_stacked" -title "Répartition des qualifications offense par règle" -Xrange "Rule_name" -Yrange $NonAlphaEscaped -seriesHeader  $closing_reason_list_cache.text <#$NonAlphaEscaped #>  -chartType "columnstacked"
    #PivotChartTable -output $output -data $KPI_closingReason_by_rule -WorkSheetName 'ClosingByRule' -TableName 'KPIClosingReasonByRule' -PivotRows 'rule_name' -PivotData 'count' -PivotColumns 'closingReason_name' -charType 'columnStacked' -PivotDataAction "Sum"

    
    
    Create_chart_excel -value $KPI_offense_by_rule -output $output -worksheetName "KPI_offense_by_rule" -title "Nombre d'offense par règle sur la période $start_date - $end_date_formated" -Xrange "rule_name" -Yrange "count" -chartType "pie" -ShowPercent
    $KPI_offense_status | export-excel -path $output -WorksheetName "offense_status" -tableName "offense_status"

    #no Charts
    #$KPI_closingReason_by_rule_transported | Export-Excel -Path $output -WorksheetName "test" -TableName "test"
    #$KPI_offense_open_by_rule | Export-Excel -Path $output -WorksheetName "KPI_offense_backlog_by_rule" -TableName "KPI_offense_backlog_by_rule"
    
    Create_chart_excel -value $KPI_offense_open_by_rule -output $output -worksheetName "KPI_offense_backlog_by_rule" -title "Répartition des offenses ouvertes par règle" -Xrange "ruleid" -Yrange "count" -chartType "BarStacked" -noLegend
    Create_chart_excel -value $KPI_closingReason -output $output -worksheetName "KPI_closingReason" -title "Répartition des qualifications par le SOC" -Xrange "type" -Yrange "count" -chartType "pie" -ShowPercent


    $KPI_offense_partial_dataset | Export-Excel -path $output -WorksheetName "partial_offense_dataset" -TableName "partial_offense_dataset"
    $KPI_offense_full_dataset| Export-Excel -path $output -WorksheetName "offense_full_dataset" -tableName "offense_full_dataset"
    $KPI_rule_and_BB_dataset| Export-Excel -path $output -WorksheetName "rule_BB_dataset" -tableName "rule_BB_dataset"
        
 }   

if($test){
    #Get-SearchResults -SearchID 'a283ca83-83f0-4ecc-bfab-e0af945a74ab'
    #Get-Search -SearchID 'bd1bab49-b8f5-49dd-8900-75828e87c106'
    Get-SearchListName
}
else {
    if ($ReferenceSet_json -ne '')
    {
        $RS_set = jsonMISP_to_Qradar_ReferenceSet -json (get-content $ReferenceSet_json | ConvertFrom-Json)
        Search-ReferenceSet -RS_set $RS_set
        
    }
    elseif ($ReferenceSet_csv -ne '')
    {
        $RS_set = csv_to_Qradar_ReferenceSet -csv (get-content $ReferenceSet_csv | ConvertFrom-Csv -Delimiter $delimiter)
        Search-ReferenceSet -RS_set $RS_set
    }
    elseif($null -ne $listReferenceSet)
    {
        if($listReferenceSet.length -gt 1){
            (Get-ReferenceSet -RSName $listReferenceSet) | select @{Name="value";Expression={$_.data.value}},@{Name="last_seen";Expression={convert_epochtime_milliseconds -epoch_time_to_convert $_.data.last_seen}},@{Name="first_seen";Expression={convert_epochtime_milliseconds -epoch_time_to_convert $_.data.first_seen}},time_to_live,timeout_type | sort last_seen
        }
        else{
        (Get-ReferenceSet) | select name,element_type,number_of_elements,@{Name="Creation_time";Expression={convert_epochtime_milliseconds -epoch_time_to_convert $_.creation_time}},time_to_live,timeout_type | sort name
        }
    }
    elseif($listSearch){
        Get-SearchListName
    }
    elseif($searchID)
    {
        if($cancelSearch){
            Update-Search -SearchID $searchID -status "CANCELED"
            Write-Host "The search $searchID has been canceled"
        }
        elseif($deleteSearch){
            Delete-Search -SearchID $searchID
        }
        elseif($checkSearchStatus){
            $result = Get-Search -SearchID $searchID
            
            Write-Host ("AQL query: "+$result.query_string)
            Write-Host ("Search ID: "+$result.search_id)
            Write-Host ("Search status: "+$result.status)
            Write-Host ("Progress: "+$result.progress +"%")
            Write-Host ("Save results: "+$result.save_result.tostring)
            Write-Host ("Record count: "+$result.record_count)
            Write-Host ("Query execution time: "+$result.query_execution_time +"ms")
            $result
        }
        else{
            if($saveSearch){
                Update-Search -SearchID $searchID -saveResults $saveSearch
                Write-Host "The search $searchID has been saved"
            }
            else {
                $result = Get-SearchResults -SearchID $searchID
                
                Write-Host ("AQL query: "+$result.query_string)
                Write-Host ("Search ID: "+$result.search_id)
                Write-Host ("Search status: "+$result.status)
                Write-Host ("Progress: "+$result.progress +"%")
                Write-Host ("Save results: "+$result.save_result.tostring)
                Write-Host ("Record count: "+$result.record_count)
                Write-Host ("Query execution time: "+$result.query_execution_time +"ms")
                $result
            }
        }
    }
    elseif ($null -ne $runSearch) {
        $result = Create-Search -AQL $runSearch
        Write-Host "AQL query: $runSearch"
        Write-Host ("Search ID: "+$result.search_id)
        Write-Host ("Search status: "+$result.status)
        Write-Host ("Progress: "+$result.progress +"%")
        Write-Host ("Save results: "+$result.save_result)
        Write-Host ("Record count: "+$result.record_count)
        Write-Host ("Query execution time: "+$result.query_execution_time +"ms")

        $result 
    }
    else{
        KPI_generation -start_date $start_date -end_date $end_date
    }
}