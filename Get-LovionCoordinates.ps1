[CmdletBinding(DefaultParameterSetName = 'Excel')]
param(
    [Parameter(Mandatory = $true, ParameterSetName = 'Excel')]
    [string]$InputExcelPath,

    [Parameter(Mandatory = $true, ParameterSetName = 'Text')]
    [string]$InputTextPath,

    [Parameter(ParameterSetName = 'Excel')]
    [string]$WorksheetName = 'Alle stations nieuw',

    [string]$OutputDirectory,

    [string]$WfsUrl = 'https://api.enexis.nl/opendata-assets/v1/wfs',

    [string]$WfsLayerName = 'Opendata:asm_e_lv_service_connection',

    [string]$WfsGeoJsonPath,

    [int]$GeocodeDelayMs = 0,

    [int]$GeocodeConcurrency = 6,

    [int]$SnapDistanceMeters = 25,

    [int]$StreetFurnitureSnapDistanceMeters = 100,

    [int]$WfsBufferMeters = 100,

    [int]$MaxRows = 0,

    [switch]$SkipGeocoding,

    [switch]$SkipWfs
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Net.Http

$scriptRoot = if ([string]::IsNullOrWhiteSpace($PSScriptRoot)) { (Get-Location).Path } else { $PSScriptRoot }
if ([string]::IsNullOrWhiteSpace($OutputDirectory)) {
    $OutputDirectory = Join-Path -Path $scriptRoot -ChildPath 'output'
}

$script:StreetFurnitureUsagePatterns = @(
    '(^|\W)(riool|fontein|putkast)',
    '(^|\W)(ov[\- ]?kast|vri[\- ]?kast|parkeer|slagboom|verkeer|brug|sluis)',
    '(^|\W)(abri|bushalte|reclame|cai|telecom|telefoon|sirene|camera|flits|bord)',
    '(^|\W)(verlichting|mast|lantaarn|marktkast|kermis|feest|afval|vuil|container|toilet|urinoir)',
    '(^|\W)(bouwaansluiting)'
)

function Write-Info {
    param([string]$Message)
    Write-Host "[INFO] $Message"
}

function Write-WarnLine {
    param([string]$Message)
    Write-Warning $Message
}

function Test-Blank {
    param($Value)

    if ($null -eq $Value) {
        return $true
    }

    return [string]::IsNullOrWhiteSpace([string]$Value)
}

function Normalize-Text {
    param($Value)

    if ($null -eq $Value) {
        return ''
    }

    $text = [string]$Value
    $text = $text.Trim().ToUpperInvariant()
    $text = $text -replace '\s+', ' '
    return $text
}

function Normalize-Postcode {
    param($Value)

    if ($null -eq $Value) {
        return ''
    }

    return ([string]$Value).ToUpperInvariant() -replace '\s+', ''
}

function Normalize-HouseNumber {
    param($Value)

    if ($null -eq $Value) {
        return ''
    }

    return ([string]$Value).ToUpperInvariant() -replace '\s+', ''
}

function Get-HouseNumberDigits {
    param($Value)

    $normalized = Normalize-HouseNumber $Value
    if ([string]::IsNullOrWhiteSpace($normalized)) {
        return ''
    }

    $match = [regex]::Match($normalized, '^\d+')
    if (-not $match.Success) {
        return $normalized
    }

    return $match.Value
}

function Build-StreetKeyFromValues {
    param(
        $Street,
        $Postcode,
        $City
    )

    $normalizedStreet = Normalize-Text $Street
    $normalizedPostcode = Normalize-Postcode $Postcode
    $normalizedCity = Normalize-Text $City

    if ([string]::IsNullOrWhiteSpace($normalizedStreet)) {
        return ''
    }

    return "$normalizedStreet|$normalizedPostcode|$normalizedCity"
}

function Test-IsStreetFurniture {
    param([pscustomobject]$Row)

    $usage = [string]$Row.Gebruiksdoel
    if ([string]::IsNullOrWhiteSpace($usage)) {
        return $false
    }

    foreach ($pattern in $script:StreetFurnitureUsagePatterns) {
        if ([regex]::IsMatch($usage, $pattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)) {
            return $true
        }
    }

    return $false
}

function Join-NonEmpty {
    param([string[]]$Parts)

    $filtered = @()
    foreach ($part in $Parts) {
        if (-not [string]::IsNullOrWhiteSpace($part)) {
            $filtered += $part.Trim()
        }
    }

    return ($filtered -join ' ').Trim()
}

function Escape-SolrPhrase {
    param([string]$Value)

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return '""'
    }

    $escaped = $Value -replace '\\', '\\\\'
    $escaped = $escaped -replace '"', '\"'
    return ('"{0}"' -f $escaped)
}

function Build-AddressQuery {
    param([pscustomobject]$Row)

    $street = [string]$Row.Straatnaam
    $houseNumber = [string]$Row.Huisnummer
    $postcode = [string]$Row.Postcode
    $city = [string]$Row.Woonplaats

    $structured = Join-NonEmpty @($street, $houseNumber, $postcode, $city)
    if (-not [string]::IsNullOrWhiteSpace($structured)) {
        return $structured
    }

    return Join-NonEmpty @([string]$Row.Adres, [string]$Row.Gemeente, [string]$Row.Woonplaats)
}

function Build-PdokStructuredQuery {
    param([pscustomobject]$Row)

    $street = [string]$Row.Straatnaam
    $postcode = Normalize-Postcode $Row.Postcode
    $city = [string]$Row.Woonplaats
    $houseNumberRaw = [string]$Row.Huisnummer
    $houseNumberDigits = Get-HouseNumberDigits $houseNumberRaw
    $houseNumberFull = Normalize-HouseNumber $houseNumberRaw
    $suffix = ''
    if (-not [string]::IsNullOrWhiteSpace($houseNumberFull) -and -not [string]::IsNullOrWhiteSpace($houseNumberDigits) -and $houseNumberFull.Length -gt $houseNumberDigits.Length) {
        $suffix = $houseNumberFull.Substring($houseNumberDigits.Length)
    }

    $clauses = New-Object System.Collections.Generic.List[string]
    if (-not [string]::IsNullOrWhiteSpace($street)) {
        $clauses.Add(('straatnaam:{0}' -f (Escape-SolrPhrase -Value $street.Trim())))
    }
    if (-not [string]::IsNullOrWhiteSpace($houseNumberDigits)) {
        $clauses.Add(('huisnummer:{0}' -f $houseNumberDigits))
    }
    if (-not [string]::IsNullOrWhiteSpace($suffix)) {
        if ($suffix -match '^[A-Z]$') {
            $clauses.Add(('huisletter:{0}' -f $suffix))
        }
        elseif ($suffix -match '^\d{1,4}$') {
            $clauses.Add(('huisnummertoevoeging:{0}' -f (Escape-SolrPhrase -Value $suffix)))
        }
    }
    if (-not [string]::IsNullOrWhiteSpace($postcode)) {
        $clauses.Add(('postcode:{0}' -f $postcode))
    }
    if (-not [string]::IsNullOrWhiteSpace($city)) {
        $clauses.Add(('woonplaatsnaam:{0}' -f (Escape-SolrPhrase -Value $city.Trim())))
    }

    if ($clauses.Count -gt 0) {
        return ($clauses -join ' and ')
    }

    return Build-AddressQuery -Row $Row
}

function Build-PdokRequestUri {
    param([pscustomobject]$Row)

    $query = Build-PdokStructuredQuery -Row $Row
    if ([string]::IsNullOrWhiteSpace($query)) {
        return ''
    }

    $pairs = @(
        'q=' + [System.Uri]::EscapeDataString($query),
        'fq=' + [System.Uri]::EscapeDataString('type:adres'),
        'fq=' + [System.Uri]::EscapeDataString('bron:BAG'),
        'rows=3',
        'fl=' + [System.Uri]::EscapeDataString('id,weergavenaam,type,straatnaam,straatnaam_verkort,huisnummer,huisletter,huisnummertoevoeging,huis_nlt,postcode,woonplaatsnaam,centroide_ll,centroide_rd,score')
    )

    return ('https://api.pdok.nl/bzk/locatieserver/search/v3_1/free?{0}' -f ($pairs -join '&'))
}

function Build-AddressKey {
    param([pscustomobject]$Row)

    $street = Normalize-Text $Row.Straatnaam
    $houseNumber = Normalize-HouseNumber $Row.Huisnummer
    $postcode = Normalize-Postcode $Row.Postcode
    $city = Normalize-Text $Row.Woonplaats

    if ([string]::IsNullOrWhiteSpace($street) -or [string]::IsNullOrWhiteSpace($houseNumber) -or [string]::IsNullOrWhiteSpace($postcode)) {
        return ''
    }

    return "$street|$houseNumber|$postcode|$city"
}

function Release-ComObject {
    param($ComObject)

    if ($null -ne $ComObject) {
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($ComObject)
    }
}

function Get-WorksheetByName {
    param(
        $Workbook,
        [string]$Name
    )

    foreach ($worksheet in $Workbook.Worksheets) {
        if ($worksheet.Name -eq $Name) {
            return $worksheet
        }
    }

    throw "Werkblad '$Name' niet gevonden in '$($Workbook.FullName)'."
}

function Read-ExcelRows {
    param(
        [string]$Path,
        [string]$SheetName
    )

    $excel = $null
    $workbook = $null
    $worksheet = $null
    $usedRange = $null
    $values = $null

    try {
        Write-Info "Lees Excel: $Path [$SheetName]"
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false

        $workbook = $excel.Workbooks.Open($Path)
        $worksheet = Get-WorksheetByName -Workbook $workbook -Name $SheetName
        $usedRange = $worksheet.UsedRange

        $rowCount = [int]$usedRange.Rows.Count
        $columnCount = [int]$usedRange.Columns.Count
        $values = $usedRange.Value2

        if ($rowCount -lt 2) {
            throw "Werkblad '$SheetName' bevat geen datarijen."
        }

        $headers = @()
        for ($columnIndex = 1; $columnIndex -le $columnCount; $columnIndex++) {
            $headerValue = [string]$values[1, $columnIndex]
            if ([string]::IsNullOrWhiteSpace($headerValue)) {
                $headerValue = "Column$columnIndex"
            }

            $headers += $headerValue.Trim()
        }

        $rows = New-Object System.Collections.Generic.List[object]
        for ($rowIndex = 2; $rowIndex -le $rowCount; $rowIndex++) {
            $record = [ordered]@{
                _RowNumber = $rowIndex
            }

            $hasValue = $false
            for ($columnIndex = 1; $columnIndex -le $columnCount; $columnIndex++) {
                $header = $headers[$columnIndex - 1]
                $cellValue = $values[$rowIndex, $columnIndex]
                $cellText = if ($null -eq $cellValue) { '' } else { [string]$cellValue }
                if (-not [string]::IsNullOrWhiteSpace($cellText)) {
                    $hasValue = $true
                }

                $record[$header] = $cellText
            }

            if ($hasValue) {
                $rows.Add([pscustomobject]$record)
            }
        }

        return $rows
    }
    finally {
        if ($null -ne $usedRange) { Release-ComObject $usedRange }
        if ($null -ne $worksheet) { Release-ComObject $worksheet }
        if ($null -ne $workbook) {
            $workbook.Close($false)
            Release-ComObject $workbook
        }
        if ($null -ne $excel) {
            $excel.Quit()
            Release-ComObject $excel
        }

        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

function Parse-WktPoint {
    param([string]$Wkt)

    if ([string]::IsNullOrWhiteSpace($Wkt)) {
        return $null
    }

    $match = [regex]::Match($Wkt, 'POINT\s*\(\s*(?<x>-?\d+(\.\d+)?)\s+(?<y>-?\d+(\.\d+)?)\s*\)', 'IgnoreCase')
    if (-not $match.Success) {
        return $null
    }

    return [pscustomobject]@{
        X = [double]$match.Groups['x'].Value
        Y = [double]$match.Groups['y'].Value
    }
}

function Convert-RdToWgs84 {
    param(
        [double]$X,
        [double]$Y
    )

    $dX = ($X - 155000.0) * 0.00001
    $dY = ($Y - 463000.0) * 0.00001

    $latSeconds =
        (3235.65389 * $dY) +
        (-32.58297 * [math]::Pow($dX, 2)) +
        (-0.2475 * [math]::Pow($dY, 2)) +
        (-0.84978 * [math]::Pow($dX, 2) * $dY) +
        (-0.0655 * [math]::Pow($dY, 3)) +
        (-0.01709 * [math]::Pow($dX, 2) * [math]::Pow($dY, 2)) +
        (-0.00738 * $dX) +
        (0.0053 * [math]::Pow($dX, 4)) +
        (-0.00039 * [math]::Pow($dX, 2) * [math]::Pow($dY, 3)) +
        (0.00033 * [math]::Pow($dX, 4) * $dY) +
        (-0.00012 * $dX * $dY)

    $lonSeconds =
        (5260.52916 * $dX) +
        (105.94684 * $dX * $dY) +
        (2.45656 * $dX * [math]::Pow($dY, 2)) +
        (-0.81885 * [math]::Pow($dX, 3)) +
        (0.05594 * $dX * [math]::Pow($dY, 3)) +
        (-0.05607 * [math]::Pow($dX, 3) * $dY) +
        (0.01199 * $dY) +
        (-0.00256 * [math]::Pow($dX, 3) * [math]::Pow($dY, 2)) +
        (0.00128 * $dX * [math]::Pow($dY, 4)) +
        (0.00022 * [math]::Pow($dY, 2)) +
        (-0.00022 * [math]::Pow($dX, 2)) +
        (0.00026 * [math]::Pow($dX, 5))

    return [pscustomobject]@{
        Latitude = 52.15517440 + ($latSeconds / 3600.0)
        Longitude = 5.38720621 + ($lonSeconds / 3600.0)
    }
}

function Search-PdokAddress {
    param(
        [string]$Query,
        [string]$RequestUri
    )

    if ([string]::IsNullOrWhiteSpace($RequestUri) -and [string]::IsNullOrWhiteSpace($Query)) {
        return @()
    }

    $uri = if ([string]::IsNullOrWhiteSpace($RequestUri)) {
        Get-PdokSearchUri -Query $Query
    }
    else {
        $RequestUri
    }
    $response = Invoke-RestMethod -Uri $uri -Method Get -Headers @{ Accept = 'application/json' }
    $docs = @($response.response.docs)

    if ($null -eq $response.response -or $docs.Count -eq 0) {
        return @()
    }

    return $docs
}

function Get-PdokSearchUri {
    param(
        [string]$Query,
        [int]$Rows = 10
    )

    $encodedQuery = [System.Uri]::EscapeDataString($Query)
    return "https://api.pdok.nl/bzk/locatieserver/search/v3_1/free?q=$encodedQuery&rows=$Rows"
}

function Convert-PdokJsonToDocs {
    param([string]$Json)

    if ([string]::IsNullOrWhiteSpace($Json)) {
        return @()
    }

    try {
        $response = $Json | ConvertFrom-Json
    }
    catch {
        return @()
    }

    $docs = @($response.response.docs)
    if ($null -eq $response.response -or $docs.Count -eq 0) {
        return @()
    }

    return $docs
}

function Invoke-PdokQueryBatch {
    param(
        [string[]]$Queries,
        [int]$MaxConcurrency = 6,
        [int]$DelayBetweenDispatchMs = 0
    )

    $queryList = @($Queries | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    $cache = @{}
    if ($queryList.Count -eq 0) {
        return $cache
    }

    if ($MaxConcurrency -le 1) {
        $completed = 0
        foreach ($query in $queryList) {
            try {
                if ($query -like 'http*://*') {
                    $cache[$query] = @(Search-PdokAddress -RequestUri $query)
                }
                else {
                    $cache[$query] = @(Search-PdokAddress -Query $query)
                }
            }
            catch {
                $cache[$query] = @()
            }

            $completed += 1
            if ($DelayBetweenDispatchMs -gt 0 -and $completed -lt $queryList.Count) {
                Start-Sleep -Milliseconds $DelayBetweenDispatchMs
            }
        }

        return $cache
    }

    [System.Net.ServicePointManager]::DefaultConnectionLimit = [Math]::Max([System.Net.ServicePointManager]::DefaultConnectionLimit, ($MaxConcurrency * 2))

    $client = [System.Net.Http.HttpClient]::new()
    $client.Timeout = [TimeSpan]::FromSeconds(30)
    $client.DefaultRequestHeaders.Accept.Clear()
    [void]$client.DefaultRequestHeaders.Accept.Add([System.Net.Http.Headers.MediaTypeWithQualityHeaderValue]::new('application/json'))

    $pending = New-Object System.Collections.ArrayList
    $nextIndex = 0

    try {
        while ($nextIndex -lt $queryList.Count -or $pending.Count -gt 0) {
            while ($nextIndex -lt $queryList.Count -and $pending.Count -lt $MaxConcurrency) {
                $query = [string]$queryList[$nextIndex]
                $uri = if ($query -like 'http*://*') { $query } else { Get-PdokSearchUri -Query $query }
                $task = $client.GetStringAsync($uri)
                [void]$pending.Add([pscustomobject]@{
                    Query = $query
                    Task = $task
                })

                $nextIndex += 1
                if ($DelayBetweenDispatchMs -gt 0 -and $pending.Count -lt $MaxConcurrency -and $nextIndex -lt $queryList.Count) {
                    Start-Sleep -Milliseconds $DelayBetweenDispatchMs
                }
            }

            if ($pending.Count -eq 0) {
                continue
            }

            $tasks = [System.Threading.Tasks.Task[]]@($pending | ForEach-Object { [System.Threading.Tasks.Task]$_.Task })
            $completedIndex = [System.Threading.Tasks.Task]::WaitAny($tasks)
            if ($completedIndex -lt 0) {
                continue
            }

            $item = $pending[$completedIndex]
            $pending.RemoveAt($completedIndex)

            $query = [string]$item.Query
            try {
                $json = $item.Task.GetAwaiter().GetResult()
                $cache[$query] = @(Convert-PdokJsonToDocs -Json $json)
            }
            catch {
                $cache[$query] = @()
            }
        }
    }
    finally {
        $client.Dispose()
    }

    return $cache
}

function Convert-PdokDocToCandidate {
    param($Doc)

    $rdPoint = Parse-WktPoint (Get-PropertyValue -Object $Doc -CandidateNames @('centroide_rd'))
    $llPoint = Parse-WktPoint (Get-PropertyValue -Object $Doc -CandidateNames @('centroide_ll'))
    $fullHouseNumber = Normalize-HouseNumber (Get-PropertyValue -Object $Doc -CandidateNames @('huis_nlt'))
    if ([string]::IsNullOrWhiteSpace($fullHouseNumber)) {
        $fullHouseNumber = Normalize-HouseNumber ("{0}{1}" -f [string](Get-PropertyValue -Object $Doc -CandidateNames @('huisnummer')), [string](Get-PropertyValue -Object $Doc -CandidateNames @('huisletter')))
    }

    return [pscustomobject]@{
        Id = [string](Get-PropertyValue -Object $Doc -CandidateNames @('id'))
        DisplayName = [string](Get-PropertyValue -Object $Doc -CandidateNames @('weergavenaam'))
        Type = [string](Get-PropertyValue -Object $Doc -CandidateNames @('type'))
        Score = [double](Get-PropertyValue -Object $Doc -CandidateNames @('score'))
        Street = Normalize-Text (Get-PropertyValue -Object $Doc -CandidateNames @('straatnaam'))
        StreetShort = Normalize-Text (Get-PropertyValue -Object $Doc -CandidateNames @('straatnaam_verkort'))
        Postcode = Normalize-Postcode (Get-PropertyValue -Object $Doc -CandidateNames @('postcode'))
        City = Normalize-Text (Get-PropertyValue -Object $Doc -CandidateNames @('woonplaatsnaam'))
        HouseNumber = Get-HouseNumberDigits (Get-PropertyValue -Object $Doc -CandidateNames @('huisnummer'))
        HouseNumberFull = $fullHouseNumber
        RdX = if ($rdPoint) { [double]$rdPoint.X } else { $null }
        RdY = if ($rdPoint) { [double]$rdPoint.Y } else { $null }
        Lon = if ($llPoint) { [double]$llPoint.X } else { $null }
        Lat = if ($llPoint) { [double]$llPoint.Y } else { $null }
    }
}

function Select-PdokCandidateForRow {
    param(
        [pscustomobject]$Row,
        [object[]]$Docs
    )

    if ($Docs.Count -eq 0) {
        return $null
    }

    $expectedStreet = Normalize-Text $Row.Straatnaam
    $expectedPostcode = Normalize-Postcode $Row.Postcode
    $expectedCity = Normalize-Text $Row.Woonplaats
    $expectedHouseNumber = Normalize-HouseNumber $Row.Huisnummer
    $expectedHouseDigits = Get-HouseNumberDigits $Row.Huisnummer
    $isStreetFurniture = Test-IsStreetFurniture -Row $Row

    $candidates = foreach ($doc in $Docs) {
        $candidate = Convert-PdokDocToCandidate -Doc $doc
        $streetMatches = $false
        if (-not [string]::IsNullOrWhiteSpace($expectedStreet)) {
            $streetMatches =
                ($candidate.Street -eq $expectedStreet) -or
                ($candidate.StreetShort -eq $expectedStreet)
        }

        $postcodeMatches = [string]::IsNullOrWhiteSpace($expectedPostcode) -or ($candidate.Postcode -eq $expectedPostcode)
        $cityMatches = [string]::IsNullOrWhiteSpace($expectedCity) -or ($candidate.City -eq $expectedCity)
        $houseDigitsMatch = [string]::IsNullOrWhiteSpace($expectedHouseDigits) -or ($candidate.HouseNumber -eq $expectedHouseDigits)
        $houseFullMatch = [string]::IsNullOrWhiteSpace($expectedHouseNumber) -or ($candidate.HouseNumberFull -eq $expectedHouseNumber)

        $strictMatch = $streetMatches -and $postcodeMatches -and $cityMatches -and $houseFullMatch

        $reason = if ($strictMatch) {
            'exact'
        }
        elseif (-not $streetMatches) {
            'street_mismatch'
        }
        elseif (-not $postcodeMatches) {
            'postcode_mismatch'
        }
        elseif (-not $houseDigitsMatch) {
            'house_number_mismatch'
        }
        elseif (-not $houseFullMatch) {
            'house_suffix_mismatch'
        }
        elseif (-not $cityMatches) {
            'city_mismatch'
        }
        else {
            'fuzzy'
        }

        $ranking =
            ($(if ($strictMatch) { 1000 } else { 0 })) +
            ($(if ($streetMatches) { 200 } else { 0 })) +
            ($(if ($postcodeMatches) { 100 } else { 0 })) +
            ($(if ($cityMatches) { 50 } else { 0 })) +
            ($(if ($houseFullMatch) { 40 } elseif ($houseDigitsMatch) { 20 } else { 0 })) +
            [math]::Round($candidate.Score, 0)

        [pscustomobject]@{
            Candidate = $candidate
            StrictMatch = $strictMatch
            Reason = $reason
            Ranking = $ranking
            StreetMatches = $streetMatches
            PostcodeMatches = $postcodeMatches
            CityMatches = $cityMatches
            HouseDigitsMatch = $houseDigitsMatch
            HouseFullMatch = $houseFullMatch
            IsStreetFurniture = $isStreetFurniture
        }
    }

    $strictCandidates = @($candidates | Where-Object { $_.StrictMatch })
    if ($strictCandidates.Count -gt 0) {
        return $strictCandidates | Sort-Object Ranking -Descending | Select-Object -First 1
    }

    return $candidates | Sort-Object Ranking -Descending | Select-Object -First 1
}

function Test-IsPdokToleratedReason {
    param([string]$Reason)

    return (([string]$Reason).Trim() -eq 'house_suffix_mismatch')
}

function New-BoundingBox {
    param(
        [object[]]$Rows,
        [int]$BufferMeters
    )

    $points = @(
        $Rows |
            Where-Object { $null -ne $_.PdokRdX -and $null -ne $_.PdokRdY } |
            ForEach-Object {
                [pscustomobject]@{
                    X = [double]$_.PdokRdX
                    Y = [double]$_.PdokRdY
                }
            }
    )

    if ($points.Count -eq 0) {
        return $null
    }

    $minX = ($points | Measure-Object -Property X -Minimum).Minimum - $BufferMeters
    $maxX = ($points | Measure-Object -Property X -Maximum).Maximum + $BufferMeters
    $minY = ($points | Measure-Object -Property Y -Minimum).Minimum - $BufferMeters
    $maxY = ($points | Measure-Object -Property Y -Maximum).Maximum + $BufferMeters

    return [pscustomobject]@{
        MinX = [double]$minX
        MinY = [double]$minY
        MaxX = [double]$maxX
        MaxY = [double]$maxY
    }
}

function Convert-HashtableToQueryString {
    param([hashtable]$Parameters)

    $pairs = foreach ($entry in $Parameters.GetEnumerator()) {
        '{0}={1}' -f [System.Uri]::EscapeDataString([string]$entry.Key), [System.Uri]::EscapeDataString([string]$entry.Value)
    }

    return ($pairs -join '&')
}

function Get-WfsFeaturesFromService {
    param(
        [string]$BaseUrl,
        [string]$LayerName,
        $BoundingBox
    )

    if ($null -eq $BoundingBox) {
        return @()
    }

    $parameters = @{
        service = 'WFS'
        version = '2.0.0'
        request = 'GetFeature'
        typeNames = $LayerName
        outputFormat = 'application/json'
        srsName = 'EPSG:28992'
        bbox = '{0},{1},{2},{3},EPSG:28992' -f $BoundingBox.MinX, $BoundingBox.MinY, $BoundingBox.MaxX, $BoundingBox.MaxY
        count = 5000
    }

    $uri = '{0}?{1}' -f $BaseUrl.TrimEnd('?'), (Convert-HashtableToQueryString -Parameters $parameters)
    Write-Info "Laad WFS-features uit bbox: $($BoundingBox.MinX),$($BoundingBox.MinY),$($BoundingBox.MaxX),$($BoundingBox.MaxY)"
    $response = Invoke-RestMethod -Uri $uri -Method Get -Headers @{ Accept = 'application/json' }

    if ($null -eq $response.features) {
        return @()
    }

    return @($response.features)
}

function Get-WfsFeaturesFromGeoJson {
    param([string]$Path)

    Write-Info "Laad lokale GeoJSON: $Path"
    $content = Get-Content -Path $Path -Raw -Encoding UTF8
    $json = $content | ConvertFrom-Json

    if ($null -eq $json.features) {
        return @()
    }

    return @($json.features)
}

function Is-NumericCoordinatePair {
    param($Node)

    if ($Node -isnot [System.Array]) {
        return $false
    }

    if ($Node.Length -lt 2) {
        return $false
    }

    $x = $Node[0]
    $y = $Node[1]

    return ($x -is [ValueType]) -and ($y -is [ValueType])
}

function Add-CoordinatePairs {
    param(
        $Node,
        [System.Collections.Generic.List[object]]$Pairs
    )

    if ($null -eq $Node) {
        return
    }

    if (Is-NumericCoordinatePair $Node) {
        $Pairs.Add([pscustomobject]@{
            X = [double]$Node[0]
            Y = [double]$Node[1]
        })
        return
    }

    if ($Node -is [System.Array]) {
        foreach ($child in $Node) {
            Add-CoordinatePairs -Node $child -Pairs $Pairs
        }
    }
}

function Get-GeometryCentroid {
    param($Geometry)

    if ($null -eq $Geometry -or $null -eq $Geometry.coordinates) {
        return $null
    }

    $pairs = New-Object System.Collections.Generic.List[object]
    Add-CoordinatePairs -Node $Geometry.coordinates -Pairs $pairs

    if ($pairs.Count -eq 0) {
        return $null
    }

    $xAverage = ($pairs | Measure-Object -Property X -Average).Average
    $yAverage = ($pairs | Measure-Object -Property Y -Average).Average

    return [pscustomobject]@{
        X = [double]$xAverage
        Y = [double]$yAverage
    }
}

function Get-PropertyValue {
    param(
        $Object,
        [string[]]$CandidateNames
    )

    if ($null -eq $Object) {
        return $null
    }

    foreach ($candidate in $CandidateNames) {
        $exact = $Object.PSObject.Properties | Where-Object { $_.Name -ieq $candidate } | Select-Object -First 1
        if ($exact) {
            return $exact.Value
        }
    }

    foreach ($candidate in $CandidateNames) {
        $partial = $Object.PSObject.Properties | Where-Object { $_.Name -imatch [regex]::Escape($candidate) } | Select-Object -First 1
        if ($partial) {
            return $partial.Value
        }
    }

    return $null
}

function Build-WfsAddressKey {
    param($Properties)

    $street = Normalize-Text (Get-PropertyValue -Object $Properties -CandidateNames @(
        'straatnaam', 'street', 'openbareruimtenaam', 'openbare_ruimte_naam', 'road'
    ))

    $houseNumber = Normalize-HouseNumber (Get-PropertyValue -Object $Properties -CandidateNames @(
        'huisnummer', 'house_number', 'nummer', 'number', 'huis_nummer'
    ))

    $postcode = Normalize-Postcode (Get-PropertyValue -Object $Properties -CandidateNames @(
        'postcode', 'postalcode', 'zip', 'zip_code', 'pc6'
    ))

    $city = Normalize-Text (Get-PropertyValue -Object $Properties -CandidateNames @(
        'woonplaats', 'woonplaatsnaam', 'city', 'town', 'place'
    ))

    if ([string]::IsNullOrWhiteSpace($street) -or [string]::IsNullOrWhiteSpace($houseNumber) -or [string]::IsNullOrWhiteSpace($postcode)) {
        return ''
    }

    return "$street|$houseNumber|$postcode|$city"
}

function Build-WfsStreetKey {
    param($Properties)

    $street = Get-PropertyValue -Object $Properties -CandidateNames @(
        'straatnaam', 'street', 'openbareruimtenaam', 'openbare_ruimte_naam', 'road'
    )

    $postcode = Get-PropertyValue -Object $Properties -CandidateNames @(
        'postcode', 'postalcode', 'zip', 'zip_code', 'pc6'
    )

    $city = Get-PropertyValue -Object $Properties -CandidateNames @(
        'woonplaats', 'woonplaatsnaam', 'city', 'town', 'place'
    )

    return Build-StreetKeyFromValues -Street $street -Postcode $postcode -City $city
}

function Measure-RdDistance {
    param(
        [double]$X1,
        [double]$Y1,
        [double]$X2,
        [double]$Y2
    )

    return [math]::Sqrt(([math]::Pow($X2 - $X1, 2)) + ([math]::Pow($Y2 - $Y1, 2)))
}

function Get-NearestWfsFeature {
    param(
        [object[]]$Features,
        [double]$BaseX,
        [double]$BaseY,
        [double]$MaxDistanceMeters = [double]::PositiveInfinity
    )

    if ($Features.Count -eq 0) {
        return $null
    }

    $ranked = $Features |
        ForEach-Object {
            [pscustomobject]@{
                Feature = $_
                Distance = Measure-RdDistance -X1 $BaseX -Y1 $BaseY -X2 $_.RdX -Y2 $_.RdY
            }
        } |
        Sort-Object Distance

    $nearest = $ranked | Select-Object -First 1
    if ($null -eq $nearest) {
        return $null
    }

    if ($nearest.Distance -gt $MaxDistanceMeters) {
        return $null
    }

    return $nearest
}

function Prepare-WfsFeatures {
    param([object[]]$RawFeatures)

    $prepared = New-Object System.Collections.Generic.List[object]
    foreach ($feature in $RawFeatures) {
        $centroid = Get-GeometryCentroid -Geometry $feature.geometry
        if ($null -eq $centroid) {
            continue
        }

        $properties = $feature.properties
        $prepared.Add([pscustomobject]@{
            FeatureId = [string]$feature.id
            GeometryType = [string]$feature.geometry.type
            RdX = [double]$centroid.X
            RdY = [double]$centroid.Y
            AddressKey = Build-WfsAddressKey -Properties $properties
            StreetKey = Build-WfsStreetKey -Properties $properties
            Properties = $properties
        })
    }

    return $prepared
}

function Export-GeoJson {
    param(
        [object[]]$Rows,
        [string]$Path
    )

    $features = New-Object System.Collections.Generic.List[object]
    foreach ($row in $Rows) {
        if ($null -eq $row.FinalRdX -or $null -eq $row.FinalRdY) {
            continue
        }

        $properties = [ordered]@{}
        foreach ($property in $row.PSObject.Properties) {
            if ($property.Name -like '_*') {
                continue
            }

            if ($property.Name -in @('FinalRdX', 'FinalRdY', 'FinalLon', 'FinalLat')) {
                continue
            }

            $properties[$property.Name] = $property.Value
        }

        $features.Add([ordered]@{
            type = 'Feature'
            geometry = [ordered]@{
                type = 'Point'
                coordinates = @([double]$row.FinalRdX, [double]$row.FinalRdY)
            }
            properties = $properties
        })
    }

    $geoJson = [ordered]@{
        type = 'FeatureCollection'
        name = 'lovion_coordinates'
        crs = [ordered]@{
            type = 'name'
            properties = [ordered]@{
                name = 'EPSG:28992'
            }
        }
        features = $features
    }

    $geoJson | ConvertTo-Json -Depth 8 | Set-Content -Path $Path -Encoding UTF8
}

function Export-CsvFile {
    param(
        [object[]]$Rows,
        [string]$Path
    )

    $selectProperties = @()
    foreach ($property in $Rows[0].PSObject.Properties.Name) {
        if ($property -notlike '_*') {
            $selectProperties += $property
        }
    }

    $Rows | Select-Object -Property $selectProperties | Export-Csv -Path $Path -NoTypeInformation -Encoding UTF8
}

function Export-AugmentedWorkbook {
    param(
        [string]$SourcePath,
        [string]$TargetPath,
        [string]$SheetName,
        [object[]]$Rows,
        [string[]]$ColumnsToWrite
    )

    Copy-Item -Path $SourcePath -Destination $TargetPath -Force

    $excel = $null
    $workbook = $null
    $worksheet = $null
    $usedRange = $null
    $headerRange = $null
    $dataRange = $null

    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $excel.ScreenUpdating = $false

        $workbook = $excel.Workbooks.Open($TargetPath)
        $worksheet = Get-WorksheetByName -Workbook $workbook -Name $SheetName
        $usedRange = $worksheet.UsedRange

        $startColumn = [int]$usedRange.Columns.Count + 1
        $headerRange = $worksheet.Range(
            $worksheet.Cells.Item(1, $startColumn),
            $worksheet.Cells.Item(1, $startColumn + $ColumnsToWrite.Count - 1)
        )

        $headerValues = [System.Array]::CreateInstance([object], @([int]1, [int]$ColumnsToWrite.Count), @([int]1, [int]1))
        for ($index = 0; $index -lt $ColumnsToWrite.Count; $index++) {
            $columnOffset = $index + 1
            $headerValues[1, $columnOffset] = $ColumnsToWrite[$index]
        }
        $headerRange.Value2 = $headerValues

        $maxRowNumber = [int](($Rows | Measure-Object -Property _RowNumber -Maximum).Maximum)
        $dataRowCount = $maxRowNumber - 1
        $dataValues = [System.Array]::CreateInstance([object], @([int]$dataRowCount, [int]$ColumnsToWrite.Count), @([int]1, [int]1))
        foreach ($row in $Rows) {
            $rowOffset = [int]$row._RowNumber - 1
            for ($index = 0; $index -lt $ColumnsToWrite.Count; $index++) {
                $columnName = $ColumnsToWrite[$index]
                $value = $row.$columnName
                $columnOffset = $index + 1
                $dataValues[$rowOffset, $columnOffset] = $value
            }
        }

        $dataRange = $worksheet.Range(
            $worksheet.Cells.Item(2, $startColumn),
            $worksheet.Cells.Item($maxRowNumber, $startColumn + $ColumnsToWrite.Count - 1)
        )
        $dataRange.Value2 = $dataValues

        $workbook.Save()
    }
    finally {
        if ($null -ne $dataRange) { Release-ComObject $dataRange }
        if ($null -ne $headerRange) { Release-ComObject $headerRange }
        if ($null -ne $usedRange) { Release-ComObject $usedRange }
        if ($null -ne $worksheet) { Release-ComObject $worksheet }
        if ($null -ne $workbook) {
            $workbook.Close($true)
            Release-ComObject $workbook
        }
        if ($null -ne $excel) {
            $excel.Quit()
            Release-ComObject $excel
        }

        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

function Convert-DutchNumberToDouble {
    param($Value)

    if ($null -eq $Value) {
        return $null
    }

    $text = ([string]$Value).Trim()
    if ([string]::IsNullOrWhiteSpace($text)) {
        return $null
    }

    return [double]::Parse($text, [System.Globalization.CultureInfo]::GetCultureInfo('nl-NL'))
}

function Read-TextFileLines {
    param([string]$Path)

    try {
        return @(Get-Content -Path $Path -Encoding UTF8)
    }
    catch {
        return @(Get-Content -Path $Path)
    }
}

function Parse-OverdrachtspuntValue {
    param([string]$Value)

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return $null
    }

    $match = [regex]::Match(
        $Value,
        'linksonder:\s*(?<x1>-?\d+(,\d+)?)\s*:\s*(?<y1>-?\d+(,\d+)?)\s*m\s*/\s*rechtsboven:\s*(?<x2>-?\d+(,\d+)?)\s*:\s*(?<y2>-?\d+(,\d+)?)\s*m',
        'IgnoreCase'
    )

    if (-not $match.Success) {
        return $null
    }

    $x1 = Convert-DutchNumberToDouble $match.Groups['x1'].Value
    $y1 = Convert-DutchNumberToDouble $match.Groups['y1'].Value
    $x2 = Convert-DutchNumberToDouble $match.Groups['x2'].Value
    $y2 = Convert-DutchNumberToDouble $match.Groups['y2'].Value

    return [pscustomobject]@{
        X1 = $x1
        Y1 = $y1
        X2 = $x2
        Y2 = $y2
        X = $x1
        Y = $y1
        IsExactPoint = (($x1 -eq $x2) -and ($y1 -eq $y2))
    }
}

function Finalize-LovionTextRecord {
    param(
        [hashtable]$Record,
        [System.Collections.Generic.List[object]]$Rows
    )

    if ($null -eq $Record -or $Record.Count -eq 0) {
        return
    }

    $ordered = [ordered]@{}
    foreach ($entry in $Record.GetEnumerator()) {
        $ordered[$entry.Key] = $entry.Value
    }

    $overdrachtspunt = Parse-OverdrachtspuntValue ([string]$ordered['Overdrachtspunt'])
    if ($overdrachtspunt) {
        $ordered['OverdrachtspuntRdX1'] = $overdrachtspunt.X1
        $ordered['OverdrachtspuntRdY1'] = $overdrachtspunt.Y1
        $ordered['OverdrachtspuntRdX2'] = $overdrachtspunt.X2
        $ordered['OverdrachtspuntRdY2'] = $overdrachtspunt.Y2
        $ordered['OverdrachtspuntIsPoint'] = $overdrachtspunt.IsExactPoint
        $ordered['FinalRdX'] = $overdrachtspunt.X1
        $ordered['FinalRdY'] = $overdrachtspunt.Y1
        $wgs84 = Convert-RdToWgs84 -X $overdrachtspunt.X1 -Y $overdrachtspunt.Y1
        $ordered['FinalLon'] = [math]::Round([double]$wgs84.Longitude, 8)
        $ordered['FinalLat'] = [math]::Round([double]$wgs84.Latitude, 8)
        $ordered['CoordinateSource'] = 'overdrachtspunt_text'
        $ordered['MatchStatus'] = if ($overdrachtspunt.IsExactPoint) { 'overdrachtspunt_exact' } else { 'overdrachtspunt_centroid' }
    }
    else {
        $ordered['OverdrachtspuntIsPoint'] = $false
        $ordered['FinalRdX'] = $null
        $ordered['FinalRdY'] = $null
        $ordered['FinalLon'] = $null
        $ordered['FinalLat'] = $null
        $ordered['CoordinateSource'] = 'none'
        $ordered['MatchStatus'] = 'overdrachtspunt_missing'
    }

    $Rows.Add([pscustomobject]$ordered)
}

function Read-LovionTextRows {
    param([string]$Path)

    Write-Info "Lees tekstexport: $Path"
    $lines = Read-TextFileLines -Path $Path
    $rows = New-Object System.Collections.Generic.List[object]
    $current = $null
    $recordNumber = 0

    foreach ($rawLine in $lines) {
        $line = [string]$rawLine
        $trimmed = $line.Trim()

        if ($trimmed -match '^LS Aansluiting\b') {
            Finalize-LovionTextRecord -Record $current -Rows $rows
            $recordNumber += 1
            $current = [ordered]@{
                _RecordNumber = $recordNumber
                RecordTitle = $trimmed
            }
            continue
        }

        if ($null -eq $current) {
            continue
        }

        if ([string]::IsNullOrWhiteSpace($trimmed)) {
            continue
        }

        if ($trimmed -match '^\((?<date>\d{1,2}-\d{1,2}-\d{4}\s+\d{2}:\d{2}:\d{2})\)$') {
            $current['ExportedAt'] = $Matches['date']
            continue
        }

        if ($trimmed -match '^(?<key>[^:]+):\s*(?<value>.*)$') {
            $key = $Matches['key'].Trim()
            $value = $Matches['value']
            $current[$key] = $value
            continue
        }

        if (-not $current.Contains('ExportedBy')) {
            $current['ExportedBy'] = $trimmed
        }
    }

    Finalize-LovionTextRecord -Record $current -Rows $rows
    return $rows
}

function Get-RowPropertyValue {
    param(
        [pscustomobject]$Row,
        [string]$Name
    )

    $property = $Row.PSObject.Properties[$Name]
    if ($property) {
        return $property.Value
    }

    return $null
}

function Get-DbfSafeFieldName {
    param(
        [string]$Name,
        [hashtable]$UsedNames
    )

    $normalized = Normalize-Text $Name
    $ascii = [regex]::Replace($normalized, '[^A-Z0-9]+', '_').Trim('_')
    if ([string]::IsNullOrWhiteSpace($ascii)) {
        $ascii = 'FIELD'
    }
    if ($ascii.Length -gt 10) {
        $ascii = $ascii.Substring(0, 10)
    }
    if ($ascii -match '^\d') {
        $ascii = 'F' + $ascii.Substring(0, [Math]::Min(9, $ascii.Length))
    }

    $candidate = $ascii
    $suffix = 1
    while ($UsedNames.ContainsKey($candidate)) {
        $suffixText = [string]$suffix
        $prefixLength = [Math]::Min(10 - $suffixText.Length, $ascii.Length)
        $candidate = $ascii.Substring(0, $prefixLength) + $suffixText
        $suffix += 1
    }

    $UsedNames[$candidate] = $true
    return $candidate
}

function New-DynamicDbfFieldDefinitions {
    param([object[]]$Rows)

    $propertyNames = New-Object System.Collections.Generic.List[string]
    foreach ($row in $Rows) {
        foreach ($property in $row.PSObject.Properties) {
            if ($property.Name -like '_*') {
                continue
            }

            if (-not $propertyNames.Contains($property.Name)) {
                $propertyNames.Add($property.Name)
            }
        }
    }

    $numericFields = @(
        'OverdrachtspuntRdX1',
        'OverdrachtspuntRdY1',
        'OverdrachtspuntRdX2',
        'OverdrachtspuntRdY2',
        'FinalRdX',
        'FinalRdY',
        'FinalLon',
        'FinalLat'
    )

    $usedNames = @{}
    $fields = New-Object System.Collections.Generic.List[object]
    foreach ($propertyName in $propertyNames) {
        $isNumeric = $numericFields -contains $propertyName
        if ($isNumeric) {
            $length = 18
            $decimals = if ($propertyName -like 'FinalL*') { 8 } else { 3 }
            $type = 'N'
        }
        else {
            $length = 1
            foreach ($row in $Rows) {
                $value = Get-RowPropertyValue -Row $row -Name $propertyName
                $length = [Math]::Max($length, [Math]::Min(254, ([string]$value).Length))
            }
            $decimals = 0
            $type = 'C'
        }

        $fields.Add([pscustomobject]@{
            Name = Get-DbfSafeFieldName -Name $propertyName -UsedNames $usedNames
            OriginalName = $propertyName
            Type = $type
            Length = $length
            Decimals = $decimals
        })
    }

    return $fields
}

function Write-DbfFileDynamic {
    param(
        [object[]]$Rows,
        [object[]]$Fields,
        [string]$Path
    )

    $recordLength = 1
    foreach ($field in $Fields) {
        $recordLength += [int]$field.Length
    }

    $headerLength = 32 + ($Fields.Count * 32) + 1
    $encoding = [System.Text.Encoding]::GetEncoding(1252)
    $stream = $null
    $writer = $null

    try {
        $stream = [System.IO.File]::Create($Path)
        $writer = New-Object System.IO.BinaryWriter($stream, $encoding)

        $now = Get-Date
        $writer.Write([byte]0x03)
        $writer.Write([byte]($now.Year - 1900))
        $writer.Write([byte]$now.Month)
        $writer.Write([byte]$now.Day)
        $writer.Write([int]$Rows.Count)
        $writer.Write([int16]$headerLength)
        $writer.Write([int16]$recordLength)
        $writer.Write([byte[]](0..17 | ForEach-Object { [byte]0 }))
        $writer.Write([byte]125)
        $writer.Write([byte]0)

        foreach ($field in $Fields) {
            $nameBytes = $encoding.GetBytes($field.Name)
            $fieldName = New-Object byte[] 11
            [array]::Copy($nameBytes, 0, $fieldName, 0, [math]::Min($nameBytes.Length, 10))
            $writer.Write($fieldName)
            $writer.Write([byte][char]$field.Type)
            $writer.Write([int]0)
            $writer.Write([byte]$field.Length)
            $writer.Write([byte]$field.Decimals)
            $writer.Write([byte[]](0..13 | ForEach-Object { [byte]0 }))
        }

        $writer.Write([byte]0x0D)

        foreach ($row in $Rows) {
            $writer.Write([byte]0x20)
            foreach ($field in $Fields) {
                $value = Get-RowPropertyValue -Row $row -Name $field.OriginalName
                $formatted = if ($field.Type -eq 'N') {
                    Format-DbfNumberValue -Value $value -Length $field.Length -Decimals $field.Decimals
                }
                else {
                    Format-DbfTextValue -Value ([string]$value) -Length $field.Length -Encoding $encoding
                }

                $writer.Write($encoding.GetBytes($formatted))
            }
        }

        $writer.Write([byte]0x1A)
    }
    finally {
        if ($writer) { $writer.Close() }
        elseif ($stream) { $stream.Close() }
    }
}

function Export-PointShapefileDynamic {
    param(
        [object[]]$Rows,
        [string]$OutputDirectory,
        [string]$BaseName
    )

    $pointRows = @($Rows | Where-Object { $null -ne $_.FinalRdX -and $null -ne $_.FinalRdY })
    if ($pointRows.Count -eq 0) {
        Write-Info 'Shapefile export overgeslagen: geen rijen met coordinaten.'
        return [pscustomobject]@{
            Directory = $null
            FieldMapPath = $null
        }
    }

    $shapeDirectory = Join-Path -Path $OutputDirectory -ChildPath ("{0}_{1}" -f $BaseName, 'shapefile')
    if (-not (Test-Path -Path $shapeDirectory)) {
        New-Item -Path $shapeDirectory -ItemType Directory | Out-Null
    }

    $shpPath = Join-Path -Path $shapeDirectory -ChildPath ($BaseName + '.shp')
    $shxPath = Join-Path -Path $shapeDirectory -ChildPath ($BaseName + '.shx')
    $dbfPath = Join-Path -Path $shapeDirectory -ChildPath ($BaseName + '.dbf')
    $prjPath = Join-Path -Path $shapeDirectory -ChildPath ($BaseName + '.prj')
    $cpgPath = Join-Path -Path $shapeDirectory -ChildPath ($BaseName + '.cpg')
    $fieldMapPath = Join-Path -Path $shapeDirectory -ChildPath ($BaseName + '_fieldmap.csv')

    $minX = ($pointRows | Measure-Object -Property FinalRdX -Minimum).Minimum
    $maxX = ($pointRows | Measure-Object -Property FinalRdX -Maximum).Maximum
    $minY = ($pointRows | Measure-Object -Property FinalRdY -Minimum).Minimum
    $maxY = ($pointRows | Measure-Object -Property FinalRdY -Maximum).Maximum

    $shpFileLengthWords = 50 + ($pointRows.Count * 14)
    $shxFileLengthWords = 50 + ($pointRows.Count * 4)

    $shpStream = $null
    $shxStream = $null
    $shpWriter = $null
    $shxWriter = $null

    try {
        $shpStream = [System.IO.File]::Create($shpPath)
        $shxStream = [System.IO.File]::Create($shxPath)
        $shpWriter = New-Object System.IO.BinaryWriter($shpStream)
        $shxWriter = New-Object System.IO.BinaryWriter($shxStream)

        Write-ShapefileHeader -Writer $shpWriter -FileLengthWords $shpFileLengthWords -ShapeType 1 -MinX $minX -MinY $minY -MaxX $maxX -MaxY $maxY
        Write-ShapefileHeader -Writer $shxWriter -FileLengthWords $shxFileLengthWords -ShapeType 1 -MinX $minX -MinY $minY -MaxX $maxX -MaxY $maxY

        $offsetWords = 50
        $recordNumber = 1
        foreach ($row in $pointRows) {
            Write-BigEndianInt32 -Writer $shxWriter -Value $offsetWords
            Write-BigEndianInt32 -Writer $shxWriter -Value 10

            Write-BigEndianInt32 -Writer $shpWriter -Value $recordNumber
            Write-BigEndianInt32 -Writer $shpWriter -Value 10
            $shpWriter.Write([int]1)
            $shpWriter.Write([double]$row.FinalRdX)
            $shpWriter.Write([double]$row.FinalRdY)

            $offsetWords += 14
            $recordNumber += 1
        }
    }
    finally {
        if ($shpWriter) { $shpWriter.Close() }
        elseif ($shpStream) { $shpStream.Close() }

        if ($shxWriter) { $shxWriter.Close() }
        elseif ($shxStream) { $shxStream.Close() }
    }

    $fields = New-DynamicDbfFieldDefinitions -Rows $pointRows
    Write-DbfFileDynamic -Rows $pointRows -Fields $fields -Path $dbfPath
    $fields | Select-Object Name, OriginalName, Type, Length, Decimals | Export-Csv -Path $fieldMapPath -NoTypeInformation -Encoding UTF8

    Set-Content -Path $cpgPath -Value '1252' -Encoding ASCII
    Set-Content -Path $prjPath -Encoding ASCII -Value 'PROJCS["Amersfoort / RD New",GEOGCS["Amersfoort",DATUM["Amersfoort",SPHEROID["Bessel 1841",6377397.155,299.1528128]],PRIMEM["Greenwich",0],UNIT["degree",0.0174532925199433]],PROJECTION["Oblique_Stereographic"],PARAMETER["latitude_of_origin",52.15616055555555],PARAMETER["central_meridian",5.38763888888889],PARAMETER["scale_factor",0.9999079],PARAMETER["false_easting",155000],PARAMETER["false_northing",463000],UNIT["metre",1]]'

    return [pscustomobject]@{
        Directory = $shapeDirectory
        FieldMapPath = $fieldMapPath
    }
}

function Write-BigEndianInt32 {
    param(
        [System.IO.BinaryWriter]$Writer,
        [int]$Value
    )

    $bytes = [System.BitConverter]::GetBytes([int]$Value)
    [array]::Reverse($bytes)
    $Writer.Write($bytes)
}

function Write-ShapefileHeader {
    param(
        [System.IO.BinaryWriter]$Writer,
        [int]$FileLengthWords,
        [int]$ShapeType,
        [double]$MinX,
        [double]$MinY,
        [double]$MaxX,
        [double]$MaxY
    )

    Write-BigEndianInt32 -Writer $Writer -Value 9994
    for ($index = 0; $index -lt 5; $index++) {
        Write-BigEndianInt32 -Writer $Writer -Value 0
    }
    Write-BigEndianInt32 -Writer $Writer -Value $FileLengthWords
    $Writer.Write([int]1000)
    $Writer.Write([int]$ShapeType)
    $Writer.Write([double]$MinX)
    $Writer.Write([double]$MinY)
    $Writer.Write([double]$MaxX)
    $Writer.Write([double]$MaxY)
    $Writer.Write([double]0)
    $Writer.Write([double]0)
    $Writer.Write([double]0)
    $Writer.Write([double]0)
}

function Format-DbfTextValue {
    param(
        [string]$Value,
        [int]$Length,
        [System.Text.Encoding]$Encoding
    )

    $text = if ($null -eq $Value) { '' } else { [string]$Value }
    while ($Encoding.GetByteCount($text) -gt $Length -and $text.Length -gt 0) {
        $text = $text.Substring(0, $text.Length - 1)
    }

    return $text.PadRight($Length)
}

function Format-DbfNumberValue {
    param(
        $Value,
        [int]$Length,
        [int]$Decimals
    )

    if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) {
        return ''.PadLeft($Length)
    }

    $number = [double]$Value
    $format = if ($Decimals -gt 0) { 'F' + $Decimals } else { 'F0' }
    $text = $number.ToString($format, [System.Globalization.CultureInfo]::InvariantCulture)
    if ($text.Length -gt $Length) {
        $text = $text.Substring(0, $Length)
    }

    return $text.PadLeft($Length)
}

function Write-DbfFile {
    param(
        [object[]]$Rows,
        [string]$Path
    )

    $fields = @(
        @{ Name = 'STATION';   Type = 'C'; Length = 20; Decimals = 0; Getter = { param($row) [string]$row.Station } },
        @{ Name = 'MASTERID';  Type = 'C'; Length = 25; Decimals = 0; Getter = { param($row) [string]$row.'Master Asset ID' } },
        @{ Name = 'ADRES';     Type = 'C'; Length = 80; Decimals = 0; Getter = { param($row) [string]$row.Adres } },
        @{ Name = 'STRAAT';    Type = 'C'; Length = 40; Decimals = 0; Getter = { param($row) [string]$row.Straatnaam } },
        @{ Name = 'POSTCODE';  Type = 'C'; Length = 10; Decimals = 0; Getter = { param($row) [string]$row.Postcode } },
        @{ Name = 'HUISNR';    Type = 'C'; Length = 12; Decimals = 0; Getter = { param($row) [string]$row.Huisnummer } },
        @{ Name = 'GEBRUIK';   Type = 'C'; Length = 25; Decimals = 0; Getter = { param($row) [string]$row.Gebruiksdoel } },
        @{ Name = 'COORDSRC';  Type = 'C'; Length = 12; Decimals = 0; Getter = { param($row) [string]$row.CoordinateSource } },
        @{ Name = 'MATCHSTAT'; Type = 'C'; Length = 24; Decimals = 0; Getter = { param($row) [string]$row.MatchStatus } },
        @{ Name = 'MATCHMODE'; Type = 'C'; Length = 18; Decimals = 0; Getter = { param($row) [string]$row.WfsMatchMode } },
        @{ Name = 'PDOKSTR';   Type = 'C'; Length = 1;  Decimals = 0; Getter = { param($row) if ($row.PdokStrictMatch) { 'Y' } else { 'N' } } },
        @{ Name = 'WFSDIST';   Type = 'N'; Length = 12; Decimals = 2; Getter = { param($row) $row.WfsDistanceM } },
        @{ Name = 'RDX';       Type = 'N'; Length = 18; Decimals = 3; Getter = { param($row) $row.FinalRdX } },
        @{ Name = 'RDY';       Type = 'N'; Length = 18; Decimals = 3; Getter = { param($row) $row.FinalRdY } }
    )

    $recordLength = 1
    foreach ($field in $fields) {
        $recordLength += [int]$field.Length
    }
    $headerLength = 32 + ($fields.Count * 32) + 1
    $encoding = [System.Text.Encoding]::GetEncoding(1252)
    $stream = $null
    $writer = $null

    try {
        $stream = [System.IO.File]::Create($Path)
        $writer = New-Object System.IO.BinaryWriter($stream, $encoding)

        $now = Get-Date
        $writer.Write([byte]0x03)
        $writer.Write([byte]($now.Year - 1900))
        $writer.Write([byte]$now.Month)
        $writer.Write([byte]$now.Day)
        $writer.Write([int]$Rows.Count)
        $writer.Write([int16]$headerLength)
        $writer.Write([int16]$recordLength)
        $writer.Write([byte[]](0..17 | ForEach-Object { [byte]0 }))
        $writer.Write([byte]125)
        $writer.Write([byte]0)

        foreach ($field in $fields) {
            $nameBytes = $encoding.GetBytes($field.Name)
            $fieldName = New-Object byte[] 11
            [array]::Copy($nameBytes, 0, $fieldName, 0, [math]::Min($nameBytes.Length, 10))
            $writer.Write($fieldName)
            $writer.Write([byte][char]$field.Type)
            $writer.Write([int]0)
            $writer.Write([byte]$field.Length)
            $writer.Write([byte]$field.Decimals)
            $writer.Write([byte[]](0..13 | ForEach-Object { [byte]0 }))
        }

        $writer.Write([byte]0x0D)

        foreach ($row in $Rows) {
            $writer.Write([byte]0x20)
            foreach ($field in $fields) {
                $value = & $field.Getter $row
                $formatted = if ($field.Type -eq 'N') {
                    Format-DbfNumberValue -Value $value -Length $field.Length -Decimals $field.Decimals
                }
                else {
                    Format-DbfTextValue -Value $value -Length $field.Length -Encoding $encoding
                }

                $writer.Write($encoding.GetBytes($formatted))
            }
        }

        $writer.Write([byte]0x1A)
    }
    finally {
        if ($writer) { $writer.Close() }
        elseif ($stream) { $stream.Close() }
    }
}

function Export-PointShapefile {
    param(
        [object[]]$Rows,
        [string]$OutputDirectory,
        [string]$BaseName
    )

    $pointRows = @($Rows | Where-Object { $null -ne $_.FinalRdX -and $null -ne $_.FinalRdY })
    if ($pointRows.Count -eq 0) {
        Write-Info 'Shapefile export overgeslagen: geen rijen met coordinaten.'
        return $null
    }

    $shapeDirectory = Join-Path -Path $OutputDirectory -ChildPath ("{0}_{1}" -f $BaseName, 'shapefile')
    if (-not (Test-Path -Path $shapeDirectory)) {
        New-Item -Path $shapeDirectory -ItemType Directory | Out-Null
    }

    $shpPath = Join-Path -Path $shapeDirectory -ChildPath ($BaseName + '.shp')
    $shxPath = Join-Path -Path $shapeDirectory -ChildPath ($BaseName + '.shx')
    $dbfPath = Join-Path -Path $shapeDirectory -ChildPath ($BaseName + '.dbf')
    $prjPath = Join-Path -Path $shapeDirectory -ChildPath ($BaseName + '.prj')
    $cpgPath = Join-Path -Path $shapeDirectory -ChildPath ($BaseName + '.cpg')

    $minX = ($pointRows | Measure-Object -Property FinalRdX -Minimum).Minimum
    $maxX = ($pointRows | Measure-Object -Property FinalRdX -Maximum).Maximum
    $minY = ($pointRows | Measure-Object -Property FinalRdY -Minimum).Minimum
    $maxY = ($pointRows | Measure-Object -Property FinalRdY -Maximum).Maximum

    $shpFileLengthWords = 50 + ($pointRows.Count * 14)
    $shxFileLengthWords = 50 + ($pointRows.Count * 4)

    $shpStream = $null
    $shxStream = $null
    $shpWriter = $null
    $shxWriter = $null

    try {
        $shpStream = [System.IO.File]::Create($shpPath)
        $shxStream = [System.IO.File]::Create($shxPath)
        $shpWriter = New-Object System.IO.BinaryWriter($shpStream)
        $shxWriter = New-Object System.IO.BinaryWriter($shxStream)

        Write-ShapefileHeader -Writer $shpWriter -FileLengthWords $shpFileLengthWords -ShapeType 1 -MinX $minX -MinY $minY -MaxX $maxX -MaxY $maxY
        Write-ShapefileHeader -Writer $shxWriter -FileLengthWords $shxFileLengthWords -ShapeType 1 -MinX $minX -MinY $minY -MaxX $maxX -MaxY $maxY

        $offsetWords = 50
        $recordNumber = 1
        foreach ($row in $pointRows) {
            Write-BigEndianInt32 -Writer $shxWriter -Value $offsetWords
            Write-BigEndianInt32 -Writer $shxWriter -Value 10

            Write-BigEndianInt32 -Writer $shpWriter -Value $recordNumber
            Write-BigEndianInt32 -Writer $shpWriter -Value 10
            $shpWriter.Write([int]1)
            $shpWriter.Write([double]$row.FinalRdX)
            $shpWriter.Write([double]$row.FinalRdY)

            $offsetWords += 14
            $recordNumber += 1
        }
    }
    finally {
        if ($shpWriter) { $shpWriter.Close() }
        elseif ($shpStream) { $shpStream.Close() }

        if ($shxWriter) { $shxWriter.Close() }
        elseif ($shxStream) { $shxStream.Close() }
    }

    Write-DbfFile -Rows $pointRows -Path $dbfPath
    Set-Content -Path $cpgPath -Value '1252' -Encoding ASCII
    Set-Content -Path $prjPath -Encoding ASCII -Value 'PROJCS["Amersfoort / RD New",GEOGCS["Amersfoort",DATUM["Amersfoort",SPHEROID["Bessel 1841",6377397.155,299.1528128]],PRIMEM["Greenwich",0],UNIT["degree",0.0174532925199433]],PROJECTION["Oblique_Stereographic"],PARAMETER["latitude_of_origin",52.15616055555555],PARAMETER["central_meridian",5.38763888888889],PARAMETER["scale_factor",0.9999079],PARAMETER["false_easting",155000],PARAMETER["false_northing",463000],UNIT["metre",1]]'

    return $shapeDirectory
}

if (-not (Test-Path -Path $OutputDirectory)) {
    New-Item -Path $OutputDirectory -ItemType Directory | Out-Null
}

if ($PSCmdlet.ParameterSetName -eq 'Text') {
    if (-not (Test-Path -Path $InputTextPath)) {
        throw "InputTextPath bestaat niet: $InputTextPath"
    }

    $rows = Read-LovionTextRows -Path $InputTextPath
    if ($MaxRows -gt 0 -and $rows.Count -gt $MaxRows) {
        $limitedRows = New-Object System.Collections.Generic.List[object]
        foreach ($item in ($rows | Select-Object -First $MaxRows)) {
            $limitedRows.Add($item)
        }
        $rows = $limitedRows
        Write-Info ("MaxRows actief, beperkt tot: {0}" -f $rows.Count)
    }

    Write-Info ("Records gelezen: {0}" -f $rows.Count)

    $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($InputTextPath)
    $csvPath = Join-Path -Path $OutputDirectory -ChildPath ("{0}_{1}.csv" -f $baseName, $timestamp)
    $geoJsonPath = Join-Path -Path $OutputDirectory -ChildPath ("{0}_{1}.geojson" -f $baseName, $timestamp)

    Export-CsvFile -Rows $rows -Path $csvPath
    Export-GeoJson -Rows $rows -Path $geoJsonPath
    $shapeResult = Export-PointShapefileDynamic -Rows $rows -OutputDirectory $OutputDirectory -BaseName ("{0}_{1}" -f $baseName, $timestamp)

    $matchedCount = @($rows | Where-Object { $_.CoordinateSource -ne 'none' }).Count

    Write-Host ''
    Write-Host 'Klaar.'
    Write-Host ("Records totaal          : {0}" -f $rows.Count)
    Write-Host ("Records met coordinaten : {0}" -f $matchedCount)
    Write-Host ("CSV output              : {0}" -f $csvPath)
    Write-Host ("GeoJSON output          : {0}" -f $geoJsonPath)
    if ($shapeResult.Directory) {
        Write-Host ("Shapefile output        : {0}" -f $shapeResult.Directory)
    }
    if ($shapeResult.FieldMapPath) {
        Write-Host ("Fieldmap output         : {0}" -f $shapeResult.FieldMapPath)
    }

    return
}

if (-not (Test-Path -Path $InputExcelPath)) {
    throw "InputExcelPath bestaat niet: $InputExcelPath"
}

$rows = Read-ExcelRows -Path $InputExcelPath -SheetName $WorksheetName
if ($MaxRows -gt 0 -and $rows.Count -gt $MaxRows) {
    $limitedRows = New-Object System.Collections.Generic.List[object]
    foreach ($item in ($rows | Select-Object -First $MaxRows)) {
        $limitedRows.Add($item)
    }
    $rows = $limitedRows
    Write-Info ("MaxRows actief, beperkt tot: {0}" -f $rows.Count)
}
Write-Info ("Datarijen gelezen: {0}" -f $rows.Count)

foreach ($row in $rows) {
    $isStreetFurniture = Test-IsStreetFurniture -Row $row
    Add-Member -InputObject $row -MemberType NoteProperty -Name AddressQuery -Value (Build-AddressQuery -Row $row)
    Add-Member -InputObject $row -MemberType NoteProperty -Name AddressKey -Value (Build-AddressKey -Row $row)
    Add-Member -InputObject $row -MemberType NoteProperty -Name StreetKey -Value (Build-StreetKeyFromValues -Street $row.Straatnaam -Postcode $row.Postcode -City $row.Woonplaats)
    Add-Member -InputObject $row -MemberType NoteProperty -Name IsStreetFurniture -Value $isStreetFurniture
    Add-Member -InputObject $row -MemberType NoteProperty -Name UsageCategory -Value $(if ($isStreetFurniture) { 'street_furniture' } else { 'service_connection' })
    Add-Member -InputObject $row -MemberType NoteProperty -Name PdokDisplayName -Value $null
    Add-Member -InputObject $row -MemberType NoteProperty -Name PdokType -Value $null
    Add-Member -InputObject $row -MemberType NoteProperty -Name PdokScore -Value $null
    Add-Member -InputObject $row -MemberType NoteProperty -Name PdokRdX -Value $null
    Add-Member -InputObject $row -MemberType NoteProperty -Name PdokRdY -Value $null
    Add-Member -InputObject $row -MemberType NoteProperty -Name PdokLon -Value $null
    Add-Member -InputObject $row -MemberType NoteProperty -Name PdokLat -Value $null
    Add-Member -InputObject $row -MemberType NoteProperty -Name PdokStrictMatch -Value $false
    Add-Member -InputObject $row -MemberType NoteProperty -Name PdokMatchReason -Value $null
    Add-Member -InputObject $row -MemberType NoteProperty -Name WfsFeatureId -Value $null
    Add-Member -InputObject $row -MemberType NoteProperty -Name WfsMatchMode -Value $null
    Add-Member -InputObject $row -MemberType NoteProperty -Name WfsDistanceM -Value $null
    Add-Member -InputObject $row -MemberType NoteProperty -Name FinalRdX -Value $null
    Add-Member -InputObject $row -MemberType NoteProperty -Name FinalRdY -Value $null
    Add-Member -InputObject $row -MemberType NoteProperty -Name FinalLon -Value $null
    Add-Member -InputObject $row -MemberType NoteProperty -Name FinalLat -Value $null
    Add-Member -InputObject $row -MemberType NoteProperty -Name CoordinateSource -Value 'none'
    Add-Member -InputObject $row -MemberType NoteProperty -Name MatchStatus -Value 'unmatched'
}

if (-not $SkipGeocoding) {
    $queries = @(
        $rows |
            ForEach-Object { Build-PdokRequestUri -Row $_ } |
            Select-Object -Unique |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    )

    Write-Info ("Geocodeer unieke adressen via PDOK: {0}" -f $queries.Count)
    try {
        $geocodeCache = Invoke-PdokQueryBatch -Queries $queries -MaxConcurrency $GeocodeConcurrency -DelayBetweenDispatchMs $GeocodeDelayMs
    }
    catch {
        Write-WarnLine "PDOK geocoding batch mislukt: $($_.Exception.Message)"
        $geocodeCache = @{}
        foreach ($query in $queries) {
            try {
                if ($query -like 'http*://*') {
                    $geocodeCache[$query] = @(Search-PdokAddress -RequestUri $query)
                }
                else {
                    $geocodeCache[$query] = @(Search-PdokAddress -Query $query)
                }
            }
            catch {
                Write-WarnLine "PDOK geocoding mislukt voor '$query': $($_.Exception.Message)"
                $geocodeCache[$query] = @()
            }

            if ($GeocodeDelayMs -gt 0) {
                Start-Sleep -Milliseconds $GeocodeDelayMs
            }
        }
    }

    foreach ($row in $rows) {
        $pdokRequestUri = Build-PdokRequestUri -Row $row
        $docs = @($geocodeCache[$pdokRequestUri])
        $result = Select-PdokCandidateForRow -Row $row -Docs $docs
        if ($null -eq $result) {
            $row.MatchStatus = 'pdok_not_found'
            continue
        }

        $candidate = $result.Candidate
        $row.PdokDisplayName = $candidate.DisplayName
        $row.PdokType = $candidate.Type
        $row.PdokScore = $candidate.Score
        $row.PdokRdX = $candidate.RdX
        $row.PdokRdY = $candidate.RdY
        $row.PdokLon = $candidate.Lon
        $row.PdokLat = $candidate.Lat
        $row.PdokStrictMatch = $result.StrictMatch
        $row.PdokMatchReason = $result.Reason
        $isPdokTolerated = Test-IsPdokToleratedReason -Reason $result.Reason

        if ($row.IsStreetFurniture) {
            if ($result.StrictMatch) {
                $row.MatchStatus = 'street_furniture_seed_exact'
            }
            elseif ($isPdokTolerated) {
                $row.MatchStatus = 'street_furniture_seed_tolerant_house_suffix'
            }
            else {
                $row.MatchStatus = 'street_furniture_seed_fuzzy'
            }
            continue
        }

        if (-not $result.StrictMatch -and -not $isPdokTolerated) {
            $row.MatchStatus = 'pdok_not_exact'
            continue
        }

        $row.FinalRdX = $candidate.RdX
        $row.FinalRdY = $candidate.RdY
        $row.FinalLon = $candidate.Lon
        $row.FinalLat = $candidate.Lat
        if ($result.StrictMatch) {
            $row.CoordinateSource = 'pdok'
            $row.MatchStatus = 'pdok_exact'
        }
        else {
            $row.CoordinateSource = 'pdok_tolerant'
            $row.MatchStatus = 'pdok_tolerant_house_suffix'
        }
    }
}
else {
    Write-Info 'PDOK geocoding overgeslagen.'
}

$wfsPrepared = @()
if (-not $SkipWfs) {
    try {
        $rawFeatures = @()
        if (-not [string]::IsNullOrWhiteSpace($WfsGeoJsonPath)) {
            if (-not (Test-Path -Path $WfsGeoJsonPath)) {
                throw "WfsGeoJsonPath bestaat niet: $WfsGeoJsonPath"
            }

            $rawFeatures = Get-WfsFeaturesFromGeoJson -Path $WfsGeoJsonPath
        }
        else {
            $bbox = New-BoundingBox -Rows $rows -BufferMeters $WfsBufferMeters
            if ($null -eq $bbox) {
                Write-WarnLine 'WFS ophalen overgeslagen: geen RD-coordinaten beschikbaar om een bbox te bouwen.'
            }
            else {
                $rawFeatures = Get-WfsFeaturesFromService -BaseUrl $WfsUrl -LayerName $WfsLayerName -BoundingBox $bbox
            }
        }

        $wfsPrepared = Prepare-WfsFeatures -RawFeatures $rawFeatures
        Write-Info ("WFS-features bruikbaar voor matching: {0}" -f $wfsPrepared.Count)
    }
    catch {
        Write-WarnLine "WFS matching overgeslagen: $($_.Exception.Message)"
        $wfsPrepared = @()
    }
}
else {
    Write-Info 'WFS matching overgeslagen.'
}

if ($wfsPrepared.Count -gt 0) {
    $wfsByAddress = @{}
    $wfsByStreet = @{}
    foreach ($feature in $wfsPrepared) {
        if (-not [string]::IsNullOrWhiteSpace($feature.AddressKey)) {
            if (-not $wfsByAddress.ContainsKey($feature.AddressKey)) {
                $wfsByAddress[$feature.AddressKey] = New-Object System.Collections.Generic.List[object]
            }

            $wfsByAddress[$feature.AddressKey].Add($feature)
        }

        if (-not [string]::IsNullOrWhiteSpace($feature.StreetKey)) {
            if (-not $wfsByStreet.ContainsKey($feature.StreetKey)) {
                $wfsByStreet[$feature.StreetKey] = New-Object System.Collections.Generic.List[object]
            }

            $wfsByStreet[$feature.StreetKey].Add($feature)
        }
    }

    foreach ($row in $rows) {
        $selectedFeature = $null
        $matchMode = $null
        $distance = $null

        if ($null -eq $row.PdokRdX -or $null -eq $row.PdokRdY) {
            continue
        }

        if ($row.IsStreetFurniture) {
            $streetCandidates = @()
            if (-not [string]::IsNullOrWhiteSpace($row.StreetKey) -and $wfsByStreet.ContainsKey($row.StreetKey)) {
                $streetCandidates = $wfsByStreet[$row.StreetKey].ToArray()
            }

            if ($streetCandidates.Count -gt 0) {
                $nearestStreet = Get-NearestWfsFeature -Features $streetCandidates -BaseX $row.PdokRdX -BaseY $row.PdokRdY -MaxDistanceMeters $StreetFurnitureSnapDistanceMeters
                if ($nearestStreet) {
                    $selectedFeature = $nearestStreet.Feature
                    $distance = $nearestStreet.Distance
                    $matchMode = 'street_furniture_street'
                }
            }

            if ($null -eq $selectedFeature) {
                $nearestAny = Get-NearestWfsFeature -Features $wfsPrepared -BaseX $row.PdokRdX -BaseY $row.PdokRdY -MaxDistanceMeters $StreetFurnitureSnapDistanceMeters
                if ($nearestAny) {
                    $selectedFeature = $nearestAny.Feature
                    $distance = $nearestAny.Distance
                    $matchMode = 'street_furniture_nearest'
                }
            }

            if ($null -eq $selectedFeature) {
                $row.FinalRdX = $null
                $row.FinalRdY = $null
                $row.FinalLon = $null
                $row.FinalLat = $null
                $row.CoordinateSource = 'none'
                $row.MatchStatus = 'street_furniture_no_wfs'
                continue
            }
        }
        else {
            if (-not $row.PdokStrictMatch -and -not (Test-IsPdokToleratedReason -Reason $row.PdokMatchReason)) {
                continue
            }

            if (-not [string]::IsNullOrWhiteSpace($row.AddressKey) -and $wfsByAddress.ContainsKey($row.AddressKey)) {
                $addressCandidates = $wfsByAddress[$row.AddressKey].ToArray()
                $nearestAddress = Get-NearestWfsFeature -Features $addressCandidates -BaseX $row.PdokRdX -BaseY $row.PdokRdY
                if ($nearestAddress) {
                    $selectedFeature = $nearestAddress.Feature
                    $distance = $nearestAddress.Distance
                    $matchMode = 'address_key'
                }
            }

            if ($null -eq $selectedFeature) {
                $nearest = Get-NearestWfsFeature -Features $wfsPrepared -BaseX $row.PdokRdX -BaseY $row.PdokRdY -MaxDistanceMeters $SnapDistanceMeters
                if ($nearest) {
                    $selectedFeature = $nearest.Feature
                    $distance = $nearest.Distance
                    $matchMode = 'nearest'
                }
            }

            if ($null -eq $selectedFeature) {
                continue
            }
        }

        $row.WfsFeatureId = $selectedFeature.FeatureId
        $row.WfsMatchMode = $matchMode
        $row.WfsDistanceM = if ($null -ne $distance) { [math]::Round([double]$distance, 2) } else { $null }
        $row.FinalRdX = $selectedFeature.RdX
        $row.FinalRdY = $selectedFeature.RdY

        $wgs84 = Convert-RdToWgs84 -X $selectedFeature.RdX -Y $selectedFeature.RdY
        $row.FinalLon = [math]::Round([double]$wgs84.Longitude, 8)
        $row.FinalLat = [math]::Round([double]$wgs84.Latitude, 8)
        $row.CoordinateSource = 'wfs'
        $row.MatchStatus = $matchMode
    }
}

$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$baseName = [System.IO.Path]::GetFileNameWithoutExtension($InputExcelPath)
$csvPath = Join-Path -Path $OutputDirectory -ChildPath ("{0}_{1}.csv" -f $baseName, $timestamp)
$xlsxPath = Join-Path -Path $OutputDirectory -ChildPath ("{0}_{1}.xlsx" -f $baseName, $timestamp)
$geoJsonPath = Join-Path -Path $OutputDirectory -ChildPath ("{0}_{1}.geojson" -f $baseName, $timestamp)

Export-CsvFile -Rows $rows -Path $csvPath
Export-GeoJson -Rows $rows -Path $geoJsonPath

$excelColumns = @(
    'UsageCategory',
    'AddressQuery',
    'AddressKey',
    'StreetKey',
    'PdokDisplayName',
    'PdokType',
    'PdokScore',
    'PdokRdX',
    'PdokRdY',
    'PdokLon',
    'PdokLat',
    'PdokStrictMatch',
    'PdokMatchReason',
    'WfsFeatureId',
    'WfsMatchMode',
    'WfsDistanceM',
    'FinalRdX',
    'FinalRdY',
    'FinalLon',
    'FinalLat',
    'CoordinateSource',
    'MatchStatus'
)

Export-AugmentedWorkbook -SourcePath $InputExcelPath -TargetPath $xlsxPath -SheetName $WorksheetName -Rows $rows -ColumnsToWrite $excelColumns
$shapePath = Export-PointShapefile -Rows $rows -OutputDirectory $OutputDirectory -BaseName ("{0}_{1}" -f $baseName, $timestamp)

$matchedCount = @($rows | Where-Object { $_.CoordinateSource -ne 'none' }).Count
$wfsCount = @($rows | Where-Object { $_.CoordinateSource -eq 'wfs' }).Count
$pdokCount = @($rows | Where-Object { $_.CoordinateSource -eq 'pdok' }).Count
$pdokTolerantCount = @($rows | Where-Object { $_.CoordinateSource -eq 'pdok_tolerant' }).Count

Write-Host ''
Write-Host 'Klaar.'
Write-Host ("Rijen totaal            : {0}" -f $rows.Count)
Write-Host ("Rijen met coordinaten   : {0}" -f $matchedCount)
Write-Host ("Waarvan uit WFS         : {0}" -f $wfsCount)
Write-Host ("Waarvan uit PDOK exact  : {0}" -f $pdokCount)
Write-Host ("Waarvan uit PDOK tol.   : {0}" -f $pdokTolerantCount)
Write-Host ("CSV output              : {0}" -f $csvPath)
Write-Host ("Excel output            : {0}" -f $xlsxPath)
Write-Host ("GeoJSON output          : {0}" -f $geoJsonPath)
if ($shapePath) {
    Write-Host ("Shapefile output        : {0}" -f $shapePath)
}
