param(
  [string]$Url = "http://api.link.heroesports.com/hfins/v2/0/anchor/anchor-profiles/list-records?page=0&size=10",
  [int]$DelayMs = 200,
  [int]$TimeoutSeconds = 10,
  [string]$TokenUrl = "http://api.link.heroesports.com/oauth/oauth/token",
  [string]$GrantType = "client_credentials",
  [string]$ClientId = "interface",
  [string]$ClientSecret = "pN2fD7vK4bM9",
  [string]$BearerToken = "cd93af67-4f48-4e9e-845d-1a43fb5640e3",
  [string]$Cookie = "",
  [string]$LogPath = ""
)

if ([string]::IsNullOrWhiteSpace($LogPath)) {
  $LogPath = Join-Path $PSScriptRoot ("continuous-request-hfins-{0}.log" -f (Get-Date -Format "yyyyMMdd-HHmmss"))
}

$total = 0
$success = 0
$failed = 0
$statusCounts = @{}
$headers = @{}

function Get-AccessToken {
  param(
    [string]$TokenUrl,
    [string]$GrantType,
    [string]$ClientId,
    [string]$ClientSecret,
    [int]$TimeoutSeconds
  )

  $body = @{
    grant_type = $GrantType
    client_id = $ClientId
    client_secret = $ClientSecret
  }

  $response = Invoke-WebRequest `
    -Uri $TokenUrl `
    -Method Post `
    -Body $body `
    -ContentType "application/x-www-form-urlencoded" `
    -TimeoutSec $TimeoutSeconds `
    -UseBasicParsing `
    -ErrorAction Stop

  $json = $response.Content | ConvertFrom-Json
  if (-not $json.access_token) {
    throw "Token response does not contain access_token. Response: $($response.Content)"
  }

  return $json.access_token
}

function Get-ErrorResponseBody {
  param($Response)

  if (-not $Response) {
    return ""
  }

  try {
    $stream = $Response.GetResponseStream()
    if (-not $stream) {
      return ""
    }

    $reader = [System.IO.StreamReader]::new($stream)
    return (($reader.ReadToEnd() -replace "`r|`n", " ") -replace '"', '""')
  } catch {
    return ""
  }
}

if ([string]::IsNullOrWhiteSpace($BearerToken)) {
  Write-Host "Fetching token: $TokenUrl"
  $BearerToken = Get-AccessToken `
    -TokenUrl $TokenUrl `
    -GrantType $GrantType `
    -ClientId $ClientId `
    -ClientSecret $ClientSecret `
    -TimeoutSeconds $TimeoutSeconds
  Write-Host "Token fetched."
}

if (-not [string]::IsNullOrWhiteSpace($BearerToken)) {
  $headers["Authorization"] = "bearer $BearerToken"
}

if (-not [string]::IsNullOrWhiteSpace($Cookie)) {
  $headers["Cookie"] = $Cookie
}

"time,index,status,latency_ms,ok,error" | Out-File -FilePath $LogPath -Encoding utf8
Write-Host "Requesting: $Url"
Write-Host "Delay: ${DelayMs}ms, Timeout: ${TimeoutSeconds}s"
Write-Host "Log: $LogPath"
if ($headers.Count -gt 0) {
  Write-Host "Headers: $($headers.Keys -join ', ')"
}
Write-Host "Press Ctrl+C to stop."
Write-Host ""

while ($true) {
  $total++
  $startedAt = Get-Date
  $sw = [System.Diagnostics.Stopwatch]::StartNew()

  try {
    $response = Invoke-WebRequest -Uri $Url -Method Get -Headers $headers -TimeoutSec $TimeoutSeconds -UseBasicParsing -ErrorAction Stop
    $sw.Stop()

    $status = [int]$response.StatusCode
    $ok = ($status -ge 200 -and $status -lt 400)

    if (-not $statusCounts.ContainsKey($status)) {
      $statusCounts[$status] = 0
    }
    $statusCounts[$status]++

    if ($ok) {
      $success++
      $line = "{0} #{1} status={2} latency={3}ms ok={4} success={5} failed={6}" -f `
        $startedAt.ToString("HH:mm:ss.fff"), $total, $status, $sw.ElapsedMilliseconds, $ok, $success, $failed
      Write-Host $line -ForegroundColor Green
    } else {
      $failed++
      $line = "{0} #{1} status={2} latency={3}ms ok={4} success={5} failed={6}" -f `
        $startedAt.ToString("HH:mm:ss.fff"), $total, $status, $sw.ElapsedMilliseconds, $ok, $success, $failed
      Write-Host $line -ForegroundColor Red
    }

    "{0},{1},{2},{3},{4}," -f $startedAt.ToString("yyyy-MM-dd HH:mm:ss.fff"), $total, $status, $sw.ElapsedMilliseconds, $ok |
      Out-File -FilePath $LogPath -Append -Encoding utf8
  } catch {
    $sw.Stop()
    $status = "ERROR"
    $errorMessage = ($_.Exception.Message -replace "`r|`n", " ")

    if ($_.Exception.Response -and $_.Exception.Response.StatusCode) {
      $status = [int]$_.Exception.Response.StatusCode
      if (-not $statusCounts.ContainsKey($status)) {
        $statusCounts[$status] = 0
      }
      $statusCounts[$status]++
    }

    $errorBody = Get-ErrorResponseBody -Response $_.Exception.Response
    $failed++

    if ([string]::IsNullOrWhiteSpace($errorBody)) {
      $line = "{0} #{1} status={2} latency={3}ms success={4} failed={5} message={6}" -f `
        $startedAt.ToString("HH:mm:ss.fff"), $total, $status, $sw.ElapsedMilliseconds, $success, $failed, $errorMessage
    } else {
      $line = "{0} #{1} status={2} latency={3}ms success={4} failed={5} message={6} body={7}" -f `
        $startedAt.ToString("HH:mm:ss.fff"), $total, $status, $sw.ElapsedMilliseconds, $success, $failed, $errorMessage, $errorBody
    }
    Write-Host $line -ForegroundColor Red

    "{0},{1},{2},{3},False,""{4} {5}""" -f $startedAt.ToString("yyyy-MM-dd HH:mm:ss.fff"), $total, $status, $sw.ElapsedMilliseconds, $errorMessage, $errorBody |
      Out-File -FilePath $LogPath -Append -Encoding utf8
  }

  if ($total % 50 -eq 0) {
    $summary = ($statusCounts.GetEnumerator() | Sort-Object Name | ForEach-Object { "$($_.Name)=$($_.Value)" }) -join " "
    Write-Host ("SUMMARY total={0} success={1} failed={2} {3}" -f $total, $success, $failed, $summary) -ForegroundColor Cyan
  }

  Start-Sleep -Milliseconds $DelayMs
}
