<#
.SYNOPSIS
    Tests connection quality to a specified IP and reports packet loss, latency stats, and jitter.

.PARAMETER IPAddress
    The target IP address to ping.

.PARAMETER PingCount
    Number of pings to send.

.PARAMETER DelaySeconds
    Seconds to wait between pings.
#>

param(
    [string]$IPAddress    = "10.10.10.11",
    [int]   $PingCount    = 10,
    [int]   $DelaySeconds = 1
)

# Array to hold each ping result
$results = @()

Write-Host "Pinging $IPAddress $PingCount time(s), one ping every $DelaySeconds second(s)...`n"

for ($i = 1; $i -le $PingCount; $i++) {
    $timestamp = Get-Date
    try {
        $reply = Test-Connection -ComputerName $IPAddress -Count 1 -ErrorAction Stop
        $rtt   = $reply.ResponseTime
    }
    catch {
        # On failure (timeout or unreachable), mark as lost
        $rtt = $null
    }

    $results += [PSCustomObject]@{
        Timestamp     = $timestamp
        ResponseTime  = $rtt     # in ms; $null if lost
    }

    Start-Sleep -Seconds $DelaySeconds
}

# Calculate statistics
$totalSent   = $PingCount
$successful  = $results | Where-Object ResponseTime -ne $null
$received    = $successful.Count
$lost        = $totalSent - $received
$lossPercent = [math]::Round(($lost / $totalSent) * 100, 2)

if ($received -gt 0) {
    $avg   = [math]::Round(($successful.ResponseTime | Measure-Object -Average).Average, 2)
    $min   = [math]::Min($successful.ResponseTime)
    $max   = [math]::Max($successful.ResponseTime)
    # Jitter = average absolute deviation from mean
    $jitter = [math]::Round(
        ($successful.ResponseTime |
            ForEach-Object { [math]::Abs($_ - $avg) } |
            Measure-Object -Average).Average
    , 2)
} else {
    $avg = $min = $max = $jitter = $null
}

# Output summary
"`n=== Connection Quality Summary ==="
"Target:         $IPAddress"
"Sent:           $totalSent"
"Received:       $received"
"Lost:           $lost ($lossPercent`%)"
"Avg Latency:    $avg ms"
"Min Latency:    $min ms"
"Max Latency:    $max ms"
"Jitter:         $jitter ms"

# (Optional) Export detailed results to CSV:
$results | Export-Csv -Path "$env:USERPROFILE\ping_results.csv" -NoTypeInformation
