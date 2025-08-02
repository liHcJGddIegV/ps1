# Define the path to your netstat output file
$netstatFile = "C:\Users\YGonzalez\OneDrive - Invenergy LLC\Desktop\netstat_established.txt"

# Read all lines from the file
$content = Get-Content $netstatFile

# Array to store unique remote IP addresses
$remoteIPs = @()

# Process each line in the file
foreach ($line in $content) {
    # Use regex to extract all IPv4 addresses in the line
    $matches = [regex]::Matches($line, "\b(\d{1,3}\.){3}\d{1,3}\b")
    if ($matches.Count -ge 2) {
        # The first IP is typically the local endpoint and the second is the remote endpoint
        $remoteIP = $matches[1].Value
        
        # Filter out local/loopback and common private addresses
        if ($remoteIP -notmatch "^127\." -and $remoteIP -notmatch "^192\.168\." -and $remoteIP -notmatch "^10\.") {
            if (-not $remoteIPs.Contains($remoteIP)) {
                $remoteIPs += $remoteIP
            }
        }
    }
}

# Display the unique remote IP addresses found
Write-Output "Found remote IP addresses:"
$remoteIPs | ForEach-Object { Write-Output $_ }
Write-Output "---------------------------------------------"

# Loop through each remote IP and query ARIN's RDAP service
foreach ($ip in $remoteIPs) {
    Write-Output "Querying RDAP for IP: $ip"
    try {
        $result = Invoke-RestMethod -Uri "https://rdap.arin.net/registry/ip/$ip"
        $org = $result.name
        $start = $result.startAddress
        $end = $result.endAddress
        Write-Output "IP: $ip is in the range $start - $end, allocated to: $org"
    }
    catch {
        Write-Output "Failed to query RDAP for IP: $ip"
    }
    Write-Output "---------------------------------------------"
}
