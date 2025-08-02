param (
    [switch]$DryRun,          # Simulate changes without saving
    [switch]$VerboseOutput    # Provide detailed logs
)

# Define the CSV file path (update if needed)
$csvPath = "C:\Users\YGonzalez\Downloads\Lazbuddie TPRG_Master.csv"

# Verify CSV file exists
if (-Not (Test-Path $csvPath)) {
    Write-Host "Error: CSV file not found at $csvPath" -ForegroundColor Red
    exit
}

if ($VerboseOutput) {
    Write-Host "CSV file found at $csvPath. Importing CSV..." -ForegroundColor Cyan
}

# Import the CSV
$vendors = Import-Csv -Path $csvPath

# Extract vendor names (column headers) excluding the first column which contains field labels
$vendorNames = $vendors | Get-Member -MemberType NoteProperty | Select-Object -ExpandProperty Name
$vendorNames = $vendorNames | Where-Object { $_ -ne "Vendor Name (Be specific, needs to Include Division)" } # Exclude first column

if ($VerboseOutput) {
    Write-Host "Vendor names extracted:" -ForegroundColor Cyan
    $vendorNames | ForEach-Object { Write-Host $_ -ForegroundColor Cyan }
}

# Define CC recipients
$ccRecipients = "sclee@aep.com; jrbrakus@aep.com; rfarrar@aep.com; Harry.Gauthier@mortenson.com; sarah.mitchell@mortenson.com; Jung.Lee@mortenson.com"

# Create an Outlook Application COM object
$outlook = New-Object -ComObject Outlook.Application

# Process each vendor
foreach ($vendorName in $vendorNames) {
    if ($VerboseOutput) {
        Write-Host "Processing vendor: $vendorName" -ForegroundColor Cyan
    }
    
    # Extract vendor-specific details into a hashtable where the key comes from the first column
    $vendorInfo = @{}
    foreach ($row in $vendors) {
        $fieldName = $row."Vendor Name (Be specific, needs to Include Division)"
        $vendorInfo[$fieldName] = $row.$vendorName
    }

    # Ensure necessary fields exist
    $primaryContactName = $vendorInfo["Vendor Primary Contact Name"]
    $primaryContactEmail = $vendorInfo["Vendor Primary Contact Email"]
    $primaryContactPhone = $vendorInfo["Vendor Primary Contact Phone"]
    $technicalContactName = $vendorInfo["Vendor Primary Technical Contact Name (We will need a network SME)"]
    $technicalContactEmail = $vendorInfo["Vendor Primary Technical Contact Email (We will need a network SME)"]
    $technicalContactPhone = $vendorInfo["Vendor Primary Technical Contact Phone (We will need a network SME)"]
    $projectName = $vendorInfo["Project Name"]
    $description = $vendorInfo["Description: Clearly define what they are providing or doing, including connectivity details, data type and volume, access method (onsite, remote, or both), and any involvement of fourth-party data traffic."]

    # Skip vendors with missing primary contact email
    if (-not $primaryContactEmail) {
        Write-Host "Skipping $vendorName due to missing email" -ForegroundColor Yellow
        continue
    }

    # Create a new Mail item (0 = olMailItem)
    $mailItem = $outlook.CreateItem(0)

    # Enable HTML formatting to ensure signature appears
    $mailItem.BodyFormat = 2  # 2 = olFormatHTML

    # Prepare the email body (now using HTML)
    $body = @"
<html>
<body>
<p>Hello $primaryContactName,</p>

<p>I hope you're doing well. I'm reaching out on behalf of the <strong>$projectName</strong> project to confirm
we have the correct information for <strong>$vendorName</strong>.</p>

<p>We want to ensure our records are accurate so we can collaborate effectively.</p>

<p><strong>Here is what we have on file:</strong></p>

<ul>
<li><strong>Vendor:</strong> $vendorName</li>
</ul>

<p><strong>Primary Contact:</strong></p>
<ul>
<li>Name: $primaryContactName</li>
<li>Email: <a href='mailto:$primaryContactEmail'>$primaryContactEmail</a></li>
<li>Phone: $primaryContactPhone</li>
</ul>

<p><strong>Technical Contact (Network SME):</strong></p>
<ul>
<li>Name: $technicalContactName</li>
<li>Email: <a href='mailto:$technicalContactEmail'>$technicalContactEmail</a></li>
<li>Phone: $technicalContactPhone</li>
</ul>

<p><strong>Project:</strong> $projectName</p>

<p><strong>Description:</strong><br>$description</p>

<p>Could you please review the above details and let me know if anything needs to be updated or corrected?</p>

</body>
</html>
"@

# Set email properties
$mailItem.Subject = "Lazbuddie Third Party Review Contacts Validation"
$mailItem.HTMLBody = $body  # Use HTML body to retain signature
$mailItem.To = $primaryContactEmail
$mailItem.CC = $ccRecipients  # Add CC recipients

    if ($DryRun) {
        Write-Host "[Dry Run] Draft email for $primaryContactName ($vendorName) would be created here." -ForegroundColor Magenta
    }
    else {
        # Display the draft email (does not send automatically)
        $mailItem.Display()
        Write-Host "Draft email created for: $primaryContactName ($vendorName)" -ForegroundColor Green
    }
}
