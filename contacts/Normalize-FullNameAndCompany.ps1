# PowerShell script to normalize Full Name and assign Company based on email domain
param (
    [switch]$DryRun,          # Simulate changes without saving
    [switch]$VerboseOutput    # Provide detailed logs
)

# Initialize Outlook Application
try {
    $Outlook = New-Object -ComObject Outlook.Application
    $Namespace = $Outlook.GetNamespace("MAPI")
    $ContactsFolder = $Namespace.GetDefaultFolder(10)  # olFolderContacts = 10
} catch {
    Write-Host "Error: Outlook is not installed or could not be accessed." -ForegroundColor Red
    exit 1
}

# Counters
$updatedCount = 0
$errorCount   = 0

Write-Host "Starting Outlook contacts cleanup..."

# Define regex pattern to extract Full Name and Company
$pattern = "^(.+),\s(.+)\s\((.+)\)$"

foreach ($Contact in $ContactsFolder.Items) {
    try {
        # Get Full Name
        $FullName = $Contact.FullName
        $Email = $Contact.Email1Address  # Assuming primary email

        if ($FullName -match $pattern) {
            $LastName = $matches[1]
            $FirstName = $matches[2]
            $CompanyName = $matches[3]

            # Construct the new Full Name
            $NewFullName = "$LastName, $FirstName"

            # Determine if Company should be assigned based on Email Domain
            $AssignCompany = $false
            if ($Email -match "@gevernova.com$" -or $Email -match "@ge.com$") {
                $AssignCompany = $true
            }

            if ($VerboseOutput) {
                Write-Host "Processing: $FullName → $NewFullName (Email: $Email)"
            }

            # Apply changes if not DryRun
            if (-not $DryRun) {
                $Contact.FullName = $NewFullName
                if ($AssignCompany) {
                    $Contact.CompanyName = $CompanyName
                }
                $Contact.Save()
            }

            # Corrected log message
            if ($AssignCompany) {
                Write-Host "Updated: '$FullName' → '$NewFullName' (Company: $CompanyName)" -ForegroundColor Green
            } else {
                Write-Host "Updated: '$FullName' → '$NewFullName' (Company: N/A)" -ForegroundColor Green
            }

            $updatedCount++
        } else {
            # Updated message: Show "Skipping due to company detected" if company already exists
            if ($Contact.CompanyName) {
                Write-Host "Skipping: '$FullName' (Due to company detected: $($Contact.CompanyName))" -ForegroundColor Yellow
            } else {
                Write-Host "Skipping: '$FullName' (No company detected)" -ForegroundColor Yellow
            }
        }
    } catch {
        $errorCount++
        Write-Host "Error updating contact: $_" -ForegroundColor Red
    }
}

Write-Host "`nProcess complete: $updatedCount contacts updated, $errorCount errors encountered."
