param (
    [switch]$DryRun,          # Simulate changes without saving
    [switch]$VerboseOutput    # Provide detailed logs
)

# Notify dry-run mode if applicable
if ($DryRun) {
    Write-Host "DryRun mode: No changes will be saved." -ForegroundColor Cyan
}

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

# Define the ExchangeLabs pattern to match
$pattern = "/o=ExchangeLabs/ou=Exchange Administrative Group \(FYDIBOHF23SPDLT\)/cn=Recipients/cn="

# Email fields to check
$EmailFields = @("Email1Address", "Email2Address", "Email3Address")

foreach ($Contact in $ContactsFolder.Items) {
    try {
        $emailCleared = $false
        foreach ($field in $EmailFields) {
            $EmailAddress = $Contact.$field
            if ($EmailAddress -match $pattern) {
                if ($VerboseOutput) {
                    Write-Verbose "Removing ExchangeLabs email from: $($Contact.FullName) ($($field): $EmailAddress)"
                } else {
                    Write-Host "Removing ExchangeLabs email from: $($Contact.FullName) ($($field): $EmailAddress)"
                }
                if (-not $DryRun) {
                    $Contact.$field = ""
                }
                $updatedCount++
                $emailCleared = $true
                Write-Host "Updated: '$($Contact.FullName)' → $field cleared" -ForegroundColor Green
            }
        }
        if ($emailCleared -and -not $DryRun) {
            $Contact.Save()
        }
        if (-not $emailCleared) {
            Write-Host "Skipping: '$($Contact.FullName)' (No ExchangeLabs email found)" -ForegroundColor Yellow
        }
    } catch {
        $errorCount++
        Write-Host "Error updating contact: $_" -ForegroundColor Red
    }
}

Write-Host "`nProcess complete: $updatedCount email fields updated, $errorCount errors encountered."

# Clean up COM objects
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($ContactsFolder) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Namespace) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Outlook) | Out-Null