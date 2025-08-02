# PowerShell script to remove ExchangeLabs email addresses from Outlook contacts
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

# Define the ExchangeLabs pattern to match
$pattern = "/o=ExchangeLabs/ou=Exchange Administrative Group \(FYDIBOHF23SPDLT\)/cn=Recipients/cn="

foreach ($Contact in $ContactsFolder.Items) {
    try {
        $EmailFields = @("Email1Address", "Email2Address", "Email3Address")
        $emailCleared = $false
        
        foreach ($field in $EmailFields) {
            $EmailAddress = $Contact.$field
            
            if ($EmailAddress -match $pattern) {
                if ($VerboseOutput) {
                    Write-Host "Removing ExchangeLabs email from: $($Contact.FullName) ( $($field) : $EmailAddress )"
                }

                # Clear the email field
                if (-not $DryRun) {
                    $Contact.$field = ""
                    $Contact.Save()
                }

                $updatedCount++
                $emailCleared = $true
                Write-Host "Updated: '$($Contact.FullName)' → $field cleared" -ForegroundColor Green
            }
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
