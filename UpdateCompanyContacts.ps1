# PowerShell script to update Outlook contacts' company field based on email domain
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

# Define email domain-to-company mapping
$CompanyMapping = @{
    "@invenergy.com" = "Invenergy LLC"
    "@aep.com" = "AEP"
    "@detect-inc.com" = "Detect, Inc"
    "@eepowersolutions.com" = "Eagle Eye Power Solutions, LLC"
    "@eciusa.com" = "Electrical Consultants, Inc."
    "@ge.com" = "GE Vernova"
    "@mortenson.com" = "Mortenson"
    "@neieng.com" = "NEI"
    "@ulteig.com" = "Ulteig Engineers, Inc."
    "@vikor.com" = "Vikor"
    "@renewablepower.org" = "Renewable Power Group"
}

# Counters
$updatedCount = 0
$errorCount   = 0

Write-Host "Starting Outlook contacts update process..."

foreach ($Contact in $ContactsFolder.Items) {
    try {
        $EmailFields = @("Email1Address", "Email2Address", "Email3Address")
        $companyUpdated = $false

        foreach ($field in $EmailFields) {
            $EmailAddress = $Contact.$field

            # Ensure the email is not null or empty
            if (![string]::IsNullOrWhiteSpace($EmailAddress)) {
                foreach ($domain in $CompanyMapping.Keys) {
                    # Proper regex matching for domains
                    if ($EmailAddress -match ($domain -replace "\.", "\.")) {
                        $NewCompany = $CompanyMapping[$domain]

                        if ($VerboseOutput) {
                            Write-Host "Updating company for: $($Contact.FullName) ( $($field) : $EmailAddress )"
                        }

                        # Update the company field
                        if (-not $DryRun) {
                            $Contact.CompanyName = $NewCompany
                            $Contact.Save()
                        }

                        $updatedCount++
                        $companyUpdated = $true
                        Write-Host "✅ Updated: '$($Contact.FullName)' → Company set to '$NewCompany'" -ForegroundColor Green
                        break  # Stop checking once a match is found
                    }
                }
            }
        }

        if (-not $companyUpdated) {
            Write-Host "⚠ Skipping: '$($Contact.FullName)' (No matching email domain found)" -ForegroundColor Yellow
        }
    } catch {
        $errorCount++
        Write-Host "❌ Error updating contact: $_" -ForegroundColor Red
    }
}

Write-Host "`n🎯 Process complete: $updatedCount contacts updated, $errorCount errors encountered."
