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
    "@examplecorp.com" = "Example Corporation"
    "@techsolutions.io" = "Tech Solutions Inc."
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

            foreach ($domain in $CompanyMapping.Keys) {
                if ($EmailAddress -match $domain) {
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
                    Write-Host "Updated: '$($Contact.FullName)' → Company set to '$NewCompany'" -ForegroundColor Green
                    break  # Stop checking once a match is found
                }
            }
        }

        if (-not $companyUpdated) {
            Write-Host "Skipping: '$($Contact.FullName)' (No matching email domain found)" -ForegroundColor Yellow
        }
    } catch {
        $errorCount++
        Write-Host "Error updating contact: $_" -ForegroundColor Red
    }
}

Write-Host "`nProcess complete: $updatedCount contacts updated, $errorCount errors encountered."
