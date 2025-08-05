# PowerShell script to update the "Display As" field with "Full Name"
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

Write-Host "Starting Outlook contacts update..."

foreach ($Contact in $ContactsFolder.Items) {
    try {
        # Retrieve the Full Name
        $FullName = $Contact.FullName
        $CurrentDisplayAs = $Contact.FileAs  # "Display As" field in Outlook
        
        if (-not [string]::IsNullOrWhiteSpace($FullName)) {
            if ($VerboseOutput) {
                Write-Host "Processing: $FullName (Current Display As: $CurrentDisplayAs)"
            }

            # Overwrite "Display As" field with "Full Name"
            if (-not $DryRun) {
                $Contact.FileAs = $FullName  # Setting FileAs to FullName
                $Contact.Save()
            }

            $updatedCount++
            Write-Host "Updated: '$FullName' → 'Display As' set to '$FullName'" -ForegroundColor Green
        } else {
            Write-Host "Skipping contact: No Full Name found." -ForegroundColor Yellow
        }
    } catch {
        $errorCount++
        Write-Host "Error updating contact: $_" -ForegroundColor Red
    }
}

Write-Host "`nProcess complete: $updatedCount contacts updated, $errorCount errors encountered."
