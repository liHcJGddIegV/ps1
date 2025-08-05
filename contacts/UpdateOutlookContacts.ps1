<#
.SYNOPSIS
    Updates Outlook contacts by ensuring the 'FileAs' field (Display As) matches the contact's name.

.DESCRIPTION
    This script iterates over contacts in your default Outlook Contacts folder. For each contact, it computes
    a desired name using the following rules:
      - Use the contact’s FullName if available.
      - If FullName is empty, combine FirstName and LastName.
    If the computed name is non-empty and does not match the current FileAs (Display As) field,
    the script will update FileAs to the computed name.

.PARAMETER DryRun
    If specified, the script will simulate the changes without saving them.

.PARAMETER VerboseOutput
    If specified, the script will output additional details during execution.

.EXAMPLE
    .\contacts\UpdateOutlookContacts.ps1 -DryRun -VerboseOutput
#>

param (
    [switch]$DryRun,
    [switch]$VerboseOutput
)

# Counters for reporting
$updatedCount = 0
$skippedCount = 0
$errorCount   = 0

Write-Host "Starting Outlook contacts update script..." -ForegroundColor Cyan

# Create an instance of the Outlook application COM object
try {
    $outlook = New-Object -ComObject Outlook.Application
} catch {
    Write-Error "Could not create the Outlook Application object. Ensure Outlook is installed on this system."
    exit 1
}

# Get the MAPI namespace and then the default Contacts folder (Folder Type 10 corresponds to Contacts)
$namespace      = $outlook.GetNamespace("MAPI")
$contactsFolder = $namespace.GetDefaultFolder(10)

# Retrieve all items from the Contacts folder
$contacts   = $contactsFolder.Items
$totalCount = $contacts.Count
Write-Host "Found $totalCount items in the Contacts folder."

# Loop through each contact
for ($i = 1; $i -le $totalCount; $i++) {
    try {
        $item = $contacts.Item($i)
        
        # Process only if the item is a ContactItem (Class 40)
        if ($item.Class -eq 40) {

            # Compute the desired name
            $desiredName = $item.FullName

            if ([string]::IsNullOrWhiteSpace($desiredName)) {
                $first = $item.FirstName
                $last  = $item.LastName
                if (-not ([string]::IsNullOrWhiteSpace($first)) -or -not ([string]::IsNullOrWhiteSpace($last))) {
                    $desiredName = ("$first $last").Trim()
                }
            }
            
            # If we have a valid desired name, check if FileAs needs to be updated
            if (-not [string]::IsNullOrWhiteSpace($desiredName)) {
                if ($item.FileAs -ne $desiredName) {
                    if ($DryRun) {
                        Write-Host "DryRun: Would update contact: '$desiredName' (current FileAs: '$($item.FileAs)')" -ForegroundColor Yellow
                    }
                    else {
                        $item.FileAs = $desiredName
                        $item.Save()
                        Write-Host "Updated contact: '$desiredName' - FileAs changed from '$($item.FileAs)' to '$desiredName'" -ForegroundColor Green
                    }
                    $updatedCount++
                }
                elseif ($VerboseOutput) {
                    Write-Host "Contact at index $i already has matching FileAs; skipping." -ForegroundColor Gray
                    $skippedCount++
                }
                else {
                    $skippedCount++
                }
            }
            else {
                Write-Host "Skipped contact at index $i due to no available name information." -ForegroundColor DarkYellow
                $skippedCount++
            }
        }
        elseif ($VerboseOutput) {
            Write-Host "Item at index $i is not a ContactItem; skipping." -ForegroundColor Gray
        }
    } catch {
        Write-Warning ("Error processing item " + $i + ": " + $_.Exception.Message)
        $errorCount++
    }
}

# Display summary
Write-Host "`nContacts update complete." -ForegroundColor Cyan
Write-Host "Total items processed: $totalCount"
Write-Host "Contacts updated:      $updatedCount"
Write-Host "Contacts skipped:      $skippedCount"
Write-Host "Errors encountered:    $errorCount"
