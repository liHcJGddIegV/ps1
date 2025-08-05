param (
    [switch]$DryRun,          # Simulate changes without moving/deleting
    [switch]$VerboseOutput    # Provide detailed logs
)

# Define file paths for logging
$LogFilePath = "C:\Users\YGonzalez\OneDrive - Invenergy LLC\Desktop\Logs\Remove-DuplicateContactsl.log"

# Function to log messages with proper UTF-8 encoding (PowerShell 5.1 Compatible)
function Write-Log {
    param (
        [string]$Message,
        [string]$Color = "White"
    )

    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogMessage = "[$Timestamp] $Message"

    # Write to console with color
    Write-Host $LogMessage -ForegroundColor $Color

    # Ensure UTF-8 encoding without BOM using [System.Text.Encoding]::UTF8
    $Utf8Encoding = New-Object System.Text.UTF8Encoding($False)
    [System.IO.File]::AppendAllText($LogFilePath, "$LogMessage`r`n", $Utf8Encoding)
}

# Notify dry-run mode if applicable
if ($DryRun) {
    Write-Log "DryRun mode: No contacts will be moved or deleted." "Cyan"
}

# Initialize Outlook Application
try {
    $Outlook = New-Object -ComObject Outlook.Application
    $Namespace = $Outlook.GetNamespace("MAPI")
    $ContactsFolder = $Namespace.GetDefaultFolder(10)  # olFolderContacts = 10

    # Create or get the "Duplicates" folder
    $DuplicatesFolder = $ContactsFolder.Folders | Where-Object { $_.Name -eq "Duplicates" }
    if (-not $DuplicatesFolder) {
        $DuplicatesFolder = $ContactsFolder.Folders.Add("Duplicates")
    }
} catch {
    Write-Log "Error: Outlook is not installed or could not be accessed." "Red"
    exit 1
}

# Counters for updates and errors
$duplicatesFound = 0
$movedCount = 0
$deletedCount = 0
$errorCount = 0

Write-Log "Starting Outlook duplicate contacts cleanup..." "White"
Write-Log "Starting duplicate cleanup at $(Get-Date)" "White"

# Hashtable to track unique contacts
$uniqueContacts = @{}

foreach ($Contact in $ContactsFolder.Items) {
    try {
        # Use Full Name, Primary Email, Mobile Number, Home Phone, Business Phone, and Company Name for uniqueness
        $fullName = if ($Contact.FullName) { $Contact.FullName.Trim() } else { "Unknown" }
        $email = if ($Contact.Email1Address) { $Contact.Email1Address.Trim() } else { "Unknown" }
        $mobile = if ($Contact.MobileTelephoneNumber) { $Contact.MobileTelephoneNumber.Trim() } else { "Unknown" }
        $homePhone = if ($Contact.HomeTelephoneNumber) { $Contact.HomeTelephoneNumber.Trim() } else { "Unknown" }
        $businessPhone = if ($Contact.BusinessTelephoneNumber) { $Contact.BusinessTelephoneNumber.Trim() } else { "Unknown" }
        $company = if ($Contact.CompanyName) { $Contact.CompanyName.Trim() } else { "Unknown" }

        # Create a unique key (ignoring case)
        $contactKey = "$fullName|$email|$mobile|$homePhone|$businessPhone|$company".ToLower()

        if ($uniqueContacts.ContainsKey($contactKey)) {
            # Duplicate found
            $msg = "Duplicate found: '$fullName' (Email: $email, Mobile: $mobile, Company: $company)"
            Write-Log $msg "Yellow"
            $duplicatesFound++

            if (-not $DryRun) {
                try {
                    # Move duplicate to "Duplicates" folder
                    $Contact.Move($DuplicatesFolder)
                    $movedCount++
                    Write-Log "Moved to 'Duplicates': '$fullName'" "Green"
                } catch {
                    # If moving fails, prompt for deletion
                    $confirmation = [System.Windows.Forms.MessageBox]::Show(
                        "A duplicate contact was found: $fullName.`nDo you want to DELETE it?",
                        "Duplicate Contact Found",
                        [System.Windows.Forms.MessageBoxButtons]::YesNo,
                        [System.Windows.Forms.MessageBoxIcon]::Warning
                    )

                    if ($confirmation -eq "Yes") {
                        $Contact.Delete()
                        $deletedCount++
                        Write-Log "Deleted duplicate: '$fullName'" "Red"
                    } else {
                        Write-Log "Skipped deletion: '$fullName'" "Cyan"
                    }
                }
            }
        } else {
            # Store unique contact
            $uniqueContacts[$contactKey] = $true
        }
    } catch {
        $errorCount++
        $errorMessage = "Error processing contact: '$fullName' - $($_.Exception.Message)"
        Write-Log $errorMessage "Red"
    }
}

Write-Log "Process complete: $duplicatesFound duplicates found, $movedCount moved, $deletedCount deleted, $errorCount errors encountered." "White"

# Clean up COM objects
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($ContactsFolder) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Namespace) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Outlook) | Out-Null

# Display the log file location message
Write-Log "Log file saved at: $LogFilePath" "White"
