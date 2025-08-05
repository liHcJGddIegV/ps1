param (
    [switch]$DryRun,          # Simulate changes without saving
    [switch]$VerboseOutput    # Provide detailed logs
)

# Notify dry-run mode if applicable
if ($DryRun) {
    Write-Host "DryRun mode: No contacts will be moved or deleted." -ForegroundColor Cyan
}

# Define log file and JSON file paths
$LogFilePath = "C:\Users\YGonzalez\OneDrive - Invenergy LLC\Desktop\Logs\Update-IMtoEmail1DisplayName.log"
$JsonFile = "C:\Users\YGonzalez\OneDrive - Invenergy LLC\Desktop\Logs\CompanyMapping.json"

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

# Load JSON file containing valid email domains
if (Test-Path $JsonFile) {
    $DomainData = Get-Content -Path $JsonFile | ConvertFrom-Json
    $AllowedDomains = $DomainData.PSObject.Properties.Name -join "|"
    $domainPattern = "@($AllowedDomains)"
    Write-Log "Loaded email domains from JSON: $AllowedDomains" "Cyan"
} else {
    Write-Log "Error: JSON file with domain data not found. Exiting script." "Red"
    exit 1
}

# Initialize Outlook Application
try {
    $Outlook = New-Object -ComObject Outlook.Application
    $Namespace = $Outlook.GetNamespace("MAPI")
    $ContactsFolder = $Namespace.GetDefaultFolder(10)  # olFolderContacts = 10
} catch {
    Write-Log "Error: Outlook is not installed or could not be accessed." "Red"
    exit 1
}

# Counters for updates and errors
$updatedCount = 0
$errorCount   = 0

Write-Log "Starting Outlook contacts update..."

foreach ($Contact in $ContactsFolder.Items) {
    try {
        # Ensure we are processing a ContactItem (skip other item types)
        if ($Contact -isnot [Microsoft.Office.Interop.Outlook.ContactItem]) {
            continue
        }

        $FullName = $Contact.FullName
        $FirstName = $Contact.FirstName
        $LastName = $Contact.LastName
        $EmailDisplayAs = $Contact.Email1DisplayName
        $IMAddress = $Contact.IMAddress
        $EmailField = "Email1Address"

        # Validate that Full Name, First Name, and Last Name exist
        if (-not [string]::IsNullOrWhiteSpace($FullName)) {

            # Ensure Full Name format is correct (convert "Last, First" to "First Last" if needed)
            if ($FullName -match "^([\w\s]+),\s([\w\s]+)") {
                $CorrectedFullName = "$($matches[2]) $($matches[1])".Trim()
            } else {
                $CorrectedFullName = $FullName.Trim()
            }

            # Extract first and last names correctly
            $NameParts = $CorrectedFullName -split "\s+"
            if ($NameParts.Count -gt 1) {
                $CorrectedFirstName = $NameParts[0]   # First word
                $CorrectedLastName = $NameParts[-1]   # Last word
            } else {
                $CorrectedFirstName = $CorrectedFullName
                $CorrectedLastName = ""
            }

            # Update First Name if needed
            if ($FirstName -ne $CorrectedFirstName) {
                if ($VerboseOutput) { Write-Log "Correcting First Name: $FirstName → $CorrectedFirstName" "Yellow" }
                if (-not $DryRun) { $Contact.FirstName = $CorrectedFirstName }
            }

            # Update Last Name if needed
            if ($LastName -ne $CorrectedLastName) {
                if ($VerboseOutput) { Write-Log "Correcting Last Name: $LastName → $CorrectedLastName" "Yellow" }
                if (-not $DryRun) { $Contact.LastName = $CorrectedLastName }
            }

            # Update Email Display As to match Full Name
            if ($EmailDisplayAs -ne $CorrectedFullName) {
                if ($VerboseOutput) { Write-Log "Updating Email Display As: $EmailDisplayAs → $CorrectedFullName" "Yellow" }
                if (-not $DryRun) { $Contact.Email1DisplayName = [string]$CorrectedFullName }
                Write-Log "Updated Email Display As for '$CorrectedFullName'" "Blue"
                $updatedCount++
            }
        }

        # Ensure Email1Address is correct based on JSON domains
        if ($IMAddress -match $domainPattern) {
            if ($VerboseOutput) { Write-Log "Updating Email1Address: $IMAddress" "Yellow" }
            if (-not $DryRun) { $Contact.$EmailField = $IMAddress }
            Write-Log "Updated Email1Address for '$CorrectedFullName' → $IMAddress" "Green"
            $updatedCount++
        }

        # Save changes if not in DryRun mode
        if (-not $DryRun) { $Contact.Save() }

    } catch {
        $errorCount++
        Write-Log "Error updating contact '$FullName': $_" "Red"
    }
}

Write-Log "`nProcess complete: $updatedCount updates made, $errorCount errors encountered." "Magenta"
Write-Log "Log file saved at: $LogFilePath" "Magenta"
