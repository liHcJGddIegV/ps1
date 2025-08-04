<#
.SYNOPSIS
Updates the Outlook contacts' "Company" field based on email domain mappings.
If a contact’s email domain isn’t found in the mapping, it updates the Company field to "To Review" and
adds a new mapping entry with that domain, persisting changes to a JSON file for future runs.
Additionally, if a contact has the business address:
    1 S. Wacker Drive Suite 1800
    Chicago, Illinois  60606
the "Company" field is updated to "Invenergy LLC".

.PARAMETER DryRun
Simulates changes without actually saving to Outlook.

.PARAMETER VerboseOutput
Provides detailed logs for each update attempt.
#>

param (
    [switch]$DryRun,
    [switch]$VerboseOutput
)

# Define file paths for logging and mapping persistence
$LogFilePath = "C:\Users\YGonzalez\OneDrive - Invenergy LLC\Desktop\Logs\Update-CompanyContactsRev1.txt"
$MappingFilePath = "C:\Users\YGonzalez\OneDrive - Invenergy LLC\Desktop\Logs\CompanyMapping.json"

# Start transcript logging
Start-Transcript -Path $LogFilePath -Append

# ---------------------------------------------
# Load or Initialize Email Domain-to-Company Mapping from JSON
# ---------------------------------------------
if (Test-Path $MappingFilePath) {
    try {
        $json = Get-Content $MappingFilePath -Raw
        $mappingObject = ConvertFrom-Json $json
        # Convert PSCustomObject to hashtable
        $CompanyMapping = @{}
        foreach ($prop in $mappingObject.PSObject.Properties) {
            $CompanyMapping[$prop.Name.ToLower()] = $prop.Value
        }
        if ($VerboseOutput) {
            Write-Host "Loaded existing mapping from $MappingFilePath"
        }
    }
    catch {
        Write-Host "Error loading mapping file. Initializing default mapping." -ForegroundColor Yellow
        $CompanyMapping = $null
    }
}

if (-not $CompanyMapping) {
    # Default mapping if file does not exist or loading failed.
    $CompanyMapping = @{
        "invenergy.com"        = "Invenergy LLC"
        "aep.com"              = "AEP"
        "detect-inc.com"       = "Detect, Inc"
        "eepowersolutions.com" = "Eagle Eye Power Solutions, LLC"
        "eciusa.com"           = "Electrical Consultants, Inc."
        "gevernova.com"        = "GE Vernova"
        "ge.com"               = "GE Renewable Energy"
        "morteson.com"         = "Mortenson"
        "neieng.com"           = "NEI"
        "ulteig.com"           = "Ulteig Engineers, Inc."
        "vikor.com"            = "Vikor"
        "burnsmcd.com"         = "Burns & McDonnell"
        "emerson.com"          = "Emerson Electric Co"
    }
    if ($VerboseOutput) {
        Write-Host "Initialized default mapping."
    }
}

# ---------------------------------------------
# Initialize Outlook and get the Contacts folder
# ---------------------------------------------
try {
    $Outlook = New-Object -ComObject Outlook.Application
    $Namespace = $Outlook.GetNamespace("MAPI")
    $ContactsFolder = $Namespace.GetDefaultFolder(10)  # 10 = olFolderContacts
}
catch {
    Write-Host "Error: Outlook is not installed or could not be accessed." -ForegroundColor Red
    Stop-Transcript
    exit 1
}

# ---------------------------------------------
# Counters
# ---------------------------------------------
$updatedCount = 0
$errorCount   = 0

Write-Host "Starting Outlook contacts update process..."

# ---------------------------------------------
# Iterate over all items in the Contacts folder
# ---------------------------------------------
foreach ($Item in $ContactsFolder.Items) {
    try {
        # Ensure the item is a contact (40 = olContact)
        if ($Item.Class -ne 40) {
            if ($VerboseOutput) {
                Write-Host "Skipping non-contact item: $($Item.Name)"
            }
            continue
        }

        $Contact = $Item

        # --------------------------------------------------
        # New Check: Update based on Business Address
        # --------------------------------------------------
        # Adjust the string comparison or use a regex if the address format may vary.
        $expectedAddress = "1 S. Wacker Drive Suite 1800`nChicago, Illinois  60606"
        if ($Contact.BusinessAddress -and $Contact.BusinessAddress -eq $expectedAddress) {
            if ($VerboseOutput) {
                Write-Host "Business Address matched for $($Contact.FullName). Setting Company to 'Invenergy LLC'."
            }
            if (-not $DryRun) {
                $Contact.CompanyName = "Invenergy LLC"
                $Contact.Save()
            }
            $updatedCount++
            Write-Host "✅ Updated via Business Address: '$($Contact.FullName)' → Company set to 'Invenergy LLC'" -ForegroundColor Green
            # Skip further processing for this contact
            continue
        }

        # --------------------------------------------------
        # Process based on Email Domains
        # --------------------------------------------------
        $EmailFields = @("Email1Address", "Email2Address", "Email3Address")
        $companyUpdated = $false

        foreach ($field in $EmailFields) {
            $EmailAddress = $Contact.$field

            if (![string]::IsNullOrWhiteSpace($EmailAddress)) {
                # Extract the domain (the part after "@")
                if ($EmailAddress -match "@(.+)$") {
                    $domain = $Matches[1].ToLower()
                    if ($VerboseOutput) {
                        Write-Host "Extracted domain '$domain' from email '$EmailAddress' for $($Contact.FullName)"
                    }

                    # Check if the extracted domain exists in the mapping
                    if ($CompanyMapping.ContainsKey($domain)) {
                        $NewCompany = $CompanyMapping[$domain]
                        if ($VerboseOutput) {
                            Write-Host "Updating company for: $($Contact.FullName) (`$field: $EmailAddress)"
                        }
                        if (-not $DryRun) {
                            $Contact.CompanyName = $NewCompany
                            $Contact.Save()
                        }
                        $updatedCount++
                        $companyUpdated = $true
                        Write-Host "✅ Updated: '$($Contact.FullName)' → Company set to '$NewCompany'" -ForegroundColor Green
                        break  # Found a match; stop checking further email fields
                    }
                    else {
                        # No mapping exists: add new entry with "To Review"
                        if ($VerboseOutput) {
                            Write-Host "No mapping found for domain '$domain'. Adding new mapping with value 'To Review'" -ForegroundColor Cyan
                        }
                        $CompanyMapping[$domain] = "To Review"
                        if ($VerboseOutput) {
                            Write-Host "Setting company for: $($Contact.FullName) (`$field: $EmailAddress) to 'To Review'"
                        }
                        if (-not $DryRun) {
                            $Contact.CompanyName = "To Review"
                            $Contact.Save()
                        }
                        $updatedCount++
                        $companyUpdated = $true
                        Write-Host "✅ Updated: '$($Contact.FullName)' → Company set to 'To Review'" -ForegroundColor Green
                        break
                    }
                }
                else {
                    if ($VerboseOutput) {
                        Write-Host "Could not extract domain from email address '$EmailAddress' for contact '$($Contact.FullName)'" -ForegroundColor Yellow
                    }
                }
            }
        }

        if (-not $companyUpdated) {
            if ($VerboseOutput) {
                Write-Host "No valid email domain found for '$($Contact.FullName)'. Skipping company update." -ForegroundColor Yellow
            }
        }
    }
    catch {
        $errorCount++
        Write-Host "❌ Error updating contact: $_" -ForegroundColor Red
    }
}

Write-Host "`n🎯 Process complete: $updatedCount contacts updated, $errorCount errors encountered."

# Optionally, list the final mapping keys (for debugging)
if ($VerboseOutput) {
    Write-Host "`nFinal Domain Mappings:" -ForegroundColor Magenta
    foreach ($key in $CompanyMapping.Keys | Sort-Object) {
        Write-Host "$key  =>  $($CompanyMapping[$key])"
    }
}

# ---------------------------------------------
# Persist the updated mapping to the JSON file for future runs
# ---------------------------------------------
try {
    $CompanyMapping | ConvertTo-Json -Depth 5 | Out-File -FilePath $MappingFilePath -Encoding UTF8
    if ($VerboseOutput) {
        Write-Host "Persisted updated mapping to $MappingFilePath" -ForegroundColor Green
    }
}
catch {
    Write-Host "❌ Error saving mapping to file: $_" -ForegroundColor Red
}

Stop-Transcript
