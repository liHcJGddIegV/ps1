param (
    [switch]$DryRun,
    [switch]$VerboseOutput
)

# Set console output encoding to UTF-8 to properly display special characters
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::UTF8

# File paths for logging and mapping persistence
$LogFilePath = "C:\Users\YGonzalez\OneDrive - Invenergy LLC\Desktop\Logs\Update-CompanyContacts.txt"
$MappingFilePath = "C:\Users\YGonzalez\OneDrive - Invenergy LLC\Desktop\Logs\CompanyMapping.json"

Start-Transcript -Path $LogFilePath -Append

# Load or initialize the email domain-to-company mapping from JSON
if (Test-Path $MappingFilePath) {
    try {
        $json = Get-Content $MappingFilePath -Raw
        $mappingObject = ConvertFrom-Json $json
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

# Initialize Outlook and get the Contacts folder
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

$updatedCount = 0
$errorCount   = 0

Write-Host "Starting Outlook contacts update process..."

foreach ($Item in $ContactsFolder.Items) {
    try {
        if ($Item.Class -ne 40) {
            if ($VerboseOutput) {
                Write-Host "Skipping non-contact item: $($Item.Name)"
            }
            continue
        }

        $Contact = $Item

        # Update based on Business Address using a regex pattern
        $expectedRegex = '1\s+S\.\s+Wacker\s+Drive\s+Suite\s+1800\s*[\r\n]+\s*Chicago,\s+Illinois\s+60606'
        if ($Contact.BusinessAddress -and $Contact.BusinessAddress -imatch $expectedRegex) {
            if ($VerboseOutput) {
                Write-Host "Business Address matched for $($Contact.FullName). Setting Company to 'Invenergy LLC'."
            }
            if (-not $DryRun) {
                $Contact.CompanyName = "Invenergy LLC"
                $Contact.Save()
            }
            $updatedCount++
            Write-Host "[OK] Updated via Business Address: '$($Contact.FullName)' -> Company set to 'Invenergy LLC'" -ForegroundColor Green
            continue
        }

        # Process based on Email Domains
        $EmailFields = @("Email1Address", "Email2Address", "Email3Address")
        $companyUpdated = $false

        foreach ($field in $EmailFields) {
            $EmailAddress = $Contact.$field
            if (![string]::IsNullOrWhiteSpace($EmailAddress)) {
                if ($EmailAddress -match "@(.+)$") {
                    $domain = $Matches[1].ToLower()
                    if ($VerboseOutput) {
                        Write-Host "Extracted domain '$domain' from email '$EmailAddress' for $($Contact.FullName)"
                    }
                    if ($CompanyMapping.ContainsKey($domain)) {
                        $NewCompany = $CompanyMapping[$domain]
                        if ($VerboseOutput) {
                            Write-Host "Updating company for: $($Contact.FullName) (Field: $field, Email: $EmailAddress)"
                        }
                        if (-not $DryRun) {
                            $Contact.CompanyName = $NewCompany
                            $Contact.Save()
                        }
                        $updatedCount++
                        $companyUpdated = $true
                        Write-Host "[OK] Updated: '$($Contact.FullName)' -> Company set to '$NewCompany'" -ForegroundColor Green
                        break
                    }
                    else {
                        if ($VerboseOutput) {
                            Write-Host "No mapping found for domain '$domain'. Deriving company name from domain..." -ForegroundColor Cyan
                        }
                        # Remove the TLD (e.g., .com) from the domain
                        $baseName = $domain -replace '\.[^.]+$', ''
                        # Insert a space before "company" (if applicable)
                        $prettyName = $baseName -replace '(?<=\w)(company)$', ' Company'
                        # Convert to Title Case (e.g., "blattnercompany" -> "Blattner Company")
                        $prettyName = [System.Globalization.CultureInfo]::CurrentCulture.TextInfo.ToTitleCase($prettyName)
                        
                        # Add the new mapping to the hashtable
                        $CompanyMapping[$domain] = $prettyName

                        if ($VerboseOutput) {
                            Write-Host "Setting company for: $($Contact.FullName) (Field: $field, Email: $EmailAddress) to '$prettyName'"
                        }
                        if (-not $DryRun) {
                            $Contact.CompanyName = $prettyName
                            $Contact.Save()
                        }
                        $updatedCount++
                        $companyUpdated = $true
                        Write-Host "[OK] Updated: '$($Contact.FullName)' -> Company set to '$prettyName'" -ForegroundColor Green
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
        Write-Host "[ERROR] Error updating contact: $_" -ForegroundColor Red
    }
}

Write-Host "`n[OK] Process complete: $updatedCount contacts updated, $errorCount errors encountered."

if ($VerboseOutput) {
    Write-Host "`nFinal Domain Mappings:" -ForegroundColor Magenta
    foreach ($key in $CompanyMapping.Keys | Sort-Object) {
        Write-Host "$key  =>  $($CompanyMapping[$key])"
    }
}

try {
    $CompanyMapping | ConvertTo-Json -Depth 5 | Out-File -FilePath $MappingFilePath -Encoding UTF8
    if ($VerboseOutput) {
        Write-Host "Persisted updated mapping to $MappingFilePath" -ForegroundColor Green
    }
}
catch {
    Write-Host "[ERROR] Error saving mapping to file: $_" -ForegroundColor Red
}

Stop-Transcript
