<# 
Example Usage:
.\contacts\Update-GEEmailContacts.ps1 -TargetDomain "invenergy.com" -Verbose
#>
[CmdletBinding()]
param(
    [switch]$DryRun,
    [string]$TargetDomain,
    [string]$LogFile = "C:\Users\YGonzalez\OneDrive - Invenergy LLC\Desktop\Logs\Update-GEEmailContacts.txt"
)

# Prompt for target domain if not provided.
if (-not $TargetDomain) {
    Write-Host "Please input the target domain in the following format: 'ge.com' (for example, 'ge.com')"
    $TargetDomain = Read-Host "Enter target domain"
}

# Start transcript for comprehensive logging.
Start-Transcript -Path $LogFile -Append

# Load .NET assembly for title-case conversion.
Add-Type -AssemblyName System.Globalization
$textInfo = [System.Globalization.CultureInfo]::CurrentCulture.TextInfo

# Helper function: Extract first and last names from an email address.
function Get-NamesFromEmail {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Email,
        [Parameter(Mandatory = $true)]
        [System.Globalization.TextInfo]$TextInfo
    )
    $username = $Email.Split('@')[0]
    if ($username.Contains('.')) {
        $parts = $username.Split('.')
        if ($parts.Count -ge 2) {
            return @{
                FirstName = $TextInfo.ToTitleCase($parts[0].ToLower())
                LastName  = $TextInfo.ToTitleCase($parts[1].ToLower())
            }
        }
    }
    return $null
}

try {
    Write-Verbose "Creating Outlook COM object..."
    # Create Outlook COM object and access the default Contacts folder.
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNameSpace("MAPI")
    $contactsFolder = $namespace.GetDefaultFolder(10)  # 10 = olFolderContacts
}
catch {
    Write-Error "Failed to access Outlook contacts. Error: $_"
    Stop-Transcript
    exit
}

Write-Verbose "Processing contacts from Outlook..."
# Process each contact.
foreach ($contact in $contactsFolder.Items) {
    try {
        $email = $contact.Email1Address

        # Ensure the email is valid.
        if (-not [string]::IsNullOrEmpty($email) -and $email.Contains('@')) {
            $domain = $email.Split('@')[1]
            if ($domain -ieq $TargetDomain) {
                Write-Verbose "Processing contact with email: $email"
                $names = Get-NamesFromEmail -Email $email -TextInfo $textInfo
                if ($names) {
                    # Update contact names.
                    $contact.FirstName = $names.FirstName
                    $contact.LastName  = $names.LastName
                    $contact.FullName  = "$($names.FirstName) $($names.LastName)"
                    Write-Verbose "Extracted names - First Name: $($names.FirstName), Last Name: $($names.LastName)"

                    # Save changes unless in DryRun mode.
                    if ($DryRun) {
                        Write-Output "DryRun: Would update contact: $($contact.FullName) (Email: $email)"
                    }
                    else {
                        $contact.Save()
                        Write-Output "Updated contact: $($contact.FullName) (Email: $email)"
                    }
                }
                else {
                    Write-Verbose "Could not extract names from email: $email"
                }
            }
            else {
                Write-Verbose "Skipping email '$email' because it is not in the '$TargetDomain' domain."
            }
        }
        else {
            Write-Verbose "Skipping contact with no valid email."
        }
    }
    catch {
        Write-Error "Error processing contact with email '$email'. Error: $_"
    }
}

# Ensure transcript is stopped even if an error occurs.
Stop-Transcript
