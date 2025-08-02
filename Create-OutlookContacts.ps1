<#
.SYNOPSIS
    Scans the Outlook Inbox for all emails, extracts potential contact entries from anywhere in the email text,
    cleans out any extraneous punctuation surrounding names and company names, and creates new contacts in Outlook if they do not already exist.

.DESCRIPTION
    The script connects to Outlook via COM and retrieves all emails from the Inbox.
    It then builds a combined text string from several fields (SenderName, SenderEmailAddress, To, CC, Subject, and Body)
    and uses a regular expression to extract “raw” names and an optional company.
    Regardless of which delimiter characters (parentheses, brackets, braces, quotes, asterisks, etc.) are used,
    the script “cleans” them so that the final contact always uses the format:
        FullName: [FirstName LastName] and Company: [Company Name].
    If the raw name contains a comma (e.g. "Liu, Ellen"), it is reordered to become "Ellen Liu".
    In‑memory deduplication (by email if available or by a combination of name and company) prevents duplicate entries.
    The script supports a DryRun mode and logs detailed information to a transcript file.

.PARAMETER DryRun
    If specified, the script will simulate contact creation without actually saving any new contacts.

.PARAMETER LogFile
    Path to the log file. Default is "C:\Users\YGonzalez\OneDrive - Invenergy LLC\Desktop\OutlookContactCreation.log".
    Ensure that the folder exists.
#>

[CmdletBinding()]
param (
    [switch]$DryRun,
    [string]$LogFile = "C:\Users\YGonzalez\OneDrive - Invenergy LLC\Desktop\OutlookContactCreation.log"
)

# =====================================
# Setup Logging: Start Transcript & Logger
# =====================================
try {
    Start-Transcript -Path $LogFile -Append -ErrorAction Stop
} catch {
    Write-Warning "Could not start transcript logging: $_"
}

function Write-Log {
    param (
        [string]$Message,
        [ValidateSet("INFO", "VERBOSE", "ERROR")]
        [string]$Level = "INFO"
    )
    $timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    $line = "$timestamp [$Level] - $Message"
    Write-Host $line
    Write-Output $line
}

# =====================================
# Helper: Format-CompanyName
# Removes any unwanted punctuation characters from the beginning and end of the company string.
# =====================================
function Format-CompanyName {
    param (
        [string]$Company
    )
    if ([string]::IsNullOrWhiteSpace($Company)) { 
        return $null 
    }
    # Remove common surrounding characters: parentheses, brackets, braces, quotes, asterisks, etc.
    $cleaned = $Company.Trim() -replace '^[\(\[\{\>"\*\s]+','' -replace '[\)\]\}\>"\*\s]+$',''
    return $cleaned
}

# =====================================
# Function: Get-OutlookFolders
# =====================================
function Get-OutlookFolders {
    try {
        $outlook = New-Object -ComObject Outlook.Application
        $namespace = $outlook.GetNameSpace("MAPI")
        $inbox = $namespace.GetDefaultFolder(6)          # 6 = olFolderInbox
        $contactsFolder = $namespace.GetDefaultFolder(10)  # 10 = olFolderContacts
        Write-Log "Connected to Outlook and retrieved Inbox and Contacts folder." "INFO"
        return [PSCustomObject]@{
            Outlook         = $outlook
            Namespace       = $namespace
            Inbox           = $inbox
            ContactsFolder  = $contactsFolder
        }
    }
    catch {
        Write-Log "Error connecting to Outlook: $_" "ERROR"
        throw $_
    }
}

# =====================================
# Function: ContactExists
# Checks whether a contact with the given email exists.
# =====================================
function ContactExists {
    param (
        [object]$ContactsFolder,
        [string]$Email
    )
    try {
        $existing = $ContactsFolder.Items.Find("[Email1Address] = '$Email'")
        return $null -ne $existing
    }
    catch {
        Write-Log "Error checking for existing contact ($Email): $_" "ERROR"
        return $false
    }
}

# =====================================
# Function: Invoke-ContactStringProcessing
#
# This function scans a text string for contact-like entries.
# It uses a regular expression to capture:
#   - rawName: A sequence of letters, digits, underscores, whitespace, commas, periods, apostrophes, or hyphens.
#   - Optionally, company: text enclosed in common delimiters (excluding angle brackets).
#
# After extracting the groups, the script:
#   1. Normalizes the text (removes newlines).
#   2. Trims and cleans the company name.
#   3. If the rawName contains a comma (e.g. "Liu, Ellen"), it splits the name and reorders it to "Ellen Liu".
#   4. Uses the email (if found) or a combination of display name and company for deduplication.
#
# Deduplication prevents duplicate contact creation.
# =====================================
function Invoke-ContactStringProcessing {
    param (
        [string]$ContactString,
        [object]$Outlook,
        [object]$ContactsFolder,
        [ref]$ProcessedContacts,  # Hashtable for deduplication keys.
        [ref]$CreatedCount,
        [ref]$SkippedCount
    )
    
    if ([string]::IsNullOrWhiteSpace($ContactString)) { 
        return 
    }
    
    # Normalize the contact string by replacing newline characters with a space.
    $ContactString = $ContactString -replace "\r?\n", " "

    # Define a regex pattern using a here-string.
    # Updated to remove angle brackets from delimiters for the optional company.
    $contactRegex = [regex]@"
(?<rawName>[-\w\s,.'"]+)(?:[\(\[\{\*\s]+(?<company>[-\w\s&]+)[\)\]\}\*\s]+)?
"@
    # Get matches (using a non-automatic variable to avoid interfering with $matches)
    $contactMatches = $contactRegex.Matches($ContactString)
    
    foreach ($match in $contactMatches) {
        if (-not $match.Success) { 
            continue 
        }
        $rawName = $match.Groups["rawName"].Value.Trim()
        if ([string]::IsNullOrWhiteSpace($rawName)) { 
            continue 
        }
        $company = $null
        if ($match.Groups["company"].Success -and -not [string]::IsNullOrWhiteSpace($match.Groups["company"].Value)) {
            $company = Format-CompanyName $match.Groups["company"].Value
        }
        $email = ""
        # Look for an email pattern anywhere in the contact string.
        $emailRegex = [regex]'[\w\.\-+]+@[\w\.\-]+\.[\w]{2,}'
        $emailMatch = $emailRegex.Match($ContactString)
        if ($emailMatch.Success) {
            $email = $emailMatch.Value.Trim()
        }
        
        # Process rawName: if it contains a comma, assume "Last, First" and swap.
        if ($rawName -match ",") {
            $parts = $rawName.Split(",") | ForEach-Object { $_.Trim() }
            if ($parts.Count -ge 2) {
                $displayName = "$($parts[1]) $($parts[0])"
            }
            else {
                $displayName = $rawName
            }
        }
        else {
            $displayName = $rawName
        }
        
        # Build deduplication key.
        if ($email -ne "") {
            $dedupKey = $email.ToLower()
        }
        else {
            $dedupKey = ("{0}|{1}" -f $displayName.ToLower(), ($company | ForEach-Object { $_.ToLower() }))
        }
        
        if ($ProcessedContacts.Value.ContainsKey($dedupKey)) {
            Write-Log "Skipping $($dedupKey): already processed in this run." "VERBOSE"
            continue
        }
        
        # Check if contact already exists in Outlook.
        if ($email -ne "") {
            if (ContactExists -ContactsFolder $ContactsFolder -Email $email) {
                Write-Log "Skipping $($email): contact already exists in Outlook." "INFO"
                $ProcessedContacts.Value[$dedupKey] = $true
                $SkippedCount.Value++
                continue
            }
        }
        else {
            try {
                $existing = $ContactsFolder.Items.Find("[FullName] = '$displayName'")
            }
            catch {
                $existing = $null
            }
            if ($null -ne $existing) {
                Write-Log "Skipping $($displayName): contact already exists in Outlook (by FullName)." "INFO"
                $ProcessedContacts.Value[$dedupKey] = $true
                $SkippedCount.Value++
                continue
            }
        }
        
        Write-Log "Parsed contact: DisplayName='$displayName'; Company='$company'; Email='$email'" "VERBOSE"
        
        if ($DryRun) {
            Write-Log "[DRY RUN] Would create contact: $displayName, Company: $company, Email: $email" "INFO"
        }
        else {
            try {
                $contact = $Outlook.CreateItem(2)   # 2 = olContactItem
                $contact.FullName = $displayName
                if ($company) {
                    $contact.CompanyName = $company
                }
                if ($email -ne "") {
                    $contact.Email1Address = $email
                }
                $contact.Save()
                Write-Log "Created contact: $displayName; Company: $company; Email: $email" "INFO"
                $CreatedCount.Value++
            }
            catch {
                Write-Log "Failed to create contact for $displayName ($email): $_" "ERROR"
            }
        }
        $ProcessedContacts.Value[$dedupKey] = $true
    }
}  # End of Invoke-ContactStringProcessing

# =====================================
# Main Script Execution
# =====================================

# Initialize counters.
$overallCreated = [ref]0
$overallSkipped = [ref]0
$overallEmailsProcessed = 0

# Hashtable to track processed contacts (for deduplication).
$ProcessedContacts = [ref]@{}

# Retrieve Outlook folders.
$folders = Get-OutlookFolders
$outlook = $folders.Outlook
$inbox = $folders.Inbox
$contactsFolder = $folders.ContactsFolder

Write-Log "Processing all emails in the Inbox folder." "INFO"

# Retrieve all items in the Inbox.
$allItems = $inbox.Items

foreach ($item in $allItems) {
    try {
        # Process only mail items.
        if ($null -ne $item -and $item.MessageClass -eq "IPM.Note") {
            $overallEmailsProcessed++
            # Build a combined text from several fields.
            $combinedText = ""
            if ($item.PSObject.Properties.Match("SenderName")) { 
                $combinedText += " " + $item.SenderName 
            }
            if ($item.PSObject.Properties.Match("SenderEmailAddress")) { 
                $combinedText += " " + $item.SenderEmailAddress 
            }
            if ($item.PSObject.Properties.Match("To")) { 
                $combinedText += " " + $item.To 
            }
            if ($item.PSObject.Properties.Match("CC")) { 
                $combinedText += " " + $item.CC 
            }
            if ($item.PSObject.Properties.Match("Subject")) { 
                $combinedText += " " + $item.Subject 
            }
            if ($item.PSObject.Properties.Match("Body")) { 
                $combinedText += " " + $item.Body 
            }
            
            Write-Log "Processing email: Subject='$($item.Subject)', ReceivedTime='$($item.ReceivedTime)'" "VERBOSE"
            Invoke-ContactStringProcessing -ContactString $combinedText -Outlook $outlook -ContactsFolder $contactsFolder -ProcessedContacts $ProcessedContacts -CreatedCount $overallCreated -SkippedCount $overallSkipped
        }
    }
    catch {
        Write-Log "Error processing an email item: $_" "ERROR"
    }
}

# =====================================
# Summary Report
# =====================================
Write-Log "Processing complete." "INFO"
Write-Log "Emails processed: $overallEmailsProcessed" "INFO"
Write-Log "New contacts created: $($overallCreated.Value)" "INFO"
Write-Log "Contacts skipped (already existed): $($overallSkipped.Value)" "INFO"

# =====================================
# Cleanup: Release COM objects
# =====================================
try {
    if ($null -ne $folders) {
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($folders.Namespace) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($folders.Inbox) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($folders.ContactsFolder) | Out-Null
    }
    if ($null -ne $outlook) {
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook) | Out-Null
    }
}
catch {
    Write-Log "Error releasing COM objects: $_" "ERROR"
}
finally {
    try {
        Stop-Transcript | Out-Null
    }
    catch {
        Write-Warning "Error stopping transcript: $_"
    }
}