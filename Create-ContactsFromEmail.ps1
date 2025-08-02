param (
    [switch]$DryRun,          # Simulate changes without saving
    [switch]$VerboseOutput    # Provide detailed logs
)

# Define the log file location
$logFile = "C:\Users\YGonzalez\OneDrive - Invenergy LLC\Desktop\Logs\CreateContactsLog.txt"

# Initialize counters
$contactsCreated = 0
$contactsSkipped = 0

# Helper function for logging messages with timestamp
function Write-Log {
    param(
        [string]$Message,
        [switch]$Verbose
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $entry = "$timestamp - $Message"
    Add-Content -Path $logFile -Value $entry
    if ($VerboseOutput -or (-not $Verbose)) {
        Write-Host $entry
    }
}

Write-Log "Starting contact extraction and creation from selected email(s)."

# Create the Outlook COM object and get the MAPI namespace
$olApp = New-Object -ComObject Outlook.Application
$ns = $olApp.GetNamespace("MAPI")
Write-Log "Outlook COM object created." -Verbose

# Get the currently selected items in Outlook
$selection = $olApp.ActiveExplorer().Selection
if ($selection.Count -lt 1) {
    Write-Log "Please select one or more emails in Outlook."
    exit
}

# Filter out only MailItem objects (this ensures that we process only valid email items)
$emailsToProcess = @()
foreach ($item in $selection) {
    if ($item.Class -eq 43) {
        $emailsToProcess += $item
    }
    else {
        Write-Log "Skipping non-mail item of type $($item.Class)" -Verbose
    }
}

if ($emailsToProcess.Count -lt 1) {
    Write-Log "No valid mail items selected. Please select one or more emails." 
    exit
}

# Build a hash table to hold unique contacts (keyed by lowercase email address)
$uniqueContacts = @{ }

# Helper function to add an AddressEntry to the hash if it has a valid email address
function Add-ContactFromAddressEntry {
    param (
        [Parameter(Mandatory=$true)]
        $addrEntry
    )
    if (-not $addrEntry) { return }
    
    $email = $addrEntry.Address
    if ($email -like "SMTP:*") {
        $email = $email.Substring(5)
    }
    
    if ([string]::IsNullOrWhiteSpace($email) -or -not ($email -match "@")) {
        Write-Log "Skipping address entry '$($addrEntry.Name)' because '$email' is not a valid email address." -Verbose
        return
    }
    
    $key = $email.ToLower()
    if (-not $uniqueContacts.ContainsKey($key)) {
        $uniqueContacts[$key] = $addrEntry
        Write-Log "Added contact: $($addrEntry.Name) with email $email" -Verbose
    }
}

# Loop over each selected email and extract contacts from sender, To, and CC fields
foreach ($mail in $emailsToProcess) {
    Write-Log "Processing email: $($mail.Subject)" -Verbose

    if ($mail.Sender) {
        Add-ContactFromAddressEntry -addrEntry $mail.Sender
    }
    
    if ($mail.Recipients.Count -gt 0) {
        foreach ($recipient in $mail.Recipients) {
            if ($recipient.AddressEntry) {
                Add-ContactFromAddressEntry -addrEntry $recipient.AddressEntry
            }
        }
    }
}

Write-Log "Total unique valid contacts found across all emails: $($uniqueContacts.Keys.Count)" -Verbose

# Get the default Contacts folder
$contactsFolder = $ns.GetDefaultFolder(10)
Write-Log "Retrieved default Contacts folder." -Verbose

# Helper function to check if a contact already exists by primary email
function ContactExists($emailAddress) {
    $found = $contactsFolder.Items | Where-Object {
        $_.Email1Address -and $_.Email1Address.ToLower() -eq $emailAddress.ToLower()
    }
    return ($null -ne $found)
}

# Process each unique contact from the collected contacts
foreach ($emailKey in $uniqueContacts.Keys) {
    $addrEntry = $uniqueContacts[$emailKey]
    $displayName = $addrEntry.Name
    Write-Log "Processing contact: ${displayName} ($emailKey)" -Verbose
    
    if (ContactExists($emailKey)) {
        Write-Log "A contact with email $emailKey already exists. Skipping." -Verbose
        $contactsSkipped++
        continue
    }
    
    # Initialize fields using the display name as the default
    $FullName = $displayName
    $LastName = ""
    $FirstName = ""
    $JobTitle = ""
    $Department = ""
    $Company = ""
    $BusinessAddress = ""
    $Mobile = ""
    $Email = $emailKey
    $EmailDisplayAs = $displayName
    
    if ($addrEntry.Type -eq "EX") {
        try {
            $exUser = $addrEntry.GetExchangeUser()
        }
        catch {
            Write-Log "Error obtaining ExchangeUser details for ${displayName}: $($_.Exception.Message)" -Verbose
            $exUser = $null
        }
        if ($exUser) {
            $FullName        = $exUser.Name
            $LastName        = $exUser.LastName
            $FirstName       = $exUser.FirstName
            $JobTitle        = $exUser.JobTitle
            $Department      = $exUser.Department
            $Company         = $exUser.CompanyName
            $BusinessAddress = $exUser.OfficeLocation
            $Mobile          = $exUser.MobileTelephoneNumber
            $Email           = $exUser.PrimarySmtpAddress
            $EmailDisplayAs  = $exUser.Name
            Write-Log "Extracted ExchangeUser details for ${displayName}" -Verbose
        }
        else {
            Write-Log "No ExchangeUser details available for ${displayName}" -Verbose
        }
    }
    else {
        Write-Log "Non-Exchange contact. Basic details will be used for ${displayName}" -Verbose
    }
    
    # --- Universal Cleanup and Splitting Logic ---
    # Remove any email address in angle brackets and substrings like ":(...)".
    $cleanName = $FullName -replace "<.*?>", "" -replace ":\(.*?\)", ""
    $cleanName = $cleanName.Trim()

    if ($cleanName -match ",") {
        # Format appears to be "Last, First ..." - split on the first comma.
        $parts = $cleanName.Split(",",2)
        $lastName = $parts[0].Trim()
        $firstPart = $parts[1].Trim()
        # Split first part into tokens and choose the candidate with more than one character.
        $tokens = $firstPart.Split(" ", [System.StringSplitOptions]::RemoveEmptyEntries)
        $firstNameCandidates = $tokens | Where-Object { $_.Length -gt 1 }
        if ($firstNameCandidates.Count -gt 0) {
            $firstName = ($firstNameCandidates | Sort-Object Length -Descending | Select-Object -First 1)
        }
        else {
            $firstName = $tokens[0]
        }
        $fullName = "$firstName $lastName"
    }
    else {
        # Assume a "First Last" format.
        $tokens = $cleanName.Split(" ", [System.StringSplitOptions]::RemoveEmptyEntries)
        if ($tokens.Count -ge 2) {
            $firstName = $tokens[0]
            $lastName = ($tokens[1..($tokens.Count-1)] -join " ")
            $fullName = "$firstName $lastName"
        }
        else {
            $firstName = $cleanName
            $lastName = ""
            $fullName = $cleanName
        }
    }
    # Use the cleaned full name for EmailDisplayAs as well.
    $EmailDisplayAs = $fullName
    # --- End Universal Cleanup ---

    $contactReport = @"
FirstName: $firstName
LastName: $lastName
Full Name: $fullName
Email: $Email
Email Display As: $EmailDisplayAs
"@
    
    if ($DryRun) {
        Write-Log "[DryRun] Would create contact:" 
        Write-Log $contactReport 
        $contactsCreated++
    }
    else {
        try {
            $newContact = $contactsFolder.Items.Add()
            $newContact.FullName = $fullName
            $newContact.FirstName = $firstName
            $newContact.LastName = $lastName
            $newContact.JobTitle = $JobTitle
            $newContact.Department = $Department
            $newContact.CompanyName = $Company
            $newContact.BusinessAddressStreet = $BusinessAddress
            $newContact.MobileTelephoneNumber = $Mobile
            $newContact.Email1Address = $Email
            $newContact.Email1DisplayName = $EmailDisplayAs
            $newContact.Email2Address = ""   # Not used
            $newContact.Email2DisplayName = ""  # Not used
            $newContact.Email3Address = ""   # Not used
            $newContact.Email3DisplayName = ""  # Not used

            $newContact.Save()
            Write-Log "Created contact:" 
            Write-Log $contactReport
            $contactsCreated++
        }
        catch {
            Write-Log "Error creating contact for ${fullName} with email ${Email}: $($_.Exception.Message)"
            $contactsSkipped++
        }
    }
}

Write-Log "Processing complete."
Write-Log "Total contacts created: $contactsCreated"
Write-Log "Total contacts skipped: $contactsSkipped"
