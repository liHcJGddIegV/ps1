param (
    [switch]$DryRun,                # Simulate saving attachments without writing files
    [switch]$VerboseOutput,         # Provide detailed logs
    [string]$AttachmentFolder = "C:\Users\YGonzalez\OneDrive - Invenergy LLC\Attachments"  # Folder to save attachments
)

# Define the log file location (adjust the path as needed)
$logFile = "C:\Users\YGonzalez\OneDrive - Invenergy LLC\Desktop\logs\AttachmentScriptLog.txt"

# Initialize counters
$attachmentsProcessed = 0
$attachmentsSaved = 0
$attachmentsSkipped = 0

# Helper function for logging messages with a timestamp
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

Write-Log "Starting attachment extraction from entire Inbox (including subfolders)."

# Create the attachment folder if it doesn't exist
if (!(Test-Path -Path $AttachmentFolder)) {
    if ($DryRun) {
        Write-Log "[DryRun] Would create folder: $AttachmentFolder" -Verbose
    }
    else {
        New-Item -ItemType Directory -Force -Path $AttachmentFolder | Out-Null
        Write-Log "Created folder: $AttachmentFolder" -Verbose
    }
}

# Create the Outlook COM object and get the MAPI namespace
try {
    $Outlook = New-Object -ComObject Outlook.Application
    $Namespace = $Outlook.GetNamespace("MAPI")
    Write-Log "Outlook COM object created." -Verbose
}
catch {
    Write-Log "Error creating Outlook COM object: $($_.Exception.Message)"
    exit
}

# Get the default Inbox folder (6 represents the Inbox)
try {
    $Inbox = $Namespace.GetDefaultFolder(6)
    Write-Log "Accessed default Inbox folder." -Verbose
}
catch {
    Write-Log "Error accessing Inbox folder: $($_.Exception.Message)"
    exit
}

# Recursive function to retrieve all MailItems from a folder and its subfolders
function Get-MailItemsFromFolder {
    param(
        [Parameter(Mandatory=$true)]
        $Folder
    )
    $mailItems = @()
    try {
        foreach ($item in $Folder.Items) {
            if ($item -is [Microsoft.Office.Interop.Outlook.MailItem]) {
                $mailItems += $item
            }
        }
    }
    catch {
        Write-Log "Error processing items in folder '$($Folder.Name)': $($_.Exception.Message)" -Verbose
    }
    foreach ($subFolder in $Folder.Folders) {
        $mailItems += Get-MailItemsFromFolder -Folder $subFolder
    }
    return $mailItems
}

# Get all mail items in the Inbox (and any subfolders)
$mailItems = Get-MailItemsFromFolder -Folder $Inbox
Write-Log "Total mail items found in Inbox and subfolders: $($mailItems.Count)" -Verbose

# Loop through each mail item and process attachments
foreach ($mail in $mailItems) {
    Write-Log "Processing email: $($mail.Subject)" -Verbose

    if ($mail.Attachments.Count -gt 0) {
        for ($i = 1; $i -le $mail.Attachments.Count; $i++) {
            $attachmentsProcessed++
            $attachment = $mail.Attachments.Item($i)
            $filePath = Join-Path $AttachmentFolder $attachment.FileName

            if ($DryRun) {
                Write-Log "[DryRun] Would save attachment '$($attachment.FileName)' from email '$($mail.Subject)' to '$filePath'" -Verbose
                $attachmentsSaved++
            }
            else {
                try {
                    $attachment.SaveAsFile($filePath)
                    Write-Log "Saved attachment '$($attachment.FileName)' from email '$($mail.Subject)' to '$filePath'" -Verbose
                    $attachmentsSaved++
                }
                catch {
                    Write-Log "Error saving attachment '$($attachment.FileName)' from email '$($mail.Subject)': $($_.Exception.Message)" -Verbose
                    $attachmentsSkipped++
                }
            }
        }
    }
    else {
        Write-Log "No attachments found in email: $($mail.Subject)" -Verbose
    }
}

Write-Log "Attachment extraction complete."
Write-Log "Total attachments processed: $attachmentsProcessed"
Write-Log "Total attachments saved: $attachmentsSaved"
Write-Log "Total attachments skipped: $attachmentsSkipped"
