param(
    [string]$SourcePath = "C:\Users\YGonzalez\OneDrive - Invenergy LLC\Attachments",
    [string[]]$Extensions = @(".xlsx", ".pdf", ".kmz", ".zip", ".7z", ".xlsm", ".xls", ".png", ".exp", ".docx", ".doc", ".jpg", ".csv", ".vsdx", ".pptx", ".msg" , ".dotx", "),
    [string]$LogFile = "C:\Users\YGonzalez\OneDrive - Invenergy LLC\Attachments\file_mover_log.txt",
    [switch]$Verbose
)

# Enable verbose output if the switch is set
if ($Verbose) {
    $VerbosePreference = "Continue"
} else {
    $VerbosePreference = "SilentlyContinue"
}

# Convert extensions to lowercase for case-insensitive comparison
$Extensions = $Extensions | ForEach-Object { $_.ToLower() }

# Clear the log file at the start
if (Test-Path -Path $LogFile) {
    Clear-Content -Path $LogFile
} else {
    New-Item -Path $LogFile -ItemType File | Out-Null
}

# Function to write log entries in structured format
function Write-Log {
    param(
        [string]$Message,
        [string]$Level = "INFO"
    )
    $logEntry = @{
        Timestamp = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
        Level     = $Level
        Message   = $Message
    }
    $logEntry | ConvertTo-Json -Compress | Add-Content -Path $LogFile
}

# Collect destination folders to exclude them from the search
$destinationFolders = @()
foreach ($extension in $Extensions) {
    $folderName = $extension.TrimStart('.')
    $destinationPath = Join-Path -Path $SourcePath -ChildPath $folderName
    $destinationFolders += $destinationPath
}

# Get all files in the source directory and subdirectories
Write-Verbose "Scanning for files in $SourcePath..."
$files = Get-ChildItem -Path $SourcePath -File -Recurse |
         Where-Object {
             ($Extensions -contains $_.Extension.ToLower()) -and
             ($destinationFolders -notcontains $_.DirectoryName)
         }

$totalFiles = $files.Count
$currentFileIndex = 0

Write-Log -Message "Found $totalFiles files to process."

foreach ($file in $files) {
    $currentFileIndex++
    Write-Progress -Activity "Processing files" -Status "Processing $($file.Name)" -PercentComplete (($currentFileIndex / $totalFiles) * 100)

    try {
        $extension = $file.Extension.ToLower()
        $folderName = $extension.TrimStart('.')
        $destinationPath = Join-Path -Path $SourcePath -ChildPath $folderName

        # Create the destination directory if it does not exist
        if (-not (Test-Path -Path $destinationPath)) {
            New-Item -Path $destinationPath -ItemType Directory | Out-Null
            Write-Verbose "Created directory $destinationPath"
        }

        # Handle duplicate file names
        $destinationFile = Join-Path -Path $destinationPath -ChildPath $file.Name
        if (Test-Path -Path $destinationFile) {
            $newFileName = "{0}_{1}" -f (Get-Date -Format "yyyyMMddHHmmssfff"), $file.Name
            $destinationFile = Join-Path -Path $destinationPath -ChildPath $newFileName
            Write-Verbose "File name conflict. Renamed to $newFileName"
        }

        # Move the file
        Move-Item -Path $file.FullName -Destination $destinationFile -Force -ErrorAction Stop

        # Log the action
        $message = "Moved '$($file.FullName)' to '$destinationFile'"
        Write-Verbose $message
        Write-Log -Message $message
    }
    catch [System.IO.IOException] {
        $errorMessage = "IO Error moving file $($file.FullName): $($_.Exception.Message)"
        Write-Warning $errorMessage
        Write-Log -Message $errorMessage -Level "ERROR"
    }
    catch [System.UnauthorizedAccessException] {
        $errorMessage = "Access denied moving file $($file.FullName): $($_.Exception.Message)"
        Write-Warning $errorMessage
        Write-Log -Message $errorMessage -Level "ERROR"
    }
    catch {
        $errorMessage = "Error moving file $($file.FullName): $($_.Exception.Message)"
        Write-Warning $errorMessage
        Write-Log -Message $errorMessage -Level "ERROR"
    }
}

Write-Progress -Activity "Processing files" -Completed
Write-Log -Message "File moving process completed."
