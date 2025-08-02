param(
    [string]$SourcePath = "C:\Users\YGonzalez\OneDrive - Invenergy LLC\Attachments",
    [string[]]$Extensions = @(".xlsx", ".pdf", ".kmz", ".zip", ".7z", ".xlsm", ".xls", ".png", ".exp", ".docx", ".doc", ".jpg", ".csv", ".vsdx", ".pptx", ".msg", ".dotx"),
    [string]$LogFile = "C:\Users\YGonzalez\OneDrive - Invenergy LLC\Attachments\file_mover_log.txt",
    [string]$DetailedLogFile = "C:\Users\YGonzalez\OneDrive - Invenergy LLC\Desktop\Logs\MoveFilesByExtension.txt",
    [switch]$Verbose,
    [switch]$DryRun,           # Perform a dry run (log actions without moving files).
    [switch]$GroupByDate,      # Group files by their creation date.
    [switch]$Dynamic         # Process all file types found in the source.
)

# Enable verbose output if the switch is set
if ($Verbose) {
    $VerbosePreference = "Continue"
} else {
    $VerbosePreference = "SilentlyContinue"
}

# In non-dynamic mode, clean up the extensions array
if (-not $Dynamic) {
    $Extensions = $Extensions | ForEach-Object { $_.ToLower().Trim() }
}

# Ensure the folder for the main log file exists
$LogFolder = Split-Path -Path $LogFile
if (-not (Test-Path -Path $LogFolder)) {
    New-Item -Path $LogFolder -ItemType Directory -Force | Out-Null
}

# Prepare main log file: clear it if it exists or create a new one
if (Test-Path -Path $LogFile) {
    Clear-Content -Path $LogFile
} else {
    New-Item -Path $LogFile -ItemType File | Out-Null
}

# Ensure the folder for the detailed log file exists, then clear/create the detailed log file.
$DetailedLogFolder = Split-Path -Path $DetailedLogFile
if (-not (Test-Path -Path $DetailedLogFolder)) {
    New-Item -Path $DetailedLogFolder -ItemType Directory -Force | Out-Null
}
if (Test-Path -Path $DetailedLogFile) {
    Clear-Content -Path $DetailedLogFile
} else {
    New-Item -Path $DetailedLogFile -ItemType File | Out-Null
}

# Function to write log entries in structured JSON format to both log files
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
    $logJson = $logEntry | ConvertTo-Json -Compress
    # Append to the main log file
    $logJson | Add-Content -Path $LogFile
    # Append to the detailed log file for troubleshooting
    $logJson | Add-Content -Path $DetailedLogFile
}

# Function to determine destination path based on file extension and optionally creation date
function Get-DestinationPath {
    param(
        [System.IO.FileInfo]$file
    )
    $extension = $file.Extension.ToLower()
    $folderName = $extension.TrimStart('.')
    $destinationPath = Join-Path -Path $SourcePath -ChildPath $folderName

    if ($GroupByDate) {
        # Create a subfolder based on file creation date (formatted as yyyy-MM-dd)
        $dateFolder = $file.CreationTime.ToString("yyyy-MM-dd")
        $destinationPath = Join-Path -Path $destinationPath -ChildPath $dateFolder
    }
    return $destinationPath
}

# Determine destination folders to exclude from processing and get the list of files.
if ($Dynamic) {
    # In dynamic mode, assume destination folders are any subfolders directly under $SourcePath.
    $destinationFolders = (Get-ChildItem -Path $SourcePath -Directory).FullName
    # Note: This prevents re-processing files already moved into a subfolder.
    $files = Get-ChildItem -Path $SourcePath -File -Recurse | Where-Object { 
        $_.Extension -and ($destinationFolders -notcontains $_.DirectoryName)
    }
}
else {
    # In static mode, only process files with extensions specified in $Extensions.
    $destinationFolders = @()
    foreach ($extension in $Extensions) {
        $folderName = $extension.TrimStart('.')
        $baseDestination = Join-Path -Path $SourcePath -ChildPath $folderName
        $destinationFolders += $baseDestination

        if ($GroupByDate -and (Test-Path $baseDestination)) {
            # Also exclude subfolders under each extension folder.
            $subFolders = Get-ChildItem -Path $baseDestination -Directory
            foreach ($sub in $subFolders) {
                $destinationFolders += $sub.FullName
            }
        }
    }
    $files = Get-ChildItem -Path $SourcePath -File -Recurse | Where-Object {
        ($Extensions -contains $_.Extension.ToLower()) -and
        ($destinationFolders -notcontains $_.DirectoryName)
    }
}

$totalFiles = $files.Count
$currentFileIndex = 0

Write-Log -Message "Found $totalFiles files to process."

foreach ($file in $files) {
    $currentFileIndex++
    Write-Progress -Activity "Processing files" -Status "Processing $($file.Name)" -PercentComplete (($currentFileIndex / $totalFiles) * 100)
    
    try {
        $destinationPath = Get-DestinationPath -file $file

        # Create the destination directory if it does not exist
        if (-not (Test-Path -Path $destinationPath)) {
            if (-not $DryRun) {
                New-Item -Path $destinationPath -ItemType Directory | Out-Null
            }
            Write-Verbose "Created directory $destinationPath"
            Write-Log -Message "Created directory $destinationPath"
        }

        # Handle duplicate file names by appending a timestamp if necessary
        $destinationFile = Join-Path -Path $destinationPath -ChildPath $file.Name
        if (Test-Path -Path $destinationFile) {
            $newFileName = "{0}_{1}" -f (Get-Date -Format "yyyyMMddHHmmssfff"), $file.Name
            $destinationFile = Join-Path -Path $destinationPath -ChildPath $newFileName
            Write-Verbose "File name conflict. Renamed to $newFileName"
            Write-Log -Message "File name conflict for '$($file.FullName)'. Renamed to $newFileName"
        }

        $actionDescription = "Moving '$($file.FullName)' (Size: $($file.Length) bytes) to '$destinationFile'"

        if ($DryRun) {
            # In dry run mode, only log the intended action.
            Write-Verbose "[Dry Run] $actionDescription"
            Write-Log -Message "[Dry Run] $actionDescription"
        }
        else {
            # Move the file
            Move-Item -Path $file.FullName -Destination $destinationFile -Force -ErrorAction Stop
            Write-Verbose $actionDescription
            Write-Log -Message $actionDescription
        }
    }
    catch [System.IO.IOException] {
        $errorMessage = "IO Error moving file '$($file.FullName)': $($_.Exception.Message)"
        Write-Warning $errorMessage
        Write-Log -Message $errorMessage -Level "ERROR"
    }
    catch [System.UnauthorizedAccessException] {
        $errorMessage = "Access denied moving file '$($file.FullName)': $($_.Exception.Message)"
        Write-Warning $errorMessage
        Write-Log -Message $errorMessage -Level "ERROR"
    }
    catch {
        $errorMessage = "Error moving file '$($file.FullName)': $($_.Exception.Message)"
        Write-Warning $errorMessage
        Write-Log -Message $errorMessage -Level "ERROR"
    }
}

Write-Progress -Activity "Processing files" -Completed
Write-Log -Message "File moving process completed."
