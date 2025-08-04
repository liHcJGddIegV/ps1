param(
    [string]$SourcePath = "C:\Users\YGonzalez\OneDrive - Invenergy LLC\Attachments",
    [string[]]$Extensions = @(".xlsx", ".pdf", ".kmz", ".zip", ".7z", ".xlsm", ".xls", ".png", ".exp", ".docx", ".doc", ".jpg", ".csv", ".vsdx", ".pptx", ".msg", ".dotx"),
    [string]$LogFile = "C:\Users\YGonzalez\OneDrive - Invenergy LLC\Attachments\file_mover_log.txt",
    [switch]$Verbose,
    [switch]$DryRun,           # New switch to perform a dry run without moving files.
    [switch]$GroupByDate       # New switch to group files by their creation date.
)

# Enable verbose output if the switch is set
if ($Verbose) {
    $VerbosePreference = "Continue"
} else {
    $VerbosePreference = "SilentlyContinue"
}

# Convert extensions to lowercase for case-insensitive comparison
$Extensions = $Extensions | ForEach-Object { $_.ToLower().Trim() }

# Prepare log file: clear it if it exists or create a new one
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
        Timestamp    = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
        Level        = $Level
        Message      = $Message
    }
    $logEntry | ConvertTo-Json -Compress | Add-Content -Path $LogFile
}

# Function to safely get destination path based on file extension and optionally creation date
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

# Collect destination folders to exclude them from the search
$destinationFolders = @()
foreach ($extension in $Extensions) {
    $folderName = $extension.TrimStart('.')
    $baseDestination = Join-Path -Path $SourcePath -ChildPath $folderName
    $destinationFolders += $baseDestination

    if ($GroupByDate -and (Test-Path $baseDestination)) {
        # If grouping by date, also exclude subfolders under each extension folder.
        $subFolders = Get-ChildItem -Path $baseDestination -Directory
        foreach ($sub in $subFolders) {
            $destinationFolders += $sub.FullName
        }
    }
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
        $destinationPath = Get-DestinationPath -file $file

        # Create the destination directory if it does not exist
        if (-not (Test-Path -Path $destinationPath)) {
            if (-not $DryRun) {
                New-Item -Path $destinationPath -ItemType Directory | Out-Null
            }
            Write-Verbose "Created directory $destinationPath"
        }

        # Handle duplicate file names
        $destinationFile = Join-Path -Path $destinationPath -ChildPath $file.Name
        if (Test-Path -Path $destinationFile) {
            $newFileName = "{0}_{1}" -f (Get-Date -Format "yyyyMMddHHmmssfff"), $file.Name
            $destinationFile = Join-Path -Path $destinationPath -ChildPath $newFileName
            Write-Verbose "File name conflict. Renamed to $newFileName"
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
