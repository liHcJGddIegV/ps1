# Set the path to the source directory where you want to search for files
$sourcePath = "C:\Users\YGonzalez\OneDrive - Invenergy LLC\Attachments"

# Define an array of file extensions to process
$extensions = @(".xlsx", ".pdf", ".kmz", ".zip", ".xlsm", ".png", ".docx", ".xls")

# Iterate through each file extension
foreach ($extension in $extensions) {
    # Set the path to the destination directory for the current extension
    $destinationPath = Join-Path -Path $sourcePath -ChildPath $extension

    # Create the destination directory if it does not exist
    if (-not (Test-Path -Path $destinationPath)) {
        New-Item -Path $destinationPath -ItemType Directory
    }

    # Get all files in the source directory with the current extension (not looking in subdirectories)
    $files = Get-ChildItem -Path $sourcePath -File -Filter "*$extension"

    # Log the number of files found that match the criteria
    Write-Host "Found $($files.Count) $extension files."

    # Iterate through each file and move it to the destination directory
    foreach ($file in $files) {
        # Move the file
        Move-Item -Path $file.FullName -Destination $destinationPath

        # Optional: Output the filename and its new location
        Write-Host "Moved '$($file.Name)' to '$destinationPath'"
    }

    # Inform the user that the process is complete for the current extension
    Write-Host "File moving process for $extension files completed."
}
