# Set the path to the source directory where you want to search for files
$sourcePath = "C:\Users\Tech\Downloads"

# Set the path to the destination directory where you want to move the files
$destinationPath = "C:\Users\Tech\Downloads\_AEP - Diversion"

# Create the destination directory if it does not exist
if (-not (Test-Path -Path $destinationPath)) {
    New-Item -Path $destinationPath -ItemType Directory
}

# Get all files in the source directory and subdirectories that contain "Diversion" in their filename
$files = Get-ChildItem -Path $sourcePath -File -Recurse | Where-Object { $_.Name -like "*Diversion*" }

# Log the number of files found that match the criteria
Write-Host "Found $($files.Count) files matching the criteria."

# Iterate through each file and move it to the destination directory
foreach ($file in $files) {
    # Move the file
    Move-Item -Path $file.FullName -Destination $destinationPath

    # Optional: Output the filename and its new location
    Write-Host "Moved '$($file.Name)' to '$destinationPath'"
}

# Inform the user that the process is complete
Write-Host "File moving process completed."
