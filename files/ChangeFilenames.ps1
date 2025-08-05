# Set the path to the directory where you want to change filenames
$path = "C:\Users\Tech\Downloads"

# Define the text pattern to find and the replacement text
$findText = "Statement Summary  Summary Summary"
$replaceText = "Statement Summary "

# Get all files in the directory and subdirectories that contain the text to find
$files = Get-ChildItem -Path $path -File -Recurse | Where-Object { $_.Name -like "*$findText*" }

# Log the number of files found that match the criteria
Write-Host "Found $($files.Count) files matching the criteria."

# Iterate through each file and rename it
foreach ($file in $files) {
    # Create the new filename by replacing the find text with the replacement text
    $newFilename = $file.Name -replace [regex]::Escape($findText), $replaceText

    # Check if the new filename is different from the original
    if ($newFilename -ne $file.Name) {
        # Rename the file
        Rename-Item -Path $file.FullName -NewName $newFilename
        Write-Host "Renamed '$($file.Name)' to '$newFilename'"
    } else {
        Write-Host "No change required for '$($file.Name)'"
    }
}

# Inform the user that the process is complete
Write-Host "File renaming process completed."
