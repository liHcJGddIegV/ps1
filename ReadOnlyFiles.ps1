# Set the path to the directory where you want to change file properties
$path = "C:\Users\Tech"

# Get all files in the directory and its subdirectories
$files = Get-ChildItem -Path $path -Recurse -File

# Iterate through each file and change its properties
foreach ($file in $files) {
    # Remove the read-only attribute
    $file.Attributes = $file.Attributes -band (-bnot [System.IO.FileAttributes]::ReadOnly)

    # Optional: Output the file name and its updated attributes
    Write-Host "$($file.Name) attributes updated to: $($file.Attributes)"
}

# Inform the user that the process is complete
Write-Host "All files in '$path' have been updated to writable."
