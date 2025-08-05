# Set to $true to simulate the removal of files without actually deleting them
$whatIf = $false

# Remove files named ".DS_Store"
Function Remove-DS_StoreFiles($path)
{
    # Get all files named ".DS_Store" in the directory and its subdirectories
    $files = Get-ChildItem -Path $path -Recurse -Force -File | Where-Object { $_.Name -eq ".DS_Store" }

    # Delete each file
    foreach ($file in $files)
    {
        Write-Host "Removing file '$($file.FullName)'"
        Remove-Item -Force -LiteralPath $file.FullName -WhatIf:$whatIf
    }
}

# Run the script
Remove-DS_StoreFiles -path "C:\Users\YGonzalez\OneDrive - Invenergy LLC"
