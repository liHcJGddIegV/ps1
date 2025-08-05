# Set to $true to simulate the removal of files without actually deleting them
$whatIf = $false

# Remove files containing "Diversion" in their filename
Function Remove-SpecificFiles($path)
{
    # Get all files containing "Diversion" in their filename in the directory and its subdirectories
    $files = Get-ChildItem -Path $path -Recurse -Force -File | Where-Object { $_.Name -like "*Diversion*" }

    # Delete each filev
    foreach ($file in $files)
    {
        Write-Host "Removing file '$($file.FullName)'"
        Remove-Item -Force -LiteralPath $file.FullName -WhatIf:$whatIf
    }
}

# Run the script
Remove-SpecificFiles -path "C:\Users\YGonzalez\Invenergy LLC\Project SCADA Engineering - Documents\General\2.- Sites and Projects\_AEP - Top Hat\"
