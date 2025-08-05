# Set to true to test the script
$whatIf = $false

# Remove hidden files, like thumbs.db
$removeHiddenFiles = $false

# Function to remove empty folders
Function Remove-EmptyFolder($path)
{
    # Process each subfolder recursively
    Foreach ($subFolder in Get-ChildItem -LiteralPath $path -Directory -Force)
    {
        Remove-EmptyFolder -path $subFolder.FullName
    }

    # Optionally remove hidden files
    if ($removeHiddenFiles) {
        Get-ChildItem -LiteralPath $path -File -Hidden -Force | ForEach-Object {
            Write-Host "Removing hidden file '$($_.FullName)'"
            if ($whatIf) {
                Remove-Item -LiteralPath $_.FullName -Force -WhatIf
            } else {
                Remove-Item -LiteralPath $_.FullName -Force
            }
        }
    }

    # Get all child items, including hidden ones if not removing them
    if ($removeHiddenFiles) {
        $subItems = Get-ChildItem -LiteralPath $path -Force
    } else {
        $subItems = Get-ChildItem -LiteralPath $path -Force | Where-Object { -not $_.Attributes.HasFlag([IO.FileAttributes]::Hidden) }
    }

    # If there are no items, delete the folder
    If ($subItems.Count -eq 0)
    {
        Write-Host "Removing empty folder '$path'"
        try {
            if ($whatIf) {
                Remove-Item -LiteralPath $path -Force -WhatIf -Confirm:$false
            } else {
                Remove-Item -LiteralPath $path -Force -Confirm:$false
            }
        } catch {
            Write-Host "Error removing folder '$path': $($_.Exception.Message)" -ForegroundColor Red
        }
    }
}

# Run the script
Remove-EmptyFolder -path "C:\Users\YGonzalez\OneDrive - Invenergy LLC\Attachments" -Verbose
