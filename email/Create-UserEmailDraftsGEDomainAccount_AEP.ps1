# Path to the CSV file with user information
$csvPath = "C:\Users\YGonzalez\Downloads\AEP_user_data.csv"

# Load user data from CSV
$users = Import-Csv -Path $csvPath

# Create an instance of the Outlook application
$Outlook = New-Object -ComObject Outlook.Application

# Iterate through each user in the CSV
foreach ($user in $users) {
    # Check for missing fields
    $missingFields = @()
    if (-not $user.'First Name') { $missingFields += "First Name" }
    if (-not $user.Username) { $missingFields += "Username" }
    if (-not $user.'Temporary Password') { $missingFields += "Temporary Password" }
    if (-not $user.Email) { $missingFields += "Email" }

    if ($missingFields.Count -gt 0) {
        Write-Warning "Incomplete data for user $($user.Username). Missing fields: $($missingFields -join ', '). Skipping."
        continue
    }

    # Use the first name from the CSV
    $firstName = $user.'First Name'

    # Customize the email subject
    $subject = "Your GE Domain Account for the Top Hat Wind Project"

    # Revised email body
    $body = @"
Dear $firstName,

We are excited to inform you that your GE domain account for the Top Hat Wind Project has been successfully created. Please find your login credentials below:

Username: [$($user.Username)]
Domain: [ILTOPHAT\$($user.Username)]
Temporary Password: [$($user.'Temporary Password')]

Important Instructions:

First-Time Login:
Upon your first login, please update your temporary password. Ensure your new password meets our security guidelines.

Accessing the System:
Use the AEP jumphost at 204.29.193.85 to access the GE system.

Password Security:
For your protection, do not share your password with anyone. If you encounter any issues or have questions, please contact our support team for assistance.

Thank you, and welcome to the Top Hat Wind Project!

Best regards,
"@

    # Create a new mail item in Outlook
    $mail = $Outlook.CreateItem(0) # 0 indicates a MailItem

    # Set the properties of the mail item
    $mail.To = $user.Email
    $mail.Subject = $subject
    $mail.Body = $body

    # Save the mail item to the Drafts folder
    $mail.Save()

    # Log or display confirmation
    Write-Host "Draft email created for $firstName <$($user.Email)>"
}

# Clean up the COM object
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Outlook) | Out-Null
Remove-Variable -Name Outlook
[GC]::Collect()
[GC]::WaitForPendingFinalizers()
