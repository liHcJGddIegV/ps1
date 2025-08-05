# Path to the CSV file with user information
$csvPath = "C:\Users\YGonzalez\Downloads\invenergy_user_data.csv"

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
    $subject = "Your Engineering Workstation Account - Diversion Wind Project"

    # Revised email body
    $body = @"
Dear $firstName,

We are pleased to inform you that your engineering workstation account for Diversion wind project has been successfully created. Please find your login credentials below:

Username: [$($user.Username)]
Domain: [CORP\]
Temporary Password: [$($user.'Temporary Password')]

Important Instructions:
First-Time Login: After you log in for the first time, please change your temporary password. Please choose a secure password that meets our security guidelines.

Accessing the System: You can log in to the Engineering workstation on this AEP jumphost from the ICC workstation [204.29.193.1].

Password Security: For security reasons, please do not share your password. If you need assistance or have any questions, contact our support team.

Thank you, and welcome to Diversion wind project!

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
