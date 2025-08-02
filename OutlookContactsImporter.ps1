# Create a new Outlook application instance
$outlook = New-Object -ComObject Outlook.Application
$namespace = $outlook.GetNamespace("MAPI")

# Function to add email to contacts
function AddEmailToContacts($email, $firstName, $lastName) {
    $contacts = $namespace.GetDefaultFolder([Microsoft.Office.Interop.Outlook.OlDefaultFolders]::olFolderContacts)
    $contact = $contacts.Items.Add()
    $contact.Email1Address = $email
    $contact.FirstName = $firstName
    $contact.LastName = $lastName
    $contact.Save()
}

# List of emails and names
$emails = @(
    @{Email="Jared.Arlt@mortenson.com"; FirstName="Jared"; LastName="Arlt"},
    @{Email="Jack.Shopbell@mortenson.com"; FirstName="Jack"; LastName="Shopbell"},
    @{Email="Justin.Speller@mortenson.com"; FirstName="Justin"; LastName="Speller"},
    @{Email="Mohit.Dwivedi@mortenson.com"; FirstName="Mohit"; LastName="Dwivedi"},
    @{Email="Logan.Runge@mortenson.com"; FirstName="Logan"; LastName="Runge"},
    @{Email="jmwilliams1@aep.com"; FirstName="James M"; LastName="Williams"},
    @{Email="jtparkison@aep.com"; FirstName="Joseph T"; LastName="Parkison"},
    @{Email="fmkarr@aep.com"; FirstName="Frank"; LastName="Karr"},
    @{Email="AColozza@invenergy.com"; FirstName="Anthony"; LastName="Colozza"},
    @{Email="gjdaft1@aep.com"; FirstName="Garrick J"; LastName="Daft"},
    @{Email="jafoster1@aep.com"; FirstName="Jacob A"; LastName="Foster"},
    @{Email="bbdelyons@aep.com"; FirstName="Bonnie J"; LastName="de Lyons"},
    @{Email="kcooper1@aep.com"; FirstName="Kenneth J"; LastName="Cooper"},
    @{Email="allowther@aep.com"; FirstName="Allowther"; LastName=""},
    @{Email="fpsouza-neto@aep.com"; FirstName="Fpsouza-neto"; LastName=""},
    @{Email="cdferguson@aep.com"; FirstName="Cdferguson"; LastName=""},
    @{Email="tsmarcum@aep.com"; FirstName="Tsmarcum"; LastName=""},
    @{Email="beshields@aep.com"; FirstName="Beshields"; LastName=""},
    @{Email="cjwise@aep.com"; FirstName="Cjwise"; LastName=""},
    @{Email="pfkeogh@aep.com"; FirstName="Pfkeogh"; LastName=""},
    @{Email="bacrowder@aep.com"; FirstName="Bacrowder"; LastName=""}
)

# Add each email to the contacts
foreach ($email in $emails) {
    AddEmailToContacts $email.Email $email.FirstName $email.LastName
}

# Cleanup
$outlook.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook)
