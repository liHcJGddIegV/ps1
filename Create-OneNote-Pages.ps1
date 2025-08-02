Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName Microsoft.VisualBasic

# Ensure OneNote is active
[Microsoft.VisualBasic.Interaction]::AppActivate("OneNote")
Start-Sleep -Seconds 2

# List of page titles
$attachments = @(
    "Exhibit A – Form of Seller Parent Guaranty",
    "Exhibit B – Form of Build-Out Agreement",
    "Exhibit C – Definition of Substantial Completion and Related Definitions",
    "Exhibit D – Project Warranty",
    "Exhibit E – Closing Title Endorsements",
    "Exhibit F-1 – Form of Land Contract Estoppel",
    "Exhibit F-2 – Form of Major Project Document Estoppel",
    "Exhibit G – Form of Notice to Proceed",
    "Exhibit H – Form of Letter of Credit",
    "Exhibit I – Seller’s Account and Wire Transfer Information",
    "Exhibit J-1 – Invenergy Start of Construction Certificate (Execution Date)",
    "Exhibit J-2 – Invenergy Start of Construction Certificate (Determination Date)",
    "Exhibit J-3 – Invenergy Start of Construction Certificate (Closing Date)",
    "Exhibit K – Form of Monthly Report",
    "Exhibit L – Form of Financing Consent",
    "Exhibit M – Form of Buyer In-House Counsel",
    "Exhibit N – Form of Membership Interest Assignment",
    "Exhibit O-1 – Terms of BOP Contract (with Insurance Attachment)",
    "Exhibit P-1 – Form of Amendment to Land Contracts",
    "Exhibit P-2 – Terms of New Land Contracts",
    "Exhibit Q – Form of Seller Release",
    "Exhibit R – Form of Excluded Assets and Liabilities Assignment Agreement",
    "Exhibit S – Form of Shared Facilities Agreement",
    "Exhibit V – Form of Cost Breakdown for FERC Purposes",
    "Exhibit W – Form of Major Project Document Counterparty Lien Waiver",
    "Exhibit X – Form of O&M Agreement"
)

foreach ($item in $attachments) {
    # Simulate Ctrl+N to create new page
    [System.Windows.Forms.SendKeys]::SendWait("^{n}")
    Start-Sleep -Milliseconds 700

    # Send the title text
    [System.Windows.Forms.SendKeys]::SendWait($item)
    [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
    Start-Sleep -Milliseconds 300
}
