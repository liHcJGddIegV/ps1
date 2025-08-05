# PowerShell Script Collection

This repository contains assorted automation scripts organized by topic.

## Directory layout

- `contacts/` – manage Outlook and other contact information.
- `email/` – scripts for creating email drafts, saving attachments, and cleaning Exchange data.
- `files/` – utilities for renaming, moving, and removing files or directories.
- `network/` – network diagnostics and related helpers.
- `python/` – miscellaneous Python tools.

Run scripts from the repository root. For example:

```powershell
# Update Outlook contacts
.\contacts\UpdateOutlookContacts.ps1 -DryRun -VerboseOutput
```

```powershell
# Test network quality
.\network\Test-NetworkQuality.ps1
```
