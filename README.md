# redmine-outlook-addin
This Outlook add-in integrates with Redmine to automatically log email activities and create or update Redmine issues based on email content.
This project is built using C# and targets the .NET Framework 4.8.

## Features
- Automatically logs emails (both sent and received) to Redmine issues based on ticket IDs in the email subject.
- Prevents duplicate comments in Redmine by checking existing journal entries.
- Provides a right click context menu to manually create a new Redmine issue from an email.
- Configurable settings for Redmine URL, API key, and email subject prefix.

## Prerequisites
- Microsoft Outlook (compatible with VSTO add-ins).
- A Redmine instance with API access enabled.
- .NET Framework 4.8 installed on the system.

## Installation

- Download ClickOnceSetup.zip from Releases
  - [https://github.com/yuukioosaka/redmine-outlook-addin/releases](https://github.com/yuukioosaka/redmine-outlook-addin/releases)
- extract and run Setup.exe
- Start Outlook(classic). The registry settings will be created automatically with default values.

## Configuration

The add-in uses Windows Registry to store its settings. All settings are stored under the following registry key:

```registry
HKEY_CURRENT_USER\Software\CrmOutlookAddIn
```

You can modify the settings using the Windows Registry Editor (regedit.exe) or through PowerShell commands.

### Registry Settings

Below are the available registry settings and their descriptions:

1. **RedmineUrl** (String)
   - Default value: `http://redmine.example.com`
   - The base URL of your Redmine instance

2. **RedmineApiKey** (String)
   - Default value: ""
   - Your Redmine API key, required for authentication

3. **idprefix** (String)
   - Default value: "[id-"
   - The prefix used in email subjects to identify Redmine ticket IDs
   - Example: With default settings, an email with subject "[id-1234] Bug Fix" will be linked to ticket #1234

4. **ReplyDelimiter1** through **ReplyDelimiter4** (String)
   - Default values:
     - ReplyDelimiter1: "From:"
     - ReplyDelimiter2: "差出人:"
     - ReplyDelimiter3: "-----Original Message-----"
     - ReplyDelimiter4: "From "
   - Regular expressions to detect quoted text in email replies
   - These help trim unnecessary content when logging email bodies to Redmine

5. **UseCurlClient** (DWORD)
   - Default value: 0 (False)
   - Set to 1 (True) if you want to bypass corporate proxy servers that block Redmine access

### PowerShell Example

You can use PowerShell to configure the settings. Here's an example:

```powershell
# Set Redmine URL
Set-ItemProperty -Path "HKCU:\Software\CrmOutlookAddIn" -Name "RedmineUrl" -Value "http://your-redmine-server.com"

# Set API Key
Set-ItemProperty -Path "HKCU:\Software\CrmOutlookAddIn" -Name "RedmineApiKey" -Value "your-api-key-here"
```

### Logging

Logs are written to a file in the `%TEMP%\CrmOutlookAddIn.log`.

## Usage

1. Start Outlook and the add-in will initialize with default registry settings
2. The add-in will automatically monitor your Inbox and Sent Items folders
3. Emails with a subject containing a ticket ID (e.g., `[id-1234] Addins Bugs Post`) will be logged to the corresponding Redmine issue
4. Use the right click context menu to manually create a "New Redmine Ticket" from a selected email

## Troubleshooting

- Check the log file (`%TEMP%\CrmOutlookAddIn.log`) for detailed error messages
- Verify that your Redmine instance is accessible and the API key is correctly configured in the registry
- Ensure the registry settings under `HKEY_CURRENT_USER\Software\CrmOutlookAddIn` are properly configured
