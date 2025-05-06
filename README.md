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
  - https://github.com/yuukioosaka/redmine-outlook-addin/releases
- extract and run Setup.exe
- Start Outlook(classic), and Close. "user.config" will create automatically.

## Configuration
Before using the add-in, you need to configure the `user.config` file. 
This file contains essential settings for connecting to Redmine and customizing the behavior of the add-in.  
you can find user.config below  
%LOCALAPPDATA%\Local\Apps\2.0\Data\[randomid]\crmo..vsto_[randomid]\Data\16.0.18730.20122  
ex)  
C:\Users\username\AppData\Local\Apps\2.0\Data\TZ13HK22.WN8\0TO1CPPV.XHW\crmo..vsto_061175295e4e6d57_0001.0000_7ac4fe303c687902\Data\16.0.18730.20122

### Configuration File: `user.config`
Below is an example configuration file and instructions for each setting:
```
<?xml version="1.0" encoding="utf-8"?>
<configuration>
    <userSettings>
        <CrmOutlookAddIn.Properties.Settings>
            <!-- Redmine URL --> 
            <setting name="RedmineUrl" serializeAs="String">
                <value>http://localhost:3000</value>
            </setting>
            <!-- Redmine API Key -->
            <setting name="RedmineApiKey" serializeAs="String">
                <value>9b3572bf4f2cccdcdd0c254371a38babeb1004c7</value>
            </setting>
            <!-- Prefix for ticket IDs in email subjects -->
            <setting name="idprefix" serializeAs="String">
                <value>id:</value>
            </setting>
            <!-- Delimiters for detecting quoted text in email replies -->
            <setting name="ReplyDelimiter1" serializeAs="String">
                <value>^On .+ wrote:</value>
            </setting>
            <setting name="ReplyDelimiter2" serializeAs="String">
                <value>^From: .+</value>
            </setting>
            <setting name="ReplyDelimiter3" serializeAs="String">
                <value>^-----Original Message-----</value>
            </setting>
            <setting name="ReplyDelimiter4" serializeAs="String">
                <value>^\d{4}年\d{1,2}月\d{1,2}日(.+) \d{1,2}:\d{2} .+ .+@.+..+:</value>
            </setting>
            <setting name="Init" serializeAs="String">
                <value>initialized</value>
            </setting>
        </CrmOutlookAddIn.Properties.Settings>
    </userSettings>
</configuration>
...
```

### Key Settings
1. **RedmineUrl**: The base URL of your Redmine instance. Example: `http://localhost:3000`.
2. **RedmineApiKey**: Your Redmine API key. This is required for authentication.
3. **idprefix**: The prefix used in email subjects to identify Redmine ticket IDs. Example Your Email Title indicate ticket id 1234: `id:1234`.
4. **ReplyDelimiterX**: Regular expressions to detect quoted text in email replies. These delimiters help trim unnecessary content when logging email bodies to Redmine.

### Logging
Logs are written to a file in the `%TEMP%CrmOutlookAddIn.log`. 

## Usage
1. Start Outlook after installing the add-in.
2. The add-in will automatically monitor your Inbox and Sent Items folders.
3. Emails with a subject containing a ticket ID (e.g., `[id:1234] Addins Bugs Post.`) will be logged to the corresponding Redmine issue.
4. Use the right click context menu to manually create a "New Redmine Tick" from a selected email.

## Troubleshooting
- Ensure the `user.config` file is correctly configured.
- Check the log file (`%TEMP%\CrmOutlookAddIn.log`) for detailed error messages.
- Verify that your Redmine instance is accessible and the API key is valid.
