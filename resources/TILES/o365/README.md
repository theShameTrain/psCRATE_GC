# o365 TILE Overview

This tile allows regular users to connect to o365 as a service account and manage Mailbox and Calendar Permissions for o365 user accounts

### Settings
Settings are located in \resources\XML\Config.xml in \<o365\>
+ msolUser
    + o365 service account that has rights to view users mailbox and canendar settings
+ encPassword
    + This is an AES encrypted string.  The string is converted to a secure string using the AESkey in the PSCredential Object
+ emailSuffix
    + This is used the string to append to a username if no '@' symbol is detected when submitting a name in the window

### Dependencies
+ AESencryped password for service account
+ AzureAD management Shell - v1 (Connect-msolService) [link to version 1](https://docs.microsoft.com/en-us/powershell/module/MSOnline/?view=azureadps-1.0)
+ Exchange Online Management Shell (Loads automatically)

### Tools

+ Get Mailbox Permissions
+ Get SendAs Permissions
+ Get SendOnBehalf Permissions
+ Get Calendar Permissions
