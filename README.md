# psCRATE

This is a WPF wrapper for various Powershell Tools. (This is a work in progress)  
The idea for this is that every toolset will run as a service account that allows regualr (non-elevated) accounts to manage specific items or tasks.  This can be used to extend helpdesk, firstlevel or regular users access to items in a controled manner without having to grant anyone any special access.  


## TODO
+ Control tiles via group membership

---

### TOOLS
Current items:
1. o365 (relies on AESencryped password for service account, o365 management Shell)
+ Get Mailbox Permissions
+ Get SendAs Permissions
+ Get SendOnBehalf Permissions
+ Get Calendar Permissions

2. Shared Folders (relies on Quest ActiveRoles ADmanagement Shell)
+ Get Self Service Service Folder Manager
+ Create Self Service Folder




### Configuration
\resources\XML\Config.xml  
this file gets loaded into the initial syncHash.  You must add your own settings to this file
