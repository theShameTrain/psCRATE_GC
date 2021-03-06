# psCRATE

This is a WPF wrapper for various Powershell Tools. (This is a work in progress)  
The idea for this is that every toolset will run as a service account that allows regualr (non-elevated) accounts to manage specific items or tasks.  This can be used to extend helpdesk, firstlevel or regular users access to items in a controled manner without having to grant anyone any special access.  

## Overview

Main script runs in a seperate runspace.  All tiles run in another runspace.

To add a tile create a folder in \resources\TILES\\_**TILE**_
  + Create a config.xml in the _**TILE**_ folder
     + Config contains Tile Settings for:
       + Name
       + Width
       + Height
       + Tile Title
       + Tile Accent Color (See MahApps [Style Guide](http://mahapps.com/guides/styles.html) for color theme options)
       + Icon (See Mahapps [IconPacks](https://github.com/MahApps/MahApps.Metro.IconPacks). The IconBrowser is a great tool!)

#### Main script psCrate.ps1
+ First create a GLOBAL synchronized hashtable
+ Read config from \resources\XML\config.xml
    + Any organization specific information should reside in this Config
+ Add WPF assemblies and DLL's for [MahApps.Metro](http://mahapps.com/)
+ Create a runspace for the main WPF and sync the SyncHash to the runspace
+ Add powershell script to runspace
  + load XAML for Main Window
  + load XML for each TILE from each folder in \resources\TILES
  + foreach tile found create the tile XAML and Append to Main XAML
  + Find all 'NAMES' in XAML and set to SyncHash
  + Create jobCleanup Runspace
  + Add the Click events for each TILE and import the click event from \resources\TILES\<TILE>\<TILE>.add_Click.ps1
  + Show the Window Dialog
+ Invoke the Parent runspace
+ Create loop to wait for Main Runspace to finish
+ When Main Runspace is done Dispose of all runspaces that are not 'Busy"

#### Tile.add_Click.ps1
+ First disable the Tile in the main window so only one child runspace per tile is available
+ Import any Modules that are required for the script 
+ Create a runspace for the child process to run in
+ Add powershell script to runspace
  + Add Function _**Update-Window**_ that allows updating the window from the new runspace
  + Add Functions _**close-SplashScreen**_ and _**start-SplashScreen**_ if required
  + Create runspace for SplashScreen if needed, start splash is required as well
  + Load XAML for for tile WPF window
  + Find all 'NAMES' in XAML and set to SyncHash
  + Add all events for script
    + Close SplashScreen in _**tile_Window.Add_Loaded**_ if required
    + In _**tile_Window.Add_Closed**_ :
      + Enable tile in MainWindow
      + Remove any items from *$syncHash* that match tile name
  + Show the Window Dialog
+ Invoke the Parent runspace

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
