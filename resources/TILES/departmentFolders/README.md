# DepartmantFolder TILE Overview

This tile allows regular users to create controlled access folders or view the manager of existing controlled access folders and associated groups.  The folders are accessed via membership in AD groups that are managed by a department user.  The script creates the folder and the associated groups and also assigns the primary manager for the groups.

### Settings
Settings are located in \resources\TILES\*<TILE>*\Config.xml
+ TILE
    + Settings related to the display of the tile in the main psCrate.ps1
+ VARS
    + deptRoot
        + The UNC path to the root of the department folders.  This is the folder where all the main department folders are created.  Each department folder must be manually created.
    + mapDrive
        + If users access share via mapped drive, use the drive letter to use in the description for the users.
    + shareLevel
        + Basically the number of "\" in the uncpath that a folder must be created in.  This is used to calculate if the path is too short or too deep and to display a ".." icon in the folder browser.
    + deptShortList
        + Name of file containing the department "OU" to "SHRTNAME" that gets loaded into a hash table.  This allows all groups for the department to have the same prefix.  All groups for department are created in the specified OU.
    + searchRoot
        + OU specifies where to start searching for groups.  If all department groups are limited to the same space use that OU to speed up searches.

### Dependencies
+ Active Roles Management Shell needs to be installed
+ Account that the script runs as must have rights to create groups
+ Account that the script runs as must have the ability to create folders in the Department Root
+ Account that the script runs as must have the ability to assign security permissions to the folders in the Department Root

### Tools

+ Get Self Service Folder Manager
+ Create Self Service Folder
