#Disable the tile when the runspace is open
$SyncHash.("tile_" + $SyncHash.tileClicked).IsEnabled = $false 

$newRunspace =[runspacefactory]::CreateRunspace()
$newRunspace.ApartmentState = "STA"
$newRunspace.ThreadOptions = "ReuseThread"          
$newRunspace.Open()
$newRunspace.SessionStateProxy.SetVariable("SyncHash",$SyncHash) 
$PowerShell = [PowerShell]::Create().AddScript( {
        $tileName = $SyncHash.tileClicked
        #region FUNCTIONS        

        #Import Functions from Modules
        Get-ChildItem ($SyncHash.scriptRoot + "\resources\MODULES" ) | ForEach-Object {Import-Module -Name $_.FullName}
        
        #endregion FUNCTIONS
    
        #Create a variable for the path to the tile
        $tilePath = Join-Path -Path $syncHash.scriptRoot -ChildPath ("\resources\Tiles\" + $tileName + "\")
    
        #Load the Tile Config.xml
        #for testing use external config file
        [xml]$tileConfig = Get-Content ("c:\config\$tileName\Config.xml")
        #unremark the following line to enable config in resources
        #[xml]$tileConfig = Get-Content ($tilePath + "config.xml")

        #Load the XAML from file and Replace the ACCENTcolor String from config.xml
        $xamlText = Get-Content ($tilePath + $tileName + ".xaml")
        $xamlText = $xamlText -replace "ACCENTcolor", $tileConfig.settings.tile.accent

        $xamlLoader = (New-Object System.Xml.XmlDocument)
        $xamlLoader.LoadXml($xamlText)
    
        #Load the XAML
        $reader = (New-Object System.Xml.XmlNodeReader $xamlLoader) 
        $SyncHash.DepartmentFolders_Window = [Windows.Markup.XamlReader]::Load($reader)
	
        #Find all of the form types and add them as members to the synchash
        $xamlLoader.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]") | ForEach-Object {
            $SyncHash.Add($_.Name, $SyncHash.($tileName + "_Window").FindName($_.Name) )
        }

        #region EVENTS
    
        #Window Loaded
        $syncHash.($tileName + "_Window").Add_Loaded( {
                $SyncHash.departmentfolders_Root = $tileConfig.settings.vars.deptRoot
            } )

        #Open the Flyout when the toggle is switched on and disable the Grid to prevent multiple toggles being selected
        $toggles = $syncHash.Keys | Where-Object {$_ -like ($tileName + "_tog*")}
        $toggles | ForEach-Object {
            $syncHash.$_.add_Checked( {
                    $SyncHash.($tileName + "_FlyOutContent").Visibility = "Visible"
                    $SyncHash.($tileName + "_gridSwitches").IsEnabled = $false 
                } )
        }

        #Flyout is Opened
        $SyncHash.DepartmentFolders_FlyOutContent.Add_IsVisibleChanged( {
                $SyncHash.DepartmentFolders_tb_Message.text = ""
                If ($SyncHash.DepartmentFolders_FlyOutContent.IsVisible -eq $true) {
                    $SyncHash.DepartmentFolders_treeView.Items.Clear()
                    $syncHash.DepartmentFolders_tb_PATH.Text = $SyncHash.Departmentfolders_Root
                    $dirList = Get-ChildItem $SyncHash.departmentfolders_Root | Where-Object {$_.PSIsContainer}
                    $dirList |Sort-Object Name| ForEach-Object {
                        $syncHash.DepartmentFolders_treeView.Items.Add( $_)
                    }
                    $SyncHash.DepartmentFolders_treeView.IsEnabled = $true
                    $SyncHash.DepartmentFolders_but_OK.IsEnabled = $false
                }

                if ($SyncHash.DepartmentFolders_tog2.IsChecked) {
                    $SyncHash.DepartmentFolders_tb_Message.text = "Please Select Department folder where you would like to`ncreate the Self Service folder."
                }
            } )

        $SyncHash.departmentFolders_treeView.add_MouseDoubleClick( {   
                $dir = $SyncHash.departmentFolders_treeView.SelectedItem.FullName
                $SyncHash.departmentFolders_tb_PATH.Text = $dir
                $SyncHash.departmentFolders_treeView.Items.Clear()
        
                #Add a dir UP button if not at the Directory root (Prevents user from browsing up from root dir)
                #the number specified used is the total count of '\' in the UNC path to the root dir, because 
                #the UNC path starts with '\\' the count should be one more than the folder levels. 
                #Do not use a trailing '\' in the deptRoot config
    
                if (($dir.ToCharArray() | Where-Object {$_ -eq '\'}).Count -gt (($tileConfig.settings.vars.deptRoot).ToCharArray() | Where-Object {$_ -eq '\'}).Count) {
                    $customObject = [PSCustomObject]@{
                        FullName = (Split-Path $dir)
                        Name     = ".."
                    }
                    $SyncHash.departmentFolders_treeView.Items.Add($customObject)
                }

                $dirList = Get-ChildItem $dir | Where-Object {$_.PSIsContainer}
                $dirList |Sort-Object Name| ForEach-Object {
                    $SyncHash.departmentFolders_treeView.Items.Add( $_ )
                }
	    
                if (($dir.ToCharArray() | Where-Object {$_ -eq '\'}).Count -ge $tileConfig.settings.deptfolders.shareLevel) {
                    #Turn on the OK button when the user has selected a folder at the proper level (ie - \\Root\Share\Department\SHAREDFOLDER)
                    $SyncHash.DepartmentFolders_but_OK.IsEnabled = $true
                    if ($SyncHash.DepartmentFolders_tog2.IsChecked) {
                        $SyncHash.DepartmentFolders_tb_Message.Text = "Self Service folders must be`ncreated at the root of the Department folder."
                    }

                }
            
                else {$SyncHash.DepartmentFolders_but_OK.IsEnabled = $false}
    
            } )

        #Flyout OK Button Clicked
        $SyncHash.DepartmentFolders_but_OK.add_Click( {
                $SyncHash.DepartmentFolders_treeView.IsEnabled = $false

                #TOGGLE 1 was selected :  Get Self Service Folder
                if ($SyncHash.DepartmentFolders_tog1.IsChecked) { 
                    #If the selected path is greater than the shareLevel, adjust the path
                    $path = $SyncHash.DepartmentFolders_tb_PATH.Text
                    if (($path.ToCharArray() | Where-Object {$_ -eq '\'}).Count -gt $tileConfig.settings.vars.shareLevel) {
                        do {
                            $path = Split-Path $path
                        }
                        until (($path.ToCharArray() | Where-Object {$_ -eq '\'}).Count -eq $tileConfig.settings.vars.shareLevel)
                        $SyncHash.DepartmentFolders_tb_Message.text = "Self-Service is only Supported on the First Set of Subfolders`n in a Department Folder. Path has been adjusted."
                    }
                    $SyncHash.DepartmentFolders_tb_PATH.Text = $path
                    $check = 0
                    (Get-Acl $SyncHash.DepartmentFolders_tb_PATH.Text).Access.IdentityReference | ForEach-Object {
                        if (($_.ToString()).Contains((Split-Path $SyncHash.DepartmentFolders_tb_PATH.Text -Leaf))) {
                            if ($_ -like "*-MOD") {
                                $check++
                                $SyncHash.DepartmentFolders_tb_groupMOD.text = $_.ToString().Split('\')[1]
                                $modManager = (Get-QADGroup $_.ToString()).ManagedBy
                                if ($modManager.Count -gt 0) {
                                    $SyncHash.DepartmentFolders_tb_ManagerMOD.text = (Get-QADUser $modManager).DisplayName
                                }
                                else { 
                                    $SyncHash.DepartmentFolders_tb_ManagerMOD.text = "No Manager Specified"
                                }
                            }
                            if ($_ -like "*-RO") {
                                $check++
                                $SyncHash.DepartmentFolders_tb_groupRO.text = $_.ToString().Split('\')[1]
                                $roManager = (Get-QADGroup $_.ToString()).ManagedBy
                                if ($roManager.Count -gt 0) {
                                    $SyncHash.DepartmentFolders_tb_ManagerRO.text = (Get-QADUser $roManager).DisplayName
                                }
                                else { 
                                    $SyncHash.DepartmentFolders_tb_ManagerRO.text = "No Manager Specified"
                                }
                            }
                        }
                        $SyncHash.zCheck = $check
                    }
                    if (!$check -gt 0) {
                        $SyncHash.DepartmentFolders_tb_Message.Text = "NOT A SELF SERVICE FOLDER!"
                    }
                }

                #TOGGLE 2 was selected : Create Self Service Folder
                if ($SyncHash.DepartmentFolders_tog2.IsChecked) { 
                    #If the selected path is greater than the (shareLevel - 1), adjust the path : Self-Service Folders need to be created in the root of the Department folder
                    $path = $SyncHash.DepartmentFolders_tb_PATH.Text
                    if (($path.ToCharArray() | Where-Object {$_ -eq '\'}).Count -gt ($tileConfig.settings.vars.shareLevel - 1)) {
                        do {
                            $path = Split-Path $path
                        }
                        until (($path.ToCharArray() | Where-Object {$_ -eq '\'}).Count -eq ($tileConfig.settings.vars.shareLevel - 1))
                        $SyncHash.DepartmentFolders_tb_Message.text = "Self-Service is only Supported in the root of a Department Folder.`n Path has been adjusted."
                    }
                    $SyncHash.DepartmentFolders_tb_PATH.Text = $path

                    #Make sure that a mapped drive to departments exists
                    if (!(Get-PSDrive -Name ($tileConfig.settings.vars.mapDrive).Replace(':', '') -ErrorAction SilentlyContinue)) {
                        New-PSDrive -Name ($tileConfig.settings.vars.mapDrive).Replace(':', '') -PSProvider FileSystem -Root $tileConfig.settings.vars.deptRoot -Persist
                    }

                    $check = 0
                    (Get-Acl $SyncHash.DepartmentFolders_tb_PATH.Text).Access.IdentityReference | ForEach-Object {
                        if (($_.ToString()).Contains((Split-Path $SyncHash.DepartmentFolders_tb_PATH.Text -Leaf))) {
                            if ($_ -like "*-MOD") {
                                $check++
                                $SyncHash.DepartmentFolders_tb_groupMOD.text = $_.ToString().Split('\')[1]
                                $modManager = (Get-QADGroup $_.ToString()).ManagedBy
                                if ($modManager.Count -gt 0) {
                                    $SyncHash.DepartmentFolders_tb_ManagerMOD.text = (Get-QADUser $modManager).DisplayName
                                }
                                else { 
                                    $SyncHash.DepartmentFolders_tb_ManagerMOD.text = "No Manager Specified"
                                }
                            }
                            if ($_ -like "*-RO") {
                                $check++
                                $SyncHash.DepartmentFolders_tb_groupRO.text = $_.ToString().Split('\')[1]
                                $roManager = (Get-QADGroup $_.ToString()).ManagedBy
                                if ($roManager.Count -gt 0) {
                                    $SyncHash.DepartmentFolders_tb_ManagerRO.text = (Get-QADUser $roManager).DisplayName
                                }
                                else { 
                                    $SyncHash.DepartmentFolders_tb_ManagerRO.text = "No Manager Specified"
                                }
                            }
                        }
                        $SyncHash.zCheck = $check
                    }
                    #if (!$check -gt 0) {
                    #     $SyncHash.DepartmentFolders_tb_Message.Text = "NOT A SELF SERVICE FOLDER!"
                    #}
                }
            } )


        #Flyout RESET Button Clicked
        $SyncHash.DepartmentFolders_but_Reset.add_Click( {
                $syncHash.DepartmentFolders_treeView.Items.Clear()
                $DepartmentFoldersTBitems = $syncHash.Keys | Where-Object {$_ -like "DepartmentFolders_tb_*"} 
                $DepartmentFoldersTBitems | ForEach-Object {$syncHash.$_.Text = ""}
                $syncHash.DepartmentFolders_tb_PATH.Text = $SyncHash.Departmentfolders_Root
                $dirList = Get-ChildItem $SyncHash.departmentfolders_Root | Where-Object {$_.PSIsContainer}
                $dirList |Sort-Object Name| ForEach-Object {
                    $syncHash.DepartmentFolders_treeView.Items.Add( $_)
                }
                $SyncHash.DepartmentFolders_treeView.IsEnabled = $true
                $SyncHash.DepartmentFolders_but_OK.IsEnabled = $false

            } )


        #Flyout Close button clicked
        $SyncHash.($tileName + "_but_flClose").add_Click( {
                $SyncHash.($tileName + "_but_Reset").RaiseEvent((New-Object -TypeName System.Windows.RoutedEventArgs -ArgumentList $([System.Windows.Controls.Button]::ClickEvent)))
                $SyncHash.($tileName + "_gridSwitches").IsEnabled = $True
                $SyncHash.($tileName + "_FlyOutContent").Visibility = "Collapsed"
                
                $tileToggles = $syncHash.Keys | Where-Object {$_ -like ($tileName + "_tog*")} 
                $tileToggles | ForEach-Object {$syncHash.$_.IsChecked = $false}
        
                #Cleanup any items that may have been populated
                $syncHash.DepartmentFolders_tb_UPN.Text = ""
                $syncHash.DepartmentFolders_TreeView.Items.Clear()
            } )

        #Window Closed
        $SyncHash.($tileName + "_Window").Add_Closed( {
                #Re-enable the tile when the Runspace closes
                Update-Window -control ("tile_" + $tileName) -property IsEnabled -value $true 
                $tileItems = $syncHash.Keys | Where-Object {$_ -like ($tileName + "*")} 
                $tileItems | ForEach-Object {$syncHash.Remove($_)}
            } )

        #endregion EVENTS


        #Show the Window 
        $SyncHash.($tileName + "_Window").ShowDialog() 
        $syncHash.Error = $Error  
    } )

$PowerShell.Runspace = $newRunspace
[void]$Jobs.Add((
    [pscustomobject]@{
        PowerShell = $PowerShell
        Runspace = $PowerShell.BeginInvoke()
    }
))