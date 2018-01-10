#Disable the tile when the runspace is open
$SyncHash.tile_DepartmentFolders.IsEnabled = $false 

$newRunspace =[runspacefactory]::CreateRunspace()
$newRunspace.ApartmentState = "STA"
$newRunspace.ThreadOptions = "ReuseThread"          
$newRunspace.Open()
$newRunspace.SessionStateProxy.SetVariable("SyncHash",$SyncHash) 
$PowerShell = [PowerShell]::Create().AddScript({
    
    #region FUNCTIONS        
    #region Update Window Function
    Function Update-Window {
        Param (
            $Control,
            $Property,
            $Value,
            [switch]$AppendContent
        )

        # This is kind of a hack, there may be a better way to do this
        If ($Property -eq "Close") {
            $syncHash.Window.Dispatcher.invoke([action]{$syncHash.Window.Close()},"Normal")
            Return
        }

        # This updates the control based on the parameters passed to the function
        $syncHash.$Control.Dispatcher.Invoke([action]{
            # This bit is only really meaningful for the TextBox control, which might be useful for logging progress steps
            If ($PSBoundParameters['AppendContent']) {
                $syncHash.$Control.AppendText($Value)
            } 
            Else {
                $syncHash.$Control.$Property = $Value
            }
        }, "Normal")
    }
    #endregion Update Window Function  
            
    function close-SplashScreen (){
        $hash.WindowSplash.Dispatcher.Invoke("Normal",[action]{ $hash.WindowSplash.close() })
        $Pwshell.EndInvoke($handle) | Out-Null
        #$runspace.Close() | Out-Null
    }

    function Start-SplashScreen{
        $Pwshell.Runspace = $runspace
        $script:handle = $Pwshell.BeginInvoke() 
    }

    #endregion FUNCTIONS

<#
    #region SplashScreen
    #Create a splash screen in a seperate runspace so it shows while the rest of the script is loading
    $hash = [hashtable]::Synchronized(@{})
    $hash.scriptRoot = $syncHash.scriptRoot  #Pass the base directory (scriptRoot) to the runspace
    $runspace = [runspacefactory]::CreateRunspace()
    $runspace.ApartmentState = "STA"
    $Runspace.ThreadOptions = "ReuseThread"
    $runspace.Open()
    $runspace.SessionStateProxy.SetVariable("hash",$hash) 
    $Pwshell = [PowerShell]::Create()

    $Pwshell.AddScript({
        #Load the Main XAML from file
        $xamlLoader = (New-Object System.Xml.XmlDocument)
        $xamlLoader.Load($hash.scriptRoot + "\resources\XML\Splash.xaml")

        $reader=(New-Object System.Xml.XmlNodeReader $xamlLoader) 
        $hash.WindowSplash = [Windows.Markup.XamlReader]::Load($reader)
	            
        $xamlLoader.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]") | ForEach-Object{
            #Find all of the form types and add them as members to the synchash
            $Hash.Add($_.Name,$Hash.WindowSplash.FindName($_.Name) )
        }
                               
        $hash.lblTitle.Content = "Loading Exchange Management Shell"
        $hash.LoadingLabel.Content = "Please Wait" 
        $hash.WindowSplash.ShowDialog() 
    
    }) | Out-Null

    #endregion SplashScreen

    #Start the splash screen while the rest of the script loads
    Start-SplashScreen
#>
 
    #Load the XAML from file
    $xamlLoader = (New-Object System.Xml.XmlDocument)
    $xamlLoader.Load($SyncHash.scriptRoot + "\resources\XML\DepartmentFolders.xaml")

    #Load the XAML and catch a failure
    $reader=(New-Object System.Xml.XmlNodeReader $xamlLoader) 
    $SyncHash.DepartmentFolders_Window = [Windows.Markup.XamlReader]::Load($reader)
	
    #[xml]$XAML = $xml
    $xamlLoader.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]") | ForEach-Object{
        #Find all of the form types and add them as members to the synchash
        $SyncHash.Add($_.Name,$SyncHash.DepartmentFolders_Window.FindName($_.Name) )
    }

    #region EVENTS
    
    #DepartmentFolders Window Loaded
    $syncHash.Departmentfolders_Window.Add_Loaded({
		$SyncHash.departmentfolders_Root =  "\\bafileadm01\departments"
        #$syncHash.DeaprtmentFolders.tb_PATH.Text = $SyncHash.departmentfolders_Root
        #$dirList = Get-ChildItem $SyncHash.departmentfolders_Root | Where-Object {$_.PSIsContainer}
	    #$dirList |sort Name| foreach {
		#    $syncHash.DepartmentFolders_treeView.Items.Add( $_)
        #}
    })

    #Open the Flyout when the toggle is switched on and disable the Grid to prevent multiple toggles being selected
    $toggles = $syncHash.Keys | Where-Object {$_ -like "departmentFolders_tog*"}
    $toggles | ForEach-Object {
        $syncHash.$_.add_Checked({
            $SyncHash.DepartmentFolders_FlyOutContent.Visibility = "Visible"
            $SyncHash.DepartmentFolders_gridSwitches.IsEnabled = $false 
        })
    }

    #Flyout is Opened
    $SyncHash.DepartmentFolders_FlyOutContent.Add_IsVisibleChanged({
        $SyncHash.DepartmentFolders_tb_Message.text = ""
        If ($SyncHash.DepartmentFolders_FlyOutContent.IsVisible -eq $true) {
            $SyncHash.DepartmentFolders_treeView.Items.Clear()
            $syncHash.DepartmentFolders_tb_PATH.Text = $SyncHash.Departmentfolders_Root
            $dirList = Get-ChildItem $SyncHash.departmentfolders_Root | Where-Object {$_.PSIsContainer}
	        $dirList |sort Name| foreach {
		        $syncHash.DepartmentFolders_treeView.Items.Add( $_)
            }
            $SyncHash.DepartmentFolders_treeView.IsEnabled = $true
            $SyncHash.DepartmentFolders_but_OK.IsEnabled = $false
        }

        if ($SyncHash.DepartmentFolders_tog2.IsChecked) {
            $SyncHash.DepartmentFolders_tb_Message.text = "Please Select Department folder where you would like to`ncreate the Self Service folder."
        }
    })

    $SyncHash.departmentFolders_treeView.add_MouseDoubleClick({
       

            $dir = $SyncHash.departmentFolders_treeView.SelectedItem.FullName
            $SyncHash.departmentFolders_tb_PATH.Text = $dir
            $SyncHash.departmentFolders_treeView.Items.Clear()
        
            #Add a dir UP button if not at the Directory root
    
            if ((Split-path $dir).Split('\').Count -gt 3) {
                $customObject = [PSCustomObject]@{
                    FullName = (Split-Path $dir)
                    Name = ".."
                }
                $SyncHash.departmentFolders_treeView.Items.Add($customObject)
            }

            $dirList = Get-ChildItem $dir | Where-Object {$_.PSIsContainer}
	        $dirList |sort Name| foreach {
		        #$observableCollection.Add( $_ )
                $SyncHash.departmentFolders_treeView.Items.Add( $_ )
	        }
	    
            if (($dir.Split('\').Count) -ge 6) {
                $SyncHash.DepartmentFolders_but_OK.IsEnabled = $true
                if ($SyncHash.DepartmentFolders_tog2.IsChecked) {
                    $SyncHash.DepartmentFolders_tb_Message.Text = "Self Service folders must be`ncreated at the root of the Department folder."
                }

            }
            
            else {$SyncHash.DepartmentFolders_but_OK.IsEnabled = $false}

    })

    #Flyout OK Button Clicked
    $SyncHash.DepartmentFolders_but_OK.add_Click({
        $SyncHash.DepartmentFolders_treeView.IsEnabled = $false
        #If the selected path is greater than 6 shorten the path
        $path = $SyncHash.DepartmentFolders_tb_PATH.Text
        if (($path).Split('\\').Count -gt 6) {
            do {
                $path = Split-Path $path
            }
            until (($path).Split('\\').Count -eq 6)
            $SyncHash.DepartmentFolders_tb_Message.text = "Self-Service is only Supported on the First Set of Subfolders`n in a Department Folder. Path has been adjusted."
        }
        $SyncHash.DepartmentFolders_tb_PATH.Text = $path
        (Get-Acl $SyncHash.DepartmentFolders_tb_PATH.Text).Access.IdentityReference | ForEach-Object {
            if (($_.ToString()).Contains((Split-Path $SyncHash.DepartmentFolders_tb_PATH.Text -Leaf))) {
                if ($_ -like "*-MOD") {
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
        }

    })


    #Flyout RESET Button Clicked
    $SyncHash.DepartmentFolders_but_Reset.add_Click({
        $syncHash.DepartmentFolders_treeView.Items.Clear()
        $DepartmentFoldersTBitems = $syncHash.Keys | Where-Object {$_ -like "DepartmentFolders_tb_*"} 
        $DepartmentFoldersTBitems | % {$syncHash.$_.Text = ""}
        $syncHash.DepartmentFolders_tb_PATH.Text = $SyncHash.Departmentfolders_Root
        $dirList = Get-ChildItem $SyncHash.departmentfolders_Root | Where-Object {$_.PSIsContainer}
	    $dirList |sort Name| foreach {
		    $syncHash.DepartmentFolders_treeView.Items.Add( $_)
        }
        $SyncHash.DepartmentFolders_treeView.IsEnabled = $true
        $SyncHash.DepartmentFolders_but_OK.IsEnabled = $false

    })


    #Flyout Close button clicked
    $SyncHash.DepartmentFolders_but_flClose.add_Click({
        $SyncHash.DepartmentFolders_but_Reset.RaiseEvent((New-Object -TypeName System.Windows.RoutedEventArgs -ArgumentList $([System.Windows.Controls.Button]::ClickEvent)))
        $SyncHash.DepartmentFolders_gridSwitches.IsEnabled = $True
        $SyncHash.DepartmentFolders_FlyOutContent.Visibility = "Collapsed"
                
        $DepartmentFoldersToggles = $syncHash.Keys | Where-Object {$_ -like "DepartmentFolders_tog*"} 
        $DepartmentFoldersToggles | % {$syncHash.$_.IsChecked = $false}
                
        #$SyncHash.DepartmentFolders_tog1.IsEnabled = $True
        #$SyncHash.DepartmentFolders_tog1.IsChecked = $false

        #Cleanup any items that may have been populated
        $syncHash.DepartmentFolders_tb_UPN.Text = ""
        $syncHash.DepartmentFolders_TreeView.Items.Clear()
    })

    #DepartmentFolders Window Closed
    $SyncHash.DepartmentFolders_Window.Add_Closed({
        #Re-enable the tile when the Runspace closes
        Update-Window -control tile_DepartmentFolders -property IsEnabled -value $true 
        $DepartmentFoldersitems = $syncHash.Keys | Where-Object {$_ -like "DepartmentFolders*"} 
        $DepartmentFoldersitems | % {$syncHash.Remove($_)}
    })

    #endregion EVENTS

    $SyncHash.DepartmentFolders_Window.ShowDialog() 
    $syncHash.Error = $Error  
})

$PowerShell.Runspace = $newRunspace
[void]$Jobs.Add((
    [pscustomobject]@{
        PowerShell = $PowerShell
        Runspace = $PowerShell.BeginInvoke()
    }
))