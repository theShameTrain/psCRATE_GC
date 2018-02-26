#Disable the tile when the runspace is open
$SyncHash.("tile_" + $SyncHash.tileClicked).IsEnabled = $false 

#Create a variable for the path to the tile
$tilePath = Join-Path -Path $syncHash.scriptRoot -ChildPath ("\resources\Tiles\" + $tileName + "\")

#Load the Tile Config.xml
#for testing use external config file
[xml]$tileConfig = Get-Content ("c:\config\$tileName\Config.xml")
#unremark the following line to enable config in resources
#[xml]$tileConfig = Get-Content ($tilePath + "config.xml")

#Add the AzureAD module (msonline) to the runspace
Import-Module "msonline"

#Get the connection credentials
If (!$SyncHash.cred) {
    $key = Get-Content ($SyncHash.config.Settings.general.keyStore + "\AESkey.aes")
    $SyncHash.msolCred = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $tileConfig.settings.vars.msolUser, ($tileConfig.settings.vars.encPassword | ConvertTo-SecureString -Key $key)
}

#Connect to AzureAD
Connect-MsolService -Credential $SyncHash.msolCred

$newRunspace = [runspacefactory]::CreateRunspace()
$newRunspace.ApartmentState = "STA"
$newRunspace.ThreadOptions = "ReuseThread"          
$newRunspace.Open()
$newRunspace.SessionStateProxy.SetVariable("SyncHash", $SyncHash) 
$PowerShell = [PowerShell]::Create().AddScript( {
    
        #region FUNCTIONS
        
        #Import Functions from Modules
        Get-ChildItem ($SyncHash.scriptRoot + "\resources\MODULES" ) | ForEach-Object {Import-Module -Name $_.FullName}
        #endregion FUNCTIONS

        #region SplashScreen
        #Create a splash screen in a seperate runspace so it shows while the rest of the script is loading
        $hash = [hashtable]::Synchronized(@{})
        $hash.scriptRoot = $syncHash.scriptRoot  #Pass the base directory (scriptRoot) to the runspace
        $runspace = [runspacefactory]::CreateRunspace()
        $runspace.ApartmentState = "STA"
        $Runspace.ThreadOptions = "ReuseThread"
        $runspace.Open()
        $runspace.SessionStateProxy.SetVariable("hash", $hash) 
        $Pwshell = [PowerShell]::Create()

        $Pwshell.AddScript( {
                #Load the Main XAML from file
                $xamlLoader = (New-Object System.Xml.XmlDocument)
                $xamlLoader.Load($hash.scriptRoot + "\resources\XML\Splash.xaml")

                $reader = (New-Object System.Xml.XmlNodeReader $xamlLoader) 
                $hash.WindowSplash = [Windows.Markup.XamlReader]::Load($reader)
	            
                $xamlLoader.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]") | ForEach-Object {
                    #Find all of the form types and add them as members to the synchash
                    $Hash.Add($_.Name, $Hash.WindowSplash.FindName($_.Name) )
                }
                               
                $hash.lblTitle.Content = "Loading Exchange Management Shell"
                $hash.LoadingLabel.Content = "Please Wait" 
                $hash.WindowSplash.ShowDialog() 
    
            }) | Out-Null

        #endregion SplashScreen

        #Start the splash screen while the rest of the script loads
        Start-SplashScreen

        #Load the Exchange Online Modules
        $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $syncHash.msolCred -Authentication Basic -AllowRedirection
        Import-PSSession $Session -AllowClobber
            
        #Load the XAML from file and Replace the ACCENTcolor String from config.xml
        $configLoader = (New-Object System.Xml.XmlDocument)
        $configLoader.Load($SyncHash.scriptRoot + "\resources\TILES\" + $syncHash.tileClicked + "\" + "config.xml")
        $xamlText = Get-Content ($SyncHash.scriptRoot + "\resources\TILES\" + $syncHash.tileClicked + "\" + $syncHash.tileClicked + ".xaml")
        $xamlText = $xamlText -replace "ACCENTcolor", $configLoader.tile.accent

        $xamlLoader = (New-Object System.Xml.XmlDocument)
        $xamlLoader.LoadXml($xamlText)

        #Load the XAML and catch a failure
        $reader = (New-Object System.Xml.XmlNodeReader $xamlLoader) 
        $SyncHash.o365_Window = [Windows.Markup.XamlReader]::Load($reader)
	
        #[xml]$XAML = $xml
        $xamlLoader.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]") | ForEach-Object {
            #Find all of the form types and add them as members to the synchash
            $SyncHash.Add($_.Name, $SyncHash.o365_Window.FindName($_.Name) )
        }

        #region EVENTS
    
        #o365 Window Loaded
        $syncHash.o365_Window.Add_Loaded( { 
                close-SplashScreen
            })

        #Open the Flyout when the toggle is switched on and disable the Grid to prevent multiple toggles being selected
        $toggles = $syncHash.Keys | Where-Object {$_ -like "o365_tog*"}
        $toggles | ForEach-Object {
            $syncHash.$_.add_Checked( { 
                    $SyncHash.o365_FlyOutContent.Visibility = "Visible"
                    $SyncHash.o365_gridSwitches.IsEnabled = $false 
                })
        }

        #Textbox Key Input
        $syncHash.o365_tb_UPN.add_KeyDown( { #Check every time a key is pressed
                
                if ($_.Key -eq "Esc") {
                    #if the user presses Escape
                    $syncHash.o365_tb_UPN.Text = ""
                    $syncHash.o365_listResults.Items.Clear()
                }
                if (($_.Key -eq "Return")) {
                    $syncHash.o365_listResults.Items.Clear()
            
                    #Add email suffix if no '@' in supplied text 
                    if (!($syncHash.o365_tb_UPN.Text).Contains('@') -eq $True) {
                        $syncHash.o365_tb_UPN.Text = $syncHash.o365_tb_UPN.Text + $syncHash.config.Settings.o365.emailSuffix
                    }

                    if ($syncHash.o365_tog1.IsChecked) {
                        #GetMailBoxPermissions                                    
                
                        $results = Get-MailboxPermission -Identity $syncHash.o365_tb_UPN.Text 
                
                        if (!$results) {
                            $newItem = Select-Object -InputObject "" ID, ACCESS
                            $newItem.ACCESS = "USER NOT FOUND"
                            $syncHash.o365_listResults.Items.Add($newItem)
                        }
                    
                        else {
                            $results = $results | Where-Object {$_.IsInherited -eq $false} | Where-Object {($_.User -notlike "NT Authority\Self") -and ($_.User -notlike "S-1-5*")} | Select-Object User, AccessRights
                            if ($results -eq $null) {
                                $newItem = Select-Object -InputObject "" ID, ACCESS
                                $newItem.ACCESS = "No Special Permissions Found"
                                $syncHash.o365_listResults.Items.Add($newItem)
                            }
                            else {
                                #$syncHash.results = $results
                                $results | ForEach-Object {
                                    $newItem = Select-Object -InputObject "" ID, ACCESS
                                    $newItem.ID = $_.User
                                    $newItem.ACCESS = $_.AccessRights
                                    $syncHash.o365_listResults.Items.Add($newItem)
                                }
                            }
                        }
                    }

                    if ($syncHash.o365_tog2.IsChecked) {
                        #GetSendAsPermissions
                        $results = Get-RecipientPermission -Identity $syncHash.o365_tb_UPN.Text 
                
                        if (!$results) {
                            $newItem = Select-Object -InputObject "" ID, ACCESS
                            $newItem.ACCESS = "USER NOT FOUND"
                            $syncHash.o365_listResults.Items.Add($newItem)
                        }

                        else {
                            $results = $results | Select-Object Trustee, AccessRights
                            if ($results.Trustee.Count -le 1) {
                                $newItem = Select-Object -InputObject "" ID, ACCESS
                                $newItem.ACCESS = "No Special Permissions Found"
                                $syncHash.o365_listResults.Items.Add($newItem)
                            }
                            else {
                                $results | ForEach-Object {
                                    $newItem = Select-Object -InputObject "" ID, ACCESS
                                    $newItem.ID = $_.Trustee
                                    $newItem.ACCESS = $_.AccessRights
                                    $syncHash.o365_listResults.Items.Add($newItem)
                                }
                            }
                        }
                    }

                    if ($syncHash.o365_tog3.IsChecked) {
                        #Get SendOnBehalfPermissions
                
                        $results = Get-Mailbox -Identity $syncHash.o365_tb_UPN.Text 

                        if (!$results) {
                            $newItem = Select-Object -InputObject "" ID, ACCESS
                            $newItem.ACCESS = "USER NOT FOUND"
                            $syncHash.o365_listResults.Items.Add($newItem)
                        }

                        else {
                            $results = $results | Select-Object GrantSendOnBehalfTo
                            if ($results.GrantSendOnBehalfTo.Count -eq 0) {
                                $newItem = Select-Object -InputObject "" ID, ACCESS
                                $newItem.ACCESS = "No Special Permissions Found"
                                $syncHash.o365_listResults.Items.Add($newItem)
                            }
                            else {
                                $results.GrantSendOnBehalfTo | ForEach-Object {
                                    $newItem = Select-Object -InputObject "" ID, ACCESS
                                    $newItem.ID = $_
                                    $newItem.ACCESS = "Send On Behalf"
                                    $syncHash.o365_listResults.Items.Add($newItem)
                                }
                            }
                        }
                    }

                    if ($syncHash.o365_tog4.IsChecked) {
                        #Get Calendar Permissions
                        $results = Get-MailboxFolderPermission -Identity ($syncHash.o365_tb_UPN.Text + ":\Calendar")

                        if (!$results) {
                            $newItem = Select-Object -InputObject "" ID, ACCESS
                            $newItem.ACCESS = "USER NOT FOUND"
                            $syncHash.o365_listResults.Items.Add($newItem)
                        }

                        else {
                            $results = $results | Where-Object {$_.AccessRights -ne "None"} | Where-Object {$_.User -notlike "NT:S-1-5*"} | Select-Object User, AccessRights  #Filter out system users 
                            if ($results.User.Count -eq 0) {
                                $newItem = Select-Object -InputObject "" ID, ACCESS
                                $newItem.ACCESS = "No Special Permissions Found"
                                $syncHash.o365_listResults.Items.Add($newItem)
                            }
                            else {
                                $results | ForEach-Object {
                                    $newItem = Select-Object -InputObject "" ID, ACCESS
                                    $newItem.ID = $_.User
                                    $newItem.ACCESS = $_.AccessRights
                                    $syncHash.o365_listResults.Items.Add($newItem)
                                }
                            }
                        }
                    }

                }
            })

        #Flyout Close button clicked
        $SyncHash.o365_but_flClose.add_Click( {
                $SyncHash.o365_gridSwitches.IsEnabled = $True
                $SyncHash.o365_FlyOutContent.Visibility = "Collapsed"
                
                $o365Toggles = $syncHash.Keys | Where-Object {$_ -like "o365_tog*"} 
                $o365Toggles | ForEach-Object {$syncHash.$_.IsChecked = $false}
                
                #$SyncHash.o365_tog1.IsEnabled = $True
                #$SyncHash.o365_tog1.IsChecked = $false

                #Cleanup any items that may have been populated
                $syncHash.o365_tb_UPN.Text = ""
                $syncHash.o365_listResults.Items.Clear()
            })

        #o365 Window Closed
        $SyncHash.o365_Window.Add_Closed( {
                #Re-enable the tile when the Runspace closes
                Update-Window -control tile_o365 -property IsEnabled -value $true 
                $o365items = $syncHash.Keys | Where-Object {$_ -like "o365*"} 
                $o365items | ForEach-Object {$syncHash.Remove($_)}
            })

        #endregion EVENTS

        $SyncHash.o365_Window.ShowDialog() 
        $syncHash.Error = $Error  
    })

$PowerShell.Runspace = $newRunspace
[void]$Jobs.Add((
        [pscustomobject]@{
            PowerShell = $PowerShell
            Runspace   = $PowerShell.BeginInvoke()
        }
    ))
        