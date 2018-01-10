#Disable the tile when the runspace is open
$SyncHash.tile_o365.IsEnabled = $false 
        
#Add the AzureAD module (msonline) to the runspace
Import-Module "msonline"

#Get the connection credentials
If (!$SyncHash.cred) {

    $encPassword = "76492d1116743f0423413b16050a5345MgB8ADMANABsADkAVgBoAEoAdQBxAGwATQBTADYATwA1AFkAcABMAFoAdABBAFEAPQA9AHwAYwBlAGYAZgA1ADcAMQAxAGIAZAAwAGYAOABjAGIAOABmADkANgA0AGUANgA3ADUAZAA5AGQANgA5ADIAYQBlADcAMABmADkAYgA2ADAAOQBhADAAMgA0ADcAOQA0ADQANQBlAGQAMgBhAGEAOAAyAGQAYQA3ADMAOABjADUAOABjAGMAZABhADEAYwBmADYAMAA2AGYANQA3AGEAZQA1ADYAMgBjAGIAZQA0ADUAYwA2ADQAMwBiAGQANgBhADIA"
    $key = Get-Content ($SyncHash.keyStore + "\AESkey.aes")
    $SyncHash.msolCred = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList "shamus.berube@georgiancollege.onmicrosoft.com", ($encPassword | ConvertTo-SecureString -Key $key)
}

#Connect to AzureAD
Connect-MsolService -Credential $SyncHash.msolCred

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

    #Load the Exchange Online Modules
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $syncHash.msolCred -Authentication Basic -AllowRedirection
    Import-PSSession $Session -AllowClobber
            
    #Load the XAML from file
    $xamlLoader = (New-Object System.Xml.XmlDocument)
    $xamlLoader.Load($SyncHash.scriptRoot + "\resources\XML\o365.xaml")

    #Load the XAML and catch a failure
    $reader=(New-Object System.Xml.XmlNodeReader $xamlLoader) 
    $SyncHash.o365_Window = [Windows.Markup.XamlReader]::Load($reader)
	
    #[xml]$XAML = $xml
    $xamlLoader.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]") | ForEach-Object{
        #Find all of the form types and add them as members to the synchash
        $SyncHash.Add($_.Name,$SyncHash.o365_Window.FindName($_.Name) )
    }

    #region EVENTS
    
    #o365 Window Loaded
    $syncHash.o365_Window.Add_Loaded({
        close-SplashScreen
    })

    #Open the Flyout when the toggle is switched on and disable the Grid to prevent multiple toggles being selected
    $toggles = $syncHash.Keys | Where-Object {$_ -like "o365_tog*"}
    $toggles | ForEach-Object {
        $syncHash.$_.add_Checked({
            $SyncHash.o365_FlyOutContent.Visibility = "Visible"
            $SyncHash.o365_gridSwitches.IsEnabled = $false 
        })
    }

    #Textbox Key Input
    $syncHash.o365_tb_UPN.add_KeyDown({ #Check every time a key is pressed
                
        if ($_.Key -eq "Esc") { #if the user presses Escape
            $syncHash.o365_tb_UPN.Text = ""
            $syncHash.o365_listResults.Items.Clear()
        }
        if (($_.Key -eq "Return")) {
            $syncHash.o365_listResults.Items.Clear()
            
            #Add email suffix if no '@' in text 
            if (!($syncHash.o365_tb_UPN.Text).Contains('@') -eq $True) {
                $syncHash.o365_tb_UPN.Text = $syncHash.o365_tb_UPN.Text + "@georgiancollege.ca"
            }

            if ($syncHash.o365_tog1.IsChecked) {  #GetMailBoxPermissions                                    
                
                $results = Get-MailboxPermission -Identity $syncHash.o365_tb_UPN.Text 
                
                if (!$results) {
                    $newItem = Select-Object -InputObject "" ID,ACCESS
                    $newItem.ACCESS = "USER NOT FOUND"
                    $syncHash.o365_listResults.Items.Add($newItem)
                }
                    
                else {
                    $results = $results | Where-Object {$_.IsInherited -eq $false} | Where-Object {($_.User -notlike "NT Authority\Self") -and ($_.User -notlike "S-1-5*")} | Select-Object User,AccessRights
                    if ($results -eq $null) {
                        $newItem = Select-Object -InputObject "" ID,ACCESS
                        $newItem.ACCESS = "No Special Permissions Found"
                        $syncHash.o365_listResults.Items.Add($newItem)
                    }
                    else {
                        #$syncHash.results = $results
                        $results | ForEach-Object {
                            $newItem = Select-Object -InputObject "" ID,ACCESS
                            $newItem.ID = $_.User
                            $newItem.ACCESS = $_.AccessRights
                            $syncHash.o365_listResults.Items.Add($newItem)
                        }
                    }
                }
            }

            if ($syncHash.o365_tog2.IsChecked) {  #GetSendAsPermissions
                $results = Get-RecipientPermission -Identity $syncHash.o365_tb_UPN.Text 
                
                if (!$results) {
                    $newItem = Select-Object -InputObject "" ID,ACCESS
                    $newItem.ACCESS = "USER NOT FOUND"
                    $syncHash.o365_listResults.Items.Add($newItem)
                }

                else {
                    $results = $results | Select-Object Trustee,AccessRights
                    if ($results.Trustee.Count -le 1) {
                        $newItem = Select-Object -InputObject "" ID,ACCESS
                        $newItem.ACCESS = "No Special Permissions Found"
                        $syncHash.o365_listResults.Items.Add($newItem)
                    }
                    else {
                        $results | ForEach-Object {
                            $newItem = Select-Object -InputObject "" ID,ACCESS
                            $newItem.ID = $_.Trustee
                            $newItem.ACCESS = $_.AccessRights
                            $syncHash.o365_listResults.Items.Add($newItem)
                        }
                    }
                }
            }

            if ($syncHash.o365_tog3.IsChecked) {  #Get SendOnBehalfPermissions
                
                $results = Get-Mailbox -Identity $syncHash.o365_tb_UPN.Text 

                if (!$results) {
                    $newItem = Select-Object -InputObject "" ID,ACCESS
                    $newItem.ACCESS = "USER NOT FOUND"
                    $syncHash.o365_listResults.Items.Add($newItem)
                }

                else {
                    $results = $results | Select-Object GrantSendOnBehalfTo
                    if ($results.GrantSendOnBehalfTo.Count -eq 0) {
                        $newItem = Select-Object -InputObject "" ID,ACCESS
                        $newItem.ACCESS = "No Special Permissions Found"
                        $syncHash.o365_listResults.Items.Add($newItem)
                    }
                    else {
                        $results.GrantSendOnBehalfTo | ForEach-Object {
                            $newItem = Select-Object -InputObject "" ID,ACCESS
                            $newItem.ID = $_
                            $newItem.ACCESS = "Send On Behalf"
                            $syncHash.o365_listResults.Items.Add($newItem)
                        }
                    }
                }
            }

            if ($syncHash.o365_tog4.IsChecked) {  #Get Calendar Permissions
                
                $results = Get-MailboxFolderPermission -Identity ($syncHash.o365_tb_UPN.Text + ":\Calendar")

                if (!$results) {
                    $newItem = Select-Object -InputObject "" ID,ACCESS
                    $newItem.ACCESS = "USER NOT FOUND"
                    $syncHash.o365_listResults.Items.Add($newItem)
                }

                else {
                    $results = $results | Where-Object {$_.AccessRights -ne "None"} | Where-Object {$_.User -notlike "NT:S-1-5*"} | Select-Object User,AccessRights
                    if ($results.User.Count -eq 0) {
                        $newItem = Select-Object -InputObject "" ID,ACCESS
                        $newItem.ACCESS = "No Special Permissions Found"
                        $syncHash.o365_listResults.Items.Add($newItem)
                    }
                    else {
                        $results | ForEach-Object {
                            $newItem = Select-Object -InputObject "" ID,ACCESS
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
    $SyncHash.o365_but_flClose.add_Click({
        $SyncHash.o365_gridSwitches.IsEnabled = $True
        $SyncHash.o365_FlyOutContent.Visibility = "Collapsed"
                
        $o365Toggles = $syncHash.Keys | Where-Object {$_ -like "o365_tog*"} 
        $o365Toggles | % {$syncHash.$_.IsChecked = $false}
                
        #$SyncHash.o365_tog1.IsEnabled = $True
        #$SyncHash.o365_tog1.IsChecked = $false

        #Cleanup any items that may have been populated
        $syncHash.o365_tb_UPN.Text = ""
        $syncHash.o365_listResults.Items.Clear()
    })

    #o365 Window Closed
    $SyncHash.o365_Window.Add_Closed({
        #Re-enable the tile when the Runspace closes
        Update-Window -control tile_o365 -property IsEnabled -value $true 
        $o365items = $syncHash.Keys | Where-Object {$_ -like "o365*"} 
        $o365items | % {$syncHash.Remove($_)}
    })

    #endregion EVENTS

    $SyncHash.o365_Window.ShowDialog() 
    $syncHash.Error = $Error  
})

$PowerShell.Runspace = $newRunspace
[void]$Jobs.Add((
    [pscustomobject]@{
        PowerShell = $PowerShell
        Runspace = $PowerShell.BeginInvoke()
    }
))
        