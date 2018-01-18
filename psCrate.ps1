###########################################################################
#
# NAME:      psCrate.ps1
#
# AUTHOR:  	 Shamus Berube
#
# COMMENT: 	 WPF Form running multiple runspaces allowing for multi-threaded like behaviour
#			 Runspace Overview - ROOT (This Script) -> PARENT (Main GUI) -> CHILD (Events)
#
# VERSION HISTORY:
#            1.0 dd/MM/yyyy - Initial release
#
###########################################################################

#Create the Parent SyncHash that each runspace can access
$Global:syncHash = [hashtable]::Synchronized(@{})

#Add some GLOBAL items to the syncHash 
$syncHash.scriptRoot = $PSScriptRoot #Pass the base directory (scriptRoot) to the runspace
#for testing use external config file
[xml]$syncHash.config = Get-Content ("c:\config\Config.xml")
#unremark the following line to enable config in resources
#[xml]$syncHash.config =  Get-Content ($syncHash.scriptRoot + "\resources\XML\Config.xml")

$newRunspace = [runspacefactory]::CreateRunspace()
$newRunspace.ApartmentState = "STA"
$newRunspace.ThreadOptions = "ReuseThread"
$newRunspace.Open()
$newRunspace.SessionStateProxy.SetVariable("syncHash", $syncHash)

#region Assemblies to Load
$assemblyList = @("PresentationFramework", "PresentationCore", "WindowsBase")
$assemblyList | Foreach-Object {Add-Type -AssemblyName $_}

# Mahapps Library Asemblies
Add-Type -Path ($syncHash.scriptRoot + "\assembly\MahApps.Metro.dll")
Add-Type -Path ($syncHash.scriptRoot + "\assembly\System.Windows.Interactivity.dll")
Add-Type -Path ($syncHash.scriptRoot + "\assembly\Mahapps.metro.iconpacks.dll")
#endregion Assemblies

#CODE set to add to the New Runspace
$parentCODE = [PowerShell]::Create().AddScript( {   
        #region XAML
        #Get the Main XAML from file
        $xamlLoader = (New-Object System.Xml.XmlDocument)
        $xamlLoader.Load($syncHash.scriptRoot + "\resources\XML\Mainform.xaml")

        #Load the XAML for each tile and add to $xamlLoader
        Get-ChildItem ($syncHash.scriptRoot + "\resources\TILES") | ForEach-Object {
            #Load the TILE Config.xml
            $tileLoader = (New-Object System.Xml.XmlDocument)
            $tileLoader.Load((Join-Path -Path $_.FullName -ChildPath "Config.xml"))
            
            $nsMahappsString = "clr-namespace:MahApps.Metro.Controls;assembly=MahApps.Metro"
            $nsPresentationString = "http://schemas.microsoft.com/winfx/2006/xaml/presentation"
            $accentString = "pack://application:,,,/MahApps.Metro;component/Styles/Accents/"

            $tileLoader.tile | ForEach-Object {
                #Add the TILE and tile attributes
                $tile = $xamlLoader.MetroWindow.Grid.WrapPanel.AppendChild($xamlLoader.CreateElement("Controls", "Tile", $nsMahappsString))
                $tile.SetAttribute("Name", ("tile_" + $_.Name ))
                $tile.SetAttribute("Width", $_.Width)
                $tile.SetAttribute("Height", $_.Height)
                $tile.SetAttribute("Foreground", "White")
                $tile.SetAttribute("Title", $_.Title)
                $tile.SetAttribute("TitleFontSize", "12")
                $tile.SetAttribute("FontWeight", "Light")
                
                #Set the TILE accent color (create XML TILE.TileResources.ResourceDictionary.ResourceDictionary.MergedDictionaries.ResourceDictionary)
                $tileResources = $tile.AppendChild($xamlLoader.CreateElement("Controls", "Tile.Resources", $nsMahappsString))
                $tileResourcesDict = $tileResources.AppendChild($xamlLoader.CreateElement("", "ResourceDictionary", $nsPresentationString))
                $tileResourceMergDict = $tileResourcesDict.AppendChild($xamlLoader.CreateElement("", "ResourceDictionary.MergedDictionaries", $nsPresentationString))
                $tileResourceMergDictRes = $tileResourceMergDict.AppendChild($xamlLoader.CreateElement("", "ResourceDictionary", $nsPresentationString))
                $tileResourceMergDictRes.SetAttribute("Source", ($accentString + $_.accent + ".xaml"))

                #Add an ICON to TILE using Rectange and Opacity Mask
                $tileRectangle = $tile.AppendChild($xamlLoader.CreateElement("", "Rectangle", $nsPresentationString))
                $tileRectangle.SetAttribute("Width", "50")
                $tileRectangle.SetAttribute("Height", "50")
                $tileRectangle.SetAttribute("Fill", "White")
                
                $tileRectangleOpMask = $tileRectangle.AppendChild($xamlLoader.CreateElement("", "Rectangle.OpacityMask", $nsPresentationString))
                $tileRectangleVisBrush = $tileRectangleOpMask.AppendChild($xamlLoader.CreateElement("", "VisualBrush", $nsPresentationString))
                $tileRectangleVisBrush.SetAttribute("Visual", ("{iconPacks:PackIcon" + $_.icon + "}"))
                $tileRectangleVisBrush.SetAttribute("Stretch", "Uniform")

            }
            
        }

        #Load the main XAML 
        $reader = (New-Object System.Xml.XmlNodeReader $xamlLoader)
        $syncHash.mainWindow = [Windows.Markup.XamlReader]::Load( $reader )


        #Select the 'Named' objects from XAML and add them to Syncronized Hash
        $xamlLoader.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]") | ForEach-Object {
            #Find all of the form types and add them as members to the synchash
            $syncHash.Add($_.Name, $syncHash.mainWindow.FindName($_.Name) )
        }


        #region Background runspace to clean up jobs
        $Global:JobCleanup = [hashtable]::Synchronized(@{})
        $Global:Jobs = [system.collections.arraylist]::Synchronized((New-Object System.Collections.ArrayList))

        $jobCleanup.Flag = $True
        $newRunspace = [runspacefactory]::CreateRunspace()
        $newRunspace.ApartmentState = "STA"
        $newRunspace.ThreadOptions = "ReuseThread"          
        $newRunspace.Open()        
        $newRunspace.SessionStateProxy.SetVariable("jobCleanup", $jobCleanup)     
        $newRunspace.SessionStateProxy.SetVariable("jobs", $jobs) 

        $jobCleanup.PowerShell = [PowerShell]::Create().AddScript( {
                #Routine to handle completed runspaces
                Do {    
                    Foreach ($runspace in $jobs) {            
                        If ($runspace.Runspace.isCompleted) {
                            [void]$runspace.powershell.EndInvoke($runspace.Runspace)
                            $runspace.powershell.dispose()
                            $runspace.Runspace = $null
                            $runspace.powershell = $null               
                        } 
                    }

                    #Clean out unused runspace jobs
                    $temphash = $jobs.clone()
                    $temphash | Where-Object {$_.runspace -eq $Null} | ForEach-Object {$jobs.remove($_)}        
                    Start-Sleep -Seconds 1     
                } 
                while ($jobCleanup.Flag)
            })

        $jobCleanup.PowerShell.Runspace = $newRunspace
        $jobCleanup.Thread = $jobCleanup.PowerShell.BeginInvoke()  
        #endregion Background runspace to clean up jobs

        #region EVENTS
        
        #Add the tile Click events which read from $scriptRoot\resources\TILES\<TILENAME>\<TILENAME>.add_Click.ps1     
        $syncHash.keys | Where-Object {$_ -like "tile*"} | ForEach-Object {
            $syncHash.($_.ToString()).add_Click( {
                    [System.Object]$sender = $args[0].Name.TrimStart("tile_")
                    Invoke-Expression (Get-Content ($syncHash.scriptRoot + "\resources\TILES\" + $sender + "\" + $sender + ".add_Click.ps1") -Raw )
                })
        }

        #Window Close Event
        $syncHash.mainWindow.Add_Closed( {
                $jobCleanup.Flag = $False

                #Stop all runspaces
                $jobCleanup.PowerShell.Dispose()
            })

        #endregion EVENTS

        #Show the GUI running in newRunspace
        $syncHash.mainWindow.ShowDialog() | Out-Null
        $syncHash.Error = $Error
    })

#Call the PARENT CODE BLOCK
$parentCODE.Runspace = $newRunspace
$data = $parentCODE.BeginInvoke()
#Wait for the main window to close then cleanup all the runspaces
While ($data.IsCompleted -ne $true) {Start-Sleep 0.5}
Get-Runspace | Where-Object {$_.RunspaceAvailability -ne "Busy"} | ForEach-Object {$_.Dispose()}