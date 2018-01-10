$Global:syncHash = [hashtable]::Synchronized(@{})
$syncHash.scriptRoot = $PSScriptRoot #Pass the base directory (scriptRoot) to the runspace
$newRunspace =[runspacefactory]::CreateRunspace()
$newRunspace.ApartmentState = "STA"
$newRunspace.ThreadOptions = "ReuseThread"
$newRunspace.Open()
$newRunspace.SessionStateProxy.SetVariable("syncHash",$syncHash)


#region Assemblies to Load
$assemblyList = @("PresentationFramework","PresentationCore","WindowsBase")
$assemblyList | foreach {Add-Type -AssemblyName $_}
# Mahapps Library Asemblies
Add-Type -Path "$PSScriptRoot\assembly\MahApps.Metro.dll"
Add-Type -Path "$PSScriptRoot\assembly\System.Windows.Interactivity.dll"
Add-Type -Path "$PSScriptRoot\assembly\Mahapps.metro.iconpacks.dll"
#endregion Assemblies


#CODE set to add to the New Runspace
$parentCODE = [PowerShell]::Create().AddScript({   
	#region XAML
    #Get the Main XAML from file
    $xamlLoader = (New-Object System.Xml.XmlDocument)
    $xamlLoader.Load($syncHash.scriptRoot + "\resources\XML\FolderBrowserDialog.xaml")
	
	#Load the XAML 
	$reader=(New-Object System.Xml.XmlNodeReader $xamlLoader)
    $syncHash.mainWindow=[Windows.Markup.XamlReader]::Load( $reader )
	
	#Select the 'Named' objects from XAML and add them to Syncronized Hash
#	[xml]$XAML = $xaml
    $xamlLoader.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]") | ForEach-Object{
        #Find all of the form types and add them as members to the synchash
        $syncHash.Add($_.Name,$syncHash.mainWindow.FindName($_.Name) )
    }
	

    

	#region EVENTS
    
    $syncHash.tile_o365.add_Click({
        Invoke-Expression ( Get-Content ( $syncHash.scriptRoot + "\resources\EVENTS\tile_o365.add_Click.ps1" ) -Raw )
    })

    $syncHash.tile_SharedFolders.add_Click({
        Invoke-Expression ( Get-Content ( $syncHash.scriptRoot + "\resources\EVENTS\tile_SharedFolders.add_Click.ps1" ) -Raw )
    })

    #Window Close 
    $syncHash.mainWindow.Add_Closed({
    
    })

	#endregion EVENTS

#>
	
	#Show the GUI running in newRunspace
	$syncHash.mainWindow.ShowDialog() | Out-Null
    $syncHash.Error = $Error
})

#Call the PARENT CODE BLOCK
$parentCODE.Runspace = $newRunspace
$data = $parentCODE.BeginInvoke()
#Wait for the main window to close then cleanup all the runspaces
While ($data.IsCompleted -ne $true) {Sleep 0.5}
Get-Runspace | Where-Object {$_.RunspaceAvailability -ne "Busy"} | % {$_.Dispose()}  
