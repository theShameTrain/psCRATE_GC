###########################################################################
#
# NAME:      psCrate.ps1
#
# AUTHOR:  	 Shamus Berube
#
# COMMENT: 	 WPF Form running multiple runspaces allowing for multi-threaded like behaviour
#			 Runspace Overview - ROOT (This Script) -> PARENT (Main GUI) -> CHILD (Events)
#
#
# ASSUMPTIONS: Microsoft Azure AD Module (msonline) is installed 
# 
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
[xml]$syncHash.config =  Get-Content ("c:\config\Config.xml")
#unremark the following line to enable config in resources
#[xml]$syncHash.config =  Get-Content ($syncHash.scriptRoot + "\resources\XML\Config.xml")

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
    $xamlLoader.Load($syncHash.scriptRoot + "\resources\XML\Mainform.xaml")
	
	#Load the XAML 
	$reader=(New-Object System.Xml.XmlNodeReader $xamlLoader)
    $syncHash.mainWindow=[Windows.Markup.XamlReader]::Load( $reader )
	
	#Select the 'Named' objects from XAML and add them to Syncronized Hash
#	[xml]$XAML = $xaml
    $xamlLoader.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]") | ForEach-Object{
        #Find all of the form types and add them as members to the synchash
        $syncHash.Add($_.Name,$syncHash.mainWindow.FindName($_.Name) )
    }
	

    #region Background runspace to clean up jobs
    $Global:JobCleanup = [hashtable]::Synchronized(@{})
    $Global:Jobs = [system.collections.arraylist]::Synchronized((New-Object System.Collections.ArrayList))
    
    $jobCleanup.Flag = $True
    $newRunspace =[runspacefactory]::CreateRunspace()
    $newRunspace.ApartmentState = "STA"
    $newRunspace.ThreadOptions = "ReuseThread"          
    $newRunspace.Open()        
    $newRunspace.SessionStateProxy.SetVariable("jobCleanup",$jobCleanup)     
    $newRunspace.SessionStateProxy.SetVariable("jobs",$jobs) 
    
    $jobCleanup.PowerShell = [PowerShell]::Create().AddScript({
        #Routine to handle completed runspaces
        Do {    
            Foreach($runspace in $jobs) {            
                If ($runspace.Runspace.isCompleted) {
                    [void]$runspace.powershell.EndInvoke($runspace.Runspace)
                    $runspace.powershell.dispose()
                    $runspace.Runspace = $null
                    $runspace.powershell = $null               
                } 
            }

            #Clean out unused runspace jobs
            $temphash = $jobs.clone()
            $temphash | Where {$_.runspace -eq $Null} | ForEach {$jobs.remove($_)}        
            Start-Sleep -Seconds 1     
        } 
        while ($jobCleanup.Flag)
    })

    $jobCleanup.PowerShell.Runspace = $newRunspace
    $jobCleanup.Thread = $jobCleanup.PowerShell.BeginInvoke()  
    #endregion Background runspace to clean up jobs

	#region EVENTS
    
    $syncHash.tile_o365.add_Click({
        Invoke-Expression ( Get-Content ( $syncHash.scriptRoot + "\resources\EVENTS\tile_o365.add_Click.ps1" ) -Raw )
    })

    $syncHash.tile_DepartmentFolders.add_Click({
        Invoke-Expression ( Get-Content ( $syncHash.scriptRoot + "\resources\EVENTS\tile_DepartmentFolders.add_Click.ps1" ) -Raw )
    })

    #Window Close 
    $syncHash.mainWindow.Add_Closed({
        Write-Verbose 'Halt runspace cleanup job processing'
        $jobCleanup.Flag = $False

        #Stop all runspaces
        $jobCleanup.PowerShell.Dispose()      
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
