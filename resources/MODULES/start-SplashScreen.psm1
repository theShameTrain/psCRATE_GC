function start-SplashScreen {
    $Pwshell.Runspace = $runspace
    $script:handle = $Pwshell.BeginInvoke() 
}