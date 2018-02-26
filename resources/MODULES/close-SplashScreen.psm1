function close-SplashScreen ($windowName) {
    $hash.$windowName.Dispatcher.Invoke( "Normal", [action] { $hash.$windowName.close() } )
    $Pwshell.EndInvoke($handle) | Out-Null
}