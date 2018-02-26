Function Update-Window {
    Param (
        $Control,
        $Property,
        $Value,
        [switch]$AppendContent
    )

    # This is kind of a hack, there may be a better way to do this

    If ($Property -eq "Close") {
        $syncHash.Window.Dispatcher.invoke([action] {$syncHash.Window.Close()}, "Normal")
        Return
    }

    # This updates the control based on the parameters passed to the function
    $syncHash.$Control.Dispatcher.Invoke([action] {
        # This bit is only really meaningful for the TextBox control, which might be useful for logging progress steps
        If ($PSBoundParameters['AppendContent']) {
            $syncHash.$Control.AppendText($Value)
        } 
        Else {
            $syncHash.$Control.$Property = $Value
        }
    }, "Normal")
}