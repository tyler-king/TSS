Public Sub refreshLocal(icontrol As IRibbonControl)
    Set dDefault = New clsCodePrep
    refreshTimeSeries
End Sub
Public Sub saveLocal(icontrol As IRibbonControl)
    
    Set dDefault = New clsCodePrep
    createTimeSeries
    
End Sub

Sub openLocalSettings(iribbon As IRibbonControl)
    settings.Show
End Sub

Sub openDatabaseSettings(iribbon As IRibbonControl)
    DatabaseSettings.Show
End Sub
