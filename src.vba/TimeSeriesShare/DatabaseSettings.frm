Private Sub TabStrip1_Change()

End Sub

Private Sub UserForm_Activate()
Dim fso As FileSystemObject, dbName As String

If TypeName(Selection) <> "Range" Then
    Unload Me
    MsgBox "Please select a database path"
    Exit Sub
End If

dbName = Selection.Value
If Len(dbName) = 0 Then
    Unload Me
    MsgBox "Choose a cell with a database path"
    Exit Sub
End If

Set fso = New FileSystemObject
If Not fso.FileExists(dbName & DB_EXTENSION) Then
    Unload Me
    MsgBox dbName & " does not exist"
    Exit Sub
End If
Me.Caption = "Database Settings: " & dbName

Dim ssettings As Scripting.Dictionary, toGet As Collection
Set toGet = New Collection
toGet.Add "settings"
Set ssettings = readJSON(dbName, "", toGet)

If Not ssettings.Exists("settings") Then
   ssettings.Add "settings", getLocalDatabaseSettings()
End If
Set ssettings.Item("settings") = getLocalDatabaseSettings(ssettings.Item("settings"))
ttop = 10
For Each okey In ssettings.Item("settings").keys
    Set octrl = Me.Controls.Add("Forms.Label.1", "Label" & okey)
    octrl.Top = ttop
    octrl.Left = 10
    octrl.Height = 12
    octrl.Width = 100
    ttop = ttop + 15
    Select Case okey
        Case SYSTEM_SETTING_TS_ROBUST
            octrl.Caption = okey & " : " & ssettings.Item("settings").Item(okey)
        Case SYSTEM_SETTING_METADATA_MAPPINGS
            octrl.Caption = "MetaData Mappings : " & "TEST"
            For Each metkey In ssettings.Item("settings").Item(okey).keys
                Set octrl = Me.Controls.Add("Forms.Label.1", "Label" & okey & metkey)
                octrl.Top = ttop
                octrl.Left = 10
                octrl.Height = 12
                octrl.Width = 100
                ttop = ttop + 15
                octrl.Caption = metkey & " : " & "TEST"
            Next metkey
        Case SYSTEM_SETTING_VERSION
            Me.Controls.Remove ("Label" & okey)
            ttop = ttop - 15
    End Select

Next okey


End Sub

