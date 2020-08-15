Option Private Module

Public Const APP_NAME = "TSS"

Public Const SETTING_DB_MISSING = "DB_MISSING"
Public Const SETTING_TS_DEF_MISSING = "TS_DEF_MISSING"
Public Const SETTING_VALUE_MISSING = "VALUE_MISSING"
Public Const SETTING_DIF_HIGHLIGHT = "DIF_HIGHLIGHT"
Public Const SETTING_SHOW_SAVE_LOG = "SHOW_SAVE_LOG"

Public Function getUserSettings() As Scripting.Dictionary
    Set getUserSettings = getSettingsFile
    Set dDefault = defaultUserSettings()
    If getUserSettings Is Nothing Then
        Set getUserSettings = dDefault
    Else
        For Each O In dDefault.keys
            If Not getUserSettings.Exists(O) Then
                getUserSettings.Add O, dDefault.Item(O)
            End If
        Next O
    End If
End Function

Public Function defaultUserSettings() As Scripting.Dictionary
    Set defaultUserSettings = New Scripting.Dictionary
    defaultUserSettings.Add SETTING_DB_MISSING, "#N/A:PATH"
    defaultUserSettings.Add SETTING_TS_DEF_MISSING, "#N/A:CODE"
    defaultUserSettings.Add SETTING_VALUE_MISSING, Null
    defaultUserSettings.Add SETTING_SHOW_SAVE_LOG, False
    defaultUserSettings.Add SETTING_DIF_HIGHLIGHT, RGB(255, 0, 0)
End Function


Public Sub insertTemplate(Optional ttype As String = "save")
Dim ts As New tssWS
If Application.Workbooks.Count > 0 Then
    Set ws = ActiveWorkbook.Sheets.Add
    Dim O As Variant
    If ttype = "save" Then
        O = ts.exampleSaveSheet
    Else
        O = ts.exampleRetrieveSheet
    End If
    Set nRng = rangeInput(ws.Range("A1"), O)
    nRng.EntireColumn.AutoFit
End If

End Sub


Public Sub saveUserSettings(odict As Scripting.Dictionary)
Dim fso As FileSystemObject, ssettings As Scripting.Dictionary
Set fso = New FileSystemObject

Set ssettings = getUserSettings
If Not ssettings Is Nothing Then
    For Each k In odict.keys
        If ssettings.Exists(k) Then
            ssettings.Item(k) = odict.Item(k)
        End If
    Next k
    If Not setSettingsFile(ssettings) Then
        MsgBox "Couldn't save user settings"
    End If
Else
    MsgBox "Can't save user settings. Can't find AppData folder"
End If

End Sub

Public Function getSettingsFile() As Scripting.Dictionary
Dim fso As FileSystemObject
Set fso = New FileSystemObject
Set getSettingsFile = Nothing
If Environ("appdata") <> "" Then
    myData = Environ("appdata") & "\" & APP_NAME
    If Not fso.FolderExists(myData) Then
        fso.CreateFolder myData
    End If
    If fso.FolderExists(myData) Then
        Set fold = fso.GetFolder(myData)
        If Not fso.FileExists(myData & "\settings_vba.json") Then
            fso.CreateTextFile myData & "\settings_vba.json", True
            Set getSettingsFile = New Scripting.Dictionary
        Else
            On Error Resume Next
            Dim txt As String
            txt = fso.OpenTextFile(myData & "\settings_vba.json").ReadAll
            If Len(txt) > 0 Then
                Set getSettingsFile = JSONConverter.ParseJson(txt)
            Else
                Set getSettingsFile = New Scripting.Dictionary
            End If
        End If
    End If
End If

End Function

Public Function setSettingsFile(odict As Scripting.Dictionary) As Boolean
Dim fso As FileSystemObject
Set fso = New FileSystemObject
setSettingsFile = False
If Environ("appdata") <> "" Then
    myData = Environ("appdata") & "\" & APP_NAME
    If Not fso.FolderExists(myData) Then
        fso.CreateFolder myData
    End If
    If fso.FolderExists(myData) Then
        fso.OpenTextFile(myData & "\settings_vba.json", ForWriting, True).Write (JSONConverter.ConvertToJson(odict))
        setSettingsFile = True
    End If
End If

End Function


