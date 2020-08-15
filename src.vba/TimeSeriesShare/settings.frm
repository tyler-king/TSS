Private Sub CommandButton1_Click()
tssSettings.insertTemplate "save"
Unload Me
End Sub

Private Sub CommandButton2_Click()
 Unload Me
End Sub

Private Sub CommandButton3_Click()
tssSettings.insertTemplate "refresh"
Unload Me
End Sub

Private Sub CommandButton4_Click()
    Dim odict As Scripting.Dictionary
    Set odict = New Scripting.Dictionary
    With Me
        odict.Add SETTING_DB_MISSING, .dbMissing.Value
        odict.Add SETTING_TS_DEF_MISSING, .scMissing.Value
        If .valMissing.Value = "" Then
            odict.Add SETTING_VALUE_MISSING, Null
        Else
            odict.Add SETTING_VALUE_MISSING, .valMissing.Value
        End If
        odict.Add SETTING_SHOW_SAVE_LOG, .displayLog.Value
        odict.Add SETTING_DIF_HIGHLIGHT, .highlightChanges.Value
    End With
    saveUserSettings odict
    Unload Me
End Sub


Private Sub UserForm_Initialize()
    Set osettings = getUserSettings
    With Me
        .dbMissing.Value = osettings.Item(SETTING_DB_MISSING)
        .scMissing.Value = osettings.Item(SETTING_TS_DEF_MISSING)
        .valMissing.Value = osettings.Item(SETTING_VALUE_MISSING)
        .displayLog.Value = osettings.Item(SETTING_SHOW_SAVE_LOG)
        .highlightChanges.Value = osettings.Item(SETTING_DIF_HIGHLIGHT)
        For Each ctrl In .Controls
            If ctrl.name Like "*CommandButton*" Then
                 ctrl.BackColor = RGB(143, 179, 203)
                 ctrl.ForeColor = vbWhite
                 ctrl.Font.Bold = True
            End If
        Next ctrl
    End With

End Sub