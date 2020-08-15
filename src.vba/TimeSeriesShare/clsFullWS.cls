Public WithEvents ThisWS As Excel.Worksheet
Public hiddenRows As Range
Public hiddenColumns As Range

Public Property Set ws(ByRef oWS As Excel.Worksheet)
    Set ThisWS = oWS
    Set hiddenColumns = Me.UnHideColumns(ThisWS.UsedRange)
    Set hiddenRows = Me.UnHideRows(ThisWS.UsedRange)
End Property
Private Sub Class_Terminate()
    HideRowsAndColumns Me.hiddenColumns, Me.hiddenRows
End Sub

Public Function UnHideColumns(ByVal rToSearch As Range) As Range
'this function unhides all columns in a WS

Dim C As Range
Dim rHiddenCols As Range

Set rHiddenCols = Nothing
For Each C In rToSearch.Columns
    If C.Hidden = True Then
        'unhide col
        C.Hidden = False

        'add to union Range
        If rHiddenCols Is Nothing Then
            Set rHiddenCols = C
        Else
            Set rHiddenCols = Application.Union(rHiddenCols, C)
        End If
    End If
Next C
Set UnHideColumns = rHiddenCols
End Function


Public Function UnHideRows(ByVal rToSearch As Range) As Range
'this function unhides all rows in a WS
Dim C As Range
Dim rHiddenRows As Range

Set rHiddenRows = Nothing
For Each C In rToSearch.Rows
    If C.Hidden = True Then
        'unhide col

        'add to union Range
        If rHiddenRows Is Nothing Then
            Set rHiddenRows = C
        Else
            Set rHiddenRows = Application.Union(rHiddenRows, C)
        End If
    End If
Next C
If Not rHiddenRows Is Nothing Then
    rHiddenRows.Rows.Hidden = False
End If
Set UnHideRows = rHiddenRows
End Function

Public Sub HideRowsAndColumns(ByVal rHiddenCols As Range, ByVal rHiddenRows As Range)
'After the DMX Refresh has been performed this procedure rehides all of the previously hidden columns and rows
Dim r As Range

If Not rHiddenCols Is Nothing Then
    For Each r In rHiddenCols.Columns
        r.Columns.Hidden = True
    Next r
End If

If Not rHiddenRows Is Nothing Then
    rHiddenRows.Rows.Hidden = True
End If
End Sub