Private colStrings As Collection
Private strDelim As String

Public Property Get Text() As String
    Dim lngNext As Long
    Dim strStrings() As String
    Dim varItem As Variant
    Dim strText As String
    
    Select Case colStrings.Count
        Case 0
            Text = vbNullString
        Case 1
            Text = colStrings(1)
        Case Else
            ReDim strStrings(1 To colStrings.Count)
            For Each varItem In colStrings
                lngNext = lngNext + 1
                strStrings(lngNext) = varItem
            Next
            Set colStrings = Nothing
            Set colStrings = New Collection
            strText = Join(strStrings, strDelim)
            colStrings.Add strText
            Text = strText
    End Select
End Property

Public Property Let Text(ByVal strNewValue As String)
    Set colStrings = Nothing
    Set colStrings = New Collection
    colStrings.Add strNewValue
End Property

Public Function Add(ByVal strNewValue As String)
    colStrings.Add strNewValue
End Function

Public Function Count() As Long
    Count = colStrings.Count
End Function

Private Sub Class_Initialize()
    Set colStrings = New Collection
End Sub

Private Sub Class_Terminate()
    Set colStrings = Nothing
End Sub

Public Property Get Delim() As String
    Delim = strDelim
End Property

Public Property Let Delim(ByVal strNewValue As String)
    strDelim = strNewValue
End Property