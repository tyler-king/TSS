Const CONSOLIDATE_ERROR = 37
Private param1 As String

Public Function apply(ovals As Scripting.Dictionary, args() As Variant) As Scripting.Dictionary
    Set apply = runConsolidate(ovals, args)
End Function

Public Function IsError(numb As Long) As Boolean
    IsError = False
    Select Case numb
        Case CONSOLIDATE_ERROR:
            IsError = True
    End Select
End Function

Public Function force(param As String)
    param1 = param 'TODO could make this an automatic time thing
End Function

Private Function runConsolidate(ovals As Scripting.Dictionary, args() As Variant) As Scripting.Dictionary
    Dim divisor As Long, div As utc
    Set runConsolidate = New Scripting.Dictionary
    If param1 = "" Then
        consol_method = args(0)
        consol_freq = args(1)
        consol_missing_freq = args(2)
    Else
        consol_method = param1
        consol_freq = args(0)
        consol_missing_freq = args(1)
    End If
    Set missing = New Scripting.Dictionary
    SortDictionaryByKey ovals  'dictionary had to be ordered prior for first/last
    For Each s In ovals.keys
        tStamp = ConvertUTCToFreq(s, CStr(consol_freq))
        If Not runConsolidate.Exists(tStamp) Then
            runConsolidate.Add tStamp, New Collection
        End If
        runConsolidate.Item(tStamp).Add ovals.Item(s)
        If consol_missing_freq <> "" Then
            If Not missing.Exists(tStamp) Then
                missing.Add tStamp, New Scripting.Dictionary
            End If
            mStamp = ConvertUTCToFreq(s, CStr(consol_missing_freq))
            If Not missing.Item(tStamp).Exists(mStamp) Then
                 missing.Item(tStamp).Add mStamp, New Collection
            End If
            missing.Item(tStamp).Item(mStamp).Add ovals.Item(s)
        End If
    Next s
    Set div = New utc
    For Each t In runConsolidate
        If consol_missing_freq <> "" Then
            stillrun = False
            div.exactUTC = t
            divisor = div.periodicity(CStr(consol_freq), CStr(consol_missing_freq))
            If missing.Exists(t) Then
                If missing.Item(t).Count = divisor Then
                    stillrun = True
                End If
            End If
        Else
             stillrun = True
        End If
        If stillrun Then
            runConsolidate.Item(t) = consolidate(runConsolidate.Item(t), CStr(consol_method))
        Else
            runConsolidate.Remove t
        End If
    Next t
End Function

Private Function consolidate(n As Collection, method As String) As Variant
    Dim l As Long, vals As Variant
    consolidate = ""
    l = n.Count
    If l > 0 Then
        vals = collectionToVariant(n, 2)
        On Error GoTo finish:
        Select Case LCase(method)
            Case "average"
                consolidate = WorksheetFunction.Average(vals)
            Case "sum"
                consolidate = WorksheetFunction.Sum(vals)
            Case "last"
                consolidate = n(n.Count)
            Case "first"
                consolidate = n(1)
            Case "count"
                consolidate = l
            Case "max"
                consolidate = WorksheetFunction.Max(vals)
            Case "min"
                consolidate = WorksheetFunction.Min(vals)
            Case "middle"
                consolidate = WorksheetFunction.Median(vals)
            Case Else
                err.Raise CONSOLIDATE_ERROR, , "Consolidation method '" & LCase(method) & "' is not available"
        End Select
    End If
    Exit Function
finish:
    consolidate = ""
    If err.Number = CONSOLIDATE_ERROR Then
        err.Raise err.Number, , err.Description
    End If
End Function

