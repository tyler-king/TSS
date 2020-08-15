Option Private Module

Public Function collectionToVariant(col As Collection, Optional forceType As Integer = 0) As Variant
Dim l As Long, vals As Variant
ReDim vals(0 To 0)
l = col.Count
If l > 0 Then
    ReDim vals(1 To l)
    If forceType = 0 Then
        For i = 1 To l
            vals(i) = col(i)
        Next i
    ElseIf forceType = 1 Then
        For i = 1 To l
            vals(i) = CStr(col(i))
        Next i
    ElseIf forceType = 2 Then
        For i = 1 To l
            vals(i) = CDbl(col(i))
        Next i
    End If
End If
collectionToVariant = vals
End Function

Public Function zip_collection(ByVal col1 As Collection, ByVal col2 As Collection) As Scripting.Dictionary
    Set zip_collection = New Scripting.Dictionary
    If col1.Count <> col2.Count Then
        Exit Function
    End If
    Set zip_collection = New Scripting.Dictionary
    For i = 1 To col1.Count
        zip_collection.Add CStr(col1(i)), col2(i)
    Next i

End Function

Public Sub MyMultiDir(dfol As String)
    Dim iDir As Integer
    Dim i As Integer
    Dim path As String
    Dim nDirs As Variant
     On Error Resume Next
    
    If dfol <> "" Then          'if directory is not nothing then
        nDirs = Split(dfol, "\")            'split up by \ to break up the different folders
        If Left(dfol, 2) = "\\" Then        'if it is a network drive, then don't try to create the first folder - it is nothing
            iDir = 3
        Else
            iDir = 1
        End If                              'start position established
 
        path = Left(dfol, InStr(iDir, dfol, "\"))           'define the new path
 
        For i = iDir To UBound(nDirs)
            path = path & nDirs(i) & "\"
            
            If Len(Dir(path, vbDirectory)) = 0 Then
            'If Dir(path, vbDirectory) = vbNullString Then       'create new directory
                MkDir path
                
            Else: End If
        Next i                                                  'loop through to each new folder
        
    End If
End Sub


Public Function unionOr(ByRef rng As Range, ByRef ocell As Range) As Range

If rng Is Nothing Then
    Set rng = ocell
Else
    Set rng = Union(ocell, rng)
End If
Set unionOr = rng
End Function

Public Function array_push(ByRef arr As Variant, l As Variant)
    arr(UBound(arr)) = l
    ReDim Preserve arr(LBound(arr) To UBound(arr) + 1)
End Function

Public Function in_collection(col As Variant, l As Variant) As Boolean
Dim obj As Variant
On Error GoTo err
    in_collection = False
    For Each oi In col
       If LCase(oi) = LCase(l) Then
           in_collection = True
           Exit Function
       End If
    Next oi
err:
    in_collection = False
End Function

Public Function in_array(ByVal arr As Variant, l As Variant)
     in_array = False
     For Each oi In arr
        If LCase(oi) = LCase(l) Then
            in_array = True
            Exit Function
        End If
     Next oi
End Function


Public Function ConvertUTCToFreq(utc As Variant, freq As String)
    
    Dim O As utc
    Set O = New utc
    O.frequency = freq
    O.exactUTC = utc
    ConvertUTCToFreq = CStr(O.getUTCByFrequency())
    
End Function

Public Function UTCToString(utc As Variant, Optional fformat As String = "mm/dd/yyyy")
    
    Dim O As utc
    Set O = New utc
    O.exactUTC = utc
    O.getUTCByFrequency
    UTCToString = Format(O.dateString, fformat)
    
End Function

Public Function UTCFromFreq(excelDate As Variant, freq As String)
    Dim O As utc
    On Error Resume Next
    dt = CDate(excelDate)
    If err.Number = 0 Then
        Set O = New utc
        O.frequency = freq
        O.exactDate = dt
        UTCFromFreq = CStr(O.getUTCByFrequency())
    End If
    
End Function


Public Function cellRef(rng As Range, ffind As String) As Range

Set cellRef = rng.Find(ffind, , , xlWhole, , , False)

End Function

Public Function getCellValue(ocell As Range) As Variant
    If IsError(ocell) Then
        getCellValue = ""
        Exit Function
    End If
    getCellValue = Trim(ocell.Value)
End Function

Public Function rangeInput(rng As Range, arr As Variant, Optional transpose = False) As Range
Application.EnableEvents = False
Application.Calculation = xlCalculationManual
Set rangeInput = Nothing
If Not transpose Then
    For j = LBound(arr) To UBound(arr)
        rng.Offset(j).Resize(1, UBound(arr(j)) + 1).Value = arr(j)
        unionOr rangeInput, rng.Offset(j).Resize(1, UBound(arr(j)) + 1)
    Next j
Else
    For j = LBound(arr) To UBound(arr)
        rng.Offset(, j).Resize(UBound(arr(j)) + 1, 1).Value = Application.transpose(arr(j))
        unionOr rangeInput, rng.Offset(, j).Resize(UBound(arr(j)) + 1, 1)
    Next j
End If
Application.Calculation = xlAutomatic
Application.EnableEvents = True
End Function

Sub SortDictionaryByKey(ByRef Dict As Scripting.Dictionary)

    'Set a reference to Microsoft Scripting Runtime by using
    'Tools > References in the Visual Basic Editor (Alt+F11)

    'Declare the variables

    Dim TempDict As Scripting.Dictionary
    Dim KeyVal As Variant
    Dim arr() As Variant
    Dim Temp As Variant
    Dim txt As String
    Dim i As Long
    Dim j As Long
    
    
    'Set the comparison mode to perform a textual comparison
    'Dict.CompareMode = TextCompare
    
    'Allocate storage space for the dynamic array
    ReDim arr(0 To Dict.Count - 1)
    
    'Fill the array with the keys from the Dictionary
    For i = 0 To Dict.Count - 1
        arr(i) = Dict.keys(i)
    Next i
    
    'Sort the array using the bubble sort method
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If arr(i) > arr(j) Then
                Temp = arr(j)
                arr(j) = arr(i)
                arr(i) = Temp
            End If
        Next j
    Next i
    
    'Create an instance of the temporary Dictionary
    Set TempDict = New Dictionary
    
    'Add the keys and items to the temporary Dictionary,
    'using the sorted keys from the array
    For i = LBound(arr) To UBound(arr)
        KeyVal = arr(i)
        TempDict.Add Key:=KeyVal, Item:=Dict.Item(KeyVal)
    Next i
    
    'Set the Dict object to the TempDict object
    Set Dict = TempDict
    
    'Build a list of keys and items from the original Dictionary
    For i = 0 To Dict.Count - 1
        txt = txt & Dict.keys(i) & vbTab & Dict.Items(i) & vbCrLf
    Next i

End Sub

Public Function utcDate(ncell As Variant, Optional tzOffset As Double = 0) As LongLong
    Dim utcD As utc
    On Error Resume Next
    dt = CDate(ncell)
    If err.Number = 0 Then
        Set utcD = New utc
        utcD.exactDate = dt
        utcDate = utcD.utcDate(tzOffset)
    End If
End Function

Public Function TimeOffset() As Double
    Dim dString As Double
    dString = DateSerial(1970, 1, 1)
    TimeOffset = ParseUtc(CDate(dString)) - CDate(dString)
End Function


