Public ThisWS As clsFullWS
Public tssHeaders As Scripting.Dictionary
Public isValidSheet As Boolean
Public startFilterCell As Range
Public isRow As Boolean
Public codeList As Scripting.Dictionary
Public indexList As Collection
Public settings As Scripting.Dictionary
Public hadErrors As Boolean
Private selectedRange As Range
Private seriesRange As Range
Private resourceRange As Range
Private functionRegistrar As Scripting.Dictionary
Private firstHeaders As Collection

Const SERIES_CODE = "Retrieve_Code"
Const DATABASE = "Retrieve_Path"
Const RESOURCE = "Resource"

Const TSS_ID_KEY = "__id"
Const TSS_VALUES_KEY = "__values"
Const TSS_TIMESTAMP_KEY = "timestamp"
Const TSS_VALUE_KEY = "value"

Const HEADER_SCALE = "scale()"
Const HEADER_FREQ = "frequency()"
Const HEADER_FORMULA = "convert()"
Const HEADER_ERRORS = "errors()"

Const SEP_CHARACTER = "~"
Const JOIN_DELIM = "|"

Const FUNCTION_NAME = "function"
Const ARGUMENTS_NAME = "arguments"
Const KEY_NAME = "keyed"
Const FORMULA_ERROR = 45


Private Sub Class_Initialize()
    isValidSheet = False
    isRefreshSheet = False
    isRow = False
    Set codeList = New Scripting.Dictionary
    Set indexList = New Collection
    Set settings = New Scripting.Dictionary
    Set seriesRange = Nothing
    Set resourceRange = Nothing
    Set startFilterCell = Nothing
    Set functionRegistrar = New Scripting.Dictionary
    Set firstHeaders = New Collection
    hadErrors = False
End Sub

Public Sub setHeaders(h1 As String, h2 As String)
    Set firstHeaders = New Collection
    firstHeaders.Add h1
    firstHeaders.Add h2
End Sub

Public Function setFunctionLibrary(functions As Scripting.Dictionary)
    If Not functions Is Nothing Then
        Set functionRegistrar = functions
    End If
End Function

Public Function setWorksheet(ByRef oWS As Excel.Worksheet, Optional sheetRange As Range)
    Set ThisWS = New clsFullWS
    Set ThisWS.ws = oWS
    If Not sheetRange Is Nothing Then
        Set selectedRange = sheetRange
    End If
    If firstHeaders.Count = 2 Then
        Set tssHeaders = getHeaders(ThisWS.ThisWS)
        If tssHeaders.Exists("required") Then
           isValidSheet = True
        End If
    Else
        setHeaders DATABASE, SERIES_CODE
        Set tssHeaders = getHeaders(ThisWS.ThisWS)
        If tssHeaders.Exists("required") Then
           isValidSheet = True
        End If
    End If
    If isValidSheet Then
        'Set settings = getUserSettings()
        getSeriesRange
    End If
End Function



Private Function getHeaders(ws As Worksheet) As Scripting.Dictionary
Dim dts As Range, ocell As Range, hds As Range, rq As Range, toSearch As Range, O As Range, shead As String, funcs As Range, intI As Range
i = 0
Set toSearch = ws.UsedRange

'CHECK REQUIRED HEADERS AND PUT IN SCRIPTING
Set getHeaders = New Scripting.Dictionary
getHeaders.Add "required", New Scripting.Dictionary
getHeaders.Add "headers", Nothing
getHeaders.Add "dates", Nothing

For Each head In firstHeaders
    shead = LCase(head)
    Set rq = cellRef(toSearch, shead)
    If Not rq Is Nothing Then
        getHeaders.Item("required").Add shead, Nothing
        Set getHeaders.Item("required").Item(shead) = rq
        If i = 0 Then
            Set toSearch = Union(rq.EntireRow, rq.EntireColumn)
            Set O = rq
            i = i + 1
        Else
            If rq.Row = O.Row Then
                Set toSearch = rq.EntireRow
                isRow = True
            Else
                Set toSearch = rq.EntireColumn
            End If
        End If
    Else
        Set getHeaders = New Scripting.Dictionary
        Exit Function
    End If
Next head

Set rng = Intersect(ws.UsedRange, toSearch)
On Error Resume Next
For Each ocell In rng.Cells
    If Not IsEmpty(ocell) Then
        err.Number = 0
        n = CDate(ocell)
        If err.Number = 0 Then
            unionOr dts, ocell
        Else
            If Right(ocell, 2) = "()" Then
                unionOr funcs, ocell
            ElseIf Not in_array(firstHeaders, ocell.Value) Then
                unionOr hds, ocell
            Else
                If getHeaders.Item("required").Item(LCase(ocell.Value)).Address <> ocell.Address Then
                    Set getHeaders = New Scripting.Dictionary
                    Exit Function
                End If
            End If
        End If
    End If
Next ocell
Set getHeaders.Item("functions") = funcs
If Not hds Is Nothing Then
     If Not selectedRange Is Nothing Then
        If isRow Then
            Set hds = Intersect(hds, selectedRange.EntireColumn)
        Else
            Set hds = Intersect(hds, selectedRange.EntireRow)
        End If
    End If
End If
Set getHeaders.Item("headers") = hds
Set getHeaders.Item("dates") = dts
End Function

Private Sub getSeriesRange()
    Dim ocell As Range
    Set rRequired = tssHeaders.Item("required")
    Set db = rRequired.Item(whichDatabase)
    Set sc = rRequired.Item(whichHeader)

    scoloffset = 0
    srowoffset = 0
    If isRow Then
        Set dbRefreshRange = Intersect(ThisWS.ThisWS.UsedRange, db.EntireColumn)
        scoloffset = sc.Column - db.Column
        If Not selectedRange Is Nothing Then
            Set dbRefreshRange = Intersect(dbRefreshRange, selectedRange.EntireRow)
        End If
    Else
        Set dbRefreshRange = Intersect(ThisWS.ThisWS.UsedRange, db.EntireRow)
        srowoffset = sc.Row - db.Row
        If Not selectedRange Is Nothing Then
            Set dbRefreshRange = Intersect(dbRefreshRange, selectedRange.EntireColumn)
        End If
    End If
    
    For Each ocell In dbRefreshRange.Cells
        If (isRow And ocell.Row <> rRequired.Item(whichDatabase).Cells(1, 1).Row) Or (Not isRow And ocell.Column <> rRequired.Item(whichDatabase).Cells(1, 1).Column) Then
            unionOr seriesRange, ocell.Offset(srowoffset, scoloffset)
            unionOr resourceRange, ocell
        End If
    Next ocell
End Sub

Public Function getCodes() As Scripting.Dictionary
    Dim ocell As Range
   
    If codeList.Count = 0 Then
        For Each ocell In seriesRange.Cells
            sseries = Array(getCellValue(ocell))
            If validSeries(sseries) Then
                jjoin = Trim(LCase(Join(sseries, "")))
                If jjoin <> "" Then
                    If isRow Then
                        resourceType = getCellValue(resourceRange.Parent.Cells(ocell.Row, resourceRange.Column))
                    Else
                        resourceType = getCellValue(resourceRange.Parent.Cells(resourceRange.Row, ocell.Column))
                    End If
                  
                    resourceType = Trim(LCase(resourceType))
                    If resourceType <> "" And Left(resourceType, 1) <> "#" Then 'comments out the line
                        If Not codeList.Exists(resourceType) Then
                            codeList.Add resourceType, New Scripting.Dictionary
                        End If
                        With codeList.Item(resourceType)
                            If Not .Exists(jjoin) Then
                                .Add jjoin, ocell
                            End If
                        End With
                        indexList.Add Array(ocell, resourceType, resourceType, jjoin)
                    End If
                End If
            End If
        Next ocell
    End If
    Set getCodes = codeList
End Function

Private Function getConvertedDates() As Scripting.Dictionary
Set dts = tssHeaders.Item("dates")
Set getConvertedDates = New Scripting.Dictionary
If Not dts Is Nothing Then
    For Each ncell In dts.Cells
        ddt = ncell.Value
        If Not getConvertedDates.Exists(ddt) Then
            getConvertedDates.Add ddt, Nothing
        Else
            Set ncell = Union(ncell, getConvertedDates.Item(ddt))
        End If
        Set getConvertedDates.Item(ddt) = ncell
    Next ncell
End If
End Function

Public Sub fillSheet(rresults As Scripting.Dictionary)
    Dim origcell As Range, dcell As Range, toHighlight As Range, backHighlight As Range, missingCell As Range, currentSeries As Scripting.Dictionary, parsedFormula As Scripting.Dictionary
    Dim currentfreq As String, eHeader As Range, i As Long, j As Long, errorList As Collection
    mainId = seriesCodeKey()
   
   
    dbMissing = getUserSettings().Item(SETTING_DB_MISSING)
    defMissing = getUserSettings().Item(SETTING_TS_DEF_MISSING)
    valuemissing = getUserSettings().Item(SETTING_VALUE_MISSING)
    metadatavaluemissing = Null
    highlightDifferent = getUserSettings().Item(SETTING_DIF_HIGHLIGHT)

    Set convDates = getConvertedDates
    datekeys = convDates.keys
    
    Set scaleMapping = defaultScale().Item("value")
    
    Set memoizer = getMemoizedDates()
    Set rng = tssHeaders.Item("headers")
    Set funcs = tssHeaders.Item("functions")
    
    Set foundCodes = New Scripting.Dictionary
    Set missingCodes = New Scripting.Dictionary
    For Each resourceType In rresults.keys
        itemId = 0
        For Each O In rresults.Item(resourceType).Item("data")
             itemId = itemId + 1
             If O.Item("type") <> "mget_miss" Then
                foundCodes.Add resourceType & SEP_CHARACTER & O.Item("type") & SEP_CHARACTER & LCase(O.Item(mainId)), itemId
             Else
                missingCodes.Add resourceType & SEP_CHARACTER & LCase(O.Item(mainId)), O.Item("attributes").Item("reason")
             End If
        Next O
    Next resourceType

    ccount = indexList.Count
    divisor = WorksheetFunction.Max(1, WorksheetFunction.Min(50, Round(ccount / 10, 0)))
    j = -1
    For Each indexItem In indexList
        j = j + 1
        rErr = ""
        Set errorList = New Collection
        requestType = indexItem(1)
        resourceType = indexItem(2)
        oitem = indexItem(3)
        Set ocell = indexItem(0)
        If j Mod ccount = 0 Then
            Application.StatusBar = "Refreshing " & j + 1 & " of " & ccount
        End If
        If foundCodes.Exists(requestType & SEP_CHARACTER & resourceType & SEP_CHARACTER & oitem) Then
            
            Set currentSeries = rresults.Item(requestType).Item("data").Item(foundCodes.Item(requestType & SEP_CHARACTER & resourceType & SEP_CHARACTER & oitem))
            If Not rng Is Nothing Then
                 For Each rngHead In rng.Cells
                    If isRow Then
                        Set dcell = ThisWS.ThisWS.Cells(ocell.Row, rngHead.Column)
                    Else
                        Set dcell = ThisWS.ThisWS.Cells(rngHead.Row, ocell.Column)
                    End If
                    If highlightDifferent <> False Then
                        Oval = dcell.Value
                    End If
                    If currentSeries.Item("attributes").Exists(LCase(rngHead.Value)) Then
                        On Error Resume Next
                        dcell.Value = currentSeries.Item("attributes").Item(LCase(rngHead.Value))
                        If err.Number <> 0 Then
                            dcell.Value = Join(collectionToVariant(currentSeries.Item("attributes").Item(LCase(rngHead.Value)), 1), JOIN_DELIM)
                        End If
                        On Error GoTo 0
                        On Error GoTo -1
                    Else
                        If LCase(rngHead.Value) = TSS_ID_KEY Then
                            If currentSeries.Exists(LCase(rngHead.Value)) Then
                                dcell.Value = currentSeries.Item(LCase(rngHead.Value))
                            End If
                        Else
                            dcell.Value = metadatavaluemissing
                        End If
                    End If
                    If highlightDifferent <> False Then
                      If Oval <> dcell.Value Then
                          unionOr toHighlight, dcell
                      Else
                          unionOr backHighlight, dcell
                      End If
                    Else
                        unionOr backHighlight, dcell
                    End If
                Next rngHead
            End If
            If currentSeries.Item("attributes").Exists("__values") Then
                sscale = 1
                currentfreq = ""
                Set parsedFormula = Nothing
                Set eHeader = Nothing
                Set ovals = Nothing
                If Not funcs Is Nothing Then
                    For Each func In funcs.Cells
                        If isRow Then
                            Set dcell = ThisWS.ThisWS.Cells(ocell.Row, func.Column)
                        Else
                            Set dcell = ThisWS.ThisWS.Cells(func.Row, ocell.Column)
                        End If
                        tempval = getCellValue(dcell)
                        shead = LCase(func.Value)
                        If tempval <> "" Then
                             If Left(tempval, 1) = "@" Then
                                tempval = currentSeries.Item("attributes").Item(Mid(LCase(tempval), 2))
                             End If
                             If shead = HEADER_SCALE Then
                                If scaleMapping.Exists(LCase(tempval)) Then
                                    sscale = scaleMapping.Item(LCase(tempval))
                                Else
                                    errorList.Add "Scale '" & tempval & "' does not exist"
                                End If
                             ElseIf shead = HEADER_FREQ Then
                                currentfreq = tempval
                             ElseIf shead = HEADER_FORMULA Then
                                Set parsedFormula = parseFormula(tempval)
                             End If
                        End If
                        If shead = HEADER_ERRORS Then
                            Set eHeader = dcell
                        End If
                    Next func
                End If
                On Error Resume Next
                Set ovals = runFormula(currentSeries.Item("attributes"), parsedFormula)
                If err.Number = FORMULA_ERROR Then
                    errorList.Add err.Description
                    rErr = Null
                ElseIf err.Number > 0 Then
                    errorList.Add "Error in convert(). " & err.Description
                    rErr = Null
                End If
                On Error GoTo 0
                On Error GoTo -1
               
                If Not ovals Is Nothing Then
                    If convDates.Count > 0 Then
                        For i = LBound(datekeys) To UBound(datekeys)
                            If isRow Then
                                Set dcell = Intersect(ocell.EntireRow, convDates.Item(datekeys(i)).EntireColumn)
                            Else
                                Set dcell = Intersect(convDates.Item(datekeys(i)).EntireRow, ocell.EntireColumn)
                            End If
    
                            If Not memoizer.Exists(datekeys(i)) Then
                                memoizer.Add datekeys(i), New Scripting.Dictionary
                            End If
                            If Not memoizer.Item(datekeys(i)).Exists(currentfreq) Then
                                memoizer.Item(datekeys(i)).Add currentfreq, UTCFromFreq(datekeys(i), currentfreq)
                            End If
                            
                            tStamp = memoizer.Item(datekeys(i)).Item(currentfreq)
                            
                            If highlightDifferent <> False Then
                              Oval = dcell.Value
                              Set origVals = New Scripting.Dictionary
                              If dcell.Cells.Count > 1 Then
                                For Each origcell In dcell
                                    origVals.Add origcell.Address, getCellValue(origcell)
                                Next origcell
                              End If
                            End If
                            If ovals.Exists(tStamp) Then
                                If IsNumeric(ovals.Item(tStamp)) Then
                                    dcell.Value = ovals.Item(tStamp) / sscale
                                Else
                                    dcell.Value = valuemissing
                                End If
                            Else
                                dcell.Value = valuemissing
                            End If
                            If highlightDifferent <> False Then
                              If dcell.Cells.Count > 1 Then
                                v = getCellValue(dcell.Cells(1, 1))
                                For Each origcell In dcell
                                    If origVals.Item(origcell.Address) <> v Then
                                      unionOr toHighlight, origcell
                                    Else
                                      unionOr backHighlight, origcell
                                    End If
                                Next origcell
                              Else
                                If IsError(Oval) And Not IsError(dcell) Then
                                    unionOr toHighlight, dcell
                                Else
                                    If Oval <> dcell.Value Then
                                      unionOr toHighlight, dcell
                                    Else
                                      unionOr backHighlight, dcell
                                    End If
                                End If
                              End If
                            Else
                                unionOr backHighlight, dcell
                            End If
                        
                        Next i
                    End If
                Else
                    Set missingCell = Nothing
                    If convDates.Count > 0 Then
                        For i = LBound(datekeys) To UBound(datekeys)
                            If isRow Then
                                unionOr missingCell, Intersect(ocell.EntireRow, convDates.Item(datekeys(i)).EntireColumn)
                            Else
                                unionOr missingCell, Intersect(convDates.Item(datekeys(i)).EntireRow, ocell.EntireColumn)
                            End If
                        Next i
                    End If
                    missingCell.Value = valuemissing
                    If highlightDifferent <> False Then
                        unionOr toHighlight, missingCell
                    Else
                        unionOr backHighlight, missingCell
                    End If
                End If
            End If
            If Not eHeader Is Nothing Then
                eHeader.Value = Join(collectionToVariant(errorList), ";")
                hadErrors = True
            End If
        Else
            rErr = defMissing
            If missingCodes.Exists(resourceType & SEP_CHARACTER & oitem) Then
                Select Case missingCodes.Item(resourceType & SEP_CHARACTER & oitem)
                    Case "database does not exist":
                        rErr = dbMissing
                End Select
            End If
            Set missingCell = Nothing
            If Not rng Is Nothing Then
                 For Each rngHead In rng.Cells
                    If isRow Then
                        unionOr missingCell, ThisWS.ThisWS.Cells(ocell.Row, rngHead.Column)
                    Else
                        unionOr missingCell, ThisWS.ThisWS.Cells(rngHead.Row, ocell.Column)
                    End If
                Next rngHead
            End If
            If convDates.Count > 0 Then
                For i = LBound(datekeys) To UBound(datekeys)
                    If isRow Then
                        unionOr missingCell, Intersect(ocell.EntireRow, convDates.Item(datekeys(i)).EntireColumn)
                    Else
                        unionOr missingCell, Intersect(convDates.Item(datekeys(i)).EntireRow, ocell.EntireColumn)
                    End If
                Next i
            End If
            missingCell.Value = rErr
            If highlightDifferent <> False Then
                unionOr toHighlight, missingCell
            Else
                unionOr backHighlight, missingCell
            End If
        End If
    Next indexItem
    
    If Not toHighlight Is Nothing Then
        toHighlight.Font.Color = vbRed
    End If
    If Not backHighlight Is Nothing Then
        backHighlight.Font.Color = vbBlack
    End If
    
End Sub

Public Function pullSheet() As Scripting.Dictionary

Dim dcell As Range, currentfreq As String, errorList As Collection, originalSeries As Scripting.Dictionary, grabValues As Scripting.Dictionary

Set convDates = getConvertedDates
datekeys = convDates.keys

Set scaleMapping = defaultScale().Item("value")

Set memoizer = getMemoizedDates()
Set rng = tssHeaders.Item("headers")
Set funcs = tssHeaders.Item("functions")
scKey = seriesCodeKey
Set originalSeries = New Scripting.Dictionary
Set foundCodes = New Scripting.Dictionary

ccount = indexList.Count
j = -1
errorIndex = errorColumn
divisor = WorksheetFunction.Max(1, WorksheetFunction.Min(50, Round(ccount / 10, 0)))

For Each indexItem In indexList
    j = j + 1
    rErr = ""
    Set errorList = New Collection
    dbName = indexItem(1)
    resourceType = indexItem(2)
    oitem = indexItem(3)
    Set ocell = indexItem(0)
    If j Mod divisor = 0 Then
        Application.StatusBar = "Saving " & j + 1 & " of " & ccount
    End If
    If Not originalSeries.Exists(dbName) Then
        foundCodes.Add dbName, New Scripting.Dictionary
        originalSeries.Add dbName, New Scripting.Dictionary
        originalSeries.Item(dbName).Add "data", New Collection
    End If
    With originalSeries.Item(dbName)
        With .Item("data")
            If Not foundCodes.Item(dbName).Exists(oitem) Then
                Set dataItem = New Scripting.Dictionary
                dataItem.Add scKey, oitem
                dataItem.Add "attributes", New Scripting.Dictionary
                foundCodes.Item(dbName).Add oitem, oitem
                
                With dataItem.Item("attributes")
                    If Not rng Is Nothing Then
                        For Each ncell In rng.Cells
                            shead = LCase(ncell.Value)
                            If isRow Then
                                Set dcell = ThisWS.ThisWS.Cells(ocell.Row, ncell.Column)
                            Else
                                Set dcell = ThisWS.ThisWS.Cells(ncell.Row, ocell.Column)
                            End If
                            tempval = getCellValue(dcell)
                            If tempval <> "" Then
                                If Not .Exists(shead) Then
                                    .Add shead, ""
                                End If
                                .Item(shead) = tempval
                            End If
                        Next ncell
                    End If
               
                   
                  
                    sscale = 1
                    currentfreq = ""
                    If Not funcs Is Nothing Then
                        For Each func In funcs.Cells
                            If isRow Then
                                Set dcell = ThisWS.ThisWS.Cells(ocell.Row, func.Column)
                            Else
                                Set dcell = ThisWS.ThisWS.Cells(func.Row, ocell.Column)
                            End If
                            tempval = getCellValue(dcell)
                            If tempval <> "" Then
                                 shead = LCase(func.Value)
                                 If shead = HEADER_SCALE Then
                                    If scaleMapping.Exists(LCase(tempval)) Then
                                        sscale = scaleMapping.Item(LCase(tempval))
                                    Else
                                        errorList.Add "Scale '" & tempval & "' does not exist"
                                    End If
                                 ElseIf shead = HEADER_FREQ Then
                                    currentfreq = tempval
                                 End If
                            End If
                        Next func
                    End If
        
                    If convDates.Count > 0 Then
                        Set grabValues = New Scripting.Dictionary
                        For i = LBound(datekeys) To UBound(datekeys)
                            If isRow Then
                                Set dcell = ThisWS.ThisWS.Cells(ocell.Row, convDates.Item(datekeys(i)).Column)
                            Else
                                Set dcell = ThisWS.ThisWS.Cells(convDates.Item(datekeys(i)).Row, ocell.Column)
                            End If
                            If Not memoizer.Exists(datekeys(i)) Then
                                memoizer.Add datekeys(i), New Scripting.Dictionary
                            End If
                            If Not memoizer.Item(datekeys(i)).Exists(currentfreq) Then
                                memoizer.Item(datekeys(i)).Add currentfreq, UTCFromFreq(datekeys(i), currentfreq)
                            End If
        
                            tStamp = memoizer.Item(datekeys(i)).Item(currentfreq)
                            If IsNumeric(dcell.Value) And Not IsEmpty(dcell) Then
                                dval = dcell.Value * sscale
                                If Not grabValues.Exists(tStamp) Then
                                    grabValues.Add tStamp, dval
                                Else
                                    grabValues.Item(tStamp) = dval
                                End If
                            ElseIf Not IsNumeric(dcell.Value) And Not IsEmpty(dcell) Then
                                cVal = getCellValue(dcell)
                                If cVal <> "" Then
                                    grabValues.Item(tStamp) = cVal
                                End If
                            End If
                        Next i
                        
                        Set currentValues = New Scripting.Dictionary
                        currentValues.Add TSS_TIMESTAMP_KEY, New Collection
                        currentValues.Add TSS_VALUE_KEY, New Collection
                        
                        For Each oldkey In grabValues.keys
                            currentValues.Item(TSS_TIMESTAMP_KEY).Add oldkey
                            currentValues.Item(TSS_VALUE_KEY).Add grabValues.Item(oldkey)
                        Next oldkey
            
                        'currentValues.Item(TSS_TIMESTAMP_KEY) = grabValues.keys
                        'currentValues.Item(TSS_VALUE_KEY) = grabValues.Items
                        .Add TSS_VALUES_KEY, currentValues
                    End If
                End With
                .Add dataItem
            Else
                errorList.Add "Duplicate Series: " & oitem
            End If
        End With
    End With
    
    
    If errorList.Count > 0 Then
        hadErrors = True
        If errorIndex > 0 Then
            If isRow Then
                Set eHeader = ThisWS.ThisWS.Cells(ocell.Row, errorIndex)
            Else
                Set eHeader = ThisWS.ThisWS.Cells(errorIndex, ocell.Column)
            End If
            If Not eHeader Is Nothing Then
                eHeader.Value = Join(collectionToVariant(errorList), ";")
            End If
        End If
    Else
        If errorIndex > 0 Then
            If isRow Then
                Set eHeader = ThisWS.ThisWS.Cells(ocell.Row, errorIndex)
            Else
                Set eHeader = ThisWS.ThisWS.Cells(errorIndex, ocell.Column)
            End If
            If Not eHeader Is Nothing Then
                eHeader.Value = ""
            End If
        End If
    End If

Next indexItem
Set pullSheet = originalSeries


End Function

Public Function seriesCodeKey() As String
    seriesCodeKey = LCase(TSS_ID_KEY)
End Function

Private Function whichHeader() As String
    whichHeader = LCase(firstHeaders.Item(2))
End Function
Private Function whichDatabase() As String
    whichDatabase = LCase(firstHeaders.Item(1))
End Function

Private Function hasHeader(headerName As String) As Boolean
    hasHeader = tssHeaders.Item("required").Exists(LCase(headerName))
End Function

Private Function validSeries(seriesInfo As Variant) As Boolean
    validSeries = True
End Function


Public Function errorColumn() As Long
errorColumn = -1
Set funcs = tssHeaders.Item("functions")
If Not funcs Is Nothing Then
    If Not funcs.Find("errors()", , , xlWhole) Is Nothing Then
       If isRow Then
            errorColumn = funcs.Find("errors()", , , xlWhole).Column
        Else
            errorColumn = funcs.Find("errors()", , , xlWhole).Row
        End If
    End If
End If
End Function


'Public Function hadErrors() As Boolean
'hadErrors = errorList.Count > 0
'End Function


Public Function getMemoizedDates() As Scripting.Dictionary
    If memoize Is Nothing Then 'must set memoize public in a module
        Set memoize = New Scripting.Dictionary
    End If
    Set getMemoizedDates = memoize
End Function

Private Function defaultScale() As Scripting.Dictionary
    Set defaultScale = New Scripting.Dictionary
    defaultScale.Add "type", "function"
    defaultScale.Add "value", New Scripting.Dictionary
    Set dScale = defaultScale.Item("value")
    With dScale
        .Add "trillion", 1000000000000#
        .Add "billion", 1000000000
        .Add "million", 1000000
        .Add "thousand", 1000
        .Add "hundred", 100
        .Add "ten", 10
        .Add "unit", 1
        .Add "tenth", 0.1
        .Add "hundredth", 0.001
        .Add "thousandth", 0.0001
        .Add "millionth", 0.0000001
        .Add "billionth", 0.0000000001
        .Add "trillionth", 0.0000000000001
    End With
End Function

Public Function parseFormula(txt As Variant) As Scripting.Dictionary

    Set parseFormula = New Scripting.Dictionary
    parseFormula.Add FUNCTION_NAME, ""
    parseFormula.Add ARGUMENTS_NAME, New Scripting.Dictionary
    parseFormula.Add KEY_NAME, New Scripting.Dictionary
    ttemp = Split(txt, "(")
    If UBound(ttemp) > 0 Then
        If Right(ttemp(UBound(ttemp)), 1) = ")" Then
            ttemp(UBound(ttemp)) = Mid(ttemp(UBound(ttemp)), 1, Len(ttemp(UBound(ttemp))) - 1)
            consols = Array("method", "frequency", "missing_frequency")
            parseFormula.Item(FUNCTION_NAME) = LCase(ttemp(0))
            ttemp = Join(ttemp, "(")
            ttemp = Mid(ttemp, Len(parseFormula.Item(FUNCTION_NAME)) + 2)
            ssplit = Split(ttemp, ",")
            For O = 0 To UBound(consols)
                parseFormula.Item(ARGUMENTS_NAME).Add consols(O), ""
                If UBound(ssplit) >= O Then
                    If Left(ssplit(O), 1) = "@" Then
                        parseFormula.Item(KEY_NAME).Add consols(O), Mid(ssplit(O), 2)
                    End If
                    parseFormula.Item(ARGUMENTS_NAME).Item(consols(O)) = ssplit(O)
                End If
            Next O
        End If
    Else
         parseFormula.Add "Error", "No () found in the formula text"
    End If
End Function

Public Function runFormula(currentSeries As Scripting.Dictionary, parsedFormula As Scripting.Dictionary) As Scripting.Dictionary
Dim ovals As Scripting.Dictionary, currentValues As Scripting.Dictionary
Set runFormula = Nothing
Set currentValues = currentSeries.Item(TSS_VALUES_KEY)
Set ovals = zip_collection(currentValues.Item(TSS_TIMESTAMP_KEY), currentValues.Item(TSS_VALUE_KEY))
On Error GoTo inv:
If parsedFormula Is Nothing Then
    Set runFormula = ovals
Else
     Dim oArgs() As Variant
     
     
     If parsedFormula.Exists("Error") Then
        err.Raise FORMULA_ERROR, , parsedFormula.Item("Error")
     End If
     ReDim oArgs(0 To parsedFormula.Item("arguments").Count - 1)
     i = 0
     For Each arg In parsedFormula.Item("arguments").keys
        If parsedFormula.Item("keyed").Exists(arg) Then
            oArgs(i) = currentSeries.Item(parsedFormula.Item("keyed").Item(arg))
        Else
            oArgs(i) = parsedFormula.Item("arguments").Item(arg)
        End If
        i = i + 1
     Next arg
   
     'TODO dot path could be here, also do arithmetic parse here, with foreach loop on each way
     functionname = parsedFormula.Item("function")
     If functionRegistrar.Exists(functionname) Then
        Set runFormula = functionRegistrar.Item(functionname).apply(ovals, oArgs)
     Else
        err.Raise FORMULA_ERROR, , "Formula '" & parsedFormula.Item("function") & "' is not available"
     End If
End If
inv:
If err.Number = FORMULA_ERROR Then
    Set runFormula = Nothing
    err.Raise err.Number, , err.Description
    Exit Function
ElseIf Not parsedFormula Is Nothing Then
    If functionRegistrar.Exists(parsedFormula.Item("function")) Then
        If functionRegistrar.Item(parsedFormula.Item("function")).IsError(err.Number) Then
            err.Raise err.Number, , err.Description
            Exit Function
        End If
    End If
End If

If err.Number > 0 Then
    Set runFormula = Nothing
    err.Raise FORMULA_ERROR, , "Formula '" & parsedFormula.Item("function") & "' is not available"
End If
End Function

Private Function operandParse()

End Function

Public Function exampleSaveSheet()

    exampleSaveSheet = Array( _
        Array("Save_Path", "Save_Code", HEADER_SCALE, HEADER_FREQ, HEADER_ERRORS, "1/1/2016", "1/1/2017", "1/1/2018"), _
        Array("#Database Path", "Unique Series Name", "Apply Scale to values", "Apply Frequency to date headers", "Shows errors", "", "", ""), _
        Array("C:\Temp\example_database", "111NGDP", "million", "m", "", 3, 4, 5) _
        )
    
   
End Function

Public Function exampleRetrieveSheet()

    exampleRetrieveSheet = Array( _
        Array("Retrieve_Path", "Retrieve_Code", HEADER_SCALE, HEADER_FREQ, HEADER_FORMULA, HEADER_ERRORS, "1/1/2016", "1/1/2017", "1/1/2018"), _
        Array("#Database Path", "Unique Series Name", "Apply Scale to values", "Apply Frequency to date headers", "Apply formula", "Shows errors", "", "", ""), _
        Array("C:\Temp\example_database", "111NGDP", "million", "m", "", "", 3, 4, 5) _
        )
    
End Function