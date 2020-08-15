'is not optimized for a sinlge series refresh among a bunch
'TODO allowing customized required fields
'TODO what happens to scale when only part of series gets resaved with scale
Option Private Module

Public Const CLEAR_CHAR = "~"
Public Const CODE_DELIM = "."
Public Const DB_EXTENSION = ".series"

Public Const SYSTEM_SETTING_TS_ROBUST = "TS_ROBUST"
Public Const SYSTEM_SETTING_METADATA_MAPPINGS = "METADATA_MAPPINGS"
Public Const SYSTEM_SETTING_VERSION = "VERSION"

Public memoize As Scripting.Dictionary

Public Function getFunctions() As Scripting.Dictionary
    Set getFunctions = New Scripting.Dictionary
    getFunctions.Add "consolidate", New clsConsolidate
    'getFunctions.Add "consolidate.average", New clsConsolidate
    'getFunctions.Item("consolidate.average").force "average"
    getFunctions.Add "math", New clsConsolidate
    'TODO how to run on multiple series?
    'formulas don't make sense in our context..., you can modify the raw series and record the modifications to backtrack...?
End Function

Public Sub refreshTimeSeries(Optional onlyRange As Boolean = False)

Set dDefault = New clsCodePrep
Dim rresults As Scripting.Dictionary, tssSheet As tssWS, resourceType As Variant, codes As Scripting.Dictionary
Set tssSheet = New tssWS
tssSheet.setFunctionLibrary getFunctions

If onlyRange Then
    If TypeName(Selection) <> "Range" Then
        MsgBox "You must select a valid Excel range"
        Exit Sub
    End If
    tssSheet.setWorksheet ActiveSheet, Selection.Cells
Else
    tssSheet.setWorksheet ActiveSheet
End If

If Not tssSheet.isValidSheet Then
    MsgBox "This isn't a valid refresh sheet"
    Exit Sub
End If

Dim scKey As String
scKey = tssSheet.seriesCodeKey
Set codes = tssSheet.getCodes()
If codes.Count > 0 Then
    Set rresults = getCodesFromFiles(codes, scKey)
    tssSheet.fillSheet rresults
End If

End Sub


Public Sub createTimeSeries(Optional onlyRange As Boolean = False)

Set dDefault = New clsCodePrep
Dim rresults As Scripting.Dictionary, tssSheet As tssWS, resourceType As Variant, codes As Scripting.Dictionary, hadErrors As Boolean, merged As Scripting.Dictionary
Set tssSheet = New tssWS
tssSheet.setHeaders "save_path", "save_code"
If onlyRange Then
    If TypeName(Selection) <> "Range" Then
        MsgBox "You must select a valid Excel range"
        Exit Sub
    End If
    tssSheet.setWorksheet ActiveSheet, Selection.Cells
Else
    tssSheet.setWorksheet ActiveSheet
End If

If Not tssSheet.isValidSheet Then
    MsgBox "This isn't a valid save sheet"
    Exit Sub
End If


Dim scKey As String
scKey = tssSheet.seriesCodeKey
Set codes = tssSheet.getCodes()
If codes.Count > 0 Then
   
    Set rresults = tssSheet.pullSheet()
    Set merged = mergeCodesFromFiles(rresults, scKey)
    
    saveSeriesToFiles merged
    hadErrors = tssSheet.hadErrors
    errorColumn = tssSheet.errorColumn
    Set tssSheet = Nothing
    If getUserSettings().Item(SETTING_SHOW_SAVE_LOG) Then
        If Not hadErrors Then
            MsgBox "Succesfully Saved"
        Else
            If errorColumn > -1 Then
                MsgBox "Saved with errors."
            Else
                MsgBox "Saved with errors. Use the errors() header to view them."
            End If
        End If
    End If
End If

End Sub

 
Public Function getCodesFromFiles(codeList As Scripting.Dictionary, scKey As String) As Scripting.Dictionary
Set getCodesFromFiles = New Scripting.Dictionary
Dim dbName As Variant, fso As FileSystemObject

If codeList.Count > 0 Then
    
    Set fso = New FileSystemObject
    Set odatabases = New Scripting.Dictionary
    Set eDatabases = New Scripting.Dictionary
   
    For Each dbName In codeList.keys
    
        Set singleFile = New Scripting.Dictionary
        singleFile.Add "data", New Collection
        For Each oitem In codeList.Item(dbName)
            scVal = oitem
            sseries = Array(scVal)
            If validSeries(sseries) Then
                rErr = ""
                On Error Resume Next
               
                If Not eDatabases.Exists(dbName) Then
                    If Not odatabases.Exists(dbName) Then
                        If fso.FileExists(dbName & DB_EXTENSION) Then
                            pwd = ""
                            odatabases.Add dbName, readJSON(dbName, pwd)
                            Set oseries = odatabases.Item(dbName).Item("timeseries")
                        Else
                            Set oseries = Nothing
                            rErr = "database does not exist"
                        End If
                    Else
                        Set oseries = odatabases.Item(dbName).Item("timeseries")
                    End If
                End If
                
                If err.Number <> 0 Then
                    Set oseries = Nothing
                    If Not eDatabases.Exists(dbName) Then
                        eDatabases.Add dbName, "Database file exists in a unreadable format"
                    End If
                End If
                If eDatabases.Exists(dbName) Then
                     rErr = eDatabases.Item(dbName)
                ElseIf oseries Is Nothing Then
                     rErr = "database does not exist"
                End If
                On Error GoTo 0
                On Error GoTo -1
                localname = LCase(Join(sseries, CODE_DELIM))
                If rErr = "" Then
                    If Not oseries.Exists(localname) Then
                        rErr = "timeseries does not exist"
                    Else
                        Set attr = New Scripting.Dictionary
                        attr.Add "attributes", New Scripting.Dictionary
                        attr.Add "type", dbName
                        attr.Add scKey, scVal
                        Set attr.Item("attributes") = oseries.Item(localname)
                        singleFile.Item("data").Add attr
                    End If
                End If
                If rErr <> "" Then
                    Set attr = New Scripting.Dictionary
                    attr.Add "attributes", New Scripting.Dictionary
                    attr.Add "type", "mget_miss"
                    attr.Add scKey, scVal
                    attr.Item("attributes").Add "reason", rErr
                    singleFile.Item("data").Add attr
                End If
            End If
        Next oitem
        getCodesFromFiles.Add dbName, New Scripting.Dictionary
        Set getCodesFromFiles.Item(dbName) = singleFile
    Next dbName
End If

End Function

Private Function mergeCodesFromFiles(originalSeries As Scripting.Dictionary, scKey As String)

Set mergeCodesFromFiles = New Scripting.Dictionary
Dim dbName As Variant, fso As FileSystemObject, settings As Scripting.Dictionary

If originalSeries.Count > 0 Then
    
    Set fso = New FileSystemObject
    Set odatabases = New Scripting.Dictionary
    Set eDatabases = New Scripting.Dictionary
    
    For Each dbName In originalSeries.keys
        noError = True

        Set singleFile = New Scripting.Dictionary
        On Error Resume Next
        If fso.FileExists(dbName & DB_EXTENSION) Then
            Set singleFile = readJSON(dbName)
            If err.Number <> 0 Then
                noError = False
               
                'fillAllSeriesWithError = "Database file exists in a unreadable format"
            End If
        End If
        On Error GoTo 0
        On Error GoTo -1
        If noError Then
            Set settings = New Scripting.Dictionary
            If Not singleFile.Exists("settings") Then
                Set settings = getLocalDatabaseSettings()
                singleFile.Add "settings", getLocalDatabaseSettings()
            Else
                Set settings = getLocalDatabaseSettings(singleFile.Item("settings"))
            End If
            Set singleFile.Item("settings") = settings
            'oitem.Item("attributes") = applySettings(oitem.Item("attributes"), settings)
            If Not singleFile.Exists("timeseries") Then
                singleFile.Add "timeseries", New Scripting.Dictionary
                For Each dataItem In originalSeries.Item(dbName).Item("data")
                    singleFile.Item("timeseries").Add dataItem.Item(scKey), New Scripting.Dictionary
                    Set singleFile.Item("timeseries").Item(dataItem.Item(scKey)) = dataItem.Item("attributes")
                Next dataItem
            Else
                For Each dataItem In originalSeries.Item(dbName).Item("data")
                    If Not singleFile.Item("timeseries").Exists(dataItem.Item(scKey)) Then
                        singleFile.Item("timeseries").Add dataItem.Item(scKey), New Scripting.Dictionary
                        Set singleFile.Item("timeseries").Item(dataItem.Item(scKey)) = dataItem.Item("attributes")
                    Else
                        Set series = singleFile.Item("timeseries").Item(dataItem.Item(scKey))
                        For Each metakey In dataItem.Item("attributes").keys
                            If metakey <> "__values" Then
                                cVal = dataItem.Item("attributes").Item(metakey)
                                If series.Exists(metakey) Then
                                    If cVal = CLEAR_CHAR Then
                                        series.Remove metakey
                                    Else
                                        series.Item(metakey) = cVal
                                    End If
                                Else
                                    If cVal <> CLEAR_CHAR Then
                                        series.Add metakey, cVal
                                    End If
                                End If
                            Else
                                If series.Exists(metakey) Then
                                    'TODO could index these out
                                    Set NewSeries = New Scripting.Dictionary
                                    Set oldSeries = New Scripting.Dictionary
                                    Set NewSeries = zip_collection(dataItem.Item("attributes").Item(metakey).Item("timestamp"), dataItem.Item("attributes").Item(metakey).Item("value"))
                                    Set oldSeries = zip_collection(series.Item(metakey).Item("timestamp"), series.Item(metakey).Item("value"))
                                    
                                    For Each newKey In NewSeries.keys
                                        If IsNumeric(NewSeries.Item(newKey)) Or NewSeries.Item(newKey) = CLEAR_CHAR Then
                                            oldSeries.Item(newKey) = NewSeries.Item(newKey)
                                        End If
                                    Next newKey
                                    
                                    Set oldSeriesArray = New Scripting.Dictionary
                                    oldSeriesArray.Add "timestamp", New Collection
                                    oldSeriesArray.Add "value", New Collection
                                    
                                    For Each oldkey In oldSeries.keys
                                        If oldSeries.Item(oldkey) <> CLEAR_CHAR Then
                                            oldSeriesArray.Item("timestamp").Add oldkey
                                            oldSeriesArray.Item("value").Add oldSeries.Item(oldkey)
                                        End If
                                    Next oldkey
                                    
                                    Set series.Item(metakey) = oldSeriesArray
                                    
                                Else
                                    series.Item(metakey) = dataItem.Item("attributes").Item(metakey)
                                End If
                            End If
                        Next metakey
                        Set singleFile.Item("timeseries").Item(dataItem.Item(scKey)) = series
                    End If
                Next dataItem
                
            End If
        End If
        If noError Then
            mergeCodesFromFiles.Add dbName, New Scripting.Dictionary
            Set mergeCodesFromFiles.Item(dbName) = singleFile
        End If
    Next dbName
End If


End Function

Private Sub saveSeriesToFiles(timeseries As Scripting.Dictionary)
    
    For Each kkey In timeseries.keys
        saveSeriesSingle CStr(kkey), timeseries.Item(kkey), ""
    Next kkey
End Sub

Private Function applySettings(attributes As Scripting.Dictionary, settings As Scripting.Dictionary)
'    Set forceMappings = settings.Item(SYSTEM_SETTING_METADATA_MAPPINGS)
'    If tempval <> "" Then
'        If forceMappings.Exists(shead) Then
'            Set mMap = forceMappings.Item(shead).Item("value")
'            If TypeName(mMap) = "Dictionary" Then
'                If Not mMap.Exists(LCase(tempval)) Then
'                    tempval = ""
'                End If
'            ElseIf TypeName(mMap) = "Collection" Then
'                If Not in_collection(mMap, LCase(tempval)) Then
'                     tempval = ""
'                End If
'            End If
'        End If
'    End If
'
'  If tempval <> prevfreq Then
'                                        If .Exists(TSS_VALUES_KEY) Then 'resets all data if frequency changes
'                                            Set grabValues = New Scripting.Dictionary
'                                        End If
'                                    End If
'                prevfreq = ""
'                If currentSeries.Exists(Left(HEADER_FREQ, Len(HEADER_FREQ) - 2)) Then
'                    prevfreq = currentSeries.Item(Left(HEADER_FREQ, Len(HEADER_FREQ) - 2))
'                End If

End Function

Private Function saveSeriesSingle(dbName As String, series As Scripting.Dictionary, Optional pwd = "")
    saveJSON cleanDatabaseName(dbName) & DB_EXTENSION, series, series.Item("settings"), pwd
End Function

Private Function cleanDatabaseName(dbName As String)
    If Right(dbName, 1) = "\" Then
        dbName = Left(dbName, Len(dbName) - 1)
    End If
    cleanDatabaseName = dbName
End Function

Private Function saveJSON(location As String, values As Scripting.Dictionary, setting As Scripting.Dictionary, Optional pwd = "")
    Dim fso As FileSystemObject, dirPath As String, CC As clsCryptoFilterBox
    Set fso = New FileSystemObject
    dirPath = fso.GetParentFolderName(location)
    MyMultiDir dirPath
    If Len(Trim(pwd)) = 0 Then
        pwd = Null
    End If
    
    On Error Resume Next
    Set ffile = fso.OpenTextFile(location, ForWriting, True)
    For Each okey In values.keys
        If Not IsNull(pwd) Then
            ffile.WriteLine "[" & okey & "]:" & encrypt(JSONConverter.ConvertToJson(values.Item(okey)), CStr(pwd))
        Else
            If setting.Item(SYSTEM_SETTING_TS_ROBUST) > -1 Then
                ffile.WriteLine "[" & okey & "]:" & JSONConverter.ConvertToJson(values.Item(okey), setting.Item(SYSTEM_SETTING_TS_ROBUST))
            Else
                ffile.WriteLine "[" & okey & "]:" & JSONConverter.ConvertToJson(values.Item(okey))
            End If
        End If
    Next okey
    If err.Number <> 0 Then
        MsgBox "Could not save to " & location & vbNewLine & "Reason:" & err.Description
    End If
    On Error GoTo 0
    On Error GoTo -1
    
End Function

Public Function readJSON(ByVal location As String, Optional pwd = "", Optional keys As Collection) As Scripting.Dictionary
    Set readJSON = New Scripting.Dictionary
    Dim fso As FileSystemObject, dirPath As String, CC As clsCryptoFilterBox, okey() As String, reader As clsBuildString
    Set fso = New FileSystemObject
    location = cleanDatabaseName(location) & DB_EXTENSION
    dirPath = fso.GetParentFolderName(location)
    MyMultiDir dirPath
    If Len(Trim(pwd)) = 0 Then
        pwd = Null
    End If
    If keys Is Nothing Then
        Set keys = New Collection
    End If
    On Error Resume Next
    ffile = Split(fso.OpenTextFile(location, ForReading, True).ReadAll, vbNewLine)
   
    Set reader = New clsBuildString
    reader.Delim = ""
    current = ""
    For Each f In ffile
        If Len(f) > 0 Then
            okey = Split(f, ":", 2)
            rkey = Mid(okey(0), 2, Len(okey(0)) - 2)
            If Left(okey(0), 1) = "[" And Right(okey(0), 1) = "]" Then
                If reader.Count > 0 Then
                    If IsNull(pwd) Then
                        readJSON.Add current, JSONConverter.ParseJson(reader.Text)
                    Else
                        readJSON.Add current, JSONConverter.ParseJson(decrypt(reader.Text, CStr(pwd)))
                    End If
                End If
                
                Set reader = New clsBuildString
                reader.Delim = ""
                current = ""
                If keys.Count > 0 Then
                    If in_collection(keys, rkey) Then
                        current = rkey
                        reader.Add okey(1)
                    End If
                Else
                    current = rkey
                    reader.Add okey(1)
                End If
            Else
                If current <> "" Then
                    reader.Add f
                End If
            End If
        End If
    Next f
    If reader.Count > 0 Then
        If IsNull(pwd) Then
            readJSON.Add current, JSONConverter.ParseJson(reader.Text)
        Else
            readJSON.Add current, JSONConverter.ParseJson(decrypt(reader.Text, CStr(pwd)))
        End If
    End If
    On Error GoTo 0
    On Error GoTo -1

End Function

Private Function decrypt(jOutput As String, pwd As String) As String
    Set CC = New clsCryptoFilterBox
    CC.Password = pwd
    CC.InBuffer = jOutput
    CC.decrypt
    decrypt = CC.OutBuffer
End Function

Private Function encrypt(jInput As String, pwd As String) As String
    Set CC = New clsCryptoFilterBox
    CC.Password = pwd
    CC.InBuffer = jInput
    CC.encrypt
    encrypt = CC.OutBuffer
End Function

Private Function validSeries(n As Variant) As Boolean
    validSeries = True
    invalidCharacters = INVALID_CHARACTERS()
    For Each i In n
        If Trim(i) = "" Then
            validSeries = False
            Exit Function
        End If
        For q = 1 To Len(i)
            If InStr(invalidCharacters, Mid(i, q, 1)) > 0 Then
                validSeries = False
                Exit Function
            End If
        Next q
    Next i
End Function

Private Function INVALID_CHARACTERS() As Variant
    INVALID_CHARACTERS = Join(Array("+", "-", "~", "/", "\", ":", "[", "]"), "")
End Function

Public Function getLocalDatabaseSettings(Optional setting As Scripting.Dictionary) As Scripting.Dictionary
    Set getLocalDatabaseSettings = New Scripting.Dictionary
    getLocalDatabaseSettings.Add SYSTEM_SETTING_TS_ROBUST, -1
    getLocalDatabaseSettings.Add SYSTEM_SETTING_METADATA_MAPPINGS, New Scripting.Dictionary
    getLocalDatabaseSettings.Add SYSTEM_SETTING_VERSION, "v1"
    Set getLocalDatabaseSettings.Item(SYSTEM_SETTING_METADATA_MAPPINGS) = exampleMappings()
    If Not setting Is Nothing Then
        For Each sett In setting
            If Not getLocalDatabaseSettings.Exists(sett) Then
                getLocalDatabaseSettings.Add sett, setting.Item(sett)
            Else
                If TypeName(getLocalDatabaseSettings.Item(sett)) = "Dictionary" Then
                    Set getLocalDatabaseSettings.Item(sett) = setting.Item(sett)
                Else
                    getLocalDatabaseSettings.Item(sett) = setting.Item(sett)
                End If
            End If
        Next sett
    End If
End Function

Public Function exampleMappings() As Scripting.Dictionary
    Set exampleMappings = New Scripting.Dictionary
End Function